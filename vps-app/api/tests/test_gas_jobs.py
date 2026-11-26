"""GAS の仕事（services/gas_jobs.py。実装仕様 7-3・7-4・12-2 の M）。偽の GasClient を入れ、外へは通信しない。"""
from __future__ import annotations

import base64
import json
import sys
import types
from datetime import datetime, timedelta, timezone
from types import SimpleNamespace

import pytest

from app import jobs, models, sheet_queue, worker
from app.jobs import JobRunner
from app.services import gas_jobs, write_gate
from app.services.gas_client import GasAuthError, GasTransportError

T0 = datetime(2026, 11, 26, 3, 0, tzinfo=timezone.utc)  # JST 12:00
APPROVED_AT = datetime(2026, 11, 26, 2, 30, 15, tzinfo=timezone.utc)  # JST 11:30:15
APPROVED_AT2 = datetime(2026, 11, 26, 2, 45, 0, tzinfo=timezone.utc)
PNG = b"\x89PNG\r\n\x1a\nfake-png-bytes"
STAGING_PING = {"ok": True, "scriptId": "staging-script-id", "spreadsheetId": "staging-db-sheet-id",
                "spreadsheetTimeZone": "", "scriptTimeZone": "Asia/Tokyo", "migrated": False,
                "settings": {"PDF_FOLDER_ID": "staging-pdf-folder-id", "LEDGER_SSID": "staging-ledger-id",
                             "MAIL_TO_SET": True},
                "ledgerKeyRows": 1, "ledgerOpenable": True, "pdfFolderOpenable": True, "auth": "bearer",
                "apiVersion": 1}


class Clock:
    def __init__(self, at=T0):
        self.at = at

    def __call__(self):
        return self.at

    def advance(self, **kw):
        self.at = self.at + timedelta(**kw)


@pytest.fixture
def clock(monkeypatch):
    from app import timeutil

    c = Clock()
    monkeypatch.setattr(timeutil, "now_utc", c)
    return c


class FakeGas:
    """GasClient.call と同じ形。replies[action] は dict・例外・それらの list（順に使う）・関数。"""

    def __init__(self, **replies):
        self.replies = replies
        self.calls: list[dict] = []

    def call(self, action, *, job_id, request_id="", expect=None, args=None, timeout=None):
        self.calls.append({"action": action, "jobId": job_id, "reqId": request_id, "expect": expect, "args": args})
        reply = self.replies.get(action)
        if isinstance(reply, list):
            reply = reply.pop(0) if len(reply) > 1 else reply[0]
        if callable(reply):
            reply = reply(action, args)
        if isinstance(reply, Exception):
            raise reply
        if reply is None:
            reply = {"ok": True}
        return {"action": action, "jobId": job_id, "reqId": request_id, **reply}


@pytest.fixture
def gas():
    fake = FakeGas()
    previous = gas_jobs.set_client_factory(lambda settings: fake)
    gas_jobs.register_worker()
    yield fake
    gas_jobs.set_client_factory(previous)


def add_request(db, req_id, *, worker_id="1001", status="承認済", approved_at=APPROVED_AT, approved_at2=None,
                sign_png=None, synced=True, **extra):
    from datetime import date

    row = models.LeaveRequest(req_id=req_id, worker_id=worker_id, status=status, approved_at=approved_at,
                              approved_at2=approved_at2, leave_date=date(2026, 11, 27), sign_png=sign_png,
                              approved_by1_email="boss1@lineworks-local.info", **extra)
    db.add(row)
    db.flush()
    sheet_queue.mark_dirty(db, "leave_requests", req_id)
    if synced:
        mark_synced(db, req_id)
    db.commit()
    return row


def mark_synced(db, req_id):
    state = db.get(models.SheetSyncState, ("leave_requests", req_id))
    state.synced_version = state.desired_version
    state.synced_values = {"STATUS": "承認済"}
    db.flush()


def ping_ok(db, at=T0, ok=True, migrated=False):
    db.add(models.SyncRun(kind="gas_ping", started_at=at, finished_at=at, ok=ok,
                          summary_json={"problems": [] if ok else ["x"], "migrationMode": migrated}))
    db.commit()


def run(session_factory, clock):
    return JobRunner(session_factory, lane="gas", now=clock).run_once()


def job(db, job_id):
    db.expire_all()
    return db.get(models.Job, job_id)


def req(db, req_id):
    db.expire_all()
    return db.get(models.LeaveRequest, req_id)


# ===== 積む（7-4 の表の行ごと） =====


def test_enqueue_after_approve1_kind_args_guard(db, clock, gas):
    add_request(db, "LV-20261126-001", sign_png=PNG)
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    row = job(db, job_id)
    assert row.kind == "internal.afterApprove1" and row.lane == "gas" and row.request_id == "LV-20261126-001"
    assert json.loads(row.args_json) == {}
    guard = json.loads(row.guard_json)
    assert guard["status"] == "承認済" and guard["approvedAt"] == APPROVED_AT.isoformat() and guard["approvedAt2"] == ""
    assert len(guard["gasJobId"]) == 36
    # PNG は jobs に入れない
    assert "base64" not in row.args_json and "PNG" not in row.args_json


def test_enqueue_after_approve2_keeps_old_pdf(db, clock, gas):
    add_request(db, "LV-20261126-002", approved_at2=APPROVED_AT2, pdf_url="https://drive.google.com/file/d/OLD1/view",
                pdf_file_id="OLD1")
    job_id = gas_jobs.enqueue_after_approve2(db, "LV-20261126-002")
    db.commit()
    row = job(db, job_id)
    assert row.kind == "internal.afterApprove2"
    assert json.loads(row.args_json) == {"oldPdfUrl": "https://drive.google.com/file/d/OLD1/view",
                                         "oldPdfFileId": "OLD1"}
    assert json.loads(row.guard_json)["approvedAt2"] == APPROVED_AT2.isoformat()


@pytest.mark.parametrize("reason,clear_sign,status", [("edit", True, "申請中"), ("cancel", False, "取消")])
def test_enqueue_after_unapprove(db, clock, gas, reason, clear_sign, status):
    add_request(db, "LV-20261126-003", status=status)
    job_id = gas_jobs.enqueue_after_unapprove(db, "LV-20261126-003", reason=reason, old_leave_date="2026-11-20",
                                              old_pdf_url="u", old_pdf_file_id="f", clear_sign=clear_sign)
    db.commit()
    row = job(db, job_id)
    assert row.kind == "internal.afterUnapprove"
    assert json.loads(row.args_json) == {"reason": reason, "oldLeaveDate": "2026-11-20", "oldPdfUrl": "u",
                                         "oldPdfFileId": "f", "clearSign": clear_sign}
    with pytest.raises(ValueError):
        gas_jobs.enqueue_after_unapprove(db, "LV-20261126-003", reason="delete", old_leave_date="", old_pdf_url="",
                                         old_pdf_file_id="", clear_sign=False)
    with pytest.raises(ValueError):
        gas_jobs.enqueue_after_approve1(db, "LV-NOPE")


def test_enqueue_send_warn_mail_once(db, clock, gas):
    first = gas_jobs.enqueue_send_warn_mail(db)
    second = gas_jobs.enqueue_send_warn_mail(db)
    db.commit()
    assert first == second
    row = job(db, first)
    assert row.kind == "internal.sendWarnMail" and row.request_id == "" and json.loads(row.requires_json) == []


def test_requires_include_same_worker_pending_rows(db, clock, gas):
    add_request(db, "LV-20261120-001", worker_id="1001", synced=False)  # 同じ作業員・反映待ち
    add_request(db, "LV-20261120-002", worker_id="1001")  # 同じ作業員・反映済み
    add_request(db, "LV-20261120-003", worker_id="2002", synced=False)  # 別の作業員
    add_request(db, "LV-20261126-001", worker_id="1001", synced=False)
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    requires = json.loads(job(db, job_id).requires_json)
    assert {r["pk"] for r in requires} == {"LV-20261126-001", "LV-20261120-001"}
    assert all(r["table"] == "leave_requests" and r["version"] >= 1 for r in requires)


# ===== 動かす前の確かめ =====


def test_does_not_run_until_requires_synced(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261120-001", synced=False)
    add_request(db, "LV-20261126-001")
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    assert run(session_factory, clock)["done"] == 0
    assert gas.calls == []
    assert job(db, job_id).state == "pending" and job(db, job_id).attempts == 0
    mark_synced(db, "LV-20261120-001")
    db.commit()
    clock.advance(seconds=6)
    assert run(session_factory, clock)["done"] == 1
    assert len(gas.calls) == 1


def test_waits_for_rows_dirtied_after_enqueue(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001")
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    sheet_queue.mark_dirty(db, "leave_requests", "LV-20261126-001")  # 積んだ後の編集
    db.commit()
    assert run(session_factory, clock)["waiting"] == 1
    assert gas.calls == [] and "row_not_synced" in job(db, job_id).last_error
    assert job(db, job_id).attempts == 0


def test_guard_mismatch_skips_and_clears_sign(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001", sign_png=PNG)
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    row = req(db, "LV-20261126-001")
    row.status = "取消"  # 承認の直後に取消
    db.commit()
    assert run(session_factory, clock)["skipped"] == 1
    assert gas.calls == []
    assert job(db, job_id).state == "skipped"
    assert req(db, "LV-20261126-001").sign_png is None


def test_guard_mismatch_keeps_png_for_later_approval(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001", sign_png=b"old")
    first = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    row = req(db, "LV-20261126-001")
    row.approved_at = APPROVED_AT + timedelta(minutes=5)  # 編集 → 承認し直し（新しいサイン）
    row.sign_png = b"new"
    second = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    run(session_factory, clock)
    assert job(db, first).state == "skipped"
    assert req(db, "LV-20261126-001").sign_png == b"new"
    run(session_factory, clock)
    assert job(db, second).state == "done"
    assert gas.calls[0]["args"]["signBase64"] == "data:image/png;base64," + base64.b64encode(b"new").decode()


def test_after_approve2_guard_on_approved_at2(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-002", approved_at2=APPROVED_AT2)
    job_id = gas_jobs.enqueue_after_approve2(db, "LV-20261126-002")
    db.commit()
    row = req(db, "LV-20261126-002")
    row.status, row.approved_at, row.approved_at2 = "申請中", None, None  # 編集で承認が外れた
    db.commit()
    run(session_factory, clock)
    assert job(db, job_id).state == "skipped" and gas.calls == []


def test_unapprove_is_never_skipped(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-003", status="申請中", approved_at=None)
    job_id = gas_jobs.enqueue_after_unapprove(db, "LV-20261126-003", reason="edit", old_leave_date="2026-11-20",
                                              old_pdf_url="", old_pdf_file_id="", clear_sign=True)
    row = req(db, "LV-20261126-003")
    row.status, row.approved_at = "承認済", APPROVED_AT  # 編集の直後に承認し直した
    db.commit()
    run(session_factory, clock)
    assert job(db, job_id).state == "done"
    assert gas.calls[0]["expect"] == {"status": "承認済"}  # 今の DB の状態で照合する


def test_gas_ping_required(session_factory, db, clock, gas):
    add_request(db, "LV-20261126-001")
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    assert run(session_factory, clock)["waiting"] == 1  # ping が無い
    row = job(db, job_id)
    assert row.run_after == T0 + timedelta(seconds=60) and row.attempts == 0 and gas.calls == []
    ping_ok(db, at=T0 - timedelta(minutes=11))  # 古い
    clock.advance(seconds=61)
    assert run(session_factory, clock)["waiting"] == 1
    ping_ok(db, at=clock.at, ok=False)  # 不合格
    clock.advance(seconds=61)
    assert run(session_factory, clock)["waiting"] == 1
    assert gas.calls == []
    ping_ok(db, at=clock.at)
    clock.advance(seconds=61)
    assert run(session_factory, clock)["done"] == 1


# ===== 呼ぶ形と結果の記録 =====


def test_after_approve1_sends_expect_args_and_records(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001", sign_png=PNG)
    version_before = req(db, "LV-20261126-001").row_version
    synced_before = dict(db.get(models.SheetSyncState, ("leave_requests", "LV-20261126-001")).synced_values)
    gas.replies["internal.afterApprove1"] = {"ok": True, "sign": {"status": "saved", "signImageUrl": "https://s/1"},
                                             "mail": {"status": "sent"},
                                             "ledger": {"status": "upserted", "sheetName": "2026年度"}}
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    guard = json.loads(job(db, job_id).guard_json)
    assert run(session_factory, clock)["done"] == 1
    call = gas.calls[0]
    assert call["action"] == "internal.afterApprove1" and call["reqId"] == "LV-20261126-001"
    assert call["jobId"] == guard["gasJobId"]
    assert call["expect"] == {"status": "承認済", "approvedAt": "2026-11-26 11:30:15",
                              "approvedBy1Email": "boss1@lineworks-local.info"}
    assert call["args"] == {"signBase64": "data:image/png;base64," + base64.b64encode(PNG).decode(),
                            "leaveDate": "2026-11-27"}
    row = req(db, "LV-20261126-001")
    assert row.sign_image_url == "https://s/1" and row.sign_png is None
    # G 列は行の反映の対象外: 版も synced_values も変えない
    assert row.row_version == version_before
    db.expire_all()
    assert db.get(models.SheetSyncState, ("leave_requests", "LV-20261126-001")).synced_values == synced_before
    stored = job(db, job_id)
    assert json.loads(stored.result_json)["sign"]["signImageUrl"] == "https://s/1"
    assert "base64" not in stored.result_json and "base64" not in stored.args_json


def test_electronic_stamp_sends_empty_sign(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001", sign_image_url="https://s/old")
    gas.replies["internal.afterApprove1"] = {"ok": True, "sign": {"status": "cleared", "signImageUrl": ""}}
    gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    run(session_factory, clock)
    assert gas.calls[0]["args"]["signBase64"] == ""
    assert req(db, "LV-20261126-001").sign_image_url == ""


def test_after_approve2_records_pdf(session_factory, db, clock, gas, monkeypatch):
    ping_ok(db)
    add_request(db, "LV-20261126-002", approved_at2=APPROVED_AT2, approved_by2_email="boss2@lineworks-local.info",
                pdf_url="https://drive.google.com/file/d/OLD1/view", pdf_file_id="OLD1")
    gas.replies["internal.afterApprove2"] = {"ok": True, "pdf": {"status": "created", "pdfFileId": "NEW1",
                                                                 "pdfUrl": "https://drive.google.com/file/d/NEW1/view"},
                                             "trash": {"status": "skipped"}, "ledger": {"status": "upserted"}}
    gas_jobs.enqueue_after_approve2(db, "LV-20261126-002")
    db.commit()
    run(session_factory, clock)
    call = gas.calls[0]
    assert call["expect"] == {"status": "承認済", "approvedAt2": "2026-11-26 11:45:00",
                              "approvedBy2Email": "boss2@lineworks-local.info"}
    assert call["args"] == {"approver1Email": "boss1@lineworks-local.info",
                            "approver2Email": "boss2@lineworks-local.info",
                            "oldPdfUrl": "https://drive.google.com/file/d/OLD1/view", "oldPdfFileId": "OLD1",
                            "cleanupMode": "asis"}
    row = req(db, "LV-20261126-002")
    assert row.pdf_url == "https://drive.google.com/file/d/NEW1/view" and row.pdf_file_id == "NEW1"


def test_cleanup_mode_from_quirks(db, clock, monkeypatch):
    fake_quirks = types.ModuleType("app.domain.quirks")
    fake_quirks.PDF_CLEANUP_MODE = "trash"
    monkeypatch.setitem(sys.modules, "app.domain.quirks", fake_quirks)
    import app.domain as domain_pkg

    monkeypatch.setattr(domain_pkg, "quirks", fake_quirks, raising=False)
    row = add_request(db, "LV-20261126-002", approved_at2=APPROVED_AT2)
    assert gas_jobs.build_call("internal.afterApprove2", row, {})[1]["cleanupMode"] == "trash"


def test_record_already_takes_id_from_url(db, clock):
    add_request(db, "LV-20261126-002")
    gas_jobs.record_gas_result(db, "LV-20261126-002", {
        "ok": True, "action": "internal.afterApprove2",
        "pdf": {"status": "already", "pdfUrl": "https://drive.google.com/file/d/ABC_def-1/view?usp=x"}})
    db.commit()
    row = req(db, "LV-20261126-002")
    assert row.pdf_url.endswith("ABC_def-1/view?usp=x") and row.pdf_file_id == "ABC_def-1"


def test_record_pdf_error_keeps_old(db, clock):
    add_request(db, "LV-20261126-002", pdf_url="u-old", pdf_file_id="f-old")
    gas_jobs.record_gas_result(db, "LV-20261126-002", {"ok": True, "pdf": {"status": "error", "error": "x"}})
    db.commit()
    row = req(db, "LV-20261126-002")
    assert (row.pdf_url, row.pdf_file_id) == ("u-old", "f-old")


def test_after_unapprove_args_and_trash(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-003", status="取消", sign_image_url="https://s/1",
                pdf_url="https://drive.google.com/file/d/P1/view", pdf_file_id="P1")
    gas.replies["internal.afterUnapprove"] = {"ok": True, "trash": {"status": "trashed"},
                                              "ledger": {"status": "removed"}}
    gas_jobs.enqueue_after_unapprove(db, "LV-20261126-003", reason="cancel", old_leave_date="2026-11-27",
                                     old_pdf_url="https://drive.google.com/file/d/P1/view", old_pdf_file_id="P1",
                                     clear_sign=False)
    db.commit()
    run(session_factory, clock)
    call = gas.calls[0]
    assert call["expect"] == {"status": "取消"}
    assert call["args"] == {"reason": "cancel", "oldLeaveDate": "2026-11-27",
                            "clearSign": False, "oldPdfUrl": "https://drive.google.com/file/d/P1/view",
                            "oldPdfFileId": "P1", "cleanupMode": "asis"}
    row = req(db, "LV-20261126-003")
    assert row.pdf_url == "" and row.pdf_file_id == "" and row.sign_image_url == "https://s/1"


def test_after_unapprove_clear_sign(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-003", status="申請中", approved_at=None, sign_image_url="https://s/1",
                pdf_url="u", pdf_file_id="f")
    gas_jobs.enqueue_after_unapprove(db, "LV-20261126-003", reason="edit", old_leave_date="2026-11-27",
                                     old_pdf_url="u", old_pdf_file_id="f", clear_sign=True)
    db.commit()
    run(session_factory, clock)
    row = req(db, "LV-20261126-003")
    assert row.sign_image_url == "" and row.pdf_url == "u"  # asis: PDF の列は残す


def test_send_warn_mail_writes_summary_first(session_factory, db, clock, gas, monkeypatch):
    ping_ok(db)
    order: list[str] = []
    fake = types.ModuleType("app.services.summary_sheet")
    fake.write = lambda db: order.append("summary") or {"ok": True, "rows": 3}
    monkeypatch.setitem(sys.modules, "app.services.summary_sheet", fake)
    import app.services as services_pkg

    monkeypatch.setattr(services_pkg, "summary_sheet", fake, raising=False)
    gas.replies["internal.sendWarnMail"] = lambda action, args: order.append("gas") or {
        "ok": True, "mail": {"status": "not_warn_day", "count": 0}}
    job_id = gas_jobs.enqueue_send_warn_mail(db)
    db.commit()
    assert run(session_factory, clock)["done"] == 1
    assert order == ["summary", "gas"]
    call = gas.calls[0]
    assert call["reqId"] == "" and call["expect"] is None and call["args"] == {"fiscalYear": "2026"}
    assert json.loads(job(db, job_id).result_json)["mail"]["status"] == "not_warn_day"


def test_send_warn_mail_summary_failure_retries(session_factory, db, clock, gas, monkeypatch):
    ping_ok(db)
    fake = types.ModuleType("app.services.summary_sheet")
    fake.write = lambda db: {"ok": False, "error": "gate"}
    monkeypatch.setitem(sys.modules, "app.services.summary_sheet", fake)
    import app.services as services_pkg

    monkeypatch.setattr(services_pkg, "summary_sheet", fake, raising=False)
    job_id = gas_jobs.enqueue_send_warn_mail(db)
    db.commit()
    assert run(session_factory, clock)["retry"] == 1
    assert gas.calls == [] and job(db, job_id).state == "pending"


# ===== 再試行 =====


def test_retry_uses_same_job_id_on_timeout(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001")
    gas.replies["internal.afterApprove1"] = [GasTransportError("timeout", timeout=True),
                                             GasTransportError("network"), {"ok": True, "replayed": True}]
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    assert run(session_factory, clock)["retry"] == 1
    assert job(db, job_id).run_after == T0 + timedelta(minutes=1)
    clock.advance(minutes=1)
    ping_ok(db, at=clock.at)
    assert run(session_factory, clock)["retry"] == 1
    assert job(db, job_id).run_after == clock.at + timedelta(minutes=2)
    clock.advance(minutes=2)
    ping_ok(db, at=clock.at)
    assert run(session_factory, clock)["done"] == 1
    assert len({c["jobId"] for c in gas.calls}) == 1 and len(gas.calls) == 3


def test_retryable_false_fails_and_records_partial(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001", sign_png=PNG)
    gas.replies["internal.afterApprove1"] = {"ok": False, "error": "ledger_header_mismatch", "retryable": False,
                                             "sign": {"status": "saved", "signImageUrl": "https://s/9"},
                                             "mail": {"status": "sent"}}
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    assert run(session_factory, clock)["failed"] == 1
    row = job(db, job_id)
    assert row.state == "failed" and "ledger_header_mismatch" in row.last_error
    assert json.loads(row.result_json)["error"] == "ledger_header_mismatch"
    r = req(db, "LV-20261126-001")
    assert r.sign_image_url == "https://s/9" and r.sign_png is None  # 済んだ段は先に記録
    exhausted = db.query(models.SyncRun).filter_by(kind="gas_job_failed").all()
    assert len(exhausted) == 1 and exhausted[0].summary_json["requestId"] == "LV-20261126-001"
    # 運用の欄から再実行できる
    assert jobs.retry_failed(db, job_id) is True


def test_retryable_true_retries(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001")
    gas.replies["internal.afterApprove1"] = {"ok": False, "error": "busy", "retryable": True}
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    assert run(session_factory, clock)["retry"] == 1
    assert job(db, job_id).state == "pending" and "busy" in job(db, job_id).last_error


def test_auth_error_is_not_retried_and_pauses_lane(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001")
    add_request(db, "LV-20261126-002", worker_id="2002")
    gas.replies["internal.afterApprove1"] = [GasAuthError("login page"), {"ok": True}]
    first = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    second = gas_jobs.enqueue_after_approve1(db, "LV-20261126-002")
    db.commit()
    assert run(session_factory, clock)["failed"] == 1
    assert job(db, first).state == "failed" and "auth" in job(db, first).last_error
    assert len(gas.calls) == 1
    # レーンは次の合格の gas_ping まで止まる
    assert run(session_factory, clock)["waiting"] == 1
    assert len(gas.calls) == 1 and job(db, second).state == "pending"
    ok, reason = gas_jobs.gas_ping_ok(db, clock.at)
    assert not ok and "入口に入れません" in reason


def test_retry_gives_up_after_24_hours(session_factory, db, clock, gas):
    add_request(db, "LV-20261126-001")
    gas.replies["internal.afterApprove1"] = GasTransportError("network")
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    results = []
    for _ in range(60):
        ping_ok(db, at=clock.at)
        out = run(session_factory, clock)
        results.append(out)
        if out["failed"]:
            break
        row = job(db, job_id)
        clock.at = row.run_after
    assert job(db, job_id).state == "failed"
    assert clock.at - T0 <= timedelta(hours=24)
    assert len(gas.calls) >= 10


def test_same_request_runs_in_order_and_failed_blocks(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001", approved_at2=APPROVED_AT2)
    add_request(db, "LV-20261126-009", worker_id="3003")
    gas.replies["internal.afterApprove1"] = {"ok": False, "error": "bad_request", "retryable": False}
    first = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    second = gas_jobs.enqueue_after_approve2(db, "LV-20261126-001")
    other = gas_jobs.enqueue_after_approve2(db, "LV-20261126-009")
    db.commit()
    run(session_factory, clock)
    assert job(db, first).state == "failed"
    run(session_factory, clock)
    run(session_factory, clock)
    assert job(db, second).state == "pending"  # 前が failed なので止まる
    assert job(db, other).state == "done"  # 別の申請は進む
    assert [c["reqId"] for c in gas.calls] == ["LV-20261126-001", "LV-20261126-009"]
    assert [j["id"] for j in jobs.list_blocked(db)] == [second]


def test_missing_request_is_skipped(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-001")
    job_id = gas_jobs.enqueue_after_approve1(db, "LV-20261126-001")
    db.commit()
    db.query(models.LeaveRequest).delete()
    db.commit()
    run(session_factory, clock)
    assert job(db, job_id).state == "skipped"


# ===== gas_ping =====


def test_gas_ping_ok_is_stored(session_factory, db, clock, gas, settings):
    gas.replies["internal.ping"] = dict(STAGING_PING, ledgerKeyRows=2)
    out = gas_jobs.run_gas_ping(session_factory, settings)
    assert out["ok"] is True and out["problems"] == []
    assert out["warnings"] and "LEDGER_SSID" in out["warnings"][0]
    row = gas_jobs.last_gas_ping(db)
    assert row.ok is True and row.summary_json["migrationMode"] is False and row.summary_json["migrated"] is False
    assert gas.calls[0]["action"] == "internal.ping" and gas.calls[0]["jobId"].startswith("ping-")
    assert gas_jobs.gas_ping_ok(db, clock.at) == (True, "")


@pytest.mark.parametrize("change", [{"scriptId": "15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i"},
                                    {"spreadsheetId": "other"}, {"spreadsheetTimeZone": "Asia/Tokyo"}])
def test_gas_ping_mismatch_stops_lane(session_factory, db, clock, gas, settings, change):
    gas.replies["internal.ping"] = dict(STAGING_PING, **change)
    out = gas_jobs.run_gas_ping(session_factory, settings)
    assert out["ok"] is False and out["problems"]
    assert gas_jobs.gas_ping_ok(db, clock.at)[0] is False


def test_gas_ping_transport_error_is_recorded(session_factory, db, clock, gas, settings):
    gas.replies["internal.ping"] = GasTransportError("network")
    out = gas_jobs.run_gas_ping(session_factory, settings)
    assert out["ok"] is False and "GAS に届きませんでした" in out["problems"][0]
    assert gas_jobs.last_gas_ping(db).ok is False


@pytest.mark.real_write_gate
def test_gas_ping_failure_closes_write_gate(session_factory, db, clock, gas, settings):
    gas.replies["internal.ping"] = dict(STAGING_PING, spreadsheetId="other")
    gas_jobs.run_gas_ping(session_factory, settings)
    allowed, reason = write_gate.check(db, settings, clock.at)
    assert allowed is False and "gas_ping" in reason
    gas.replies["internal.ping"] = dict(STAGING_PING)
    gas_jobs.run_gas_ping(session_factory, settings)
    assert write_gate.check(db, settings, clock.at) == (True, "")
    # 合格の後に照合が落ち続けると、10 分で関所が閉じる
    gas.replies["internal.ping"] = dict(STAGING_PING, spreadsheetId="other")
    clock.advance(minutes=5)
    gas_jobs.run_gas_ping(session_factory, settings)
    clock.advance(minutes=6)
    assert write_gate.check(db, settings, clock.at)[0] is False


@pytest.mark.real_write_gate
def test_production_gate_follows_migrated(session_factory, db, clock, gas, settings):
    prod = SimpleNamespace(is_staging=False, sheet_write_ping_max_minutes=10)
    gas.replies["internal.ping"] = dict(STAGING_PING, migrated=False)
    gas_jobs.run_gas_ping(session_factory, settings)
    assert write_gate.check(db, prod, clock.at)[0] is False
    gas.replies["internal.ping"] = dict(STAGING_PING, migrated=True)
    gas_jobs.run_gas_ping(session_factory, settings)
    assert write_gate.check(db, prod, clock.at) == (True, "")


# ===== 本物の GasClient を通した本文の形（G の src/Internal.gs と同じキー） =====


def test_wire_body_matches_internal_gs(session_factory, db, clock, settings):
    import re

    import httpx

    from app.services.gas_client import GasClient

    sent: list[dict] = []

    def handler(request: httpx.Request) -> httpx.Response:
        if request.method == "POST":
            sent.append(json.loads(request.content))
            return httpx.Response(302, headers={"location": "https://script.googleusercontent.com/macros/echo?x=1"})
        body = sent[-1]
        reply = {"ok": True, "action": body["action"], "jobId": body["jobId"], "reqId": body["reqId"],
                 "sign": {"status": "cleared", "signImageUrl": ""}, "trash": {"status": "skipped"},
                 "ledger": {"status": "removed", "sheetName": "2026年度"}, "reason": "edit"}
        return httpx.Response(200, headers={"content-type": "application/json"}, content=json.dumps(reply).encode())

    client = GasClient(settings.gas_internal_url, settings.gas_internal_token, settings=settings,
                       auth_mode="anonymous",
                       http_factory=lambda **kw: httpx.Client(transport=httpx.MockTransport(handler), **kw))
    previous = gas_jobs.set_client_factory(lambda s: client)
    try:
        gas_jobs.register_worker()
        ping_ok(db)
        add_request(db, "LV-20261126-003", status="申請中", approved_at=None, sign_image_url="https://s/1")
        gas_jobs.enqueue_after_unapprove(db, "LV-20261126-003", reason="edit", old_leave_date="2026-11-27",
                                         old_pdf_url="", old_pdf_file_id="", clear_sign=True)
        db.commit()
        assert run(session_factory, clock)["done"] == 1
    finally:
        gas_jobs.set_client_factory(previous)
    body = sent[0]
    assert body["internal"] is True and body["action"] == "internal.afterUnapprove"
    assert body["reqId"] == "LV-20261126-003" and "requestId" not in body
    assert re.fullmatch(r"[A-Za-z0-9._:-]{1,100}", body["jobId"])
    assert body["expect"] == {"status": "申請中"}
    assert set(body["args"]) == {"reason", "oldLeaveDate", "clearSign", "oldPdfUrl", "oldPdfFileId", "cleanupMode"}
    assert req(db, "LV-20261126-003").sign_image_url == ""


def test_pdf_failure_from_gas_is_retried_and_ledger_kept(session_factory, db, clock, gas):
    ping_ok(db)
    add_request(db, "LV-20261126-002", approved_at2=APPROVED_AT2, pdf_url="u-old", pdf_file_id="f-old")
    gas.replies["internal.afterApprove2"] = {"ok": False, "error": "internal", "message": "pdf: boom",
                                             "retryable": True, "failedStage": "pdf",
                                             "pdf": {"status": "error", "error": "boom"},
                                             "trash": {"status": "skipped"}, "ledger": {"status": "upserted"}}
    job_id = gas_jobs.enqueue_after_approve2(db, "LV-20261126-002")
    db.commit()
    assert run(session_factory, clock)["retry"] == 1
    assert json.loads(job(db, job_id).result_json)["failedStage"] == "pdf"
    assert req(db, "LV-20261126-002").pdf_url == "u-old"


# ===== 登録 =====


def test_register_worker(gas):
    assert set(jobs.registered_kinds("gas")) == set(gas_jobs.KINDS)
    sched = worker.schedules()["gas_ping"]
    assert sched.every_minutes == 5 and sched.run_at_startup is True and sched.task is gas_jobs.run_gas_ping


def test_worker_loads_registration():
    loaded = worker.load_registrations()
    assert "app.services.gas_jobs" in loaded
    assert "gas_ping" in worker.schedules()
