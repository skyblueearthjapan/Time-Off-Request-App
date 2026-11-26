"""運用の欄の中身（services/ops_status.get_ops_status。担当 M）。"""
from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone

import pytest

from app import jobs, models, sheet_queue
from app.services import ops_status

T0 = datetime(2026, 11, 26, 3, 0, tzinfo=timezone.utc)  # JST 12:00


@pytest.fixture
def frozen(monkeypatch):
    from app import timeutil

    monkeypatch.setattr(timeutil, "now_utc", lambda: T0)


def _req(db, req_id, synced=True, first_dirty_at=None):
    db.add(models.LeaveRequest(req_id=req_id, worker_id="1001", status="承認済", sign_png=b"secret-png"))
    db.flush()
    sheet_queue.mark_dirty(db, "leave_requests", req_id)
    state = db.get(models.SheetSyncState, ("leave_requests", req_id))
    if synced:
        state.synced_version = state.desired_version
    if first_dirty_at is not None:
        state.first_dirty_at = first_dirty_at


def test_empty_status(db, frozen):
    out = ops_status.get_ops_status(db)
    assert out["now"] == "2026-11-26 12:00:00"
    assert out["rowSync"]["pending"] == 0 and out["rowSync"]["blocked"] == []
    assert out["gasJobs"]["counts"]["failed"] == 0 and out["gasJobs"]["failed"] == []
    assert out["gasPing"]["lanePaused"] is True and out["gasPing"]["fresh"] is False
    assert out["sheetWriteGate"]["blocked"] is True and "gas_ping" in out["sheetWriteGate"]["reason"]
    assert set(out["lastRuns"]) == set(ops_status.LAST_RUN_KINDS) and out["exhausted"] == []
    json.dumps(out)  # JSON にできる


def test_status_shows_jobs_ping_rows_and_runs(db, frozen):
    _req(db, "LV-1", synced=False, first_dirty_at=T0 - timedelta(minutes=7))
    _req(db, "LV-2")
    blocked_state = db.get(models.SheetSyncState, ("leave_requests", "LV-2"))
    blocked_state.state, blocked_state.last_error = "blocked", "REQ_ID が見つかりません"
    failed = jobs.enqueue(db, "internal.afterApprove1", {}, request_id="LV-3")
    after = jobs.enqueue(db, "internal.afterApprove2", {}, request_id="LV-3")
    waiting = jobs.enqueue(db, "internal.afterApprove1", {}, request_id="LV-4")
    db.flush()
    row = db.get(models.Job, failed)
    row.state, row.last_error, row.attempts = "failed", "PermanentJobError: ledger_header_mismatch", 1
    db.get(models.Job, waiting).last_error = "row_not_synced（行の反映待ち）: LV-4"
    db.add(models.SyncRun(kind="gas_ping", started_at=T0 - timedelta(minutes=2), finished_at=T0 - timedelta(minutes=2),
                          ok=False, summary_json={"problems": ["scriptId が本番の GAS です"], "warnings": ["w"],
                                                  "migrated": False}))
    db.add(models.SyncRun(kind="gas_job_failed", started_at=T0, finished_at=T0, ok=False,
                          summary_json={"jobId": failed, "kind": "internal.afterApprove1", "requestId": "LV-3",
                                        "error": "x"}))
    db.add(models.SyncRun(kind="schedule:sync_masters", started_at=T0 - timedelta(hours=6), finished_at=T0, ok=True,
                          summary_json={"ok": True, "rows": 50}))
    db.commit()

    out = ops_status.get_ops_status(db)
    assert out["rowSync"]["pending"] == 1 and out["rowSync"]["delayed"] == 1
    assert out["rowSync"]["oldestPendingSeconds"] == 420
    assert out["rowSync"]["blocked"][0]["pk"] == "LV-2"
    g = out["gasJobs"]
    assert g["counts"]["failed"] == 1 and g["counts"]["pending"] == 2
    assert g["failed"][0]["id"] == failed and g["failed"][0]["action"] == "internal.afterApprove1"
    assert "ledger_header_mismatch" in g["failed"][0]["reason"]
    assert [b["id"] for b in g["blocked"]] == [after]
    assert [w["id"] for w in g["waiting"]] == [waiting]
    p = out["gasPing"]
    assert p["ok"] is False and p["lanePaused"] is True and p["problems"] == ["scriptId が本番の GAS です"]
    assert p["at"] == "2026-11-26 11:58:00" and p["warnings"] == ["w"]
    assert out["exhausted"][0]["requestId"] == "LV-3"
    assert out["lastRuns"]["sync_masters"]["ok"] is True
    assert out["lastRuns"]["sync_masters"]["summary"]["rows"] == 50
    assert out["lastRuns"]["sync_masters"]["finishedAt"] == "2026-11-26 12:00:00"
    assert out["sheetWriteGate"]["allowed"] is False  # 合格の ping が無い
    text = json.dumps(out, ensure_ascii=False)
    assert "secret-png" not in text and "token" not in text.lower()


def test_fresh_ping_unpauses_lane(db, frozen):
    db.add(models.SyncRun(kind="gas_ping", started_at=T0 - timedelta(minutes=1), finished_at=T0 - timedelta(minutes=1),
                          ok=True, summary_json={"problems": [], "migrated": True, "auth": "bearer", "apiVersion": 1}))
    db.commit()
    p = ops_status.get_ops_status(db)["gasPing"]
    assert p["ok"] is True and p["fresh"] is True and p["lanePaused"] is False
    assert p["migrated"] is True and p["auth"] == "bearer" and p["apiVersion"] == 1
    assert ops_status.get_ops_status(db)["sheetWriteGate"]["allowed"] is True
