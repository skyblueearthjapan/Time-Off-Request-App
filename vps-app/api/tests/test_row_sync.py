"""services/row_sync: 差分の列だけ書く・G 列と N 列を書かない・新しい行は最後の行の次（REQ_ID が空の半端な行を上書きしない）・
古い行の型を保つ・統合業務メールと同じ条件で読める（休暇日が日付のセル）・REQ_ID の確かめと探し直し・見出しの違い・
時間帯の違い・書き込みの関所・行やタブを消す要求を出さない。"""
from __future__ import annotations

from datetime import date, datetime, timezone

import pytest

from app import sheet_queue
from app.models import LeaveRequest, SheetSyncState, SyncRun
from app.services import row_sync, write_gate
from app.services import sheet_codec as codec
from conftest import freeze_now
from d_sheet_fixtures import HEADERS, NOW, R_HALF, R_OLD, R_SUB, SIGN_URL, STAGING_SID, make_fake, migrate


@pytest.fixture(autouse=True)
def _lane():
    row_sync.reset_lane()
    yield
    row_sync.reset_lane()


@pytest.fixture
def env(engine, session_factory, settings, monkeypatch):
    freeze_now(monkeypatch, NOW)
    fake = make_fake()
    migrate(fake, engine, settings, app_env="production")
    fake.calls.clear()
    book = fake.books[STAGING_SID]

    def run(**kw):
        return row_sync.run_once(session_factory, api=fake.api(settings), app_settings=settings, **kw)

    return type("Env", (), {"fake": fake, "book": book, "run": staticmethod(run), "sf": session_factory})


def cell(book, row, header):
    return book.tabs["T_LEAVE_REQUEST"].get(row - 1, HEADERS.index(header))


def touched(fake):
    out = []
    for req in fake.batch_requests():
        body = req.get("updateCells")
        if body:
            out.append((body["start"]["rowIndex"] + 1, HEADERS[body["start"]["columnIndex"]]))
    return out


def assert_only_allowed_requests(fake):
    for req in fake.batch_requests():
        (kind,) = req.keys()
        assert kind in ("updateCells", "appendDimension")
    for _m, url, _b in fake.writes():
        assert url.endswith(":batchUpdate") and STAGING_SID in url and "values" not in url  # values.append を使わない


def new_request(db, req_id, *, leave=date(2026, 12, 10)):
    db.add(LeaveRequest(req_id=req_id, fiscal_year="2026", status="申請中", applied_at=NOW, dept_id="組立",
                        dept_name="組立", worker_id="1002", worker_name="試験 二郎", leave_date=leave,
                        leave_kubun="全日休", leave_type="有給消化（私用）", created_by_email="shared@lineworks-local.info",
                        cell_kinds={"fiscal_year": "number", "worker_id": "number", "leave_date": "serial",
                                    "applied_at": "serial"}))
    db.flush()
    sheet_queue.mark_dirty(db, "leave_requests", req_id)


def test_nothing_dirty_reads_only(env):
    assert env.run()["written"] == 0
    assert all(m == "GET" for m, *_ in env.fake.calls)
    env.fake.calls.clear()
    assert env.run()["written"] == 0 and env.fake.calls == []  # 時間帯は 1 時間ごと


def test_approve_writes_only_changed_v_columns_and_not_g_columns(env, db):
    row = db.get(LeaveRequest, R_SUB)
    row.status = "承認済"
    row.approved_at = datetime(2026, 11, 26, 2, 0, tzinfo=timezone.utc)
    row.approved_by1_email = "boss@lineworks-local.info"
    row.sign_image_url = "https://drive.google.com/file/d/NEW/view"  # G 列（DB の値）があってもシートには書かない
    sheet_queue.mark_dirty(db, "leave_requests", R_SUB)
    db.commit()
    result = env.run()
    assert result["ok"] and result["written"] == 1
    assert sorted(h for _, h in touched(env.fake)) == ["APPROVED_AT", "APPROVED_BY1_EMAIL", "STATUS"]
    assert cell(env.book, 5, "STATUS").value == "承認済"
    approved = cell(env.book, 5, "APPROVED_AT")
    assert isinstance(approved.value, float) and approved.formatted() == "2026-11-26 02:00"  # 列の表示形式を写す
    assert cell(env.book, 5, "SIGN_IMAGE_URL").value in (None, "")
    st = db.get(SheetSyncState, ("leave_requests", R_SUB))
    db.refresh(st)
    assert st.synced_version == st.desired_version == 2 and st.state == "ok"
    assert_only_allowed_requests(env.fake)
    env.fake.calls.clear()
    assert env.run()["written"] == 0 and env.fake.writes() == []


def test_unapprove_does_not_touch_gas_columns_or_old_leave_date(env, db):
    before_leave = cell(env.book, 3, "LEAVE_DATE").value
    row = db.get(LeaveRequest, R_HALF)
    row.status = "申請中"
    row.approved_at = None
    row.sign_image_url = ""  # W が DB の G 列を空にしても、シートの G 列は GAS が空にする
    sheet_queue.mark_dirty(db, "leave_requests", R_HALF)
    db.commit()
    assert env.run()["written"] == 1
    assert sorted(h for _, h in touched(env.fake)) == ["APPROVED_AT", "STATUS"]
    assert cell(env.book, 3, "APPROVED_AT").value is None
    assert cell(env.book, 3, "APPROVED_AT").fmt is not None  # 空にしても表示形式は残る
    assert cell(env.book, 3, "SIGN_IMAGE_URL").value == SIGN_URL
    assert cell(env.book, 3, "LEAVE_DATE").value == before_leave


def test_old_row_types_are_kept_and_changed_dates_use_jst_midnight(env, db):
    old_leave = cell(env.book, 2, "LEAVE_DATE").value
    assert old_leave == int(old_leave)  # GMT 0 時＝JST 9 時の古い行
    row = db.get(LeaveRequest, R_OLD)
    row.half_day_type = "午前"
    sheet_queue.mark_dirty(db, "leave_requests", R_OLD)
    db.commit()
    env.run()
    assert cell(env.book, 2, "LEAVE_DATE").value == old_leave  # 変えていない列はそのまま
    assert cell(env.book, 2, "FY_START_YEAR").value == 2025
    env.fake.calls.clear()
    row = db.get(LeaveRequest, R_OLD)
    row.leave_date = date(2026, 3, 23)
    row.fiscal_year = "2026"
    sheet_queue.mark_dirty(db, "leave_requests", R_OLD)
    db.commit()
    env.run()
    assert cell(env.book, 2, "LEAVE_DATE").value == codec.date_to_serial(date(2026, 3, 23), "Etc/GMT")
    assert cell(env.book, 2, "LEAVE_DATE").value % 1 == 0.625
    assert cell(env.book, 2, "FY_START_YEAR").value == 2026


def test_new_rows_go_after_the_last_row_without_overwriting_half_rows(env, db):
    # 最後の REQ_ID の行の下に、REQ_ID が空の半端な行がある（values.append なら上書きしかねない）
    tab = env.book.tabs["T_LEAVE_REQUEST"]
    tab.cell(6, HEADERS.index("WORKER_NAME")).value = "末尾の半端な行"
    tab.row_count = 7
    new_request(db, "LV-20261126-002")
    new_request(db, "LV-20261126-003", leave=date(2026, 12, 11))
    db.commit()
    result = env.run()
    assert result["written"] == 2, result
    assert cell(env.book, 4, "WORKER_NAME").value == "メモ"  # 途中の半端な行はそのまま
    assert cell(env.book, 7, "WORKER_NAME").value == "末尾の半端な行"
    assert cell(env.book, 8, "REQ_ID").value == "LV-20261126-002"
    assert cell(env.book, 9, "REQ_ID").value == "LV-20261126-003"
    # 統合業務メールと同じ条件で読める: 休暇日は日付のセル（数＋日付の表示形式）、状態の文字、2 次承認日時は空
    leave = cell(env.book, 8, "LEAVE_DATE")
    assert isinstance(leave.value, float) and leave.fmt["type"] == "DATE" and leave.value % 1 == 0.625
    assert codec.serial_to_jst_date(leave.value, "Etc/GMT") == date(2026, 12, 10)
    assert cell(env.book, 8, "STATUS").value == "申請中"
    assert cell(env.book, 8, "WORKER_ID").value == 1002 and cell(env.book, 8, "FY_START_YEAR").value == 2026
    assert cell(env.book, 8, "APPROVED_AT2").value is None and cell(env.book, 8, "PDF_URL").value is None
    assert tab.row_count >= 9  # appendDimension で末尾に足した
    assert [db.get(LeaveRequest, k).sheet_row for k in ("LV-20261126-002", "LV-20261126-003")] == [8, 9]
    assert_only_allowed_requests(env.fake)


def test_moved_row_is_found_again_or_blocked(env, db):
    tab = env.book.tabs["T_LEAVE_REQUEST"]
    tab.rows[4], tab.rows[5] = tab.rows[5], tab.rows[4]  # 人が行を並べ替えた（R_SUB が 6 行目へ）
    row = db.get(LeaveRequest, R_SUB)
    row.worker_name = "試験 二郎（改）"
    sheet_queue.mark_dirty(db, "leave_requests", R_SUB)
    db.commit()
    assert env.run()["written"] == 1
    assert cell(env.book, 6, "WORKER_NAME").value == "試験 二郎（改）"
    row = db.get(LeaveRequest, R_SUB)
    db.refresh(row)
    assert row.sheet_row == 6
    # REQ_ID がどこにも無い → 書かずに blocked
    tab.cell(5, 0).value = "LV-99999999-999"
    row.worker_name = "x"
    sheet_queue.mark_dirty(db, "leave_requests", R_SUB)
    db.commit()
    env.fake.calls.clear()
    assert env.run()["blocked"] == 1 and env.fake.writes() == []
    assert db.get(SheetSyncState, ("leave_requests", R_SUB)).state == "blocked"
    assert row_sync.status(db)["blocked"][0]["pk"] == R_SUB


def test_header_mismatch_halts_lane(env, db):
    env.book.tabs["T_LEAVE_REQUEST"].cell(0, HEADERS.index("PDF_URL")).value = "PDF"
    new_request(db, "LV-20261126-002")
    db.commit()
    result = env.run()
    assert result["ok"] is False and "見出し" in result["halted"]
    assert env.fake.writes() == []
    assert db.query(SyncRun).filter_by(kind="row_sync").one().summary_json["halted"] is True


def test_timezone_mismatch_halts_lane(env, db):
    env.book.tz = "Asia/Tokyo"
    new_request(db, "LV-20261126-002")
    db.commit()
    result = env.run()
    assert "時間帯" in result["halted"] and env.fake.writes() == []


def test_write_gate_closed_halts_lane(env, db, monkeypatch):
    monkeypatch.setattr(write_gate, "check", lambda *a, **k: (False, "移転モードではありません"))
    new_request(db, "LV-20261126-002")
    db.commit()
    result = env.run()
    assert "関所" in result["halted"] and env.fake.calls == []


def test_failed_row_retries_later(env, db):
    env.fake.fail_write_rows.add(("T_LEAVE_REQUEST", 5))
    row = db.get(LeaveRequest, R_SUB)
    row.worker_name = "x"
    sheet_queue.mark_dirty(db, "leave_requests", R_SUB)
    db.commit()
    result = env.run()
    assert result["failed"] == 1
    st = db.get(SheetSyncState, ("leave_requests", R_SUB))
    db.refresh(st)
    assert st.state == "retry" and st.attempts == 1 and st.next_try_at is not None
