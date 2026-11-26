"""行の反映の印付け（sheet_queue）と監査（audit）。"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest

from app import audit, sheet_queue, timeutil
from app.models import AuditLog, LeaveRequest, SheetSyncState
from conftest import freeze_now

T0 = datetime(2026, 11, 26, 0, 0, tzinfo=timezone.utc)
T = "leave_requests"


def add(db, req_id, **kw):
    db.add(LeaveRequest(req_id=req_id, **kw))
    db.flush()


def test_mark_dirty_new_rows_get_state_in_creation_order(db, monkeypatch):
    freeze_now(monkeypatch, T0)
    add(db, "LV-20261126-001")
    add(db, "LV-20261126-002")
    assert sheet_queue.mark_dirty(db, T, "LV-20261126-001") == 2
    assert sheet_queue.mark_dirty(db, T, "LV-20261126-002") == 2
    db.commit()
    s1 = db.get(SheetSyncState, (T, "LV-20261126-001"))
    s2 = db.get(SheetSyncState, (T, "LV-20261126-002"))
    assert (s1.desired_version, s1.synced_version, s1.state, s1.first_dirty_at) == (2, 0, "ok", T0)
    assert s2.seq == s1.seq + 1
    assert sheet_queue.current_version(db, T, "LV-20261126-001") == 2


def test_new_row_defaults(db):
    add(db, "LV-1")
    row = db.get(LeaveRequest, "LV-1")
    assert row.status == "申請中" and row.row_version == 1 and row.sheet_row is None and row.sign_png is None
    assert row.pdf_file_id == "" and row.cell_kinds == {}


def test_mark_dirty_again_keeps_first_dirty_until_synced(db, monkeypatch):
    freeze_now(monkeypatch, T0)
    add(db, "R1")
    sheet_queue.mark_dirty(db, T, "R1")
    freeze_now(monkeypatch, T0 + timedelta(seconds=5))
    assert sheet_queue.mark_dirty(db, T, "R1") == 3
    state = db.get(SheetSyncState, (T, "R1"))
    assert state.desired_version == 3 and state.first_dirty_at == T0
    # 反映が済んだ後の書き込みは、新しく「汚れた」時刻を持つ
    state.synced_version = 3
    state.state = "retry"
    freeze_now(monkeypatch, T0 + timedelta(minutes=1))
    sheet_queue.mark_dirty(db, T, "R1")
    assert state.first_dirty_at == T0 + timedelta(minutes=1) and state.state == "ok"
    assert sheet_queue.is_synced(db, T, "R1", 3)
    assert not sheet_queue.is_synced(db, T, "R1", 4)
    assert sheet_queue.requires_for(db, (T, "R1")) == [{"table": T, "pk": "R1", "version": 4}]


def test_blocked_state_stays_blocked(db):
    add(db, "R2")
    sheet_queue.mark_dirty(db, T, "R2")
    state = db.get(SheetSyncState, (T, "R2"))
    state.state = "blocked"
    sheet_queue.mark_dirty(db, T, "R2")
    assert state.state == "blocked"


def test_mark_dirty_rejects_unknown(db):
    with pytest.raises(ValueError):
        sheet_queue.mark_dirty(db, "settings", "ADMIN_EMAILS")
    with pytest.raises(ValueError):
        sheet_queue.mark_dirty(db, "requests", "R1")
    with pytest.raises(ValueError):
        sheet_queue.mark_dirty(db, T, "missing")


def test_audit_record(db):
    audit.record(db, " Tanaka@lineworks-local.info ", "api_approveLeaveRequest", ["R1", "", None], ok=True)
    db.commit()
    row = db.query(AuditLog).one()
    assert row.user_email == "tanaka@lineworks-local.info" and row.request_ids == '["R1"]' and row.ok
    assert row.at.tzinfo is not None and abs(row.at - timeutil.now_utc()) < timedelta(seconds=5)
