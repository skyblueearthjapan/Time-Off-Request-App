"""store.lock_request（仕様書 5-4・10-2）: 申請 1 件の行を FOR UPDATE で読み直す。"""
from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.dialects import postgresql

from app import store
from app.models import LeaveRequest


def test_lock_request_returns_row_or_none(db):
    db.add(LeaveRequest(req_id="LV-20261126-001", status="申請中"))
    db.commit()
    row = store.lock_request(db, "LV-20261126-001")
    assert row is not None and row.status == "申請中"
    assert store.lock_request(db, "LV-20261126-999") is None
    assert store.lock_request(db, "") is None
    assert store.lock_request(db, None) is None  # type: ignore[arg-type]
    assert store.lock_request(db, " LV-20261126-001") is None  # 引数は正規化しない（完全一致）


def test_lock_request_reads_fresh_values(db, session_factory):
    db.add(LeaveRequest(req_id="LV-20261126-002", status="申請中"))
    db.commit()
    assert db.get(LeaveRequest, "LV-20261126-002").status == "申請中"
    with session_factory() as other:
        other.get(LeaveRequest, "LV-20261126-002").status = "承認済"
        other.commit()
    assert store.lock_request(db, "LV-20261126-002").status == "承認済"  # 手元の古い値でなく読み直す


def test_lock_request_uses_for_update(db, monkeypatch):
    seen = []
    original = db.scalars

    def spy(stmt, *a, **k):
        seen.append(str(stmt.compile(dialect=postgresql.dialect())))
        return original(stmt, *a, **k)

    monkeypatch.setattr(db, "scalars", spy)
    store.lock_request(db, "LV-x")
    assert seen and "FOR UPDATE" in seen[0]
    assert select(LeaveRequest) is not None
