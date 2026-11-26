"""申請の行ロック（仕様書 5-4。GAS のスクリプトのロックの代わり）。

書き込みの業務関数（担当 W）は、申請 1 件ごとに lock_request(db, req_id) で leave_requests の行を
SELECT … FOR UPDATE してから状態を読み直して決める。別の申請は待たない。SQLite では FOR UPDATE は無視される。
行の探し方は「DB の req_id と引数の完全一致」（引数を正規化しない＝GAS の leaveRequest.gs と同じ）。
"""
from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.orm import Session

from app.models import LeaveRequest


def lock_request(db: Session, req_id: str) -> LeaveRequest | None:
    """その申請の行をロックして読み直す（無ければ None）。ロックはトランザクションの終わりまで続く。"""
    if not isinstance(req_id, str) or not req_id:
        return None
    stmt = (select(LeaveRequest).where(LeaveRequest.req_id == req_id)
            .with_for_update().execution_options(populate_existing=True))
    return db.scalars(stmt).first()
