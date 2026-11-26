"""leave_requests の行の読み書きの小道具（GAS の getSheetHeaderIndex_ で列を引いて getValues() で読む所の代わり）。

- シートの並び（GAS が上から読む順）は sheet_row の順。まだシートに書いていない新しい行（sheet_row が NULL）は
  末尾に、申請日時・REQ_ID の順で並べる。
- セルの型（仕様 4-1）: 日付・日時の列は、値があり cell_kinds が 'string' でなければ Date のセル。'string' なら文字の
  セルで、元の文字は *_text の列。GAS の `x instanceof Date ? 書式 : String(x || '')` は date_cell／text_cell で書く。
- 文字の列へ書くときは GAS の setValues と同じく、数に見える文字は数のセルになる（DB には読み直した値と型を入れる）。
- 書いたら mark（sheet_queue.mark_dirty）。commit はしない。
"""
from __future__ import annotations

from datetime import date, datetime
from typing import Any

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app import sheet_queue, store as base_store, timeutil
from app.jsutil import js_string, normalize_gas, sheet_readback
from app.models import LeaveRequest

from .jsdate import js_new_date

TABLE = "leave_requests"


def norm(row: LeaveRequest, name: str) -> str:
    """normalize_(row[列])（文字の列）。"""
    return normalize_gas(getattr(row, name))


def kind_of(row: LeaveRequest, name: str) -> str | None:
    return (row.cell_kinds or {}).get(name)


def set_kind(row: LeaveRequest, name: str, kind: str | None) -> None:
    kinds = dict(row.cell_kinds or {})
    if kind is None:
        kinds.pop(name, None)
    else:
        kinds[name] = kind
    row.cell_kinds = kinds


def is_date_cell(row: LeaveRequest, name: str) -> bool:
    return getattr(row, name) is not None and kind_of(row, name) != "string"


def date_cell(row: LeaveRequest, name: str):
    """Date のセルならその値（日付の列は date、日時の列は UTC の datetime）、ほかは None。"""
    return getattr(row, name) if is_date_cell(row, name) else None


def text_cell(row: LeaveRequest, name: str) -> str:
    """Date でないセルの元の文字（空のセルは ''）。"""
    if is_date_cell(row, name):
        return ""
    return js_string(getattr(row, f"{name}_text") or "")


def raw_cell(row: LeaveRequest, name: str) -> Any:
    """GAS の row[列]: Date のセルは date／datetime、ほかは文字（空は ''）。"""
    value = date_cell(row, name)
    return value if value is not None else text_cell(row, name)


def cell_moment(value: Any) -> datetime | None:
    """`x instanceof Date ? x : new Date(x)` の瞬間（空・読めなければ None）。日付は日本時間の 0 時。"""
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return timeutil.jst_midnight(value)
    if not value:
        return None
    return js_new_date(value)


def fmt(value: date | datetime, pattern: str = "yyyy-MM-dd") -> str:
    """fmtDate_（日本時間）。"""
    return timeutil.to_jst_text(value, pattern)


def fiscal_year_text(row: LeaveRequest) -> str:
    """年度のセルを String(値) にした文字（数の 2026 は '2026'）。"""
    return js_string(row.fiscal_year or "")


def fiscal_year_value(row: LeaveRequest) -> Any:
    """年度のセルの値と型（数のセルなら数、文字のセルなら文字、空は ''）。"""
    t = fiscal_year_text(row)
    if t != "" and kind_of(row, "fiscal_year") != "string":
        try:
            n = float(t)
        except ValueError:
            return t
        return int(n) if n.is_integer() else n
    return t


# ===== 書く =====


def set_text(row: LeaveRequest, name: str, value: Any) -> None:
    """文字の列へ書く（setValue／setValues と同じく、数に見える文字は数のセル）。"""
    text = js_string(value)
    if text == "":
        setattr(row, name, "")
        set_kind(row, name, None)
        return
    stored, kind = sheet_readback(text)
    setattr(row, name, stored)
    set_kind(row, name, kind if kind == "number" else None)


def set_number(row: LeaveRequest, name: str, value: int) -> None:
    setattr(row, name, js_string(value))
    set_kind(row, name, "number")


def set_date(row: LeaveRequest, name: str, value: date | datetime | None) -> None:
    """日付・日時の列へ Date を書く（None は空にする）。"""
    setattr(row, name, value)
    setattr(row, f"{name}_text", "")
    set_kind(row, name, "serial" if value is not None else None)


def seconds(moment: datetime) -> datetime:
    """新しく入れる日時（申請・1 次・2 次の承認）は秒で切り捨てる。

    GAS 側の照合（guard の approvedAt など）は秒まで比べる。ミリ秒が残ると、シートの値と DB の値の
    ミリ秒のずれで row_not_synced を繰り返す。シートへ書くシリアル値もこの値から作るので秒の精度になる。
    既にある行（移行した古い行）の値は変えない（ここを通るのは新しく入れるときだけ）。
    """
    return moment.replace(microsecond=0)


def mark(db: Session, req_id: str) -> int:
    db.flush()
    return sheet_queue.mark_dirty(db, TABLE, req_id)


# ===== 読む =====


def _order():
    return (LeaveRequest.sheet_row.is_(None), LeaveRequest.sheet_row, LeaveRequest.applied_at, LeaveRequest.req_id)


def rows_in_sheet_order(db: Session) -> list[LeaveRequest]:
    return list(db.scalars(select(LeaveRequest).order_by(*_order())))


def row_count(db: Session) -> int:
    return int(db.scalar(select(func.count()).select_from(LeaveRequest)) or 0)


def find(db: Session, req_id: Any) -> LeaveRequest | None:
    """normalize_(行の REQ_ID) === 引数（引数は正規化しない。DB の req_id は正規化済み）。"""
    if not isinstance(req_id, str) or not req_id:
        return None
    return db.get(LeaveRequest, req_id)


def lock(db: Session, req_id: Any) -> LeaveRequest | None:
    """行ロック（SELECT … FOR UPDATE）して読み直す（無ければ None）。"""
    if not isinstance(req_id, str) or not req_id:
        return None
    return base_store.lock_request(db, req_id)
