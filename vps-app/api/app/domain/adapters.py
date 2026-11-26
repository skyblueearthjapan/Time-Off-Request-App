"""担当 D・M の関数とマスタの表を呼ぶ口（薄い合わせ。ここの 1 か所だけ）。

- マスタ: GAS が getDataRange().getValues() で読む形に直して返す（worker_sheet・calendar_entries・stamp_rows・
  dept_approver_rows・lookup_rows）。
- D・M の関数（実装仕様 10-2）: services.gas_jobs.enqueue_*・summary_sheet.mark_dirty／rebuild_now・
  ops_status.get_ops_status・config_sync.sync_sheet_config・master_sync.sync_masters。
  モジュールがまだ無いブランチでは「まだ用意されていません」の RpcError（試験では偽物を入れる）。
"""
from __future__ import annotations

import importlib
from types import ModuleType
from typing import Any

from sqlalchemy import select, text
from sqlalchemy.orm import Session

from app import models
from app.jsutil import normalize_gas
from app.rpc import RpcError

# ===== マスタの表（GAS が読む形に直す） =====


def worker_sheet(db: Session) -> list[list[Any]]:
    """作業員マスタの getDataRange().getValues()（1 行目が見出し。行が無ければ []）。"""
    rows = list(db.scalars(select(models.Worker).order_by(models.Worker.sort_order)))
    if not rows:
        return []
    headers: list[str] = []
    for r in rows:
        for h in (r.cells or {}):
            if h not in headers:
                headers.append(h)
    return [list(headers)] + [[(r.cells or {}).get(h, "") for h in headers] for r in rows]


def calendar_entries(db: Session) -> list[tuple[Any, str]]:
    """M_CALENDAR のデータ行 [(日付（読めなければ None）, 区分（normalize_ 済み）)]（上から順）。

    タブが無い・2 行目が無い・日付／区分の見出しが無いときは、D の取り込みが表を空にする（仕様 8-3）。
    """
    rows = db.scalars(select(models.CalendarDay).order_by(models.CalendarDay.sort_order))
    return [(r.date, normalize_gas(r.kubun)) for r in rows]


def stamp_rows(db: Session) -> list[tuple[str, str, str]]:
    """M_STAMP のデータ行 [(A メール, B stampFileId, C 備考)]（生の文字）。"""
    rows = db.scalars(select(models.Stamp).order_by(models.Stamp.sort_order))
    return [(r.col_a, r.col_b, r.col_c) for r in rows]


def dept_approver_rows(db: Session) -> list[tuple[str, str, str]]:
    """M_DEPT_APPROVERS のデータ行 [(A dept, B approverEmails, C approverName)]（生の文字）。"""
    rows = db.scalars(select(models.DeptApprover).order_by(models.DeptApprover.sort_order))
    return [(r.col_a, r.col_b, r.col_c) for r in rows]


def lookup_rows(db: Session) -> list[tuple[int, str, str, str]]:
    """M_LOOKUP のデータ行 [(シートの行番号, カテゴリ, 値, SORT)]（生の文字。本番の見出しは カテゴリ・値・SORT）。"""
    rows = db.scalars(select(models.Lookup).order_by(models.Lookup.sheet_row))
    return [(r.sheet_row, r.category, r.value, r.sort_text) for r in rows]


def advisory_lock(db: Session, key: int, sub: str | None = None) -> None:
    """pg_advisory_xact_lock（PostgreSQL だけ。トランザクションの終わりまで。SQLite は何もしない）。"""
    if db.get_bind().dialect.name != "postgresql":
        return
    if sub is None:
        db.execute(text("SELECT pg_advisory_xact_lock(:k)"), {"k": key})
    else:
        db.execute(text("SELECT pg_advisory_xact_lock(:k, hashtext(:s))"), {"k": key, "s": sub})


# ===== 担当 D・M の関数（10-2） =====


def _service(name: str) -> ModuleType:
    try:
        return importlib.import_module(f"app.services.{name}")
    except ModuleNotFoundError as exc:
        if exc.name == f"app.services.{name}":
            raise RpcError(f"この機能はまだ用意されていません。（services.{name}）") from exc
        raise


def enqueue_after_approve1(db: Session, req_id: str) -> int:
    return _service("gas_jobs").enqueue_after_approve1(db, req_id)


def enqueue_after_approve2(db: Session, req_id: str) -> int:
    return _service("gas_jobs").enqueue_after_approve2(db, req_id)


def enqueue_after_unapprove(db: Session, req_id: str, *, reason: str, old_leave_date: str, old_pdf_url: str,
                            old_pdf_file_id: str, clear_sign: bool) -> int:
    return _service("gas_jobs").enqueue_after_unapprove(
        db, req_id, reason=reason, old_leave_date=old_leave_date, old_pdf_url=old_pdf_url,
        old_pdf_file_id=old_pdf_file_id, clear_sign=clear_sign)


def enqueue_send_warn_mail(db: Session) -> int:
    return _service("gas_jobs").enqueue_send_warn_mail(db)


def summary_mark_dirty(db: Session) -> None:
    _service("summary_sheet").mark_dirty(db)


def summary_rebuild_now(db: Session) -> dict:
    return _service("summary_sheet").rebuild_now(db)


def get_ops_status(db: Session) -> dict:
    return _service("ops_status").get_ops_status(db)


def sync_sheet_config(db: Session) -> dict:
    return _service("config_sync").sync_sheet_config(db)


def sync_masters(db: Session) -> dict:
    return _service("master_sync").sync_masters(db)
