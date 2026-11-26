"""担当 W の試験の小道具（試験は入っていない）。申請・マスタ・設定を DB に置き、業務関数を直接呼ぶ。

D・M の関数（GAS の仕事・集計タブの予約など）は install_fakes で偽物にし、呼ばれた順を記録する。
"""
from __future__ import annotations

import itertools
from datetime import date, datetime

from app import sheet_queue, timeutil
from app.auth import CurrentUser
from app.domain import adapters
from app.models import CalendarDay, DeptApprover, LeaveRequest, Lookup, Setting, Stamp, Worker
from app.rpc import RpcContext

OWNER = "owner@lineworks-local.info"  # BOOTSTRAP_ADMINS
ADMIN = "admin@lineworks-local.info"
SOMU = "somu@lineworks-local.info"
BOSS = "boss@lineworks-local.info"
USER = "shared@lineworks-local.info"

_row = itertools.count(2)


def jst(y, m, d, h=0, mi=0, s=0) -> datetime:
    return datetime(y, m, d, h, mi, s, tzinfo=timeutil.JST)


NOW = jst(2026, 11, 26, 10, 0)


def ctx(db, email=USER, now=None, settings=None) -> RpcContext:
    from conftest import make_settings

    return RpcContext(db=db, user=CurrentUser(email=email, name=email.split("@")[0]),
                      settings=settings or make_settings(), now=(now or NOW).astimezone(timeutil.UTC))


def add_request(db, req_id, *, status="承認済", fiscal_year="2026", fy_kind="number", dept="製造1課", worker_id="101",
                worker_kind="number", worker_name="山田太郎", leave_date=date(2026, 11, 27), leave_text="",
                kubun="全日休", half="", leave_type="有給消化（私用）", substitute=None, substitute_text="",
                applied_at=None, applied_text="", approved_at=None, approved_at2=None, approved_at2_text="",
                approved_by1="", approved_by2="", approved_by2_email="", sign_image_url="", pdf_url="", pdf_file_id="",
                sheet_row="auto", **extra) -> LeaveRequest:
    kinds = {}
    if fy_kind:
        kinds["fiscal_year"] = fy_kind
    if worker_kind:
        kinds["worker_id"] = worker_kind
    if leave_date is not None:
        kinds["leave_date"] = "serial"
    elif leave_text:
        kinds["leave_date"] = "string"
    if substitute is not None:
        kinds["substitute_date"] = "serial"
    row = LeaveRequest(
        req_id=req_id, status=status, fiscal_year=fiscal_year, dept_id=dept, dept_name=dept, worker_id=worker_id,
        worker_name=worker_name, leave_date=leave_date, leave_date_text=leave_text, leave_kubun=kubun,
        half_day_type=half, leave_type=leave_type, substitute_date=substitute, substitute_date_text=substitute_text,
        applied_at=applied_at if applied_at is not None else (jst(2026, 11, 20, 9) if not applied_text else None),
        applied_at_text=applied_text, approved_at=approved_at, approved_at2=approved_at2,
        approved_at2_text=approved_at2_text, approved_by1_email=approved_by1, approved_by2=approved_by2,
        approved_by2_email=approved_by2_email, sign_image_url=sign_image_url, pdf_url=pdf_url,
        pdf_file_id=pdf_file_id, created_by_email=USER, cell_kinds=kinds, row_version=1,
        sheet_row=next(_row) if sheet_row == "auto" else sheet_row, **extra)
    db.add(row)
    db.commit()
    return row


def add_workers(db, rows: list[tuple[str, str, str]], extra: dict | None = None) -> None:
    """[(作業員コード, 氏名, 部署)]。"""
    for i, (code, name, dept) in enumerate(rows, start=1):
        cells = {"作業員コード": code, "氏名": name, "部署": dept, "担当業務": "", "拠点": "", "スタッフ種類": ""}
        cells.update((extra or {}).get(code, {}))
        db.add(Worker(sort_order=i, cells=[[h, v] for h, v in cells.items()]))
    db.commit()


def add_calendar(db, rows: list[tuple[date | None, str]]) -> None:
    for i, (d, kubun) in enumerate(rows, start=1):
        db.add(CalendarDay(sort_order=i, date=d, kubun=kubun, cells=[]))
    db.commit()


def add_stamps(db, rows: list[tuple[str, str, str]]) -> None:
    for i, (a, b, c) in enumerate(rows, start=1):
        db.add(Stamp(sort_order=i, col_a=a, col_b=b, col_c=c, cells=[a, b, c]))
    db.commit()


def add_dept_approvers(db, rows: list[tuple[str, str, str]]) -> None:
    for i, (a, b, c) in enumerate(rows, start=1):
        db.add(DeptApprover(sort_order=i, col_a=a, col_b=b, col_c=c))
    db.commit()


def add_lookup(db, rows: list[tuple[str, str, str]]) -> None:
    for i, (cat, value, sort) in enumerate(rows, start=2):
        db.add(Lookup(sheet_row=i, category=cat, value=value, sort_text=sort))
    db.commit()


def set_setting(db, key: str, value: str) -> None:
    row = db.get(Setting, key)
    if row is None:
        db.add(Setting(key=key, value=value, sheet_row=2))
    else:
        row.value = value
    db.commit()


def standard_masters(db) -> None:
    """管理者・総務・部署承認者・電子印の設定。"""
    set_setting(db, "ADMIN_EMAILS", ADMIN)
    set_setting(db, "SOMU_EMAILS", SOMU)
    set_setting(db, "ALLOW_ADMIN_PAGE_EMAILS", ADMIN)
    add_dept_approvers(db, [("製造1課", BOSS, "上司 一郎"), ("製造2課", f"{BOSS}, other@lineworks-local.info", "")])
    add_stamps(db, [(SOMU, "stamp-1", "総務 花子"), (ADMIN, "stamp-2", "")])


class Recorder:
    def __init__(self):
        self.calls: list[tuple] = []

    def names(self) -> list[str]:
        return [c[0] for c in self.calls]


def install_fakes(monkeypatch) -> Recorder:
    """D・M の関数を偽物にする（呼ばれた順と引数を記録）。"""
    rec = Recorder()
    counter = itertools.count(1)

    def fake(name, result=None):
        def _f(db, *args, **kwargs):
            rec.calls.append((name, args, kwargs))
            return next(counter) if result is None else result
        return _f

    monkeypatch.setattr(adapters, "enqueue_after_approve1", fake("approve1"))
    monkeypatch.setattr(adapters, "enqueue_after_approve2", fake("approve2"))
    monkeypatch.setattr(adapters, "enqueue_after_unapprove", fake("unapprove"))
    monkeypatch.setattr(adapters, "enqueue_send_warn_mail", fake("warn_mail"))
    monkeypatch.setattr(adapters, "summary_mark_dirty", fake("summary", result="x"))
    monkeypatch.setattr(adapters, "summary_rebuild_now", fake("rebuild", result={"ok": True, "rows": 3}))
    monkeypatch.setattr(adapters, "get_ops_status", fake("status", result={"ok": True}))
    monkeypatch.setattr(adapters, "sync_sheet_config", fake("config_sync", result={"ok": True}))
    monkeypatch.setattr(adapters, "sync_masters", fake("master_sync", result={"ok": True}))

    original = sheet_queue.mark_dirty

    def mark_dirty(db, table, pk):
        rec.calls.append(("mark", (table, pk), {}))
        return original(db, table, pk)

    monkeypatch.setattr(sheet_queue, "mark_dirty", mark_dirty)
    return rec
