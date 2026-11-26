"""申請・編集・取消・一覧・1 件（leaveRequest.gs:6-463・852-1006 の写し。実装仕様 8-1・8-5・8-6・4-4-b〜d）。

書き込みの関数は行ロック（store.lock）で状態を読み直してから決める。commit はしない。
"""
from __future__ import annotations

import functools
import re
from collections.abc import Callable
from datetime import date
from typing import Any

from sqlalchemy import select
from sqlalchemy.orm import Session

from app import timeutil
from app.constants import (
    DAY_FULL,
    DAY_HALF,
    HALF_AM,
    HALF_PM,
    KIND_SUBSTITUTE,
    STATUS_APPROVED,
    STATUS_CANCELED,
    STATUS_REJECTED,
    STATUS_SUBMITTED,
)
from app.jsutil import collate_default, js_string, normalize_gas
from app.models import LeaveRequest
from app.rpc import RpcContext, RpcError

from . import adapters, calendar, quirks, store
from .fiscal import days_ago, fiscal_year, is_retroactive, is_within_current_fy, monthly_period
from .jsdate import js_new_date, js_truthy, jst_ymd_date

YMD = re.compile(r"^\d{4}-\d{2}-\d{2}$")
LOCK_REQ_ID = 739410001
LOCK_SUBSTITUTE = 739410002


def _today(ctx: RpcContext) -> date:
    return timeutil.jst_date_of(ctx.now)


def _obj(value: Any) -> dict:
    """data.x を読む相手（null・文字・配列の .x は undefined）。"""
    return value if isinstance(value, dict) else {}


def _or_empty(value: Any) -> Any:
    """`値 || ''`。"""
    return value if js_truthy(value) else ""


def _sorted_by(items: list[dict], key: str, reverse: bool = False) -> list[dict]:
    sign = -1 if reverse else 1
    return sorted(items, key=functools.cmp_to_key(
        lambda a, b: sign * collate_default(a.get(key) or "", b.get(key) or "")))


# ===== 採番（generateReqId_） =====


def generate_req_id(db: Session, now) -> str:
    """LV-<申請した日 JST>-NNN（その日の最大の連番＋1。3 桁に満たなければ 0 で埋める）。"""
    adapters.advisory_lock(db, LOCK_REQ_ID)
    prefix = "LV-" + timeutil.to_jst_text(now, "yyyyMMdd") + "-"
    max_seq = 0
    for rid in db.scalars(select(LeaveRequest.req_id).where(LeaveRequest.req_id.startswith(prefix, autoescape=True))):
        m = re.match(r"^\s*([+-]?\d+)", rid[len(prefix):])  # parseInt(…, 10)。数字で始まらなければ数えない
        if m:
            max_seq = max(max_seq, int(m.group(1)))
    return prefix + str(max_seq + 1).zfill(3)


# ===== 月次期間・振替元の使い回し =====


def monthly_period_error(leave: date, sub: date) -> str | None:
    """validateSubstituteMonthlyPeriod_。同じ期間なら None。"""
    lp, sp = monthly_period(leave), monthly_period(sub)
    if lp[0] == sp[0]:
        return None
    f = date.isoformat
    return ("振替休暇は同一月次期間（前月16日〜当月15日）内で消化する必要があります。\n"
            f"休暇日: {f(leave)}（期間: {f(lp[0])}〜{f(lp[1])}）\n"
            f"振替元: {f(sub)}（期間: {f(sp[0])}〜{f(sp[1])}）")


def substitute_used_error(db: Session, worker_id: Any, sub_str: str, exclude_req_id: str | None) -> str | None:
    """validateSubstituteUniqueness_（同じ作業員の振替休暇・取消でない行の振替元と同じ日なら文）。"""
    if not js_truthy(worker_id) or not sub_str:
        return None
    adapters.advisory_lock(db, LOCK_SUBSTITUTE, normalize_gas(worker_id))
    if sub_str in calendar.used_substitute_dates(db, worker_id, exclude_req_id):
        return f"指定された振替元出勤日（{sub_str}）は既に別の振替休暇申請で使用されています。"
    return None


# ===== 申請（api_submitLeaveRequest） =====

_TEXT_FIELDS = (
    ("deptId", "dept_id", "部署ID"), ("deptName", "dept_name", "部署名"), ("workerId", "worker_id", "作業員ID"),
    ("workerName", "worker_name", "作業員名"), ("leaveKubun", "leave_kubun", "休暇区分"),
    ("halfDayType", "half_day_type", "半日区分"), ("leaveType", "leave_type", "休暇種類"),
    ("specialReason", "special_reason", "特別理由"), ("paidDetail", "paid_detail", "有給詳細"),
)


def api_submitLeaveRequest(ctx: RpcContext, data=None):
    data = _obj(data)
    db = ctx.db
    today = _today(ctx)
    leave_str = normalize_gas(_or_empty(data.get("leaveDate")))
    if not leave_str or not YMD.match(leave_str):
        return {"ok": False, "error": "休暇日の形式が不正です（yyyy-MM-dd）。"}

    if data.get("isRetroactive") is True:
        if not is_within_current_fy(leave_str, today):
            return {"ok": False, "error": "過去日申請は今年度内（年度開始の3/16から本日まで）の日付のみ指定できます。"}
    elif leave_str not in calendar.working_days(db, today):
        return {"ok": False, "error": "指定された休暇日は勤務日ではありません。過去日の場合は「過去申請」から選択してください。"}

    leave_date = jst_ymd_date(leave_str)
    if leave_date is None:
        # GAS は Invalid Date を書いて失敗する見込み（受け入れる差。4/32 のような日）
        raise RpcError(f"休暇日の形式が不正です: {leave_str}")

    if data.get("leaveType") == KIND_SUBSTITUTE:
        sub_str = normalize_gas(_or_empty(data.get("substituteDate")))
        if not sub_str or not YMD.match(sub_str):
            return {"ok": False, "error": "振替休暇の場合は振替元出勤日を選択してください。"}
        sub_date = jst_ymd_date(sub_str)
        if sub_date is None:
            raise RpcError(f"振替元出勤日の形式が不正です: {sub_str}")
        err = monthly_period_error(leave_date, sub_date)
        if err:
            return {"ok": False, "error": err}
        err = substitute_used_error(db, data.get("workerId"), sub_str, None)
        if err:
            return {"ok": False, "error": err}

    # 振替元は data.substituteDate が真なら種類に関係なく書く（正規化しない元の値）
    sub_value = None
    raw_sub = data.get("substituteDate")
    if js_truthy(raw_sub):
        moment = js_new_date(f"{js_string(raw_sub)}T00:00:00+09:00")
        if moment is None:
            raise RpcError(f"振替元出勤日の形式が不正です: {js_string(raw_sub)}")
        sub_value = timeutil.jst_date_of(moment)

    texts = {}
    for key, col, label in _TEXT_FIELDS:
        value = js_string(_or_empty(data.get(key)))
        if len(value) > quirks.MAX_CELL_CHARS:
            raise RpcError(f"{label}が長すぎます（{quirks.MAX_CELL_CHARS}字まで）。")
        if "\x00" in value:
            raise RpcError(f"{label}に使えない文字が含まれています。")
        texts[col] = value

    req_id = generate_req_id(db, ctx.now)
    row = LeaveRequest(req_id=req_id, status=STATUS_SUBMITTED, cell_kinds={}, row_version=0)
    store.set_number(row, "fiscal_year", fiscal_year(leave_date))
    for col, value in texts.items():
        store.set_text(row, col, value)
    store.set_date(row, "leave_date", leave_date)
    store.set_date(row, "substitute_date", sub_value)
    store.set_date(row, "applied_at", store.ms(ctx.now))
    store.set_text(row, "created_by_email", ctx.user.email or "")
    db.add(row)
    store.mark(db, req_id)
    return {"ok": True, "reqId": req_id}


# ===== 1 件（api_getLeaveRequestById） =====


def _fy_cell_text(row: LeaveRequest) -> str:
    """String(row[年度] || '')（数の 0 は ''）。"""
    t = store.fiscal_year_text(row)
    if store.kind_of(row, "fiscal_year") != "string" and t in ("0", "-0"):
        return ""
    return t


def _date_or_text(row: LeaveRequest, col: str, pattern: str) -> str:
    v = store.date_cell(row, col)
    return store.fmt(v, pattern) if v is not None else store.text_cell(row, col)


def request_view(row: LeaveRequest) -> dict:
    at2 = store.date_cell(row, "approved_at2")
    return {
        "reqId": row.req_id,
        "fiscalYear": _fy_cell_text(row),
        "deptId": store.norm(row, "dept_id"),
        "deptName": store.norm(row, "dept_name"),
        "workerId": store.norm(row, "worker_id"),
        "workerName": store.norm(row, "worker_name"),
        "leaveKubun": store.norm(row, "leave_kubun"),
        "halfDayType": store.norm(row, "half_day_type"),
        "leaveDate": _date_or_text(row, "leave_date", "yyyy-MM-dd"),
        "leaveType": store.norm(row, "leave_type"),
        "substituteDate": _date_or_text(row, "substitute_date", "yyyy-MM-dd"),
        "specialReason": store.norm(row, "special_reason"),
        "paidDetail": store.norm(row, "paid_detail"),
        "additionalDetail": store.norm(row, "additional_detail"),
        "submittedAt": _date_or_text(row, "applied_at", "yyyy-MM-dd HH:mm:ss"),
        "approvedAt": _date_or_text(row, "approved_at", "yyyy-MM-dd HH:mm:ss"),
        "status": store.norm(row, "status"),
        "signImageUrl": store.norm(row, "sign_image_url"),
        "pdfUrl": store.norm(row, "pdf_url"),
        "approvedBy1Email": store.norm(row, "approved_by1_email"),
        "approvedBy2": store.norm(row, "approved_by2"),
        "approvedBy2Email": store.norm(row, "approved_by2_email"),
        "approvedAt2": store.fmt(at2, "yyyy-MM-dd HH:mm:ss") if at2 is not None else "",
        "rowNo": row.sheet_row,
    }


def api_getLeaveRequestById(ctx: RpcContext, req_id=None):
    if not js_truthy(req_id):
        return None
    row = store.find(ctx.db, req_id)
    return request_view(row) if row is not None else None


# ===== 一覧 =====


def _leave_iso(row: LeaveRequest) -> str:
    d = store.date_cell(row, "leave_date")
    return store.fmt(d) if d is not None else ""


def upcoming_leaves(db: Session, today: date, dept_id: Any, worker_id: Any) -> list[dict]:
    out = []
    for row in store.rows_in_sheet_order(db):
        if not normalize_gas(row.req_id):
            continue
        status = store.norm(row, "status")
        if status == STATUS_CANCELED:
            continue
        if js_truthy(dept_id) and store.norm(row, "dept_id") != dept_id:
            continue
        if js_truthy(worker_id) and store.norm(row, "worker_id") != worker_id:
            continue
        leave = store.date_cell(row, "leave_date")
        if leave is not None and leave < today:
            continue
        iso = _leave_iso(row)
        out.append({
            "reqId": normalize_gas(row.req_id),
            "deptName": store.norm(row, "dept_name"),
            "workerName": store.norm(row, "worker_name"),
            "leaveKubun": store.norm(row, "leave_kubun"),
            "halfDayType": store.norm(row, "half_day_type"),
            "leaveDate": store.fmt(leave, "yyyy/MM/dd") if leave is not None else "",
            "leaveType": store.norm(row, "leave_type"),
            "status": status,
            "isRetroactive": is_retroactive(store.raw_cell(row, "applied_at"), iso),
            "daysAgo": days_ago(iso, today),
        })
    return _sorted_by(out, "leaveDate")


def api_getUpcomingLeaves(ctx: RpcContext, dept_id=None, worker_id=None):
    return upcoming_leaves(ctx.db, _today(ctx), dept_id, worker_id)


def api_getAllPendingRequests(ctx: RpcContext):
    today = _today(ctx)
    out = []
    moments = []
    for row in store.rows_in_sheet_order(ctx.db):
        rid = normalize_gas(row.req_id)
        if not rid:
            continue
        status = store.norm(row, "status")
        if status != STATUS_SUBMITTED:
            continue
        leave = store.date_cell(row, "leave_date")
        submitted = store.date_cell(row, "applied_at")
        iso = _leave_iso(row)
        out.append({
            "reqId": rid,
            "deptName": store.norm(row, "dept_name"),
            "workerName": store.norm(row, "worker_name"),
            "leaveKubun": store.norm(row, "leave_kubun"),
            "halfDayType": store.norm(row, "half_day_type"),
            "leaveDate": store.fmt(leave, "yyyy/MM/dd") if leave is not None else "",
            "leaveType": store.norm(row, "leave_type"),
            "submittedAt": store.fmt(submitted, "MM/dd HH:mm") if submitted is not None else "",
            "status": status,
            "isRetroactive": is_retroactive(store.raw_cell(row, "applied_at"), iso),
            "daysAgo": days_ago(iso, today),
        })
        moments.append(submitted)
    if not quirks.PENDING_SORT_WITH_YEAR:
        return _sorted_by(out, "submittedAt", reverse=True)
    # K-10: 本当の申請日時の新しい順（日時の無い行は後ろ。同じなら元の順）
    order = sorted(range(len(out)), key=lambda i: (moments[i] is None,
                                                   -(moments[i].timestamp()) if moments[i] is not None else 0))
    return [out[i] for i in order]


# ===== 編集（api_editLeaveRequest） =====


def _ymd_for_gas(leave_date: str) -> str:
    """GAS の内部用の入口へ渡す休暇日（必ず 'yyyy-MM-dd'。文字のセルは日付として読む。読めなければ ''）。"""
    if YMD.match(leave_date):
        return leave_date
    moment = store.cell_moment(leave_date)
    return store.fmt(moment) if moment is not None else ""


def _unapprove_job(db: Session, row: LeaveRequest, *, reason: str, old_leave_date: str, clear_sign: bool) -> None:
    adapters.enqueue_after_unapprove(db, row.req_id, reason=reason, old_leave_date=old_leave_date,
                                     old_pdf_url=row.pdf_url or "", old_pdf_file_id=row.pdf_file_id or "",
                                     clear_sign=clear_sign)


def api_editLeaveRequest(ctx: RpcContext, req_id=None, patch=None):
    db = ctx.db
    row = store.lock(db, req_id)
    if row is None:
        raise RpcError("申請が見つかりません。")
    req = request_view(row)
    if req["status"] in (STATUS_CANCELED, STATUS_REJECTED):
        raise RpcError("取消/却下済みの申請は編集できません。")

    patch = _obj(patch)
    leave_str = normalize_gas(_or_empty(patch.get("leaveDate")))
    kubun = normalize_gas(_or_empty(patch.get("leaveKubun")))
    half = normalize_gas(_or_empty(patch.get("halfDayType")))
    sub_str = normalize_gas(_or_empty(patch.get("substituteDate")))

    if leave_str and not YMD.match(leave_str):
        raise RpcError(f"休暇日の形式が不正です: {leave_str}")
    if kubun and kubun not in (DAY_FULL, DAY_HALF):
        raise RpcError(f"休暇区分が不正です: {kubun}")
    if kubun == DAY_HALF and half not in (HALF_AM, HALF_PM):
        raise RpcError("半日区分を選択してください。")
    if sub_str and not YMD.match(sub_str):
        raise RpcError(f"振替元出勤日の形式が不正です: {sub_str}")

    if req["leaveType"] == KIND_SUBSTITUTE:
        eff_leave = leave_str or req["leaveDate"]
        eff_sub = sub_str or req["substituteDate"]
        if not eff_sub:
            raise RpcError("振替休暇には振替元出勤日が必須です。編集画面から振替元を指定してください。")
        if not eff_leave:
            raise RpcError("振替休暇には休暇日が必須です。")
        ld, sd = jst_ymd_date(eff_leave), jst_ymd_date(eff_sub)
        if ld is None:
            raise RpcError(f"休暇日の形式が不正です: {eff_leave}")
        if sd is None:
            raise RpcError(f"振替元出勤日の形式が不正です: {eff_sub}")
        err = monthly_period_error(ld, sd)
        if err:
            raise RpcError(err)
        if sub_str and sub_str != req["substituteDate"]:
            err = substitute_used_error(db, req["workerId"], sub_str, row.req_id)
            if err:
                raise RpcError(err)

    if leave_str:
        new_leave = jst_ymd_date(leave_str)
        if new_leave is None:
            raise RpcError(f"休暇日の形式が不正です: {leave_str}")
        store.set_date(row, "leave_date", new_leave)
        store.set_number(row, "fiscal_year", fiscal_year(new_leave))
    if kubun:
        store.set_text(row, "leave_kubun", kubun)
        if kubun == DAY_FULL:
            store.set_text(row, "half_day_type", "")
        elif half:
            store.set_text(row, "half_day_type", half)
    if sub_str and req["leaveType"] == KIND_SUBSTITUTE:
        store.set_date(row, "substitute_date", jst_ymd_date(sub_str))

    new_status = req["status"]
    was_approved = req["status"] == STATUS_APPROVED
    if was_approved:
        new_status = STATUS_SUBMITTED
        row.status = STATUS_SUBMITTED
        store.set_date(row, "approved_at", None)
        for col in ("approved_by1_email", "sign_image_url", "approved_by2", "approved_by2_email"):
            store.set_text(row, col, "")
        store.set_date(row, "approved_at2", None)
        row.sign_png = None
        # pdf_url は残す（'asis'。'trash' のときも、空にするのは GAS の返事を受けた M の record_gas_result）
    store.mark(db, row.req_id)
    if was_approved:
        _unapprove_job(db, row, reason="edit", old_leave_date=_ymd_for_gas(req["leaveDate"]), clear_sign=True)
    adapters.summary_mark_dirty(db)
    return {"ok": True, "newStatus": new_status}


# ===== 取消（api_deleteLeaveRequest） =====


def api_deleteLeaveRequest(ctx: RpcContext, req_id=None):
    db = ctx.db
    row = store.lock(db, req_id)
    if row is None:
        raise RpcError("申請が見つかりません。")
    req = request_view(row)
    if req["status"] == STATUS_CANCELED:
        raise RpcError("既に取消済みです。")
    row.status = STATUS_CANCELED
    store.mark(db, row.req_id)
    _unapprove_job(db, row, reason="cancel", old_leave_date=_ymd_for_gas(req["leaveDate"]), clear_sign=False)
    adapters.summary_mark_dirty(db)
    return {"ok": True}


HANDLERS: dict[str, Callable] = {
    "api_submitLeaveRequest": api_submitLeaveRequest,
    "api_getLeaveRequestById": api_getLeaveRequestById,
    "api_getUpcomingLeaves": api_getUpcomingLeaves,
    "api_getAllPendingRequests": api_getAllPendingRequests,
    "api_editLeaveRequest": api_editLeaveRequest,
    "api_deleteLeaveRequest": api_deleteLeaveRequest,
}
