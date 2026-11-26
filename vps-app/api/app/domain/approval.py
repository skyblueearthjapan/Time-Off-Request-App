"""1 次承認・2 次承認（leaveRequest.gs:471-726 の写し。実装仕様 8-7・8-8・4-4-e・K-4）。

- 行ロックの後で状態を読み直す。1 次承認は「申請中」だけ、2 次承認は「承認済で 2 次承認がまだ」だけ（変えるもの）。
- PDF・サイン画像・台帳・承認完了メールは GAS の仕事（担当 M）。サインの PNG は leave_requests.sign_png に置く。
- 呼び順（7-4）: 行を書く → 行の反映の印 → GAS の仕事 → 集計タブの予約（1 次承認だけ）。
"""
from __future__ import annotations

import base64
import binascii
import re
from collections.abc import Callable

from app import permissions, timeutil
from app.constants import STATUS_APPROVED, STATUS_SUBMITTED
from app.jsutil import js_string, js_trim, normalize_gas
from app.rpc import RpcContext, RpcError

from . import adapters, quirks, store
from .fiscal import days_ago, is_retroactive
from .jsdate import js_truthy
from .masters import stamp_name
from .requests import _sorted_by

ERR_NO_PERMISSION = "権限がありません（総務または管理者のみ）"
_DATA_URL_PREFIX = re.compile(r"^data:image/\w+;base64,")


def _require_somu_or_admin(ctx: RpcContext) -> None:
    if not permissions.is_somu(ctx.db, ctx.user.email) and not permissions.is_admin(ctx.db, ctx.user.email):
        raise RpcError(ERR_NO_PERMISSION)


def _find_locked(ctx: RpcContext, req_id):
    if store.row_count(ctx.db) == 0:
        raise RpcError("申請データがありません。")
    row = store.lock(ctx.db, req_id)
    if row is None:
        raise RpcError(f"申請が見つかりません: {js_string(req_id)}")
    return row


def decode_sign(sign_base64) -> bytes | None:
    """saveSignImage_ と同じく、先頭の data:image/…;base64, を外して読む（読めなければ None＝GAS も URL 空で続ける）。"""
    if not isinstance(sign_base64, str) or not sign_base64:
        return None
    body = _DATA_URL_PREFIX.sub("", sign_base64, count=1)
    try:
        data = base64.b64decode(body, validate=True)
    except (binascii.Error, ValueError):
        return None
    return data or None


def _approved_at2_set(row) -> bool:
    """2 次承認日時が空でない（Date のセルか、文字を trim して空でない）。"""
    return store.date_cell(row, "approved_at2") is not None or js_trim(store.text_cell(row, "approved_at2")) != ""


# ===== 1 次承認（api_approveLeaveRequest） =====


def api_approveLeaveRequest(ctx: RpcContext, req_id=None, sign_base64=None, approver_email=None):
    if not js_truthy(req_id):
        raise RpcError("REQ_IDが指定されていません。")
    db = ctx.db
    row = _find_locked(ctx, req_id)
    status = store.norm(row, "status")
    if status == STATUS_APPROVED:
        raise RpcError("この申請はすでに承認済みです。")
    if status != STATUS_SUBMITTED:
        raise RpcError("この申請は取消済みです。")
    if isinstance(sign_base64, str) and len(sign_base64) > quirks.SIGN_MAX_CHARS:
        raise RpcError("サイン画像が大きすぎます。")

    row.status = STATUS_APPROVED
    store.set_date(row, "approved_at", store.ms(ctx.now))
    store.set_text(row, "approved_by1_email", approver_email if js_truthy(approver_email) else "")
    row.sign_image_url = ""  # シートは GAS が URL か空を書く
    row.sign_png = decode_sign(sign_base64) if js_truthy(sign_base64) else None
    store.mark(db, row.req_id)
    adapters.enqueue_after_approve1(db, row.req_id)
    adapters.summary_mark_dirty(db)
    return {"ok": True, "reqId": req_id}


# ===== 2 次承認待ち（api_getPending2ndApprovals） =====


def api_getPending2ndApprovals(ctx: RpcContext):
    _require_somu_or_admin(ctx)
    today = timeutil.jst_date_of(ctx.now)
    out = []
    for row in store.rows_in_sheet_order(ctx.db):
        rid = normalize_gas(row.req_id)
        if not rid:
            continue
        if store.norm(row, "status") != STATUS_APPROVED:
            continue
        if _approved_at2_set(row):
            continue
        leave = store.date_cell(row, "leave_date")
        approved = store.date_cell(row, "approved_at")
        iso = store.fmt(leave) if leave is not None else ""
        out.append({
            "reqId": rid,
            "deptName": store.norm(row, "dept_name"),
            "workerName": store.norm(row, "worker_name"),
            "leaveDate": store.fmt(leave, "yyyy/MM/dd") if leave is not None else store.text_cell(row, "leave_date"),
            "leaveKubun": store.norm(row, "leave_kubun"),
            "leaveType": store.norm(row, "leave_type"),
            "approvedAt": store.fmt(approved, "yyyy/MM/dd HH:mm") if approved is not None else "",
            "isRetroactive": is_retroactive(store.raw_cell(row, "applied_at"), iso),
            "daysAgo": days_ago(iso, today),
        })
    return _sorted_by(out, "leaveDate")


# ===== 2 次承認（api_approveLeaveRequest2） =====


def api_approveLeaveRequest2(ctx: RpcContext, req_id=None, approver2_email=None):
    if not js_truthy(req_id):
        raise RpcError("REQ_IDが指定されていません。")
    _require_somu_or_admin(ctx)
    db = ctx.db
    row = _find_locked(ctx, req_id)
    if store.norm(row, "status") != STATUS_APPROVED:
        raise RpcError("1次承認されていない申請です。")
    if _approved_at2_set(row):
        raise RpcError("この申請はすでに2次承認済みです。")

    email = ctx.user.email if quirks.SECOND_APPROVER_FROM_LOGIN else approver2_email
    name = email if js_truthy(email) else ""
    if js_truthy(email) and isinstance(email, str):
        try:
            found = stamp_name(db, email)
            if found is not None:
                name = found or email
        except Exception:  # noqa: BLE001 - GAS も握る
            pass
    store.set_text(row, "approved_by2", name)
    store.set_text(row, "approved_by2_email", email if email is not None else "")
    store.set_date(row, "approved_at2", store.ms(ctx.now))
    store.mark(db, row.req_id)
    adapters.enqueue_after_approve2(db, row.req_id)
    return {"ok": True, "reqId": req_id}


HANDLERS: dict[str, Callable] = {
    "api_approveLeaveRequest": api_approveLeaveRequest,
    "api_getPending2ndApprovals": api_getPending2ndApprovals,
    "api_approveLeaveRequest2": api_approveLeaveRequest2,
}
