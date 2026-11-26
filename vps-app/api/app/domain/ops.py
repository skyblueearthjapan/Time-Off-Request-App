"""運用者だけの関数（実装仕様 4-2・06 の K-9）。permissions.is_operator が偽なら 403（rpc も同じ確かめをする）。

担当 D・M の仕事は adapters 経由で 10-2 の名前で呼ぶ。
"""
from __future__ import annotations

from collections.abc import Callable

from app import auth, jobs, permissions
from app.rpc import RpcContext, RpcError, RpcForbidden

from . import adapters
from .jsdate import js_truthy


def _require_operator(ctx: RpcContext) -> None:
    if not permissions.is_operator(ctx.db, ctx.user.email, ctx.settings):
        raise RpcForbidden()


def ops_getStatus(ctx: RpcContext):
    _require_operator(ctx)
    return adapters.get_ops_status(ctx.db)


def ops_retryJob(ctx: RpcContext, job_id=None):
    _require_operator(ctx)
    if isinstance(job_id, int) and not isinstance(job_id, bool):
        jid = job_id
    elif isinstance(job_id, str) and job_id.isascii() and job_id.isdigit():
        jid = int(job_id)
    else:
        raise RpcError("仕事の番号が正しくありません。")
    if not jobs.retry_failed(ctx.db, jid):
        raise RpcError("失敗した仕事が見つかりません。")
    return {"ok": True}


def ops_syncSheetConfigNow(ctx: RpcContext):
    _require_operator(ctx)
    return adapters.sync_sheet_config(ctx.db)


def ops_runMasterSync(ctx: RpcContext):
    _require_operator(ctx)
    return adapters.sync_masters(ctx.db)


def ops_rebuildSummaryNow(ctx: RpcContext):
    _require_operator(ctx)
    return adapters.summary_rebuild_now(ctx.db)


def ops_sendWarnMail(ctx: RpcContext):
    _require_operator(ctx)
    return {"ok": True, "jobId": adapters.enqueue_send_warn_mail(ctx.db)}


def ops_revokeSessions(ctx: RpcContext, email=None):
    _require_operator(ctx)
    if not js_truthy(email) or not isinstance(email, str):
        raise RpcError("メールアドレスを指定してください。")
    return {"ok": True, "revoked": auth.revoke_sessions(ctx.db, email)}


HANDLERS: dict[str, Callable] = {
    "ops_getStatus": ops_getStatus,
    "ops_retryJob": ops_retryJob,
    "ops_syncSheetConfigNow": ops_syncSheetConfigNow,
    "ops_runMasterSync": ops_runMasterSync,
    "ops_rebuildSummaryNow": ops_rebuildSummaryNow,
    "ops_sendWarnMail": ops_sendWarnMail,
    "ops_revokeSessions": ops_revokeSessions,
}
