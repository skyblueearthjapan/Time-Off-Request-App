"""34 関数の登録（rpc.load_domain_handlers が register_all(rpc.register) を呼ぶ。実装仕様 4-2・4-3）。

権限の確かめは業務関数の中で行う（GAS の確かめの順と 400 の文を保つため role は付けない。ops_* は 403）。
"""
from __future__ import annotations

from collections.abc import Callable

from . import approval, masters, ops, requests, summaries

MODULES = (masters, requests, approval, summaries, ops)

# 書く関数（冪等キー・監査）と、監査に残す申請 ID の場所
WRITES: dict[str, str | None] = {
    "api_submitLeaveRequest": None,  # 返事の reqId を残す
    "api_editLeaveRequest": "first",
    "api_deleteLeaveRequest": "first",
    "api_approveLeaveRequest": "first",
    "api_approveLeaveRequest2": "first",
    "ops_retryJob": None,
    "ops_syncSheetConfigNow": None,
    "ops_runMasterSync": None,
    "ops_rebuildSummaryNow": None,
    "ops_sendWarnMail": None,
    "ops_revokeSessions": None,
}


def all_handlers() -> dict[str, Callable]:
    out: dict[str, Callable] = {}
    for module in MODULES:
        for name, fn in module.HANDLERS.items():
            if name in out:
                raise RuntimeError(f"二重に登録しています: {name}")
            out[name] = fn
    return out


def register_all(register: Callable) -> list[str]:
    handlers = all_handlers()
    for name, fn in handlers.items():
        if name in WRITES:
            register(name, fn, write=True, request_id_arg=WRITES[name])
        else:
            register(name, fn)
    return list(handlers)
