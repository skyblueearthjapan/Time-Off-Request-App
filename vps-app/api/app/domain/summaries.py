"""有給の集計・警告・期間別・取得状況の画面（leaveSummary.gs・leaveStatusApi.gs の写し。実装仕様 6-3・8-4・4-4-f〜i）。

- summary_rows(db, now): rebuildLeaveSummary_ の集計（V_LEAVE_SUMMARY の行。担当 D が書く）。
- last_paid_date_display(value, style): 画面へ返す「最終有給日」の文字（quirks.LAST_PAID_DATE_STYLE）。
- 画面の関数は、GAS が「集計タブを書き直してから読み直した値」を DB から同じ計算で作って返す（シートは書かない）。
  読み直しの癖（数に見える文字は数、最終有給日は日付のセル）も写す。
"""
from __future__ import annotations

import functools
from collections.abc import Callable
from datetime import date, datetime
from typing import Any
from zoneinfo import ZoneInfo

from sqlalchemy.orm import Session

from app import permissions, timeutil
from app.config import get_settings
from app.constants import DAY_HALF, KIND_PAID, KIND_SPECIAL, KIND_SUBSTITUTE, STATUS_APPROVED
from app.jsutil import collate_ja, js_sort_key, js_string, normalize_gas, sheet_readback
from app.rpc import RpcContext, RpcError

from . import quirks, store
from .fiscal import fiscal_year
from .jsdate import js_date_string, js_local, js_number, js_truthy, out_number
from .masters import dept_names, workers_by_dept
from .requests import upcoming_leaves

# config.gs:43-58
PAID_LEAVE_WARN = (
    {"month": 6, "day": 30, "required": 1, "label": "6月末"},
    {"month": 8, "day": 31, "required": 2, "label": "8月末"},
    {"month": 12, "day": 31, "required": 3, "label": "12月末"},
    {"month": 2, "day": 15, "required": 4, "label": "2/15"},
    {"month": 3, "day": 15, "required": 5, "label": "年度末"},
)
PAID_LEAVE_QUARTERS = (
    {"id": "Q1", "label": "春期(3/16-6月)", "sm": 3, "sd": 16, "em": 6, "ed": 30, "target": 1},
    {"id": "SUMMER", "label": "夏期(7-8月)", "sm": 7, "sd": 1, "em": 8, "ed": 31, "target": 2},
    {"id": "Q3", "label": "秋期(9-11月)", "sm": 9, "sd": 1, "em": 11, "ed": 30, "target": 3},
    {"id": "WINTER", "label": "冬期(12-2/15)", "sm": 12, "sd": 1, "em": 2, "ed": 15, "target": 4},
    {"id": "YEAREND", "label": "年度末(2/16-3/15)", "sm": 2, "sd": 16, "em": 3, "ed": 15, "target": 5},
)


def _fy_year(month: int, day: int, fy: float) -> float:
    """3/16 以降の月日は fy 年、ほかは fy+1 年。"""
    return fy if (month > 3 or (month == 3 and day >= 16)) else fy + 1


def _lt(a: datetime | None, b: datetime | None) -> bool:
    """JS の a < b（Invalid Date との比べは偽）。"""
    return a is not None and b is not None and a < b


def compute_warn_level(paid: float, fiscal_year_value: Any, now: datetime) -> str:
    """computeWarnLevel_。"""
    fy = js_number(fiscal_year_value)
    for w in reversed(PAID_LEAVE_WARN):
        deadline = js_local(_fy_year(w["month"], w["day"], fy), w["month"] - 1, w["day"], 23, 59, 59)
        if _lt(now, deadline):
            continue
        if paid < w["required"]:
            deficit = w["required"] - paid
            if deficit >= 3:
                return "危険"
            if deficit >= 2:
                return "警告"
            return "注意"
        break
    return ""


def compute_quarter_status(dates: list[date], fiscal_year_value: Any, now: datetime) -> list[dict]:
    """computeQuarterStatus_（期間に入る休暇日の件数。半日も 1）。"""
    fy = js_number(fiscal_year_value)
    out = []
    cumulative = 0
    moments = [timeutil.jst_midnight(d) for d in dates]
    for q in PAID_LEAVE_QUARTERS:
        start = js_local(_fy_year(q["sm"], q["sd"], fy), q["sm"] - 1, q["sd"], 0, 0, 0)
        end = js_local(_fy_year(q["em"], q["ed"], fy), q["em"] - 1, q["ed"], 23, 59, 59)
        count = 0
        if start is not None and end is not None:
            count = sum(1 for d in moments if start <= d <= end)
        cumulative += count
        is_current = start is not None and end is not None and start <= now <= end
        is_past = end is not None and now > end
        is_future = start is not None and now < start
        shortage = max(0, q["target"] - cumulative)
        level = ""
        if is_current or is_past:
            level = "danger" if shortage >= 2 else ("warn" if shortage == 1 else "ok")
        out.append({"id": q["id"], "label": q["label"], "targetCount": q["target"], "periodCount": count,
                    "cumulativeCount": cumulative, "shortage": shortage, "level": level, "taken": count > 0,
                    "overdue": shortage > 0 and is_past, "isCurrent": is_current, "isPast": is_past,
                    "isFuture": is_future})
    return out


def build_milestones(paid: float, fiscal_year_value: Any, now: datetime) -> list[dict]:
    """buildMilestones_。"""
    fy = js_number(fiscal_year_value)
    out = []
    for w in PAID_LEAVE_WARN:
        m, day = w["month"], w["day"]
        deadline = js_local(_fy_year(m, day, fy), m - 1, day, 23, 59, 59)
        passed = deadline is not None and now >= deadline
        met = paid >= w["required"]
        deficit = max(0, w["required"] - paid)
        if met:
            state = "done"
        elif passed:
            state = "risk" if deficit >= 3 else ("warn" if deficit >= 2 else "note")
        else:
            state = "future"
        out.append({"label": w["label"], "month": m, "day": day, "dateText": f"{m}/{day}",
                    "required": w["required"], "met": met, "passed": passed, "deficit": out_number(float(deficit)),
                    "remaining": out_number(float(deficit)), "state": state})
    return out


# ===== V_LEAVE_SUMMARY（rebuildLeaveSummary_） =====


def summary_rows(db: Session, now: datetime) -> list[dict]:
    """集計の行（キー＝年度_作業員ID の JS の既定の並び）。fiscalYear は元のセルの値と型、lastPaidDate は 'yyyy-MM-dd'。"""
    summary: dict[str, dict] = {}
    for row in store.rows_in_sheet_order(db):
        if not normalize_gas(row.req_id):
            continue
        if store.norm(row, "status") != STATUS_APPROVED:
            continue
        worker_id = store.norm(row, "worker_id")
        key = store.fiscal_year_text(row) + "_" + worker_id
        s = summary.get(key)
        if s is None:
            s = summary[key] = {
                "fiscalYear": store.fiscal_year_value(row), "deptId": store.norm(row, "dept_id"),
                "deptName": store.norm(row, "dept_name"), "workerId": worker_id,
                "workerName": store.norm(row, "worker_name"), "paidCount": 0.0, "paidHalfCount": 0,
                "substituteCount": 0, "specialCount": 0, "totalCount": 0, "lastPaidDate": "",
            }
        s["totalCount"] += 1
        leave_type = store.norm(row, "leave_type")
        if leave_type == KIND_PAID:
            if store.norm(row, "leave_kubun") == DAY_HALF:
                s["paidHalfCount"] += 1
                s["paidCount"] += 0.5
            else:
                s["paidCount"] += 1
            leave = store.date_cell(row, "leave_date")
            if leave is not None:
                ld = store.fmt(leave)
                if not s["lastPaidDate"] or ld > s["lastPaidDate"]:
                    s["lastPaidDate"] = ld
        elif leave_type == KIND_SUBSTITUTE:
            s["substituteCount"] += 1
        elif leave_type == KIND_SPECIAL:
            s["specialCount"] += 1
    out = []
    for key in sorted(summary, key=js_sort_key):
        s = summary[key]
        s["warnLevel"] = compute_warn_level(s["paidCount"], s["fiscalYear"], now)
        s["paidCount"] = out_number(s["paidCount"])
        out.append(s)
    return out


def last_paid_date_display(value: str, style: str | None = None, sheet_tz: str | None = None) -> str:
    """画面に返す「最終有給日」。'gas' は集計タブの日付のセル（シートの時間帯の 0 時）を String(Date) にした文字。"""
    style = style or quirks.LAST_PAID_DATE_STYLE
    if style != "gas" or not value:
        return value
    try:
        d = date.fromisoformat(value)
    except ValueError:
        return value
    tz = ZoneInfo(sheet_tz or get_settings().db_sheet_timezone)
    return js_date_string(datetime(d.year, d.month, d.day, tzinfo=tz))


def _readback_text(value: Any) -> str:
    """集計タブへ書いた文字を GAS が normalize_ で読み直した文字（数に見える文字は数）。"""
    return normalize_gas(sheet_readback(value)[0]) if js_string(value) != "" else ""


def summary_sheet_view(db: Session, now: datetime, sheet_tz: str) -> list[dict]:
    """V_LEAVE_SUMMARY を GAS が読み直した形（年度は Number、回数は数、文字は読み直し）。"""
    out = []
    for s in summary_rows(db, now):
        out.append({
            "fy": js_number(s["fiscalYear"]),
            "deptId": _readback_text(s["deptId"]), "deptName": _readback_text(s["deptName"]),
            "workerId": _readback_text(s["workerId"]), "workerName": _readback_text(s["workerName"]),
            "paidCount": s["paidCount"], "paidHalfCount": s["paidHalfCount"],
            "substituteCount": s["substituteCount"], "specialCount": s["specialCount"],
            "totalCount": s["totalCount"],
            "lastPaidDate": normalize_gas(last_paid_date_display(s["lastPaidDate"], None, sheet_tz)),
            "warnLevel": normalize_gas(s["warnLevel"]),
        })
    return out


# ===== 画面（buildAdminViewData_ ほか） =====


def _today(ctx: RpcContext) -> date:
    return timeutil.jst_date_of(ctx.now)


def _sheet_tz(ctx: RpcContext) -> str:
    return ctx.settings.db_sheet_timezone


def build_admin_view(ctx: RpcContext, depts: list, fy_year: Any) -> list[dict]:
    """buildAdminViewData_。"""
    db, now = ctx.db, ctx.now
    fy = js_number(fy_year)
    by_worker: dict[str, dict] = {}
    for s in summary_sheet_view(db, now, _sheet_tz(ctx)):
        if s["fy"] != fy:
            continue
        by_worker[s["workerId"]] = s
    paid_dates: dict[str, list[date]] = {}
    for row in store.rows_in_sheet_order(db):
        if not normalize_gas(row.req_id) or store.norm(row, "status") != STATUS_APPROVED:
            continue
        if js_number(store.fiscal_year_text(row)) != fy:
            continue
        if store.norm(row, "leave_type") != KIND_PAID:
            continue
        leave = store.date_cell(row, "leave_date")
        if leave is None:
            continue
        paid_dates.setdefault(store.norm(row, "worker_id"), []).append(leave)

    wanted = {normalize_gas(d) for d in depts}
    out = []
    for dept, workers in workers_by_dept(db).items():
        if normalize_gas(dept) not in wanted:
            continue
        for w in workers:
            wid = w["workerId"]
            s = by_worker.get(wid)
            paid = s["paidCount"] if s else 0
            quarters = compute_quarter_status(paid_dates.get(wid, []), fy, now)
            shortage = max([q["shortage"] for q in quarters if q["isCurrent"] or q["isPast"]], default=0)
            out.append({
                "deptName": dept, "workerId": wid, "workerName": w["workerName"],
                "paidCount": paid,
                "paidHalfCount": s["paidHalfCount"] if s else 0,
                "substituteCount": s["substituteCount"] if s else 0,
                "specialCount": s["specialCount"] if s else 0,
                "totalCount": s["totalCount"] if s else 0,
                "lastPaidDate": s["lastPaidDate"] if s else "",
                "warnLevel": compute_warn_level(js_number(paid), fy, now),
                "quarters": quarters,
                "currentShortage": shortage,
            })

    def cmp(a: dict, b: dict) -> int:
        if a["paidCount"] != b["paidCount"]:
            return -1 if a["paidCount"] < b["paidCount"] else 1
        if a["currentShortage"] != b["currentShortage"]:
            return -1 if b["currentShortage"] < a["currentShortage"] else 1
        c = collate_ja(a["deptName"], b["deptName"])
        return c if c else collate_ja(a["workerName"], b["workerName"])

    return sorted(out, key=functools.cmp_to_key(cmp))


def _fy_or_today(ctx: RpcContext, fy_year: Any) -> Any:
    return fy_year if js_truthy(fy_year) else fiscal_year(_today(ctx))


def api_getSomuAdminView(ctx: RpcContext, fy_year=None, dept_filter=None):
    if not permissions.is_somu(ctx.db, ctx.user.email) and not permissions.is_admin(ctx.db, ctx.user.email):
        raise RpcError("権限がありません（総務または管理者のみ）")
    fy_year = _fy_or_today(ctx, fy_year)
    if not js_truthy(dept_filter) or normalize_gas(dept_filter) in ("ALL", ""):
        depts = dept_names(ctx.db)
    else:
        depts = [dept_filter]
    return build_admin_view(ctx, depts, fy_year)


def api_getDeptAdminView(ctx: RpcContext, fy_year=None, selected_depts=None):
    admin = permissions.is_admin(ctx.db, ctx.user.email)
    if not admin and not permissions.is_dept_approver(ctx.db, ctx.user.email):
        raise RpcError("部署承認者の権限がありません。")
    selected = selected_depts if isinstance(selected_depts, list) else []
    if selected and admin:
        depts = selected
    elif selected:
        mine = permissions.dept_approver_depts(ctx.db, ctx.user.email)
        depts = [d for d in selected if d in mine]
    else:
        depts = permissions.dept_approver_depts(ctx.db, ctx.user.email)
    if not depts:
        return {"workers": [], "deptNames": []}
    fy = _fy_or_today(ctx, fy_year)
    return {"workers": build_admin_view(ctx, depts, fy), "deptNames": depts}


def admin_boss_view(ctx: RpcContext, dept_id: Any, fy_year: Any) -> list[dict]:
    fy = js_number(_fy_or_today(ctx, fy_year))
    out = []
    for s in summary_sheet_view(ctx.db, ctx.now, _sheet_tz(ctx)):
        if s["fy"] != fy:
            continue
        if js_truthy(dept_id) and s["deptId"] != dept_id:
            continue
        out.append({k: s[k] for k in ("deptId", "deptName", "workerId", "workerName", "paidCount", "paidHalfCount",
                                       "substituteCount", "specialCount", "totalCount", "lastPaidDate", "warnLevel")})
    return sorted(out, key=functools.cmp_to_key(
        lambda a, b: collate_ja(a["deptName"] + a["workerName"], b["deptName"] + b["workerName"])))


def _require_admin_page(ctx: RpcContext) -> None:
    if not permissions.can_access_admin(ctx.db, ctx.user.email):
        raise RpcError("権限がありません")


def api_getAdminBossView(ctx: RpcContext, dept_id=None, fy_year=None):
    _require_admin_page(ctx)
    return admin_boss_view(ctx, dept_id, fy_year)


def api_getAdminSomuView(ctx: RpcContext, fy_year=None, dept_filter=None):
    _require_admin_page(ctx)
    return admin_boss_view(ctx, dept_filter if js_truthy(dept_filter) else "", fy_year)


def api_getLeaveStatusSearch(ctx: RpcContext, fy_year=None, dept_id=None, worker_id=None):
    today = _today(ctx)
    fy = js_number(fy_year) if js_truthy(fy_year) else float(fiscal_year(today))
    depts = [dept_id] if js_truthy(dept_id) else dept_names(ctx.db)
    workers = build_admin_view(ctx, depts, fy)
    fy_out = out_number(fy)
    if js_truthy(worker_id):
        me = next((w for w in workers if w["workerId"] == worker_id), None)
        upcoming = (upcoming_leaves(ctx.db, today, me["deptName"], worker_id)
                    if me is not None and fy == fiscal_year(today) else [])
        return {"mode": "personal", "fyYear": fy_out, "worker": me,
                "milestones": build_milestones(js_number(me["paidCount"]), fy, ctx.now) if me else [],
                "upcoming": upcoming}
    return {"mode": "list", "fyYear": fy_out, "deptFilter": dept_id if js_truthy(dept_id) else "", "workers": workers}


HANDLERS: dict[str, Callable] = {
    "api_getSomuAdminView": api_getSomuAdminView,
    "api_getDeptAdminView": api_getDeptAdminView,
    "api_getAdminSomuView": api_getAdminSomuView,
    "api_getAdminBossView": api_getAdminBossView,
    "api_getLeaveStatusSearch": api_getLeaveStatusSearch,
}
