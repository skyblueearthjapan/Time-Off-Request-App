"""部署・作業員・LOOKUP・電子印・権限の関数（masters.gs:8-210・auth.gs:104-203・config.gs:233 の写し）。"""
from __future__ import annotations

import functools
from collections.abc import Callable
from typing import Any

from sqlalchemy.orm import Session

from app import permissions, timeutil
from app.jsutil import collate_default, collate_ja, js_sort_key, js_string, js_trim, normalize_gas
from app.rpc import RpcContext, RpcError

from . import adapters, calendar
from .jsdate import js_number

# ===== 作業員マスタ（loadWorkersByDeptMap_） =====


def _header_index(header: list[str]) -> dict[str, int]:
    def find(name: str) -> int:
        return header.index(name) if name in header else -1

    idx = {
        "code": find("作業員コード"),
        "name": find("氏名"),
        "dept": find("部署"),
        "active": find("在籍フラグ"),
        "email": next((i for i, h in enumerate(header) if "Googleアカウント" in h), -1),
    }
    if idx["code"] < 0:
        idx["code"] = find("作業員ID")
    if idx["dept"] < 0:
        idx["dept"] = find("部署ID")
    return idx


def _cell(row: list[Any], i: int) -> Any:
    return row[i] if 0 <= i < len(row) else None


def _active(row: list[Any], idx: dict[str, int]) -> bool:
    if idx["active"] < 0:
        return True
    active = normalize_gas(_cell(row, idx["active"]))
    return not active or active == "1" or active.lower() == "true"


def workers_by_dept(db: Session) -> dict[str, list[dict]]:
    """部署 → [{workerId, workerName, deptId}]（部署が最初に出た順。部署の中は workerId の localeCompare）。"""
    values = adapters.worker_sheet(db)
    if len(values) < 2:
        return {}
    header = [normalize_gas(h) for h in values[0]]
    idx = _header_index(header)
    if idx["code"] < 0 or idx["name"] < 0 or idx["dept"] < 0:
        raise RpcError("作業員マスタのヘッダが想定と違います（作業員コード/氏名/部署）。実際: " + ", ".join(header))
    out: dict[str, list[dict]] = {}
    for row in values[1:]:
        dept = normalize_gas(_cell(row, idx["dept"]))
        code = normalize_gas(_cell(row, idx["code"]))
        name = normalize_gas(_cell(row, idx["name"]))
        if not dept or not code or not name:
            continue
        if not _active(row, idx):
            continue
        out.setdefault(dept, []).append({"workerId": code, "workerName": name, "deptId": dept})
    for workers in out.values():
        workers.sort(key=functools.cmp_to_key(lambda a, b: collate_default(a["workerId"], b["workerId"])))
    return out


def dept_list(db: Session) -> list[dict]:
    return [{"deptId": d, "deptName": d} for d in sorted(workers_by_dept(db), key=js_sort_key)]


def dept_names(db: Session) -> list[str]:
    return [d["deptName"] for d in dept_list(db)]


def api_getDeptList(ctx: RpcContext):
    return dept_list(ctx.db)


def api_getWorkersByDept(ctx: RpcContext, dept_id=None):
    if not dept_id:
        return []
    if not isinstance(dept_id, str):
        return []  # Map のキーは文字なので、数などの引数は当たらない
    return workers_by_dept(ctx.db).get(dept_id, [])


def api_getWorkerInfo(ctx: RpcContext):
    email = (ctx.user.email or "").lower()
    if not email:
        return None
    values = adapters.worker_sheet(ctx.db)
    if not values:
        return None
    header = [normalize_gas(h) for h in values[0]]
    idx = _header_index(header)
    if idx["email"] < 0:
        return None
    for row in values[1:]:
        if not _active(row, idx):
            continue
        if normalize_gas(_cell(row, idx["email"])).lower() == email:
            dept = normalize_gas(_cell(row, idx["dept"]))
            return {"workerId": normalize_gas(_cell(row, idx["code"])), "workerName": normalize_gas(_cell(row, idx["name"])),
                    "deptId": dept, "deptName": dept, "email": email}
    return None


# ===== LOOKUP（getLookupValues_） =====


def lookup_values(db: Session, category: str) -> list[str]:
    out = []
    for sheet_row, cat, value, sort_text in adapters.lookup_rows(db):
        if normalize_gas(cat) == category:
            out.append((normalize_gas(value), js_number(sort_text or 0)))

    def cmp(a, b) -> int:
        d = a[1] - b[1]
        return 0 if d != d else (d > 0) - (d < 0)  # NaN の比べは 0（V8 と同じ）

    out.sort(key=functools.cmp_to_key(cmp))
    return [v for v, _ in out]


def api_getLookupData(ctx: RpcContext):
    return {
        "leaveKubun": ["全日休", "半日休"],
        "halfDayType": ["午前", "午後"],
        "leaveTypeFull": ["有給消化（私用）", "振替休暇", "特別休暇"],
        "leaveTypeHalf": ["有給消化（私用）", "特別休暇"],
        "specialReasons": lookup_values(ctx.db, "特別休暇理由"),
    }


# ===== カレンダー =====


def api_getWorkingDays(ctx: RpcContext):
    return calendar.working_days(ctx.db, _today(ctx))


def api_getHolidayDays(ctx: RpcContext, worker_id=None):
    return calendar.holiday_days(ctx.db, _today(ctx), worker_id)


def api_getPastWorkingDays(ctx: RpcContext, year_month=None):
    return calendar.past_working_days(ctx.db, _today(ctx), year_month)


def _today(ctx: RpcContext):
    return timeutil.jst_date_of(ctx.now)


# ===== 電子印（M_STAMP） =====


def stamp_name(db: Session, email: str) -> str | None:
    """M_STAMP の A 列（normalize_＋小文字）が email（小文字）と一致した最初の行の C 列（normalize_）。無ければ None。"""
    target = email.lower()
    for a, _b, c in adapters.stamp_rows(db):
        if normalize_gas(js_string(a or "")).lower() == target:
            return normalize_gas(js_string(c or ""))
    return None


def api_getApproverList(ctx: RpcContext):
    try:
        out = []
        for a, _b, c in adapters.stamp_rows(ctx.db):
            email = js_trim(js_string(a or ""))
            if email:
                out.append({"email": email, "name": js_trim(js_string(c or ""))})
        return out
    except Exception:  # noqa: BLE001 - GAS も例外のときは [] を返す
        return []


def api_getCurrentApproverInfo(ctx: RpcContext):
    email = ctx.user.email.lower()
    name = email
    try:
        found = stamp_name(ctx.db, email)
        if found is not None:
            name = found or email
    except Exception:  # noqa: BLE001 - GAS も握る
        pass
    return {"email": email, "name": name}


# ===== 権限（auth.gs） =====


def api_isAdmin(ctx: RpcContext):
    return permissions.is_admin(ctx.db, ctx.user.email)


def api_isSomu(ctx: RpcContext):
    return permissions.is_somu(ctx.db, ctx.user.email)


def api_isDeptApprover(ctx: RpcContext):
    try:
        return permissions.is_dept_approver(ctx.db, ctx.user.email)
    except Exception:  # noqa: BLE001 - GAS も false
        return False


def approver_profiles(db: Session) -> list[dict]:
    """M_DEPT_APPROVERS の承認者 [{email, name, depts}]（メールが最初に出た順。api_getApproverProfiles の profileMap）。"""
    profiles: dict[str, dict] = {}
    for a, b, c in adapters.dept_approver_rows(db):
        dept = js_trim(normalize_gas(js_string(a or "")))
        emails = normalize_gas(js_string(b or ""))
        name = normalize_gas(js_string(c or ""))
        if not dept or not emails:
            continue
        for e in [s.strip().lower() for s in emails.split(",")]:
            if not e:
                continue
            p = profiles.setdefault(e, {"email": e, "name": name or e, "depts": []})
            if dept not in p["depts"]:
                p["depts"].append(dept)
            if name and p["name"] == e:
                p["name"] = name
    return list(profiles.values())


def api_getApproverProfiles(ctx: RpcContext):
    email = ctx.user.email.lower()
    admin = permissions.is_admin(ctx.db, email)
    plist = approver_profiles(ctx.db)
    if admin:
        plist.sort(key=functools.cmp_to_key(lambda x, y: collate_ja(x["name"], y["name"])))
        return {"isAdmin": True, "currentEmail": email, "profiles": plist}
    mine = next((p for p in plist if p["email"] == email), None)
    return {"isAdmin": False, "currentEmail": email, "profiles": [mine] if mine else []}


HANDLERS: dict[str, Callable] = {
    "api_isDeptApprover": api_isDeptApprover,
    "api_isAdmin": api_isAdmin,
    "api_isSomu": api_isSomu,
    "api_getDeptList": api_getDeptList,
    "api_getWorkerInfo": api_getWorkerInfo,
    "api_getLookupData": api_getLookupData,
    "api_getWorkingDays": api_getWorkingDays,
    "api_getWorkersByDept": api_getWorkersByDept,
    "api_getHolidayDays": api_getHolidayDays,
    "api_getPastWorkingDays": api_getPastWorkingDays,
    "api_getApproverList": api_getApproverList,
    "api_getCurrentApproverInfo": api_getCurrentApproverInfo,
    "api_getApproverProfiles": api_getApproverProfiles,
}
