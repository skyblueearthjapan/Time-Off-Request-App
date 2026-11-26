"""年度・月次期間・過去申請の印（実装仕様 8-2。日付はすべて日本時間）。

- fiscal_year(d): computeFiscalYear_（leaveRequest.gs:34）。3/16 始まり。
- fy_start(d): getFiscalYearStartDate_（config.gs:184）。
- monthly_period(d): computeMonthlyPeriod_（leaveRequest.gs:50）。16 日始まり。
- is_within_current_fy(s, today): isWithinCurrentFiscalYear_（config.gs:193）。文字で比べる。
- is_retroactive(applied_at, leave_date_iso): computeIsRetroactive_（config.gs:206）。
- days_ago(leave_date_iso, today): computeDaysAgo_（config.gs:219）。
"""
from __future__ import annotations

import math
import re
from datetime import date, datetime
from typing import Any

from app import timeutil
from app.jsutil import js_string

from .jsdate import js_local_date, js_new_date

_YMD = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}\Z")  # JS の /^\d{4}-\d{2}-\d{2}$/ と同じ（半角の数字だけ、改行を許さない）


def _as_date(d: date | datetime) -> date:
    return timeutil.jst_date_of(d) if isinstance(d, datetime) else d


def fiscal_year(d: date | datetime) -> int:
    day = _as_date(d)
    if day.month < 3 or (day.month == 3 and day.day < 16):
        return day.year - 1
    return day.year


def fy_start(d: date | datetime) -> date:
    return date(fiscal_year(d), 3, 16)


def monthly_period(d: date | datetime) -> tuple[date, date]:
    day = _as_date(d)
    y, m0 = day.year, day.month - 1
    if day.day >= 16:
        return js_local_date(y, m0, 16), js_local_date(y, m0 + 1, 15)
    return js_local_date(y, m0 - 1, 16), js_local_date(y, m0, 15)


def is_within_current_fy(s: Any, today: date) -> bool:
    if not s or not isinstance(s, str) or not _YMD.match(s):
        return False
    return timeutil.to_jst_text(fy_start(today), "yyyy-MM-dd") <= s <= timeutil.to_jst_text(today, "yyyy-MM-dd")


def is_retroactive(applied_at: Any, leave_date_iso: Any) -> bool:
    """申請日時（Date のセルは datetime、ほかは文字）の日本時間の日付 > 休暇日の文字の先頭 10 字。"""
    if not applied_at or not leave_date_iso:
        return False
    applied = applied_at if isinstance(applied_at, datetime) else js_new_date(applied_at)
    if applied is None:
        return False
    if isinstance(leave_date_iso, (date, datetime)):
        ld = timeutil.to_jst_text(leave_date_iso, "yyyy-MM-dd")
    else:
        ld = js_string(leave_date_iso)[:10]
    return timeutil.to_jst_text(applied, "yyyy-MM-dd") > ld


def days_ago(leave_date_iso: Any, today: date) -> int:
    if not leave_date_iso:
        return 0
    if isinstance(leave_date_iso, (date, datetime)):
        ld_str = timeutil.to_jst_text(leave_date_iso, "yyyy-MM-dd")
    else:
        ld_str = js_string(leave_date_iso)[:10]
    if not _YMD.match(ld_str):
        return 0
    ld = js_new_date(f"{ld_str}T00:00:00+09:00")
    if ld is None:
        return 0
    t0 = timeutil.jst_midnight(today)
    diff_days = math.floor((t0 - ld).total_seconds() / 86400)
    return diff_days if diff_days > 0 else 0
