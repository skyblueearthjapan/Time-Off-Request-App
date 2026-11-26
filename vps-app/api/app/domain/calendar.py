"""勤務日・休日・祝日（実装仕様 8-3。masters.gs:220-560 の写し。日付はすべて日本時間）。

- japanese_holidays(today): getJapanesePublicHolidays_。今日の年度 fy と fy+1 の 2 年だけ。振替休日・国民の休日つき。
- working_days(db, today): api_getWorkingDays。今日から年度末 3/15 まで。
- holiday_days(db, today, worker_id): api_getHolidayDays。年度開始 3/16 から今日＋7 日まで。
- past_working_days(db, today, year_month): api_getPastWorkingDays。
カレンダー（M_CALENDAR）の同じ日の 2 行は下の行が勝つ。
"""
from __future__ import annotations

import math
import re
from datetime import date, timedelta
from typing import Any

from sqlalchemy.orm import Session

from app.constants import KIND_SUBSTITUTE, STATUS_CANCELED
from app.jsutil import js_string, normalize_gas

from . import adapters, store
from .fiscal import fiscal_year, fy_start
from .jsdate import js_local_date, js_truthy, jst_dow

DOW_NAMES = ("日", "月", "火", "水", "木", "金", "土")


def _ymd(d: date) -> str:
    return d.isoformat()


def holidays_for_year(y: int) -> dict[str, str]:
    """getHolidaysForYear_: 固定・ハッピーマンデー・春分・秋分。"""
    h: dict[str, str] = {}

    def add(m: int, d: int, name: str) -> None:
        h[f"{y}-{m:02d}-{d:02d}"] = name

    def nth_monday(month: int, n: int) -> int:
        d = date(y, month, 1)
        count = 0
        while count < n:
            if jst_dow(d) == 1:
                count += 1
            if count < n:
                d += timedelta(days=1)
        return d.day

    add(1, 1, "元日")
    add(2, 11, "建国記念の日")
    add(2, 23, "天皇誕生日")
    add(4, 29, "昭和の日")
    add(5, 3, "憲法記念日")
    add(5, 4, "みどりの日")
    add(5, 5, "こどもの日")
    add(8, 11, "山の日")
    add(11, 3, "文化の日")
    add(11, 23, "勤労感謝の日")
    add(1, nth_monday(1, 2), "成人の日")
    add(7, nth_monday(7, 3), "海の日")
    add(9, nth_monday(9, 3), "敬老の日")
    add(10, nth_monday(10, 2), "スポーツの日")
    shunbun = math.floor(20.8431 + 0.242194 * (y - 1980) - math.floor((y - 1980) / 4))
    add(3, shunbun, "春分の日")
    shubun = math.floor(23.2488 + 0.242194 * (y - 1980) - math.floor((y - 1980) / 4))
    add(9, shubun, "秋分の日")
    return h


def japanese_holidays(today: date) -> dict[str, str]:
    """getJapanesePublicHolidays_（日付の文字 → 祝日名）。"""
    fy = fiscal_year(today)
    holidays: dict[str, str] = {}
    for y in (fy, fy + 1):
        holidays.update(holidays_for_year(y))

    keys = sorted(holidays)
    # 振替休日: 日曜の祝日ごとに、翌日から祝日でない日まで進める（その時点の祝日を見る）
    for key in keys:
        dt = date.fromisoformat(key)
        if jst_dow(dt) == 0:
            sub = dt + timedelta(days=1)
            while _ymd(sub) in holidays:
                sub += timedelta(days=1)
            holidays[_ymd(sub)] = "振替休日"
    # 国民の休日: 振替休日を足す前の並びで、隣どうしの差がちょうど 2 日、間の日が祝日でも土日でもない
    for i in range(len(keys) - 1):
        d1, d2 = date.fromisoformat(keys[i]), date.fromisoformat(keys[i + 1])
        if (d2 - d1).days == 2:
            between = d1 + timedelta(days=1)
            if _ymd(between) not in holidays and jst_dow(between) not in (0, 6):
                holidays[_ymd(between)] = "国民の休日"
    return holidays


def calendar_map(db: Session) -> dict[str, str]:
    """M_CALENDAR の 日付 → 区分（読めない行は飛ばす。同じ日は下の行が勝つ）。"""
    out: dict[str, str] = {}
    for day, kubun in adapters.calendar_entries(db):
        if day is None:
            continue
        out[_ymd(day)] = kubun
    return out


def _is_work_day(d: date, kubun: str, public_holidays: dict[str, str]) -> bool:
    if kubun in ("休日", "祝日"):
        return False
    if kubun == "出勤土曜":
        return True
    return 1 <= jst_dow(d) <= 5 and _ymd(d) not in public_holidays


def working_days(db: Session, today: date) -> list[str]:
    entries = adapters.calendar_entries(db)
    if not entries:
        return []
    cal = calendar_map(db)
    public = japanese_holidays(today)
    end = date(fiscal_year(today) + 1, 3, 15)
    out: list[str] = []
    cursor = today
    while cursor <= end:
        if _is_work_day(cursor, cal.get(_ymd(cursor), ""), public):
            out.append(_ymd(cursor))
        cursor += timedelta(days=1)
    return out


def used_substitute_dates(db: Session, worker_id: Any, exclude_req_id: str | None = None) -> set[str]:
    """同じ作業員の振替休暇（取消でない。却下も使用済み）の振替元出勤日（'yyyy-MM-dd'）。"""
    target = normalize_gas(worker_id)
    used: set[str] = set()
    for row in store.rows_in_sheet_order(db):
        if exclude_req_id and normalize_gas(row.req_id) == normalize_gas(exclude_req_id):
            continue
        if store.norm(row, "leave_type") != KIND_SUBSTITUTE:
            continue
        if store.norm(row, "status") == STATUS_CANCELED:
            continue
        if store.norm(row, "worker_id") != target:
            continue
        moment = store.cell_moment(store.raw_cell(row, "substitute_date"))
        if moment is None:
            continue
        used.add(store.fmt(moment))
    return used


def holiday_days(db: Session, today: date, worker_id: Any = None) -> list[dict]:
    cal = calendar_map(db)
    public = japanese_holidays(today)
    start = fy_start(today)
    end = today + timedelta(days=7)
    result: list[dict] = []
    cursor = start
    while cursor <= end:
        key = _ymd(cursor)
        dow = jst_dow(cursor)
        if key in cal:
            kubun = cal[key]
            if kubun in ("休日", "祝日"):
                result.append({"date": key, "label": kubun})
            elif kubun == "出勤土曜":
                pass
            elif dow in (0, 6):
                result.append({"date": key, "label": "日曜" if dow == 0 else "土曜"})
        elif key in public:
            result.append({"date": key, "label": f"祝日（{public[key]}）"})
        elif dow == 0:
            result.append({"date": key, "label": "日曜"})
        elif dow == 6:
            result.append({"date": key, "label": "土曜"})
        cursor += timedelta(days=1)
    if js_truthy(worker_id):
        used = used_substitute_dates(db, worker_id)
        result = [c for c in result if c["date"] not in used]
    return result


# JS の \d・$ と同じ（Python の \d は全角などの数字も含み、$ は末尾の改行の前でも合う）
_YM_DASH = re.compile(r"^[0-9]{4}-[0-9]{2}\Z")
_YM_PLAIN = re.compile(r"^[0-9]{6}\Z")


def past_working_days(db: Session, today: date, year_month: Any = None) -> list[dict]:
    ym = normalize_gas("" if year_month is None else js_string(year_month))
    if _YM_DASH.match(ym):
        y, m = int(ym[0:4]), int(ym[5:7])
    elif _YM_PLAIN.match(ym):
        y, m = int(ym[0:4]), int(ym[4:6])
    else:
        y, m = today.year, today.month
    month_start = js_local_date(y, m - 1, 1)
    month_end = js_local_date(y, m, 0)
    range_start = max(month_start, fy_start(today))
    range_end = min(month_end, today)
    if range_start > range_end:
        return []
    cal = calendar_map(db)
    public = japanese_holidays(today)
    out: list[dict] = []
    cursor = range_start
    while cursor <= range_end:
        key = _ymd(cursor)
        dow = jst_dow(cursor)
        if _is_work_day(cursor, cal.get(key, ""), public):
            out.append({"date": key, "dow": dow, "label": f"{key}（{DOW_NAMES[dow]}）"})
        cursor += timedelta(days=1)
    return out
