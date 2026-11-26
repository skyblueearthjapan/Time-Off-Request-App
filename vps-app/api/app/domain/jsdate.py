"""GAS（JavaScript の V8。スクリプトの時間帯は Asia/Tokyo）と同じ結果を出す小道具（日付・真偽・数）。

- js_truthy(v): JS の真偽（''・0・None・False・NaN が偽。空の配列・辞書は真）。
- js_number(v): Number(v)（'' と None は 0、読めなければ NaN）。
- js_new_date(v): new Date(v)。読めなければ None（Invalid Date）。V8 の読み方:
    'yyyy-MM-dd' は UTC の 0 時、'yyyy-MM-ddTHH:mm(:ss)' は時差が無ければ日本時間、'yyyy/M/d'・'M/d/yyyy'・空白入りは
    日本時間。日は 1〜31（4/31 は 5/1 に繰り越す）。
- js_local(y, m0, d, h, mi, s): new Date(y, m0, d, h, mi, s)（日本時間。月・日のはみ出しは繰り越す。0〜99 年は 1900 年代）。
- js_date_string(moment): String(Date)（'Fri Nov 27 2026 09:00:00 GMT+0900 (Japan Standard Time)'）。
- out_number(x): JSON にしたとき JS と同じ見た目（整数の値は int、NaN は None＝JSON の null）。
"""
from __future__ import annotations

import math
import re
from datetime import date, datetime, timedelta, timezone
from typing import Any

from app import timeutil
from app.jsutil import js_trim

JST = timeutil.JST
UTC = timeutil.UTC
EPOCH = datetime(1970, 1, 1, tzinfo=UTC)
DAY_NAMES = ("日", "月", "火", "水", "木", "金", "土")
_EN_DAYS = ("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
_EN_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

YMD_RE = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}\Z")  # JS の /^\d{4}-\d{2}-\d{2}$/ と同じ


def js_truthy(value: Any) -> bool:
    if value is None or value is False:
        return False
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return not (value == 0 or (isinstance(value, float) and math.isnan(value)))
    if isinstance(value, str):
        return value != ""
    return True


_NUM_RE = re.compile(r"^[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?\Z")


def js_number(value: Any) -> float:
    """Number(value)。"""
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, list):
        if not value:
            return 0.0
        return js_number(value[0]) if len(value) == 1 else math.nan
    if not isinstance(value, str):
        return math.nan
    text = js_trim(value)
    if text == "":
        return 0.0
    if text in ("Infinity", "+Infinity"):
        return math.inf
    if text == "-Infinity":
        return -math.inf
    low = text.lower()
    for prefix, base in (("0x", 16), ("0o", 8), ("0b", 2)):
        if low.startswith(prefix):
            try:
                return float(int(text[2:], base)) if text[2:].isalnum() else math.nan
            except ValueError:
                return math.nan
    if _NUM_RE.match(text):
        return float(text)
    return math.nan


def out_number(value: float) -> int | float | None:
    """JSON にしたとき JS と同じ見た目にする（1.0 → 1、NaN → null）。"""
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        if value.is_integer():
            return int(value)
    return value


# ===== new Date(...) =====

_ISO_DATE = re.compile(r"^([0-9]{4})(?:-([0-9]{2})(?:-([0-9]{2}))?)?\Z")
_ISO_DATETIME = re.compile(
    r"^([0-9]{4})-([0-9]{2})-([0-9]{2})T([0-9]{2}):([0-9]{2})(?::([0-9]{2})(?:\.([0-9]{1,3}))?)?"
    r"(Z|[+-][0-9]{2}:[0-9]{2})?\Z"
)
_LEGACY_YMD = re.compile(
    r"^([0-9]{4})[/-]([0-9]{1,2})[/-]([0-9]{1,2})(?:[ T]+([0-9]{1,2}):([0-9]{2})(?::([0-9]{2}))?)?\Z"
)
_LEGACY_MDY = re.compile(r"^([0-9]{1,2})/([0-9]{1,2})/([0-9]{4})(?:[ ]+([0-9]{1,2}):([0-9]{2})(?::([0-9]{2}))?)?\Z")


def _build(y: int, mo: int, d: int, h: int, mi: int, s: int, ms: int, tz) -> datetime | None:
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return None
    if h == 24:
        if mi or s or ms:
            return None
    elif not (0 <= h <= 23):
        return None
    if not (0 <= mi <= 59 and 0 <= s <= 59):
        return None
    try:
        base = datetime(y, mo, 1, tzinfo=tz) + timedelta(days=d - 1)
        return base + timedelta(hours=h, minutes=mi, seconds=s, milliseconds=ms)
    except (ValueError, OverflowError):
        return None


def js_new_date(value: Any) -> datetime | None:
    """new Date(value)。Invalid Date は None。返す datetime は時差付き。"""
    if value is None:
        return EPOCH
    if isinstance(value, bool):
        return EPOCH + timedelta(milliseconds=1 if value else 0)
    if isinstance(value, datetime):
        return value if value.tzinfo else value.replace(tzinfo=UTC)
    if isinstance(value, date):
        return timeutil.jst_midnight(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return EPOCH + timedelta(milliseconds=int(value))
    if not isinstance(value, str):
        return None
    text = value.strip()
    m = _ISO_DATE.match(text)
    if m:
        return _build(int(m.group(1)), int(m.group(2) or 1), int(m.group(3) or 1), 0, 0, 0, 0, UTC)
    m = _ISO_DATETIME.match(text)
    if m:
        y, mo, d, h, mi = (int(m.group(i)) for i in range(1, 6))
        s = int(m.group(6) or 0)
        ms = int((m.group(7) or "0").ljust(3, "0"))
        off = m.group(8)
        if off is None:
            tz = JST
        elif off == "Z":
            tz = UTC
        else:
            sign = 1 if off[0] == "+" else -1
            tz = timezone(sign * timedelta(hours=int(off[1:3]), minutes=int(off[4:6])))
        return _build(y, mo, d, h, mi, s, ms, tz)
    m = _LEGACY_YMD.match(text)
    if m:
        h, mi, s = int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0)
        return _build(int(m.group(1)), int(m.group(2)), int(m.group(3)), h, mi, s, 0, JST)
    m = _LEGACY_MDY.match(text)
    if m:
        h, mi, s = int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0)
        return _build(int(m.group(3)), int(m.group(1)), int(m.group(2)), h, mi, s, 0, JST)
    return None


def jst_ymd_date(ymd: str) -> date | None:
    """new Date(ymd + 'T00:00:00+09:00') の日本時間の日付（4/31 は 5/1。読めなければ None）。"""
    moment = js_new_date(f"{ymd}T00:00:00+09:00")
    return moment.astimezone(JST).date() if moment is not None else None


def js_local(y: float, m0: float, d: float, h: int = 0, mi: int = 0, s: int = 0) -> datetime | None:
    """new Date(y, m0, d, h, mi, s)（日本時間）。NaN の年は None（Invalid Date）。"""
    if isinstance(y, float) and (math.isnan(y) or math.isinf(y)):
        return None
    yi, mi0, di = int(y), int(m0), int(d)
    if 0 <= yi <= 99:
        yi += 1900
    total = yi * 12 + mi0
    year, month0 = divmod(total, 12)
    try:
        base = datetime(year, month0 + 1, 1, tzinfo=JST)
        return base + timedelta(days=di - 1, hours=h, minutes=mi, seconds=s)
    except (ValueError, OverflowError):
        return None


def js_local_date(y: int, m0: int, d: int) -> date:
    """new Date(y, m0, d) の日付（月・日のはみ出しを繰り越す）。"""
    moment = js_local(y, m0, d)
    if moment is None:  # pragma: no cover - 整数の年は必ず読める
        raise ValueError((y, m0, d))
    return moment.date()


def js_date_string(moment: datetime) -> str:
    """String(Date)（V8・時間帯 Asia/Tokyo）。"""
    t = moment.astimezone(JST)
    dow = _EN_DAYS[(t.weekday() + 1) % 7]
    return (f"{dow} {_EN_MONTHS[t.month - 1]} {t.day:02d} {t.year:04d} "
            f"{t.hour:02d}:{t.minute:02d}:{t.second:02d} GMT+0900 (Japan Standard Time)")


def jst_dow(d: date) -> int:
    """getDay()（日曜 0〜土曜 6）。"""
    return (d.weekday() + 1) % 7
