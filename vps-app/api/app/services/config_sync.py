"""設定とマスタのタブの取り込み（実装仕様 6-5・5-2。`sync_sheet_config`、SHEET_CONFIG_SYNC_MINUTES ごとと手動。担当 D）。

DB スプレッドシートの 6 タブを読み、DB の表を全置換する。タブには書かない（読むだけ）。

| タブ | 表 | 読み方 |
|---|---|---|
| M_SYSTEM_SETTING | settings | 2 行目から A=キー（String(v||'').trim()）・B=値（生の文字）。同じキーは下の行が勝つ（値が空でも勝つ。config.gs:107-119） |
| M_LOOKUP | lookup | 見出し（カテゴリ・値・SORT）の文字で列を引く |
| 作業員マスタ | workers | 見出しと値の組の配列（cells＝[[見出し, 値], …]。シートの列の順のまま。同じ見出しも残す） |
| M_CALENDAR | calendar | 見出し（日付・区分・曜日・備考）の文字で列を引く（同じ見出しは前の列）。日付はセルの瞬間の JST の日付。cells は workers と同じ組の配列 |
| M_STAMP | stamps | 位置（A メール・B stampFileId・C 備考） |
| M_DEPT_APPROVERS | dept_approvers | 位置（A〜E） |

- タブごとに、読めない・空・有効な行 0 なら上書きしない（今の表を残す）。
- Settings の PDF_FOLDER_ID・LEDGER_SSID・MASTER_SOURCE_SSID・CALENDAR_SOURCE_SSID が .env の値と違えば mismatches に出す
  （MASTER・CALENDAR は空なら GAS がコードの既定＝同じ ID を使うので食い違いにしない）。LEDGER_SSID の行が 2 つ以上なら警告。
- 戻り値 {ok, counts, overlaps, mismatches, skipped, warnings}。sync_runs（kind='sync_sheet_config'）に残す。commit は呼ぶ側。

セルの値は GAS の getValues が返す形に合わせる: 文字はそのまま、数は数、日付のセル（数で表示が日付）は JST の
ISO の文字（'yyyy-MM-ddTHH:mm:ss+09:00'）。文字にするときは JS の String(値)（日付のセルは表示の文字）。

表の列の名前は実装仕様 5-2（マスタは sort_order＝データ行の順で 1 から。見出しの行は持たない）。
"""
from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Any

from sqlalchemy import delete
from sqlalchemy.orm import Session

from app import timeutil
from app.constants import (
    TAB_CALENDAR,
    TAB_DEPT_APPROVERS,
    TAB_LOOKUP,
    TAB_SETTINGS,
    TAB_STAMP,
    TAB_WORKER,
)
from app.jsutil import js_string, js_trim, normalize_gas
from app.models import CalendarDay, DeptApprover, Lookup, Setting, Stamp, SyncRun, Worker
from app.services import sheet_codec as codec
from app.services.sheets_api import SheetsApi, quote_sheet, sheet_titles

log = logging.getLogger(__name__)

RUN_KIND = "sync_sheet_config"
TABS = (TAB_SETTINGS, TAB_LOOKUP, TAB_WORKER, TAB_CALENDAR, TAB_STAMP, TAB_DEPT_APPROVERS)


# ===== セルの読み方 =====


def _cell(row: list, i: int) -> Any:
    return row[i] if 0 <= i < len(row) else ""


def _blank_js(v: Any) -> bool:
    """`String(v || '').trim() === ''`（0・false・null・'' は空）。"""
    if v is None or v is False or v == "" or (codec.is_number(v) and v == 0):
        return True
    return js_trim(js_string(v)) == ""


def _is_date_cell(u: Any, f: Any) -> bool:
    return codec.is_number(u) and codec.looks_like_date_display(f)


def gas_value(u: Any, f: Any, tz: Any) -> Any:
    """getValues の値（日付のセルは JST の ISO の文字）。"""
    if u is None:
        return ""
    if _is_date_cell(u, f):
        return codec.serial_to_datetime(u, tz).astimezone(codec.JST).isoformat(timespec="seconds")
    return u


def gas_text(u: Any, f: Any) -> str:
    """String(値)。日付のセルは表示の文字。"""
    if _is_date_cell(u, f):
        return js_string(f)
    return js_string("" if u is None else u)


def _header_col(header: list, *labels: str) -> int:
    for label in labels:
        for i, h in enumerate(header):
            if js_trim(js_string(h)) == label:
                return i
    return -1


def _rows(u: list, f: list):
    """データ行（2 行目から）: (行番号, UNFORMATTED の行, FORMATTED の行)。"""
    for i in range(1, len(u)):
        yield i + 1, list(u[i]), list(f[i]) if i < len(f) else []


def _has_any(row: list) -> bool:
    return any(not codec.is_blank(v) for v in row)


# ===== タブごとの読み方（移行の道具・マスタ同期も使う） =====


def parse_settings(u: list, f: list | None = None) -> tuple[dict[str, tuple[str, int]], list[dict]]:
    """M_SYSTEM_SETTING → ({key: (value, sheet_row)}, 重なったキー)。下の行が勝つ（空の値でも）。"""
    f = f or []
    out: dict[str, tuple[str, int]] = {}
    rows_by_key: dict[str, list[int]] = {}
    for r, urow, frow in _rows(u, f):
        k = _cell(urow, 0)
        key = "" if _blank_js(k) else js_trim(js_string(k))
        if not key:
            continue
        out[key] = (gas_text(_cell(urow, 1), _cell(frow, 1)), r)
        rows_by_key.setdefault(key, []).append(r)
    overlaps = [{"key": k, "rows": rows, "used": rows[-1]} for k, rows in rows_by_key.items() if len(rows) > 1]
    return out, overlaps


def parse_lookup(u: list, f: list | None = None) -> list[dict]:
    f = f or []
    header = list(u[0]) if u else []
    cc, cv, cs = _header_col(header, "カテゴリ"), _header_col(header, "値"), _header_col(header, "SORT")
    out = []
    for r, urow, frow in _rows(u, f):
        if not _has_any(urow):
            continue
        out.append({"sheet_row": r, "category": gas_text(_cell(urow, cc), _cell(frow, cc)),
                    "value": gas_text(_cell(urow, cv), _cell(frow, cv)),
                    "sort_text": gas_text(_cell(urow, cs), _cell(frow, cs)),
                    "cells": [gas_text(v, _cell(frow, i)) for i, v in enumerate(urow)]})
    return out


def _width(u: list) -> int:
    """getDataRange の列の数（いちばん長い行。Sheets API は行の末尾の空のセルを返さない）。"""
    return max((len(r) for r in u), default=0)


def header_pairs(header: list, urow: list, frow: list, width: int, tz: Any) -> list[list]:
    """1 行を [[見出し, 値], …]（左の列から順。見出しの無い列・同じ見出しの列も残す）にする。

    PostgreSQL の JSONB はオブジェクトのキーの順を保たず、同じキーを 1 つにまとめる。シートの列の並び
    （GAS の getValues の並び）を保つため、dict ではなく配列で持つ。
    """
    return [[js_trim(js_string(_cell(header, i))), gas_value(_cell(urow, i), _cell(frow, i), tz)]
            for i in range(width)]


def parse_workers(u: list, f: list | None, tz: Any) -> list[dict]:
    f = f or []
    header = [js_trim(js_string(h)) for h in (u[0] if u else [])]
    width = _width(u)

    def at(urow, frow, *labels):
        return gas_text(_cell(urow, _header_col(header, *labels)), _cell(frow, _header_col(header, *labels)))

    out = []
    for r, urow, frow in _rows(u, f):
        if not _has_any(urow):
            continue
        out.append({"sort_order": len(out) + 1, "sheet_row": r, "cells": header_pairs(header, urow, frow, width, tz),
                    "worker_code": at(urow, frow, "作業員コード", "作業員ID")})
    return out


def calendar_date(u: Any, f: Any, tz: Any):
    """M_CALENDAR の日付のセル → JST の日付（読めなければ None）。

    Date のセルはその瞬間の JST の日付。文字は JS の new Date(文字) と同じ読み（'yyyy-MM-dd' は UTC の 0 時＝JST で
    同じ日、'M/d/yyyy' は JST の 0 時）。
    """
    if codec.is_number(u):
        return codec.serial_to_jst_date(u, tz)
    return codec.parse_date_text(js_string(u)) if not codec.is_blank(u) else None


def parse_calendar(u: list, f: list | None, tz: Any) -> list[dict]:
    f = f or []
    header = list(u[0]) if u else []
    cd, ck = _header_col(header, "日付"), _header_col(header, "区分")
    width = _width(u)

    out = []
    for r, urow, frow in _rows(u, f):
        if not _has_any(urow):
            continue
        day = calendar_date(_cell(urow, cd), _cell(frow, cd), tz) if cd >= 0 else None
        out.append({"sort_order": len(out) + 1, "sheet_row": r, "date": day,
                    "kubun": normalize_gas(gas_text(_cell(urow, ck), _cell(frow, ck))) if ck >= 0 else "",
                    "cells": header_pairs(header, urow, frow, width, tz)})
    return out


def parse_positional(u: list, f: list | None, width: int, tz: Any) -> list[dict]:
    """位置で読むタブ（M_STAMP は 3 列、M_DEPT_APPROVERS は 5 列）→ col_a…（生の文字）。"""
    f = f or []
    names = ["col_a", "col_b", "col_c", "col_d", "col_e"][:width]
    out = []
    for r, urow, frow in _rows(u, f):
        if not _has_any(urow):
            continue
        rec = {"sort_order": len(out) + 1, "sheet_row": r, "cells": [gas_value(v, _cell(frow, i), tz) for i, v in enumerate(urow)]}
        for i, n in enumerate(names):
            rec[n] = gas_text(_cell(urow, i), _cell(frow, i))
        out.append(rec)
    return out


def valid_rows(rows: list[dict], key: str) -> int:
    """有効な行（1 列目・キーの列が空でない）の数。"""
    return sum(1 for r in rows if js_trim(js_string(r.get(key, ""))) != "")


def _fit(model: Any, record: dict) -> dict:
    """表に無いキー（読み方の手がかりの sheet_row・worker_code など）を捨てる。"""
    cols = {c.name for c in model.__table__.columns}
    return {k: v for k, v in record.items() if k in cols}


# 表 → (モデル, 読み方, 有効な行のキー)
def _parsers(tz: Any) -> dict[str, tuple[Any, Callable[[list, list], list[dict]], str]]:
    return {
        TAB_LOOKUP: (Lookup, lambda u, f: parse_lookup(u, f), "category"),
        TAB_WORKER: (Worker, lambda u, f: parse_workers(u, f, tz), "worker_code"),
        TAB_CALENDAR: (CalendarDay, lambda u, f: parse_calendar(u, f, tz), "date"),
        TAB_STAMP: (Stamp, lambda u, f: parse_positional(u, f, 3, tz), "col_a"),
        TAB_DEPT_APPROVERS: (DeptApprover, lambda u, f: parse_positional(u, f, 5, tz), "col_a"),
    }


def build_rows(tab: str, u: list, f: list, tz: Any) -> tuple[list[dict], int]:
    """タブの値 → (表の行（仕様の名前）, 有効な行の数)。"""
    if tab == TAB_SETTINGS:
        parsed, _ = parse_settings(u, f)
        return [{"key": k, "value": v, "sheet_row": r} for k, (v, r) in parsed.items()], len(parsed)
    _model, parse, key = _parsers(tz)[tab]
    rows = parse(u, f)
    if tab == TAB_WORKER:  # compactRows_ と同じく 1 列目が空でない行
        return rows, sum(1 for r in rows if not _blank_js(r["cells"][0][1] if r["cells"] else ""))
    if tab == TAB_CALENDAR:
        return rows, sum(1 for r in rows if r["date"] is not None)
    return rows, valid_rows(rows, key)


def model_for(tab: str):
    return {TAB_SETTINGS: Setting, TAB_LOOKUP: Lookup, TAB_WORKER: Worker, TAB_CALENDAR: CalendarDay,
            TAB_STAMP: Stamp, TAB_DEPT_APPROVERS: DeptApprover}[tab]


def fitted(tab: str, rows: list[dict]) -> list[dict]:
    model = model_for(tab)
    return [_fit(model, r) for r in rows]


def replace_table(db: Session, tab: str, rows: list[dict]) -> None:
    """表を全置換する（同じトランザクション。commit は呼ぶ側）。"""
    model = model_for(tab)
    with db.begin_nested():
        db.execute(delete(model))
        db.add_all(model(**r) for r in fitted(tab, rows))


def import_tab(db: Session, tab: str, u: list, f: list, tz: Any) -> tuple[int, str]:
    """1 つのタブを取り込む。(入れた行の数, 飛ばした理由)。有効な行が 0 なら上書きしない。"""
    if not u:
        return 0, "タブが空です（今の値を残します）"
    rows, valid = build_rows(tab, u, f, tz)
    if valid == 0:
        return 0, "有効な行がありません（今の値を残します）"
    replace_table(db, tab, rows)
    return len(rows), ""


def env_mismatches(settings_map: dict[str, str], app_settings: Any) -> list[dict]:
    """Settings の ID が .env と違う点（書き込み先・読み先は .env が決める）。"""
    expected = {
        "PDF_FOLDER_ID": (app_settings.pdf_folder_id, False),
        "LEDGER_SSID": (app_settings.ledger_spreadsheet_id, False),
        "MASTER_SOURCE_SSID": (app_settings.master_source_spreadsheet_id, True),
        "CALENDAR_SOURCE_SSID": (app_settings.calendar_source_spreadsheet_id, True),
    }
    out = []
    for key, (want, empty_ok) in expected.items():
        have = normalize_gas(settings_map.get(key, ""))
        if empty_ok and have == "":
            continue
        if have != want:
            out.append({"key": key, "sheet": have, "env": want})
    return out


def settings_warnings(overlaps: list[dict]) -> list[str]:
    out = []
    for o in overlaps:
        if o["key"] == "LEDGER_SSID":
            out.append(f"LEDGER_SSID の行が {len(o['rows'])} つあります（行 {o['rows']}。下の行を使います。"
                       "GAS が台帳を作った跡の恐れ）")
    return out


def read_tabs(api: SheetsApi, sid: str, tabs: tuple[str, ...] = TABS) -> tuple[dict, list[str], Any]:
    """タブを UNFORMATTED と FORMATTED で読む。({タブ: (u, f)}, 無いタブ, 時間帯)。"""
    meta = api.get_spreadsheet(sid)
    titles = set(sheet_titles(meta))
    present = [t for t in tabs if t in titles]
    absent = [t for t in tabs if t not in titles]
    reads: dict = {}
    if present:
        u, f = api.read_both(sid, [quote_sheet(t) for t in present])
        reads = {t: (u[i], f[i]) for i, t in enumerate(present)}
    tz = str((meta.get("properties") or {}).get("timeZone") or "")
    return reads, absent, tz


def sync_sheet_config(db: Session, api: SheetsApi | None = None, app_settings: Any = None) -> dict:
    """6 タブを取り込む（読むだけ。commit は呼ぶ側）。戻り値 {ok, counts, overlaps, mismatches, skipped, warnings}。"""
    from app.config import get_settings

    settings = app_settings or get_settings()
    started = timeutil.now_utc()
    result: dict = {"ok": True, "counts": {}, "overlaps": [], "mismatches": [], "skipped": {}, "warnings": []}
    sid = settings.db_spreadsheet_id
    try:
        if api is None:
            api = SheetsApi.for_db(db, settings)
        reads, absent, _sheet_tz = read_tabs(api, sid)
    except Exception as error:  # noqa: BLE001 - 委任が無い・読めないなど。今の表を残す
        result.update({"ok": False, "error": f"{type(error).__name__}: {error}"[:500]})
        _record(db, started, result)
        return result
    tz = settings.db_sheet_tz
    for t in absent:
        result["skipped"][t] = "タブがありません（今の値を残します）"
    for tab in TABS:
        if tab not in reads:
            continue
        u, f = reads[tab]
        try:
            if tab == TAB_SETTINGS:
                parsed, overlaps = parse_settings(u, f)
                result["overlaps"] = overlaps
                result["warnings"].extend(settings_warnings(overlaps))
                result["mismatches"] = env_mismatches({k: v for k, (v, _r) in parsed.items()}, settings)
            n, reason = import_tab(db, tab, u, f, tz)
            if reason:
                result["skipped"][tab] = reason
            else:
                result["counts"][tab] = n
        except Exception as error:  # noqa: BLE001 - タブごとに握り、ほかのタブは続ける
            result["skipped"][tab] = f"読めません: {type(error).__name__}: {error}"[:500]
    result["ok"] = not result["skipped"]
    _record(db, started, result)
    if result["overlaps"] or result["mismatches"]:
        log.warning("sync_sheet_config: overlaps=%s mismatches=%s", result["overlaps"], result["mismatches"])
    return result


def _record(db: Session, started, result: dict) -> None:
    db.add(SyncRun(kind=RUN_KIND, started_at=started, finished_at=timeutil.now_utc(), ok=bool(result.get("ok")),
                   summary_json=dict(result)))
    db.flush()


def _scheduled(session_factory) -> dict:
    with session_factory() as db:
        result = sync_sheet_config(db)
        db.commit()
    return result


def register_worker() -> None:
    from app import worker
    from app.config import get_settings

    worker.register_schedule("sync_sheet_config", task=_scheduled,
                             every_minutes=get_settings().sheet_config_sync_minutes)
