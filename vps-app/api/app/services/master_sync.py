"""マスタ同期（決定 D・実装仕様 6-4。`sync_masters`、毎日 6:00 JST と手動。担当 D）。

src/syncMasters.gs の syncAllMasters と同じ規則で、元（読むだけ）を DB スプレッドシートのタブへ写す。

| 書き先 | 元の ID（.env） | 元のタブ | 見出しの確かめ（違えば止める） |
|---|---|---|---|
| 作業員マスタ | MASTER_SOURCE_SPREADSHEET_ID | Settings MASTER_SOURCE_WORKER_SHEET（空なら 作業員マスタ） | 作業員コード（か 作業員ID）・氏名・部署（か 部署ID） |
| M_CALENDAR | CALENDAR_SOURCE_SPREADSHEET_ID | Settings CALENDAR_SOURCE_SHEET（空なら 社内カレンダーマスタ） | 日付・区分 |
| M_STAMP | 同上 | StampMap | 1〜3 列目の見出しが今の書き先の 1〜3 列目と 1 字ずつ同じ（位置で読むため） |
| M_DEPT_APPROVERS | 同上 | DeptApprovers | 1〜5 列目が同じ |

GAS と同じにすること:
- 読む範囲は 1 行目の値がある最後の列まで × 最後の行まで（detectLastColByHeader_・readSheetData_）。読み取りの失敗は
  2 秒おきに 2 回まで読み直す（3 回）。
- 1 列目が空（String(v || '').trim() === ''。0 も空）の行を捨て、見出しは残す（compactRows_）。
- 1 行目から書く。文字は USER_ENTERED（setValues と同じ。'=' '+' '-' '@' で始まり数でない文字は ' を付けて式にしない）。
  数・真偽はそのまま。日付のセルは元のシリアル値を、元と書き先の時間帯が違えば瞬間を保つように直して数で書き、
  元の表示形式を写す（updateCells で表示形式だけ先に当ててから値を書く）。
- タブごとに失敗を握り、ほかのタブは続ける。

変えたこと（実装仕様 11-1）:
- 元が読めない・空・見出しだけ（有効な行 0）・見出しの確かめが合わない、なら そのタブは上書きしない。
- 余った行・列は値だけ消す（values.batchClear）。行・列・タブを消したり作り直したりしない。書き先のタブが無ければ作らずに飛ばす。
- 読むのは .env の ID だけ。Settings の MASTER_SOURCE_SSID・CALENDAR_SOURCE_SSID が違えば mismatches に出す。
- 書いた後、書いたタブを読み直して DB の表（workers・calendar・stamps・dept_approvers）を全置換する（config_sync と同じ読み方）。

戻り値 {ok, tabs:{名前:{ok, rows, skipped_reason}}, mismatches}。sync_runs（kind='sync_masters'）に残す。commit は呼ぶ側。
"""
from __future__ import annotations

import logging
import time
from collections.abc import Callable
from typing import Any

from sqlalchemy import select
from sqlalchemy.orm import Session

from app import timeutil
from app.constants import TAB_CALENDAR, TAB_DEPT_APPROVERS, TAB_STAMP, TAB_WORKER
from app.jsutil import js_string, js_trim, normalize_gas
from app.models import Setting, SyncRun
from app.services import config_sync
from app.services import sheet_codec as codec
from app.services import write_gate
from app.services.sheets_api import (
    SheetsApi,
    column_letter,
    quote_sheet,
    sheet_id_by_title,
    sheet_titles,
    spreadsheet_timezone,
)

log = logging.getLogger(__name__)

RUN_KIND = "sync_masters"
READ_ATTEMPTS = 3
READ_WAIT_SECONDS = 2.0
GRID_FIELDS = "sheets(data(rowData(values(userEnteredFormat(numberFormat),effectiveFormat(numberFormat)))))"
DEFAULT_DATE_FORMAT = {"type": "DATE", "pattern": "M/d/yyyy"}
DEFAULT_DATETIME_FORMAT = {"type": "DATE_TIME", "pattern": "M/d/yyyy H:mm:ss"}
FORMULA_LEADS = ("=", "+", "-", "@")

sleep: Callable[[float], None] = time.sleep  # 試験で差し替える


def plan(settings_map: dict[str, str]) -> list[dict]:
    """4 タブの同期の決まり（元のタブ名は Settings、空なら GAS の既定）。"""
    worker_tab = normalize_gas(settings_map.get("MASTER_SOURCE_WORKER_SHEET", "")) or "作業員マスタ"
    calendar_tab = normalize_gas(settings_map.get("CALENDAR_SOURCE_SHEET", "")) or "社内カレンダーマスタ"
    return [
        {"dst": TAB_WORKER, "source": "master", "src": worker_tab, "check": "worker"},
        {"dst": TAB_CALENDAR, "source": "calendar", "src": calendar_tab, "check": "calendar"},
        {"dst": TAB_STAMP, "source": "calendar", "src": "StampMap", "check": ("position", 3)},
        {"dst": TAB_DEPT_APPROVERS, "source": "calendar", "src": "DeptApprovers", "check": ("position", 5)},
    ]


# ===== GAS の読み方の写し =====


def _blank_js(v: Any) -> bool:
    if v is None or v is False or v == "" or (codec.is_number(v) and v == 0):
        return True
    return js_trim(js_string(v)) == ""


def detect_last_col(header: list) -> int:
    for c in range(len(header), 0, -1):
        if js_trim(js_string(header[c - 1])) != "":
            return c
    return 1


def compact(unformatted: list, formatted: list) -> tuple[list[list], list[list]]:
    """見出しの最後の列で切り、1 列目が空の行を捨てる（見出しは残す）。"""
    if not unformatted:
        return [], []
    last_col = detect_last_col(list(unformatted[0]))
    u_out, f_out = [], []
    for r, row in enumerate(unformatted):
        cells = (list(row) + [""] * last_col)[:last_col]
        frow = list(formatted[r]) if r < len(formatted) else []
        frow = (frow + [""] * last_col)[:last_col]
        if r > 0 and _blank_js(cells[0]):
            continue
        u_out.append(cells)
        f_out.append(frow)
    return u_out, f_out


def _header(row: list) -> list[str]:
    return [js_trim(js_string(h)) for h in row]


def check_headers(spec: Any, src_header: list, dst_header: list) -> str:
    """見出しの確かめ。合えば ''、違えば理由。"""
    h = _header(src_header)
    if spec == "worker":
        missing = [n for n in (("作業員コード", "作業員ID"), ("氏名",), ("部署", "部署ID")) if not any(x in h for x in n)]
        return f"元の見出しに {['か'.join(n) for n in missing]} がありません" if missing else ""
    if spec == "calendar":
        missing = [n for n in ("日付", "区分") if n not in h]
        return f"元の見出しに {missing} がありません" if missing else ""
    _kind, width = spec
    src = [js_string(v) for v in (list(src_header) + [""] * width)[:width]]
    dst = [js_string(v) for v in (list(dst_header) + [""] * width)[:width]]
    if src != dst:
        return f"1〜{width} 列目の見出しが書き先と違います（元 {src} ／ 書き先 {dst}）。位置で読む列なので止めました"
    return ""


def formula_safe(value: Any) -> Any:
    if isinstance(value, str) and value[:1] in FORMULA_LEADS and not codec.NUMBER_TEXT_RE.match(value.strip()):
        return "'" + value
    return value


def _source_formats(api: SheetsApi, source_id: str, tab: str) -> list[list[dict]]:
    meta = api.get_spreadsheet_row_metadata(source_id, [quote_sheet(tab)], GRID_FIELDS)
    out: list[list[dict]] = []
    for s in meta.get("sheets", []):
        for d in s.get("data", []):
            for row in d.get("rowData", []):
                line = []
                for cd in row.get("values", []):
                    nf = ((cd.get("effectiveFormat") or {}).get("numberFormat")
                          or (cd.get("userEnteredFormat") or {}).get("numberFormat") or {})
                    line.append(nf)
                out.append(line)
    return out


def prepare(unformatted: list, formatted: list, src_tz: Any, dst_tz: Any, row_map: list[int],
            formats: list[list[dict]] | None) -> tuple[list[list], list[tuple[int, int, dict]]]:
    """書く値（USER_ENTERED）と、日付のセルの表示形式 [(行0, 列0, numberFormat)]。

    row_map は compact 後の行 → 元の行（0 始まり）。formats は元の表示形式（無ければ既定）。
    """
    values: list[list] = []
    date_cells: list[tuple[int, int, dict]] = []
    for r, (urow, frow) in enumerate(zip(unformatted, formatted)):
        line = []
        for c, v in enumerate(urow):
            f = frow[c] if c < len(frow) else ""
            if r > 0 and codec.is_number(v) and codec.looks_like_date_display(f):
                moment = codec.serial_to_datetime(v, src_tz)
                serial = codec.datetime_to_serial(moment, dst_tz)
                src_r = row_map[r]
                fmt = (formats[src_r][c] if formats and src_r < len(formats) and c < len(formats[src_r]) else {}) or {}
                if fmt.get("type") not in codec.DATE_FORMAT_TYPES:
                    fmt = DEFAULT_DATE_FORMAT if float(v) == int(float(v)) else DEFAULT_DATETIME_FORMAT
                date_cells.append((r, c, {"type": fmt["type"], "pattern": fmt.get("pattern", "")}))
                line.append(serial)
            elif v is None:
                line.append("")
            elif isinstance(v, str):
                line.append(formula_safe(v))
            else:
                line.append(v)
        values.append(line)
    return values, date_cells


# ===== 読む・書く =====


def read_source(api: SheetsApi, source_id: str, tab: str) -> tuple[list, list]:
    last: Exception | None = None
    for attempt in range(READ_ATTEMPTS):
        try:
            u, f = api.read_both(source_id, [quote_sheet(tab)])
            return u[0], f[0]
        except Exception as error:  # noqa: BLE001 - GAS と同じく読み直す
            last = error
            if attempt + 1 < READ_ATTEMPTS:
                sleep(READ_WAIT_SECONDS)
    raise RuntimeError(f"Read failed after all retries: {last}")


def _row_map(unformatted: list) -> list[int]:
    out = []
    for r, row in enumerate(unformatted):
        if r == 0 or not _blank_js(row[0] if row else ""):
            out.append(r)
    return out


def write_tab(api: SheetsApi, sid: str, dst: str, values: list[list], date_cells: list, meta: dict) -> None:
    rows, cols = len(values), len(values[0])
    q = quote_sheet(dst)
    grid = {}
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == dst:
            grid = s["properties"].get("gridProperties") or {}
    row_count = int(grid.get("rowCount") or rows)
    col_count = int(grid.get("columnCount") or cols)
    if date_cells:
        sheet_id = sheet_id_by_title(meta, dst)
        reqs = [{"updateCells": {"start": {"sheetId": sheet_id, "rowIndex": r, "columnIndex": c},
                                 "rows": [{"values": [{"userEnteredFormat": {"numberFormat": fmt}}]}],
                                 "fields": "userEnteredFormat.numberFormat"}}
                for r, c, fmt in date_cells if r < row_count and c < col_count]
        if reqs:
            api.batch_update(sid, reqs, tab=dst, purpose="master_sync")
    api.values_batch_update(sid, [{"range": f"{q}!A1:{column_letter(cols - 1)}{rows}", "values": values}],
                            tab=dst, purpose="master_sync", input_option="USER_ENTERED")
    if date_cells:  # 行・列が増えて先に表示形式を当てられなかったセル
        sheet_id = sheet_id_by_title(meta, dst)
        late = [{"updateCells": {"start": {"sheetId": sheet_id, "rowIndex": r, "columnIndex": c},
                                 "rows": [{"values": [{"userEnteredFormat": {"numberFormat": fmt}}]}],
                                 "fields": "userEnteredFormat.numberFormat"}}
                for r, c, fmt in date_cells if r >= row_count or c >= col_count]
        if late:
            api.batch_update(sid, late, tab=dst, purpose="master_sync")
    clear = []
    last_col = column_letter(max(col_count, cols) - 1)
    if row_count > rows:
        clear.append(f"{q}!A{rows + 1}:{last_col}")
    if col_count > cols:
        clear.append(f"{q}!{column_letter(cols)}1:{last_col}{rows}")
    if clear:
        api.values_batch_clear(sid, clear, tab=dst, purpose="master_sync")


def _settings_map(db: Session) -> dict[str, str]:
    return {s.key: s.value for s in db.scalars(select(Setting))}


def source_mismatches(settings_map: dict[str, str], app_settings: Any) -> list[dict]:
    out = []
    for key, want in (("MASTER_SOURCE_SSID", app_settings.master_source_spreadsheet_id),
                      ("CALENDAR_SOURCE_SSID", app_settings.calendar_source_spreadsheet_id)):
        have = normalize_gas(settings_map.get(key, ""))
        if have and have != want:
            out.append({"key": key, "sheet": have, "env": want, "used": "env"})
    return out


def sync_masters(db: Session, api: SheetsApi | None = None, app_settings: Any = None) -> dict:
    """元 → DB のタブと表。戻り値 {ok, tabs:{名前:{ok, rows, skipped_reason}}, mismatches}。commit は呼ぶ側。"""
    from app.config import get_settings

    settings = app_settings or get_settings()
    started = timeutil.now_utc()
    smap = _settings_map(db)
    result: dict = {"ok": False, "tabs": {}, "mismatches": source_mismatches(smap, settings)}
    sources = {"master": settings.master_source_spreadsheet_id, "calendar": settings.calendar_source_spreadsheet_id}
    target_id = settings.db_spreadsheet_id
    allowed, reason = write_gate.check(db, settings)
    if not allowed:
        result.update({"error": "書き込みの関所で止めました: " + reason, "writeBlocked": reason})
        _record(db, started, result)
        return result
    try:
        if api is None:
            api = SheetsApi.for_db(db, settings)
        target_meta = api.get_spreadsheet(target_id)
        target_titles = set(sheet_titles(target_meta))
        dst_tz = codec.as_tz(spreadsheet_timezone(target_meta) or settings.db_sheet_timezone)
        source_meta = {k: api.get_spreadsheet(v) for k, v in sources.items()}
    except Exception as error:  # noqa: BLE001 - 結果で返す
        result["error"] = f"{type(error).__name__}: {error}"[:500]
        _record(db, started, result)
        return result

    for m in plan(smap):
        tab = {"ok": False, "rows": 0, "skipped_reason": ""}
        result["tabs"][m["dst"]] = tab
        try:
            smeta = source_meta[m["source"]]
            if m["src"] not in set(sheet_titles(smeta)):
                tab["skipped_reason"] = f"元シートなし: {m['src']}"
                continue
            if m["dst"] not in target_titles:
                tab["skipped_reason"] = "転記先シートなし（タブは作りません）"
                continue
            source_id = sources[m["source"]]
            u_raw, f_raw = read_source(api, source_id, m["src"])
            if not u_raw:
                tab["skipped_reason"] = "空（今のデータを残します）"
                continue
            u, f = compact(u_raw, f_raw)
            if len(u) <= 1:
                tab["skipped_reason"] = "見出しだけ（有効な行 0。今のデータを残します）"
                continue
            dst_header_rows = api.get_values(target_id, f"{quote_sheet(m['dst'])}!1:1", "UNFORMATTED_VALUE")
            problem = check_headers(m["check"], u[0], dst_header_rows[0] if dst_header_rows else [])
            if problem:
                tab["skipped_reason"] = problem
                continue
            src_tz = codec.as_tz(spreadsheet_timezone(smeta) or "Etc/GMT")
            needs_formats = any(codec.is_number(v) and codec.looks_like_date_display(fv)
                                for urow, frow in zip(u[1:], f[1:]) for v, fv in zip(urow, frow))
            formats = _source_formats(api, source_id, m["src"]) if needs_formats else None
            values, date_cells = prepare(u, f, src_tz, dst_tz, _row_map(u_raw), formats)
            write_tab(api, target_id, m["dst"], values, date_cells, target_meta)
            # 書いたタブを GAS が読む形で読み直し、DB の表を全置換する
            back_u, back_f = api.read_both(target_id, [quote_sheet(m["dst"])])
            n, why = config_sync.import_tab(db, m["dst"], back_u[0], back_f[0], dst_tz)
            if why:
                tab["skipped_reason"] = "DB の表: " + why
                continue
            tab.update({"ok": True, "rows": len(values) - 1, "cols": len(values[0])})
        except Exception as error:  # noqa: BLE001 - 1 タブの失敗でほかを止めない（GAS と同じ）
            tab["skipped_reason"] = f"ERROR {type(error).__name__}: {error}"[:500]
            log.warning("sync_masters %s: %s", m["dst"], error)
    result["ok"] = all(t["ok"] for t in result["tabs"].values())
    _record(db, started, result)
    log.info("sync_masters: %s", {k: (v["ok"], v["rows"], v["skipped_reason"]) for k, v in result["tabs"].items()})
    return result


def _record(db: Session, started, result: dict) -> None:
    db.add(SyncRun(kind=RUN_KIND, started_at=started, finished_at=timeutil.now_utc(), ok=bool(result.get("ok")),
                   summary_json=dict(result)))
    db.flush()


def _scheduled(session_factory) -> dict:
    with session_factory() as db:
        result = sync_masters(db)
        db.commit()
    return result


def register_worker() -> None:
    from app import worker

    worker.register_schedule("sync_masters", task=_scheduled, at=(6, 0))
