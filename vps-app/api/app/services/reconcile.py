"""毎晩の照合（実装仕様 6-7。`sheet_reconcile`、2:10 JST。担当 D）。

- T_LEAVE_REQUEST を UNFORMATTED（SERIAL_NUMBER）と FORMATTED で全行読み、DB と 1 セルずつ比べる。
  日付は JST の日付（GMT 15:00＝.625 の行も、GMT 0:00＝JST 9:00 の古い行も同じ日）、日時は秒、文字は GAS が読む形。
  V 列は型（日付のセルか）と表示形式（表示が日付か。統合業務メールは Date の行だけを読む）も見る。
- G 列（SIGN_IMAGE_URL・PDF_URL）と N 列は比べるが書かない（差は gDiffs・nDiffs に出すだけ）。
- 反映待ち・blocked・GAS の仕事の途中の行は比べない。
- 差があれば DB の値でその行の V 列を書き直し（purpose='reconcile'）、直した行・列・前後の値を sync_runs に残す。
- シートにあって DB に無い行（REQ_ID が空の半端な行を含む）は消さず書かずに報告する（notInDb・extraRows）。
- 歯止め（罠 49）: DB の leave_requests が 0 行か、シートの REQ_ID のある行より DB の行が少なければ、1 行も書かずに止める。
- V_LEAVE_SUMMARY も DB から作った値と比べ、違えば summary_sheet.write で書き直す（--check では数えるだけ）。
- check_only=True（`python -m app.worker run sheet_reconcile --check`）は書かずに数えるだけ。

compare_rows は移行の道具（移した直後の照合）も使う。
"""
from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Any

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app import timeutil
from app.jsutil import js_string, normalize_gas
from app.models import Job, LeaveRequest, SheetSyncState, SyncRun
from app.safety import UnsafeWriteError
from app.services import row_sync, summary_sheet
from app.services import sheet_codec as codec
from app.services import write_gate
from app.services.sheets_api import SheetsApi, SheetsError, column_letter, quote_sheet, sheet_titles

log = logging.getLogger(__name__)

RUN_KIND = "sheet_reconcile"
MAX_LISTED = 200


def _row(grid: list, r: int) -> list:
    return grid[r - 1] if 0 < r <= len(grid) else []


def _display(u: Any, f: Any) -> Any:
    if codec.is_blank(u):
        return ""
    return f if isinstance(f, str) and f != "" else u


def compare_rows(unformatted: list, formatted: list, db_rows: list[dict], tz) -> dict:
    """T_LEAVE_REQUEST の全行と DB の行を比べる（読むだけ）。

    db_rows: [{"pk", "sheet_row", "version", "values"(見出し → 論理の値), "kinds", "skip"}]
    戻り値: {"diffs"(V 列), "gDiffs", "nDiffs", "idMismatch", "notInDb", "extraRows", "sheetRows", "dbRows", "compared"}
    """
    header = _row(unformatted, codec.HEADER_ROW)
    missing = codec.missing_headers(header)
    if missing:
        raise row_sync.LaneHalt(f"{codec.TAB}: 決まりの見出し {missing} がありません")
    col = codec.rule_columns(header)
    id_col = col[codec.ID_RULE.header]
    diffs: list[dict] = []
    g_diffs: list[dict] = []
    n_diffs: list[dict] = []
    id_mismatch: list[dict] = []
    claimed = {d["sheet_row"] for d in db_rows if d.get("sheet_row")}
    compared = 0
    for d in db_rows:
        r = d.get("sheet_row")
        if not r or d.get("skip"):
            continue
        urow, frow = _row(unformatted, r), _row(formatted, r)
        key = normalize_gas(codec.cell_at(urow, id_col))
        if key != normalize_gas(d["pk"]):
            id_mismatch.append({"row": r, "pk": d["pk"], "sheet": key})
            continue
        compared += 1
        for rule in codec.COLUMN_RULES:
            c = col[rule.header]
            u, f = codec.cell_at(urow, c), codec.cell_at(frow, c)
            expected = d["values"].get(rule.header)
            entry = {"row": r, "pk": d["pk"], "header": rule.header, "sheet": _display(u, f),
                     "db": codec.canonical(rule, expected)}
            value, kind = codec.decode_cell(rule, u, f, tz)
            target = diffs if rule.cls == "V" else (g_diffs if rule.cls == "G" else n_diffs)
            if not codec.same_value(rule, value, expected):
                target.append({**entry, "reason": "value"})
                continue
            if rule.cls != "V" or codec.is_blank(expected) or not rule.date_like:
                continue
            want = d.get("kinds", {}).get(rule.attr)
            if want in (codec.SERIAL, codec.STRING) and kind != want:
                diffs.append({**entry, "reason": "kind"})
            elif kind == codec.SERIAL and not codec.looks_like_date_display(f):
                diffs.append({**entry, "reason": "format"})
    db_keys = {normalize_gas(d["pk"]) for d in db_rows}
    extra, not_in_db = [], []
    sheet_rows = 0
    for r in range(codec.DATA_START_ROW, len(unformatted) + 1):
        urow = _row(unformatted, r)
        key = normalize_gas(codec.cell_at(urow, id_col))
        if key:
            sheet_rows += 1
            if key not in db_keys:
                not_in_db.append({"row": r, "reqId": key})
        elif any(not codec.is_blank(v) for v in urow) and r not in claimed:
            extra.append({"row": r, "cells": {column_letter(i): js_string(v)
                                              for i, v in enumerate(urow) if not codec.is_blank(v)}})
    return {"diffs": diffs, "gDiffs": g_diffs, "nDiffs": n_diffs, "idMismatch": id_mismatch, "notInDb": not_in_db,
            "extraRows": extra, "sheetRows": sheet_rows, "dbRows": len(db_rows), "compared": compared}


def gas_busy_request_ids(db: Session) -> set[str]:
    rows = db.scalars(select(Job.request_id).where(Job.lane == "gas", Job.state.in_(("pending", "running")),
                                                   Job.request_id != "")).all()
    return {normalize_gas(r) for r in rows}


def db_rows_for(db: Session) -> list[dict]:
    states = {s.pk: s for s in db.scalars(select(SheetSyncState).where(SheetSyncState.table_name == codec.TABLE))}
    busy = gas_busy_request_ids(db)
    out = []
    for row in db.scalars(select(LeaveRequest)):
        pk = str(row.req_id)
        st = states.get(pk)
        dirty = bool(st and st.desired_version > st.synced_version)
        blocked = bool(st and st.state == "blocked") or normalize_gas(pk) in busy
        synced = (st.synced_values if st else {}) or {}
        out.append({"pk": pk, "sheet_row": row.sheet_row, "version": int(row.row_version),
                    "values": codec.db_values(row), "kinds": dict(row.cell_kinds or {}), "synced": dict(synced),
                    "skip": dirty or blocked})
    return out


def _record(session_factory, ok: bool, summary: dict) -> None:
    now = timeutil.now_utc()
    with session_factory() as db:
        db.add(SyncRun(kind=RUN_KIND, started_at=summary.get("_started", now), finished_at=now, ok=ok,
                       summary_json={k: v for k, v in summary.items() if not k.startswith("_")}))
        db.commit()


def _fix(api: SheetsApi, sid: str, meta: dict, diffs: list[dict], rows: list[dict], session_factory, tz) -> int:
    """差のある V 列を DB の値で書き直す。書き直した行の数。"""
    by_pk = {d["pk"]: d for d in rows}
    headers_by_pk: dict[str, list[str]] = {}
    for diff in diffs:
        hs = headers_by_pk.setdefault(diff["pk"], [])
        if diff["header"] not in hs:
            hs.append(diff["header"])
    with session_factory() as db:
        current = {}
        for pk in headers_by_pk:
            st = db.get(SheetSyncState, (codec.TABLE, pk))
            row = db.get(LeaveRequest, pk)
            current[pk] = (row is not None and int(row.row_version) == by_pk[pk]["version"]
                           and not (st and st.desired_version > st.synced_version))
    works = []
    for pk, headers in headers_by_pk.items():
        if not current.get(pk):
            continue
        d = by_pk[pk]
        works.append(row_sync.RowWork(pk=pk, version=d["version"], values=d["values"], kinds=d["kinds"],
                                      synced=d["synced"], sheet_row=d["sheet_row"], seq=0, first_dirty_at=None,
                                      attempts=0, target_row=d["sheet_row"], headers=headers))
    if not works:
        return 0
    ctx = row_sync.load_tab(api, sid, meta, need_tail=False)
    works = [w for w in works if ctx.id_rows.get(w.sheet_row) == w.key]  # 比べた後に並べ替えられた行には書かない
    if not works:
        return 0
    row_sync.load_refs(api, sid, ctx, timeutil.now_utc())
    row_sync.write_rows(api, sid, ctx, works, tz, purpose="reconcile")
    fixed = 0
    with session_factory() as db:
        for w in works:
            if w.error:
                continue
            st = db.get(SheetSyncState, (codec.TABLE, w.pk))
            if st is not None and st.desired_version <= st.synced_version:
                st.synced_values = row_sync.merge_synced(st.synced_values, w.written or {})
            fixed += 1
        db.commit()
    return fixed


def _strip(summary: dict) -> dict:
    return {k: v for k, v in summary.items() if not k.startswith("_")}


def run(session_factory: Callable[[], Session], *, check_only: bool = False, api: SheetsApi | None = None,
        app_settings: Any = None) -> dict:
    """照合を 1 回。check_only なら書かない。戻り値は sync_runs の summary と同じ形（ok を含む）。"""
    from app.config import get_settings

    settings = app_settings or get_settings()
    started = timeutil.now_utc()
    summary: dict = {"checkOnly": check_only, "_started": started}
    with session_factory() as db:
        count = int(db.scalar(select(func.count()).select_from(LeaveRequest)) or 0)
    if count == 0:
        summary.update({"ok": False, "stopped": "DB の leave_requests が 0 行のため、照合を止めました（書きません）。"})
        _record(session_factory, False, summary)
        return _strip(summary)
    sid = settings.db_spreadsheet_id
    tz = settings.db_sheet_tz
    try:
        if api is None:
            api = row_sync.api_factory(session_factory, settings)
        meta = row_sync.check_timezone(api, sid, settings)
        u, f = api.read_both(sid, [quote_sheet(codec.TAB)])
        with session_factory() as db:
            rows = db_rows_for(db)
        result = compare_rows(u[0], f[0], rows, tz)
        summary.update({
            "dbRows": result["dbRows"], "sheetRows": result["sheetRows"], "compared": result["compared"],
            "diffCount": len(result["diffs"]), "diffs": result["diffs"][:MAX_LISTED],
            "gDiffs": result["gDiffs"][:MAX_LISTED], "nDiffs": result["nDiffs"][:MAX_LISTED],
            "idMismatch": result["idMismatch"][:MAX_LISTED], "idMismatchCount": len(result["idMismatch"]),
            "notInDb": result["notInDb"][:MAX_LISTED], "extraRows": result["extraRows"][:MAX_LISTED],
            "skipped": sum(1 for d in rows if d["skip"]),
        })
        if result["dbRows"] < result["sheetRows"]:
            summary.update({"ok": False, "stopped": f"DB の行（{result['dbRows']}）がシートの REQ_ID のある行"
                                                    f"（{result['sheetRows']}）より少ないため、1 行も書かずに止めました。"})
            _record(session_factory, False, summary)
            return _strip(summary)
        # V_LEAVE_SUMMARY
        summary_diffs = _summary_diffs(api, sid, meta, session_factory, tz)
        summary["summaryDiffCount"] = len(summary_diffs) if summary_diffs is not None else None
        summary["summaryDiffs"] = (summary_diffs or [])[:MAX_LISTED]
        fixed = 0
        rewrote = False
        if not check_only and (result["diffs"] or summary_diffs):
            with session_factory() as db:
                allowed, reason = write_gate.check(db, settings)
            if not allowed:
                summary.update({"ok": False, "fixedRows": 0, "writeBlocked": reason,
                                "stopped": "書き込みの関所で止めました（差は報告だけ）: " + reason})
                _record(session_factory, False, summary)
                return _strip(summary)
            if result["diffs"]:
                fixed = _fix(api, sid, meta, result["diffs"], rows, session_factory, tz)
            if summary_diffs:
                with session_factory() as db:
                    w = summary_sheet.write(db, api=api, app_settings=settings, gate=False)
                    db.commit()
                rewrote = bool(w.get("ok"))
        summary["fixedRows"] = fixed
        summary["summaryRewritten"] = rewrote
        if check_only:
            summary["ok"] = not result["diffs"] and not result["idMismatch"] and not summary_diffs
        else:
            summary["ok"] = not result["idMismatch"]
    except (row_sync.LaneHalt, UnsafeWriteError, SheetsError) as error:
        summary.update({"ok": False, "stopped": str(error)[:1000]})
        _record(session_factory, False, summary)
        return _strip(summary)
    _record(session_factory, summary["ok"], summary)
    return _strip(summary)


def _summary_diffs(api: SheetsApi, sid: str, meta: dict, session_factory, tz) -> list[dict] | None:
    """V_LEAVE_SUMMARY と DB から作った集計の差（W の集計がまだ無い・タブが無いときは None）。"""
    if summary_sheet.TAB not in set(sheet_titles(meta)):
        return None
    try:
        with session_factory() as db:
            expected = summary_sheet.summary_rows(db, timeutil.now_utc())
    except ImportError:
        return None
    values = api.get_values(sid, quote_sheet(summary_sheet.TAB), "UNFORMATTED_VALUE")
    return summary_sheet.compare(expected, values, tz)


def _scheduled(session_factory, check_only: bool = False) -> dict:
    return run(session_factory, check_only=check_only)


def register_worker() -> None:
    from app import worker

    worker.register_schedule("sheet_reconcile", task=_scheduled, at=(2, 10))


__all__ = ["compare_rows", "db_rows_for", "run", "register_worker"]
