"""戻しの道具: 切り替え前の控え（tools/cutover_backup.py が作った複製）から、名前を書いたタブの値を
この環境の DB スプレッドシートへ写し戻す（担当 D）。何もしない確かめ（dry-run）が既定。

使い方（VPS。docs/cutover-runbook.md の戻しの章）:
  docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api \
      python tools/restore_tabs_from_copy.py --from-copy <控えの ID> --tabs all [--apply]

安全の決まり:
  - 書き先はこの環境の DB スプレッドシートだけ（safety.assert_restore_target）。控えの ID と同じ・許可リストの外は止める。
  - 控えは名前が「【切替前の控え】[<環境>] 休暇届 DB」で始まるスプレッドシートだけ（台帳の控えや別のファイルを拒む）。
    名前の環境が今の APP_ENV と違えば止める（本番にステージングの控えを戻さない）。本番では、ステージングの資源の ID
    （config の STAGING_*_KNOWN と .env の STAGING_*）を --from-copy に渡しても止める。
  - タブは --tabs に名前を書いたものだけ。`all` は T_LEAVE_REQUEST・V_LEAVE_SUMMARY・マスタ 4 タブ。
    M_SYSTEM_SETTING・M_LOOKUP は `all` に入らず、名前を書いたときだけ戻す。ほかのタブは書かない。
  - T_LEAVE_REQUEST は V 列だけを、見出しの文字で列を引いて戻す。G 列（SIGN_IMAGE_URL・PDF_URL）と N 列は戻さない
    （GAS が書いた値を残す）。行の位置が控えと違う（同じ行の REQ_ID が違う）なら止める。控えに無い REQ_ID の行
    （控えの後に足された申請）があれば --apply を止める（その行の V 列が消えるため）。それでも戻すときだけ
    --drop-rows-not-in-copy を付ける（先にその行の中身を控えてから）。
  - --apply が無ければ書かない（タブごとの行数・違うセルの数だけを出す）。
  - 値は UNFORMATTED_VALUE（日時はシリアル値）で読み、RAW で書く（文字は式にならない）。表示形式は今のタブのまま。
    控えより下・右の余った値は消す（値だけ。行・列・タブは消さない・作らない）。
  - worker を止めてから使う（行の反映・照合が同じタブを書かないように）。
終了コード: 0＝成功（dry-run を含む）、2＝止めた。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Optional

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "api"))

from app import safety  # noqa: E402
from app.jsutil import normalize_gas  # noqa: E402
from app.services import sheet_codec as codec  # noqa: E402
from app.services.sheets_api import API_BASE, column_letter, quote_sheet, sheet_titles  # noqa: E402

BACKUP_DB_PREFIX = "【切替前の控え】[{env}] 休暇届 DB"
ALL_TABS = ("T_LEAVE_REQUEST", "V_LEAVE_SUMMARY", "作業員マスタ", "M_CALENDAR", "M_STAMP", "M_DEPT_APPROVERS")


class RestoreStop(RuntimeError):
    """戻しを止める理由。"""


def staging_ids(settings: Any) -> frozenset[str]:
    """ステージングの資源の ID（config の既知の値と .env の STAGING_*）。"""
    from app import config as _config

    ids = set(_config.KNOWN_STAGING_IDS) | {settings.staging_spreadsheet_id, settings.staging_ledger_spreadsheet_id,
                                            settings.staging_pdf_folder_id}
    return frozenset(i for i in ids if i)


def check_backup_title(title: str, app_env: str) -> None:
    """控えの名前が「【切替前の控え】[<今の APP_ENV>] 休暇届 DB」で始まるか（環境の違う控え・旧い名前は止める）。"""
    want = BACKUP_DB_PREFIX.format(env=app_env)
    if title.startswith(want):
        return
    for env in ("staging", "production"):
        if env != app_env and title.startswith(BACKUP_DB_PREFIX.format(env=env)):
            raise RestoreStop(f"控えの環境（{env}）が今の APP_ENV（{app_env}）と違います")
    raise RestoreStop(f"控えの名前が「{want}」で始まりません")


def parse_tabs(text: str) -> tuple[list[str], set[str]]:
    named: list[str] = []
    for part in (text or "").split(","):
        name = part.strip()
        if name and name not in named:
            named.append(name)
    if not named:
        raise RestoreStop("--tabs にタブの名前を書いてください（all は T_LEAVE_REQUEST・V_LEAVE_SUMMARY・マスタ 4 タブ）")
    tabs: list[str] = []
    explicit: set[str] = set()
    for name in named:
        if name == "all":
            tabs.extend(t for t in ALL_TABS if t not in tabs)
        else:
            explicit.add(name)
            if name not in tabs:
                tabs.append(name)
    return tabs, explicit


def _grid_size(meta: dict, tab: str) -> tuple[int, int]:
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == tab:
            g = s["properties"].get("gridProperties") or {}
            return int(g.get("rowCount") or 0), int(g.get("columnCount") or 0)
    return 0, 0


def _trim(grid: list[list]) -> list[list]:
    rows = [list(r) for r in grid]
    while rows and all(v in ("", None) for v in rows[-1]):
        rows.pop()
    return rows


def _pad(grid: list[list], width: int) -> list[list]:
    return [[("" if v is None else v) for v in (r + [""] * (width - len(r)))[:width]] for r in grid]


def _cell(grid: list[list], r: int, c: int) -> Any:
    return grid[r][c] if r < len(grid) and 0 <= c < len(grid[r]) else ""


def _formula_cells(api: Any, copy_id: str, tab: str) -> int:
    formulas = api.get_values(copy_id, quote_sheet(tab), "FORMULA")
    return sum(1 for row in formulas for v in row if isinstance(v, str) and v.startswith("="))


def plan_leave_tab(api: Any, copy_id: str, target_id: str) -> dict:
    """T_LEAVE_REQUEST の V 列だけを戻す計画（列ごと）。"""
    tab = codec.TAB
    q = quote_sheet(tab)
    src = _trim(api.get_values(copy_id, q, "UNFORMATTED_VALUE"))
    cur = _trim(api.get_values(target_id, q, "UNFORMATTED_VALUE"))
    src_cols = codec.rule_columns(src[0] if src else [])
    cur_cols = codec.rule_columns(cur[0] if cur else [])
    missing = [r.header for r in codec.V_RULES if r.header not in src_cols or r.header not in cur_cols]
    if missing:
        raise RestoreStop(f"{tab}: 控えか DB に見出しがありません: {missing}")
    sid, cid = src_cols[codec.ID_RULE.header], cur_cols[codec.ID_RULE.header]
    moved = []
    for r in range(1, min(len(src), len(cur))):
        a, b = normalize_gas(_cell(src, r, sid)), normalize_gas(_cell(cur, r, cid))
        if a and b and a != b:
            moved.append({"row": r + 1, "copy": a, "current": b})
    in_copy = {normalize_gas(_cell(src, r, sid)) for r in range(1, len(src))}
    not_in_copy = [{"row": r + 1, "reqId": normalize_gas(_cell(cur, r, cid))} for r in range(1, len(cur))
                   if normalize_gas(_cell(cur, r, cid)) and normalize_gas(_cell(cur, r, cid)) not in in_copy]
    columns, diff = [], 0
    height = len(src)
    for rule in codec.V_RULES:
        cs, ct = src_cols[rule.header], cur_cols[rule.header]
        values = [["" if _cell(src, r, cs) is None else _cell(src, r, cs)] for r in range(1, height)]
        diff += sum(1 for r in range(1, max(height, len(cur))) if _cell(src, r, cs) != _cell(cur, r, ct))
        columns.append({"header": rule.header, "col": ct, "values": values})
    return {"tab": tab, "mode": "columns", "columns": columns, "rows": height, "currentRows": len(cur),
            "diffCells": diff, "formulaCells": _formula_cells(api, copy_id, tab), "rowsNotInCopy": not_in_copy,
            "movedRows": moved, "skippedColumns": [r.header for r in codec.G_RULES + codec.N_RULES]}


def plan_tab(api: Any, copy_id: str, target_id: str, tab: str) -> dict:
    """タブ全体を戻す計画（V_LEAVE_SUMMARY・マスタ・名前を書いた設定のタブ）。"""
    q = quote_sheet(tab)
    src = _trim(api.get_values(copy_id, q, "UNFORMATTED_VALUE"))
    cur = _trim(api.get_values(target_id, q, "UNFORMATTED_VALUE"))
    width = max([len(r) for r in src] + [1])
    values = _pad(src, width)
    cur_width = max([len(r) for r in cur] + [0])
    diff = 0
    for r in range(max(len(values), len(cur))):
        for c in range(max(width, cur_width)):
            if _cell(values, r, c) not in (_cell(cur, r, c),) and not (
                    _cell(values, r, c) in ("", None) and _cell(cur, r, c) in ("", None)):
                diff += 1
    return {"tab": tab, "mode": "whole", "values": values, "rows": len(values), "cols": width, "currentRows": len(cur),
            "currentCols": cur_width, "diffCells": diff, "formulaCells": _formula_cells(api, copy_id, tab),
            "rowsNotInCopy": [], "movedRows": []}


def _post(api: Any, target_id: str, path: str, body: dict) -> None:
    # 書く関数は SheetsApi の目的の許可リストを通らないので、assert_restore_target で確かめてから呼ぶ
    api._call("POST", f"{API_BASE}/{target_id}/{path}", body=body)


def apply_leave_tab(api: Any, target_id: str, plan: dict, meta: dict) -> None:
    q = quote_sheet(plan["tab"])
    grid_rows, _ = _grid_size(meta, plan["tab"])
    data, clear = [], []
    last = max(grid_rows, plan["currentRows"], plan["rows"])
    for col in plan["columns"]:
        letter = column_letter(col["col"])
        if col["values"]:
            data.append({"range": f"{q}!{letter}2:{letter}{plan['rows']}", "values": col["values"]})
        if last > plan["rows"]:
            clear.append(f"{q}!{letter}{plan['rows'] + 1}:{letter}")
    if data:
        _post(api, target_id, "values:batchUpdate", {"valueInputOption": "RAW", "data": data})
    if clear:
        safety.assert_ranges_in_tab(clear, plan["tab"])
        _post(api, target_id, "values:batchClear", {"ranges": clear})


def apply_tab(api: Any, target_id: str, plan: dict, meta: dict) -> None:
    tab = plan["tab"]
    q = quote_sheet(tab)
    rows, cols = plan["rows"], plan["cols"]
    grid_rows, grid_cols = _grid_size(meta, tab)
    if rows:
        _post(api, target_id, "values:batchUpdate",
              {"valueInputOption": "RAW", "data": [{"range": f"{q}!A1:{column_letter(cols - 1)}{rows}",
                                                    "values": plan["values"]}]})
    clear = []
    last_col = column_letter(max(grid_cols, cols, plan["currentCols"], 1) - 1)
    if max(grid_rows, plan["currentRows"]) > rows:
        clear.append(f"{q}!A{rows + 1}:{last_col}")
    if max(grid_cols, plan["currentCols"]) > cols and rows:
        clear.append(f"{q}!{column_letter(cols)}1:{last_col}{rows}")
    if clear:
        safety.assert_ranges_in_tab(clear, tab)
        _post(api, target_id, "values:batchClear", {"ranges": clear})


def run(api: Any, settings: Any, copy_id: str, tabs_arg: str, *, apply: bool,
        drop_rows_not_in_copy: bool = False) -> dict:
    copy_id = (copy_id or "").strip()
    target_id = safety.writable_spreadsheet_id(settings)
    tabs, explicit = parse_tabs(tabs_arg)
    if not copy_id:
        raise RestoreStop("--from-copy が空です")
    if copy_id == target_id or copy_id in safety.never_writable_ids(settings):
        raise RestoreStop("--from-copy に DB・台帳・マスタ元そのものは使えません（控えの ID を渡す）")
    if not settings.is_staging and copy_id in staging_ids(settings):
        raise RestoreStop("本番では、ステージングの資源の ID は --from-copy に使えません")
    for tab in tabs:
        try:
            safety.assert_restore_target(target_id, tab, settings, protected_listed=tab in explicit)
        except safety.UnsafeWriteError as error:
            raise RestoreStop(str(error)) from error
    copy_meta = api.get_spreadsheet(copy_id)
    title = str((copy_meta.get("properties") or {}).get("title") or "")
    check_backup_title(title, settings.app_env)
    target_meta = api.get_spreadsheet(target_id)
    missing = [t for t in tabs if t not in set(sheet_titles(copy_meta)) or t not in set(sheet_titles(target_meta))]
    if missing:
        raise RestoreStop(f"控えか DB に無いタブがあります（タブは作りません）: {missing}")
    plans = [plan_leave_tab(api, copy_id, target_id) if t == codec.TAB else plan_tab(api, copy_id, target_id, t)
             for t in tabs]
    report: dict = {"appEnv": settings.app_env, "apply": apply, "copyId": copy_id, "targetId": target_id,
                    "tabs": [{k: v for k, v in p.items() if k not in ("values", "columns")} for p in plans]}
    moved = {p["tab"]: p["movedRows"] for p in plans if p["movedRows"]}
    if apply and moved:
        raise RestoreStop(f"行の位置が控えと違います（同じ行の REQ_ID が違う）。列ごとに戻すと G 列とずれるので止めました: "
                          f"{ {t: rows[:20] for t, rows in moved.items()} }")
    lost = {p["tab"]: p["rowsNotInCopy"] for p in plans if p["rowsNotInCopy"]}
    if apply and lost and not drop_rows_not_in_copy:
        raise RestoreStop("控えに無い REQ_ID の行があります（写し戻すとその行の V 列が消えます）: "
                          + "; ".join(f"{t}: " + ", ".join(f"{x['row']} 行目 {x['reqId']}" for x in rows[:20])
                                      for t, rows in lost.items())
                          + "。中身を控えてから、それでも戻すときだけ --drop-rows-not-in-copy を付けます")
    if apply:
        for p in plans:
            safety.assert_restore_target(target_id, p["tab"], settings, protected_listed=p["tab"] in explicit)
            if p["mode"] == "columns":
                apply_leave_tab(api, target_id, p, target_meta)
            else:
                apply_tab(api, target_id, p, target_meta)
        report["applied"] = [p["tab"] for p in plans]
    return report


def main(argv: Optional[list] = None, *, api: Any = None, settings: Any = None) -> int:
    parser = argparse.ArgumentParser(description="控えからタブの値を DB スプレッドシートへ写し戻す（既定は dry-run）")
    parser.add_argument("--from-copy", required=True, help="tools/cutover_backup.py が出した DB の控えの ID")
    parser.add_argument("--tabs", required=True, help="カンマ区切りのタブの名前。all は T_LEAVE_REQUEST・V_LEAVE_SUMMARY・マスタ 4 タブ")
    parser.add_argument("--apply", action="store_true", help="本当に書く（無ければ確かめるだけ）")
    parser.add_argument("--drop-rows-not-in-copy", action="store_true",
                        help="控えに無い T_LEAVE_REQUEST の行の V 列が消えてもよいときだけ付ける")
    args = parser.parse_args(argv)
    try:
        if settings is None:
            from app.config import get_settings

            settings = get_settings()
        if api is None:
            from sqlalchemy.orm import Session

            from app.db import get_engine
            from app.services.sheets_api import SheetsApi

            with Session(get_engine()) as db:
                api = SheetsApi.for_db(db, settings)
        report = run(api, settings, args.from_copy, args.tabs, apply=args.apply,
                     drop_rows_not_in_copy=args.drop_rows_not_in_copy)
    except RestoreStop as error:
        print(f"止めました: {error}", file=sys.stderr)
        return 2
    except Exception as error:  # noqa: BLE001 - 委任が無い・API の失敗など
        print(f"止めました: {type(error).__name__}: {str(error)[:300]}", file=sys.stderr)
        return 2
    print(json.dumps(report, ensure_ascii=False, indent=1, default=str))
    if not args.apply:
        print("確かめだけです（書いていません）。書くときは --apply を付けます。", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
