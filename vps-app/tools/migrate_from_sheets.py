"""本番 DB スプレッドシート → VPS の DB の移行（スプレッドシートは読むだけ。実装仕様 6-6。担当 D）。

移すタブ:
  T_LEAVE_REQUEST → leave_requests（見出し 1 行目、データ 2 行目から。見出しの文字で列を引く）
  M_SYSTEM_SETTING・M_LOOKUP・作業員マスタ・M_CALENDAR・M_STAMP・M_DEPT_APPROVERS → 5-2 の表（config_sync と同じ読み方）
  V_LEAVE_SUMMARY は移さない（DB から作る）。照合にだけ使う。

使い方（VPS。tools はイメージに入っていないので付けて動かす＝罠 5）:
  docker compose run --rm -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api \
      python tools/migrate_from_sheets.py --source sheets --dry-run --report-out /tmp/report.json
  docker compose run --rm -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api \
      python tools/migrate_from_sheets.py --source sheets [--snapshot-out /tmp/snapshot.json]
--source sheets は DB に保存した委任トークンで読むので、--dry-run でも DB と .env が要る。

安全の決まり:
  - 読む元は本番 DB スプレッドシート（config.PROD_SPREADSHEET_ID_FIXED）だけ。ほかの ID は止める（staging への移行も本番を
    読む＝読むだけ）。Sheets API で UNFORMATTED（SERIAL_NUMBER）と FORMATTED を一緒に読む（罠 11）。
  - 見出しの欠けは止める（missingColumns）。決まりに無い見出しは droppedColumns に出す（触らない）。
  - REQ_ID が空の行は飛ばし、行番号と値のあるセルを skippedBlankIdRows に出す（本番の途中に 52 行）。
    REQ_ID の重なり・LV-\\d{8}-\\d{3,} でない REQ_ID・決まりに無い状態は止める。
  - 古い行の年度・日付の型は計算し直さずに保つ（年度はセルの値のまま。休暇日・振替元は JST の日付にして持ち、型は
    cell_kinds、元のシリアル値は sheet_sync_state.synced_values の __raw__ に残す。シートの行は書き換えない）。
    年度と computeFiscalYear_(休暇日) が違う行の数は報告に出すだけ。
  - sheet_row＝その行、row_version=1、sheet_sync_state は反映済み（synced_values＝今シートにある値）。
  - pdf_file_id は PDF_URL の /d/<id>/ から取る。
  - APP_ENV=staging では PDF_URL・SIGN_IMAGE_URL・pdf_file_id を空にして入れる（罠 72。元の値は列があれば source_* に）。
  - 移行先の表（業務・設定とマスタ・sheet_sync_state・jobs・idempotency_keys・sync_runs・sessions）のどれかに行があれば
    止める。--truncate を付けたときだけ空にしてから入れる。1 トランザクションで入れる。
  - T_LEAVE_REQUEST か M_SYSTEM_SETTING が 0 行なら止める（--dry-run でも。罠 49）。
  - 照合: 件数・状態別・2 次承認済み・休暇日と振替元（JST の日付）・申請日時と承認日時（秒）・作業員 ID（GAS が読む形）・
    年度を、シートから読んだ行と DB から読み直した行で比べる。移した直後の reconcile.compare_rows が差 0。
    DB から作った集計（W の domain.summaries）と本番の V_LEAVE_SUMMARY が 12 列とも同じ（W がまだ無ければ報告だけ）。
終了コード: 0＝成功、2＝止めた、1＝入れた後の照合が合わない。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "api"))

from app import timeutil  # noqa: E402
from app.config import PROD_SPREADSHEET_ID_FIXED  # noqa: E402
from app.constants import LEAVE_STATUSES, TAB_LEAVE_SUMMARY, TAB_SETTINGS  # noqa: E402
from app.jsutil import js_string, normalize_gas  # noqa: E402
from app.services import config_sync, reconcile, summary_sheet  # noqa: E402
from app.services import sheet_codec as codec  # noqa: E402
from app.services.sheets_api import SheetsError, quote_sheet, sheet_titles, spreadsheet_timezone  # noqa: E402

CONFIG_TABLES = {"M_SYSTEM_SETTING": "settings", "M_LOOKUP": "lookup", "作業員マスタ": "workers",
                 "M_CALENDAR": "calendar", "M_STAMP": "stamps", "M_DEPT_APPROVERS": "dept_approvers"}
TARGET_TABLES = ("leave_requests",) + tuple(CONFIG_TABLES.values()) + ("sheet_sync_state",)
TRUNCATE_TABLES = TARGET_TABLES + ("jobs", "idempotency_keys", "sync_runs", "sessions")
REPORTED_SETTING_KEYS = ("APP_URL", "PDF_FOLDER_ID", "LEDGER_SSID", "MASTER_SOURCE_SSID", "CALENDAR_SOURCE_SSID",
                         "MASTER_SOURCE_WORKER_SHEET", "CALENDAR_SOURCE_SHEET")


class MigrationStop(RuntimeError):
    """移行を止める理由（安全装置）。"""


# ===== 読み込み元 =====


def check_source_id(spreadsheet_id: str) -> None:
    if spreadsheet_id != PROD_SPREADSHEET_ID_FIXED:
        raise MigrationStop(f"移行元が本番 DB スプレッドシートの ID ではありません: {spreadsheet_id}")


def read_sheets(api: Any, spreadsheet_id: str) -> dict:
    """本番 DB を読む（読むだけ）。{"tz", "tabs": {タブ: (UNFORMATTED, FORMATTED)}, "absentTabs"}。"""
    check_source_id(spreadsheet_id)
    meta = api.get_spreadsheet(spreadsheet_id)
    titles = set(sheet_titles(meta))
    for required in (codec.TAB, TAB_SETTINGS):
        if required not in titles:
            raise MigrationStop(f"タブがありません: {required}")
    wanted = [codec.TAB] + [t for t in CONFIG_TABLES if t in titles]
    if TAB_LEAVE_SUMMARY in titles:
        wanted.append(TAB_LEAVE_SUMMARY)
    u, f = api.read_both(spreadsheet_id, [quote_sheet(t) for t in wanted])
    return {"tz": spreadsheet_timezone(meta), "tabs": {t: (u[i], f[i]) for i, t in enumerate(wanted)},
            "absentTabs": [t for t in CONFIG_TABLES if t not in titles]}


# ===== 読んだ値 → 表の行 =====


def _row(grid: list, r: int) -> list:
    return grid[r - 1] if 0 < r <= len(grid) else []


def decode_requests(unformatted: list, formatted: list, tz: Any) -> dict:
    header = _row(unformatted, codec.HEADER_ROW)
    normalized = [normalize_gas(h) for h in header]
    known = set(codec.HEADERS) | {r.alias for r in codec.COLUMN_RULES if r.alias}
    dup = sorted(h for h, n in Counter(x for x in normalized if x in known).items() if n > 1)
    missing = codec.missing_headers(header)
    col = codec.rule_columns(header)
    id_col = col.get(codec.ID_RULE.header, 0)
    rows, raws, blank_rows, empty = [], {}, [], 0
    for r in range(codec.DATA_START_ROW, len(unformatted) + 1):
        urow, frow = _row(unformatted, r), _row(formatted, r)
        if all(codec.is_blank(v) for v in urow):
            empty += 1
            continue
        if not normalize_gas(codec.cell_at(urow, id_col)):
            blank_rows.append({"row": r, "cells": {normalize_gas(codec.cell_at(header, i)) or f"#{i + 1}": js_string(v)
                                                   for i, v in enumerate(urow) if not codec.is_blank(v)}})
            continue
        record, raw = codec.decode_row(col, urow, frow, tz, r)
        record["pdf_file_id"] = codec.pdf_file_id_from_url(record.get("pdf_url"))
        rows.append(record)
        raws[record["req_id"]] = raw
    return {"rows": rows, "raws": raws, "missing": missing, "duplicated": dup,
            "dropped": codec.dropped_headers(header), "skippedBlankIdRows": blank_rows, "emptyRows": empty}


def prepare(reads: dict) -> dict:
    """読んだタブを確かめ、入れる行と報告の材料を作る。止める理由があれば MigrationStop。"""
    tz = codec.as_tz(reads["tz"] or "Etc/GMT")
    u, f = reads["tabs"][codec.TAB]
    d = decode_requests(u, f, tz)
    problems = []
    if d["missing"]:
        problems.append(f"{codec.TAB}: 決まりの見出しがありません（missingColumns）: {d['missing']}")
    if d["duplicated"]:
        problems.append(f"{codec.TAB}: 同じ見出しが重なっています: {d['duplicated']}")
    if problems:
        raise MigrationStop("\n".join(problems))
    rows = d["rows"]
    bad_id = [(r["sheet_row"], r["req_id"]) for r in rows if not codec.REQ_ID_RE.match(js_string(r["req_id"]))]
    if bad_id:
        raise MigrationStop(f"{codec.TAB}: LV-yyyyMMdd-NNN の形でない REQ_ID があります: {bad_id[:20]}")
    dup = sorted(k for k, n in Counter(r["req_id"] for r in rows).items() if n > 1)
    if dup:
        raise MigrationStop(f"{codec.TAB}: 同じ REQ_ID が重なっています: {dup[:20]}")
    bad_status = Counter(r["status"] for r in rows if r["status"] not in LEAVE_STATUSES)
    if bad_status:
        raise MigrationStop(f"{codec.TAB}: 決まりに無い状態があります: {dict(bad_status)}")

    config: dict[str, list[dict]] = {}
    valid: dict[str, int] = {}
    for tab, table in CONFIG_TABLES.items():
        if tab in reads["tabs"]:
            cu, cf = reads["tabs"][tab]
            config[table], valid[tab] = config_sync.build_rows(tab, cu, cf, tz)
        else:
            config[table], valid[tab] = [], 0
    settings_parsed, overlaps = config_sync.parse_settings(*reads["tabs"][TAB_SETTINGS])
    empty = [n for n, c in ((codec.TAB, len(rows)), (TAB_SETTINGS, len(settings_parsed))) if c == 0]
    if empty:
        raise MigrationStop(f"{'・'.join(empty)} が 0 行です（空のシートを移しません。罠 49）")

    warnings = []
    if overlaps:
        warnings.append(f"M_SYSTEM_SETTING: 同じキーが 2 行以上あります（下の行を使います）: {overlaps}")
    warnings.extend(config_sync.settings_warnings(overlaps))
    for tab, n in valid.items():
        if n == 0:
            warnings.append(f"{tab}: 有効な行がありません（表は空で入れます）")
    fy_mismatch = []
    for r in rows:
        ld = r.get("leave_date")
        if isinstance(ld, date) and js_string(r.get("fiscal_year")) != str(timeutil.fy_start_0316(ld).year):
            fy_mismatch.append({"reqId": r["req_id"], "fiscalYear": js_string(r.get("fiscal_year")),
                                "computed": timeutil.fy_start_0316(ld).year})
    return {"rows": rows, "raws": d["raws"], "config": config, "tz": reads["tz"], "warnings": warnings,
            "settingsMap": {k: v for k, (v, _r) in settings_parsed.items()}, "settingsOverlaps": overlaps,
            "droppedColumns": d["dropped"], "skippedBlankIdRows": d["skippedBlankIdRows"],
            "emptyRows": d["emptyRows"], "fiscalYearMismatch": fy_mismatch, "validConfigRows": valid}


# ===== 照合の数（シートから読んだ行と DB から読み直した行の両方に同じ関数を使う） =====


def _get(row: dict, name: str) -> Any:
    """仕様の属性名の dict でも、実物の列の名前の dict でも読む。"""
    real = codec.model_attr(name)
    if real is not None and real in row and real != name:
        return row[real]
    return row.get(name, codec.get_attr(row, name))


def _sec(v: Any) -> str:
    if isinstance(v, datetime):
        v = v if v.tzinfo else v.replace(tzinfo=timeutil.UTC)
        return v.astimezone(timeutil.UTC).replace(microsecond=0).isoformat()
    return js_string(v or "")


def _day(v: Any, text: Any) -> str:
    if isinstance(v, date):
        return v.isoformat()
    return js_string(text or "")


def summarize(rows: list[dict]) -> dict:
    by_id = {js_string(_get(r, "req_id")): r for r in rows}
    return {
        "count": len(rows),
        "statusCounts": dict(sorted(Counter(js_string(_get(r, "status")) for r in rows).items())),
        "approved2Count": sum(1 for r in rows if _get(r, "approved_at2") is not None
                              or js_string(_get(r, "approved_at2_text") or "")),
        "leaveDates": {k: _day(_get(r, "leave_date"), _get(r, "leave_date_text")) for k, r in sorted(by_id.items())},
        "substituteDates": {k: _day(_get(r, "substitute_date"), _get(r, "substitute_date_text"))
                            for k, r in sorted(by_id.items())},
        "appliedAt": {k: _sec(_get(r, "applied_at")) for k, r in sorted(by_id.items())},
        "approvedAt": {k: _sec(_get(r, "approved_at")) for k, r in sorted(by_id.items())},
        "approvedAt2": {k: _sec(_get(r, "approved_at2")) for k, r in sorted(by_id.items())},
        "workerIds": {k: normalize_gas(_get(r, "worker_id")) for k, r in sorted(by_id.items())},
        "fiscalYears": {k: js_string(_get(r, "fiscal_year")) for k, r in sorted(by_id.items())},
        "sheetRows": {k: _get(r, "sheet_row") for k, r in sorted(by_id.items())},
    }


def _brief(s: dict) -> dict:
    return {"count": s["count"], "statusCounts": s["statusCounts"], "approved2Count": s["approved2Count"]}


# ===== DB への書き込み =====


def _metadata_tables() -> dict:
    from app import models  # noqa: F401 - すべての表を読み込む
    from app.db import Base

    return {t: Base.metadata.tables[t] for t in TRUNCATE_TABLES}


def build_inserts(prepared: dict, app_env: str) -> tuple[list[dict], list[dict], int]:
    """leave_requests の行（実物の列の名前）・sheet_sync_state の行・ステージングで空にした行の数。"""
    rows, states, blanked = [], [], 0
    cols = codec._leave_columns()
    for rec in prepared["rows"]:
        record = dict(rec)
        columns = codec.to_columns(record)
        synced = codec.synced_values_for(columns, prepared["raws"].get(rec["req_id"]))
        if app_env == "staging":
            originals = {"pdf_url": record.get("pdf_url", ""), "sign_image_url": record.get("sign_image_url", ""),
                         "pdf_file_id": record.get("pdf_file_id", "")}
            if any(originals.values()):
                blanked += 1
            for k, v in originals.items():
                columns[k] = ""
                if "source_" + k in cols:
                    columns["source_" + k] = v
        rows.append(columns)
        states.append({"table_name": codec.TABLE, "pk": str(rec["req_id"]), "desired_version": 1, "synced_version": 1,
                       "synced_values": synced, "state": "ok", "attempts": 0, "next_try_at": None, "last_error": "",
                       "first_dirty_at": None, "last_synced_at": None, "seq": 0})
    return rows, states, blanked


def insert_tables(engine: Any, rows: list[dict], states: list[dict], config: dict, *, truncate: bool) -> None:
    """1 つのトランザクションで入れる。表に行があれば止める（--truncate なら先に空にする）。"""
    from sqlalchemy import delete, func, select, text

    targets = _metadata_tables()
    with engine.begin() as conn:
        counts = {t: conn.execute(select(func.count()).select_from(targets[t])).scalar_one() for t in TRUNCATE_TABLES}
        not_empty = {t: n for t, n in counts.items() if n}
        if not_empty and not truncate:
            raise MigrationStop("移行先の DB が空ではありません（--truncate を付けると、業務・設定とマスタの表・jobs・"
                                "idempotency_keys・sync_runs・sessions を空にしてから入れます）: "
                                + ", ".join(f"{t}={n}" for t, n in not_empty.items()))
        if truncate:
            if conn.dialect.name == "postgresql":
                conn.execute(text("TRUNCATE TABLE " + ", ".join(TRUNCATE_TABLES) + " RESTART IDENTITY"))
            else:
                for t in TRUNCATE_TABLES:
                    conn.execute(delete(targets[t]))
        if rows:
            conn.execute(targets["leave_requests"].insert(), rows)
        for tab, table in CONFIG_TABLES.items():
            fitted = config_sync.fitted(tab, config.get(table, []))
            if fitted:
                conn.execute(targets[table].insert(), fitted)
        if states:
            conn.execute(targets["sheet_sync_state"].insert(), states)
        if conn.dialect.name == "postgresql":
            conn.execute(text("SELECT setval('sheet_sync_state_seq', "
                              "COALESCE((SELECT max(seq) FROM sheet_sync_state), 0) + 1, false)"))


def verify_in_db(engine: Any, prepared: dict, reads: dict, app_env: str) -> tuple[list, dict]:
    """入れた後に DB から読み直し、照合の数と 1 セルずつの照合（差 0）と集計の照合を確かめる。"""
    from sqlalchemy import select
    from sqlalchemy.orm import Session

    problems: list[str] = []
    targets = _metadata_tables()
    with engine.connect() as conn:
        db_rows = [dict(r._mapping) for r in conn.execute(select(targets["leave_requests"]))]
    expected, actual = summarize(prepared["rows"]), summarize(db_rows)
    for key in expected:
        if expected[key] != actual[key]:
            bad = ([k for k in expected[key] if expected[key].get(k) != actual[key].get(k)][:10]
                   if isinstance(expected[key], dict) else [])
            problems.append(f"{key}: シートと DB が違います {bad or (expected[key], actual[key])}")
    tz = codec.as_tz(reads["tz"] or "Etc/GMT")
    u, f = reads["tabs"][codec.TAB]
    detail: dict = {}
    with Session(engine) as db:
        result = reconcile.compare_rows(u, f, reconcile.db_rows_for(db), tz)
        n = len(result["diffs"]) + len(result["idMismatch"])
        if app_env != "staging":
            n += len(result["gDiffs"])
        detail["postMigrationDiffs"] = n
        detail["extraRows"] = len(result["extraRows"])
        if n:
            problems.append(f"移した直後の照合で差があります: {result['diffs'][:5]} {result['gDiffs'][:3]} "
                            f"{result['idMismatch'][:5]}")
        if TAB_LEAVE_SUMMARY in reads["tabs"]:
            try:
                rows = summary_sheet.summary_rows(db, timeutil.now_utc())
            except ImportError:
                detail["summary"] = "W の domain.summaries がまだ無いため、集計の照合は飛ばしました"
            else:
                diffs = summary_sheet.compare(rows, reads["tabs"][TAB_LEAVE_SUMMARY][0], tz)
                detail["summaryDiffs"] = len(diffs)
                if diffs:
                    problems.append(f"V_LEAVE_SUMMARY と DB から作った集計が違います: {diffs[:5]}")
    return problems, detail


# ===== 実行 =====


def run(reads: dict, engine: Optional[Any], *, dry_run: bool, truncate: bool = False,
        expected_tz: str | None = None, app_env: str = "production") -> dict:
    if expected_tz and reads["tz"] != expected_tz:
        raise MigrationStop(f"DB スプレッドシートの時間帯（{reads['tz']}）が DB_SHEET_TIMEZONE（{expected_tz}）と違います")
    prepared = prepare(reads)
    rows, states, blanked = build_inserts(prepared, app_env)
    kinds = Counter(f"{k}:{v}" for r in prepared["rows"] for k, v in (r.get("cell_kinds") or {}).items())
    report = {
        "dryRun": dry_run,
        "appEnv": app_env,
        "sheetTimeZone": reads["tz"],
        "summary": _brief(summarize(prepared["rows"])),
        "cellKinds": dict(sorted(kinds.items())),
        "warnings": prepared["warnings"],
        "missingColumns": [],
        "droppedColumns": prepared["droppedColumns"],
        "skippedBlankIdRows": prepared["skippedBlankIdRows"],
        "emptyRows": prepared["emptyRows"],
        "fiscalYearMismatchCount": len(prepared["fiscalYearMismatch"]),
        "fiscalYearMismatch": prepared["fiscalYearMismatch"][:50],
        "settingsOverlaps": prepared["settingsOverlaps"],
        "settingsForEnv": {k: prepared["settingsMap"].get(k, "") for k in REPORTED_SETTING_KEYS},
        "configCounts": {t: len(v) for t, v in prepared["config"].items()},
        "absentTabs": reads.get("absentTabs", []),
        "stagingPdfAndSignBlanked": blanked if app_env == "staging" else None,
    }
    if not dry_run:
        if engine is None:
            raise MigrationStop("移行先の DB（--database-url か DATABASE_URL）が未設定です。")
        insert_tables(engine, rows, states, prepared["config"], truncate=truncate)
        report["verifyProblems"], detail = verify_in_db(engine, prepared, reads, app_env)
        report.update(detail)
    return report


def _json_default(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value)


def _sheets_api(engine: Any, settings: Any) -> Any:
    from sqlalchemy.orm import Session

    from app.services.sheets_api import SheetsApi

    with Session(engine) as db:
        return SheetsApi.for_db(db, settings)


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="本番 DB スプレッドシート → DB の移行（読むだけ）")
    parser.add_argument("--source", choices=["sheets"], required=True)
    parser.add_argument("--spreadsheet-id", default=PROD_SPREADSHEET_ID_FIXED)
    parser.add_argument("--database-url", default=os.environ.get("DATABASE_URL", ""))
    parser.add_argument("--dry-run", action="store_true", help="確かめと照合の報告だけ。DB に書かない")
    parser.add_argument("--truncate", action="store_true",
                        help="移行先の表と jobs・idempotency_keys・sync_runs・sessions を空にしてから入れる")
    parser.add_argument("--report-out", help="報告の JSON の書き出し先")
    parser.add_argument("--snapshot-out", help="読んだシートの値（UNFORMATTED・FORMATTED）の JSON の書き出し先")
    args = parser.parse_args(argv)
    try:
        check_source_id(args.spreadsheet_id)  # トークンを取る前に止める
        if not args.database_url:
            raise MigrationStop("移行先の DB（--database-url か DATABASE_URL）が未設定です"
                                "（委任トークンを DB から読むので --dry-run でも要ります）。")
        from app.config import get_settings
        from app.db import make_engine

        settings = get_settings()
        engine = make_engine(args.database_url)
        reads = read_sheets(_sheets_api(engine, settings), args.spreadsheet_id)
        if args.snapshot_out:
            Path(args.snapshot_out).write_text(json.dumps(reads, ensure_ascii=False, default=_json_default),
                                               encoding="utf-8")
        report = run(reads, None if args.dry_run else engine, dry_run=args.dry_run, truncate=args.truncate,
                     expected_tz=settings.db_sheet_timezone, app_env=settings.app_env)
    except (MigrationStop, SheetsError) as error:
        print(f"止めました: {error}", file=sys.stderr)
        return 2
    except Exception as error:  # noqa: BLE001 - 委任が無いなど
        print(f"止めました: {type(error).__name__}: {error}", file=sys.stderr)
        return 2
    text = json.dumps(report, ensure_ascii=False, indent=1, default=_json_default)
    if args.report_out:
        Path(args.report_out).write_text(text, encoding="utf-8")
    print(text)
    return 1 if report.get("verifyProblems") else 0


if __name__ == "__main__":
    sys.exit(main())
