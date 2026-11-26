"""担当 D の試験の材料: 本番に似せた DB スプレッドシート（06 2-5・5 節の実物の形）と、移行の道具での取り込み。

- 時間帯は Etc/GMT（地域 en_US）。T_LEAVE_REQUEST は英語の見出し 26 列（本番の並び）。
- 休暇日は「JST 0 時」＝GMT の前日 15:00（.625）の行と、古い「GMT 0 時」＝JST 9 時（整数）の行がある。
- REQ_ID が空で 1 セルだけ値がある半端な行が途中にある（4 行目）。
- WORKER_ID・FY_START_YEAR は数のセル。日時は時刻つきのシリアル値。APPROVED_AT2 の表示は M/d/yyyy。
"""
from __future__ import annotations

import sys
from datetime import date, datetime, timezone
from pathlib import Path

from app import config as _config
from app.jsutil import js_sort_key, normalize_gas
from app.services import sheet_codec as codec
from d_fake_sheets import FakeSheets

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))

PROD_SID = _config.PROD_SPREADSHEET_ID_FIXED
STAGING_SID = "staging-db-sheet-id"
GMT = "Etc/GMT"
HEADERS = list(codec.HEADERS)
DATE_FMT = {"type": "DATE", "pattern": "yyyy-mm-dd"}
DT_FMT = {"type": "DATE_TIME", "pattern": "yyyy-mm-dd hh:mm"}
AT2_FMT = {"type": "DATE", "pattern": "M/d/yyyy"}

R_OLD = "LV-20260305-001"  # 古い行: 年度 2025（4/1 区切り）、休暇日は GMT 0 時＝JST 9 時の整数
R_HALF = "LV-20261120-001"  # 半日の有給、休暇日 .625、サイン画像 URL あり
R_SUB = "LV-20261125-001"  # 振替、申請中
R_CANCEL = "LV-20261126-001"  # 取消
PDF_URL = "https://drive.google.com/file/d/PDFfile_ID-1/view?usp=drivesdk"
SIGN_URL = "https://drive.google.com/file/d/SIGNfile-2/view"


def jst_day(d: date) -> float:
    """今の GAS が書く休暇日（JST 0 時）の GMT のシリアル値（.625）。"""
    return codec.date_to_serial(d, GMT)


def old_day(d: date) -> float:
    """古い行の休暇日（new Date('yyyy-MM-dd')＝UTC 0 時）の GMT のシリアル値（整数）。"""
    return float((d - codec.EPOCH.date()).days)


def at(y, mo, d, h, mi, s=0) -> float:
    """JST の日時 → GMT のシリアル値。"""
    return codec.datetime_to_serial(datetime(y, mo, d, h, mi, s, tzinfo=codec.JST), GMT)


def leave_row(**kw) -> list:
    row = [""] * len(HEADERS)
    for h, v in kw.items():
        row[HEADERS.index(h)] = v
    return row


LEAVE_ROWS = [
    leave_row(REQ_ID=R_OLD, FY_START_YEAR=2025, STATUS="承認済", APPLIED_AT=at(2026, 3, 5, 10, 0), DEPT_ID="機械加工",
              DEPT_NAME="機械加工", WORKER_ID=1001, WORKER_NAME="試験 一郎", LEAVE_DATE=old_day(date(2026, 3, 20)),
              DAY_TYPE="全日休", LEAVE_KIND="有給消化（私用）", APPROVED_AT=at(2026, 3, 6, 9, 0), PDF_URL=PDF_URL,
              CREATED_BY_EMAIL="shared@lineworks-local.info", APPROVED_BY1_EMAIL="boss@lineworks-local.info",
              APPROVED_BY2="総務 花子", APPROVED_BY2_EMAIL="somu@lineworks-local.info",
              APPROVED_AT2=at(2026, 3, 7, 11, 30)),
    leave_row(REQ_ID=R_HALF, FY_START_YEAR=2026, STATUS="承認済", APPLIED_AT=at(2026, 11, 20, 8, 15, 12),
              DEPT_ID="機械加工", DEPT_NAME="機械加工", WORKER_ID=1001, WORKER_NAME="試験 一郎",
              LEAVE_DATE=jst_day(date(2026, 11, 27)), DAY_TYPE="半日休", HALF_TYPE="午後", LEAVE_KIND="有給消化（私用）",
              PAIDLEAVE_DETAIL="私用のため\n午後", APPROVED_AT=at(2026, 11, 21, 17, 5), SIGN_IMAGE_URL=SIGN_URL,
              CREATED_BY_EMAIL="shared@lineworks-local.info"),
    leave_row(WORKER_NAME="メモ"),  # REQ_ID が空の半端な行（本番の 11 行目あたりの形）
    leave_row(REQ_ID=R_SUB, FY_START_YEAR=2026, STATUS="申請中", APPLIED_AT=at(2026, 11, 25, 7, 0), DEPT_ID="組立",
              DEPT_NAME="組立", WORKER_ID=1002, WORKER_NAME="試験 二郎", LEAVE_DATE=jst_day(date(2026, 12, 1)),
              DAY_TYPE="全日休", LEAVE_KIND="振替休暇", SUBSTITUTE_FOR_WORK_DATE=jst_day(date(2026, 11, 21)),
              CREATED_BY_EMAIL="shared@lineworks-local.info"),
    leave_row(REQ_ID=R_CANCEL, FY_START_YEAR=2026, STATUS="取消", APPLIED_AT=at(2026, 11, 26, 10, 12), DEPT_ID="組立",
              DEPT_NAME="組立", WORKER_ID=1002, WORKER_NAME="試験 二郎", LEAVE_DATE=jst_day(date(2026, 12, 3)),
              DAY_TYPE="全日休", LEAVE_KIND="特別休暇", SPECIAL_REASON_TEXT="=SUM(A1)",
              CREATED_BY_EMAIL="shared@lineworks-local.info"),
]


def leave_formats(rows: list[list]) -> dict:
    fmts = {}
    for i, row in enumerate(rows, start=1):
        for h, fmt in (("APPLIED_AT", DT_FMT), ("APPROVED_AT", DT_FMT), ("LEAVE_DATE", DATE_FMT),
                       ("SUBSTITUTE_FOR_WORK_DATE", DATE_FMT), ("APPROVED_AT2", AT2_FMT)):
            if row[HEADERS.index(h)] != "":
                fmts[(i, HEADERS.index(h))] = dict(fmt)
    return fmts


SUMMARY_HEADER = ["年度", "部署ID", "部署名", "作業員ID", "作業員名", "有給回数", "有給半日回数", "振替回数", "特別回数",
                  "合計回数", "最終有給日", "警告レベル"]
SUMMARY_ROWS = [
    [2025, "機械加工", "機械加工", 1001, "試験 一郎", 1, 0, 0, 0, 1, jst_day(date(2026, 3, 20)), ""],
    [2026, "機械加工", "機械加工", 1001, "試験 一郎", 0.5, 1, 0, 0, 1, jst_day(date(2026, 11, 27)), ""],
]


def config_tabs() -> dict:
    return {
        "M_SYSTEM_SETTING": [["KEY", "VALUE", "MEMO"], ["APP_URL", "https://script.google.com/x/exec", ""],
                             ["ADMIN_EMAILS", "admin@lineworks-local.info", ""],
                             ["SOMU_EMAILS", "somu@lineworks-local.info", ""],
                             ["PDF_FOLDER_ID", "staging-pdf-folder-id", ""], ["LEDGER_SSID", "staging-ledger-id", ""],
                             ["MASTER_SOURCE_WORKER_SHEET", "作業員マスタ", ""]],
        "M_LOOKUP": [["カテゴリ", "値", "SORT"], ["特別休暇理由", "結婚休暇", 1], ["特別休暇理由", "忌引休暇", 2]],
        "作業員マスタ": [["作業員コード", "氏名", "部署", "担当業務", "拠点", "スタッフ種類"],
                       [1001, "試験 一郎", "機械加工", "旋盤", "本社", "社員"],
                       [1002, "試験 二郎", "組立", "組立", "本社", "社員"]],
        "M_CALENDAR": [["日付", "区分", "曜日", "備考"], [old_day(date(2026, 11, 28)), "休日", "Sat", ""],
                       [old_day(date(2026, 11, 23)), "祝日", "Mon", "勤労感謝の日"]],
        "M_STAMP": [["メール", "stampFileId", "備考"], ["somu@lineworks-local.info", "stamp-1", "総務 花子"]],
        "M_DEPT_APPROVERS": [["dept", "approverEmails", " approverName", "役職", "通知メール"],
                             ["機械加工", "boss@lineworks-local.info", "上長 太郎", "課長", ""]],
    }


def add_db_book(fake: FakeSheets, sid: str, *, leave_rows: list[list] | None = None) -> None:
    rows = LEAVE_ROWS if leave_rows is None else leave_rows
    book = fake.add_book(sid, tz=GMT)
    book.title = "休暇届けapp"
    book.add_tab("T_LEAVE_REQUEST", [HEADERS] + [list(r) for r in rows], formats=leave_formats(rows),
                 row_count=len(rows) + 3, col_count=26)
    for title, values in config_tabs().items():
        fmts = {(r, 0): dict(DATE_FMT) for r in range(1, len(values))} if title == "M_CALENDAR" else None
        book.add_tab(title, values, formats=fmts)
    book.add_tab("V_LEAVE_SUMMARY", [SUMMARY_HEADER] + [list(r) for r in SUMMARY_ROWS],
                 formats={(r, 10): dict(DATE_FMT) for r in range(1, len(SUMMARY_ROWS) + 1)})
    book.add_tab("PDF_TEMPLATE", [["実行日時"]])


def make_fake(*, leave_rows: list[list] | None = None) -> FakeSheets:
    fake = FakeSheets()
    add_db_book(fake, PROD_SID, leave_rows=leave_rows)
    add_db_book(fake, STAGING_SID, leave_rows=leave_rows)
    return fake


def migrate(fake: FakeSheets, engine, settings, *, app_env: str | None = None) -> dict:
    """移行の道具で本番の写し（PROD_SID の本）を DB に入れる。"""
    import migrate_from_sheets as tool

    reads = tool.read_sheets(fake.api(settings), PROD_SID)
    return tool.run(reads, engine, dry_run=False, expected_tz=GMT, app_env=app_env or settings.app_env)


# ===== W の domain.summaries の偽物（実装仕様 10-2 の形。数え方は 6-3 の要点だけ） =====


def fake_summary_rows(db, now):
    from sqlalchemy import select

    from app.models import LeaveRequest

    out: dict[str, dict] = {}
    order = []
    for row in db.scalars(select(LeaveRequest).order_by(LeaveRequest.sheet_row)):
        if normalize_gas(row.status) != "承認済":
            continue
        fy = row.fiscal_year
        if (row.cell_kinds or {}).get("fiscal_year") == "number":
            fy = int(float(fy))  # 元のセルの型（数なら数）
        key = f"{fy}_{normalize_gas(row.worker_id)}"
        s = out.get(key)
        if s is None:
            s = out[key] = {"fiscalYear": fy, "deptId": row.dept_id, "deptName": row.dept_name,
                            "workerId": normalize_gas(row.worker_id), "workerName": row.worker_name, "paidCount": 0,
                            "paidHalfCount": 0, "substituteCount": 0, "specialCount": 0, "totalCount": 0,
                            "lastPaidDate": "", "warnLevel": ""}
            order.append(key)
        s["totalCount"] += 1
        kind = codec.get_attr(row, "leave_type")
        if kind == "有給消化（私用）":
            if codec.get_attr(row, "leave_kubun") == "半日休":
                s["paidHalfCount"] += 1
                s["paidCount"] += 0.5
            else:
                s["paidCount"] += 1
            if row.leave_date and (row.cell_kinds or {}).get("leave_date") != "string":
                ld = row.leave_date.isoformat()
                if not s["lastPaidDate"] or ld > s["lastPaidDate"]:
                    s["lastPaidDate"] = ld
        elif kind == "振替休暇":
            s["substituteCount"] += 1
        elif kind == "特別休暇":
            s["specialCount"] += 1
    return [out[k] for k in sorted(out, key=js_sort_key)]


def install_fake_summaries(monkeypatch) -> None:
    import types

    module = types.ModuleType("app.domain.summaries")
    module.summary_rows = fake_summary_rows
    monkeypatch.setitem(sys.modules, "app.domain.summaries", module)
    import app.domain as domain

    monkeypatch.setattr(domain, "summaries", module, raising=False)


NOW = datetime(2026, 11, 26, 3, 0, tzinfo=timezone.utc)
