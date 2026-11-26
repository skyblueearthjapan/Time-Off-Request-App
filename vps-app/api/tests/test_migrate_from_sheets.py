"""tools/migrate_from_sheets.py: 本番 ID 以外を拒む・REQ_ID が空の行を飛ばして一覧に出す・重なりで止まる・年度と古い型を保つ・
staging で PDF とサインの URL を空・pdf_file_id の取り出し・移した直後の照合が差 0・空の DB でなければ止まる。"""
from __future__ import annotations

from datetime import date

import pytest
from sqlalchemy import select

from app.models import CalendarDay, DeptApprover, LeaveRequest, Setting, SheetSyncState, Stamp, Worker
from d_sheet_fixtures import (
    HEADERS,
    LEAVE_ROWS,
    PROD_SID,
    R_CANCEL,
    R_HALF,
    R_OLD,
    STAGING_SID,
    install_fake_summaries,
    leave_row,
    make_fake,
    migrate,
)

import migrate_from_sheets as tool  # noqa: E402 - d_sheet_fixtures が tools を読み込み先に足す


@pytest.fixture(autouse=True)
def _summaries(monkeypatch):
    install_fake_summaries(monkeypatch)


def test_refuses_other_spreadsheets(settings):
    fake = make_fake()
    with pytest.raises(tool.MigrationStop):
        tool.read_sheets(fake.api(settings), STAGING_SID)
    assert tool.main(["--source", "sheets", "--spreadsheet-id", STAGING_SID, "--dry-run"]) == 2
    assert fake.calls == []


def test_migrates_and_verifies_with_zero_diffs(engine, db, settings):
    fake = make_fake()
    report = migrate(fake, engine, settings, app_env="production")
    assert report["verifyProblems"] == [] and report["postMigrationDiffs"] == 0
    assert report["summaryDiffs"] == 0  # DB から作った集計と V_LEAVE_SUMMARY が 12 列とも同じ
    assert report["summary"]["count"] == 4
    assert report["summary"]["statusCounts"] == {"取消": 1, "承認済": 2, "申請中": 1}
    assert report["summary"]["approved2Count"] == 1
    # REQ_ID が空の行は飛ばして、行番号と値のあるセルを出す
    assert report["skippedBlankIdRows"] == [{"row": 4, "cells": {"WORKER_NAME": "メモ"}}]
    # 年度は計算し直さない（4/1 区切りの頃の行）
    assert report["fiscalYearMismatchCount"] == 1 and report["fiscalYearMismatch"][0]["reqId"] == R_OLD
    old = db.get(LeaveRequest, R_OLD)
    assert old.fiscal_year == "2025" and old.sheet_row == 2 and old.row_version == 1
    assert old.leave_date == date(2026, 3, 20)  # GMT 0 時＝JST 9 時の整数の行も同じ日
    assert old.cell_kinds["leave_date"] == "serial" and old.cell_kinds["worker_id"] == "number"
    assert old.pdf_url.startswith("https://drive.google.com/") and old.pdf_file_id == "PDFfile_ID-1"
    half = db.get(LeaveRequest, R_HALF)
    assert half.leave_date == date(2026, 11, 27) and half.sheet_row == 3 and half.sign_image_url
    assert db.get(LeaveRequest, R_CANCEL).sheet_row == 6
    st = db.get(SheetSyncState, ("leave_requests", R_OLD))
    assert st.desired_version == st.synced_version == 1
    assert st.synced_values["__raw__"]["LEAVE_DATE"] == LEAVE_ROWS[0][HEADERS.index("LEAVE_DATE")]
    # 設定とマスタの表
    assert db.get(Setting, "ADMIN_EMAILS").value == "admin@lineworks-local.info"
    assert [w.cells["氏名"] for w in db.scalars(select(Worker).order_by(Worker.sort_order))] == ["試験 一郎", "試験 二郎"]
    assert db.scalars(select(CalendarDay)).first().date == date(2026, 11, 28)
    assert db.get(Stamp, 1).col_c == "総務 花子"
    assert db.get(DeptApprover, 1).col_a == "機械加工"
    # シートは読むだけ
    assert all(m == "GET" for m, *_ in fake.calls)


def test_staging_blanks_pdf_and_sign(engine, db, settings):
    fake = make_fake()
    report = migrate(fake, engine, settings, app_env="staging")
    assert report["stagingPdfAndSignBlanked"] == 2 and report["verifyProblems"] == []
    for req_id in (R_OLD, R_HALF):
        row = db.get(LeaveRequest, req_id)
        assert row.pdf_url == "" and row.sign_image_url == "" and row.pdf_file_id == ""


def test_stops_on_duplicate_or_bad_ids_and_statuses(engine, settings):
    for bad in (LEAVE_ROWS + [LEAVE_ROWS[0]], LEAVE_ROWS + [leave_row(REQ_ID="X-1", STATUS="申請中")],
                LEAVE_ROWS + [leave_row(REQ_ID="LV-20261126-009", STATUS="保留")]):
        fake = make_fake(leave_rows=bad)
        reads = tool.read_sheets(fake.api(settings), PROD_SID)
        with pytest.raises(tool.MigrationStop):
            tool.run(reads, engine, dry_run=True, expected_tz="Etc/GMT")


def test_stops_on_missing_header_and_timezone(engine, settings):
    fake = make_fake()
    reads = tool.read_sheets(fake.api(settings), PROD_SID)
    with pytest.raises(tool.MigrationStop):
        tool.run(reads, engine, dry_run=True, expected_tz="Asia/Tokyo")
    u, f = reads["tabs"]["T_LEAVE_REQUEST"]
    u[0] = [h if h != "APPROVED_AT2" else "" for h in u[0]]
    with pytest.raises(tool.MigrationStop, match="missingColumns"):
        tool.run(reads, engine, dry_run=True, expected_tz="Etc/GMT")


def test_stops_when_empty_sheet_or_db_not_empty(engine, settings):
    fake = make_fake(leave_rows=[])
    reads = tool.read_sheets(fake.api(settings), PROD_SID)
    with pytest.raises(tool.MigrationStop, match="0 行"):
        tool.run(reads, engine, dry_run=True, expected_tz="Etc/GMT")
    fake = make_fake()
    migrate(fake, engine, settings)
    with pytest.raises(tool.MigrationStop, match="空ではありません"):
        migrate(fake, engine, settings)


def test_dry_run_writes_nothing(engine, db, settings):
    fake = make_fake()
    reads = tool.read_sheets(fake.api(settings), PROD_SID)
    report = tool.run(reads, engine, dry_run=True, expected_tz="Etc/GMT")
    assert report["dryRun"] and "verifyProblems" not in report
    assert db.scalars(select(LeaveRequest)).first() is None
