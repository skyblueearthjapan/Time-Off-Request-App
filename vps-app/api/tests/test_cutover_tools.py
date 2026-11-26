"""tools/cutover_backup.py・tools/restore_tabs_from_copy.py: 許可リストの元だけを複製する・行数の照合・
控えから T_LEAVE_REQUEST の V 列だけ（G 列は戻さない）とマスタ・集計を戻す・設定のタブは名前を書いたときだけ・
控えに無い行・ずれた行で止まる・既定は書かない。"""
from __future__ import annotations

import json

import httpx
import pytest

from app import config as _config
from conftest import make_settings
from d_sheet_fixtures import HEADERS, R_SUB, STAGING_SID, make_fake

import cutover_backup  # noqa: E402,I001 - d_sheet_fixtures が tools を読み込み先に足す
import restore_tabs_from_copy as restore  # noqa: E402

COPY_ID = "copy-of-staging-db"
PROD_DEPLOY = "AKfycbTESTprodInternal0123456789abcdefXYZ"


# ===== 控え =====


def drive_transport(seen: list):
    def handler(request: httpx.Request) -> httpx.Response:
        seen.append((request.method, str(request.url)))
        url = str(request.url)
        if url.endswith("/copy?supportsAllDrives=true&fields=id%2Cname"):
            src = url.split("/files/")[1].split("/copy")[0]
            return httpx.Response(200, json={"id": "copy-" + src, "name": json.loads(request.content)["name"]})
        if ":batchGet" in url:
            return httpx.Response(200, json={"valueRanges": [{"values": [["a", "b", "c"]]}]})
        return httpx.Response(200, json={"sheets": [{"properties": {"title": "T_LEAVE_REQUEST"}}]})

    return httpx.MockTransport(handler)


def test_backup_copies_only_allowed_sources(settings):
    seen: list = []
    http = cutover_backup.Http(lambda: "tok", transport=drive_transport(seen))
    report = cutover_backup.run(http, settings)
    assert report["ok"] and [b["sourceId"] for b in report["backups"]] == ["staging-db-sheet-id", "staging-ledger-id"]
    assert all(b["name"].startswith("【切替前の控え】休暇届 ") for b in report["backups"])
    copies = [u for m, u in seen if m == "POST"]
    assert len(copies) == 2 and not any(i in u for u in copies for i in _config.PRODUCTION_IDS)
    with pytest.raises(cutover_backup.BackupStop):
        cutover_backup.copy_file(http, _config.PROD_SPREADSHEET_ID_FIXED, "x", settings)


def test_backup_production_uses_fixed_ids():
    prod = make_settings(app_env="production", staging_spreadsheet_id="", staging_pdf_folder_id="",
                         staging_ledger_spreadsheet_id="", gas_internal_url=f"https://script.google.com/macros/s/{PROD_DEPLOY}/exec",
                         prod_gas_internal_deployment_id=PROD_DEPLOY)
    assert [i for _, i in cutover_backup.allowed_sources(prod)] == [_config.PROD_SPREADSHEET_ID_FIXED,
                                                                   _config.PROD_LEDGER_SPREADSHEET_ID_FIXED]


# ===== 戻し =====


@pytest.fixture
def books(settings):
    fake = make_fake()
    from d_sheet_fixtures import add_db_book

    add_db_book(fake, COPY_ID)
    fake.books[COPY_ID].title = "【切替前の控え】休暇届 DB 2026-11-26 1830"
    return fake


def leave(fake, sid, row, header):
    return fake.books[sid].tabs["T_LEAVE_REQUEST"].get(row - 1, HEADERS.index(header)).value


def test_dry_run_by_default_and_v_columns_only(books, settings):
    target = books.books[STAGING_SID].tabs["T_LEAVE_REQUEST"]
    target.cell(4, HEADERS.index("STATUS")).value = "承認済"  # 切り替えの後に VPS が書いた
    target.cell(4, HEADERS.index("PDF_URL")).value = "https://drive.google.com/file/d/NEWPDF/view"  # GAS が書いた
    report = restore.run(books.api(settings), settings, COPY_ID, "all", apply=False)
    assert report["apply"] is False and books.writes() == []
    assert report["tabs"][0]["tab"] == "T_LEAVE_REQUEST" and report["tabs"][0]["diffCells"] == 1
    restore.run(books.api(settings), settings, COPY_ID, "all", apply=True)
    assert leave(books, STAGING_SID, 5, "STATUS") == "申請中"
    assert leave(books, STAGING_SID, 5, "PDF_URL") == "https://drive.google.com/file/d/NEWPDF/view"  # G 列は戻さない
    assert leave(books, STAGING_SID, 5, "REQ_ID") == R_SUB
    for _m, url, _b in books.writes():
        assert STAGING_SID in url and COPY_ID not in url


def test_settings_tabs_need_explicit_names(books, settings):
    with pytest.raises(restore.RestoreStop):
        restore.run(books.api(settings), settings, COPY_ID, "PDF_TEMPLATE", apply=False)
    report = restore.run(books.api(settings), settings, COPY_ID, "M_SYSTEM_SETTING", apply=False)
    assert report["tabs"][0]["tab"] == "M_SYSTEM_SETTING"


def test_stops_on_rows_not_in_copy_or_moved_rows(books, settings):
    target = books.books[STAGING_SID].tabs["T_LEAVE_REQUEST"]
    target.cell(6, 0).value = "LV-20261127-001"
    with pytest.raises(restore.RestoreStop, match="控えに無い"):
        restore.run(books.api(settings), settings, COPY_ID, "T_LEAVE_REQUEST", apply=True)
    target.rows[4], target.rows[5] = target.rows[5], target.rows[4]
    with pytest.raises(restore.RestoreStop, match="行の位置"):
        restore.run(books.api(settings), settings, COPY_ID, "T_LEAVE_REQUEST", apply=True,
                    drop_rows_not_in_copy=True)
    assert books.writes() == []


def test_refuses_non_backup_or_target_itself(books, settings):
    books.books[COPY_ID].title = "休暇届けapp"
    with pytest.raises(restore.RestoreStop, match="控えの名前"):
        restore.run(books.api(settings), settings, COPY_ID, "all", apply=False)
    with pytest.raises(restore.RestoreStop):
        restore.run(books.api(settings), settings, STAGING_SID, "all", apply=False)


def test_whole_tab_restore_clears_rest(books, settings):
    tab = books.books[STAGING_SID].tabs["M_STAMP"]
    tab.cell(3, 0).value = "後から足した行"
    restore.run(books.api(settings), settings, COPY_ID, "M_STAMP", apply=True)
    assert tab.values() == books.books[COPY_ID].tabs["M_STAMP"].values()
