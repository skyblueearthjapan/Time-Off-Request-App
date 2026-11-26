"""services/config_sync: 6 タブ・同じキーは下の行が勝つ（空でも勝つ）・空なら上書きしない・.env との食い違い・読むだけ。"""
from __future__ import annotations

from datetime import date

from sqlalchemy import select

from app.models import CalendarDay, DeptApprover, Lookup, Setting, Stamp, SyncRun, Worker
from app.services import config_sync
from d_sheet_fixtures import STAGING_SID, make_fake


def run(db, fake, settings):
    result = config_sync.sync_sheet_config(db, api=fake.api(settings), app_settings=settings)
    db.commit()
    return result


def test_imports_six_tabs_read_only(db, settings):
    fake = make_fake()
    result = run(db, fake, settings)
    assert result["ok"], result
    assert set(result["counts"]) == set(config_sync.TABS)
    assert db.get(Setting, "SOMU_EMAILS").value == "somu@lineworks-local.info"
    assert [lk.value for lk in db.scalars(select(Lookup).order_by(Lookup.sheet_row))] == ["結婚休暇", "忌引休暇"]
    w = db.get(Worker, 1)
    assert w.cells == [["作業員コード", 1001], ["氏名", "試験 一郎"], ["部署", "機械加工"], ["担当業務", "旋盤"],
                       ["拠点", "本社"], ["スタッフ種類", "社員"]]
    cal = db.get(CalendarDay, 1)
    assert cal.date == date(2026, 11, 28) and cal.kubun == "休日" and dict(cal.cells)["曜日"] == "Sat"
    assert (db.get(Stamp, 1).col_a, db.get(Stamp, 1).col_b) == ("somu@lineworks-local.info", "stamp-1")
    assert db.get(DeptApprover, 1).col_c == "上長 太郎"
    assert result["mismatches"] == []
    assert all(m == "GET" for m, *_ in fake.calls)
    assert db.scalars(select(SyncRun).where(SyncRun.kind == "sync_sheet_config")).one().ok


def test_lower_row_wins_even_when_empty(db, settings):
    fake = make_fake()
    tab = fake.books[STAGING_SID].tabs["M_SYSTEM_SETTING"]
    tab.cell(7, 0).value = " SOMU_EMAILS "
    tab.cell(7, 1).value = ""
    tab.cell(8, 0).value = "LEDGER_SSID"
    tab.cell(8, 1).value = "other-ledger"
    result = run(db, fake, settings)
    assert db.get(Setting, "SOMU_EMAILS").value == "" and db.get(Setting, "SOMU_EMAILS").sheet_row == 8
    keys = {o["key"] for o in result["overlaps"]}
    assert keys == {"SOMU_EMAILS", "LEDGER_SSID"}
    assert any("LEDGER_SSID" in w for w in result["warnings"])
    assert {"key": "LEDGER_SSID", "sheet": "other-ledger", "env": "staging-ledger-id"} in result["mismatches"]


def test_empty_or_missing_tabs_keep_current_rows(db, settings):
    run(db, make_fake(), settings)
    fake = make_fake()
    book = fake.books[STAGING_SID]
    book.tabs["M_STAMP"].rows = book.tabs["M_STAMP"].rows[:1]  # 見出しだけ
    del book.tabs["M_LOOKUP"]
    result = run(db, fake, settings)
    assert result["ok"] is False
    assert "M_STAMP" in result["skipped"] and "M_LOOKUP" in result["skipped"]
    assert db.get(Stamp, 1).col_a == "somu@lineworks-local.info"
    assert db.scalars(select(Lookup)).first() is not None


def test_workers_keep_sheet_column_order_and_duplicate_headers():
    """作業員マスタは列の順のまま [[見出し, 値], …] で持つ（同じ見出し・見出しの無い列も残す）。"""
    u = [["氏名", "作業員コード", "部署", "", "部署"], ["甲", 101, "製造1課", "メモ", "品質"], ["乙", 102]]
    rows, valid = config_sync.build_rows("作業員マスタ", u, [], "Asia/Tokyo")
    assert rows[0]["cells"] == [["氏名", "甲"], ["作業員コード", 101], ["部署", "製造1課"], ["", "メモ"], ["部署", "品質"]]
    assert rows[1]["cells"] == [["氏名", "乙"], ["作業員コード", 102], ["部署", ""], ["", ""], ["部署", ""]]
    assert valid == 2


def test_register_worker(settings):
    from app import worker

    config_sync.register_worker()
    assert worker.schedules()["sync_sheet_config"].every_minutes == 5
