"""services/master_sync: 1 列目が空の行を捨てる・見出しの確かめ・位置で読むタブの見出しの違いで止まる・USER_ENTERED・
数式にしない・日付のシリアル値（時間帯の違いは瞬間を保つ）と表示形式・余りは値だけ消す・空なら上書きしない・
タブごとの失敗・読む元は .env の ID だけ（元には書かない）・タブを作らない。"""
from __future__ import annotations

from datetime import date, datetime

import pytest
from sqlalchemy import select

from app import config as _config
from app.models import CalendarDay, DeptApprover, Setting, Stamp, SyncRun, Worker
from app.services import master_sync, write_gate
from app.services import sheet_codec as codec
from d_fake_sheets import FormulaCell
from d_sheet_fixtures import STAGING_SID, make_fake

MASTER = _config.MASTER_SOURCE_SPREADSHEET_ID_FIXED
CAL = _config.CALENDAR_SOURCE_SPREADSHEET_ID_FIXED
DATE_FMT = {"type": "DATE", "pattern": "M/d/yyyy"}


@pytest.fixture(autouse=True)
def _no_sleep(monkeypatch):
    monkeypatch.setattr(master_sync, "sleep", lambda s: None)


def make_sources(fake):
    m = fake.add_book(MASTER, tz="Asia/Tokyo")
    m.add_tab("作業員マスタ", [["作業員コード", "氏名", "部署", "担当業務", "拠点", "スタッフ種類", ""],
                          ["0042", "新人 三郎", "組立", "=手配", "本社", "社員"],
                          ["", "空の行", "", "", "", ""],
                          [1001, "試験 一郎", "機械加工", "旋盤", "本社", "社員"]])
    c = fake.add_book(CAL, tz="Etc/GMT")
    jst_midnight = codec.datetime_to_serial(datetime(2026, 12, 5, tzinfo=codec.JST), "Etc/GMT")
    c.add_tab("社内カレンダーマスタ", [["日付", "区分", "曜日", "備考"], [jst_midnight, "休日", "Sat", ""],
                                 [0, "祝日", "", "0 は空の行"]], formats={(1, 0): dict(DATE_FMT)})
    c.add_tab("StampMap", [["メール", "stampFileId", "備考"], ["somu@lineworks-local.info", "stamp-9", "総務 花子"]])
    c.add_tab("DeptApprovers", [["dept", "approverEmails", " approverName", "役職", "通知メール"],
                                ["組立", "kumi@lineworks-local.info", "組立 長", "係長", ""]])
    return m, c


def run(db, fake, settings):
    result = master_sync.sync_masters(db, api=fake.api(settings), app_settings=settings)
    db.commit()
    return result


def test_syncs_four_tabs_like_gas(db, settings):
    fake = make_fake()
    make_sources(fake)
    book = fake.books[STAGING_SID]
    book.tabs["作業員マスタ"].cell(5, 0).value = "古い余り"
    result = run(db, fake, settings)
    assert result["ok"], result
    assert {k: v["rows"] for k, v in result["tabs"].items()} == {"作業員マスタ": 2, "M_CALENDAR": 1, "M_STAMP": 1,
                                                                "M_DEPT_APPROVERS": 1}
    w = book.tabs["作業員マスタ"].values()
    assert w[1][:4] == [42, "新人 三郎", "組立", "=手配"]  # USER_ENTERED（'0042' は数）、= の文字は文字のまま
    assert not isinstance(book.tabs["作業員マスタ"].get(1, 3).value, FormulaCell)
    assert len(w) == 3  # 1 列目が空の行を捨て、余りの行は値だけ消した
    cal = book.tabs["M_CALENDAR"].get(1, 0)
    assert cal.fmt == DATE_FMT and codec.serial_to_jst_date(cal.value, "Etc/GMT") == date(2026, 12, 5)
    assert book.tabs["M_STAMP"].get(1, 1).value == "stamp-9"
    # DB の表も同じデータで置き換わる
    assert db.get(Worker, 1).cells[0] == ["作業員コード", 42]
    assert db.get(CalendarDay, 1).date == date(2026, 12, 5)
    assert db.get(Stamp, 1).col_b == "stamp-9" and db.get(DeptApprover, 1).col_a == "組立"
    # 元には書かない
    for _m, url, _b in fake.writes():
        assert STAGING_SID in url and MASTER not in url and CAL not in url
    assert db.scalars(select(SyncRun).where(SyncRun.kind == "sync_masters")).one().ok


def test_position_header_mismatch_stops_that_tab_only(db, settings):
    fake = make_fake()
    _m, c = make_sources(fake)
    c.tabs["StampMap"].cell(0, 1).value = "fileId"  # 位置で読む列の見出しが違う
    before = fake.books[STAGING_SID].tabs["M_STAMP"].values()
    result = run(db, fake, settings)
    assert result["ok"] is False
    assert result["tabs"]["M_STAMP"]["ok"] is False and "見出し" in result["tabs"]["M_STAMP"]["skipped_reason"]
    assert result["tabs"]["作業員マスタ"]["ok"] and result["tabs"]["M_DEPT_APPROVERS"]["ok"]
    assert fake.books[STAGING_SID].tabs["M_STAMP"].values() == before


def test_empty_or_header_only_or_unreadable_source_does_not_overwrite(db, settings):
    fake = make_fake()
    m, c = make_sources(fake)
    m.tabs["作業員マスタ"].rows = m.tabs["作業員マスタ"].rows[:1]
    c.tabs["DeptApprovers"].rows = []
    del c.tabs["社内カレンダーマスタ"]
    book = fake.books[STAGING_SID]
    before = {t: book.tabs[t].values() for t in ("作業員マスタ", "M_CALENDAR", "M_DEPT_APPROVERS")}
    result = run(db, fake, settings)
    for t, v in before.items():
        assert book.tabs[t].values() == v and result["tabs"][t]["ok"] is False
    assert result["tabs"]["M_STAMP"]["ok"]


def test_worker_header_check_and_missing_target_tab(db, settings):
    fake = make_fake()
    m, _c = make_sources(fake)
    m.tabs["作業員マスタ"].cell(0, 1).value = "名前"
    del fake.books[STAGING_SID].tabs["M_DEPT_APPROVERS"]
    result = run(db, fake, settings)
    assert "氏名" in result["tabs"]["作業員マスタ"]["skipped_reason"]
    assert "作りません" in result["tabs"]["M_DEPT_APPROVERS"]["skipped_reason"]
    assert "M_DEPT_APPROVERS" not in fake.books[STAGING_SID].tabs


def test_settings_source_ids_are_reported_but_env_is_used(db, settings):
    db.add(Setting(key="CALENDAR_SOURCE_SSID", value="some-other-id", sheet_row=9))
    db.commit()
    fake = make_fake()
    make_sources(fake)
    result = run(db, fake, settings)
    assert result["mismatches"] == [{"key": "CALENDAR_SOURCE_SSID", "sheet": "some-other-id", "env": CAL,
                                     "used": "env"}]
    assert result["tabs"]["M_STAMP"]["ok"]


def test_read_retries_then_fails_one_tab(db, settings, monkeypatch):
    fake = make_fake()
    make_sources(fake)
    calls = {"n": 0}
    real = master_sync.read_source

    def flaky(api, sid, tab):
        if tab == "StampMap":
            calls["n"] += 1
            raise RuntimeError("boom")
        return real(api, sid, tab)

    monkeypatch.setattr(master_sync, "read_source", flaky)
    result = run(db, fake, settings)
    assert result["tabs"]["M_STAMP"]["ok"] is False and result["tabs"]["M_CALENDAR"]["ok"]


def test_read_source_retries_three_times(settings):
    class Api:
        n = 0

        def read_both(self, sid, ranges):
            Api.n += 1
            raise RuntimeError("x")

    with pytest.raises(RuntimeError, match="all retries"):
        master_sync.read_source(Api(), MASTER, "作業員マスタ")
    assert Api.n == 3


def test_gate_closed_writes_nothing(db, settings, monkeypatch):
    monkeypatch.setattr(write_gate, "check", lambda *a, **k: (False, "閉"))
    fake = make_fake()
    make_sources(fake)
    result = run(db, fake, settings)
    assert result["writeBlocked"] == "閉" and fake.calls == []


def test_register_worker():
    from app import worker

    master_sync.register_worker()
    assert worker.schedules()["sync_masters"].at == (6, 0)
