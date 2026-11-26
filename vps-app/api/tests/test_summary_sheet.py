"""services/summary_sheet: 見出し 12 列・並び・型（年度と作業員 ID と回数は数、最終有給日はいつも整数のシリアル値の日付のセル）・
書いてから余りを消す・0 行で書かない・数式にしない・タブを作らない・予約の間引き・関所。"""
from __future__ import annotations

from datetime import date, timedelta

import pytest
from sqlalchemy import delete, select

from app import jobs
from app.domain.summaries import last_paid_date_display  # 本物（env は偽の summaries を差し込む）
from app.models import Job, LeaveRequest
from app.services import summary_sheet, write_gate
from app.services import sheet_codec as codec
from conftest import freeze_now
from d_sheet_fixtures import NOW, STAGING_SID, SUMMARY_HEADER, install_fake_summaries, make_fake, migrate


@pytest.fixture
def env(engine, settings, monkeypatch):
    freeze_now(monkeypatch, NOW)
    install_fake_summaries(monkeypatch)
    fake = make_fake()
    migrate(fake, engine, settings, app_env="production")
    fake.calls.clear()
    return type("Env", (), {"fake": fake, "tab": fake.books[STAGING_SID].tabs["V_LEAVE_SUMMARY"],
                            "api": fake.api(settings)})


def rows(rows_):
    return [{"fiscalYear": 2026, "deptId": "D", "deptName": "D", "workerId": w, "workerName": "n", "paidCount": p,
             "paidHalfCount": 0, "substituteCount": 0, "specialCount": 0, "totalCount": 1, "lastPaidDate": lp,
             "warnLevel": lv} for w, p, lp, lv in rows_]


def test_writes_12_columns_with_types_and_clears_rest(env, db, settings, monkeypatch):
    env.tab.cell(5, 0).value = "古い行"
    env.tab.cell(0, 13).value = "右の余り"
    monkeypatch.setattr(summary_sheet, "summary_rows", lambda db, now: rows(
        [("1001", 1.5, "2026-11-27", ""), ("0042", 0, "", "=危険"), ("1003", 2, "2026-06-01", "-x")]))
    result = summary_sheet.write(db, api=env.api, app_settings=settings)
    assert result["ok"] and result["rows"] == 3
    values = env.tab.values()
    assert values[0] == SUMMARY_HEADER
    assert values[1][:6] == [2026, "D", "D", 1001, "n", 1.5]  # 年度・作業員 ID・回数は数
    assert values[2][3] == 42  # USER_ENTERED（GAS の setValues と同じく数字だけなら数）
    assert values[2][11] == "=危険" and not isinstance(values[2][11], type(None))
    assert type(values[2][11]).__name__ == "str"  # 式にならない（' を付けて文字）
    assert values[3][11] == "-x"
    # 最終有給日は、今のタブが .625 の形でも、いつも整数のシリアル値の日付のセル（空は空）
    lp = env.tab.get(1, 10)
    assert lp.value == (date(2026, 11, 27) - codec.EPOCH.date()).days and lp.fmt["type"] == "DATE"
    assert codec.serial_to_jst_date(lp.value, "Etc/GMT") == date(2026, 11, 27)
    assert values[2][10] in ("", None)
    assert len(values) == 4 and all(len(r) <= 12 for r in values)  # 余りは消えた
    # 書いてから消す（values.update → updateCells（表示形式）→ batchClear の順）
    kinds = [u.split("/")[-1].split(":")[-1] if m == "POST" else "update" for m, u, _ in env.fake.writes()]
    assert kinds[0] == "update" and kinds[-1] == "batchClear"


def test_last_paid_date_is_integer_serial_and_matches_screen_text(env, db, settings, monkeypatch):
    """GAS の setValues('yyyy-MM-dd') をシート（Etc/GMT）が日付のセルにした結果と同じ整数のシリアル値で書く。
    画面へ返す文字（シートの時間帯の 0 時の String(Date)＝JST 9:00）と同じ瞬間になる。"""
    monkeypatch.setattr(summary_sheet, "summary_rows", lambda db, now: rows([("1001", 1, "2026-11-27", "")]))
    assert summary_sheet.write(db, api=env.api, app_settings=settings)["ok"]
    cell = env.tab.get(1, 10)
    assert cell.value == (date(2026, 11, 27) - codec.EPOCH.date()).days and cell.fmt["type"] == "DATE"
    instant = codec.serial_to_datetime(cell.value, "Etc/GMT")
    assert instant.astimezone(codec.JST).hour == 9
    assert last_paid_date_display("2026-11-27", "gas", "Etc/GMT") == "Fri Nov 27 2026 09:00:00 GMT+0900 (Japan Standard Time)"


def test_compare_matches_the_current_tab(env, db):
    expected = summary_sheet.summary_rows(db, NOW)
    assert summary_sheet.compare(expected, env.tab.values(), "Etc/GMT") == []
    expected[0]["paidCount"] = 2
    assert summary_sheet.compare(expected, env.tab.values(), "Etc/GMT")[0]["col"] == "有給回数"


def test_no_rows_or_no_tab_or_gate_closed_writes_nothing(env, db, settings, monkeypatch):
    db.execute(delete(LeaveRequest))
    db.commit()
    assert summary_sheet.write(db, api=env.api, app_settings=settings)["ok"] is False
    assert env.fake.writes() == []
    monkeypatch.setattr(write_gate, "check", lambda *a, **k: (False, "閉"))
    assert summary_sheet.write(db, api=env.api, app_settings=settings)["writeBlocked"] == "閉"


def test_missing_tab_is_not_created(engine, db, settings, monkeypatch):
    install_fake_summaries(monkeypatch)
    fake = make_fake()
    migrate(fake, engine, settings)
    del fake.books[STAGING_SID].tabs["V_LEAVE_SUMMARY"]
    result = summary_sheet.write(db, api=fake.api(settings), app_settings=settings)
    assert result["ok"] is False and "作りません" in result["error"] and fake.writes() == []


def test_mark_dirty_debounces_into_one_job(db, monkeypatch):
    freeze_now(monkeypatch, NOW)
    summary_sheet.mark_dirty(db)
    freeze_now(monkeypatch, NOW + timedelta(seconds=5))
    summary_sheet.mark_dirty(db)
    db.commit()
    pending = db.scalars(select(Job).where(Job.kind == summary_sheet.JOB_KIND)).all()
    assert len(pending) == 1 and pending[0].lane == "default"
    assert pending[0].run_after.replace(tzinfo=None) == (NOW + timedelta(seconds=15)).replace(tzinfo=None)
    for i in range(20):  # 予約が続いても最初の予約から 60 秒で動く
        freeze_now(monkeypatch, NOW + timedelta(seconds=10 + 5 * i))
        summary_sheet.mark_dirty(db)
    db.commit()
    db.refresh(pending[0])
    assert pending[0].run_after.replace(tzinfo=None) == (NOW + timedelta(seconds=60)).replace(tzinfo=None)


def test_register_worker(monkeypatch):
    from app import worker

    summary_sheet.register_worker()
    assert worker.schedules()["summary_sheet"].at == (6, 50)
    assert jobs.get_kind(summary_sheet.JOB_KIND).lane == "default"
