"""services/reconcile: JST の日付で比べる（.625 と古い整数の行）・日時は秒・V 列の差は直す・G 列は書かず報告だけ・
半端な行は触らない・DB が少なければ書かない・--check は書かない・V_LEAVE_SUMMARY も比べて書き直す。"""
from __future__ import annotations

import pytest
from sqlalchemy import delete

from app.models import LeaveRequest, SyncRun
from app.services import reconcile, row_sync
from conftest import freeze_now
from d_sheet_fixtures import HEADERS, NOW, R_HALF, R_OLD, R_SUB, STAGING_SID, install_fake_summaries, make_fake, migrate


@pytest.fixture(autouse=True)
def _lane():
    row_sync.reset_lane()
    yield
    row_sync.reset_lane()


@pytest.fixture
def env(engine, session_factory, settings, monkeypatch):
    freeze_now(monkeypatch, NOW)
    install_fake_summaries(monkeypatch)
    fake = make_fake()
    migrate(fake, engine, settings, app_env="production")
    fake.calls.clear()
    book = fake.books[STAGING_SID]

    def run(check_only=False):
        return reconcile.run(session_factory, check_only=check_only, api=fake.api(settings), app_settings=settings)

    return type("Env", (), {"fake": fake, "book": book, "run": staticmethod(run), "tab": book.tabs["T_LEAVE_REQUEST"]})


def col(h):
    return HEADERS.index(h)


def test_clean_after_migration_and_check_writes_nothing(env):
    result = env.run(check_only=True)
    assert result["ok"] and result["diffCount"] == 0 and result["summaryDiffCount"] == 0
    assert result["extraRows"] == [{"row": 4, "cells": {"H": "メモ"}}]  # 半端な行は報告だけ
    assert env.fake.writes() == []


def test_old_9am_row_and_625_row_are_same_jst_day(env):
    # 古い行の休暇日（整数＝JST 9 時）を .625（JST 0 時）にしても同じ日 → 差にしない
    cell = env.tab.cell(1, col("LEAVE_DATE"))
    cell.value = cell.value - 0.375
    assert env.run(check_only=True)["diffCount"] == 0
    cell.value = cell.value + 1  # 1 日ずれた
    result = env.run(check_only=True)
    assert result["diffCount"] == 1 and result["diffs"][0]["pk"] == R_OLD and result["ok"] is False


def test_fixes_v_columns_but_never_g_columns(env, session_factory):
    env.tab.cell(4, col("STATUS")).value = "承認済"  # R_SUB は DB では申請中
    env.tab.cell(2, col("SIGN_IMAGE_URL")).value = "https://drive.google.com/file/d/OTHER/view"  # G 列の差
    env.tab.cell(2, col("APPLIED_AT")).value += 0.4 / 86400  # 1 秒未満の差は同じ
    result = env.run()
    assert result["fixedRows"] == 1 and result["diffs"][0]["pk"] == R_SUB
    assert env.tab.get(4, col("STATUS")).value == "申請中"
    assert result["gDiffs"][0]["pk"] == R_HALF
    assert env.tab.get(2, col("SIGN_IMAGE_URL")).value == "https://drive.google.com/file/d/OTHER/view"
    assert env.tab.get(3, col("WORKER_NAME")).value == "メモ"
    for req in env.fake.batch_requests():
        (kind,) = req.keys()
        assert kind == "updateCells" and HEADERS[req[kind]["start"]["columnIndex"]] not in ("SIGN_IMAGE_URL", "PDF_URL")


def test_summary_tab_is_rewritten_when_different(env):
    env.book.tabs["V_LEAVE_SUMMARY"].cell(1, 5).value = 9
    result = env.run(check_only=True)
    assert result["summaryDiffCount"] == 1 and result["ok"] is False
    result = env.run()
    assert result["summaryRewritten"] and env.book.tabs["V_LEAVE_SUMMARY"].get(1, 5).value == 1


def test_stops_when_db_has_fewer_rows_than_sheet(env, db):
    db.execute(delete(LeaveRequest).where(LeaveRequest.req_id == R_SUB))
    db.commit()
    env.tab.cell(1, col("STATUS")).value = "取消"
    result = env.run()
    assert result["ok"] is False and "少ない" in result["stopped"] and env.fake.writes() == []


def test_stops_when_db_is_empty(env, db, session_factory):
    db.execute(delete(LeaveRequest))
    db.commit()
    result = env.run()
    assert result["ok"] is False and "0 行" in result["stopped"] and env.fake.calls == []
    assert db.query(SyncRun).filter_by(kind="sheet_reconcile").count() == 1


def test_register_worker():
    from app import worker

    reconcile.register_worker()
    assert worker.schedules()["sheet_reconcile"].at == (2, 10)
