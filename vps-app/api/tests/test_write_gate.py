"""シートへ書く前の関所（services/write_gate）。

本番は「最後に合格した gas_ping が N 分以内・GAS が移転モード（migrationMode=true）」でなければ、行の反映・照合・
集計タブ・マスタ同期はシートへ書かない。ステージングは、ステージングの GAS への合格した ping が N 分以内なら書ける。
（row_sync などがこの関所で止まることの試験は、それぞれの担当 D が足す。）
"""
from __future__ import annotations

import types
from datetime import datetime, timedelta, timezone

import pytest

from app.models import SyncRun
from app.services import write_gate

pytestmark = pytest.mark.real_write_gate

NOW = datetime(2026, 11, 26, 3, 0, tzinfo=timezone.utc)
PROD = types.SimpleNamespace(is_staging=False, sheet_write_ping_max_minutes=10)
STAGING = types.SimpleNamespace(is_staging=True, sheet_write_ping_max_minutes=10)


def add_ping(db, *, at=NOW, ok=True, migrated=True):
    db.add(SyncRun(kind="gas_ping", started_at=at, finished_at=at, ok=ok,
                   summary_json={"problems": [] if ok else ["x"], "migrationMode": migrated}))
    db.commit()


def test_no_ping_blocks(db):
    ok, reason = write_gate.evaluate(db, PROD, NOW)
    assert not ok and "gas_ping" in reason


def test_production_requires_fresh_ok_ping_with_migrated_true(db):
    add_ping(db, at=NOW - timedelta(minutes=3), migrated=True)
    assert write_gate.evaluate(db, PROD, NOW) == (True, "")
    assert not write_gate.evaluate(db, PROD, NOW + timedelta(minutes=8))[0]  # 11 分前は古い


@pytest.mark.parametrize("migrated", [False, None, "true"])
def test_production_blocks_when_gas_not_in_migration_mode(db, migrated):
    add_ping(db, migrated=migrated)
    ok, reason = write_gate.evaluate(db, PROD, NOW)
    assert not ok and "移転モード" in reason


def test_failed_latest_ping_falls_back_to_last_ok_one(db):
    add_ping(db, at=NOW - timedelta(minutes=4), migrated=True)
    add_ping(db, at=NOW - timedelta(minutes=1), ok=False, migrated=None)
    assert write_gate.evaluate(db, PROD, NOW)[0] is True
    add_ping(db, at=NOW, ok=True, migrated=False)  # 戻した（移転モードを解いた）
    assert write_gate.evaluate(db, PROD, NOW)[0] is False


def test_staging_needs_fresh_staging_ping_but_not_migration_mode(db):
    assert not write_gate.evaluate(db, STAGING, NOW)[0]
    add_ping(db, migrated=False)
    assert write_gate.evaluate(db, STAGING, NOW) == (True, "")


def test_status_shape(db):
    add_ping(db, migrated=True)
    st = write_gate.status(db, PROD, NOW)
    assert st["allowed"] is True and st["blocked"] is False and st["migrated"] is True and st["lastOkPingAt"]


def test_check_is_evaluate(db):
    assert write_gate.check(db, PROD, NOW) == write_gate.evaluate(db, PROD, NOW)
