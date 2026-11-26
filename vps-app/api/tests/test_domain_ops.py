"""運用者だけの関数 ops_*（実装仕様 4-2・K-9）と 34 関数の登録。"""
from __future__ import annotations

import pytest
from test_domain_helpers import ADMIN, OWNER, SOMU, USER, ctx, install_fakes, standard_masters

from app import jobs, rpc
from app.domain import ops, register
from app.models import Job, UserSession
from app.rpc import RpcError, RpcForbidden


def test_register_all_matches_closed_list():
    names = register.register_all(rpc.register)
    assert sorted(names) == sorted(rpc.RPC_FUNCTIONS)
    assert rpc.get_spec("api_approveLeaveRequest").write is True
    assert rpc.get_spec("api_approveLeaveRequest").request_id_arg == "first"
    assert rpc.get_spec("api_getDeptList").write is False
    assert all(rpc.get_spec(n).role is None for n in names)


@pytest.mark.parametrize("fn", ["ops_getStatus", "ops_syncSheetConfigNow", "ops_runMasterSync",
                                "ops_rebuildSummaryNow", "ops_sendWarnMail"])
def test_ops_forbidden_for_non_operator(db, monkeypatch, fn):
    standard_masters(db)
    install_fakes(monkeypatch)
    with pytest.raises(RpcForbidden):
        ops.HANDLERS[fn](ctx(db, USER))


def test_ops_for_operators(db, monkeypatch):
    standard_masters(db)
    rec = install_fakes(monkeypatch)
    for email in (SOMU, ADMIN, OWNER):  # K-9: 総務・管理者・BOOTSTRAP_ADMINS
        assert ops.ops_getStatus(ctx(db, email)) == {"ok": True}
    assert ops.ops_syncSheetConfigNow(ctx(db, SOMU)) == {"ok": True}
    assert ops.ops_runMasterSync(ctx(db, SOMU)) == {"ok": True}
    assert ops.ops_rebuildSummaryNow(ctx(db, SOMU)) == {"ok": True, "rows": 3}
    res = ops.ops_sendWarnMail(ctx(db, SOMU))
    assert res["ok"] is True and isinstance(res["jobId"], int)
    assert rec.names() == ["status"] * 3 + ["config_sync", "master_sync", "rebuild", "warn_mail"]


def test_ops_send_warn_mail_real_job(db):
    standard_masters(db)
    res = ops.ops_sendWarnMail(ctx(db, SOMU))
    job = db.get(Job, res["jobId"])
    assert job.kind == "internal.sendWarnMail" and job.lane == "gas"
    assert ops.ops_sendWarnMail(ctx(db, SOMU))["jobId"] == res["jobId"]  # 待ちがあれば同じ仕事


def test_ops_get_status_real(db):
    standard_masters(db)
    st = ops.ops_getStatus(ctx(db, SOMU))
    assert "gasJobs" in st and "rowSync" in st


def test_ops_retry_job(db):
    standard_masters(db)
    jid = jobs.enqueue(db, "internal.afterApprove1", {}, request_id="LV-1")
    job = db.get(Job, jid)
    job.state = "failed"
    db.commit()
    for bad in (None, True, "1.5", " 3 ", 1.0):
        with pytest.raises(RpcError, match="^仕事の番号が正しくありません。$"):
            ops.ops_retryJob(ctx(db, SOMU), bad)
    assert ops.ops_retryJob(ctx(db, SOMU), str(jid)) == {"ok": True}
    assert db.get(Job, jid).state == "pending"
    with pytest.raises(RpcError, match="^失敗した仕事が見つかりません。$"):
        ops.ops_retryJob(ctx(db, SOMU), jid)


def test_ops_revoke_sessions(db, settings):
    from datetime import timedelta

    from app import timeutil

    standard_masters(db)
    now = timeutil.now_utc()
    db.add(UserSession(token_hash="a" * 64, email=USER, name="", created_at=now, expires_at=now + timedelta(days=1)))
    db.commit()
    assert ops.ops_revokeSessions(ctx(db, ADMIN), USER) == {"ok": True, "revoked": 1}
    with pytest.raises(RpcError, match="^メールアドレスを指定してください。$"):
        ops.ops_revokeSessions(ctx(db, ADMIN), "")


def test_ops_over_http_is_403(client, db, login_as):
    register.register_all(rpc.register)
    standard_masters(db)
    login_as(USER)
    r = client.post("/api/rpc/ops_getStatus", json=[])
    assert r.status_code == 403 and r.json() == {"ok": False, "error": "管理者権限がありません。"}
    # 画面の関数の権限の拒否は 400（GAS と同じ文）
    r = client.post("/api/rpc/api_getPending2ndApprovals", json=[])
    assert r.status_code == 400 and r.json() == {"ok": False, "error": "権限がありません（総務または管理者のみ）"}
