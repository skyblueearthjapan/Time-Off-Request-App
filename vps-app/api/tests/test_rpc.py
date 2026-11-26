"""RPC の入口と呼び出し: 閉じたリスト（34 関数）・Origin・415・413・404・400・501・401・403・冪等キーの再生と 409・監査。"""
from __future__ import annotations

import json
import types

import pytest
from sqlalchemy import func, select

from app import rpc
from app.models import AuditLog, DeptApprover, IdempotencyKey, Setting
from conftest import make_admin

ME = "tanaka@lineworks-local.info"
ORIGIN = "https://leave.lineworks-local.net"


@pytest.fixture
def logged_in(client, login_as):
    login_as(ME)
    return client


def count(db, model):
    return db.execute(select(func.count()).select_from(model)).scalar()


# ===== 閉じたリスト =====


def test_closed_list_starts_empty():
    assert rpc.registered_functions() == []
    assert rpc.implemented_functions() == []


def test_closed_list_is_the_34_functions_of_the_spec():
    """仕様書 4-3 の 27 関数＋4-2 の ops_* 7 関数。保守の関数・画面が呼ばない関数は入れない。"""
    assert len(rpc.SCREEN_FUNCTIONS) == 27 and len(set(rpc.SCREEN_FUNCTIONS)) == 27
    assert len(rpc.OPS_FUNCTIONS) == 7
    assert len(rpc.RPC_FUNCTIONS) == 34
    assert all(n.startswith("api_") for n in rpc.SCREEN_FUNCTIONS)
    assert all(n.startswith("ops_") for n in rpc.OPS_FUNCTIONS)
    for name in ("api_canAccessAdmin", "api_getDeptApproverDepts", "api_approveLeaveRequest2Batch", "setupAllSheets",
                 "setupAllTriggers", "syncAllMasters", "rebuildLeaveSummary", "checkAndSendWarnMails",
                 "debug_checkLeaveRequestSheet", "test_generatePdf", "migrateLeaveRequestHeaders_"):
        assert name not in rpc.RPC_FUNCTIONS, name


@pytest.mark.parametrize("name", ["api_canAccessAdmin", "api_approveLeaveRequest2Batch", "syncAllMasters", "api_x"])
def test_register_rejects_names_outside_closed_list(name):
    with pytest.raises(ValueError):
        rpc.register(name, lambda ctx: None)


def test_register_builds_closed_list():
    def handler(ctx):
        return ["A"]

    assert rpc.register("api_getDeptList", handler) is handler
    rpc.register("api_submitLeaveRequest", write=True)
    assert rpc.registered_functions() == ["api_getDeptList", "api_submitLeaveRequest"]
    assert rpc.get_handler("api_getDeptList") is handler
    assert rpc.implemented_functions() == ["api_getDeptList"]
    assert rpc.get_spec("api_submitLeaveRequest").write is True
    assert rpc.get_spec("api_getDeptList").write is False


@pytest.mark.parametrize("kwargs", [{"role": "everyone"}, {"request_id_arg": "second"}])
def test_register_rejects_unknown_options(kwargs):
    with pytest.raises(ValueError):
        rpc.register("api_x", lambda ctx: None, **kwargs)


@pytest.mark.parametrize("name", ["", "__init__", "a.b", "1abc", "api-x", "doGet;x"])
def test_register_rejects_bad_names(name):
    with pytest.raises(ValueError):
        rpc.register(name, lambda ctx: None)


def test_register_twice_is_error():
    rpc.register("api_getDeptList", lambda ctx: None)
    with pytest.raises(ValueError):
        rpc.register("api_getDeptList", lambda ctx: None)


def test_load_domain_handlers_registers_whole_closed_list():
    # W の app/domain/register.py が入った後は、閉じたリストの全部に中身が登録される
    loaded = rpc.load_domain_handlers()
    assert sorted(loaded) == sorted(rpc.RPC_FUNCTIONS)


def test_load_domain_handlers_calls_register_all(monkeypatch):
    import sys

    module = types.ModuleType("app.domain.register")

    def register_all(register):
        register("api_getDeptList", lambda ctx: ["A"])
        register("api_isAdmin")

    module.register_all = register_all
    monkeypatch.setitem(sys.modules, "app.domain.register", module)
    import app.domain

    monkeypatch.setattr(app.domain, "register", module, raising=False)
    assert rpc.load_domain_handlers() == ["api_getDeptList"]
    assert rpc.registered_functions() == ["api_getDeptList", "api_isAdmin"]


# ===== 入口 =====


def test_unknown_function_is_404(logged_in):
    assert logged_in.post("/api/rpc/api_getFormUrl", json=[]).status_code == 404
    assert logged_in.post("/api/rpc/__init__", json=[]).status_code == 404
    assert logged_in.post("/api/rpc/api_approveLeaveRequest2Batch", json=[[]]).status_code == 404


def test_declared_but_unimplemented_function_is_501(logged_in):
    rpc.register("api_getUpcomingLeaves")
    resp = logged_in.post("/api/rpc/api_getUpcomingLeaves", json=[])
    assert resp.status_code == 501
    assert resp.json()["ok"] is False


@pytest.mark.parametrize("name", sorted(rpc.RPC_FUNCTIONS))
def test_every_listed_function_is_501_until_registered(client, login_as, name):
    login_as("owner@lineworks-local.info")  # 運用者（ops_* の 403 を通る）
    assert client.post(f"/api/rpc/{name}", json=[]).status_code == 501


def test_login_required_is_401_json(client):
    rpc.register("api_getDeptList", lambda ctx: [])
    resp = client.post("/api/rpc/api_getDeptList", json=[])
    assert resp.status_code == 401
    assert resp.json() == {"ok": False, "error": "ログインが必要です。", "login": "/auth/login"}


def test_success_shape_and_context(logged_in):
    seen = {}

    def upcoming(ctx):
        seen["user"], seen["fn"], seen["now"] = ctx.user, ctx.fn, ctx.now
        return {"items": []}

    rpc.register("api_getUpcomingLeaves", upcoming)
    resp = logged_in.post("/api/rpc/api_getUpcomingLeaves", json=[])
    assert resp.status_code == 200
    assert resp.json() == {"ok": True, "result": {"items": []}}
    assert resp.headers["cache-control"] == "no-store"
    assert seen["user"].email == ME and seen["fn"] == "api_getUpcomingLeaves"
    assert seen["now"].tzinfo is not None


@pytest.mark.parametrize("result", [None, {"error": "x"}, {"ok": False, "error": "休暇日を入力してください"},
                                    {"ok": True, "results": [{"reqId": "a", "ok": False, "error": "e"}]}])
def test_failure_shapes_inside_success_are_passed_through(logged_in, result):
    rpc.register("api_submitLeaveRequest", lambda ctx: result, write=True)
    resp = logged_in.post("/api/rpc/api_submitLeaveRequest", json=[])
    assert resp.status_code == 200 and resp.json() == {"ok": True, "result": result}


def test_business_error_is_400_and_rolls_back(logged_in, db):
    def delete_request(ctx, req_id):
        ctx.db.add(Setting(key="X", value="x"))
        ctx.db.flush()
        raise rpc.RpcError("この申請はすでに取消済みです。")

    rpc.register("api_deleteLeaveRequest", delete_request, write=True, request_id_arg="first")
    resp = logged_in.post("/api/rpc/api_deleteLeaveRequest", json=["LV-20261126-001"])
    assert resp.status_code == 400
    assert resp.json() == {"ok": False, "error": "この申請はすでに取消済みです。"}
    db.expire_all()
    assert db.get(Setting, "X") is None
    # 失敗も監査に残る（業務の書き込みは戻した後で）
    row = db.scalars(select(AuditLog)).one()
    assert row.fn == "api_deleteLeaveRequest" and row.ok is False
    assert json.loads(row.request_ids) == ["LV-20261126-001"]
    assert row.error == "この申請はすでに取消済みです。"


def test_unexpected_error_is_500_without_details(logged_in):
    def boom(ctx):
        raise KeyError("secret detail")

    rpc.register("api_getDeptList", boom)
    resp = logged_in.post("/api/rpc/api_getDeptList", json=[])
    assert resp.status_code == 500
    assert "secret" not in resp.text
    assert resp.json()["error"] == "サーバーでエラーが起きました。時間をおいてもう一度お試しください。"


def test_optional_args_and_wrong_arity(logged_in):
    rpc.register("api_getWorkingDays", lambda ctx, year, month=None: [year, month])
    url = "/api/rpc/api_getWorkingDays"
    assert logged_in.post(url, json=[2026]).json()["result"] == [2026, None]
    assert logged_in.post(url, json=[2026, 11]).json()["result"] == [2026, 11]
    # 足りない引数は GAS（JS）の undefined と同じく None で受ける（仕様書 4-1）
    assert logged_in.post(url, json=[]).json()["result"] == [None, None]
    # 余った引数は GAS（JS）と同じく捨てる
    assert logged_in.post(url, json=[2026, None, "extra"]).json()["result"] == [2026, None]


@pytest.mark.parametrize("body", [b"{}", b"\"x\"", b"not json", b"123", b"null"])
def test_body_must_be_json_array(logged_in, body):
    rpc.register("api_getWorkersByDept", lambda ctx, *a: None)
    resp = logged_in.post("/api/rpc/api_getWorkersByDept", content=body, headers={"content-type": "application/json"})
    assert resp.status_code == 400


def test_empty_body_means_no_args(logged_in):
    rpc.register("api_getDeptList", lambda ctx: "ok")
    resp = logged_in.post("/api/rpc/api_getDeptList", content=b"", headers={"content-type": "application/json"})
    assert resp.json() == {"ok": True, "result": "ok"}


def test_large_sign_png_payload_is_accepted(logged_in):
    """手書きサインの PNG（base64）を 1 次承認の引数で送れる（api は 3 MB、nginx は 4m まで）。"""
    data = "data:image/png;base64," + "A" * 1_000_000
    rpc.register("api_approveLeaveRequest", lambda ctx, req_id, sign: len(sign), write=True, request_id_arg="first")
    resp = logged_in.post("/api/rpc/api_approveLeaveRequest", json=["LV-1", data])
    assert resp.status_code == 200 and resp.json()["result"] == len(data)


def test_body_too_large_is_413(logged_in):
    """本文の上限: api_approveLeaveRequest だけ 3 MB、ほかは 256 KB（仕様書 2-2）。"""
    rpc.register("api_approveLeaveRequest", lambda ctx, req_id, sign: "ok", write=True)
    rpc.register("api_submitLeaveRequest", lambda ctx, data: "ok", write=True)
    big = "A" * (3 * 1024 * 1024)
    assert logged_in.post("/api/rpc/api_approveLeaveRequest", json=["LV-1", big]).status_code == 413
    assert logged_in.post("/api/rpc/api_submitLeaveRequest", json=["A" * (256 * 1024)]).status_code == 413
    assert logged_in.post("/api/rpc/api_submitLeaveRequest", json=["A" * (200 * 1024)]).status_code == 200


def test_content_type_must_be_json_415(logged_in):
    rpc.register("api_getDeptList", lambda ctx: "ok")
    assert logged_in.post("/api/rpc/api_getDeptList").status_code == 415
    for ct in ("text/plain", "application/x-www-form-urlencoded", "multipart/form-data; boundary=x"):
        assert logged_in.post("/api/rpc/api_getDeptList", content=b"[]", headers={"content-type": ct}).status_code == 415
    ok = logged_in.post("/api/rpc/api_getDeptList", content=b"[]",
                        headers={"content-type": "application/json; charset=utf-8"})
    assert ok.status_code == 200


def test_missing_origin_requires_same_origin_fetch_403(logged_in):
    rpc.register("api_getDeptList", lambda ctx: "ok")
    for site in ("cross-site", "same-site", "none"):
        assert logged_in.post("/api/rpc/api_getDeptList", json=[], headers={"sec-fetch-site": site}).status_code == 403
    del logged_in.headers["sec-fetch-site"]
    assert logged_in.post("/api/rpc/api_getDeptList", json=[]).status_code == 403
    assert logged_in.post("/api/rpc/api_getDeptList", json=[], headers={"origin": ORIGIN}).status_code == 200


@pytest.mark.parametrize("depth", [33, 2000, 20_000])  # 2000 段から json の再帰の上限を超える（本文は 256 KB 以内）
def test_deeply_nested_json_is_400_not_500(logged_in, depth):
    """[[[[…]]]] の深い入れ子は、json の再帰の上限（RecursionError）でも 500 にせず形の誤り（400）。"""
    rpc.register("api_getDeptList", lambda ctx, *a: "ok")
    body = ("[" * depth + "]" * depth).encode()
    resp = logged_in.post("/api/rpc/api_getDeptList", content=body, headers={"content-type": "application/json"})
    assert resp.status_code == 400 and resp.json() == {"ok": False, "error": "呼び出しの形が正しくありません。"}
    deep_obj = ('[' + '{"a":' * depth + '1' + '}' * depth + ']').encode()
    resp = logged_in.post("/api/rpc/api_getDeptList", content=deep_obj, headers={"content-type": "application/json"})
    assert resp.status_code == 400


def test_normal_nesting_is_accepted(logged_in):
    rpc.register("api_getDeptList", lambda ctx, *a: "ok")
    resp = logged_in.post("/api/rpc/api_getDeptList", json=[{"a": [{"b": [1]}]}, [[["x"]]]])
    assert resp.status_code == 200


def test_cross_origin_is_403(logged_in):
    rpc.register("api_getDeptList", lambda ctx: "ok")
    for origin in ("https://evil.example.com", "http://leave.lineworks-local.net", "null",
                   "https://leave.lineworks-local.net.evil.example.com"):
        resp = logged_in.post("/api/rpc/api_getDeptList", json=[], headers={"origin": origin})
        assert resp.status_code == 403
        assert resp.json() == {"ok": False, "error": "この画面からは操作できません。"}
    assert logged_in.post("/api/rpc/api_getDeptList", json=[], headers={"origin": ORIGIN + "/"}).status_code == 200


# ===== 権限（画面の関数は 400＝GAS の throw。ops_* は運用者だけで 403） =====


def _register_probe(name, role, **kw):
    called = []
    rpc.register(name, lambda ctx, *a: called.append(1) or {"ok": True}, role=role, **kw)
    return called


def test_admin_role(logged_in, db):
    called = _register_probe("api_getAdminBossView", "admin")
    resp = logged_in.post("/api/rpc/api_getAdminBossView", json=[])
    assert resp.status_code == 400 and resp.json() == {"ok": False, "error": "権限がありません"}
    assert called == []
    make_admin(db, ME)
    assert logged_in.post("/api/rpc/api_getAdminBossView", json=[]).status_code == 200
    assert called == [1]


def test_somu_or_admin_role(logged_in, db):
    called = _register_probe("api_getPending2ndApprovals", "somu_or_admin")
    resp = logged_in.post("/api/rpc/api_getPending2ndApprovals", json=[])
    assert resp.status_code == 400
    assert resp.json()["error"] == "権限がありません（総務または管理者のみ）"
    db.add(Setting(key="SOMU_EMAILS", value=ME.upper(), sheet_row=3))
    db.commit()
    assert logged_in.post("/api/rpc/api_getPending2ndApprovals", json=[]).status_code == 200
    assert called == [1]


def test_dept_approver_or_admin_role(logged_in, db):
    _register_probe("api_getDeptAdminView", "dept_approver_or_admin")
    resp = logged_in.post("/api/rpc/api_getDeptAdminView", json=[])
    assert resp.status_code == 400 and resp.json()["error"] == "部署承認者の権限がありません。"
    db.add(DeptApprover(sort_order=1, col_a="製造1課", col_b=f"x@lineworks-local.info, {ME}", col_c="田中"))
    db.commit()
    assert logged_in.post("/api/rpc/api_getDeptAdminView", json=[]).status_code == 200


def test_admin_page_role_and_custom_message(logged_in, db):
    _register_probe("api_getAdminSomuView", "admin_page", forbidden_message="権限がありません。")
    resp = logged_in.post("/api/rpc/api_getAdminSomuView", json=[])
    assert resp.status_code == 400 and resp.json()["error"] == "権限がありません。"
    make_admin(db, ME)  # 管理者でも ALLOW_ADMIN_PAGE_EMAILS に無ければだめ（GAS の canAccessAdmin_ と同じ）
    assert logged_in.post("/api/rpc/api_getAdminSomuView", json=[]).status_code == 400
    db.add(Setting(key="ALLOW_ADMIN_PAGE_EMAILS", value=ME, sheet_row=4))
    db.commit()
    assert logged_in.post("/api/rpc/api_getAdminSomuView", json=[]).status_code == 200


def test_bootstrap_admin_is_not_screen_admin(client, login_as):
    """BOOTSTRAP_ADMINS は画面の管理者（ADMIN_EMAILS）に足さない（仕様書 8-9）。運用者ではある。"""
    rpc.register("api_getAdminBossView", lambda ctx: {"ok": True}, role="admin")
    rpc.register("ops_getStatus", lambda ctx: {"ok": True})
    login_as("owner@lineworks-local.info")
    assert client.post("/api/rpc/api_getAdminBossView", json=[]).status_code == 400
    assert client.post("/api/rpc/ops_getStatus", json=[]).status_code == 200


def test_ops_functions_only_for_operators(logged_in, db):
    """K-9: ops_* は総務・管理者・BOOTSTRAP_ADMINS だけ。ほかは 403「管理者権限がありません。」。"""
    called = []
    rpc.register("ops_getStatus", lambda ctx: called.append(1) or {"ok": True})
    resp = logged_in.post("/api/rpc/ops_getStatus", json=[])
    assert resp.status_code == 403 and resp.json() == {"ok": False, "error": "管理者権限がありません。"}
    assert called == []
    db.add(Setting(key="ALLOW_ADMIN_PAGE_EMAILS", value=ME, sheet_row=5))  # 古い管理画面だけでは足りない
    db.commit()
    assert logged_in.post("/api/rpc/ops_getStatus", json=[]).status_code == 403
    db.add(Setting(key="SOMU_EMAILS", value=ME, sheet_row=6))
    db.commit()
    assert logged_in.post("/api/rpc/ops_getStatus", json=[]).status_code == 200
    assert called == [1]


def test_ops_function_for_admin(logged_in, db):
    rpc.register("ops_retryJob", lambda ctx, job_id: {"ok": True}, write=True)
    assert logged_in.post("/api/rpc/ops_retryJob", json=[1]).status_code == 403
    make_admin(db, ME)
    assert logged_in.post("/api/rpc/ops_retryJob", json=[1]).status_code == 200


def test_handler_side_check_keeps_gas_order(logged_in, db):
    """確かめの順を GAS と同じにしたい関数（21: 先に reqId、次に権限）は role を付けず業務関数の中で確かめる。"""

    def approve_second(ctx, req_id, approver2_email):
        if not req_id:
            raise rpc.RpcError("REQ_IDが指定されていません。")
        if not ctx.has_role("somu_or_admin"):
            raise rpc.RpcError("権限がありません（総務または管理者のみ）")
        return {"ok": True, "reqId": req_id}

    rpc.register("api_approveLeaveRequest2", approve_second, write=True, request_id_arg="first")
    url = "/api/rpc/api_approveLeaveRequest2"
    resp = logged_in.post(url, json=["", ME])
    assert resp.status_code == 400 and resp.json()["error"] == "REQ_IDが指定されていません。"
    resp = logged_in.post(url, json=["LV-1", ME])
    assert resp.status_code == 400 and resp.json()["error"] == "権限がありません（総務または管理者のみ）"
    make_admin(db, ME)
    assert logged_in.post(url, json=["LV-1", ME]).status_code == 200
    with pytest.raises(rpc.RpcError):  # require_admin も 400（RpcError）
        rpc.RpcContext(db=db, user=types.SimpleNamespace(email="x@lineworks-local.info"),
                       settings=None).require_admin()


# ===== 冪等キー =====


def test_idempotency_key_replays_first_result(logged_in, db):
    calls = []

    def submit(ctx, data):
        calls.append(data)
        ctx.db.add(Setting(key=f"k{len(calls)}", value="v"))
        return {"ok": True, "reqId": f"LV-20261126-00{len(calls)}"}

    rpc.register("api_submitLeaveRequest", submit, write=True)
    h = {"X-Idempotency-Key": "abcdefgh-1234"}
    first = logged_in.post("/api/rpc/api_submitLeaveRequest", json=[{"a": 1}], headers=h)
    second = logged_in.post("/api/rpc/api_submitLeaveRequest", json=[{"a": 1}], headers=h)
    assert first.status_code == second.status_code == 200
    assert first.json() == second.json() == {"ok": True, "result": {"ok": True, "reqId": "LV-20261126-001"}}
    assert len(calls) == 1
    assert count(db, Setting) == 1  # 二重押しで 2 件できない
    row = db.get(IdempotencyKey, (ME, "abcdefgh-1234"))
    assert row.state == "done" and row.fn == "api_submitLeaveRequest"
    third = logged_in.post("/api/rpc/api_submitLeaveRequest", json=[{"a": 1}],
                           headers={"X-Idempotency-Key": "abcdefgh-5678"})
    assert third.json()["result"]["reqId"] == "LV-20261126-002"
    # 同じキーを別の関数に使うと 400
    rpc.register("api_deleteLeaveRequest", lambda ctx, rid: {"ok": True}, write=True)
    assert logged_in.post("/api/rpc/api_deleteLeaveRequest", json=["R"], headers=h).status_code == 400
    # 監査: 成功は 2 回分（再生は数えない）。申請 ID は返事の reqId
    rows = db.scalars(select(AuditLog).where(AuditLog.ok.is_(True)).order_by(AuditLog.id)).all()
    assert [json.loads(r.request_ids) for r in rows] == [["LV-20261126-001"], ["LV-20261126-002"]]


def test_in_progress_key_is_409(logged_in, db):
    from app.timeutil import now_utc

    rpc.register("api_approveLeaveRequest", lambda ctx, rid: {"ok": True}, write=True)
    db.add(IdempotencyKey(user_email=ME, key="busy-key-0001", fn="api_approveLeaveRequest", state="running",
                          created_at=now_utc()))
    db.commit()
    resp = logged_in.post("/api/rpc/api_approveLeaveRequest", json=["R"],
                          headers={"X-Idempotency-Key": "busy-key-0001"})
    assert resp.status_code == 409
    assert resp.json()["ok"] is False


def test_read_only_functions_ignore_key_and_are_not_audited(logged_in, db):
    rpc.register("api_getUpcomingLeaves", lambda ctx: {"n": 1})
    logged_in.post("/api/rpc/api_getUpcomingLeaves", json=[], headers={"X-Idempotency-Key": "read-key-0001"})
    assert count(db, IdempotencyKey) == 0
    assert count(db, AuditLog) == 0


def test_business_error_does_not_keep_key(logged_in, db):
    state = {"n": 0}

    def approve(ctx, rid):
        state["n"] += 1
        if state["n"] == 1:
            raise rpc.RpcError("申請が見つかりません")
        return {"ok": True}

    rpc.register("api_approveLeaveRequest", approve, write=True, request_id_arg="first")
    h = {"X-Idempotency-Key": "retry-key-0001"}
    assert logged_in.post("/api/rpc/api_approveLeaveRequest", json=["R"], headers=h).status_code == 400
    assert db.get(IdempotencyKey, (ME, "retry-key-0001")) is None  # 鍵が残らない（もう一度押せる）
    assert logged_in.post("/api/rpc/api_approveLeaveRequest", json=["R"], headers=h).json() == {
        "ok": True, "result": {"ok": True}}
    assert count(db, IdempotencyKey) == 1


def test_crash_does_not_keep_key(logged_in, db):
    def crash(ctx, rid):
        raise RuntimeError("db down")

    rpc.register("api_editLeaveRequest", crash, write=True)
    resp = logged_in.post("/api/rpc/api_editLeaveRequest", json=["R"], headers={"X-Idempotency-Key": "crash-key-0001"})
    assert resp.status_code == 500
    assert count(db, IdempotencyKey) == 0


def test_bad_idempotency_key_is_400(logged_in):
    rpc.register("api_submitLeaveRequest", lambda ctx, p: "x", write=True)
    for key in ("short", "has space 123", "x" * 129, "semi;colon-key1"):
        resp = logged_in.post("/api/rpc/api_submitLeaveRequest", json=[{}], headers={"X-Idempotency-Key": key})
        assert resp.status_code == 400


def test_idempotency_is_per_user(client, login_as):
    calls = []
    rpc.register("api_submitLeaveRequest", lambda ctx, p: calls.append(ctx.user.email) or len(calls), write=True)
    h = {"X-Idempotency-Key": "shared-key-0001"}
    login_as("a@lineworks-local.info")
    assert client.post("/api/rpc/api_submitLeaveRequest", json=[{}], headers=h).json()["result"] == 1
    client.cookies.clear()
    login_as("b@lineworks-local.info")
    assert client.post("/api/rpc/api_submitLeaveRequest", json=[{}], headers=h).json()["result"] == 2


def test_purge_idempotency_keys(db):
    from datetime import timedelta

    from app.timeutil import now_utc

    now = now_utc()
    db.add_all([IdempotencyKey(user_email=ME, key="old-key-0001", fn="f", state="done",
                               created_at=now - timedelta(hours=25)),
                IdempotencyKey(user_email=ME, key="new-key-0001", fn="f", state="done",
                               created_at=now - timedelta(hours=23))])
    db.commit()
    assert rpc.purge_idempotency_keys(db, now) == 1
    db.commit()
    assert [k.key for k in db.scalars(select(IdempotencyKey))] == ["new-key-0001"]


# ===== 監査 =====


def test_batch_request_ids_are_audited(logged_in, db):
    rpc.register("api_approveLeaveRequest2", lambda ctx, ids, email=None: {"ok": True}, write=True,
                 request_id_arg="first_list")
    logged_in.post("/api/rpc/api_approveLeaveRequest2", json=[["A", "B"]])
    row = db.scalars(select(AuditLog)).one()
    assert row.ok and row.user_email == ME and json.loads(row.request_ids) == ["A", "B"]


def test_audit_is_in_same_transaction_as_business_write(logged_in, db):
    def edit(ctx, rid, patch):
        ctx.db.add(Setting(key="E", value="1"))
        return {"ok": True}

    rpc.register("api_editLeaveRequest", edit, write=True, request_id_arg="first")
    assert logged_in.post("/api/rpc/api_editLeaveRequest", json=["R9", {"leaveDate": "2026-12-01"}]).status_code == 200
    db.expire_all()
    assert db.get(Setting, "E") is not None
    assert json.loads(db.scalars(select(AuditLog)).one().request_ids) == ["R9"]


# ===== healthz =====


def test_healthz(client):
    resp = client.get("/healthz")
    assert resp.status_code == 200 and resp.json() == {"ok": True, "env": "staging"}


def test_healthz_detail_only_for_admin(client, db, login_as):
    assert client.get("/healthz?detail=1").json() == {"ok": True, "env": "staging"}
    login_as("someone@lineworks-local.info")
    assert "detail" not in client.get("/healthz?detail=1").json()
    make_admin(db, "someone@lineworks-local.info")
    body = client.get("/healthz?detail=1").json()
    assert body["detail"]["rowSync"] == {"pending": 0, "blocked": 0, "oldestSeconds": 0, "delayed": False}
    assert body["detail"]["gasPing"] is None
    from app import jsutil

    assert "icu" in body["detail"] and body["detail"]["icu"] == jsutil.icu_version()
