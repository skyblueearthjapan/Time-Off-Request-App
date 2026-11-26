"""二重承認・二重押し・同時の申請（実装仕様 12-2 W・K-4）。D・M の実物（gas_jobs・summary_sheet）を呼ぶ。

SQLite では順番に呼んで「後の方が止まる」ことを、PostgreSQL（TEST_DATABASE_URL）では本当に同時に走らせる。
"""
from __future__ import annotations

import os
import threading
from datetime import date

import pytest
from test_domain_helpers import SOMU, USER, add_calendar, add_request, ctx, standard_masters

from app import rpc
from app.domain import approval, register, requests
from app.models import Job, LeaveRequest
from app.rpc import RpcError

postgres_only = pytest.mark.skipif(not os.environ.get("TEST_DATABASE_URL"), reason="PostgreSQL の結合試験だけ")


def gas_jobs(db, kind=None):
    q = db.query(Job).filter(Job.lane == "gas")
    if kind:
        q = q.filter(Job.kind == kind)
    return q.all()


def test_two_screens_approve1_one_job(db):
    add_request(db, "LV-1", status="申請中")
    approval.api_approveLeaveRequest(ctx(db), "LV-1", "", "a@x")
    db.commit()
    with pytest.raises(RpcError, match="^この申請はすでに承認済みです。$"):
        approval.api_approveLeaveRequest(ctx(db), "LV-1", "", "b@x")
    db.rollback()
    assert len(gas_jobs(db, "internal.afterApprove1")) == 1
    assert db.query(Job).filter(Job.kind == "summary_sheet.write").count() == 1
    assert db.get(LeaveRequest, "LV-1").approved_by1_email == "a@x"


def test_two_screens_approve2_one_job(db):
    standard_masters(db)
    add_request(db, "LV-1")
    approval.api_approveLeaveRequest2(ctx(db, SOMU), "LV-1", SOMU)
    db.commit()
    with pytest.raises(RpcError, match="^この申請はすでに2次承認済みです。$"):
        approval.api_approveLeaveRequest2(ctx(db, SOMU), "LV-1", SOMU)
    db.rollback()
    jobs = gas_jobs(db, "internal.afterApprove2")
    assert len(jobs) == 1


def test_edit_and_cancel_real_jobs_args(db):
    import json

    add_request(db, "LV-1", status="承認済", leave_date=None, leave_text="2026/12/03")
    requests.api_editLeaveRequest(ctx(db), "LV-1", {})
    add_request(db, "LV-2", status="申請中", leave_date=date(2026, 12, 4))
    requests.api_deleteLeaveRequest(ctx(db), "LV-2")
    db.commit()
    args = [json.loads(j.args_json) for j in gas_jobs(db, "internal.afterUnapprove")]
    assert [(a["reason"], a["oldLeaveDate"], a["clearSign"]) for a in args] == [
        ("edit", "2026-12-03", True), ("cancel", "2026-12-04", False)]


def test_double_submit_same_idempotency_key_one_row(client, db, login_as):
    register.register_all(rpc.register)
    add_calendar(db, [(date(2026, 1, 1), "祝日")])
    login_as(USER)
    body = [{"leaveDate": "2026-12-01", "leaveType": "有給消化（私用）", "workerId": "101"}]
    headers = {"X-Idempotency-Key": "key-0123456789"}
    r1 = client.post("/api/rpc/api_submitLeaveRequest", json=body, headers=headers)
    r2 = client.post("/api/rpc/api_submitLeaveRequest", json=body, headers=headers)
    assert r1.status_code == 200 and r1.json()["result"]["ok"] is True
    assert r2.json() == r1.json()
    db.expire_all()
    assert db.query(LeaveRequest).count() == 1


def test_rpc_submit_error_is_200_and_not_found_400(client, db, login_as):
    register.register_all(rpc.register)
    login_as(USER)
    r = client.post("/api/rpc/api_submitLeaveRequest", json=[{"leaveDate": "x"}])
    assert r.status_code == 200 and r.json()["result"] == {"ok": False, "error": "休暇日の形式が不正です（yyyy-MM-dd）。"}
    r = client.post("/api/rpc/api_deleteLeaveRequest", json=["LV-NOPE"])
    assert r.status_code == 400 and r.json() == {"ok": False, "error": "申請が見つかりません。"}
    r = client.post("/api/rpc/api_getLeaveRequestById", json=["LV-NOPE"])
    assert r.status_code == 200 and r.json() == {"ok": True, "result": None}


def _run_parallel(session_factory, fns):
    barrier = threading.Barrier(len(fns))
    results: list = [None] * len(fns)

    def run(i, fn):
        with session_factory() as s:
            try:
                barrier.wait()
                results[i] = fn(s)
                s.commit()
            except BaseException as exc:  # noqa: BLE001
                s.rollback()
                results[i] = exc

    threads = [threading.Thread(target=run, args=(i, fn)) for i, fn in enumerate(fns)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    return results


@pytest.mark.postgres
@postgres_only
def test_concurrent_approve1_postgres(session_factory):
    with session_factory() as s:
        add_request(s, "LV-1", status="申請中")
    results = _run_parallel(session_factory, [
        lambda s: approval.api_approveLeaveRequest(ctx(s), "LV-1", "", "a@x"),
        lambda s: approval.api_approveLeaveRequest(ctx(s), "LV-1", "", "b@x"),
    ])
    errors = [r for r in results if isinstance(r, RpcError)]
    assert len(errors) == 1 and errors[0].message == "この申請はすでに承認済みです。"
    with session_factory() as s:
        assert len(gas_jobs(s, "internal.afterApprove1")) == 1


@pytest.mark.postgres
@postgres_only
def test_concurrent_submit_unique_req_ids_postgres(session_factory):
    with session_factory() as s:
        add_calendar(s, [(date(2026, 1, 1), "祝日")])
    results = _run_parallel(session_factory, [
        lambda s: requests.api_submitLeaveRequest(ctx(s), {"leaveDate": "2026-12-01"}) for _ in range(4)])
    ids = [r["reqId"] for r in results]
    assert sorted(ids) == [f"LV-20261126-00{i}" for i in range(1, 5)]


@pytest.mark.postgres
@postgres_only
def test_concurrent_same_substitute_only_one_postgres(session_factory):
    with session_factory() as s:
        add_calendar(s, [(date(2026, 1, 1), "祝日")])
    data = {"leaveDate": "2026-12-01", "leaveType": "振替休暇", "substituteDate": "2026-11-28", "workerId": "101"}
    results = _run_parallel(session_factory, [
        lambda s: requests.api_submitLeaveRequest(ctx(s), dict(data)) for _ in range(2)])
    assert sorted(r["ok"] for r in results) == [False, True]
