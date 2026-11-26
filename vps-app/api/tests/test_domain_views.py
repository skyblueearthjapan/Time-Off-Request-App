"""取得状況の画面（somuAdmin・deptAdmin・admin・leaveStatus。実装仕様 4-4-f・h・i）。シートは書かない。"""
from __future__ import annotations

from datetime import date

import pytest
from test_domain_helpers import ADMIN, BOSS, SOMU, USER, add_request, add_workers, ctx, jst, standard_masters

from app.domain import summaries
from app.models import Job
from app.rpc import RpcError

NOW = jst(2026, 11, 26, 10)
LONG = "2026-11-27"  # 既定は iso（2026-11-26 オーナーの決定）。名前は元のまま


@pytest.fixture
def data(db):
    standard_masters(db)
    add_workers(db, [("101", "山田太郎", "製造1課"), ("102", "佐藤次郎", "製造1課"), ("0201", "鈴木", "製造2課")])
    add_request(db, "LV-1", worker_id="101", leave_date=date(2026, 11, 27))
    add_request(db, "LV-2", worker_id="101", leave_date=date(2026, 7, 1), kubun="半日休")
    add_request(db, "LV-3", worker_id="0201", worker_kind="string", dept="製造2課", worker_name="鈴木",
                leave_date=date(2026, 12, 1), status="申請中")
    add_request(db, "LV-4", worker_id="0201", worker_kind="string", dept="製造2課", worker_name="鈴木",
                leave_date=date(2026, 4, 1))
    return db


def c(db, email=SOMU):
    return ctx(db, email, now=NOW)


def test_somu_admin_view(data):
    with pytest.raises(RpcError, match=r"^権限がありません（総務または管理者のみ）$"):
        summaries.api_getSomuAdminView(c(data, USER), 2026, "ALL")
    res = summaries.api_getSomuAdminView(c(data), None, "ALL")
    # 有給の少ない順 → 今か過去の期間の不足の多い順（佐藤 3・鈴木 2）
    assert [w["workerId"] for w in res] == ["102", "0201", "101"]
    yamada = res[2]
    assert yamada["paidCount"] == 1.5 and yamada["lastPaidDate"] == LONG and yamada["warnLevel"] == "注意"
    assert yamada["quarters"][1]["periodCount"] == 1 and yamada["currentShortage"] == 1
    sato = res[0]
    assert sato["paidCount"] == 0 and sato["lastPaidDate"] == "" and sato["warnLevel"] == "警告"
    # 集計タブを読み直すと '0201' は数の 201 になり、作業員マスタの '0201' と当たらない（今のまま）
    assert res[1]["paidCount"] == 0 and res[1]["quarters"][0]["periodCount"] == 1 and res[1]["currentShortage"] == 2
    assert [w["workerId"] for w in summaries.api_getSomuAdminView(c(data, ADMIN), "2026", "製造2課")] == ["0201"]
    assert data.query(Job).count() == 0  # シートも仕事も書かない


def test_dept_admin_view(data):
    with pytest.raises(RpcError, match="^部署承認者の権限がありません。$"):
        summaries.api_getDeptAdminView(c(data, USER), 2026, [])
    res = summaries.api_getDeptAdminView(c(data, BOSS), 2026, None)
    assert res["deptNames"] == ["製造1課", "製造2課"] and len(res["workers"]) == 3
    res = summaries.api_getDeptAdminView(c(data, BOSS), 2026, ["製造2課", "品質"])
    assert res["deptNames"] == ["製造2課"]
    res = summaries.api_getDeptAdminView(c(data, ADMIN), 2026, ["品質"])
    assert res == {"workers": [], "deptNames": ["品質"]}
    assert summaries.api_getDeptAdminView(c(data, ADMIN), 2026, []) == {"workers": [], "deptNames": []}


def test_admin_boss_and_somu_view(data):
    with pytest.raises(RpcError, match="^権限がありません$"):
        summaries.api_getAdminBossView(c(data, SOMU), "", 2026)
    res = summaries.api_getAdminBossView(c(data, ADMIN), "", None)
    assert [(r["deptName"], r["workerId"]) for r in res] == [("製造1課", "101"), ("製造2課", "201")]
    assert res[0] == {"deptId": "製造1課", "deptName": "製造1課", "workerId": "101", "workerName": "山田太郎",
                      "paidCount": 1.5, "paidHalfCount": 1, "substituteCount": 0, "specialCount": 0,
                      "totalCount": 2, "lastPaidDate": LONG, "warnLevel": "注意"}
    assert summaries.api_getAdminSomuView(c(data, ADMIN), 2026, "製造2課")[0]["workerId"] == "201"
    assert summaries.api_getAdminBossView(c(data, ADMIN), "製造1課", 2025) == []


def test_leave_status_search(data):
    res = summaries.api_getLeaveStatusSearch(c(data, USER), None, "", "")
    assert res["mode"] == "list" and res["fyYear"] == 2026 and res["deptFilter"] == "" and len(res["workers"]) == 3
    res = summaries.api_getLeaveStatusSearch(c(data, USER), "2026", "製造1課", "101")
    assert res["mode"] == "personal" and res["worker"]["workerId"] == "101"
    assert [m["state"] for m in res["milestones"]][:2] == ["done", "note"]
    assert [u["reqId"] for u in res["upcoming"]] == ["LV-1"]
    res = summaries.api_getLeaveStatusSearch(c(data, USER), 2025, "製造1課", "101")
    assert res["upcoming"] == []  # 今の年度でなければ予定は出さない
    res = summaries.api_getLeaveStatusSearch(c(data, USER), 2026, "", "999")
    assert res == {"mode": "personal", "fyYear": 2026, "worker": None, "milestones": [], "upcoming": []}
