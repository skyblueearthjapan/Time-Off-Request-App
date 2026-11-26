"""部署・作業員・LOOKUP・電子印・権限の関数（実装仕様 4-3 の 1〜8・17・19・23、4-4-a・g）。"""
from __future__ import annotations

import json

import pytest
from test_domain_helpers import (
    ADMIN,
    BOSS,
    SOMU,
    USER,
    add_lookup,
    add_stamps,
    add_workers,
    ctx,
    standard_masters,
)

from app.domain import masters
from app.models import Worker
from app.rpc import RpcError


def test_dept_list_and_workers(db):
    add_workers(db, [("B02", "佐藤", "製造2課"), ("a01", "鈴木", "製造1課"), ("A10", "田中", "製造1課"),
                     ("", "空", "製造1課"), ("Ｃ０３", "高橋", " 品質 ")])
    assert masters.api_getDeptList(ctx(db)) == [{"deptId": "品質", "deptName": "品質"},
                                                {"deptId": "製造1課", "deptName": "製造1課"},
                                                {"deptId": "製造2課", "deptName": "製造2課"}]
    assert masters.api_getWorkersByDept(ctx(db), "製造1課") == [
        {"workerId": "a01", "workerName": "鈴木", "deptId": "製造1課"},
        {"workerId": "A10", "workerName": "田中", "deptId": "製造1課"}]
    assert masters.api_getWorkersByDept(ctx(db), "品質")[0]["workerId"] == "C03"
    assert masters.api_getWorkersByDept(ctx(db), "") == []
    assert masters.api_getWorkersByDept(ctx(db), "無い") == []


def test_active_flag_and_numbers(db):
    for i, cells in enumerate([{"作業員ID": 101, "氏名": "甲", "部署ID": "A", "在籍フラグ": "TRUE"},
                               {"作業員ID": 102, "氏名": "乙", "部署ID": "A", "在籍フラグ": "0"},
                               {"作業員ID": 103, "氏名": "丙", "部署ID": "A", "在籍フラグ": ""}], start=1):
        db.add(Worker(sort_order=i, cells=[[h, v] for h, v in cells.items()]))
    db.commit()
    assert [w["workerId"] for w in masters.api_getWorkersByDept(ctx(db), "A")] == ["101", "103"]


def _like_jsonb(value):
    """PostgreSQL の JSONB のまね: オブジェクトのキーの順を保たない（ここでは並べ替える）。配列の順は保つ。"""
    return json.loads(json.dumps(value, sort_keys=True))


def test_worker_header_error(db):
    db.add(Worker(sort_order=1, cells=_like_jsonb([["コード", "1"], ["氏名", "甲"]])))
    db.commit()
    with pytest.raises(RpcError, match="^作業員マスタのヘッダが想定と違います（作業員コード/氏名/部署）。実際: コード, 氏名$"):
        masters.api_getDeptList(ctx(db))


def test_worker_header_order_survives_db_that_reorders_keys(db):
    """見出しの並び（シートの列の順）は、キーの順を保たない DB を通っても変わらない。"""
    # dict で持っていた頃は {"氏名", "コード"} がキーの並べ替えで「コード, 氏名」に変わった
    assert list(_like_jsonb({"氏名": "甲", "コード": "1"})) == ["コード", "氏名"]
    db.add(Worker(sort_order=1, cells=_like_jsonb([["氏名", "甲"], ["コード", "1"]])))
    db.commit()
    with pytest.raises(RpcError, match="実際: 氏名, コード$"):
        masters.api_getDeptList(ctx(db))


def test_worker_duplicate_header_reads_left_column(db):
    """同じ見出しの列が 2 つあれば、GAS の indexOf と同じく前（左）の列を読む。"""
    cols = [["作業員コード", "101"], ["氏名", "甲"], ["部署", "製造1課"], ["部署", "品質"], ["氏名", "乙"]]
    db.add(Worker(sort_order=1, cells=_like_jsonb(cols)))
    db.commit()
    assert masters.api_getWorkersByDept(ctx(db), "製造1課") == [
        {"workerId": "101", "workerName": "甲", "deptId": "製造1課"}]
    assert masters.api_getWorkersByDept(ctx(db), "品質") == []


def test_worker_info(db):
    add_workers(db, [("101", "甲", "A")])
    assert masters.api_getWorkerInfo(ctx(db)) is None  # 本番は Googleアカウントの列が無い
    db.query(Worker).delete()
    add_workers(db, [("101", "甲", "A")], {"101": {"Googleアカウント（任意）": " Shared@LineWorks-Local.info "}})
    assert masters.api_getWorkerInfo(ctx(db)) == {"workerId": "101", "workerName": "甲", "deptId": "A",
                                                  "deptName": "A", "email": USER}


def test_lookup_data(db):
    add_lookup(db, [("特別休暇理由", "忌引休暇", "2"), ("特別休暇理由", "結婚休暇", "1"), ("ほか", "x", "0"),
                    ("特別休暇理由", "その他", ""), ("特別休暇理由", "Ｘ", " 3 ")])
    data = masters.api_getLookupData(ctx(db))
    assert data["leaveKubun"] == ["全日休", "半日休"]
    assert data["leaveTypeHalf"] == ["有給消化（私用）", "特別休暇"]
    # SORT は Number(値 || 0) の昇順（空は 0）。値は normalize_
    assert data["specialReasons"] == ["その他", "結婚休暇", "忌引休暇", "X"]


def test_approver_list_and_current_info(db):
    add_stamps(db, [(" A@x.jp ", "s1", " 名前A "), ("", "s2", "空"), ("b@x.jp", "s3", "")])
    assert masters.api_getApproverList(ctx(db)) == [{"email": "A@x.jp", "name": "名前A"},
                                                     {"email": "b@x.jp", "name": ""}]
    assert masters.api_getCurrentApproverInfo(ctx(db, "a@x.jp")) == {"email": "a@x.jp", "name": "名前A"}
    assert masters.api_getCurrentApproverInfo(ctx(db, "b@x.jp")) == {"email": "b@x.jp", "name": "b@x.jp"}
    assert masters.api_getCurrentApproverInfo(ctx(db, "c@x.jp")) == {"email": "c@x.jp", "name": "c@x.jp"}


def test_roles(db):
    standard_masters(db)
    assert masters.api_isAdmin(ctx(db, ADMIN)) is True
    assert masters.api_isAdmin(ctx(db, "owner@lineworks-local.info")) is False  # BOOTSTRAP は足さない
    assert masters.api_isSomu(ctx(db, SOMU)) is True
    assert masters.api_isDeptApprover(ctx(db, BOSS)) is True
    assert masters.api_isDeptApprover(ctx(db, USER)) is False


def test_approver_profiles(db):
    standard_masters(db)
    res = masters.api_getApproverProfiles(ctx(db, ADMIN))
    assert res["isAdmin"] is True and res["currentEmail"] == ADMIN
    by_email = {p["email"]: p for p in res["profiles"]}
    assert by_email[BOSS] == {"email": BOSS, "name": "上司 一郎", "depts": ["製造1課", "製造2課"]}
    assert by_email["other@lineworks-local.info"]["name"] == "other@lineworks-local.info"
    assert masters.api_getApproverProfiles(ctx(db, BOSS)) == {
        "isAdmin": False, "currentEmail": BOSS, "profiles": [by_email[BOSS]]}
    assert masters.api_getApproverProfiles(ctx(db, USER))["profiles"] == []
