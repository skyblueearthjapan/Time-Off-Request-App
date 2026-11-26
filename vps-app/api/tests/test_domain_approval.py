"""1 次承認・2 次承認・2 次承認待ち（実装仕様 8-7・8-8・4-4-e・K-4）。"""
from __future__ import annotations

import base64
from datetime import date

import pytest
from test_domain_helpers import ADMIN, SOMU, USER, add_request, ctx, install_fakes, jst, standard_masters

from app.domain import approval, quirks
from app.models import LeaveRequest
from app.rpc import RpcError

PNG = b"\x89PNG\r\n\x1a\nfake"
DATA_URL = "data:image/png;base64," + base64.b64encode(PNG).decode()


@pytest.fixture
def rec(db, monkeypatch):
    standard_masters(db)
    return install_fakes(monkeypatch)


def err(fn, *args) -> str:
    with pytest.raises(RpcError) as e:
        fn(*args)
    return e.value.message


# ===== 1 次承認 =====


def test_approve1_errors(db, rec):
    c = ctx(db)
    assert err(approval.api_approveLeaveRequest, c, "") == "REQ_IDが指定されていません。"
    assert err(approval.api_approveLeaveRequest, c, "LV-1") == "申請データがありません。"
    add_request(db, "LV-1", status="承認済")
    add_request(db, "LV-2", status="取消")
    add_request(db, "LV-3", status="却下")
    assert err(approval.api_approveLeaveRequest, c, "LV-9") == "申請が見つかりません: LV-9"
    assert err(approval.api_approveLeaveRequest, c, 5) == "申請が見つかりません: 5"
    assert err(approval.api_approveLeaveRequest, c, "LV-1") == "この申請はすでに承認済みです。"
    assert err(approval.api_approveLeaveRequest, c, "LV-2") == "この申請は取消済みです。"
    assert err(approval.api_approveLeaveRequest, c, "LV-3") == "この申請は取消済みです。"
    assert rec.calls == []


def test_approve1_handwritten_sign(db, rec):
    add_request(db, "LV-1", status="申請中", sign_image_url="old")
    res = approval.api_approveLeaveRequest(ctx(db), "LV-1", DATA_URL, "boss@x")
    assert res == {"ok": True, "reqId": "LV-1"}
    row = db.get(LeaveRequest, "LV-1")
    assert row.status == "承認済" and row.approved_by1_email == "boss@x" and row.sign_image_url == ""
    assert row.approved_at.astimezone(jst(2026, 1, 1).tzinfo) == jst(2026, 11, 26, 10)
    assert row.cell_kinds["approved_at"] == "serial"
    assert row.sign_png == PNG
    assert rec.names() == ["mark", "approve1", "summary"]


def test_approve1_stamp_and_bad_sign(db, rec):
    add_request(db, "LV-1", status="申請中")
    add_request(db, "LV-2", status="申請中")
    approval.api_approveLeaveRequest(ctx(db), "LV-1", "", None)
    row = db.get(LeaveRequest, "LV-1")
    assert row.sign_png is None and row.approved_by1_email == ""
    approval.api_approveLeaveRequest(ctx(db), "LV-2", "data:image/png;base64,@@@", "a@x")
    assert db.get(LeaveRequest, "LV-2").sign_png is None  # 読めなければ入れない（GAS も URL 空で続ける）


def test_approve1_sign_too_large(db, rec, monkeypatch):
    monkeypatch.setattr(quirks, "SIGN_MAX_CHARS", 10)
    add_request(db, "LV-1", status="申請中")
    assert err(approval.api_approveLeaveRequest, ctx(db), "LV-1", DATA_URL, "a@x") == "サイン画像が大きすぎます。"


def test_decode_sign_without_prefix():
    assert approval.decode_sign(base64.b64encode(PNG).decode()) == PNG
    assert approval.decode_sign(123) is None


# ===== 2 次承認待ち =====


def test_pending2(db, rec):
    add_request(db, "LV-1", leave_date=date(2026, 12, 2), approved_at=jst(2026, 11, 21, 8, 5),
                applied_at=jst(2026, 12, 3, 9))
    add_request(db, "LV-2", leave_date=date(2026, 12, 1), approved_at2=jst(2026, 11, 22))
    add_request(db, "LV-3", leave_date=date(2026, 12, 1), approved_at2_text="  ")  # 空白だけは未承認
    add_request(db, "LV-4", leave_date=None, leave_text="いつか", approved_at2_text="済")
    add_request(db, "LV-5", status="申請中")
    assert err(approval.api_getPending2ndApprovals, ctx(db, USER)) == "権限がありません（総務または管理者のみ）"
    res = approval.api_getPending2ndApprovals(ctx(db, SOMU))
    assert [r["reqId"] for r in res] == ["LV-3", "LV-1"]
    assert res[1] == {"reqId": "LV-1", "deptName": "製造1課", "workerName": "山田太郎", "leaveDate": "2026/12/02",
                      "leaveKubun": "全日休", "leaveType": "有給消化（私用）", "approvedAt": "2026/11/21 08:05",
                      "isRetroactive": True, "daysAgo": 0}
    assert res[0]["approvedAt"] == ""
    assert approval.api_getPending2ndApprovals(ctx(db, ADMIN))


# ===== 2 次承認 =====


def test_approve2_errors_in_gas_order(db, rec):
    assert err(approval.api_approveLeaveRequest2, ctx(db, USER), "", "x") == "REQ_IDが指定されていません。"
    assert err(approval.api_approveLeaveRequest2, ctx(db, USER), "LV-1", "x") == "権限がありません（総務または管理者のみ）"
    assert err(approval.api_approveLeaveRequest2, ctx(db, SOMU), "LV-1", "x") == "申請データがありません。"
    add_request(db, "LV-1", status="申請中")
    add_request(db, "LV-2", approved_at2=jst(2026, 11, 22))
    assert err(approval.api_approveLeaveRequest2, ctx(db, SOMU), "LV-9", "x") == "申請が見つかりません: LV-9"
    assert err(approval.api_approveLeaveRequest2, ctx(db, SOMU), "LV-1", "x") == "1次承認されていない申請です。"
    assert err(approval.api_approveLeaveRequest2, ctx(db, SOMU), "LV-2", "x") == "この申請はすでに2次承認済みです。"
    assert rec.calls == []


def test_approve2_uses_login_email_and_stamp_name(db, rec):
    add_request(db, "LV-1", approved_at=jst(2026, 11, 21))
    res = approval.api_approveLeaveRequest2(ctx(db, SOMU), "LV-1", "someone-else@x")
    assert res == {"ok": True, "reqId": "LV-1"}
    row = db.get(LeaveRequest, "LV-1")
    assert row.approved_by2 == "総務 花子" and row.approved_by2_email == SOMU
    assert row.approved_at2 is not None and row.cell_kinds["approved_at2"] == "serial"
    assert rec.names() == ["mark", "approve2"]  # 集計は予約しない
    assert err(approval.api_approveLeaveRequest2, ctx(db, SOMU), "LV-1", "x") == "この申請はすでに2次承認済みです。"


def test_approve2_name_falls_back_to_email(db, rec, monkeypatch):
    add_request(db, "LV-1")
    approval.api_approveLeaveRequest2(ctx(db, ADMIN), "LV-1", None)  # 印の名前が空 → メール
    assert db.get(LeaveRequest, "LV-1").approved_by2 == ADMIN
    monkeypatch.setattr(quirks, "SECOND_APPROVER_FROM_LOGIN", False)
    add_request(db, "LV-2")
    approval.api_approveLeaveRequest2(ctx(db, ADMIN), "LV-2", "SOMU@lineworks-local.info")
    row = db.get(LeaveRequest, "LV-2")
    assert row.approved_by2 == "総務 花子" and row.approved_by2_email == "SOMU@lineworks-local.info"
