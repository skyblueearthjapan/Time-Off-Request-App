"""申請・編集・取消・1 件・一覧（実装仕様 8-5・8-6・4-4-b〜d。leaveRequest.gs）。"""
from __future__ import annotations

from datetime import date

import pytest
from test_domain_helpers import add_calendar, add_request, ctx, install_fakes, jst

from app.domain import requests
from app.models import LeaveRequest, SheetSyncState
from app.rpc import RpcError


@pytest.fixture
def rec(db, monkeypatch):
    add_calendar(db, [(date(2026, 1, 1), "祝日")])
    return install_fakes(monkeypatch)


def submit(db, **data):
    return requests.api_submitLeaveRequest(ctx(db), data)


# ===== 申請 =====


def test_submit_errors_are_200_results(db, rec):
    assert submit(db) == {"ok": False, "error": "休暇日の形式が不正です（yyyy-MM-dd）。"}
    assert submit(db, leaveDate="2026/11/30")["error"] == "休暇日の形式が不正です（yyyy-MM-dd）。"
    assert submit(db, leaveDate="2026-11-29") == {
        "ok": False, "error": "指定された休暇日は勤務日ではありません。過去日の場合は「過去申請」から選択してください。"}
    assert submit(db, leaveDate="2026-11-20")["error"].startswith("指定された休暇日は勤務日ではありません。")
    assert submit(db, leaveDate="2026-03-15", isRetroactive=True) == {
        "ok": False, "error": "過去日申請は今年度内（年度開始の3/16から本日まで）の日付のみ指定できます。"}
    # 真偽値の true だけが過去申請
    assert submit(db, leaveDate="2026-11-20", isRetroactive="true")["error"].startswith("指定された休暇日は勤務日")
    assert submit(db, leaveDate="2026-11-30", leaveType="振替休暇") == {
        "ok": False, "error": "振替休暇の場合は振替元出勤日を選択してください。"}
    assert requests.api_submitLeaveRequest(ctx(db), None)["ok"] is False
    assert db.query(LeaveRequest).count() == 0


def test_submit_substitute_monthly_period_message(db, rec):
    res = submit(db, leaveDate="2026-12-16", leaveType="振替休暇", substituteDate="2026-12-12")
    assert res == {"ok": False, "error": (
        "振替休暇は同一月次期間（前月16日〜当月15日）内で消化する必要があります。\n"
        "休暇日: 2026-12-16（期間: 2026-12-16〜2027-01-15）\n"
        "振替元: 2026-12-12（期間: 2026-11-16〜2026-12-15）")}


def test_submit_substitute_uniqueness(db, rec):
    add_request(db, "LV-20261120-001", leave_type="振替休暇", substitute=date(2026, 11, 28), worker_id="101",
                status="却下")
    res = submit(db, leaveDate="2026-12-01", leaveType="振替休暇", substituteDate="2026-11-28", workerId="101")
    assert res == {"ok": False, "error": "指定された振替元出勤日（2026-11-28）は既に別の振替休暇申請で使用されています。"}
    assert submit(db, leaveDate="2026-12-01", leaveType="振替休暇", substituteDate="2026-11-28",
                  workerId="102")["ok"] is True


def test_submit_writes_row_like_gas(db, rec):
    res = submit(db, leaveDate="2026-11-30", leaveType="有給消化（私用）", leaveKubun="半日休", halfDayType="午後",
                 deptId="製造1課", deptName="製造1課", workerId="0042", workerName="<b>山田</b>",
                 specialReason="", paidDetail="私用 ", substituteDate="2026-11-29", additionalDetail="捨てる")
    assert res == {"ok": True, "reqId": "LV-20261126-001"}
    row = db.get(LeaveRequest, "LV-20261126-001")
    assert row.status == "申請中" and row.fiscal_year == "2026" and row.cell_kinds["fiscal_year"] == "number"
    assert row.worker_id == "42" and row.cell_kinds["worker_id"] == "number"  # setValues と同じく数のセル
    assert row.worker_name == "<b>山田</b>"  # 入力はそのまま
    assert row.paid_detail == "私用 "  # 正規化しない
    assert row.leave_date == date(2026, 11, 30) and row.cell_kinds["leave_date"] == "serial"
    assert row.substitute_date == date(2026, 11, 29)  # 種類に関係なく書く
    assert row.additional_detail == ""
    assert row.applied_at == jst(2026, 11, 26, 10).astimezone(row.applied_at.tzinfo)
    assert row.created_by_email == "shared@lineworks-local.info"
    assert row.sheet_row is None
    assert db.get(SheetSyncState, ("leave_requests", "LV-20261126-001")).desired_version == 1
    assert rec.names() == ["mark"]  # 集計・GAS の仕事は積まない


def test_submit_retroactive_and_fiscal_year_of_leave_date(db, rec):
    res = submit(db, leaveDate="2026-03-16", isRetroactive=True)
    assert res["ok"] is True
    assert db.get(LeaveRequest, res["reqId"]).fiscal_year == "2026"


def test_submit_too_long_and_bad_substitute(db, rec):
    with pytest.raises(RpcError, match="^特別理由が長すぎます（50000字まで）。$"):
        submit(db, leaveDate="2026-11-30", specialReason="あ" * 50001)
    with pytest.raises(RpcError, match="^振替元出勤日の形式が不正です: abc$"):
        submit(db, leaveDate="2026-11-30", substituteDate="abc")


# ===== 1 件 =====


def test_get_by_id_shape_and_cells(db, rec):
    add_request(db, "LV-20261120-001", status="承認済", fiscal_year="2026", worker_id="101", dept="製造１課",
                approved_at=jst(2026, 11, 21, 8, 5, 9), approved_at2_text="11/22/2026", sheet_row=12,
                substitute=None, substitute_text="", special_reason=" 理由 ", paid_detail="詳細")
    r = requests.api_getLeaveRequestById(ctx(db), "LV-20261120-001")
    assert r["reqId"] == "LV-20261120-001" and r["fiscalYear"] == "2026" and r["deptId"] == "製造1課"
    assert r["leaveDate"] == "2026-11-27" and r["substituteDate"] == ""
    assert r["submittedAt"] == "2026-11-20 09:00:00" and r["approvedAt"] == "2026-11-21 08:05:09"
    assert r["approvedAt2"] == ""  # 文字のセルは ''
    assert r["specialReason"] == "理由" and r["rowNo"] == 12
    assert "isRetroactive" not in r
    assert requests.api_getLeaveRequestById(ctx(db), "LV-NOPE") is None
    assert requests.api_getLeaveRequestById(ctx(db), "") is None
    assert requests.api_getLeaveRequestById(ctx(db), " LV-20261120-001") is None  # 引数は正規化しない


def test_get_by_id_text_cells(db, rec):
    add_request(db, "LV-20260328-001", leave_date=None, leave_text="2026-03-28", applied_text="3/28/2026 9:00",
                fiscal_year="0", fy_kind="number")
    r = requests.api_getLeaveRequestById(ctx(db), "LV-20260328-001")
    assert r["leaveDate"] == "2026-03-28" and r["submittedAt"] == "3/28/2026 9:00" and r["fiscalYear"] == ""


# ===== 一覧 =====


def test_upcoming_leaves(db, rec):
    add_request(db, "LV-1", leave_date=date(2026, 11, 26), applied_at=jst(2026, 11, 27, 9), status="申請中")
    add_request(db, "LV-2", leave_date=date(2026, 11, 25))  # 昨日は除く
    add_request(db, "LV-3", leave_date=date(2026, 12, 1), status="取消")
    add_request(db, "LV-4", leave_date=None, leave_text="いつか", status="却下")  # Date でなければ載せる
    add_request(db, "LV-5", leave_date=date(2026, 11, 30), dept="製造2課", worker_id="102")
    res = requests.api_getUpcomingLeaves(ctx(db))
    assert [r["reqId"] for r in res] == ["LV-4", "LV-1", "LV-5"]
    assert res[0]["leaveDate"] == "" and res[1]["leaveDate"] == "2026/11/26"
    assert res[1]["isRetroactive"] is True and res[1]["daysAgo"] == 0
    assert set(res[1]) == {"reqId", "deptName", "workerName", "leaveKubun", "halfDayType", "leaveDate", "leaveType",
                           "status", "isRetroactive", "daysAgo"}
    assert [r["reqId"] for r in requests.api_getUpcomingLeaves(ctx(db), "製造2課", None)] == ["LV-5"]
    assert [r["reqId"] for r in requests.api_getUpcomingLeaves(ctx(db), "", "101")] == ["LV-4", "LV-1"]
    assert requests.api_getUpcomingLeaves(ctx(db), None, 101) == []  # 数の引数は文字と一致しない


def test_pending_sorted_by_real_datetime_desc(db, rec):
    add_request(db, "LV-A", status="申請中", applied_at=jst(2026, 12, 31, 9))
    add_request(db, "LV-B", status="申請中", applied_at=jst(2027, 1, 2, 9))
    add_request(db, "LV-C", status="申請中", applied_at=None, applied_text="x")
    add_request(db, "LV-D", status="承認済")
    res = requests.api_getAllPendingRequests(ctx(db, now=jst(2027, 1, 3)))
    assert [r["reqId"] for r in res] == ["LV-B", "LV-A", "LV-C"]  # K-10: 年をまたいでも新しい順
    assert res[0]["submittedAt"] == "01/02 09:00" and res[2]["submittedAt"] == ""


def test_pending_gas_string_sort_when_quirk_off(db, rec, monkeypatch):
    from app.domain import quirks

    monkeypatch.setattr(quirks, "PENDING_SORT_WITH_YEAR", False)
    add_request(db, "LV-A", status="申請中", applied_at=jst(2026, 12, 31, 9))
    add_request(db, "LV-B", status="申請中", applied_at=jst(2027, 1, 2, 9))
    res = requests.api_getAllPendingRequests(ctx(db, now=jst(2027, 1, 3)))
    assert [r["reqId"] for r in res] == ["LV-A", "LV-B"]


# ===== 編集 =====


def test_edit_errors(db, rec):
    add_request(db, "LV-1", status="取消")
    add_request(db, "LV-2", status="申請中")
    c = ctx(db)
    for args, msg in [
        (("LV-X", {}), "申請が見つかりません。"),
        ((None, {}), "申請が見つかりません。"),
        (("LV-1", {}), "取消/却下済みの申請は編集できません。"),
        (("LV-2", {"leaveDate": "2026/12/01"}), "休暇日の形式が不正です: 2026/12/01"),
        (("LV-2", {"leaveKubun": "時間休"}), "休暇区分が不正です: 時間休"),
        (("LV-2", {"leaveKubun": "半日休"}), "半日区分を選択してください。"),
        (("LV-2", {"substituteDate": "x"}), "振替元出勤日の形式が不正です: x"),
    ]:
        with pytest.raises(RpcError) as e:
            requests.api_editLeaveRequest(c, *args)
        assert e.value.message == msg


def test_edit_substitute_checks(db, rec):
    add_request(db, "LV-S", status="申請中", leave_type="振替休暇", leave_date=date(2026, 12, 1),
                substitute=date(2026, 11, 28))
    add_request(db, "LV-T", leave_type="振替休暇", substitute=date(2026, 11, 29))
    add_request(db, "LV-N", status="申請中", leave_type="振替休暇", substitute=None)
    c = ctx(db)
    with pytest.raises(RpcError, match="^振替休暇には振替元出勤日が必須です。"):
        requests.api_editLeaveRequest(c, "LV-N", {})
    with pytest.raises(RpcError, match="^振替休暇は同一月次期間"):
        requests.api_editLeaveRequest(c, "LV-S", {"leaveDate": "2026-12-16"})
    with pytest.raises(RpcError, match=r"^指定された振替元出勤日（2026-11-29）は既に"):
        requests.api_editLeaveRequest(c, "LV-S", {"substituteDate": "2026-11-29"})
    # 自分の振替元と同じなら確かめない
    assert requests.api_editLeaveRequest(c, "LV-S", {"substituteDate": "2026-11-28"})["ok"] is True
    assert requests.api_editLeaveRequest(c, "LV-S", {"substituteDate": "2026-11-22"}) == {"ok": True,
                                                                                          "newStatus": "申請中"}
    assert db.get(LeaveRequest, "LV-S").substitute_date == date(2026, 11, 22)


def test_edit_approved_resets_and_enqueues(db, rec):
    add_request(db, "LV-1", status="承認済", approved_at=jst(2026, 11, 21), approved_by1="a@x", approved_by2="総務",
                approved_by2_email="s@x", approved_at2=jst(2026, 11, 22), sign_image_url="https://s",
                pdf_url="https://drive/d/abc/view", pdf_file_id="abc", half="午後", kubun="半日休",
                leave_date=date(2026, 12, 3))
    row = db.get(LeaveRequest, "LV-1")
    row.sign_png = b"png"
    db.commit()
    res = requests.api_editLeaveRequest(ctx(db), "LV-1", {"leaveDate": "2027-03-16", "leaveKubun": "全日休"})
    assert res == {"ok": True, "newStatus": "申請中"}
    row = db.get(LeaveRequest, "LV-1")
    assert row.status == "申請中" and row.leave_date == date(2027, 3, 16) and row.fiscal_year == "2027"
    assert row.half_day_type == "" and row.leave_kubun == "全日休"
    assert row.approved_at is None and row.approved_at2 is None and row.approved_by1_email == ""
    assert row.approved_by2 == "" and row.approved_by2_email == "" and row.sign_image_url == ""
    assert row.sign_png is None
    assert row.pdf_url == "https://drive/d/abc/view"  # asis: 残す
    assert rec.names() == ["mark", "unapprove", "summary"]
    _, args, kwargs = rec.calls[1]
    assert args == ("LV-1",) and kwargs == {"reason": "edit", "old_leave_date": "2026-12-03",
                                            "old_pdf_url": "https://drive/d/abc/view", "old_pdf_file_id": "abc",
                                            "clear_sign": True}


def test_edit_pending_without_changes_still_marks(db, rec):
    add_request(db, "LV-1", status="申請中", half="午前")
    res = requests.api_editLeaveRequest(ctx(db), "LV-1", None)
    assert res == {"ok": True, "newStatus": "申請中"}
    assert rec.names() == ["mark", "summary"]
    # 半日区分が空なら書かない
    requests.api_editLeaveRequest(ctx(db), "LV-1", {"leaveKubun": "半日休", "halfDayType": "午後"})
    assert db.get(LeaveRequest, "LV-1").half_day_type == "午後"


# ===== 取消 =====


def test_delete(db, rec):
    add_request(db, "LV-1", status="承認済", pdf_url="https://p", pdf_file_id="f1", leave_date=date(2026, 12, 3))
    add_request(db, "LV-2", status="取消")
    add_request(db, "LV-3", status="却下")
    assert requests.api_deleteLeaveRequest(ctx(db), "LV-1") == {"ok": True}
    row = db.get(LeaveRequest, "LV-1")
    assert row.status == "取消" and row.pdf_url == "https://p"
    assert rec.names() == ["mark", "unapprove", "summary"]
    assert rec.calls[1][2] == {"reason": "cancel", "old_leave_date": "2026-12-03", "old_pdf_url": "https://p",
                               "old_pdf_file_id": "f1", "clear_sign": False}
    with pytest.raises(RpcError, match="^既に取消済みです。$"):
        requests.api_deleteLeaveRequest(ctx(db), "LV-2")
    with pytest.raises(RpcError, match="^申請が見つかりません。$"):
        requests.api_deleteLeaveRequest(ctx(db), "LV-9")
    assert requests.api_deleteLeaveRequest(ctx(db), "LV-3") == {"ok": True}  # 却下も取消できる


# ===== 入力の形（JS の \d・$ と同じ: 半角の数字だけ、末尾の改行を許さない） =====


def test_input_patterns_match_js_semantics():
    from app.domain import calendar, fiscal, jsdate
    from app.services import sheet_codec

    for pattern in (requests.YMD, fiscal._YMD, jsdate.YMD_RE):
        assert pattern.match("2026-11-27")
        for bad in ("2026-11-27\n", "２０２６-11-27", "2026-11-2٧", "2026-11-27 "):
            assert not pattern.match(bad), (pattern.pattern, bad)
    assert calendar._YM_DASH.match("2026-11") and not calendar._YM_DASH.match("2026-11\n")
    assert calendar._YM_PLAIN.match("202611") and not calendar._YM_PLAIN.match("２０２６11")
    assert sheet_codec.REQ_ID_RE.match("LV-20261126-1000") and not sheet_codec.REQ_ID_RE.match("LV-20261126-001\n")
    assert not sheet_codec.NUMBER_TEXT_RE.match("１２")
