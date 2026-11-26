"""年度・月次期間・過去申請・採番の日本時間の境目（実装仕様 8-0〜8-2・12-2 W）。"""
from __future__ import annotations

from datetime import date

from test_domain_helpers import add_calendar, add_request, ctx, install_fakes, jst

from app.domain import requests, summaries
from app.domain.fiscal import days_ago, fiscal_year, fy_start, is_retroactive, is_within_current_fy, monthly_period


def test_fiscal_year_boundary_3_15_and_3_16():
    assert fiscal_year(date(2026, 3, 15)) == 2025
    assert fiscal_year(date(2026, 3, 16)) == 2026
    assert fiscal_year(date(2027, 1, 1)) == 2026
    assert fy_start(date(2027, 2, 1)) == date(2026, 3, 16)


def test_fiscal_year_of_jst_moment_at_0_30():
    # 3/16 0:30 JST は UTC では 3/15 15:30。年度は JST の日付で決める
    assert fiscal_year(jst(2026, 3, 16, 0, 30)) == 2026
    assert fiscal_year(jst(2026, 3, 15, 23, 30)) == 2025


def test_monthly_period_15_and_16():
    assert monthly_period(date(2026, 4, 15)) == (date(2026, 3, 16), date(2026, 4, 15))
    assert monthly_period(date(2026, 4, 16)) == (date(2026, 4, 16), date(2026, 5, 15))
    # 15 日 23:30 JST と 16 日 0:30 JST
    assert monthly_period(jst(2026, 4, 15, 23, 30)) == (date(2026, 3, 16), date(2026, 4, 15))
    assert monthly_period(jst(2026, 4, 16, 0, 30)) == (date(2026, 4, 16), date(2026, 5, 15))
    # 年またぎ
    assert monthly_period(date(2027, 1, 10)) == (date(2026, 12, 16), date(2027, 1, 15))
    assert monthly_period(date(2026, 12, 20)) == (date(2026, 12, 16), date(2027, 1, 15))


def test_is_within_current_fy_ends_and_string_compare():
    today = date(2026, 11, 26)
    assert is_within_current_fy("2026-03-16", today)
    assert not is_within_current_fy("2026-03-15", today)
    assert is_within_current_fy("2026-11-26", today)
    assert not is_within_current_fy("2026-11-27", today)
    assert is_within_current_fy("2026-04-31", today)  # 文字で比べる（今のまま）
    assert not is_within_current_fy("2026/11/01", today)
    assert not is_within_current_fy(None, today)


def test_is_retroactive_and_days_ago():
    assert is_retroactive(jst(2026, 11, 26, 0, 30), "2026-11-25")
    assert not is_retroactive(jst(2026, 11, 26, 0, 30), "2026-11-26")
    assert not is_retroactive("", "2026-11-25")
    assert is_retroactive("2026/11/26 10:00", "2026-11-25")
    assert not is_retroactive("読めない", "2026-11-25")
    today = date(2026, 11, 26)
    assert days_ago("2026-11-20", today) == 6
    assert days_ago("2026-11-26", today) == 0
    assert days_ago("2026-11-30", today) == 0
    assert days_ago("", today) == 0
    assert days_ago("abc", today) == 0


def test_req_id_at_0_30_jst_uses_jst_date(db, monkeypatch):
    install_fakes(monkeypatch)
    add_calendar(db, [(date(2026, 1, 1), "祝日")])
    now = jst(2026, 11, 27, 0, 30)  # UTC では 11/26 15:30
    res = requests.api_submitLeaveRequest(ctx(db, now=now), {"leaveDate": "2026-11-30", "leaveType": "有給消化（私用）"})
    assert res == {"ok": True, "reqId": "LV-20261127-001"}


def test_req_id_next_number_counts_cancelled_and_pads(db, monkeypatch):
    install_fakes(monkeypatch)
    add_calendar(db, [(date(2026, 1, 1), "祝日")])
    add_request(db, "LV-20261126-007", status="取消")
    add_request(db, "LV-20261126-x01")  # 数字で始まらない連番は数えない
    add_request(db, "LV-20261125-050")
    res = requests.api_submitLeaveRequest(ctx(db), {"leaveDate": "2026-11-30"})
    assert res["reqId"] == "LV-20261126-008"
    add_request(db, "LV-20261126-999")
    res = requests.api_submitLeaveRequest(ctx(db), {"leaveDate": "2026-11-30"})
    assert res["reqId"] == "LV-20261126-1000"


def test_warn_deadline_23_59_59_and_next_day():
    # 6/30 23:59:59 JST まではまだ、7/1 0:00 は過ぎた
    assert summaries.compute_warn_level(0, 2026, jst(2026, 6, 30, 23, 59, 58)) == ""
    assert summaries.compute_warn_level(0, 2026, jst(2026, 6, 30, 23, 59, 59)) == "注意"
    assert summaries.compute_warn_level(0, 2026, jst(2026, 7, 1, 0, 0)) == "注意"
    assert summaries.compute_warn_level(1, 2026, jst(2026, 7, 1, 0, 0)) == ""
    # 2/15（翌年）の期限
    assert summaries.compute_warn_level(1, "2026", jst(2027, 2, 16, 0, 0)) == "危険"
    assert summaries.compute_warn_level(2, 2026, jst(2027, 2, 16, 0, 0)) == "警告"
