"""勤務日・休日・祝日（実装仕様 8-3。masters.gs:220-560）。"""
from __future__ import annotations

from datetime import date

from test_domain_helpers import add_calendar, add_request, ctx, jst

from app.domain import calendar, masters


def test_holidays_only_two_years_of_current_fy():
    h = calendar.japanese_holidays(date(2026, 11, 26))  # 年度 2026 → 2026 年と 2027 年
    assert "2026-01-01" in h and "2027-11-03" in h
    assert "2025-11-03" not in h and "2028-01-01" not in h
    h = calendar.japanese_holidays(date(2026, 3, 15))  # 年度 2025 → 2025 年と 2026 年
    assert "2025-01-01" in h and "2027-01-01" not in h


def test_substitute_holiday_and_golden_week():
    h = calendar.japanese_holidays(date(2026, 11, 26))
    # 2026-05-03（日）憲法記念日 → 5/4・5/5 は祝日なので 5/6 が振替休日
    assert h["2026-05-06"] == "振替休日"
    # 2027-03-21（日）春分の日 → 3/22 振替休日
    assert h["2027-03-21"] == "春分の日" and h["2027-03-22"] == "振替休日"
    # 2026-11-23 は月曜（振替なし）
    assert "2026-11-24" not in h


def test_kokumin_no_kyujitsu_2026_september():
    h = calendar.japanese_holidays(date(2026, 11, 26))
    assert h["2026-09-21"] == "敬老の日"
    assert h["2026-09-22"] == "国民の休日"
    assert h["2026-09-23"] == "秋分の日"


def test_happy_monday_and_equinox_formula():
    h = calendar.holidays_for_year(2027)
    assert h["2027-01-11"] == "成人の日"
    assert h["2027-07-19"] == "海の日"
    assert h["2027-10-11"] == "スポーツの日"
    assert h["2027-09-23"] == "秋分の日"


def test_working_days_empty_calendar(db):
    assert calendar.working_days(db, date(2026, 11, 26)) == []


def test_working_days_range_and_kubun(db):
    add_calendar(db, [
        (date(2026, 11, 27), "休日"),     # 金曜でも休日
        (date(2026, 11, 28), "出勤土曜"),  # 土曜でも勤務日
        (date(2026, 12, 1), "祝日"),
        (None, "休日"),                   # 読めない行は飛ばす
        (date(2026, 12, 2), "休日"),
        (date(2026, 12, 2), ""),          # 同じ日は下の行が勝つ
    ])
    days = calendar.working_days(db, date(2026, 11, 26))
    assert days[:4] == ["2026-11-26", "2026-11-28", "2026-11-30", "2026-12-02"]
    assert days[-1] == "2027-03-15"  # 年度末 3/15 まで（月曜）
    assert "2027-03-16" not in days
    assert "2027-01-01" not in days  # 計算の祝日
    assert "2026-11-29" not in days  # 日曜


def test_working_days_from_3_15_stops_same_day(db):
    add_calendar(db, [(date(2026, 1, 1), "祝日")])
    assert calendar.working_days(db, date(2027, 3, 15)) == ["2027-03-15"]
    assert calendar.working_days(db, date(2027, 3, 13)) == ["2027-03-15"]  # 土曜から


def test_api_working_days_uses_jst_today(db):
    add_calendar(db, [(date(2026, 1, 1), "祝日")])
    # 11/27 0:30 JST（UTC では 11/26）→ 今日は 11/27
    res = masters.api_getWorkingDays(ctx(db, now=jst(2026, 11, 27, 0, 30)))
    assert res[0] == "2026-11-27"
    res = masters.api_getWorkingDays(ctx(db, now=jst(2026, 11, 26, 8, 59)))
    assert res[0] == "2026-11-26"


def test_holiday_days_labels_and_range(db):
    add_calendar(db, [(date(2026, 3, 21), "休日"), (date(2026, 3, 28), "出勤土曜"), (date(2026, 3, 22), "")])
    res = calendar.holiday_days(db, date(2026, 3, 30))
    assert res[0] == {"date": "2026-03-20", "label": "祝日（春分の日）"}
    assert {"date": "2026-03-21", "label": "休日"} in res
    assert {"date": "2026-03-22", "label": "日曜"} in res  # カレンダーにある日で区分が空の日曜
    assert all(r["date"] != "2026-03-28" for r in res)  # 出勤土曜
    assert res[-1]["date"] <= "2026-04-06"  # 今日＋7 日まで
    assert {"date": "2026-04-05", "label": "日曜"} in res
    assert {"date": "2026-04-04", "label": "土曜"} in res


def test_holiday_days_beyond_fiscal_year_end(db):
    res = calendar.holiday_days(db, date(2027, 3, 14))
    assert res[-1]["date"] == "2027-03-21"  # 年度末を越えても止めない


def test_holiday_days_excludes_used_substitute(db):
    add_request(db, "LV-20261101-001", leave_type="振替休暇", substitute=date(2026, 11, 22), worker_id="101",
                status="却下")
    add_request(db, "LV-20261101-002", leave_type="振替休暇", substitute=date(2026, 11, 15), worker_id="101",
                status="取消")
    add_request(db, "LV-20261101-003", leave_type="振替休暇", substitute=None, substitute_text="2026-11-21",
                worker_id="101", status="申請中")
    add_request(db, "LV-20261101-004", leave_type="振替休暇", substitute=date(2026, 11, 14), worker_id="102")
    res = [r["date"] for r in calendar.holiday_days(db, date(2026, 11, 26), "101")]
    assert "2026-11-22" not in res  # 却下も使用済み
    assert "2026-11-21" not in res  # 文字のセルも日付として読む
    assert "2026-11-15" in res      # 取消は再利用できる
    assert "2026-11-14" in res      # ほかの作業員
    res_num = [r["date"] for r in calendar.holiday_days(db, date(2026, 11, 26), 101)]
    assert "2026-11-22" not in res_num  # 数の作業員 ID も文字にして比べる


def test_past_working_days(db):
    add_calendar(db, [(date(2026, 11, 2), "休日")])
    res = calendar.past_working_days(db, date(2026, 11, 26), "2026-11")
    assert res[0] == {"date": "2026-11-04", "dow": 3, "label": "2026-11-04（水）"}  # 11/2 休日、11/3 文化の日
    assert res[-1]["date"] == "2026-11-26"
    assert calendar.past_working_days(db, date(2026, 11, 26), 202603)[0]["date"] == "2026-03-16"
    assert calendar.past_working_days(db, date(2026, 11, 26), "2026-02") == []  # 年度より前
    assert calendar.past_working_days(db, date(2026, 11, 26), "2026-12") == []  # 今日より後
    assert calendar.past_working_days(db, date(2026, 11, 26), None)[-1]["date"] == "2026-11-26"  # 今月
    assert calendar.past_working_days(db, date(2026, 11, 26), "2026-13")== []  # 2027-01 に繰り越す
