"""集計（V_LEAVE_SUMMARY）・警告・期間別・マイルストーン（実装仕様 6-3・8-4。leaveSummary.gs・leaveStatusApi.gs）。"""
from __future__ import annotations

from datetime import date

from test_domain_helpers import add_request, jst

from app.domain import summaries


def test_summary_rows_counts_and_keys(db):
    add_request(db, "LV-1", worker_id="101", leave_date=date(2026, 5, 1), sheet_row=2, dept="製造1課")
    add_request(db, "LV-2", worker_id="101", leave_date=date(2026, 7, 1), kubun="半日休", sheet_row=3,
                dept="製造2課")  # 部署は最初の行の値
    add_request(db, "LV-3", worker_id="101", leave_type="振替休暇", sheet_row=4)
    add_request(db, "LV-4", worker_id="101", leave_type="特別休暇", sheet_row=5)
    add_request(db, "LV-5", worker_id="101", status="申請中", sheet_row=6)
    add_request(db, "LV-6", worker_id="101", leave_date=None, leave_text="いつか", sheet_row=7)  # Date でない
    add_request(db, "LV-7", worker_id="101", fiscal_year="2025", leave_date=date(2026, 3, 20), sheet_row=8)  # 列の値
    add_request(db, "LV-8", worker_id="99", fiscal_year="2026", fy_kind="string", sheet_row=9)
    rows = summaries.summary_rows(db, jst(2026, 11, 26, 10))
    assert [(r["fiscalYear"], r["workerId"]) for r in rows] == [(2025, "101"), (2026, "101"), ("2026", "99")]
    r = rows[1]
    assert r == {"fiscalYear": 2026, "deptId": "製造1課", "deptName": "製造1課", "workerId": "101",
                 "workerName": "山田太郎", "paidCount": 2.5, "paidHalfCount": 1, "substituteCount": 1,
                 "specialCount": 1, "totalCount": 5, "lastPaidDate": "2026-07-01", "warnLevel": ""}  # 8/31 の 2 回は満たす
    assert rows[0]["paidCount"] == 1 and isinstance(rows[0]["paidCount"], int)


def test_warn_levels_after_deadlines():
    now = jst(2027, 1, 10)  # 12/31 の期限（3 回）を過ぎた
    assert summaries.compute_warn_level(0, 2026, now) == "危険"
    assert summaries.compute_warn_level(1, 2026, now) == "警告"
    assert summaries.compute_warn_level(2.5, 2026, now) == "注意"
    assert summaries.compute_warn_level(3, 2026, now) == ""
    assert summaries.compute_warn_level(0, 2026, jst(2026, 6, 1)) == ""  # 過ぎた期限が無い
    assert summaries.compute_warn_level(0, 2027, jst(2027, 1, 10)) == ""


def test_quarter_status_half_day_counts_one_and_11_30():
    dates = [date(2026, 4, 1), date(2026, 7, 1), date(2026, 11, 30), date(2026, 12, 1)]
    qs = summaries.compute_quarter_status(dates, 2026, jst(2026, 12, 1, 0, 0))
    assert [q["id"] for q in qs] == ["Q1", "SUMMER", "Q3", "WINTER", "YEAREND"]
    assert [q["periodCount"] for q in qs] == [1, 1, 1, 1, 0]
    assert [q["cumulativeCount"] for q in qs] == [1, 2, 3, 4, 4]
    q3, winter = qs[2], qs[3]
    assert q3["isPast"] and not q3["isCurrent"] and q3["level"] == "ok" and q3["shortage"] == 0
    assert winter["isCurrent"] and winter["level"] == "ok"
    assert qs[4]["isFuture"] and qs[4]["level"] == "" and qs[4]["label"] == "年度末(2/16-3/15)"
    qs = summaries.compute_quarter_status([], 2026, jst(2026, 9, 1))
    assert qs[0]["level"] == "warn" and qs[0]["overdue"] is True
    assert qs[1]["level"] == "danger" and qs[2]["isCurrent"] and qs[2]["level"] == "danger"
    assert set(qs[0]) == {"id", "label", "targetCount", "periodCount", "cumulativeCount", "shortage", "level", "taken",
                          "overdue", "isCurrent", "isPast", "isFuture"}


def test_milestones():
    ms = summaries.build_milestones(1.5, 2026, jst(2026, 9, 1))
    assert ms[0] == {"label": "6月末", "month": 6, "day": 30, "dateText": "6/30", "required": 1, "met": True,
                     "passed": True, "deficit": 0, "remaining": 0, "state": "done"}
    assert ms[1]["state"] == "note" and ms[1]["deficit"] == 0.5
    assert ms[2]["state"] == "future" and ms[3]["dateText"] == "2/15"
    assert summaries.build_milestones(0, 2026, jst(2027, 1, 1))[2]["state"] == "risk"


def test_last_paid_date_display():
    assert summaries.last_paid_date_display("2026-11-27", "gas", "Etc/GMT") == \
        "Fri Nov 27 2026 09:00:00 GMT+0900 (Japan Standard Time)"
    assert summaries.last_paid_date_display("2026-03-05", "gas", "Asia/Tokyo") == \
        "Thu Mar 05 2026 00:00:00 GMT+0900 (Japan Standard Time)"
    assert summaries.last_paid_date_display("2026-11-27", "iso") == "2026-11-27"
    assert summaries.last_paid_date_display("", "gas") == ""
    # 既定は iso（2026-11-26 オーナーの決定: 最終有給日は yyyy-MM-dd で出す）
    assert summaries.last_paid_date_display("2026-11-27") == "2026-11-27"
