"""業務の決まった値（src/config.gs の定数の写し）と、DB の enum の値。

業務の計算に使う定数（マイルストーン・期間・警告メールの日など）は担当 W が domain/ に置く。
"""
from __future__ import annotations

# ===== シートのタブ名（src/config.gs:4-14 の SHEET） =====
TAB_LEAVE_REQUEST = "T_LEAVE_REQUEST"
TAB_LEAVE_SUMMARY = "V_LEAVE_SUMMARY"
TAB_SETTINGS = "M_SYSTEM_SETTING"
TAB_LOOKUP = "M_LOOKUP"
TAB_WORKER = "作業員マスタ"
TAB_CALENDAR = "M_CALENDAR"
TAB_STAMP = "M_STAMP"
TAB_DEPT_APPROVERS = "M_DEPT_APPROVERS"

# ===== 承認状態（src/config.gs:16-21 の STATUS） =====
STATUS_SUBMITTED = "申請中"
STATUS_APPROVED = "承認済"
STATUS_CANCELED = "取消"
STATUS_REJECTED = "却下"
LEAVE_STATUSES: tuple[str, ...] = (STATUS_SUBMITTED, STATUS_APPROVED, STATUS_CANCELED, STATUS_REJECTED)

# ===== 休暇区分・半日区分・休暇種類（src/config.gs:24-40） =====
DAY_FULL = "全日休"
DAY_HALF = "半日休"
HALF_AM = "午前"
HALF_PM = "午後"
KIND_PAID = "有給消化（私用）"
KIND_SUBSTITUTE = "振替休暇"
KIND_SPECIAL = "特別休暇"

# 仕事（jobs）
JOB_STATES: tuple[str, ...] = ("pending", "running", "done", "failed", "skipped")
JOB_LANES: tuple[str, ...] = ("gas", "default")
PRIORITY_NORMAL = 0

# 行の反映（sheet_sync_state）
SYNC_STATES: tuple[str, ...] = ("ok", "retry", "blocked")

# 冪等キー
IDEMPOTENCY_STATES: tuple[str, ...] = ("running", "done")
