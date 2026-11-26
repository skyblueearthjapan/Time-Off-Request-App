"""権限。GAS の src/auth.gs（isAdmin_・isSomu_・canAccessAdmin_・getDeptApproverDeptsForEmail_）を写す（仕様書 8-9）。

- 管理者 is_admin = M_SYSTEM_SETTING の ADMIN_EMAILS だけ（BOOTSTRAP_ADMINS は足さない＝画面の振る舞いを変えない）。
- 総務 is_somu = SOMU_EMAILS。
- 古い管理画面 can_access_admin（?page=admin）= ALLOW_ADMIN_PAGE_EMAILS。
- 部署承認者 = dept_approvers の B 列（approverEmails、カンマ区切り）にメールがある。
- 運用者 is_operator = SOMU_EMAILS・ADMIN_EMAILS・BOOTSTRAP_ADMINS のどれかに含まれる（06 の K-9。ops_* と運用の欄だけ）。
- 一覧が空なら誰も当てはまらない（空で全員にしない。罠 35）。メールの空も偽。
- 一覧は GAS と同じく normalize_ → split(',') → trim → toLowerCase。
- 表は config_sync・master_sync（担当 D）が取り込んだもの（settings・dept_approvers）。
"""
from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.orm import Session

from app.config import Settings, get_settings
from app.jsutil import js_string, normalize_gas
from app.models import DeptApprover, Setting

KEY_ADMIN = "ADMIN_EMAILS"
KEY_SOMU = "SOMU_EMAILS"
KEY_ADMIN_PAGE = "ALLOW_ADMIN_PAGE_EMAILS"


def _norm_email(email: str | None) -> str:
    return normalize_gas(email).lower()


def _split_lower(raw: str) -> list[str]:
    """カンマ区切り → 前後の空白を除き小文字（空を除く。重なりなし、出てきた順）。"""
    out: list[str] = []
    for part in (raw or "").split(","):
        email = part.strip().lower()
        if email and email not in out:
            out.append(email)
    return out


def settings_map(db: Session) -> dict[str, str]:
    """M_SYSTEM_SETTING の key → value（config_sync が同じキーは下の行を残している）。"""
    return {row.key: row.value for row in db.scalars(select(Setting))}


def setting_emails(db: Session, key: str) -> list[str]:
    """M_SYSTEM_SETTING の 1 つのキーのメール一覧（GAS の normalize_ を通してから分ける）。"""
    row = db.get(Setting, key)
    return _split_lower(normalize_gas(js_string(row.value if row else "")))


def _in_list(db: Session, key: str, email: str | None) -> bool:
    target = _norm_email(email)
    return bool(target) and target in setting_emails(db, key)


def is_admin(db: Session, email: str) -> bool:
    """管理者か（ADMIN_EMAILS だけ）。メールが空なら偽。"""
    return _in_list(db, KEY_ADMIN, email)


def is_somu(db: Session, email: str) -> bool:
    """総務か（SOMU_EMAILS）。メールが空なら偽。"""
    return _in_list(db, KEY_SOMU, email)


def can_access_admin(db: Session, email: str) -> bool:
    """古い管理画面（?page=admin）を開けるか（ALLOW_ADMIN_PAGE_EMAILS）。メールが空なら偽。"""
    return _in_list(db, KEY_ADMIN_PAGE, email)


def is_operator(db: Session, email: str, settings: Settings | None = None) -> bool:
    """運用者か（K-9: SOMU_EMAILS・ADMIN_EMAILS・BOOTSTRAP_ADMINS のどれか）。ops_* と運用の欄だけに使う。"""
    target = _norm_email(email)
    if not target:
        return False
    if target in (settings or get_settings()).bootstrap_admin_list:
        return True
    return is_admin(db, target) or is_somu(db, target)


def _dept_rows(db: Session) -> list[DeptApprover]:
    """データ行（sort_order の順。見出しの行は表に無い）。"""
    return list(db.scalars(select(DeptApprover).order_by(DeptApprover.sort_order)))


def dept_approver_depts(db: Session, email: str) -> list[str]:
    """そのメールが承認者になっている部署名（出てきた順。GAS は同じ部署が 2 行あれば 2 回入れる）。例外は []。"""
    target = _norm_email(email)
    if not target:
        return []
    try:
        rows = _dept_rows(db)
    except Exception:  # noqa: BLE001 - GAS も例外のときは [] を返す
        return []
    depts: list[str] = []
    for row in rows:
        dept = normalize_gas(js_string(row.col_a)).strip()
        raw = normalize_gas(js_string(row.col_b))
        if not dept or not raw:
            continue
        if target in _split_lower(raw):
            depts.append(dept)
    return depts


def is_dept_approver(db: Session, email: str) -> bool:
    """どこかの部署の承認者か（部署管理の入口。管理者は別に判定する）。"""
    return bool(dept_approver_depts(db, email))


def dept_approver_map(db: Session) -> dict[str, list[str]]:
    """部署 → 承認者のメール（小文字、重なりなし、出てきた順）。"""
    result: dict[str, list[str]] = {}
    for row in _dept_rows(db):
        dept = normalize_gas(js_string(row.col_a)).strip()
        raw = normalize_gas(js_string(row.col_b))
        if not dept or not raw:
            continue
        bucket = result.setdefault(dept, [])
        for email in _split_lower(raw):
            if email not in bucket:
                bucket.append(email)
    return result
