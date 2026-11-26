"""権限。GAS の src/auth.gs（isAdmin_・isSomu_・canAccessAdmin_・getDeptApproverDeptsForEmail_・
api_getApproverProfiles）を写す。

- 管理者 = M_SYSTEM_SETTING の ADMIN_EMAILS ＋ BOOTSTRAP_ADMINS。
- 総務 = SOMU_EMAILS。
- 古い管理画面（?page=admin）= ALLOW_ADMIN_PAGE_EMAILS。
- 部署承認者 = M_DEPT_APPROVERS の B 列（approverEmails、カンマ区切り）にメールがある（1 行目は見出し）。
- 一覧が空なら誰も当てはまらない（空で全員にしない。罠 35）。メールの空も偽。
- メールは小文字にし、前後の空白を除いて比べる（GAS: split(',') → trim → toLowerCase）。
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
    return (email or "").strip().lower()


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


def admin_emails(db: Session, settings: Settings | None = None) -> list[str]:
    """ADMIN_EMAILS ＋ BOOTSTRAP_ADMINS（小文字、重なりなし）。"""
    s = settings or get_settings()
    out = setting_emails(db, KEY_ADMIN)
    for email in s.bootstrap_admin_list:
        if email not in out:
            out.append(email)
    return out


def is_admin(db: Session, email: str, settings: Settings | None = None) -> bool:
    """管理者か。メールが空なら偽。"""
    target = _norm_email(email)
    return bool(target) and target in admin_emails(db, settings)


def is_somu(db: Session, email: str) -> bool:
    """総務か（SOMU_EMAILS）。メールが空なら偽。"""
    target = _norm_email(email)
    return bool(target) and target in setting_emails(db, KEY_SOMU)


def can_access_admin_page(db: Session, email: str) -> bool:
    """古い管理画面（?page=admin）を開けるか（ALLOW_ADMIN_PAGE_EMAILS）。メールが空なら偽。"""
    target = _norm_email(email)
    return bool(target) and target in setting_emails(db, KEY_ADMIN_PAGE)


def _dept_rows(db: Session) -> list[DeptApprover]:
    """データ行（2 行目から）。"""
    return list(db.scalars(select(DeptApprover).where(DeptApprover.sheet_row >= 2).order_by(DeptApprover.sheet_row)))


def dept_approver_depts(db: Session, email: str) -> list[str]:
    """そのメールが承認者になっている部署名（出てきた順。GAS は同じ部署が 2 行あれば 2 回入れる）。"""
    target = _norm_email(normalize_gas(email))
    if not target:
        return []
    depts: list[str] = []
    for row in _dept_rows(db):
        dept = normalize_gas(js_string(row.dept))
        raw = normalize_gas(js_string(row.emails))
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
        dept = normalize_gas(js_string(row.dept))
        raw = normalize_gas(js_string(row.emails))
        if not dept or not raw:
            continue
        bucket = result.setdefault(dept, [])
        for email in _split_lower(raw):
            if email not in bucket:
                bucket.append(email)
    return result


def approver_profiles(db: Session) -> list[dict]:
    """M_DEPT_APPROVERS の承認者一覧 [{email, name, depts}]（api_getApproverProfiles の profileMap と同じ作り。
    名前での並べ替え（localeCompare 'ja'）と、管理者以外は自分だけに絞るのは業務の側＝担当 W）。"""
    profiles: dict[str, dict] = {}
    for row in _dept_rows(db):
        dept = normalize_gas(js_string(row.dept))
        raw = normalize_gas(js_string(row.emails))
        name = normalize_gas(js_string(row.name))
        if not dept or not raw:
            continue
        for email in _split_lower(raw):
            profile = profiles.setdefault(email, {"email": email, "name": name or email, "depts": []})
            if dept not in profile["depts"]:
                profile["depts"].append(dept)
            # 名前が空だった場合、後の行で名前が見つかれば上書き（GAS と同じ）
            if name and profile["name"] == email:
                profile["name"] = name
    return list(profiles.values())
