"""権限（src/auth.gs の写し。仕様書 8-9・06 の K-9）: 管理者・総務・古い管理画面・部署承認者・運用者。
一覧が空でも全員にしない。メールは小文字・前後の空白を除いて比べる。"""
from __future__ import annotations

import pytest

from app import permissions
from app.models import DeptApprover, Setting

OWNER = "owner@lineworks-local.info"  # conftest の BOOTSTRAP_ADMINS


def add_settings(db, **kv):
    for i, (k, v) in enumerate(kv.items()):
        db.add(Setting(key=k, value=v, sheet_row=2 + i))
    db.commit()


def add_dept(db, *rows):
    for i, (dept, emails, name) in enumerate(rows):
        db.add(DeptApprover(sort_order=1 + i, col_a=dept, col_b=emails, col_c=name))
    db.commit()


def test_admin_is_admin_emails_only(db):
    add_settings(db, ADMIN_EMAILS=" A@lineworks-local.info , b@lineworks-local.info",
                 SOMU_EMAILS="somu@lineworks-local.info")
    for email in ("a@lineworks-local.info", "b@lineworks-local.info", "A@lineworks-local.info",
                  "  b@lineworks-local.info  "):
        assert permissions.is_admin(db, email), email
    assert not permissions.is_admin(db, OWNER)  # BOOTSTRAP_ADMINS は is_admin に入らない（画面を変えない）
    assert not permissions.is_admin(db, "somu@lineworks-local.info")  # 総務は管理者ではない
    assert not permissions.is_admin(db, "someone@lineworks-local.info")
    assert not permissions.is_admin(db, "")


def test_operator_is_somu_admin_or_bootstrap(db, settings):
    """K-9: 運用者は SOMU_EMAILS・ADMIN_EMAILS・BOOTSTRAP_ADMINS のどれか。"""
    add_settings(db, ADMIN_EMAILS="a@lineworks-local.info", SOMU_EMAILS="Somu@lineworks-local.info",
                 ALLOW_ADMIN_PAGE_EMAILS="page@lineworks-local.info")
    for email in (OWNER, OWNER.upper(), "a@lineworks-local.info", " somu@lineworks-local.info "):
        assert permissions.is_operator(db, email, settings), email
    assert not permissions.is_operator(db, "page@lineworks-local.info", settings)  # 古い管理画面だけの人は違う
    assert not permissions.is_operator(db, "someone@lineworks-local.info", settings)
    assert not permissions.is_operator(db, "", settings)


def test_operator_with_empty_lists_is_only_bootstrap(db, settings):
    add_settings(db, ADMIN_EMAILS="", SOMU_EMAILS=" , ")
    assert permissions.is_operator(db, OWNER, settings)
    assert not permissions.is_operator(db, "someone@lineworks-local.info", settings)


def test_somu_and_admin_page_lists(db):
    add_settings(db, SOMU_EMAILS="Somu@lineworks-local.info,  s2@lineworks-local.info",
                 ALLOW_ADMIN_PAGE_EMAILS="page@lineworks-local.info")
    assert permissions.is_somu(db, "somu@lineworks-local.info")
    assert permissions.is_somu(db, " S2@lineworks-local.info ")
    assert not permissions.is_somu(db, "page@lineworks-local.info")
    assert permissions.can_access_admin(db, "PAGE@lineworks-local.info")
    assert not permissions.can_access_admin(db, "somu@lineworks-local.info")
    assert not permissions.can_access_admin(db, OWNER)  # BOOTSTRAP_ADMINS は古い管理画面に足さない


@pytest.mark.parametrize("value", ["", " , ", None])
def test_empty_lists_never_make_everyone_allowed(db, value):
    if value is not None:
        add_settings(db, ADMIN_EMAILS=value, SOMU_EMAILS=value, ALLOW_ADMIN_PAGE_EMAILS=value)
    someone = "someone@lineworks-local.info"
    for check in (permissions.is_admin, permissions.is_somu, permissions.can_access_admin):
        assert not check(db, someone)
        assert not check(db, "")
    assert permissions.setting_emails(db, "ADMIN_EMAILS") == []


def test_full_width_letters_are_normalized_like_gas(db):
    add_settings(db, SOMU_EMAILS="ｓｏｍｕ@lineworks-local.info")
    assert permissions.is_somu(db, "somu@lineworks-local.info")


def test_dept_approvers(db):
    add_dept(db, ("製造１課", "X@lineworks-local.info, y@lineworks-local.info", "田中"),
             ("製造1課", "z@lineworks-local.info", ""),
             ("", "ignored@lineworks-local.info", ""),
             ("総務", "x@lineworks-local.info", "田中"))
    assert permissions.dept_approver_map(db) == {
        "製造1課": ["x@lineworks-local.info", "y@lineworks-local.info", "z@lineworks-local.info"],
        "総務": ["x@lineworks-local.info"],
    }
    assert permissions.dept_approver_depts(db, " X@lineworks-local.info ") == ["製造1課", "総務"]
    assert permissions.is_dept_approver(db, "y@lineworks-local.info")
    assert not permissions.is_dept_approver(db, "ignored@lineworks-local.info")  # 部署が空の行は数えない
    assert not permissions.is_dept_approver(db, "")
    assert not permissions.is_dept_approver(db, "someone@lineworks-local.info")


def test_no_dept_approver_rows_means_nobody(db):
    assert permissions.dept_approver_map(db) == {}
    assert not permissions.is_dept_approver(db, "someone@lineworks-local.info")
