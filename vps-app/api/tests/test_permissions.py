"""権限（src/auth.gs の写し）: 管理者・総務・古い管理画面・部署承認者。一覧が空でも全員にしない。
メールは小文字・前後の空白を除いて比べる。"""
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
    db.add(DeptApprover(sheet_row=1, dept="dept", emails="approverEmails", name=" approverName",
                        title="役職", notify_emails="通知メール",
                        cells=["dept", "approverEmails", " approverName", "役職", "通知メール"]))
    for i, (dept, emails, name) in enumerate(rows):
        db.add(DeptApprover(sheet_row=2 + i, dept=dept, emails=emails, name=name, cells=[dept, emails, name]))
    db.commit()


def test_admin_is_admin_emails_and_bootstrap(db, settings):
    add_settings(db, ADMIN_EMAILS=" A@lineworks-local.info , b@lineworks-local.info",
                 SOMU_EMAILS="somu@lineworks-local.info")
    for email in ("a@lineworks-local.info", "b@lineworks-local.info", OWNER, "A@lineworks-local.info",
                  "  b@lineworks-local.info  ", OWNER.upper()):
        assert permissions.is_admin(db, email, settings), email
    assert not permissions.is_admin(db, "somu@lineworks-local.info", settings)  # 総務は管理者ではない
    assert not permissions.is_admin(db, "someone@lineworks-local.info", settings)
    assert not permissions.is_admin(db, "", settings)


def test_somu_and_admin_page_lists(db):
    add_settings(db, SOMU_EMAILS="Somu@lineworks-local.info,  s2@lineworks-local.info",
                 ALLOW_ADMIN_PAGE_EMAILS="page@lineworks-local.info")
    assert permissions.is_somu(db, "somu@lineworks-local.info")
    assert permissions.is_somu(db, " S2@lineworks-local.info ")
    assert not permissions.is_somu(db, "page@lineworks-local.info")
    assert permissions.can_access_admin_page(db, "PAGE@lineworks-local.info")
    assert not permissions.can_access_admin_page(db, "somu@lineworks-local.info")
    assert not permissions.can_access_admin_page(db, OWNER)  # BOOTSTRAP_ADMINS は古い管理画面に足さない


@pytest.mark.parametrize("value", ["", " , ", None])
def test_empty_lists_never_make_everyone_allowed(db, settings, value):
    if value is not None:
        add_settings(db, ADMIN_EMAILS=value, SOMU_EMAILS=value, ALLOW_ADMIN_PAGE_EMAILS=value)
    someone = "someone@lineworks-local.info"
    assert not permissions.is_admin(db, someone, settings)
    assert not permissions.is_somu(db, someone)
    assert not permissions.can_access_admin_page(db, someone)
    assert not permissions.is_somu(db, "")
    assert not permissions.can_access_admin_page(db, "")
    assert permissions.admin_emails(db, settings) == [OWNER]


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
    assert not permissions.is_dept_approver(db, "approveremails")  # 見出しの行は読まない
    assert not permissions.is_dept_approver(db, "")
    assert not permissions.is_dept_approver(db, "someone@lineworks-local.info")


def test_no_dept_approver_rows_means_nobody(db):
    assert permissions.dept_approver_map(db) == {}
    assert not permissions.is_dept_approver(db, "someone@lineworks-local.info")


def test_approver_profiles(db):
    add_dept(db, ("製造1課", "x@lineworks-local.info,y@lineworks-local.info", ""),
             ("総務", "x@lineworks-local.info", "田中太郎"),
             ("組立", "", "空"),
             ("製造1課", "x@lineworks-local.info", "後の名前"))
    assert permissions.approver_profiles(db) == [
        {"email": "x@lineworks-local.info", "name": "田中太郎", "depts": ["製造1課", "総務"]},
        {"email": "y@lineworks-local.info", "name": "y@lineworks-local.info", "depts": ["製造1課"]},
    ]
