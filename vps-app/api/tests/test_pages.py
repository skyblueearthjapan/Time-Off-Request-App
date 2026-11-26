"""画面の入口 `GET /`（app/pages.py）の試験。

ログイン・権限・設定は偽物に差し替え、pages の振り分けとテンプレートの差し込みだけを確かめる。
土台（app.auth・app.permissions・app.config・app.db。担当 F）がまだ無い作業フォルダでも動くよう、
読み込めないモジュールは最小の偽物を入れてから pages を読む。
"""
from __future__ import annotations

import importlib
import json
import re
import sys
import types
from html import unescape
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import parse_qs, quote, urlsplit

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient


def _ensure_module(name: str, **attrs) -> None:
    try:
        importlib.import_module(name)
    except ModuleNotFoundError as exc:
        if exc.name != name:
            raise
        module = types.ModuleType(name)
        for key, value in attrs.items():
            setattr(module, key, value)
        sys.modules[name] = module
        setattr(sys.modules["app"], name.split(".")[-1], module)


def _fake_safe_next(value: str | None) -> str:
    text = value or "/"
    return text if text.startswith("/") and not text.startswith("//") else "/"


importlib.import_module("app")
_ensure_module("app.config", Settings=SimpleNamespace, get_settings=lambda: None)
_ensure_module("app.db", get_db=lambda: None)
_ensure_module("app.auth", get_optional_user=lambda: None, safe_next=_fake_safe_next)
_ensure_module("app.permissions")

from app import auth, pages  # noqa: E402
from app.config import get_settings  # noqa: E402
from app.db import get_db  # noqa: E402

ADMIN = "imaizumi@lineworks-local.info"
SOMU = "somu@lineworks-local.info"
APPROVER = "shonin@lineworks-local.info"
OLD_ADMIN = "kanri@lineworks-local.info"   # ALLOW_ADMIN_PAGE_EMAILS だけにいる人
STAFF = "genba@lineworks-local.info"
PORTAL = "https://script.google.com/a/macros/lineworks-local.info/s/PORTAL_TEST/exec"
ORIGIN = "https://leave.lineworks-local.net"

TEMPLATES = Path(pages.__file__).resolve().parent / "templates" / "pages"
TITLES = {
    "top": "休暇届 申請TOP", "preview": "休暇届 プレビュー", "sign": "休暇届 サイン", "somuAdmin": "休暇届 総務管理",
    "deptAdmin": "休暇届 部署管理", "admin": "休暇届 管理", "leaveStatus": "休暇届 取得状況",
}


@pytest.fixture
def make_client(monkeypatch):
    def _make(email: str | None, portal: str = PORTAL):
        app = FastAPI()
        app.include_router(pages.router)
        user = None if email is None else SimpleNamespace(email=email, name=email.split("@")[0])
        app.dependency_overrides[auth.get_optional_user] = lambda: user
        app.dependency_overrides[get_db] = lambda: None
        app.dependency_overrides[get_settings] = lambda: SimpleNamespace(public_origin=ORIGIN)
        perm = pages.permissions
        monkeypatch.setattr(perm, "is_admin", lambda db, e: e == ADMIN, raising=False)
        monkeypatch.setattr(perm, "is_operator", lambda db, e: e == ADMIN, raising=False)
        monkeypatch.setattr(perm, "is_somu", lambda db, e: e == SOMU, raising=False)
        monkeypatch.setattr(perm, "is_dept_approver", lambda db, e: e == APPROVER, raising=False)
        monkeypatch.setattr(perm, "can_access_admin", lambda db, e: e in (ADMIN, OLD_ADMIN), raising=False)
        monkeypatch.setattr(perm, "settings_map", lambda db: {"PORTAL_URL": portal, "APP_URL": "https://script.google.com/x/exec"},
                            raising=False)
        return TestClient(app, base_url="https://testserver")

    return _make


def title_of(html: str) -> str:
    m = re.search(r"<title>(.*?)</title>", html)
    return m.group(1) if m else ""


def data_of(html: str) -> dict:
    m = re.search(r'<script type="application/json" id="gas-template-data">(.*?)</script>', html, re.S)
    assert m, "テンプレートの値の塊が無い"
    return json.loads(m.group(1))


def page_js(html: str) -> str:
    srcs = re.findall(r'<script src="/static/js/([^"]+)\.js"></script>', html)
    assert srcs[0] == "gas-shim"
    return srcs[1]


def no_auth_msg(html: str) -> str:
    return unescape(re.search(r'<div class="msg">(.*?)</div>', html, re.S).group(1))


# ---- ログイン ----

@pytest.mark.parametrize("path", ["/", "/?page=preview&reqId=LV-20261126-001", "/?page=sign&reqId=LV-20261126-001&x=1",
                                  "/?page=somuAdmin"])
def test_not_logged_in_redirects_to_login_with_next(make_client, path):
    r = make_client(None).get(path, follow_redirects=False)
    assert r.status_code == 302
    loc = urlsplit(r.headers["location"])
    assert loc.path == "/auth/login"
    assert parse_qs(loc.query)["next"] == [path]
    assert r.headers["location"] == "/auth/login?next=" + quote(path, safe="")
    assert r.headers["cache-control"] == "no-store"


# ---- 振り分け ----

@pytest.mark.parametrize("query,expected", [
    ("", "top"), ("?page=", "top"), ("?page=top", "top"), ("?page=unknown", "top"),
    ("?page=preview", "preview"), ("?page=sign", "sign"), ("?page=somuAdmin", "somuAdmin"),
    ("?page=deptAdmin", "deptAdmin"), ("?page=admin", "admin"), ("?page=leaveStatus", "leaveStatus"),
    ("?page=no_auth", "top"), ("?page=../admin", "top"), ("?page=SOMUADMIN", "top"),
    ("?page=admin&page=top", "admin"),  # GAS の e.parameter と同じく最初の値
])
def test_page_routing_for_admin(make_client, query, expected):
    r = make_client(ADMIN).get("/" + query)
    assert r.status_code == 200
    assert page_js(r.text) == expected
    assert data_of(r.text)["PAGE"] == expected
    assert title_of(r.text) == TITLES[expected]
    assert r.headers["cache-control"] == "no-store"


@pytest.mark.parametrize("page", ["top", "preview", "sign", "leaveStatus"])
def test_open_pages_need_no_role(make_client, page):
    r = make_client(STAFF).get("/?page=" + page)
    assert r.status_code == 200
    assert page_js(r.text) == page


@pytest.mark.parametrize("page,allowed,denied", [
    ("somuAdmin", [ADMIN, SOMU], [APPROVER, OLD_ADMIN, STAFF]),
    ("deptAdmin", [ADMIN, APPROVER], [SOMU, OLD_ADMIN, STAFF]),
    ("admin", [ADMIN, OLD_ADMIN], [SOMU, APPROVER, STAFF]),
])
def test_role_pages(make_client, page, allowed, denied):
    for email in allowed:
        r = make_client(email).get("/?page=" + page)
        assert page_js(r.text) == page, email
    title, message = pages.NO_AUTH[page]
    for email in denied:
        r = make_client(email).get("/?page=" + page)
        assert r.status_code == 200
        assert title_of(r.text) == title, email
        assert no_auth_msg(r.text) == message
        assert "gas-shim.js" not in r.text
        assert f'href="{ORIGIN}/?page=top"' in r.text


def test_no_auth_titles_and_messages_are_the_gas_ones():
    assert pages.NO_AUTH == {
        "admin": ("アクセス権限エラー", "このページは管理者・総務のみアクセスできます。"),
        "deptAdmin": ("アクセス権限なし", "このページは部署承認者または管理者のみアクセスできます。"),
        "somuAdmin": ("アクセス権限なし", "このページは総務部・管理者のみアクセスできます。"),
    }


# ---- テンプレートの値 ----

@pytest.mark.parametrize("value,expected", [
    ("LV-20261126-001", "LV-20261126-001"),
    ("LV-20261126-1", ""),
    ("lv-20261126-001", ""),
    ("LV-２０２６1126-001", ""),
    ("LV-20261126-001 ", ""),
    ("LV-20261126-001\n", ""),
    ("", ""),
    ("</script><script>alert(1)</script>", ""),
    ("LV-20261126-001';alert(1)//", ""),
])
def test_req_id_is_validated(make_client, value, expected):
    for page in ("preview", "sign"):
        r = make_client(STAFF).get("/", params={"page": page, "reqId": value})
        assert data_of(r.text)["REQ_ID"] == expected
        assert "alert(1)" not in r.text
    # preview・sign のほかは REQ_ID を渡さない（GAS の doGet と同じ）
    assert "REQ_ID" not in data_of(make_client(STAFF).get("/", params={"page": "top", "reqId": value}).text)


def test_json_block_cannot_close_script():
    text = pages.json_for_script({"x": "</script><script>alert(1)</script>&"})
    assert "<" not in text and ">" not in text and "&" not in text
    assert json.loads(text) == {"x": "</script><script>alert(1)</script>&"}


def test_template_values_top(make_client):
    r = make_client(STAFF).get("/?page=top")
    d = data_of(r.text)
    assert d["APP_URL"] == ORIGIN + "/"          # 自分の公開 URL（Settings の APP_URL は使わない）
    assert d["PORTAL_URL"] == PORTAL
    for p in ("leaveStatus", "deptAdmin", "somuAdmin"):
        assert f'href="{ORIGIN}/?page={p}"' in r.text
    assert f'href="{PORTAL}"' in r.text
    assert "{{" not in r.text and "{%" not in r.text and "<?" not in r.text


def test_portal_url_is_normalized_and_may_be_empty(make_client):
    r = make_client(STAFF, portal="  ｈｔｔｐｓ://portal.example/exec ").get("/?page=top")
    assert data_of(r.text)["PORTAL_URL"] == "https://portal.example/exec"
    r = make_client(STAFF, portal="").get("/?page=top")
    assert data_of(r.text)["PORTAL_URL"] == ""
    # ボタンは今と同じく style="display:none" で出し、JS が PORTAL_URL のあるときだけ見せる
    assert 'id="portalLink" style="display:none;"' in r.text


def test_url_values_are_escaped_in_attributes(make_client):
    r = make_client(STAFF, portal='https://x/"><img src=x onerror=alert(1)>').get("/?page=top")
    assert "<img" not in r.text
    assert data_of(r.text)["PORTAL_URL"] == 'https://x/"><img src=x onerror=alert(1)>'


def test_ops_panel_only_for_admin(make_client):
    assert 'id="opsPanel"' in make_client(ADMIN).get("/?page=somuAdmin").text
    assert 'id="opsPanel"' not in make_client(SOMU).get("/?page=somuAdmin").text


# ---- テンプレートのファイル（port-pages.js が作ったもの） ----

@pytest.mark.parametrize("name", sorted(p.name for p in TEMPLATES.glob("*.html")))
def test_templates_have_no_inline_script_or_on_attributes(name):
    html = (TEMPLATES / name).read_text(encoding="utf-8")
    assert not re.search(r"<[A-Za-z][^>]*\son[a-z]+\s*=", html, re.I)
    for m in re.finditer(r"<script\b([^>]*)>(.*?)</script>", html, re.S):
        attrs, body = m.group(1), m.group(2)
        if 'type="application/json"' in attrs:
            assert body == "{{ data_json|safe }}"
        else:
            assert re.fullmatch(r' src="/static/js/[\w-]+\.js"', attrs), attrs
            assert body == ""
    assert "<?" not in html
    assert '<link rel="stylesheet" href="/static/css/mobile.css">' in html


def test_eight_templates():
    assert sorted(p.stem for p in TEMPLATES.glob("*.html")) == sorted(list(pages.PAGES) + ["no_auth"])
