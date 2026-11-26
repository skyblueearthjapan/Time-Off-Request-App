"""土台の仮の画面 `GET /`（担当 U が app/pages.py を差し替える）: ログイン必須・next を付けたログインへの 302。
U が差し替えた後も、この 2 つの決まりは保つ。"""
from __future__ import annotations

from urllib.parse import parse_qs, urlparse


def test_root_requires_login_and_keeps_query_as_next(client):
    resp = client.get("/?page=preview&reqId=LV-20261126-001")
    assert resp.status_code == 302
    loc = urlparse(resp.headers["location"])
    assert loc.path == "/auth/login"
    assert parse_qs(loc.query)["next"] == ["/?page=preview&reqId=LV-20261126-001"]


def test_root_after_login_is_200_html(client, login_as):
    login_as("someone@lineworks-local.info")
    resp = client.get("/")
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("text/html")
    assert resp.headers["cache-control"] == "no-store"


def test_foundation_routes_answer(client):
    assert client.get("/healthz").status_code == 200
    assert client.get("/auth/me").status_code == 401
    assert client.get("/auth/login").status_code == 302
    assert client.get("/auth/google-delegate/start").status_code == 302  # 未ログインはログインへ
    assert client.post("/api/rpc/api_getDeptList", json=[]).status_code == 401
    assert client.post("/api/sebastian", content=b"{}").status_code in (404, 405)  # このアプリには無い
