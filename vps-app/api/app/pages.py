"""画面の入口 `GET /`（GAS の doGet の代わり）。土台（担当 F）の仮のページで、担当 U がこのファイルごと差し替える。

U が作るもの: `?page=` の振り分け（top・preview・sign・admin・deptAdmin・somuAdmin・leaveStatus、不明は top）、
入口の判定（somuAdmin＝総務か管理者、deptAdmin＝管理者か部署承認者、admin＝ALLOW_ADMIN_PAGE_EMAILS。
permissions.is_somu・is_admin・is_dept_approver・can_access_admin_page を使う）、題名、テンプレート変数
（APP_URL は settings.app_url、PORTAL_URL は M_SYSTEM_SETTING の PORTAL_URL＝permissions.settings_map(db)）。
差し替えても、`router`（APIRouter）を持つこと・全ページでログインが要ること（未ログインは
`/auth/login?next=<今のパスとクエリ>` へ 302）は保つ（main.py が router を取り込む）。
"""
from __future__ import annotations

import html
from urllib.parse import quote

from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response

from app import auth

router = APIRouter()

_NO_STORE = {"Cache-Control": "no-store"}


def login_redirect(request: Request) -> RedirectResponse:
    """未ログインのとき、今のパスとクエリを next にしてログインへ。"""
    target = request.url.path + (f"?{request.url.query}" if request.url.query else "")
    return RedirectResponse(f"{auth.LOGIN_URL}?next={quote(target, safe='')}", status_code=302, headers=_NO_STORE)


@router.get("/")
def index(request: Request, user: auth.CurrentUser | None = Depends(auth.get_optional_user)) -> Response:
    if user is None:
        return login_redirect(request)
    body = (
        "<!doctype html><html lang=\"ja\"><head><meta charset=\"utf-8\">"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">"
        "<title>休暇届</title></head><body>"
        f"<p>休暇届（準備中）。ログイン中: {html.escape(user.email)}</p>"
        "</body></html>"
    )
    return HTMLResponse(body, headers=_NO_STORE)
