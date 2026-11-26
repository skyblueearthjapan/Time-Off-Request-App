"""画面の入口 `GET /`（GAS の doGet の代わり。`src/menu.gs:3-89` と同じ振り分け）。

- `?page=` で 7 ページを振り分ける。無指定・空・一覧に無い値は top（`menu.gs:4, 81`）。
- 全ページでログインが要る。未ログインは `/auth/login?next=<今のパスとクエリ>` へ（`page`・`reqId` を保つ）。
- 画面の決まりは docs/vps-implementation-spec.md の 3 章。
- somuAdmin は総務か管理者、deptAdmin は管理者か部署承認者、admin は古い管理画面の一覧（`canAccessAdmin_`）。
  だめなら権限エラーの画面（題名と文は GAS と同じ 2 種類）。
- テンプレートは GAS の HTML をそのまま写したもの（`vps-app/web/dev/port-pages.js` が作る）。
  差し込む値は `{{ APP_URL }}`・`{{ PORTAL_URL }}`・`{{ message }}`・`{{ title }}`（自動エスケープ）と、
  JS へ渡す値の JSON の塊 `{{ data_json|safe }}`（`<`・`>`・`&` を \\u 形式にしてあるので `</script>` を閉じられない）。
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any
from urllib.parse import quote

import jinja2
from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from sqlalchemy.orm import Session

from app import auth, permissions
from app.config import Settings, get_settings
from app.db import get_db

router = APIRouter()

TEMPLATES_DIR = Path(__file__).resolve().parent / "templates" / "pages"
_env = jinja2.Environment(
    loader=jinja2.FileSystemLoader(str(TEMPLATES_DIR)),
    autoescape=True,
    keep_trailing_newline=True,
)

PAGES = ("top", "preview", "sign", "somuAdmin", "deptAdmin", "admin", "leaveStatus")

# 権限エラーの題名と文（GAS の setTitle と t.message。admin だけ題名が違う）
NO_AUTH = {
    "admin": ("アクセス権限エラー", "このページは管理者・総務のみアクセスできます。"),
    "deptAdmin": ("アクセス権限なし", "このページは部署承認者または管理者のみアクセスできます。"),
    "somuAdmin": ("アクセス権限なし", "このページは総務部・管理者のみアクセスできます。"),
}

# 連番は 3 桁以上（その日に 1000 件を超えれば 4 桁になる）。JS の \d・$ と同じく半角の数字だけ・末尾の改行なし
_REQ_ID_RE = re.compile(r"^LV-[0-9]{8}-[0-9]{3,}\Z")
_FULLWIDTH_ALNUM = re.compile(r"[Ａ-Ｚａ-ｚ０-９]")
_NO_STORE = {"Cache-Control": "no-store"}


def validate_req_id(value: str | None) -> str:
    """`?reqId=` は `LV-yyyyMMdd-NNN`（連番 3 桁以上）の形のときだけその文字。ほかは ''。

    仕様 3-2 は「形を確かめない」だが、sign は REQ_ID を承認後の innerHTML のリンクに入れるので、
    形の違う値は入れない（本番の 279 件はすべてこの形）。形の違う値は「REQ_IDが指定されていません。」になる。
    """
    text = value or ""
    return text if _REQ_ID_RE.match(text) else ""


def normalize(value: Any) -> str:
    """`src/config.gs` の normalize_: 前後の空白を除き、全角英数字を半角に。"""
    text = "" if value is None else str(value)
    return _FULLWIDTH_ALNUM.sub(lambda m: chr(ord(m.group(0)) - 0xFEE0), text.strip())


def json_for_script(data: Any) -> str:
    """`<script type="application/json">` に入れてよい JSON。"""
    text = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return text.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")


def _first(request: Request, name: str) -> str | None:
    """GAS の e.parameter と同じく、同じ名前が複数あれば最初の値。"""
    values = request.query_params.getlist(name)
    return values[0] if values else None


def _render(template: str, **context: Any) -> HTMLResponse:
    html = _env.get_template(template).render(**context)
    return HTMLResponse(html, headers=_NO_STORE)


def _app_url(settings: Settings) -> str:
    """自分の公開 URL（M_SYSTEM_SETTING の APP_URL は GAS の /exec なので使わない）。"""
    return settings.public_origin.rstrip("/") + "/"


def _portal_url(db: Session) -> str:
    """Settings（DB に取り込んだ M_SYSTEM_SETTING）の PORTAL_URL。空ならボタンを出さない（画面の JS が決める）。

    https:// で始まる値だけ使う（javascript: などをリンクにしない）。違えば空＝ボタンを出さない。
    """
    url = normalize(permissions.settings_map(db).get("PORTAL_URL", ""))
    return url if url.startswith("https://") else ""


def _allowed(page: str, db: Session, email: str) -> bool:
    if page == "admin":
        return permissions.can_access_admin(db, email)
    if page == "deptAdmin":
        return permissions.is_admin(db, email) or permissions.is_dept_approver(db, email)
    if page == "somuAdmin":
        return permissions.is_somu(db, email) or permissions.is_admin(db, email)
    return True


@router.get("/", include_in_schema=False)
def page_entry(
    request: Request,
    user: Any = Depends(auth.get_optional_user),
    db: Session = Depends(get_db),
    settings: Settings = Depends(get_settings),
) -> Response:
    if user is None:
        here = request.url.path + (("?" + request.url.query) if request.url.query else "")
        target = "/auth/login?next=" + quote(auth.safe_next(here), safe="")
        return RedirectResponse(target, status_code=302, headers=_NO_STORE)

    page = _first(request, "page") or "top"
    safe_page = page if page in PAGES else "top"
    email = user.email
    app_url = _app_url(settings)

    if not _allowed(safe_page, db, email):
        title, message = NO_AUTH[safe_page]
        return _render("no_auth.html", title=title, message=message, APP_URL=app_url)

    portal_url = _portal_url(db)
    data: dict[str, Any] = {"PAGE": safe_page, "APP_URL": app_url, "PORTAL_URL": portal_url}
    if safe_page in ("preview", "sign"):
        data["REQ_ID"] = validate_req_id(_first(request, "reqId"))
    context: dict[str, Any] = {"APP_URL": app_url, "PORTAL_URL": portal_url, "data_json": json_for_script(data)}
    if safe_page == "somuAdmin":
        # 一番下の「運用（管理者）」欄は運用者のときだけ出す（仕様 3-11）
        context["SHOW_OPS"] = permissions.is_operator(db, email)
    return _render(safe_page + ".html", **context)
