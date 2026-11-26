"""画面から呼ぶ関数（RPC）の登録と呼び出し。`POST /api/rpc/{fn}`、本文は引数の JSON 配列。

入口で確かめること: Origin が PUBLIC_ORIGIN（無ければ Sec-Fetch-Site: same-origin）、Content-Type: application/json、
fn が閉じたリスト（起動時に業務の側が register で登録した名前だけ）にある、本文が配列。

応答（gas-shim.js の取り決め）:
- 成功        200 {"ok": true, "result": R}（R が null・{error}・{ok:false,…} でもそのまま）
- 業務エラー   400 {"ok": false, "error": "GAS と同じ文"}（RpcError を raise する）
- 未ログイン   401 {"ok": false, "error": "ログインが必要です。", "login": "/auth/login"}（main.py）
- 権限が無い   403 {"ok": false, "error": <RpcForbidden の文>}
- 別サイト     403、JSON でない 415、閉じたリストに無い関数 404、名前だけ決めて中身がまだの関数 501
- 冪等キーが処理中 409、想定外のエラー 500（詳細はログだけ）

閉じたリストへの登録（担当 W）: `app/domain/register.py` に `register_all(register)` を置く。起動時に
rpc.load_domain_handlers がそれを 1 回呼ぶ。登録していない名前は 404 で、外から呼べる関数はここに並べたものだけ。

    # app/domain/register.py
    from app.rpc import RpcContext, RpcError

    def api_submitLeaveRequest(ctx: RpcContext, payload):
        ...                       # ctx.db・ctx.user.email（小文字）・ctx.now（UTC。1 回の呼び出しで同じ）・ctx.settings
        raise RpcError("休暇日を入力してください")

    def register_all(register):
        register("api_getDeptList", api_getDeptList)                                  # 読むだけ
        register("api_submitLeaveRequest", api_submitLeaveRequest, write=True)        # 書く（冪等キー・監査）
        register("api_approveLeaveRequest", api_approveLeaveRequest, write=True, request_id_arg="first")
        register("api_approveLeaveRequest2", api_approveLeaveRequest2, write=True, request_id_arg="first",
                 role="somu_or_admin")                                               # 総務か管理者だけ（403）
        register("api_approveLeaveRequest2Batch", fn, write=True, request_id_arg="first_list")
        register("api_x")         # 名前だけ先に決める（中身ができるまで 501）

register(name, handler=None, *, write=False, role=None, request_id_arg=None, forbidden_message=""):
- write=True: 書く関数。X-Idempotency-Key を受け（同じ人の同じキーの 2 回目は 1 回目の返事を返す）、
  audit_log を同じトランザクションで残す。
- role: rpc が呼ぶ前に確かめる権限（ROLE_CHECKS）。"admin"（管理者）・"somu_or_admin"（総務か管理者）・
  "dept_approver_or_admin"（部署承認者か管理者）・"admin_page"（ALLOW_ADMIN_PAGE_EMAILS）。None は誰でも（ログインだけ）。
  403 の文は ROLE_MESSAGES（GAS の文）。関数ごとに違う文は forbidden_message で渡す。
  GAS の確かめの順を保ちたい関数は role を付けず、業務関数の中で ctx.require_admin() などを呼ぶ（RpcForbidden → 403）。
- request_id_arg: 監査に残す申請 ID の場所。"first"（第 1 引数が REQ_ID）・"first_list"（第 1 引数が REQ_ID の配列）。
  返事が dict で "reqId" を持てばそれも残す。
- 引数は GAS と同じ並びの位置引数。省略できる引数には既定値を持たせる。余った引数は捨てる（JS と同じ）。
- 業務関数は commit しない（rpc が冪等キー・監査と一緒に確定する）。
"""
from __future__ import annotations

import importlib
import inspect
import json
import logging
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any

from fastapi import APIRouter, Depends, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import Response
from sqlalchemy import delete
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from app import audit, permissions, timeutil
from app.auth import CurrentUser, require_user
from app.config import Settings, get_settings
from app.db import get_db
from app.models import IdempotencyKey
from app.serialize import json_default

log = logging.getLogger(__name__)

REQUEST_ID_ARGS = ("first", "first_list")


@dataclass(frozen=True)
class FunctionSpec:
    """閉じたリストの 1 行（register が作る）。"""

    name: str
    write: bool = False
    role: str | None = None
    request_id_arg: str | None = None
    forbidden_message: str = ""


IDEMPOTENCY_HEADER = "X-Idempotency-Key"
IDEMPOTENCY_TTL = timedelta(hours=24)
_IDEMPOTENCY_KEY_RE = re.compile(r"^[A-Za-z0-9._:-]{8,128}\Z")

ERR_SHAPE = "呼び出しの形が正しくありません。"
ERR_FORBIDDEN_ORIGIN = "この画面からは操作できません。"
ERR_ARGS = "引数の数が正しくありません。"
ERR_INTERNAL = "サーバーでエラーが起きました。時間をおいてもう一度お試しください。"
ERR_IN_PROGRESS = "同じ操作を処理しています。少し待ってから画面を読み込み直してください。"
ERR_KEY_REUSED = "同じ操作キーが別の操作に使われています。画面を読み込み直してください。"
ERR_UNKNOWN_FN = "その操作はありません。"
ERR_NOT_ADMIN = "権限がありません"
# role ごとの 403 の文（GAS の throw の文。関数ごとに違えば register の forbidden_message で変える）
ROLE_MESSAGES: dict[str, str] = {
    "admin": "権限がありません",
    "somu_or_admin": "権限がありません（総務または管理者のみ）",
    "dept_approver_or_admin": "部署承認者の権限がありません。",
    "admin_page": "権限がありません",
}


class RpcError(Exception):
    """業務エラー（400）。message は画面にそのまま出る（GAS の throw new Error('...') と同じ文にする）。"""

    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


class RpcForbidden(RpcError):
    """権限の無い人が呼んだ（403）。"""

    def __init__(self, message: str = ERR_NOT_ADMIN):
        super().__init__(message)


class RpcNotImplemented(Exception):
    pass


@dataclass
class RpcContext:
    """業務関数に渡す。

    - db: この呼び出しのセッション（commit しない）
    - user: CurrentUser(email, name)。email は Google ログインのメール（小文字・空にならない）
    - now: この呼び出しの時刻（UTC。1 回の呼び出しの中で同じ値を使う）
    - settings: app.config.Settings
    - fn: 呼ばれた関数名
    """

    db: Session
    user: CurrentUser
    settings: Settings
    now: datetime = field(default_factory=lambda: timeutil.now_utc())
    request: Request | None = None
    fn: str = ""
    _is_admin: bool | None = field(default=None, repr=False)

    @property
    def is_admin(self) -> bool:
        if self._is_admin is None:
            self._is_admin = permissions.is_admin(self.db, self.user.email, self.settings)
        return self._is_admin

    def require_admin(self) -> str:
        """管理者でなければ RpcForbidden（403）。メールを返す。"""
        if not self.is_admin:
            raise RpcForbidden()
        return self.user.email

    def has_role(self, role: str) -> bool:
        """ROLES のどれかに当てはまるか。"""
        return ROLE_CHECKS[role](self)


Handler = Callable[..., Any]

ROLE_CHECKS: dict[str, Callable[[RpcContext], bool]] = {
    "admin": lambda ctx: ctx.is_admin,
    "somu_or_admin": lambda ctx: ctx.is_admin or permissions.is_somu(ctx.db, ctx.user.email),
    "dept_approver_or_admin": lambda ctx: ctx.is_admin or permissions.is_dept_approver(ctx.db, ctx.user.email),
    "admin_page": lambda ctx: permissions.can_access_admin_page(ctx.db, ctx.user.email),
}


def _not_implemented(ctx: RpcContext, *args: Any) -> Any:
    raise RpcNotImplemented()


_specs: dict[str, FunctionSpec] = {}
_registry: dict[str, Handler] = {}


def reset_registry() -> None:
    """閉じたリストを空に戻す（試験用）。"""
    _specs.clear()
    _registry.clear()


def register(name: str, handler: Handler | None = None, *, write: bool = False, role: str | None = None,
             request_id_arg: str | None = None, forbidden_message: str = "") -> Handler | None:
    """閉じたリストに 1 つ足す。handler が無ければ名前だけ（501）。同じ名前の二重登録は ValueError。"""
    if not name or not name.replace("_", "").isalnum() or not name[0].isalpha():
        raise ValueError(f"公開できない関数名です: {name!r}")
    if name in _specs:
        raise ValueError(f"二重に登録しています: {name}")
    if role is not None and role not in ROLE_CHECKS:
        raise ValueError(f"知らない権限です: {role}")
    if request_id_arg is not None and request_id_arg not in REQUEST_ID_ARGS:
        raise ValueError(f"request_id_arg は {REQUEST_ID_ARGS} のどれか: {request_id_arg}")
    _specs[name] = FunctionSpec(name=name, write=bool(write), role=role, request_id_arg=request_id_arg,
                                forbidden_message=forbidden_message)
    _registry[name] = handler if handler is not None else _not_implemented
    return handler


def get_spec(name: str) -> FunctionSpec | None:
    return _specs.get(name)


def get_handler(name: str) -> Handler | None:
    return _registry.get(name)


def registered_functions() -> list[str]:
    """閉じたリスト（登録の順）。"""
    return list(_specs)


def is_implemented(name: str) -> bool:
    return _registry.get(name, _not_implemented) is not _not_implemented


def implemented_functions() -> list[str]:
    return [name for name in _specs if is_implemented(name)]


def load_domain_handlers() -> list[str]:
    """担当 W の app.domain.register.register_all(register) を呼ぶ（無ければ何もしない＝閉じたリストは空）。"""
    try:
        domain_register = importlib.import_module("app.domain.register")
    except ModuleNotFoundError as exc:
        if exc.name in ("app.domain", "app.domain.register"):
            log.warning("rpc: app.domain.register がまだありません（呼べる関数はありません）")
            return []
        raise
    domain_register.register_all(register)
    return implemented_functions()


# ===== JSON =====


def dumps(body: Any) -> str:
    return json.dumps(body, ensure_ascii=False, default=json_default, separators=(",", ":"))


def json_response(body: dict[str, Any] | str, status_code: int = 200) -> Response:
    content = body if isinstance(body, str) else dumps(body)
    return Response(
        content=content,
        status_code=status_code,
        media_type="application/json; charset=utf-8",
        headers={"Cache-Control": "no-store"},
    )


def error_response(message: str, status_code: int) -> Response:
    return json_response({"ok": False, "error": message}, status_code)


# ===== 呼び出し =====


def _call(ctx: RpcContext, handler: Handler, args: list[Any]) -> Any:
    sig = inspect.signature(handler)
    params = list(sig.parameters.values())
    if not any(p.kind == p.VAR_POSITIONAL for p in params):
        # GAS（JS）は余った引数を黙って捨てるので、同じく切り詰める（正しさ 低-7）
        positional = [p for p in params if p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)]
        args = list(args)[:max(0, len(positional) - 1)]
    try:
        bound = sig.bind(ctx, *args)
    except TypeError as exc:
        raise RpcError(ERR_ARGS) from exc
    return handler(*bound.args, **bound.kwargs)


def _request_ids(spec: FunctionSpec, args: list[Any], result: Any = None) -> list[str]:
    ids: list[str] = []
    if args:
        first = args[0]
        if spec.request_id_arg == "first" and isinstance(first, (str, int)):
            ids.append(str(first))
        elif spec.request_id_arg == "first_list" and isinstance(first, list):
            ids.extend(str(x) for x in first if isinstance(x, (str, int)))
    if isinstance(result, dict) and isinstance(result.get("reqId"), str) and result["reqId"] not in ids:
        ids.append(result["reqId"])
    return ids


def _replay(db: Session, user_email: str, key: str, fn: str) -> Response | None:
    row = db.get(IdempotencyKey, (user_email, key))
    if row is None:
        return None
    if row.fn != fn:
        return error_response(ERR_KEY_REUSED, 400)
    if row.state != "done":
        return error_response(ERR_IN_PROGRESS, 409)
    return json_response(row.response_json)


def purge_idempotency_keys(db: Session, now: datetime | None = None) -> int:
    """24 時間より古い冪等キーを消す（worker の purge）。消した件数を返す。commit は呼び出し側。"""
    cutoff = (now or timeutil.now_utc()) - IDEMPOTENCY_TTL
    return int(db.execute(delete(IdempotencyKey).where(IdempotencyKey.created_at < cutoff)).rowcount or 0)


def _audit_failure(db: Session, user: CurrentUser, spec: FunctionSpec, args: list[Any], message: str) -> None:
    """失敗した書き込みの記録（業務の書き込みを戻した後、別のトランザクションで）。"""
    fn = spec.name
    try:
        audit.record(db, user.email, fn, _request_ids(spec, args), ok=False, error=message)
        db.commit()
    except Exception:  # noqa: BLE001 - 記録の失敗で応答を変えない
        db.rollback()
        log.exception("rpc %s: could not write audit_log", fn)


def dispatch(fn: str, args: list[Any], *, db: Session, user: CurrentUser, settings: Settings,
             request: Request | None = None, idempotency_key: str | None = None) -> Response:
    handler = _registry.get(fn)
    spec = _specs.get(fn)
    if handler is None or spec is None:
        return error_response(ERR_UNKNOWN_FN, 404)
    ctx = RpcContext(db=db, user=user, settings=settings, now=timeutil.now_utc(), request=request, fn=fn)
    writes = spec.write
    use_key = bool(idempotency_key) and writes
    key_row: IdempotencyKey | None = None

    try:
        if spec.role is not None and not ctx.has_role(spec.role):
            db.rollback()
            return error_response(spec.forbidden_message or ROLE_MESSAGES[spec.role], 403)
        if use_key:
            replay = _replay(db, user.email, idempotency_key, fn)  # type: ignore[arg-type]
            if replay is not None:
                db.rollback()
                return replay
            key_row = IdempotencyKey(user_email=user.email, key=idempotency_key, fn=fn, state="running",
                                     created_at=ctx.now)
            db.add(key_row)
            try:
                # PostgreSQL では、同じキーを処理中の別の呼び出しがあれば、その確定まで待ってから重複になる
                db.flush()
            except IntegrityError:
                db.rollback()
                replay = _replay(db, user.email, idempotency_key, fn)  # type: ignore[arg-type]
                db.rollback()
                return replay or error_response(ERR_IN_PROGRESS, 409)

        result = _call(ctx, handler, args)
        body = dumps({"ok": True, "result": result})
        if key_row is not None:
            key_row.response_json = body
            key_row.state = "done"
        if writes:
            audit.record(db, user.email, fn, _request_ids(spec, args, result), ok=True)
        db.commit()
        return json_response(body)
    except RpcForbidden as exc:
        db.rollback()
        return error_response(exc.message, 403)
    except RpcError as exc:
        db.rollback()
        if writes:
            _audit_failure(db, user, spec, args, exc.message)
        return error_response(exc.message, 400)
    except RpcNotImplemented:
        db.rollback()
        return error_response(f"この機能はまだ用意されていません。（{fn}）", 501)
    except Exception as exc:  # noqa: BLE001
        log.exception("rpc %s failed", fn)
        db.rollback()
        if writes:
            _audit_failure(db, user, spec, args, f"{type(exc).__name__}")
        return error_response(ERR_INTERNAL, 500)


router = APIRouter()


def check_same_origin(request: Request, settings: Settings) -> bool:
    """Origin があれば PUBLIC_ORIGIN と照合し、無ければブラウザの Sec-Fetch-Site: same-origin を求める。"""
    origin = request.headers.get("origin")
    if origin is not None:
        return origin.rstrip("/") == settings.public_origin
    return request.headers.get("sec-fetch-site") == "same-origin"


@router.post("/api/rpc/{fn}")
async def call_rpc(
    fn: str,
    request: Request,
    db: Session = Depends(get_db),
    settings: Settings = Depends(get_settings),
    user: CurrentUser = Depends(require_user),
) -> Response:
    # 入口の確認: 同じサイトの画面からの JSON の呼び出しだけを受ける（ほかのサイトからの書き込みを止める）
    if not check_same_origin(request, settings):
        return error_response(ERR_FORBIDDEN_ORIGIN, 403)
    content_type = request.headers.get("content-type", "").split(";", 1)[0].strip().lower()
    if content_type != "application/json":
        return error_response(ERR_SHAPE, 415)
    if fn not in _specs:
        return error_response(ERR_UNKNOWN_FN, 404)

    raw = await request.body()
    if raw.strip():
        try:
            args = json.loads(raw)
        except ValueError:
            return error_response(ERR_SHAPE, 400)
    else:
        args = []
    if not isinstance(args, list):
        return error_response(ERR_SHAPE, 400)

    key = request.headers.get(IDEMPOTENCY_HEADER)
    if key is not None and not _IDEMPOTENCY_KEY_RE.match(key):
        return error_response(ERR_SHAPE, 400)

    return await run_in_threadpool(
        dispatch, fn, args, db=db, user=user, settings=settings, request=request, idempotency_key=key
    )
