"""GAS の内部用の入口を呼ぶ。担当 F の土台。担当 M（services/gas_jobs.py）が使う。

本文（仕様書 7-1）: {"internal": true, "token": GAS_INTERNAL_TOKEN, "action": "internal.<名前>", "jobId": "<uuid>",
       "reqId": "LV-...", "expect": {...}, "args": {...}}
返事: いつも HTTP 200 の JSON（成功 {"ok": true, ...}、失敗 {"ok": false, "error": "<コード>", "retryable": ...}）。

このアプリの GAS の Web アプリは access: DOMAIN（ドメインの利用者だけ）。GAS_INTERNAL_AUTH で呼び方を選ぶ:
- bearer（第 1 案・既定）: 最初の POST（script.google.com）にだけ `Authorization: Bearer <委任のアクセストークン>`
  （google_delegate.get_access_token）を付ける。
- anonymous（第 2 案）: 本文の token だけ（内部用のデプロイだけ ANYONE_ANONYMOUS の版）。

GAS の /exec への POST は 302 で script.googleusercontent.com へ送られ、そこを GET で読むと JSON が返る。
転送は自分で追う（httpx の自動追従は使わない）:
- 転送先には Authorization を付けない（委任のトークンを Google の別のホストへ渡さない）。
- 転送先は https の script.googleusercontent.com・script.google.com だけ。それ以外（accounts.google.com の
  ログインなど）は GasAuthError。
- 転送は MAX_REDIRECTS 回まで。
- 返事が JSON でなく HTML（ログイン画面・権限の画面）なら GasAuthError（はっきりした失敗。同じ呼び方で送り直しても通らない）。

例外:
- GasTransportError: 通信の失敗・時間切れ・200 以外・JSON でない返事（「結果が分からない」。担当 M は同じ jobId で
  送り直す。GAS は jobId で同じ返事を返す＝冪等）。
- GasAuthError（GasTransportError の子）: GAS の入口に入れない（ログイン画面・委任のトークンが得られない）。
  送り直しても直らないので、M は GAS のレーンを止めて管理者の欄に出す。
- token・アクセストークンはログにも例外の文にも出さない。
- 呼ぶ前に URL が環境の許可リストに合うか確かめる（staging は本番の GAS を呼ばない）。

呼び方（担当 M）:

    client = GasClient.from_settings(get_settings())
    reply = client.call("internal.afterApprove1", job_id=str(job.id), request_id=req_id,
                        expect={"status": "承認済"}, args={...})
"""
from __future__ import annotations

import json
import logging
import time
from collections.abc import Callable, Mapping
from typing import Any
from urllib.parse import urljoin, urlsplit

import httpx

from app import config as _config
from app.config import Settings

log = logging.getLogger(__name__)

# G の src/Internal.gs と取り決めた action（仕様書 7-2 の閉じたリスト）。足すときは G・M と取り決めてからここに書く。
ACTIONS = frozenset({
    "internal.ping",            # 5 分ごと・起動時（scriptId・紐づくシート・設定・時間帯・移転モード。何も書かない）
    "internal.afterApprove1",   # 1 次承認の後（sign → mail → ledger）
    "internal.afterApprove2",   # 2 次承認の後（pdf →（trash）→ ledger）
    "internal.afterUnapprove",  # 編集で承認が外れた後・取消の後（sign を空に →（trash）→ ledger の行を消す）
    "internal.sendWarnMail",    # 有給警告メール（集計タブは VPS が直前に書く）
})

AUTH_BEARER = "bearer"
AUTH_ANONYMOUS = "anonymous"
AUTH_MODES = (AUTH_BEARER, AUTH_ANONYMOUS)
MAX_REDIRECTS = 5
REDIRECT_STATUSES = frozenset({301, 302, 303, 307, 308})
REDIRECT_HOSTS = frozenset({"script.googleusercontent.com", "script.google.com"})


class GasTransportError(RuntimeError):
    """GAS に届かなかった・返事を読めなかった（結果は分からない）。timeout=True は時間切れ。"""

    def __init__(self, message: str, *, timeout: bool = False, status: int | None = None):
        super().__init__(message)
        self.timeout = timeout
        self.status = status


class GasAuthError(GasTransportError):
    """GAS の入口に入れない（ログイン画面・権限の画面が返った、委任のアクセストークンが得られない、
    ログインへ転送された）。送り直しても直らない。"""


class GasClientConfigError(RuntimeError):
    """URL が環境の許可リストに合わない（staging から本番の GAS を呼ぶなど）。"""


def assert_gas_url_allowed(url: str, settings: Settings) -> str:
    deployment = _config.gas_deployment_id(url)
    if not deployment:
        raise GasClientConfigError("GAS_INTERNAL_URL の形が正しくありません")
    prod_ids = set(_config.PRODUCTION_IDS) | set(getattr(settings, "prod_gas_deployment_ids", ()))
    if settings.is_staging and deployment in prod_ids:
        raise GasClientConfigError("ステージングから本番の GAS は呼べません")
    internal = getattr(settings, "prod_gas_internal_deployment_id", "")
    if not settings.is_staging and internal and deployment != internal:
        raise GasClientConfigError("本番は PROD_GAS_INTERNAL_DEPLOYMENT_ID のデプロイしか呼べません")
    if not settings.is_staging and deployment in _config.KNOWN_STAGING_IDS:
        raise GasClientConfigError("本番からステージングの GAS は呼べません")
    return deployment


def _delegate_token_provider(settings: Settings) -> Callable[[], str]:
    def _provide() -> str:
        from app.google_delegate import get_access_token

        return get_access_token(settings=settings)

    return _provide


def _looks_like_html(resp: httpx.Response) -> bool:
    ctype = resp.headers.get("content-type", "").lower()
    if "text/html" in ctype:
        return True
    head = resp.content[:256].lstrip().lower()
    return head.startswith(b"<!doctype") or head.startswith(b"<html")


class GasClient:
    """http_factory は httpx.Client を作る関数（試験では MockTransport を入れる）。

    auth_mode: "bearer" なら access_token_provider（引数なしでアクセストークンを返す関数）が要る。
    from_settings は google_delegate.get_access_token を使う。
    """

    def __init__(self, url: str, token: str, *, timeout: float = 330.0, settings: Settings | None = None,
                 http_factory: Callable[..., httpx.Client] | None = None, monotonic: Callable[[], float] = time.monotonic,
                 auth_mode: str = AUTH_ANONYMOUS, access_token_provider: Callable[[], str] | None = None):
        if settings is not None:
            assert_gas_url_allowed(url, settings)
        if not token:
            raise GasClientConfigError("GAS_INTERNAL_TOKEN が空です")
        if auth_mode not in AUTH_MODES:
            raise GasClientConfigError(f"GAS_INTERNAL_AUTH が正しくありません: {auth_mode}")
        if auth_mode == AUTH_BEARER and access_token_provider is None:
            raise GasClientConfigError("bearer にはアクセストークンを得る関数が要ります")
        self.url = url
        self._token = token
        self.timeout = float(timeout)
        self.auth_mode = auth_mode
        self._access_token_provider = access_token_provider
        self._http_factory = http_factory or (lambda **kw: httpx.Client(**kw))
        self._monotonic = monotonic

    def __repr__(self) -> str:  # token を出さない
        return f"GasClient(url={self.url!r}, timeout={self.timeout}, auth={self.auth_mode})"

    @classmethod
    def from_settings(cls, settings: Settings, **kwargs: Any) -> "GasClient":
        mode = getattr(settings, "gas_internal_auth", AUTH_BEARER)
        if mode == AUTH_BEARER:
            kwargs.setdefault("access_token_provider", _delegate_token_provider(settings))
        return cls(settings.gas_internal_url, settings.gas_internal_token,
                   timeout=settings.gas_call_timeout_seconds, settings=settings, auth_mode=mode, **kwargs)

    def _first_headers(self, action: str) -> dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if self.auth_mode == AUTH_BEARER:
            try:
                access_token = self._access_token_provider()  # type: ignore[misc]
            except Exception as exc:  # noqa: BLE001 - 委任が無い・更新に失敗した（中身は出さない）
                raise GasAuthError(f"委任のアクセストークンを得られません（{action}）: {type(exc).__name__}") from None
            if not access_token:
                raise GasAuthError(f"委任のアクセストークンが空です（{action}）")
            headers["Authorization"] = f"Bearer {access_token}"
        return headers

    def call(self, action: str, *, job_id: str, request_id: str = "", expect: Mapping[str, Any] | None = None,
             args: Mapping[str, Any] | None = None, timeout: float | None = None) -> dict:
        """1 件呼んで返事の JSON（dict）を返す。通信の失敗・時間切れは GasTransportError、入れないときは GasAuthError。"""
        if action not in ACTIONS:
            raise ValueError(f"知らない action です: {action}")
        if not job_id:
            raise ValueError("job_id が空です（GAS は jobId で二重の処理を防ぐ）")
        body: dict[str, Any] = {"internal": True, "token": self._token, "action": action, "jobId": str(job_id),
                                "reqId": request_id or ""}
        if expect is not None:
            body["expect"] = dict(expect)
        if args is not None:
            body["args"] = dict(args)
        limit = float(timeout if timeout is not None else self.timeout)
        headers = self._first_headers(action)
        payload = json.dumps(body, ensure_ascii=False).encode("utf-8")
        started = self._monotonic()
        try:
            with self._http_factory(timeout=httpx.Timeout(limit, connect=min(30.0, limit)),
                                    follow_redirects=False) as client:
                resp = client.post(self.url, content=payload, headers=headers)
                hops = 0
                while resp.status_code in REDIRECT_STATUSES:
                    hops += 1
                    if hops > MAX_REDIRECTS:
                        raise GasTransportError(f"GAS の転送が {MAX_REDIRECTS} 回を超えました（{action}）",
                                                status=resp.status_code)
                    location = resp.headers.get("location", "")
                    target = urljoin(str(resp.request.url), location)
                    target_parts = urlsplit(target)
                    if not location or target_parts.scheme != "https" or target_parts.hostname not in REDIRECT_HOSTS:
                        raise GasAuthError(f"GAS の入口がほかの場所へ転送しました（{action}、{target_parts.hostname or '不明'}）。"
                                           "ログインを求められています", status=resp.status_code)
                    # 転送先には Authorization を付けない。307・308 は本文ごと送り直す（GAS は 302 を返す）
                    if resp.status_code in (307, 308):
                        resp = client.post(target, content=payload, headers={"Content-Type": "application/json"})
                    else:
                        resp = client.get(target)
        except httpx.TimeoutException:
            raise GasTransportError(f"GAS の呼び出しが時間切れになりました（{action}、{limit:.0f} 秒）",
                                    timeout=True) from None
        except httpx.HTTPError as exc:
            raise GasTransportError(f"GAS に届きませんでした（{action}）: {type(exc).__name__}") from None
        elapsed = self._monotonic() - started
        if resp.status_code in (401, 403):
            raise GasAuthError(f"GAS の入口に入れません（{action}）: {resp.status_code}", status=resp.status_code)
        if resp.status_code != 200:
            raise GasTransportError(f"GAS の返事が 200 ではありません（{action}）: {resp.status_code}",
                                    status=resp.status_code)
        if _looks_like_html(resp):
            raise GasAuthError(f"GAS の返事が JSON ではなく HTML（ログイン画面など）です（{action}）。"
                               "GAS_INTERNAL_AUTH・委任の許可・デプロイの access を確かめてください", status=200)
        try:
            data = resp.json()
        except ValueError:
            raise GasTransportError(f"GAS の返事が JSON ではありません（{action}）") from None
        if not isinstance(data, dict):
            raise GasTransportError(f"GAS の返事が JSON のオブジェクトではありません（{action}）")
        log.info("gas %s job=%s request=%s ok=%s %.1fs", action, job_id, request_id, data.get("ok"), elapsed)
        return data
