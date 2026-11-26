"""スパイク S-1（実装仕様 14 章）: GAS の内部用の入口をどの呼び方で通れるかを確かめる道具（担当 M）。

ステージングの api コンテナの中で動かす（秘密は出さない。token・アクセストークンは表示もログもしない）:

    docker compose exec -T api env PYTHONPATH=/app python - --mode bearer < tools/spike_gas_auth.py
    （コンテナに tools/ を置いたなら: docker compose exec api python tools/spike_gas_auth.py --mode both）

何を試すか（.env の GAS_INTERNAL_URL・GAS_INTERNAL_TOKEN を使う）:
  bearer    : internal.ping を、委任のアクセストークンを Authorization: Bearer に付けて呼ぶ（--repeat 回）
              ＋ token 違い（unauthorized のはず）＋ Authorization なし（JSON でない＝入れないはず）
              ＋ internal.afterApprove1 の偽の REQ_ID（row_not_found のはず。何も書かない）
  anonymous : 同じことを Authorization なしで（内部用の版が ANYONE_ANONYMOUS のとき）。
              ＋ 匿名の GET で画面が出ないこと（7-1 の第 2 案の守り）
各回について、状態コード（転送を含む）・転送先のホスト・返事が JSON か HTML か・かかった秒数・ok／error を出す。
最後に 14 章 S-1 の合格の条件で「合格／不合格」を出す（終了コード 0／1）。

本番を呼ばない: APP_ENV=staging のときだけ動き、GAS_INTERNAL_URL は gas_client の許可リストで確かめる。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any
from urllib.parse import urljoin, urlsplit

if os.path.isdir("/app") and "/app" not in sys.path:  # api コンテナの中で `python - < tools/…` でも app を読める
    sys.path.insert(0, "/app")

import httpx  # noqa: E402

FAKE_REQ_ID = "LV-00000000-000"
REDIRECT_STATUSES = (301, 302, 303, 307, 308)
MAX_REDIRECTS = 5
PASS_SECONDS = 5.0
ANONYMOUS_PAGE_TEXT = "このアドレスでは画面を開けません"  # 7-1 の第 2 案の守り（G の doGet）


@dataclass
class Probe:
    label: str
    hops: list[int] = field(default_factory=list)  # 状態コードの並び（転送を含む）
    hosts: list[str] = field(default_factory=list)  # 訪ねたホスト
    kind: str = ""  # json / html / other / error
    seconds: float = 0.0
    body: dict | None = None
    error: str = ""
    text: str = ""  # 返事の先頭（表示しない。匿名の GET の確かめだけに使う）

    def line(self) -> str:
        info = ""
        if self.body is not None:
            keys = ("ok", "error", "retryable", "apiVersion", "auth", "scriptId", "spreadsheetId", "migrated")
            info = " ".join(f"{k}={self.body.get(k)!r}" for k in keys if k in self.body)
        return (f"[{self.label}] status={'→'.join(map(str, self.hops)) or '-'} hosts={'→'.join(self.hosts) or '-'} "
                f"reply={self.kind} {self.seconds:.2f}s {info} {self.error}").rstrip()


def _kind_of(resp: httpx.Response) -> str:
    ctype = resp.headers.get("content-type", "").lower()
    head = resp.content[:256].lstrip().lower()
    if "text/html" in ctype or head.startswith(b"<!doctype") or head.startswith(b"<html"):
        return "html"
    try:
        json.loads(resp.content or b"")
        return "json"
    except ValueError:
        return "other"


def probe(label: str, url: str, *, method: str = "POST", body: dict | None = None, bearer: str = "",
          http_factory: Callable[..., httpx.Client] | None = None, timeout: float = 60.0,
          monotonic: Callable[[], float] = time.monotonic) -> Probe:
    """1 回呼んで記録する。転送は自分で追い、転送先には Authorization を付けない（gas_client と同じ）。"""
    result = Probe(label)
    factory = http_factory or (lambda **kw: httpx.Client(**kw))
    headers = {"Content-Type": "application/json"}
    if bearer:
        headers["Authorization"] = f"Bearer {bearer}"
    payload = json.dumps(body or {}, ensure_ascii=False).encode("utf-8")
    started = monotonic()
    try:
        with factory(timeout=httpx.Timeout(timeout, connect=min(30.0, timeout)), follow_redirects=False) as client:
            result.hosts.append(urlsplit(url).hostname or "")
            resp = client.post(url, content=payload, headers=headers) if method == "POST" else client.get(url)
            result.hops.append(resp.status_code)
            while resp.status_code in REDIRECT_STATUSES and len(result.hops) <= MAX_REDIRECTS:
                target = urljoin(str(resp.request.url), resp.headers.get("location", ""))
                result.hosts.append(urlsplit(target).hostname or "")
                resp = client.get(target)
                result.hops.append(resp.status_code)
            result.kind = _kind_of(resp)
            result.text = resp.content[:4000].decode("utf-8", "replace")
            if result.kind == "json":
                data = resp.json()
                result.body = data if isinstance(data, dict) else {"_": data}
    except httpx.TimeoutException:
        result.kind, result.error = "error", "timeout"
    except httpx.HTTPError as exc:
        result.kind, result.error = "error", type(exc).__name__
    result.seconds = monotonic() - started
    return result


def _body(token: str, action: str, **extra: Any) -> dict:
    out = {"internal": True, "token": token, "action": action, "jobId": "spike-" + uuid.uuid4().hex}
    out.update(extra)
    return out


def run_mode(mode: str, url: str, token: str, access_token: Callable[[], str], *, repeat: int = 10,
             http_factory: Callable[..., httpx.Client] | None = None,
             monotonic: Callable[[], float] = time.monotonic, out: Callable[[str], None] = print) -> bool:
    """1 つの呼び方を試し、S-1 の合格の条件に合えば True。"""
    bearer = access_token() if mode == "bearer" else ""
    kw = {"http_factory": http_factory, "monotonic": monotonic}
    out(f"== mode={mode} ==")
    pings = [probe(f"ping#{i + 1}", url, body=_body(token, "internal.ping"), bearer=bearer, **kw)
             for i in range(max(1, repeat))]
    wrong = probe("wrong-token", url, body=_body("x" * 40, "internal.ping"), bearer=bearer, **kw)
    fake = probe("afterApprove1-fake", url, bearer=bearer, **kw,
                 body=_body(token, "internal.afterApprove1", reqId=FAKE_REQ_ID,
                            expect={"status": "承認済"}, args={"signBase64": "", "leaveDate": "2000-01-01"}))
    checks: list[tuple[str, bool]] = []
    for p in pings + [wrong, fake]:
        out(p.line())
    checks.append(("ping が毎回 200 の JSON で ok:true",
                   all(p.kind == "json" and p.hops[-1:] == [200] and (p.body or {}).get("ok") is True for p in pings)))
    checks.append((f"ping が毎回 {PASS_SECONDS:.0f} 秒以内", all(p.seconds <= PASS_SECONDS for p in pings)))
    checks.append(("302 の後の GET で読めた", any(len(p.hops) >= 2 and p.hops[0] == 302 for p in pings)))
    checks.append(("token 違いは unauthorized", (wrong.body or {}).get("error") == "unauthorized"))
    checks.append(("偽の REQ_ID は row_not_found", (fake.body or {}).get("error") == "row_not_found"))
    if mode == "bearer":
        bare = probe("no-authorization", url, body=_body(token, "internal.ping"), **kw)
        out(bare.line())
        checks.append(("Authorization なしは JSON でない（入れない）", bare.kind != "json"))
    else:
        page = probe("anonymous-GET", url, method="GET", **kw)
        out(page.line())
        checks.append(("匿名の GET で画面が出ない",
                       page.kind != "html" or ANONYMOUS_PAGE_TEXT in page.text))
    if pings and pings[0].body:
        b = pings[0].body
        out(f"scriptId={b.get('scriptId')} spreadsheetId={b.get('spreadsheetId')} auth={b.get('auth')} "
            f"apiVersion={b.get('apiVersion')}")
    for name, ok in checks:
        out(f"  {'OK ' if ok else 'NG '} {name}")
    passed = all(ok for _, ok in checks)
    out(f"mode={mode}: {'合格' if passed else '不合格'}")
    return passed


def main(argv: list[str] | None = None) -> int:  # pragma: no cover - ステージングで動かす
    parser = argparse.ArgumentParser(description="GAS の内部用の入口の呼び方を確かめる（S-1）")
    parser.add_argument("--mode", choices=("bearer", "anonymous", "both"), default="both")
    parser.add_argument("--repeat", type=int, default=10)
    args = parser.parse_args(argv)
    from app.config import get_settings
    from app.services.gas_client import assert_gas_url_allowed

    settings = get_settings()
    if not settings.is_staging:
        print("APP_ENV=staging のときだけ動きます（本番の GAS を試さない）")
        return 2
    assert_gas_url_allowed(settings.gas_internal_url, settings)

    def access_token() -> str:
        from app.google_delegate import get_access_token

        return get_access_token(settings=settings)

    modes = ("bearer", "anonymous") if args.mode == "both" else (args.mode,)
    results = {m: run_mode(m, settings.gas_internal_url, settings.gas_internal_token, access_token,
                           repeat=args.repeat) for m in modes}
    print("結果: " + ", ".join(f"{m}={'合格' if ok else '不合格'}" for m, ok in results.items()))
    return 0 if any(results.values()) else 1


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
