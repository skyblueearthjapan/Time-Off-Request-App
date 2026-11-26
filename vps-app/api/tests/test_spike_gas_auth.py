"""スパイクの道具 tools/spike_gas_auth.py（S-1）を偽の GAS で試す（両方の案。外へは通信しない）。"""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

import httpx
import pytest

TOOL = Path(__file__).resolve().parents[2] / "tools" / "spike_gas_auth.py"
URL = "https://script.google.com/macros/s/AKfycbTESTstagingInternal/exec"
TOKEN = "test-gas-internal-token-0123456789abcdef"
ACCESS = "ya29.fake-access-token"


@pytest.fixture
def tool(monkeypatch):
    if not TOOL.exists():  # api のイメージ（docker build --target test）には tools/ を入れない
        pytest.skip("tools/spike_gas_auth.py がこのイメージにありません")
    spec = importlib.util.spec_from_file_location("spike_gas_auth", TOOL)
    module = importlib.util.module_from_spec(spec)
    monkeypatch.setitem(sys.modules, "spike_gas_auth", module)  # dataclass が自分のモジュールを引くため
    spec.loader.exec_module(module)
    return module


class FakeGas:
    """mode='bearer': DOMAIN の版（Authorization が無ければログインへ）。mode='anonymous': 誰でも入れる版。"""

    def __init__(self, mode: str):
        self.mode = mode
        self.pending: dict[str, dict] = {}
        self.seen: list[tuple[str, str, str]] = []  # (method, host, Authorization)

    def handler(self, request: httpx.Request) -> httpx.Response:
        auth = request.headers.get("authorization", "")
        self.seen.append((request.method, request.url.host, auth))
        host = request.url.host
        if host == "accounts.google.com":
            return httpx.Response(200, headers={"content-type": "text/html"}, content=b"<html>login</html>")
        if host == "script.googleusercontent.com":
            key = request.url.params.get("k", "")
            return httpx.Response(200, headers={"content-type": "application/json"},
                                  content=json.dumps(self.pending.pop(key)).encode())
        if request.method == "GET":
            return httpx.Response(200, headers={"content-type": "text/html"},
                                  content="<html>このアドレスでは画面を開けません。</html>".encode())
        if self.mode == "bearer" and auth != f"Bearer {ACCESS}":
            return httpx.Response(302, headers={"location": "https://accounts.google.com/ServiceLogin"})
        body = json.loads(request.content)
        if body.get("token") != TOKEN:
            reply = {"ok": False, "error": "unauthorized", "retryable": False}
        elif body["action"] == "internal.ping":
            reply = {"ok": True, "scriptId": "staging-script", "spreadsheetId": "staging-db-sheet-id",
                     "auth": self.mode, "apiVersion": 1}
        else:
            reply = {"ok": False, "error": "row_not_found", "retryable": True}
        key = str(len(self.pending) + len(self.seen))
        self.pending[key] = reply
        return httpx.Response(302, headers={"location": f"https://script.googleusercontent.com/macros/echo?k={key}"})

    def factory(self, **kw):
        return httpx.Client(transport=httpx.MockTransport(self.handler), **kw)


def _run(tool, fake, mode, lines):
    return tool.run_mode(mode, URL, TOKEN, lambda: ACCESS, repeat=3, http_factory=fake.factory, out=lines.append)


def test_bearer_mode_passes_and_hides_secrets(tool):
    fake = FakeGas("bearer")
    lines: list[str] = []
    assert _run(tool, fake, "bearer", lines) is True
    text = "\n".join(lines)
    assert "mode=bearer: 合格" in text and "status=302→200" in text and "reply=json" in text
    assert "script.google.com→script.googleusercontent.com" in text
    assert TOKEN not in text and ACCESS not in text
    # 転送先（googleusercontent）には Authorization を付けない
    assert all(auth == "" for method, host, auth in fake.seen if host != "script.google.com")


def test_bearer_mode_fails_when_domain_requires_login_even_with_token(tool):
    fake = FakeGas("bearer")
    lines: list[str] = []
    assert tool.run_mode("bearer", URL, TOKEN, lambda: "wrong", repeat=2, http_factory=fake.factory,
                         out=lines.append) is False
    assert "reply=html" in "\n".join(lines) and "mode=bearer: 不合格" in lines[-1]


def test_anonymous_mode_passes_and_checks_get_page(tool):
    fake = FakeGas("anonymous")
    lines: list[str] = []
    assert _run(tool, fake, "anonymous", lines) is True
    assert any("anonymous-GET" in line for line in lines)
    assert all(auth == "" for _, _, auth in fake.seen)  # 第 2 案では Authorization を付けない


def test_anonymous_mode_fails_when_get_shows_screen(tool):
    fake = FakeGas("anonymous")
    original = fake.handler

    def handler(request):
        if request.method == "GET" and request.url.host == "script.google.com":
            return httpx.Response(200, headers={"content-type": "text/html"}, content=b"<html>top page</html>")
        return original(request)

    fake.handler = handler
    lines: list[str] = []
    assert _run(tool, fake, "anonymous", lines) is False
    assert any("NG  匿名の GET で画面が出ない" in line for line in lines)


def test_timeout_is_reported(tool):

    def handler(request):
        raise httpx.ReadTimeout("slow", request=request)

    p = tool.probe("x", URL, body={}, http_factory=lambda **kw: httpx.Client(transport=httpx.MockTransport(handler), **kw))
    assert p.kind == "error" and p.error == "timeout"
