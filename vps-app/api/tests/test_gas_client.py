"""services/gas_client.GasClient（GAS の内部用の入口）: 本文の形・委任の Bearer・転送の扱い・時間切れ・
200 以外・JSON でない・ログイン画面の HTML・URL の許可リスト。外へは通信しない（MockTransport）。"""
from __future__ import annotations

import json

import httpx
import pytest

from app.config import PROD_GAS_DEPLOYMENT_IDS_FIXED
from app.services import gas_client as gc
from app.services.gas_client import (
    GasAuthError,
    GasClient,
    GasClientConfigError,
    GasTransportError,
    assert_gas_url_allowed,
)
from conftest import TEST_ENV, make_settings

URL = TEST_ENV["GAS_INTERNAL_URL"]
TOKEN = TEST_ENV["GAS_INTERNAL_TOKEN"]
ACCESS = "ya29.test-access-token"
ECHO = "https://script.googleusercontent.com/macros/echo?user_content_key=x"


def _factory(handler):
    transport = httpx.MockTransport(handler)
    return lambda **opts: httpx.Client(transport=transport, **opts)  # noqa: E731


def client_with(handler, settings=None, **kw):
    if settings is not None:
        kw.setdefault("access_token_provider", lambda: ACCESS)
        return GasClient.from_settings(settings, http_factory=_factory(handler), **kw)
    return GasClient(URL, TOKEN, http_factory=_factory(handler), **kw)


def test_call_sends_body_with_bearer_and_drops_it_on_redirect(settings):
    seen = []

    def handler(request: httpx.Request):
        seen.append((request.method, request.url.host, request.headers.get("authorization"), request.content))
        if request.url.host == "script.google.com":
            return httpx.Response(302, headers={"location": ECHO})
        return httpx.Response(200, json={"ok": True, "action": "internal.afterApprove2", "jobId": "j-1"})

    gas = client_with(handler, settings)
    assert gas.auth_mode == "bearer"
    result = gas.call("internal.afterApprove2", job_id="j-1", request_id="LV-20261126-001",
                      expect={"status": "承認済"}, args={"x": 1})
    assert result["ok"] is True
    (m1, h1, a1, body1), (m2, h2, a2, body2) = seen
    assert (m1, h1, a1) == ("POST", "script.google.com", f"Bearer {ACCESS}")
    assert (m2, h2, a2, body2) == ("GET", "script.googleusercontent.com", None, b"")
    assert json.loads(body1) == {"internal": True, "token": TOKEN, "action": "internal.afterApprove2",
                                 "jobId": "j-1", "reqId": "LV-20261126-001", "expect": {"status": "承認済"},
                                 "args": {"x": 1}}
    assert gas.timeout == 330


def test_auth_anonymous_sends_no_authorization(settings):
    seen = []

    def handler(request):
        seen.append(request.headers.get("authorization"))
        return httpx.Response(200, json={"ok": True})

    gas = client_with(handler, make_settings(gas_internal_auth="anonymous"))
    gas.call("internal.ping", job_id="p-1")
    assert seen == [None]


def test_from_settings_uses_delegate_access_token(settings, monkeypatch):
    import app.google_delegate as gd

    calls = []
    monkeypatch.setattr(gd, "get_access_token", lambda scopes=None, settings=None: calls.append(settings) or "tok")
    seen = []

    def handler(request):
        seen.append(request.headers.get("authorization"))
        return httpx.Response(200, json={"ok": True})

    gas = GasClient.from_settings(settings, http_factory=_factory(handler))
    gas.call("internal.ping", job_id="p")
    assert seen == ["Bearer tok"] and calls == [settings]


def test_access_token_failure_is_auth_error(settings):
    def provider():
        raise RuntimeError("secret-detail")

    called = []
    gas = client_with(lambda r: called.append(r) or httpx.Response(200, json={}), settings,
                      access_token_provider=provider)
    with pytest.raises(GasAuthError) as exc:
        gas.call("internal.ping", job_id="p")
    assert "secret-detail" not in str(exc.value) and called == []
    assert isinstance(exc.value, GasTransportError)


def test_login_page_html_is_auth_error(settings):
    html = httpx.Response(200, headers={"content-type": "text/html; charset=utf-8"},
                          text="<!DOCTYPE html><html><title>Google アカウント</title></html>")
    with pytest.raises(GasAuthError) as exc:
        client_with(lambda r: html, settings).call("internal.ping", job_id="p")
    assert TOKEN not in str(exc.value) and ACCESS not in str(exc.value)


def test_redirect_to_login_host_is_auth_error(settings):
    def handler(request):
        return httpx.Response(302, headers={"location": "https://accounts.google.com/ServiceLogin?continue=x"})

    with pytest.raises(GasAuthError):
        client_with(handler, settings).call("internal.ping", job_id="p")


@pytest.mark.parametrize("location", ["http://script.googleusercontent.com/x", ""])
def test_redirect_must_be_https_known_host(settings, location):
    with pytest.raises(GasAuthError):
        client_with(lambda r: httpx.Response(302, headers={"location": location}), settings).call(
            "internal.ping", job_id="p")


def test_redirect_limit(settings):
    hops = []

    def handler(request):
        hops.append(request.url)
        return httpx.Response(302, headers={"location": ECHO})

    with pytest.raises(GasTransportError) as exc:
        client_with(handler, settings).call("internal.ping", job_id="p")
    assert not isinstance(exc.value, GasAuthError)
    assert len(hops) == gc.MAX_REDIRECTS + 1


def test_307_resends_body_without_authorization(settings):
    seen = []

    def handler(request):
        seen.append((request.method, request.headers.get("authorization"), request.content))
        if len(seen) == 1:
            return httpx.Response(307, headers={"location": ECHO})
        return httpx.Response(200, json={"ok": True})

    client_with(handler, settings).call("internal.ping", job_id="p")
    assert seen[1][0] == "POST" and seen[1][1] is None and seen[1][2] == seen[0][2]


@pytest.mark.parametrize("status", [401, 403])
def test_forbidden_is_auth_error(settings, status):
    with pytest.raises(GasAuthError):
        client_with(lambda r: httpx.Response(status, text="no"), settings).call("internal.ping", job_id="p")


def test_failure_json_is_returned_as_is():
    gas = client_with(lambda r: httpx.Response(200, json={"ok": False, "error": "row_not_synced", "retryable": True}))
    assert gas.call("internal.ping", job_id="p-1") == {"ok": False, "error": "row_not_synced", "retryable": True}


def test_timeout_is_transport_error():
    def handler(request):
        raise httpx.ReadTimeout("slow", request=request)

    with pytest.raises(GasTransportError) as exc:
        client_with(handler).call("internal.afterApprove2", job_id="j", timeout=5)
    assert exc.value.timeout is True
    assert TOKEN not in str(exc.value)


def test_connect_error_is_transport_error():
    def handler(request):
        raise httpx.ConnectError("down", request=request)

    with pytest.raises(GasTransportError) as exc:
        client_with(handler).call("internal.ping", job_id="j")
    assert exc.value.timeout is False


@pytest.mark.parametrize("response", [httpx.Response(500, text="err"), httpx.Response(200, text="not json"),
                                      httpx.Response(200, json=[1, 2])])
def test_bad_responses_are_transport_errors(response):
    with pytest.raises(GasTransportError) as exc:
        client_with(lambda r: response).call("internal.ping", job_id="j")
    assert not isinstance(exc.value, GasAuthError)


def test_arguments_are_checked():
    gas = client_with(lambda r: httpx.Response(200, json={"ok": True}))
    with pytest.raises(ValueError):
        gas.call("internal.unknown", job_id="j")
    with pytest.raises(ValueError):
        gas.call("internal.ping", job_id="")
    assert TOKEN not in repr(gas)


def test_actions_are_the_agreed_five():
    assert gc.ACTIONS == {"internal.ping", "internal.afterApprove1", "internal.afterApprove2",
                          "internal.afterUnapprove", "internal.sendWarnMail"}


def test_config_errors(settings):
    with pytest.raises(GasClientConfigError):
        GasClient(URL, "", settings=settings)
    with pytest.raises(GasClientConfigError):
        GasClient(URL, TOKEN, auth_mode="bearer")  # トークンを得る関数が無い
    with pytest.raises(GasClientConfigError):
        GasClient(URL, TOKEN, auth_mode="basic")


def test_url_allow_list(settings):
    for deployment in PROD_GAS_DEPLOYMENT_IDS_FIXED:
        with pytest.raises(GasClientConfigError):
            assert_gas_url_allowed(f"https://script.google.com/macros/s/{deployment}/exec", settings)
    prod = make_settings(app_env="production", staging_spreadsheet_id="", staging_pdf_folder_id="",
                         staging_ledger_spreadsheet_id="",
                         gas_internal_url="https://script.google.com/macros/s/AKfycbTESTprod/exec")
    with pytest.raises(GasClientConfigError):
        assert_gas_url_allowed("https://example.com/", prod)
    assert assert_gas_url_allowed(prod.gas_internal_url, prod) == "AKfycbTESTprod"
    pinned = make_settings(app_env="production", staging_spreadsheet_id="", staging_pdf_folder_id="",
                           staging_ledger_spreadsheet_id="",
                           gas_internal_url="https://script.google.com/macros/s/AKfycbTESTprodInternal0123456789abcdef/exec",
                           prod_gas_internal_deployment_id="AKfycbTESTprodInternal0123456789abcdef")
    with pytest.raises(GasClientConfigError):
        assert_gas_url_allowed("https://script.google.com/macros/s/AKfycbOTHER/exec", pinned)


def test_token_and_access_token_are_not_logged(settings, caplog):
    """本文の token・委任のアクセストークンをログに出さない（仕様書 12-2 F）。"""
    import logging

    caplog.set_level(logging.DEBUG)
    client_with(lambda r: httpx.Response(200, json={"ok": True}), settings).call(
        "internal.afterApprove1", job_id="j-9", request_id="LV-20261126-009")
    assert caplog.records
    assert TOKEN not in caplog.text and ACCESS not in caplog.text
