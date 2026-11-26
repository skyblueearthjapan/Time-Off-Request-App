"""設定（fail-closed）と安全装置。本番とステージングの ID の取り違えを全部の組み合わせで拒む。"""
from __future__ import annotations

import re
from pathlib import Path

import pytest
from pydantic import ValidationError

from app import config, safety
from app.config import (
    CALENDAR_SOURCE_SPREADSHEET_ID_FIXED,
    KNOWN_STAGING_IDS,
    MASTER_SOURCE_SPREADSHEET_ID_FIXED,
    PROD_GAS_DEPLOYMENT_IDS_FIXED,
    PROD_GAS_SCRIPT_ID_FIXED,
    PROD_LEDGER_SPREADSHEET_ID_FIXED,
    PROD_PDF_FOLDER_ID_FIXED,
    PROD_SPREADSHEET_ID_FIXED,
    PRODUCTION_IDS,
)
from app.safety import UnsafeWriteError
from conftest import TEST_ENV, make_settings

APP_DIR = Path(__file__).resolve().parents[1] / "app"
STAGING_SHEET = TEST_ENV["STAGING_SPREADSHEET_ID"]
STAGING_PDF = TEST_ENV["STAGING_PDF_FOLDER_ID"]
STAGING_LEDGER = TEST_ENV["STAGING_LEDGER_SPREADSHEET_ID"]


def gas_url(deployment: str) -> str:
    return f"https://script.google.com/macros/s/{deployment}/exec"


def prod_settings(**overrides):
    values = {"app_env": "production", "staging_spreadsheet_id": "", "staging_pdf_folder_id": "",
              "staging_ledger_spreadsheet_id": "", "gas_internal_url": gas_url("AKfycbTESTprodInternal")}
    values.update(overrides)
    return make_settings(**values)


# ===== 必須の値・形 =====


@pytest.mark.parametrize("key", [
    "app_env", "database_url", "public_origin", "google_oauth_client_id", "google_oauth_client_secret",
    "session_secret", "allowed_domain", "bootstrap_admins", "fernet_key", "delegate_email", "prod_spreadsheet_id",
    "master_source_spreadsheet_id", "calendar_source_spreadsheet_id", "prod_pdf_folder_id",
    "prod_ledger_spreadsheet_id", "gas_internal_url", "gas_internal_token", "db_sheet_timezone",
])
def test_missing_required_value_refuses_to_start(key, monkeypatch):
    monkeypatch.delenv(key.upper(), raising=False)
    values = {k.lower(): v for k, v in TEST_ENV.items() if k.lower() != key}
    with pytest.raises(ValidationError):
        config.Settings(_env_file=None, **values)  # type: ignore[call-arg]


def test_defaults(settings, monkeypatch):
    assert settings.session_max_age_days == 90
    assert settings.session_idle_hours == 2160  # 90 日（オーナーの決定 K-6）
    assert settings.gas_call_timeout_seconds == 330
    assert settings.gas_internal_auth == "delegate_bearer"
    assert settings.sheet_config_sync_minutes == 5
    assert settings.app_url == "https://leave.lineworks-local.net/"
    assert settings.oauth_redirect_uri == "https://leave.lineworks-local.net/auth/callback"
    assert settings.delegate_redirect_uri == "https://leave.lineworks-local.net/auth/google-delegate/callback"
    monkeypatch.delenv("GAS_INTERNAL_AUTH", raising=False)
    values = {k.lower(): v for k, v in TEST_ENV.items() if k != "GAS_INTERNAL_AUTH"}
    assert config.Settings(_env_file=None, **values).gas_internal_auth == "delegate_bearer"  # type: ignore[call-arg]


def test_settings_have_no_sebastian_or_portal_keys(settings):
    names = set(type(settings).model_fields)
    assert not {n for n in names if "sebastian" in n or n == "portal_url"}


@pytest.mark.parametrize("overrides", [
    {"app_env": "dev"},
    {"public_origin": "http://leave.lineworks-local.net"},
    {"public_origin": "https://leave.lineworks-local.net/"},
    {"session_secret": "short"},
    {"session_max_age_days": 91},
    {"fernet_key": "not-a-key"},
    {"bootstrap_admins": ""},
    {"bootstrap_admins": " , "},
    {"bootstrap_admins": "not-an-email"},
    {"delegate_email": "owner@gmail.com"},
    {"gas_internal_url": "https://script.google.com/a/macros/lineworks-local.info/s/AKfycbX/exec"},
    {"gas_internal_url": "https://evil.example.com/macros/s/AKfycbX/exec"},
    {"gas_internal_url": "https://script.google.com/macros/s/AKfycbX/dev"},
    {"gas_internal_token": "too-short"},
    {"gas_internal_auth": "basic"},
    {"db_sheet_timezone": "Mars/Base"},
])
def test_bad_values_refuse_to_start(overrides):
    with pytest.raises(ValidationError):
        make_settings(**overrides)


@pytest.mark.parametrize("key,wrong", [
    ("prod_spreadsheet_id", STAGING_SHEET),
    ("prod_spreadsheet_id", "1xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"),
    ("master_source_spreadsheet_id", PROD_SPREADSHEET_ID_FIXED),
    ("calendar_source_spreadsheet_id", PROD_SPREADSHEET_ID_FIXED),
    ("calendar_source_spreadsheet_id", MASTER_SOURCE_SPREADSHEET_ID_FIXED),
    ("prod_pdf_folder_id", STAGING_PDF),
    ("prod_ledger_spreadsheet_id", STAGING_LEDGER),
])
@pytest.mark.parametrize("env", ["staging", "production"])
def test_production_ids_must_equal_constants(env, key, wrong):
    base = prod_settings if env == "production" else make_settings
    with pytest.raises(ValidationError):
        base(**{key: wrong})


def test_constants_are_this_apps_ids():
    """読み先（カレンダー・電子印・部署承認者）は本番 DB と別のファイル。"""
    assert CALENDAR_SOURCE_SPREADSHEET_ID_FIXED != PROD_SPREADSHEET_ID_FIXED
    assert len(PRODUCTION_IDS) == 6 + len(PROD_GAS_DEPLOYMENT_IDS_FIXED)


# ===== staging は本番の ID をどこにも書けない =====


@pytest.mark.parametrize("key", ["staging_spreadsheet_id", "staging_pdf_folder_id", "staging_ledger_spreadsheet_id"])
@pytest.mark.parametrize("prod_id", sorted(PRODUCTION_IDS))
def test_staging_refuses_every_production_id(key, prod_id):
    with pytest.raises(ValidationError):
        make_settings(**{key: prod_id})


@pytest.mark.parametrize("key", ["staging_spreadsheet_id", "staging_pdf_folder_id", "staging_ledger_spreadsheet_id"])
def test_staging_requires_its_ids(key):
    with pytest.raises(ValidationError):
        make_settings(**{key: ""})


def test_staging_ids_must_not_overlap():
    with pytest.raises(ValidationError):
        make_settings(staging_ledger_spreadsheet_id=STAGING_SHEET)


@pytest.mark.parametrize("deployment", PROD_GAS_DEPLOYMENT_IDS_FIXED)
def test_staging_refuses_production_gas_deployment(deployment):
    with pytest.raises(ValidationError):
        make_settings(gas_internal_url=gas_url(deployment))


def test_staging_settings_point_to_staging(settings):
    assert settings.db_spreadsheet_id == STAGING_SHEET
    assert settings.ledger_spreadsheet_id == STAGING_LEDGER
    assert settings.pdf_folder_id == STAGING_PDF
    safety.verify_startup(settings)


# ===== production はステージングの ID を指せない =====


def test_production_refuses_known_staging_ids(monkeypatch):
    """ステージングを作って config に ID を足したら、本番はそれを指せない。"""
    known_sheet, known_deploy = "known-staging-sheet", "AKfycbKNOWNstagingDeployment0123456789abcdef"
    monkeypatch.setattr(config, "KNOWN_STAGING_IDS", frozenset({known_sheet, known_deploy}))
    with pytest.raises(ValidationError):
        prod_settings(gas_internal_url=gas_url(known_deploy))
    s = prod_settings(staging_spreadsheet_id=known_sheet)
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(known_sheet, "T_LEAVE_REQUEST", "row_sync", s)


@pytest.mark.parametrize("key", ["staging_spreadsheet_id", "staging_pdf_folder_id", "staging_ledger_spreadsheet_id"])
@pytest.mark.parametrize("prod_id", sorted(PRODUCTION_IDS))
def test_production_refuses_staging_fields_pointing_to_production(key, prod_id):
    with pytest.raises(ValidationError):
        prod_settings(**{key: prod_id})


def test_production_does_not_need_staging_values():
    s = prod_settings()
    assert s.db_spreadsheet_id == PROD_SPREADSHEET_ID_FIXED
    assert s.pdf_folder_id == PROD_PDF_FOLDER_ID_FIXED
    assert s.ledger_spreadsheet_id == PROD_LEDGER_SPREADSHEET_ID_FIXED
    safety.verify_startup(s)


# ===== 書き込みの直前の許可リスト =====

A_TABS = ["T_LEAVE_REQUEST"]
S_TABS = ["V_LEAVE_SUMMARY"]
C_TABS = ["作業員マスタ", "M_CALENDAR", "M_STAMP", "M_DEPT_APPROVERS"]
FORBIDDEN_TABS = ["M_SYSTEM_SETTING", "M_LOOKUP", "BatchLogs", "社内カレンダーマスタ", "StampMap", "DeptApprovers",
                  "2026年度", "Sheet1", "TMP_PDF_x"]
PURPOSES = ["row_sync", "reconcile", "summary_sync", "master_sync"]


@pytest.mark.parametrize("tab", A_TABS)
@pytest.mark.parametrize("purpose", PURPOSES)
def test_a_tabs_writable_only_by_row_sync_and_reconcile(settings, tab, purpose):
    if purpose in ("row_sync", "reconcile"):
        assert safety.assert_writable(STAGING_SHEET, tab, purpose, settings) == STAGING_SHEET
    else:
        with pytest.raises(UnsafeWriteError):
            safety.assert_writable(STAGING_SHEET, tab, purpose, settings)


@pytest.mark.parametrize("tab", S_TABS)
@pytest.mark.parametrize("purpose", PURPOSES)
def test_summary_tab_writable_only_by_summary_sync(settings, tab, purpose):
    if purpose == "summary_sync":
        assert safety.assert_writable(STAGING_SHEET, tab, purpose, settings)
    else:
        with pytest.raises(UnsafeWriteError):
            safety.assert_writable(STAGING_SHEET, tab, purpose, settings)


@pytest.mark.parametrize("tab", C_TABS)
@pytest.mark.parametrize("purpose", PURPOSES)
def test_c_tabs_writable_only_by_master_sync(settings, tab, purpose):
    if purpose == "master_sync":
        assert safety.assert_writable(STAGING_SHEET, tab, purpose, settings)
    else:
        with pytest.raises(UnsafeWriteError):
            safety.assert_writable(STAGING_SHEET, tab, purpose, settings)


@pytest.mark.parametrize("tab", FORBIDDEN_TABS)
@pytest.mark.parametrize("purpose", PURPOSES)
def test_other_tabs_are_never_writable(settings, tab, purpose):
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(STAGING_SHEET, tab, purpose, settings)


def test_unknown_purpose_is_refused(settings):
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(STAGING_SHEET, "T_LEAVE_REQUEST", "export", settings)


ALL_KNOWN_IDS = sorted(PRODUCTION_IDS | KNOWN_STAGING_IDS | {STAGING_LEDGER, STAGING_PDF, "some-other-sheet"})


@pytest.mark.parametrize("sid", ALL_KNOWN_IDS)
def test_staging_writes_only_to_staging_db(settings, sid):
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(sid, "T_LEAVE_REQUEST", "row_sync", settings)


@pytest.mark.parametrize("sid", sorted((PRODUCTION_IDS | KNOWN_STAGING_IDS | {"some-other-sheet", STAGING_SHEET})
                                       - {PROD_SPREADSHEET_ID_FIXED}))
def test_production_writes_only_to_production_db(sid):
    s = prod_settings()
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(sid, "T_LEAVE_REQUEST", "row_sync", s)
    assert safety.assert_writable(PROD_SPREADSHEET_ID_FIXED, "T_LEAVE_REQUEST", "row_sync", s)


@pytest.mark.parametrize("sid", [MASTER_SOURCE_SPREADSHEET_ID_FIXED, CALENDAR_SOURCE_SPREADSHEET_ID_FIXED,
                                 PROD_LEDGER_SPREADSHEET_ID_FIXED, PROD_PDF_FOLDER_ID_FIXED])
def test_read_sources_ledger_and_pdf_are_never_writable(settings, sid):
    assert sid in safety.never_writable_ids(settings)
    for s in (settings, prod_settings()):
        with pytest.raises(UnsafeWriteError):
            safety.assert_writable(sid, "M_CALENDAR", "master_sync", s)


@pytest.mark.parametrize("sid", ["", "   "])
def test_empty_id_is_refused(settings, sid):
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(sid, "T_LEAVE_REQUEST", "row_sync", settings)


@pytest.mark.parametrize("kind", ["deleteDimension", "insertDimension", "deleteSheet", "addSheet", "duplicateSheet",
                                  "sortRange", "moveDimension", "updateSheetProperties", "deleteRange",
                                  "repeatCell", "pasteData", "clearBasicFilter"])
def test_forbidden_batch_requests_are_refused(kind):
    with pytest.raises(UnsafeWriteError):
        safety.assert_write_requests([{kind: {"range": {"sheetId": 7}}}], "T_LEAVE_REQUEST", 7)


def test_allowed_batch_requests():
    safety.assert_write_requests([{"updateCells": {"range": {"sheetId": 7}, "rows": [], "fields": "*"}},
                                  {"appendDimension": {"sheetId": 7, "dimension": "ROWS", "length": 3}}],
                                 "T_LEAVE_REQUEST", 7)
    with pytest.raises(UnsafeWriteError):  # 列は足さない
        safety.assert_write_requests([{"appendDimension": {"sheetId": 7, "dimension": "COLUMNS", "length": 1}}],
                                     "T_LEAVE_REQUEST", 7)
    with pytest.raises(UnsafeWriteError):  # ほかのシート
        safety.assert_write_requests([{"updateCells": {"range": {"sheetId": 8}}}], "T_LEAVE_REQUEST", 7)
    with pytest.raises(UnsafeWriteError):  # 読むだけのタブに行を足さない
        safety.assert_write_requests([{"appendDimension": {"sheetId": 7, "dimension": "ROWS", "length": 1}}],
                                     "M_SYSTEM_SETTING", 7)
    with pytest.raises(UnsafeWriteError):
        safety.assert_write_requests([], "T_LEAVE_REQUEST", 7)


def test_write_methods():
    for tab in C_TABS + S_TABS:
        safety.assert_write_method("values.batchClear", tab)
    for tab in A_TABS + ["M_SYSTEM_SETTING", "M_LOOKUP"]:
        with pytest.raises(UnsafeWriteError):  # 行単位のタブ・読むだけのタブの clear は使わない
            safety.assert_write_method("values.batchClear", tab)
    safety.assert_write_method("values.append", "T_LEAVE_REQUEST")
    for tab in S_TABS + C_TABS:
        with pytest.raises(UnsafeWriteError):
            safety.assert_write_method("values.append", tab)
    with pytest.raises(UnsafeWriteError):
        safety.assert_write_method("values.clear", "M_CALENDAR")


@pytest.mark.parametrize("a1,tab", [("'T_LEAVE_REQUEST'!A1:B2", "T_LEAVE_REQUEST"), ("T_LEAVE_REQUEST!A1", "T_LEAVE_REQUEST"),
                                    ("'It''s'!A1", "It's"), ("'作業員マスタ'", "作業員マスタ")])
def test_a1_sheet_name(a1, tab):
    assert safety.a1_sheet_name(a1) == tab


def test_ranges_must_stay_in_tab():
    safety.assert_ranges_in_tab(["'M_CALENDAR'!A1:B2"], "M_CALENDAR")
    with pytest.raises(UnsafeWriteError):
        safety.assert_ranges_in_tab(["'M_CALENDAR'!A1", "'M_SYSTEM_SETTING'!A1"], "M_CALENDAR")


def test_restore_targets(settings):
    for tab in A_TABS + S_TABS + C_TABS:
        assert safety.assert_restore_target(STAGING_SHEET, tab, settings)
    with pytest.raises(UnsafeWriteError):
        safety.assert_restore_target(STAGING_SHEET, "M_SYSTEM_SETTING", settings)
    assert safety.assert_restore_target(STAGING_SHEET, "M_SYSTEM_SETTING", settings, protected_listed=True)
    with pytest.raises(UnsafeWriteError):
        safety.assert_restore_target(PROD_SPREADSHEET_ID_FIXED, "T_LEAVE_REQUEST", settings)


# ===== GAS の照合（verify_gas_ping） =====


def staging_ping(**overrides):
    ping = {"ok": True, "scriptId": "staging-script-id", "spreadsheetId": STAGING_SHEET,
            "spreadsheetTimeZone": "Etc/GMT", "migrationMode": False,
            "settings": {"PDF_FOLDER_ID": STAGING_PDF, "LEDGER_SSID": STAGING_LEDGER,
                         "MASTER_SOURCE_SSID": MASTER_SOURCE_SPREADSHEET_ID_FIXED,
                         "CALENDAR_SOURCE_SSID": CALENDAR_SOURCE_SPREADSHEET_ID_FIXED}}
    ping.update(overrides)
    return ping


def prod_ping(**overrides):
    ping = {"ok": True, "scriptId": PROD_GAS_SCRIPT_ID_FIXED, "spreadsheetId": PROD_SPREADSHEET_ID_FIXED,
            "spreadsheetTimeZone": "Etc/GMT", "migrationMode": True,
            "settings": {"PDF_FOLDER_ID": PROD_PDF_FOLDER_ID_FIXED, "LEDGER_SSID": PROD_LEDGER_SPREADSHEET_ID_FIXED}}
    ping.update(overrides)
    return ping


def with_setting(ping, key, value):
    ping = dict(ping)
    ping["settings"] = {**ping["settings"], key: value}
    return ping


def test_gas_ping_passes_when_everything_matches(settings):
    assert safety.verify_gas_ping(staging_ping(), settings) == []
    assert safety.verify_gas_ping(prod_ping(), prod_settings()) == []
    # 読み先の設定が空なら照合しない（GAS はコードの既定値を使う）
    assert safety.verify_gas_ping(with_setting(staging_ping(), "MASTER_SOURCE_SSID", ""), settings) == []


@pytest.mark.parametrize("ping", [
    None, "x", {"ok": False, "error": "unauthorized"},
    staging_ping(spreadsheetId=PROD_SPREADSHEET_ID_FIXED),
    staging_ping(scriptId=PROD_GAS_SCRIPT_ID_FIXED),
    staging_ping(scriptId=""),
    staging_ping(spreadsheetTimeZone="Asia/Tokyo"),
    staging_ping(spreadsheetId="another-sheet"),
    with_setting(staging_ping(), "PDF_FOLDER_ID", PROD_PDF_FOLDER_ID_FIXED),
    with_setting(staging_ping(), "LEDGER_SSID", PROD_LEDGER_SPREADSHEET_ID_FIXED),
    with_setting(staging_ping(), "LEDGER_SSID", ""),
    with_setting(staging_ping(), "MASTER_SOURCE_SSID", "other-master"),
    with_setting(staging_ping(), "CALENDAR_SOURCE_SSID", PROD_SPREADSHEET_ID_FIXED),
])
def test_gas_ping_staging_mismatches(settings, ping):
    assert safety.verify_gas_ping(ping, settings)


@pytest.mark.parametrize("ping", [
    prod_ping(spreadsheetId=STAGING_SHEET),
    prod_ping(scriptId="staging-script-id"),
    prod_ping(spreadsheetTimeZone="Asia/Tokyo"),
    with_setting(prod_ping(), "PDF_FOLDER_ID", STAGING_PDF),
    with_setting(prod_ping(), "LEDGER_SSID", "some-other-ledger"),
])
def test_gas_ping_production_mismatches(ping):
    assert safety.verify_gas_ping(ping, prod_settings())


def test_gas_ping_reports_production_id_by_name(settings):
    problems = safety.verify_gas_ping(with_setting(staging_ping(), "PDF_FOLDER_ID", PROD_PDF_FOLDER_ID_FIXED),
                                      settings)
    assert any("PDF_FOLDER_ID が本番の ID" in p for p in problems)


def test_gas_ping_empty_timezone_is_accepted_only_for_gmt_family(settings):
    """Apps Script は GMT のシートで getSpreadsheetTimeZone() に空を返す。
    GMT の系統を DB_SHEET_TIMEZONE にしたときだけ、空を同じ時間帯として合格にする。"""
    assert settings.db_sheet_timezone == "Etc/GMT"
    assert safety.verify_gas_ping(staging_ping(spreadsheetTimeZone=""), settings) == []
    assert safety.verify_gas_ping(staging_ping(spreadsheetTimeZone="Asia/Tokyo"), settings)
    tokyo = settings.model_copy(update={"db_sheet_timezone": "Asia/Tokyo"})
    assert safety.verify_gas_ping(staging_ping(spreadsheetTimeZone=""), tokyo)


# ===== ほかのモジュールに ID を直書きしない =====

# ID の頭 8 文字（定数から作る。この試験のファイルにも ID を直書きしない）と、オーナーのメールの頭
ID_PATTERNS = sorted({i[:8] for i in PRODUCTION_IDS | KNOWN_STAGING_IDS} | {"AKfycb", "imaizumi" + "@"})


def test_ids_are_written_only_in_config_py():
    offenders = []
    for path in APP_DIR.rglob("*.py"):
        if path.name == "config.py" and path.parent == APP_DIR:
            continue
        text = path.read_text(encoding="utf-8")
        for pattern in ID_PATTERNS:
            if pattern in text:
                offenders.append(f"{path.relative_to(APP_DIR)}: {pattern}")
    assert offenders == []


def test_config_defines_the_fixed_constants():
    for name in ("PROD_SPREADSHEET_ID_FIXED", "MASTER_SOURCE_SPREADSHEET_ID_FIXED",
                 "CALENDAR_SOURCE_SPREADSHEET_ID_FIXED", "PROD_GAS_SCRIPT_ID_FIXED",
                 "PROD_PDF_FOLDER_ID_FIXED", "PROD_LEDGER_SPREADSHEET_ID_FIXED"):
        assert re.fullmatch(r"[A-Za-z0-9_-]{25,}", getattr(config, name))
    assert not (PRODUCTION_IDS & KNOWN_STAGING_IDS)
    assert "" not in KNOWN_STAGING_IDS


# ===== 本番の内部用の GAS のデプロイ =====

PROD_INTERNAL = "AKfycbPRODinternalDeployment0123456789abcdefXYZ"


def _prod_values(**kw):
    values = {k.lower(): v for k, v in TEST_ENV.items()}
    values.update(app_env="production", staging_spreadsheet_id="", staging_pdf_folder_id="",
                  staging_ledger_spreadsheet_id="",
                  gas_internal_url=f"https://script.google.com/macros/s/{PROD_INTERNAL}/exec")
    values.update(kw)
    return values


def test_prod_internal_deployment_id_must_match_in_production(monkeypatch):
    for k in ("APP_ENV", "GAS_INTERNAL_URL"):
        monkeypatch.delenv(k, raising=False)
    s = config.Settings(_env_file=None, **_prod_values(prod_gas_internal_deployment_id=PROD_INTERNAL))
    assert PROD_INTERNAL in s.prod_gas_deployment_ids
    with pytest.raises(ValidationError, match="PROD_GAS_INTERNAL_DEPLOYMENT_ID"):
        config.Settings(_env_file=None, **_prod_values(
            prod_gas_internal_deployment_id=PROD_INTERNAL,
            gas_internal_url="https://script.google.com/macros/s/AKfycbOTHERdeployment0123456789abcdefgh/exec"))


def test_staging_refuses_the_configured_prod_internal_deployment():
    with pytest.raises(ValidationError, match="本番の GAS のデプロイ"):
        make_settings(prod_gas_internal_deployment_id=PROD_INTERNAL,
                      gas_internal_url=f"https://script.google.com/macros/s/{PROD_INTERNAL}/exec")
    assert make_settings(prod_gas_internal_deployment_id=PROD_INTERNAL).prod_gas_internal_deployment_id == PROD_INTERNAL


@pytest.mark.parametrize("bad", ["not-an-id", "AKfycshort"])
def test_prod_internal_deployment_id_is_validated(bad):
    with pytest.raises(ValidationError):
        make_settings(prod_gas_internal_deployment_id=bad)


def test_gas_client_refuses_other_prod_deployment(monkeypatch):
    from app.services.gas_client import GasClientConfigError, assert_gas_url_allowed

    for k in ("APP_ENV", "GAS_INTERNAL_URL"):
        monkeypatch.delenv(k, raising=False)
    s = config.Settings(_env_file=None, **_prod_values(prod_gas_internal_deployment_id=PROD_INTERNAL))
    with pytest.raises(GasClientConfigError):
        assert_gas_url_allowed("https://script.google.com/macros/s/AKfycbOTHERdeployment0123456789abcdefgh/exec", s)
    staging = make_settings(prod_gas_internal_deployment_id=PROD_INTERNAL)
    with pytest.raises(GasClientConfigError):
        assert_gas_url_allowed(f"https://script.google.com/macros/s/{PROD_INTERNAL}/exec", staging)
