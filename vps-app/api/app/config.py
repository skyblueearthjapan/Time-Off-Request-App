"""環境変数の読み込み。必須の値が欠けている・危ない組み合わせのときは起動しない（fail-closed）。

業務用の ID を読むのはこのファイルだけ。本番の ID と、取り違えを防ぐために知っておくステージングの ID は、
コードにも定数として持つ（ほかのモジュールに ID を直書きしない。tests/test_config_safety.py が拒む）。

- APP_ENV=production: 本番の ID が定数と同じでなければ起動しない。ステージングの ID（既知の値）を
  どこかで指していれば起動しない。
- APP_ENV=staging: STAGING_* が必須。本番の ID（定数）をどこかで指していれば起動しない。
"""
from __future__ import annotations

import logging
import re
from functools import lru_cache
from typing import Literal
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from cryptography.fernet import Fernet
from pydantic import Field, field_validator, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

# ===== 本番の ID（.env の値と違えば起動しない） =====
# 本番 DB スプレッドシート「休暇届けapp」（GAS が紐づくファイル）
PROD_SPREADSHEET_ID_FIXED = "16v1tQNCFkDTk2mnAZ4E2eEdHeUQ2Bepdcsn5vBAmllw"
# マスタ元（作業員マスタの元。読むだけ。どの環境からも書かない）
MASTER_SOURCE_SPREADSHEET_ID_FIXED = "1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ"
# カレンダー・電子印・部署承認者の読み先（社内カレンダーマスタ・StampMap・DeptApprovers を持つ別アプリの DB。
# 読むだけ。どの環境からも書かない）
CALENDAR_SOURCE_SPREADSHEET_ID_FIXED = "1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ"
# 本番 GAS の scriptId（internal.ping の照合）
PROD_GAS_SCRIPT_ID_FIXED = "15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i"
# 本番 PDF フォルダ「休暇届け申請書／管理フォルダ」（VPS は書かない。internal.ping の照合だけ）
PROD_PDF_FOLDER_ID_FIXED = "1MQibR08wWsE_OGurO9RocEEDGgDOwo-w"
# 休暇届台帳（総務確認用）（VPS は書かない。GAS が書く）
PROD_LEDGER_SPREADSHEET_ID_FIXED = "1B7tnsT4lr80o3pFMBxB2zEhsx2UW2PCL48YTQc9i1eU"
# 本番 GAS の画面のデプロイ（@106 本番・@104 ブックマーク）。GAS_INTERNAL_URL がこれを指していれば、どの環境も起動しない。
PROD_GAS_SCREEN_DEPLOYMENT_IDS: tuple[str, ...] = (
    "AKfycbx7sB9x7iLigEqZX9LuErK95R2ZC0cMOyU9efr6pIbpnqAnnDC3o5_DGFbtwgSh64j2",
    "AKfycbwHI1uz6tcer59sRC4dYEACDaevSn9a_eJa3J0E8LZgjdWCGPJVaqr_Qa9L-HXVKI8k",
)
# 本番 GAS のデプロイ（画面の 2 つ＋@HEAD）。ステージングの GAS_INTERNAL_URL がこれを指していれば起動しない。
# 本番の内部用のデプロイの deploymentId は、作った後に .env の PROD_GAS_INTERNAL_DEPLOYMENT_ID で渡す
# （両方の環境の .env に書く。Settings.prod_gas_deployment_ids がこの定数と合わせて使う）。
PROD_GAS_DEPLOYMENT_IDS_FIXED: tuple[str, ...] = PROD_GAS_SCREEN_DEPLOYMENT_IDS + (
    "AKfycbw_GAlPaXbETZTNpkK84k0YRwW5koclx1AHFpaqDL8",
)

# ===== ステージングの ID（本番がこれを指していれば起動しない・書かない） =====
# ステージング（DB の「コピーを作成」・台帳のコピー・試験用 PDF フォルダ・写しの GAS）を作ったら、
# ディレクターがその ID をここに足す。空のあいだは .env の STAGING_* と本番の定数の照合だけで守る。
STAGING_SPREADSHEET_ID_KNOWN = "1SG4FV9raYuFu8jaWzCVpXvKqo1LPEO7aR44KifWrYqw"
STAGING_LEDGER_SPREADSHEET_ID_KNOWN = "1FpeButODbprm7Bri2WQWlDgwzYDwiwLb8FvWu1XO2YM"
STAGING_PDF_FOLDER_ID_KNOWN = "1T6ZZ8t-G7fq8BxMUoB7yw2I8i2VXqRCm"
STAGING_GAS_SCRIPT_ID_KNOWN = "1T3bWvTY3bv8R3UeWolurVSRtq1-20NESypGegxh-iP7PPo5xcJqzXyA3"
STAGING_GAS_DEPLOYMENT_IDS_KNOWN: tuple[str, ...] = ("AKfycbz6KZ0tAP6wpDaC-2TZMsiQ_Q7UlnDbBmISOaUwF8KkEOt457euvSuFElbO_aMsZR71tQ",)

# 本番のファイル・フォルダ・スクリプトの ID（ステージングはどれも指してはいけない）
PRODUCTION_IDS: frozenset[str] = frozenset({
    PROD_SPREADSHEET_ID_FIXED,
    MASTER_SOURCE_SPREADSHEET_ID_FIXED,
    CALENDAR_SOURCE_SPREADSHEET_ID_FIXED,
    PROD_GAS_SCRIPT_ID_FIXED,
    PROD_PDF_FOLDER_ID_FIXED,
    PROD_LEDGER_SPREADSHEET_ID_FIXED,
    *PROD_GAS_DEPLOYMENT_IDS_FIXED,
})
# 既知のステージングの ID（本番はどれも指してはいけない。空の値は入れない）
KNOWN_STAGING_IDS: frozenset[str] = frozenset(i for i in (
    STAGING_SPREADSHEET_ID_KNOWN,
    STAGING_LEDGER_SPREADSHEET_ID_KNOWN,
    STAGING_PDF_FOLDER_ID_KNOWN,
    STAGING_GAS_SCRIPT_ID_KNOWN,
    *STAGING_GAS_DEPLOYMENT_IDS_KNOWN,
) if i)

log = logging.getLogger(__name__)

GAS_DEPLOYMENT_ID_RE = re.compile(r"^AKfyc[A-Za-z0-9_-]{30,}\Z")
GAS_INTERNAL_URL_RE = re.compile(r"^https://script\.google\.com/macros/s/([A-Za-z0-9_-]+)/exec\Z")
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\Z")


def split_emails(value: str | None) -> list[str]:
    """カンマ・改行・セミコロン区切りのメール。小文字にして空を除く。"""
    return [e.strip().lower() for e in re.split(r"[,\n;]", value or "") if e.strip()]


def gas_deployment_id(url: str) -> str:
    """GAS_INTERNAL_URL の deploymentId（形が違えば ''）。"""
    m = GAS_INTERNAL_URL_RE.match(url or "")
    return m.group(1) if m else ""


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=None, extra="ignore", case_sensitive=False)

    app_env: Literal["staging", "production"]
    database_url: str = Field(min_length=1)
    public_origin: str = Field(min_length=1)

    # Google ログイン（このアプリ専用の OAuth クライアント）
    google_oauth_client_id: str = Field(min_length=1)
    google_oauth_client_secret: str = Field(min_length=1)
    session_secret: str = Field(min_length=32)
    # オーナーの決定 K-6: ログインの持ち時間は 90 日（無操作でも 90 日は切らない）
    session_max_age_days: int = Field(default=90, ge=1, le=90)
    session_idle_hours: int = Field(default=2160, ge=1, le=24 * 90)
    allowed_domain: str = Field(min_length=1)
    # 管理者（M_SYSTEM_SETTING の ADMIN_EMAILS に足す）。委任の開始もこの人だけ。空なら起動しない
    bootstrap_admins: str = Field(min_length=1)

    # 委任（DELEGATE_EMAIL の権限で Sheets を使う）
    fernet_key: str = Field(min_length=1)
    delegate_email: str = Field(min_length=1)

    # 業務用の ID（読むのはこのファイルだけ）
    prod_spreadsheet_id: str = Field(min_length=1)
    staging_spreadsheet_id: str = ""
    master_source_spreadsheet_id: str = Field(min_length=1)
    prod_pdf_folder_id: str = Field(min_length=1)
    staging_pdf_folder_id: str = ""
    prod_ledger_spreadsheet_id: str = Field(min_length=1)
    staging_ledger_spreadsheet_id: str = ""
    # カレンダー・電子印・部署承認者の読み先（読むだけ。定数と違えば起動しない）
    calendar_source_spreadsheet_id: str = Field(min_length=1)

    # GAS の内部用の入口
    gas_internal_url: str = Field(min_length=1)
    # 合言葉は GAS を呼ぶ worker だけに渡す（compose）。api は空で起動する。worker は空なら起動しない
    # （worker.require_gas_token）。値があるときは 32 字以上。
    gas_internal_token: str = ""
    gas_call_timeout_seconds: float = Field(default=330, gt=0, le=600)
    # bearer（第 1 案・既定）: 委任のアクセストークンを Authorization: Bearer で付ける（Web アプリが access: DOMAIN）。
    # anonymous（第 2 案）: 本文の token だけ（内部用のデプロイだけ ANYONE_ANONYMOUS の版）。仕様書 7-1
    gas_internal_auth: Literal["bearer", "anonymous"] = "bearer"
    # 本番の内部用の GAS のデプロイの deploymentId（作ったら両方の環境に書く。空なら未設定）。
    # production: GAS_INTERNAL_URL がこれと違えば起動しない。staging: GAS_INTERNAL_URL がこれなら起動しない。
    prod_gas_internal_deployment_id: str = ""

    # シートへ書く前の関所（services/write_gate）: 合格した gas_ping がこの分数より古ければ書かない
    sheet_write_ping_max_minutes: int = Field(default=10, ge=1, le=60)

    # DB スプレッドシートの時間帯（IANA 名。シートの値と違えば行の反映を止める）
    db_sheet_timezone: str = Field(min_length=1)
    # 共有タブ（M_SYSTEM_SETTING・M_LOOKUP）の取り込みの間隔（分）。
    # 社内ポータルの URL は .env に持たず、M_SYSTEM_SETTING の PORTAL_URL を使う。
    sheet_config_sync_minutes: int = Field(default=5, ge=1, le=60)

    @field_validator("public_origin")
    @classmethod
    def _origin(cls, v: str) -> str:
        v = v.strip()
        if not v.startswith("https://") or v.endswith("/"):
            raise ValueError("PUBLIC_ORIGIN は https:// で始め、末尾に / を付けないでください")
        return v

    @field_validator("allowed_domain", "delegate_email")
    @classmethod
    def _lower(cls, v: str) -> str:
        return v.strip().lower()

    @field_validator(
        "prod_spreadsheet_id", "staging_spreadsheet_id", "master_source_spreadsheet_id", "prod_pdf_folder_id",
        "staging_pdf_folder_id", "prod_ledger_spreadsheet_id", "staging_ledger_spreadsheet_id", "gas_internal_url",
        "calendar_source_spreadsheet_id",
    )
    @classmethod
    def _strip(cls, v: str) -> str:
        return v.strip()

    @field_validator("prod_gas_internal_deployment_id")
    @classmethod
    def _prod_deployment(cls, v: str) -> str:
        v = v.strip()
        if v and not GAS_DEPLOYMENT_ID_RE.match(v):
            raise ValueError("PROD_GAS_INTERNAL_DEPLOYMENT_ID が GAS の deploymentId の形（AKfyc…）ではありません")
        if v in KNOWN_STAGING_IDS:
            raise ValueError("PROD_GAS_INTERNAL_DEPLOYMENT_ID がステージングの GAS のデプロイです")
        return v

    @field_validator("gas_internal_token")
    @classmethod
    def _gas_token(cls, v: str) -> str:
        if v and len(v) < 32:
            raise ValueError("GAS_INTERNAL_TOKEN は 32 字以上にしてください")
        return v

    @field_validator("gas_internal_url")
    @classmethod
    def _gas_url(cls, v: str) -> str:
        if not GAS_INTERNAL_URL_RE.match(v):
            raise ValueError("GAS_INTERNAL_URL は https://script.google.com/macros/s/<deploymentId>/exec の形にしてください")
        return v

    @field_validator("db_sheet_timezone")
    @classmethod
    def _tz(cls, v: str) -> str:
        v = v.strip()
        try:
            ZoneInfo(v)
        except (ZoneInfoNotFoundError, ValueError) as exc:
            raise ValueError(f"DB_SHEET_TIMEZONE が時間帯の名前（IANA）ではありません: {v}") from exc
        return v

    @field_validator("fernet_key")
    @classmethod
    def _fernet(cls, v: str) -> str:
        try:
            Fernet(v.encode())
        except Exception as exc:  # noqa: BLE001 - 形式違いはすべて起動失敗にする
            raise ValueError("FERNET_KEY が Fernet の鍵の形式ではありません") from exc
        return v

    @model_validator(mode="after")
    def _safety(self) -> "Settings":
        if self.prod_spreadsheet_id != PROD_SPREADSHEET_ID_FIXED:
            raise ValueError("PROD_SPREADSHEET_ID が本番 DB スプレッドシートの ID と違います")
        if self.master_source_spreadsheet_id != MASTER_SOURCE_SPREADSHEET_ID_FIXED:
            raise ValueError("MASTER_SOURCE_SPREADSHEET_ID がマスタ元の ID と違います")
        if self.prod_pdf_folder_id != PROD_PDF_FOLDER_ID_FIXED:
            raise ValueError("PROD_PDF_FOLDER_ID が本番 PDF フォルダの ID と違います")
        if self.prod_ledger_spreadsheet_id != PROD_LEDGER_SPREADSHEET_ID_FIXED:
            raise ValueError("PROD_LEDGER_SPREADSHEET_ID が休暇届台帳の ID と違います")
        if self.calendar_source_spreadsheet_id != CALENDAR_SOURCE_SPREADSHEET_ID_FIXED:
            raise ValueError("CALENDAR_SOURCE_SPREADSHEET_ID がカレンダー・電子印・部署承認者の読み先の ID と違います")
        if not self.bootstrap_admin_list:
            raise ValueError("BOOTSTRAP_ADMINS が空です")
        for email in self.bootstrap_admin_list:
            if not _EMAIL_RE.match(email):
                raise ValueError("BOOTSTRAP_ADMINS にメールアドレスでない値があります")
        if not self.delegate_email.endswith("@" + self.allowed_domain):
            raise ValueError("DELEGATE_EMAIL は社内ドメインのアドレスにしてください")

        deployment = gas_deployment_id(self.gas_internal_url)
        if deployment in PROD_GAS_SCREEN_DEPLOYMENT_IDS:
            raise ValueError("GAS_INTERNAL_URL が画面のデプロイ（@106・@104）を指しています（内部用のデプロイにしてください）")
        staging_values = {
            "STAGING_SPREADSHEET_ID": self.staging_spreadsheet_id,
            "STAGING_PDF_FOLDER_ID": self.staging_pdf_folder_id,
            "STAGING_LEDGER_SPREADSHEET_ID": self.staging_ledger_spreadsheet_id,
        }
        if self.app_env == "staging":
            missing = [name for name, value in staging_values.items() if not value]
            if missing:
                raise ValueError("ステージングに必要な値がありません: " + ", ".join(missing))
            for name, value in staging_values.items():
                if value in PRODUCTION_IDS:
                    raise ValueError(f"{name} が本番の ID を指しています")
            if len(set(staging_values.values())) != len(staging_values):
                raise ValueError("ステージングの DB・台帳・PDF フォルダの ID が重なっています")
            if deployment in PRODUCTION_IDS or deployment in self.prod_gas_deployment_ids:
                raise ValueError("ステージングの GAS_INTERNAL_URL が本番の GAS のデプロイを指しています")
        else:
            if deployment in KNOWN_STAGING_IDS:
                raise ValueError("本番の GAS_INTERNAL_URL がステージングの GAS のデプロイを指しています")
            # 本番は呼んでよい GAS のデプロイを必ず 1 つに決める（空のままでは GAS_INTERNAL_URL を照合できない）
            if not self.prod_gas_internal_deployment_id:
                raise ValueError("本番は PROD_GAS_INTERNAL_DEPLOYMENT_ID（本番の内部用の GAS のデプロイ）が要ります")
            if deployment != self.prod_gas_internal_deployment_id:
                raise ValueError("本番の GAS_INTERNAL_URL が PROD_GAS_INTERNAL_DEPLOYMENT_ID のデプロイと違います")
            for name, value in staging_values.items():
                if value and value in PRODUCTION_IDS:
                    raise ValueError(f"{name} が本番の ID を指しています")
        return self

    @property
    def prod_gas_deployment_ids(self) -> tuple[str, ...]:
        """本番 GAS のデプロイ（定数の画面用＋設定した内部用）。"""
        extra = (self.prod_gas_internal_deployment_id,) if self.prod_gas_internal_deployment_id else ()
        return PROD_GAS_DEPLOYMENT_IDS_FIXED + extra

    @property
    def bootstrap_admin_list(self) -> list[str]:
        return split_emails(self.bootstrap_admins)

    @property
    def is_staging(self) -> bool:
        return self.app_env == "staging"

    @property
    def db_spreadsheet_id(self) -> str:
        """この環境の DB スプレッドシート（行の反映・照合・マスタ同期の書き込み先）。"""
        return self.staging_spreadsheet_id if self.is_staging else self.prod_spreadsheet_id

    @property
    def pdf_folder_id(self) -> str:
        """この環境の PDF フォルダ（GAS が書く。internal.ping の照合だけに使う）。"""
        return self.staging_pdf_folder_id if self.is_staging else self.prod_pdf_folder_id

    @property
    def ledger_spreadsheet_id(self) -> str:
        """この環境の休暇届台帳（GAS が書く。internal.ping の照合だけに使う）。"""
        return self.staging_ledger_spreadsheet_id if self.is_staging else self.prod_ledger_spreadsheet_id

    @property
    def app_url(self) -> str:
        """画面のテンプレート変数 APP_URL（PUBLIC_ORIGIN + '/'）。"""
        return self.public_origin + "/"

    @property
    def oauth_redirect_uri(self) -> str:
        return self.public_origin + "/auth/callback"

    @property
    def delegate_redirect_uri(self) -> str:
        return self.public_origin + "/auth/google-delegate/callback"

    @property
    def db_sheet_tz(self) -> ZoneInfo:
        return ZoneInfo(self.db_sheet_timezone)


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()  # type: ignore[call-arg]
