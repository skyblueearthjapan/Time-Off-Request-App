"""DB の表。担当 F が全表を作る（alembic 0001_initial_schema と同じ形）。足りない列は後の担当が 0002 で足す。

共通:
- 文字の列は NOT NULL DEFAULT ''（空とシートの空を同じに扱う）。
- 日付は date、日時は UTCDateTime（timestamptz。時差のない値は拒む）。
- enum は CheckConstraint。
- 業務の表（A 区分: leave_requests＝T_LEAVE_REQUEST）は row_version（書くたびに +1。sheet_queue.mark_dirty）、
  sheet_row（シートの行番号。まだ書いていない新しい行は NULL）、cell_kinds（型が混ざる列の行ごとの型）を持つ。
- マスタ・M_LOOKUP の表は、元のタブの行番号（sheet_row。1 行目の見出しも行として持つ）と、元の列を全部
  （cells: 1 行の値を左から順に並べた JSON の配列）持つ。よく使う列は名前の付いた列にも写す。
- シートの値は正規化せずに生のまま入れる（GAS が normalize_ するところは業務の側で jsutil.normalize_gas を通す）。
"""
from __future__ import annotations

from datetime import date, datetime
from typing import Any

from sqlalchemy import (
    JSON,
    BigInteger,
    Boolean,
    CheckConstraint,
    Date,
    Index,
    Integer,
    LargeBinary,
    Sequence,
    String,
    Text,
)
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import Mapped, mapped_column

from app.constants import (
    IDEMPOTENCY_STATES,
    JOB_LANES,
    JOB_STATES,
    LEAVE_STATUSES,
    SYNC_STATES,
)
from app.db import Base, UTCDateTime

# 自動で増える整数の主キー（PostgreSQL は BIGINT、SQLite は INTEGER でないと自動採番されない）
BigIntPK = BigInteger().with_variant(Integer, "sqlite")
# PostgreSQL は jsonb、SQLite は JSON（文字）
JsonB = JSON().with_variant(JSONB(), "postgresql")

# sheet_sync_state.seq（新しい行を作った順）の連番。PostgreSQL だけで使う（SQLite は sheet_queue が max+1）
SHEET_SYNC_SEQ = Sequence("sheet_sync_state_seq", metadata=Base.metadata)


def _in(column: str, values: tuple[str, ...], allow_empty: bool = False) -> str:
    allowed = list(values) + ([""] if allow_empty else [])
    quoted = ", ".join("'" + v.replace("'", "''") + "'" for v in allowed)
    return f"{column} IN ({quoted})"


def _text(**kw: Any):
    return mapped_column(Text, nullable=False, default="", server_default="", **kw)


def _int_null():
    return mapped_column(Integer, nullable=True)


def _row_version():
    return mapped_column(Integer, nullable=False, default=1, server_default="1")


def _sheet_row():
    return mapped_column(Integer, nullable=True)


def _cell_kinds():
    return mapped_column(JsonB, nullable=False, default=dict, server_default="{}")


def _cells():
    return mapped_column(JsonB, nullable=False, default=list, server_default="[]")


def _date_null():
    return mapped_column(Date, nullable=True)


def _ts_null():
    return mapped_column(UTCDateTime, nullable=True)


# ===== 業務の表（A 区分。行単位でシートへ反映する） =====


class LeaveRequest(Base):
    """シート T_LEAVE_REQUEST（見出し 1 行目。本番は英語の見出し 26 列。列は見出しの文字で対応させる）。

    シートに無い列: pdf_file_id（GAS が PDF を作ったときの Drive の ID）、sign_png（手書きサインの PNG を GAS へ
    渡すまで置く）、*_text（日付にならない元の文字）、source_*（ステージングへの移行で空にした本番の値）。
    """

    __tablename__ = "leave_requests"

    req_id: Mapped[str] = mapped_column(String(64), primary_key=True)  # A REQ_ID（LV-yyyyMMdd-NNN）
    fy_start_year: Mapped[int | None] = _int_null()  # B FY_START_YEAR（年度。3/16 始まりの年）
    status: Mapped[str] = mapped_column(Text, nullable=False, default="申請中", server_default="申請中")  # C STATUS
    applied_at: Mapped[datetime | None] = _ts_null()  # D APPLIED_AT
    dept_id: Mapped[str] = _text()  # E DEPT_ID
    dept_name: Mapped[str] = _text()  # F DEPT_NAME
    worker_id: Mapped[str] = _text()  # G WORKER_ID
    worker_name: Mapped[str] = _text()  # H WORKER_NAME
    leave_date: Mapped[date | None] = _date_null()  # I LEAVE_DATE
    leave_date_text: Mapped[str] = _text()  # 日付にならない元の文字
    day_type: Mapped[str] = _text()  # J DAY_TYPE（全日休・半日休）
    half_type: Mapped[str] = _text()  # K HALF_TYPE（午前・午後。全日休でも値が残る行がある）
    leave_kind: Mapped[str] = _text()  # L LEAVE_KIND（有給消化（私用）・振替休暇・特別休暇）
    substitute_for_work_date: Mapped[date | None] = _date_null()  # M SUBSTITUTE_FOR_WORK_DATE
    substitute_for_work_date_text: Mapped[str] = _text()
    special_reason_ref: Mapped[str] = _text()  # N SPECIAL_REASON_REF
    special_reason_text: Mapped[str] = _text()  # O SPECIAL_REASON_TEXT
    paidleave_detail: Mapped[str] = _text()  # P PAIDLEAVE_DETAIL
    detail_free: Mapped[str] = _text()  # Q DETAIL_FREE
    approved_at: Mapped[datetime | None] = _ts_null()  # R APPROVED_AT
    sign_image_url: Mapped[str] = _text()  # S SIGN_IMAGE_URL（GAS が書く）
    pdf_url: Mapped[str] = _text()  # T PDF_URL（GAS が書く）
    created_by_email: Mapped[str] = _text()  # U CREATED_BY_EMAIL
    updated_at: Mapped[datetime | None] = _ts_null()  # V UPDATED_AT
    approved_by1_email: Mapped[str] = _text()  # W APPROVED_BY1_EMAIL
    approved_by2: Mapped[str] = _text()  # X APPROVED_BY2（名前）
    approved_by2_email: Mapped[str] = _text()  # Y APPROVED_BY2_EMAIL
    approved_at2: Mapped[datetime | None] = _ts_null()  # Z APPROVED_AT2

    # シートに無い列
    pdf_file_id: Mapped[str] = _text()  # GAS が PDF を作ったときの Drive の ID（ゴミ箱へ移すため）
    # PDF の列を空にした操作の、まだゴミ箱へ移していない前の ID（GAS の仕事が使う）
    old_pdf_file_id_pending: Mapped[str] = _text()
    # 手書きサインの PNG（1 次承認で受け取り、GAS がサイン画像を保存したら消す）
    sign_png: Mapped[bytes | None] = mapped_column(LargeBinary, nullable=True)
    # ステージングへの移行で空にした本番の値（参照用。仕事は使わない。シートには書かない）
    source_pdf_file_id: Mapped[str] = _text()
    source_pdf_url: Mapped[str] = _text()
    source_sign_image_url: Mapped[str] = _text()

    row_version: Mapped[int] = _row_version()
    sheet_row: Mapped[int | None] = _sheet_row()
    cell_kinds: Mapped[dict] = _cell_kinds()

    __table_args__ = (
        CheckConstraint(_in("status", LEAVE_STATUSES), name="ck_leave_requests_status"),
        Index("ix_leave_requests_status", "status"),
        Index("ix_leave_requests_leave_date", "leave_date"),
        Index("ix_leave_requests_worker_id", "worker_id"),
        Index("ix_leave_requests_dept_name", "dept_name"),
    )


# ===== 共有タブの表（読むだけ。config_sync が取り込む。VPS は書かない） =====


class Setting(Base):
    """M_SYSTEM_SETTING（1 行目見出し、2 行目から A=key、B=value）。同じキーは下の行が勝つ（取り込みで残すのは勝った行）。"""

    __tablename__ = "settings"

    key: Mapped[str] = mapped_column(String(256), primary_key=True)  # 前後の空白を取る
    value: Mapped[str] = _text()
    sheet_row: Mapped[int | None] = _sheet_row()


class Lookup(Base):
    """M_LOOKUP（1 行目見出し: カテゴリ・値・SORT）。"""

    __tablename__ = "lookup"

    sheet_row: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    category: Mapped[str] = _text()  # A カテゴリ
    value: Mapped[str] = _text()  # B 値
    sort: Mapped[str] = _text()  # C SORT（文字のまま。数は業務の側で読む）
    cells: Mapped[list] = _cells()


# ===== マスタの表（C 区分。元から作り、DB のタブにも同じ形で書く） =====


class Worker(Base):
    """作業員マスタ（元: マスタ元の「作業員マスタ」。見出し: 作業員コード・氏名・部署・担当業務・拠点・スタッフ種類）。"""

    __tablename__ = "workers"

    sheet_row: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    worker_code: Mapped[str] = _text()  # 作業員コード
    worker_name: Mapped[str] = _text()  # 氏名
    dept: Mapped[str] = _text()  # 部署
    duty: Mapped[str] = _text()  # 担当業務
    site: Mapped[str] = _text()  # 拠点
    staff_type: Mapped[str] = _text()  # スタッフ種類
    cells: Mapped[list] = _cells()


class CalendarDay(Base):
    """M_CALENDAR（元: 読み先の「社内カレンダーマスタ」。見出し: 日付・区分・曜日・備考）。"""

    __tablename__ = "calendar"

    sheet_row: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    day: Mapped[date | None] = _date_null()  # A 日付
    day_text: Mapped[str] = _text()  # 日付にならない元の文字（見出しの行も）
    kind: Mapped[str] = _text()  # B 区分（休日・祝日・出勤土曜）
    weekday: Mapped[str] = _text()  # C 曜日（Sat など。文字のまま）
    note: Mapped[str] = _text()  # D 備考
    cells: Mapped[list] = _cells()


class Stamp(Base):
    """M_STAMP（元: 読み先の「StampMap」。見出し: メール・stampFileId・備考）。"""

    __tablename__ = "stamps"

    sheet_row: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    email: Mapped[str] = _text()  # A メール
    stamp_file_id: Mapped[str] = _text()  # B stampFileId
    note: Mapped[str] = _text()  # C 備考
    cells: Mapped[list] = _cells()


class DeptApprover(Base):
    """M_DEPT_APPROVERS（元: 読み先の「DeptApprovers」。見出し: dept・approverEmails・ approverName・役職・通知メール）。"""

    __tablename__ = "dept_approvers"

    sheet_row: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    dept: Mapped[str] = _text()  # A
    emails: Mapped[str] = _text()  # B（生の文字。カンマ区切り）
    name: Mapped[str] = _text()  # C
    title: Mapped[str] = _text()  # D 役職
    notify_emails: Mapped[str] = _text()  # E 通知メール
    cells: Mapped[list] = _cells()


# ===== 土台の表 =====


class UserSession(Base):
    """ログインのセッション。クッキーには署名付きの乱数、DB には その SHA-256 だけ。"""

    __tablename__ = "sessions"

    token_hash: Mapped[str] = mapped_column(String(64), primary_key=True)
    email: Mapped[str] = mapped_column(String(256), nullable=False, index=True)
    name: Mapped[str] = _text()
    created_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    # min(最後に使った時刻 + SESSION_IDLE_HOURS, created_at + SESSION_MAX_AGE_DAYS)。90 日
    expires_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False, index=True)


class GoogleCredential(Base):
    """委任トークン（DELEGATE_EMAIL の Sheets・Drive の権限）。リフレッシュトークンは Fernet で暗号化。"""

    __tablename__ = "google_credentials"

    name: Mapped[str] = mapped_column(String(32), primary_key=True)  # 'delegate'
    account_email: Mapped[str] = mapped_column(String(256), nullable=False, default="", server_default="")
    refresh_token_encrypted: Mapped[str] = mapped_column(Text, nullable=False)
    scopes: Mapped[str] = _text()  # 空白区切り（実際に許可されたスコープ）
    granted_by_email: Mapped[str] = mapped_column(String(256), nullable=False, default="", server_default="")
    granted_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)


class IdempotencyKey(Base):
    """二重押し対策。同じ人の同じキーの 2 回目は、1 回目の結果を返す。24 時間で消す。"""

    __tablename__ = "idempotency_keys"

    user_email: Mapped[str] = mapped_column(String(256), primary_key=True)
    key: Mapped[str] = mapped_column(String(128), primary_key=True)
    fn: Mapped[str] = mapped_column(String(64), nullable=False)
    state: Mapped[str] = mapped_column(String(16), nullable=False, default="running", server_default="running")
    response_json: Mapped[str] = _text()
    created_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False, index=True)

    __table_args__ = (CheckConstraint(_in("state", IDEMPOTENCY_STATES), name="ck_idempotency_keys_state"),)


class Job(Base):
    """後回しの仕事（outbox）。業務の書き込みと同じトランザクションで積む（app/jobs.py）。

    lane: 'gas'（GAS の内部用の入口を呼ぶ。全体で同時に 1 本）／'default'。
    kind: GAS の action 名（internal.afterApprove1 など）か、D・M が登録した名前。
    requires_json: [{table, pk, version}]。この版まで行の反映が済むまで動かさない。
    guard_json: 積んだ時点の status など（担当 M が実行前に照合する）。
    """

    __tablename__ = "jobs"

    id: Mapped[int] = mapped_column(BigIntPK, primary_key=True, autoincrement=True)
    kind: Mapped[str] = mapped_column(String(64), nullable=False)
    lane: Mapped[str] = mapped_column(String(16), nullable=False, default="gas", server_default="gas")
    priority: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    request_id: Mapped[str] = mapped_column(String(64), nullable=False, default="", server_default="")
    args_json: Mapped[str] = mapped_column(Text, nullable=False, default="{}", server_default="{}")
    requires_json: Mapped[str] = mapped_column(Text, nullable=False, default="[]", server_default="[]")
    guard_json: Mapped[str] = mapped_column(Text, nullable=False, default="{}", server_default="{}")
    state: Mapped[str] = mapped_column(String(16), nullable=False, default="pending", server_default="pending")
    attempts: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    run_after: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    leased_until: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    last_error: Mapped[str] = _text()
    result_json: Mapped[str] = _text()
    batch_id: Mapped[str] = mapped_column(String(64), nullable=False, default="", server_default="")
    created_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    finished_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)

    __table_args__ = (
        CheckConstraint(_in("state", JOB_STATES), name="ck_jobs_state"),
        CheckConstraint(_in("lane", JOB_LANES), name="ck_jobs_lane"),
        Index("ix_jobs_lane_state_run_after", "lane", "state", "run_after"),
        Index("ix_jobs_request_id_state", "request_id", "state"),
        Index("ix_jobs_batch_id", "batch_id"),
    )


class SheetSyncState(Base):
    """行の反映の状態（決定 A 相当）。desired_version > synced_version の行を worker が書く。"""

    __tablename__ = "sheet_sync_state"

    table_name: Mapped[str] = mapped_column(String(32), primary_key=True)
    pk: Mapped[str] = mapped_column(String(64), primary_key=True)
    desired_version: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    synced_version: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    # 最後にシートへ書いた、またはシートにあると分かっている値（列ごと）
    synced_values: Mapped[dict] = mapped_column(JsonB, nullable=False, default=dict, server_default="{}")
    state: Mapped[str] = mapped_column(String(16), nullable=False, default="ok", server_default="ok")
    attempts: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    next_try_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    last_error: Mapped[str] = _text()
    first_dirty_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    last_synced_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    # 新しい行を作った順（PostgreSQL は sheet_sync_state_seq、SQLite は max+1。sheet_queue が振る）
    seq: Mapped[int] = mapped_column(BigInteger, nullable=False, default=0, server_default="0")

    __table_args__ = (
        CheckConstraint(_in("state", SYNC_STATES), name="ck_sheet_sync_state_state"),
        Index("ix_sheet_sync_state_pending", "state", "next_try_at"),
        Index("ix_sheet_sync_state_seq", "seq"),
    )


class SyncRun(Base):
    """定期処理・同期・照合・gas_ping の記録。worker は kind='schedule:<名前>' で予定の枠を判定する。"""

    __tablename__ = "sync_runs"

    id: Mapped[int] = mapped_column(BigIntPK, primary_key=True, autoincrement=True)
    kind: Mapped[str] = mapped_column(String(64), nullable=False)
    started_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    finished_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    ok: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="false")
    summary_json: Mapped[dict] = mapped_column(JsonB, nullable=False, default=dict, server_default="{}")

    __table_args__ = (Index("ix_sync_runs_kind_started", "kind", "started_at"),)


class AuditLog(Base):
    """書き込みの RPC の記録（rpc.py が業務の書き込みと同じトランザクションで残す。消さない）。"""

    __tablename__ = "audit_log"

    id: Mapped[int] = mapped_column(BigIntPK, primary_key=True, autoincrement=True)
    at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False, index=True)
    user_email: Mapped[str] = mapped_column(String(256), nullable=False, default="", server_default="")
    fn: Mapped[str] = mapped_column(String(64), nullable=False)
    request_ids: Mapped[str] = mapped_column(Text, nullable=False, default="[]", server_default="[]")  # JSON 配列
    ok: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="false")
    error: Mapped[str] = _text()


# sheet_queue が扱う A 区分の表（table_name → モデル）
SYNCED_TABLES: dict[str, type[Base]] = {
    "leave_requests": LeaveRequest,
}

__all__ = [
    "AuditLog",
    "CalendarDay",
    "DeptApprover",
    "GoogleCredential",
    "IdempotencyKey",
    "Job",
    "LeaveRequest",
    "Lookup",
    "SHEET_SYNC_SEQ",
    "SYNCED_TABLES",
    "Setting",
    "SheetSyncState",
    "Stamp",
    "SyncRun",
    "UserSession",
    "Worker",
]
