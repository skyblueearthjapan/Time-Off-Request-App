"""DB の表。担当 F が全表を作る（alembic 0001_initial_schema と同じ形。列名は仕様書 5 章が取り決め）。

共通:
- 文字の列は NOT NULL DEFAULT ''（空とシートの空を同じに扱う）。
- 日付は date、日時は UTCDateTime（timestamptz。時差のない値は拒む）。
- enum は CheckConstraint。
- 業務の表（leave_requests＝T_LEAVE_REQUEST）は row_version（書くたびに +1。sheet_queue.mark_dirty）、
  sheet_row（シートの行番号。まだ書いていない新しい行は NULL）、cell_kinds（型が混ざる列の行ごとの型。
  serial／string／number）を持つ。
- 設定とマスタの表（5-2）はデータ行だけを持つ（見出しの行は持たない）。
- シートの値は正規化せずに生のまま入れる（GAS が normalize_ するところは業務の側で jsutil.normalize_gas を通す）。
"""
from __future__ import annotations

from datetime import date, datetime
from typing import Any

_Date = date  # 列名 date（calendar.date）と型の名前が重ならないように

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
    """シート T_LEAVE_REQUEST（見出し 1 行目。本番は英語の見出し 26 列。列は見出しの文字で対応させる。仕様書 5-1）。

    - 日付・日時の列は、セルが日付（シリアル値）なら date／timestamptz に、文字のセルなら *_text に元の文字を入れる。
      どちらだったかは cell_kinds（{"leave_date": "serial"|"string", ...}）に持つ。
    - fiscal_year は JS の String(値)（cell_kinds.fiscal_year が number／string）。計算し直さずに移す。
    シートに無い列: pdf_file_id（GAS の返事の pdfFileId）、sign_png（手書きサインの PNG。afterApprove1 が済んだら NULL）。
    """

    __tablename__ = "leave_requests"

    req_id: Mapped[str] = mapped_column(String(64), primary_key=True)  # REQ_ID（LV-yyyyMMdd-NNN）
    fiscal_year: Mapped[str] = _text()  # FY_START_YEAR（年度）
    status: Mapped[str] = mapped_column(Text, nullable=False, default="申請中", server_default="申請中")  # STATUS
    applied_at: Mapped[datetime | None] = _ts_null()  # APPLIED_AT（申請日時）
    applied_at_text: Mapped[str] = _text()
    dept_id: Mapped[str] = _text()  # DEPT_ID
    dept_name: Mapped[str] = _text()  # DEPT_NAME
    worker_id: Mapped[str] = _text()  # WORKER_ID（cell_kinds で数か文字か）
    worker_name: Mapped[str] = _text()  # WORKER_NAME
    leave_date: Mapped[date | None] = _date_null()  # LEAVE_DATE（休暇日）
    leave_date_text: Mapped[str] = _text()
    leave_kubun: Mapped[str] = _text()  # DAY_TYPE（休暇区分: 全日休・半日休）
    half_day_type: Mapped[str] = _text()  # HALF_TYPE（半日区分: 午前・午後。全日休でも値が残る行がある）
    leave_type: Mapped[str] = _text()  # LEAVE_KIND（休暇種類: 有給消化（私用）・振替休暇・特別休暇）
    substitute_date: Mapped[date | None] = _date_null()  # SUBSTITUTE_FOR_WORK_DATE（振替元出勤日）
    substitute_date_text: Mapped[str] = _text()
    special_reason_ref: Mapped[str] = _text()  # SPECIAL_REASON_REF（N 区分）
    special_reason: Mapped[str] = _text()  # SPECIAL_REASON_TEXT（特別理由）
    paid_detail: Mapped[str] = _text()  # PAIDLEAVE_DETAIL（有給詳細）
    additional_detail: Mapped[str] = _text()  # DETAIL_FREE（追加詳細。N 区分）
    approved_at: Mapped[datetime | None] = _ts_null()  # APPROVED_AT（承認日時）
    approved_at_text: Mapped[str] = _text()
    sign_image_url: Mapped[str] = _text()  # SIGN_IMAGE_URL（G 区分: GAS が書く）
    pdf_url: Mapped[str] = _text()  # PDF_URL（G 区分: GAS が書く）
    created_by_email: Mapped[str] = _text()  # CREATED_BY_EMAIL
    updated_at_text: Mapped[str] = _text()  # UPDATED_AT（全行が空。持つだけ。N 区分）
    approved_by1_email: Mapped[str] = _text()  # APPROVED_BY1_EMAIL
    approved_by2: Mapped[str] = _text()  # APPROVED_BY2（2 次承認者名）
    approved_by2_email: Mapped[str] = _text()  # APPROVED_BY2_EMAIL
    approved_at2: Mapped[datetime | None] = _ts_null()  # APPROVED_AT2（2 次承認日時）
    approved_at2_text: Mapped[str] = _text()

    # シートに無い列
    pdf_file_id: Mapped[str] = _text()  # GAS の返事の pdfFileId（移行では PDF_URL の /d/<id>/ から）
    # 手書きサインの PNG（1 次承認で入れ、internal.afterApprove1 が済んだら NULL に戻す）
    sign_png: Mapped[bytes | None] = mapped_column(LargeBinary, nullable=True)

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
    """M_LOOKUP（1 行目見出し: カテゴリ・値・SORT）。データ行だけを持つ。"""

    __tablename__ = "lookup"

    sheet_row: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    category: Mapped[str] = _text()  # カテゴリ（生）
    value: Mapped[str] = _text()  # 値（生）
    sort_text: Mapped[str] = _text()  # SORT（文字のまま。数は業務の側で Number(値 or 0)）


# ===== マスタの表（M 区分。master_sync が元から DB のタブへ写し、config_sync が DB のタブから取り込む） =====
# sort_order はデータ行の順（2 行目＝1 から）。見出しの行は持たない。


class Worker(Base):
    """作業員マスタ。cells は見出し付きの 1 行（{見出し: 値}、GAS が読む形）。4-4-a は見出しの文字で読む。"""

    __tablename__ = "workers"

    sort_order: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    cells: Mapped[dict] = _cell_kinds()


class CalendarDay(Base):
    """M_CALENDAR（見出し: 日付・区分・曜日・備考）。"""

    __tablename__ = "calendar"

    sort_order: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    date: Mapped[_Date | None] = _date_null()  # 日付のセルを JST の日付にした値（読めない行は NULL）
    kubun: Mapped[str] = _text()  # 区分（normalize_gas 済み。休日・祝日・出勤土曜）
    cells: Mapped[dict] = _cell_kinds()  # 見出し付きの 1 行


class Stamp(Base):
    """M_STAMP（A メール・B stampFileId・C 備考＝名前。位置で読む）。"""

    __tablename__ = "stamps"

    sort_order: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    col_a: Mapped[str] = _text()  # メール（生）
    col_b: Mapped[str] = _text()  # stampFileId（生）
    col_c: Mapped[str] = _text()  # 備考＝名前（生）
    cells: Mapped[list] = _cells()  # 1 行の値を左から順に


class DeptApprover(Base):
    """M_DEPT_APPROVERS（A dept・B approverEmails・C  approverName・D 役職・E 通知メール。位置で読む）。"""

    __tablename__ = "dept_approvers"

    sort_order: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    col_a: Mapped[str] = _text()  # dept（生）
    col_b: Mapped[str] = _text()  # approverEmails（生。カンマ区切り）
    col_c: Mapped[str] = _text()  # approverName（生）
    col_d: Mapped[str] = _text()  # 役職（生）
    col_e: Mapped[str] = _text()  # 通知メール（生）


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
