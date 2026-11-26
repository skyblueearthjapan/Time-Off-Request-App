"""alembic の 0001_initial_schema が models と食い違っていないか（SQLite で upgrade → 比較 → downgrade）。"""
from pathlib import Path

from alembic import command
from alembic.autogenerate import compare_metadata
from alembic.config import Config
from alembic.migration import MigrationContext
from alembic.script import ScriptDirectory
from sqlalchemy import create_engine, inspect

from app.db import Base

API_DIR = Path(__file__).resolve().parents[1]

ALL_TABLES = {
    "leave_requests", "settings", "lookup", "workers", "calendar", "stamps", "dept_approvers",
    "sessions", "google_credentials", "idempotency_keys", "jobs", "sheet_sync_state", "sync_runs", "audit_log",
}
# T_LEAVE_REQUEST の 26 列（本番の英語の見出し）と DB の列の対応
SHEET_COLUMNS = {  # 仕様書 5-1 の表
    "REQ_ID": "req_id", "FY_START_YEAR": "fiscal_year", "STATUS": "status", "APPLIED_AT": "applied_at",
    "DEPT_ID": "dept_id", "DEPT_NAME": "dept_name", "WORKER_ID": "worker_id", "WORKER_NAME": "worker_name",
    "LEAVE_DATE": "leave_date", "DAY_TYPE": "leave_kubun", "HALF_TYPE": "half_day_type", "LEAVE_KIND": "leave_type",
    "SUBSTITUTE_FOR_WORK_DATE": "substitute_date", "SPECIAL_REASON_REF": "special_reason_ref",
    "SPECIAL_REASON_TEXT": "special_reason", "PAIDLEAVE_DETAIL": "paid_detail",
    "DETAIL_FREE": "additional_detail", "APPROVED_AT": "approved_at", "SIGN_IMAGE_URL": "sign_image_url",
    "PDF_URL": "pdf_url", "CREATED_BY_EMAIL": "created_by_email", "UPDATED_AT": "updated_at_text",
    "APPROVED_BY1_EMAIL": "approved_by1_email", "APPROVED_BY2": "approved_by2",
    "APPROVED_BY2_EMAIL": "approved_by2_email", "APPROVED_AT2": "approved_at2",
}
# 日付・日時のセルが文字だったときの元の文字（cell_kinds が string の行）
TEXT_TWINS = ("applied_at_text", "leave_date_text", "substitute_date_text", "approved_at_text", "approved_at2_text")
# 5-2 の表の列
MASTER_COLUMNS = {
    "settings": {"key", "value", "sheet_row"},
    "lookup": {"sheet_row", "category", "value", "sort_text"},
    "workers": {"sort_order", "cells"},
    "calendar": {"sort_order", "date", "kubun", "cells"},
    "stamps": {"sort_order", "col_a", "col_b", "col_c", "cells"},
    "dept_approvers": {"sort_order", "col_a", "col_b", "col_c", "col_d", "col_e"},
}
# 5-3 の jobs・sheet_sync_state の列
JOB_COLUMNS = {"kind", "lane", "priority", "request_id", "args_json", "requires_json", "guard_json", "state",
               "attempts", "run_after", "leased_until", "last_error", "result_json", "created_at", "updated_at"}
SYNC_STATE_COLUMNS = {"table_name", "pk", "desired_version", "synced_version", "synced_values", "state", "attempts",
                      "next_try_at", "last_error", "first_dirty_at", "last_synced_at", "seq"}


def _config(url: str) -> Config:
    cfg = Config(str(API_DIR / "alembic.ini"))
    cfg.set_main_option("script_location", str(API_DIR / "alembic"))
    cfg.set_main_option("sqlalchemy.url", url)
    cfg.attributes["configure_logger"] = False
    return cfg


def test_revision_chain_is_linear():
    script = ScriptDirectory.from_config(_config("sqlite://"))
    revisions = list(script.walk_revisions())
    assert revisions[-1].revision == "0001" and revisions[-1].down_revision is None
    assert Path(revisions[-1].path).name == "0001_initial_schema.py"
    assert len(script.get_heads()) == 1


def test_upgrade_head_matches_models_and_downgrades(tmp_path):
    url = f"sqlite:///{(tmp_path / 'm.db').as_posix()}"
    cfg = _config(url)
    command.upgrade(cfg, "head")

    engine = create_engine(url)
    with engine.connect() as conn:
        diff = compare_metadata(MigrationContext.configure(conn), Base.metadata)
        tables = set(inspect(conn).get_table_names())
        sync_cols = {c["name"] for c in inspect(conn).get_columns("leave_requests")}
    engine.dispose()
    assert diff == []
    assert ALL_TABLES <= tables
    assert set(Base.metadata.tables) == ALL_TABLES
    assert {"row_version", "sheet_row", "cell_kinds", "pdf_file_id", "sign_png", "leave_date_text"} <= sync_cols
    assert set(SHEET_COLUMNS.values()) <= sync_cols
    assert set(TEXT_TWINS) <= sync_cols

    command.downgrade(cfg, "base")
    engine = create_engine(url)
    with engine.connect() as conn:
        assert set(inspect(conn).get_table_names()) <= {"alembic_version"}
    engine.dispose()


def test_leave_requests_types():
    from sqlalchemy import Date, LargeBinary, Text

    cols = Base.metadata.tables["leave_requests"].c
    assert len(SHEET_COLUMNS) == 26
    for name in ("leave_date", "substitute_date"):
        assert isinstance(cols[name].type, Date) and cols[name].nullable
    for name in ("applied_at", "approved_at", "approved_at2"):
        assert cols[name].type.__class__.__name__ == "UTCDateTime" and cols[name].nullable
    for name in ("fiscal_year", "updated_at_text", *TEXT_TWINS):
        assert isinstance(cols[name].type, Text) and not cols[name].nullable
    assert isinstance(cols.sign_png.type, LargeBinary) and cols.sign_png.nullable
    assert cols.req_id.primary_key


def test_master_tables_match_spec_columns():
    for name, expected in MASTER_COLUMNS.items():
        assert set(Base.metadata.tables[name].c.keys()) == expected, name
    for name in ("workers", "calendar", "stamps", "dept_approvers"):
        assert Base.metadata.tables[name].c.sort_order.primary_key, name
    assert Base.metadata.tables["lookup"].c.sheet_row.primary_key
    assert Base.metadata.tables["settings"].c.key.primary_key


def test_foundation_tables_have_spec_columns():
    assert JOB_COLUMNS <= set(Base.metadata.tables["jobs"].c.keys())
    assert SYNC_STATE_COLUMNS <= set(Base.metadata.tables["sheet_sync_state"].c.keys())
    assert "mail_log" not in Base.metadata.tables


def test_a_tables_have_row_sync_columns():
    for name in ("leave_requests",):
        cols = Base.metadata.tables[name].c
        assert cols.row_version.nullable is False and cols.row_version.server_default.arg == "1"
        assert cols.sheet_row.nullable is True
        assert cols.cell_kinds.nullable is False


def test_text_columns_are_not_null_default_empty():
    for table in Base.metadata.tables.values():
        for col in table.c:
            if col.type.__class__.__name__ == "Text" and not col.primary_key and col.name not in (
                    "refresh_token_encrypted",):
                assert col.nullable is False, f"{table.name}.{col.name}"
                assert col.server_default is not None, f"{table.name}.{col.name}"
