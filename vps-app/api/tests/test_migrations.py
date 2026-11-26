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
SHEET_COLUMNS = {
    "REQ_ID": "req_id", "FY_START_YEAR": "fy_start_year", "STATUS": "status", "APPLIED_AT": "applied_at",
    "DEPT_ID": "dept_id", "DEPT_NAME": "dept_name", "WORKER_ID": "worker_id", "WORKER_NAME": "worker_name",
    "LEAVE_DATE": "leave_date", "DAY_TYPE": "day_type", "HALF_TYPE": "half_type", "LEAVE_KIND": "leave_kind",
    "SUBSTITUTE_FOR_WORK_DATE": "substitute_for_work_date", "SPECIAL_REASON_REF": "special_reason_ref",
    "SPECIAL_REASON_TEXT": "special_reason_text", "PAIDLEAVE_DETAIL": "paidleave_detail",
    "DETAIL_FREE": "detail_free", "APPROVED_AT": "approved_at", "SIGN_IMAGE_URL": "sign_image_url",
    "PDF_URL": "pdf_url", "CREATED_BY_EMAIL": "created_by_email", "UPDATED_AT": "updated_at",
    "APPROVED_BY1_EMAIL": "approved_by1_email", "APPROVED_BY2": "approved_by2",
    "APPROVED_BY2_EMAIL": "approved_by2_email", "APPROVED_AT2": "approved_at2",
}


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

    command.downgrade(cfg, "base")
    engine = create_engine(url)
    with engine.connect() as conn:
        assert set(inspect(conn).get_table_names()) <= {"alembic_version"}
    engine.dispose()


def test_leave_requests_types():
    from sqlalchemy import Date, Integer, LargeBinary

    cols = Base.metadata.tables["leave_requests"].c
    assert len(SHEET_COLUMNS) == 26
    for name in ("leave_date", "substitute_for_work_date"):
        assert isinstance(cols[name].type, Date) and cols[name].nullable
    for name in ("applied_at", "approved_at", "updated_at", "approved_at2"):
        assert cols[name].type.__class__.__name__ == "UTCDateTime" and cols[name].nullable
    assert isinstance(cols.fy_start_year.type, Integer)
    assert isinstance(cols.sign_png.type, LargeBinary) and cols.sign_png.nullable
    assert cols.req_id.primary_key


def test_master_tables_keep_row_order_and_all_cells():
    for name in ("workers", "calendar", "stamps", "dept_approvers", "lookup"):
        cols = Base.metadata.tables[name].c
        assert cols.sheet_row.primary_key, name
        assert cols.cells.nullable is False, name


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
