"""V_LEAVE_SUMMARY の書き直し（実装仕様 6-3。担当 D。数え方は W の domain.summaries.summary_rows）。

- write(db): 見出し 12 列と全行を values.batchUpdate（A1:L<n+1>、USER_ENTERED）で書き、その後で余りの値を
  values.batchClear（A<n+2>:ZZ と M1:ZZ<n+1>）で消す（書いてから消すので、空の瞬間が無い）。行・タブは消さない・作らない。
  - 年度は元のセルの型（数なら数）、回数・作業員 ID は USER_ENTERED（数字だけなら数＝GAS の setValues と同じ）。
  - 最終有給日: 今のタブの最終有給日のシリアル値の小数部を見て、同じ形で書く（06 5 節: 日付のシリアル値。.625 か整数かは
    実物で決まる）。整数の形（既定）は 'yyyy-MM-dd' の文字を USER_ENTERED（GAS の setValues と同じ）。小数の形は
    「その小数部で、瞬間の JST の日付が最終有給日になる」シリアル値を数で書き、今の列の表示形式を写す。
  - '=' '+' '-' '@' で始まり数でない文字は、先頭に ' を付ける（式にしない）。
  - 歯止め: leave_requests が 0 行なら書かない。タブが無ければ作らずに止めて知らせる。書き込みの関所（write_gate）を守る。
- mark_dirty(db): 書き込みの後の予約（1 次承認・編集・取消。W が同じトランザクションで呼ぶ）。lane='default' の仕事
  'summary_sheet.write' を最後の予約から 10 秒後（最初の予約から最大 60 秒）に 1 回動かす（間引き）。
- rebuild_now(db): すぐ書く（ops_rebuildSummaryNow・有給警告メールの前）。
- 予定 'summary_sheet': 毎日 6:50 JST（統合業務メールの 7:00 の前）。
"""
from __future__ import annotations

import logging
from collections import Counter
from collections.abc import Callable
from datetime import date, datetime, timedelta
from typing import Any

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app import jobs, safety, timeutil
from app.constants import TAB_LEAVE_SUMMARY
from app.jsutil import js_string, js_trim
from app.models import Job, LeaveRequest, SyncRun
from app.services import sheet_codec as codec
from app.services import write_gate
from app.services.sheets_api import SheetsApi, quote_sheet, sheet_id_by_title, sheet_titles

log = logging.getLogger(__name__)

TAB = TAB_LEAVE_SUMMARY
RUN_KIND = "summary_sheet"
JOB_KIND = "summary_sheet.write"
DEBOUNCE = timedelta(seconds=10)
MAX_DELAY = timedelta(seconds=60)
HEADER = ["年度", "部署ID", "部署名", "作業員ID", "作業員名", "有給回数", "有給半日回数", "振替回数", "特別回数",
          "合計回数", "最終有給日", "警告レベル"]
KEYS = ["fiscalYear", "deptId", "deptName", "workerId", "workerName", "paidCount", "paidHalfCount", "substituteCount",
        "specialCount", "totalCount", "lastPaidDate", "warnLevel"]
LAST_PAID_COL = 10  # K
DEFAULT_DATE_FORMAT = {"type": "DATE", "pattern": "yyyy-mm-dd"}
FORMULA_LEADS = ("=", "+", "-", "@")
GRID_FIELDS = "sheets(data(rowData(values(userEnteredFormat(numberFormat)))))"


def purpose() -> str:
    """safety の書き込みの目的の名前（仕様の summary_sheet。土台は summary_sync）。"""
    return "summary_sheet" if "summary_sheet" in safety.PURPOSE_TABS else "summary_sync"


def summary_rows(db: Session, now: datetime) -> list[dict]:
    """W の domain.summaries.summary_rows（実装仕様 10-2）。"""
    from app.domain import summaries

    return summaries.summary_rows(db, now)


# ===== 値 =====


def formula_safe(value: Any) -> Any:
    if isinstance(value, str) and value[:1] in FORMULA_LEADS and not codec.NUMBER_TEXT_RE.match(value.strip()):
        return "'" + value
    return value


def _number(v: Any) -> Any:
    if isinstance(v, float) and v == int(v):
        return int(v)
    return v


def user_entered(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if codec.is_number(value):
        return _number(value)
    return formula_safe(js_string(value))


def detect_last_paid_fraction(column: list, tz: Any) -> float | None:
    """今の最終有給日の列（K2:K の UNFORMATTED）の数のセルで多い小数部。整数なら None。数のセルが無ければ None。"""
    fracs = []
    for row in column:
        v = row[0] if row else ""
        if codec.is_number(v):
            frac = float(v) - int(float(v) // 1)
            fracs.append(int(frac * 1_000_000 + 0.5) / 1_000_000)
    if not fracs:
        return None
    best = Counter(fracs).most_common(1)[0][0]
    return None if best in (0.0, 1.0) else best


def last_paid_serial(d: date, frac: float, tz: Any) -> float:
    """小数部 frac で、瞬間の JST の日付が d になるシリアル値。"""
    base = (d - codec.EPOCH.date()).days
    for day in (base - 1, base, base + 1):
        s = day + frac
        if codec.serial_to_jst_date(s, tz) == d:
            return s
    return base + frac  # pragma: no cover - 時間帯がずれすぎている


def build_grid(rows: list[dict], frac: float | None, tz: Any) -> list[list]:
    """見出し＋全行（USER_ENTERED で渡す値）。"""
    out = [list(HEADER)]
    for r in rows:
        line = [user_entered(r.get(k)) for k in KEYS]
        lp = js_trim(js_string(r.get("lastPaidDate") or ""))
        parsed = codec.parse_date_text(lp) if lp else None
        if frac is not None and parsed is not None:
            line[LAST_PAID_COL] = last_paid_serial(parsed, frac, tz)
        else:
            line[LAST_PAID_COL] = formula_safe(lp)
        out.append(line)
    return out


# ===== 比べる（照合・移行の道具） =====


def _norm(col: int, value: Any, tz: Any, *, from_sheet: bool) -> str:
    if codec.is_blank(value):
        return ""
    if col == LAST_PAID_COL:
        if codec.is_number(value):  # 日付のセル（.625 の行も整数の行も、瞬間の JST の日付）
            return codec.serial_to_jst_date(value, tz).isoformat()
        d = codec.parse_date_text(js_string(value).lstrip("'"))
        return d.isoformat() if d else js_string(value)
    if codec.is_number(value):
        return js_string(_number(value))
    text = js_string(value)
    if not from_sheet:
        text = text[1:] if text.startswith("'") else text
        return codec.sheet_readback(text)[0]
    return text


def compare(expected_rows: list[dict], sheet_values: list[list], tz: Any) -> list[dict]:
    """DB から作った集計と、タブの UNFORMATTED の値（1 行目から）を 12 列で比べる。差の一覧。"""
    expected = build_grid(expected_rows, None, tz)
    diffs = []
    height = max(len(expected), len(sheet_values))
    for r in range(height):
        e = expected[r] if r < len(expected) else []
        s = sheet_values[r] if r < len(sheet_values) else []
        for c in range(len(HEADER)):
            ev = e[c] if c < len(e) else ""
            sv = s[c] if c < len(s) else ""
            a, b = _norm(c, ev, tz, from_sheet=False), _norm(c, sv, tz, from_sheet=True)
            if a != b:
                diffs.append({"row": r + 1, "col": HEADER[c], "db": a, "sheet": b})
        for c in range(len(HEADER), len(s)):
            if not codec.is_blank(s[c]):
                diffs.append({"row": r + 1, "col": f"#{c + 1}", "db": "", "sheet": js_string(s[c])})
    return diffs


# ===== 書く =====


def _settings(app_settings: Any):
    if app_settings is not None:
        return app_settings
    from app.config import get_settings

    return get_settings()


def _grid_size(meta: dict) -> tuple[int, int]:
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == TAB:
            g = s["properties"].get("gridProperties") or {}
            return int(g.get("rowCount") or 0), int(g.get("columnCount") or 0)
    return 0, 0


def _column_format(api: SheetsApi, sid: str) -> dict:
    meta = api.get_spreadsheet_row_metadata(sid, [f"{quote_sheet(TAB)}!K2:K"], GRID_FIELDS)
    for s in meta.get("sheets", []):
        for d in s.get("data", []):
            for row in d.get("rowData", []):
                for cd in row.get("values", []):
                    nf = (cd.get("userEnteredFormat") or {}).get("numberFormat") or {}
                    if nf.get("type") in codec.DATE_FORMAT_TYPES:
                        return {"type": nf["type"], "pattern": nf.get("pattern", "")}
    return dict(DEFAULT_DATE_FORMAT)


def _record(db: Session, started: datetime, result: dict) -> None:
    db.add(SyncRun(kind=RUN_KIND, started_at=started, finished_at=timeutil.now_utc(), ok=bool(result.get("ok")),
                   summary_json=dict(result)))
    db.flush()


def write(db: Session, *, api: SheetsApi | None = None, app_settings: Any = None, now: datetime | None = None,
          gate: bool = True) -> dict:
    """V_LEAVE_SUMMARY を DB から書き直す。戻り値 {ok, rows, lastPaidStyle} か {ok: False, ...}。commit は呼ぶ側。"""
    settings = _settings(app_settings)
    started = timeutil.now_utc()
    now = now or started
    result: dict = {"ok": False}
    if gate:
        allowed, reason = write_gate.check(db, settings)
        if not allowed:
            result.update({"writeBlocked": reason, "error": "書き込みの関所で止めました: " + reason})
            _record(db, started, result)
            return result
    if int(db.scalar(select(func.count()).select_from(LeaveRequest)) or 0) == 0:
        result["error"] = "leave_requests が 0 行のため、書きません（歯止め）"
        _record(db, started, result)
        return result
    rows = summary_rows(db, now)
    sid = settings.db_spreadsheet_id
    tz = settings.db_sheet_tz
    if api is None:
        api = SheetsApi.for_db(db, settings)
    meta = api.get_spreadsheet(sid)
    if TAB not in set(sheet_titles(meta)):
        result["error"] = f"タブ「{TAB}」がありません（タブは作りません）"
        _record(db, started, result)
        return result
    frac = detect_last_paid_fraction(api.get_values(sid, f"{quote_sheet(TAB)}!K2:K", "UNFORMATTED_VALUE"), tz)
    grid = build_grid(rows, frac, tz)
    n = len(grid)
    q = quote_sheet(TAB)
    p = purpose()
    api.values_batch_update(sid, [{"range": f"{q}!A1:L{n}", "values": grid}], tab=TAB, purpose=p,
                            input_option="USER_ENTERED")
    if frac is not None and n > 1:
        fmt = _column_format(api, sid)
        sheet_id = sheet_id_by_title(meta, TAB)
        cells = [{"values": [{"userEnteredFormat": {"numberFormat": fmt}}]} for _ in range(n - 1)]
        api.batch_update(sid, [{"updateCells": {"start": {"sheetId": sheet_id, "rowIndex": 1, "columnIndex": LAST_PAID_COL},
                                                "rows": cells, "fields": "userEnteredFormat.numberFormat"}}],
                         tab=TAB, purpose=p)
    api.values_batch_clear(sid, [f"{q}!A{n + 1}:ZZ", f"{q}!M1:ZZ{n}"], tab=TAB, purpose=p)
    result.update({"ok": True, "rows": n - 1, "lastPaidStyle": "integer" if frac is None else f"fraction {frac}"})
    _record(db, started, result)
    return result


def rebuild_now(db: Session, **kw: Any) -> dict:
    """すぐ書く（運用の欄・有給警告メールの前）。"""
    return write(db, **kw)


def mark_dirty(db: Session) -> None:
    """書き込みの後の予約（W が業務の書き込みと同じトランザクションで呼ぶ。commit は呼ぶ側）。"""
    now = timeutil.now_utc()
    pending = db.scalars(select(Job).where(Job.kind == JOB_KIND, Job.state == "pending")
                         .order_by(Job.id).limit(1)).first()
    if pending is None:
        job_id = jobs.enqueue(db, JOB_KIND, {}, lane="default")
        pending = db.get(Job, job_id)
        created = now
    else:
        created = pending.created_at if pending.created_at.tzinfo else pending.created_at.replace(tzinfo=timeutil.UTC)
    pending.run_after = min(created + MAX_DELAY, now + DEBOUNCE)
    pending.updated_at = now
    db.flush()


def _run_job(ctx: jobs.JobContext) -> dict:
    result = write(ctx.db, app_settings=ctx.app_settings, now=ctx.now)  # Sheets の失敗は例外のまま（jobs が送り直す）
    if result.get("writeBlocked"):
        raise jobs.RetryLater(60, result["writeBlocked"])
    if not result.get("ok"):
        raise jobs.SkipJob(result.get("error", ""))
    return result


def _scheduled(session_factory: Callable[[], Session]) -> dict:
    with session_factory() as db:
        result = write(db)
        db.commit()
    return result


def register_worker() -> None:
    from app import worker

    jobs.register_kind(JOB_KIND, _run_job, lane="default")
    worker.register_schedule("summary_sheet", task=_scheduled, at=(6, 50))


__all__ = ["write", "rebuild_now", "mark_dirty", "compare", "build_grid", "register_worker"]
