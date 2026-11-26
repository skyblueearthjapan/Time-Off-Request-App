"""T_LEAVE_REQUEST への行単位の反映（決定 A・実装仕様 6-2。担当 D）。

worker の行の反映のレーン（全体で 1 本。2 秒ごと）が run_once(session_factory, limit=50) を呼ぶ。

1. sheet_sync_state の desired_version > synced_version の行を選ぶ（blocked と再試行の待ちは除く）。
   新しい行（sheet_row が NULL）は seq 順（前の新しい行が待ちなら後ろも待つ＝作った順を守る）、ほかは first_dirty_at 順。
2. DB の今の値を synced_values（最後にシートへ書いた値）と比べ、違う V 列だけを書く。G 列（SIGN_IMAGE_URL・PDF_URL）と
   N 列は書かない（GAS が書いた値を古い値で上書きしない）。新しい行だけは、行の全列を書く（G・N 列は空）。
3. 書く前に 1 行目の見出しを読み、決まりの見出し（別名を含む）が 1 つでも無ければレーンを止める（列を足さない）。
   スプレッドシートの時間帯（properties.timeZone）が DB_SHEET_TIMEZONE と違えばレーンを止める（起動時と 1 時間ごと）。
   安全装置の拒否・書き込みの関所（write_gate）もレーンを止める。止めたことは sync_runs（kind='row_sync'）に残す。
4. 今ある行は sheet_row に書く。書く前に REQ_ID の列を読んで同じか確かめ、違えば REQ_ID の列から探し直す。
   見つからなければ書かずに state='blocked'。
   新しい行は「REQ_ID の列で値のある最後の行」と「どれかの列に値がある最後の行」（GAS の getLastRow）の遅い方の次へ、
   作った順に書く。values.append は使わない（本番には REQ_ID が空の半端な行が 52 行あり、append は表の切れ目を
   読み違えるおそれがある）。行の数が足りなければ appendDimension（ROWS、末尾だけ）。書いた位置を sheet_row に入れる。
5. タブにまとめて 1 回の spreadsheets.batchUpdate（updateCells だけ）。失敗したら行ごとに書き直し、失敗した行だけ
   10 秒・30 秒・1 分・2 分・5 分、以後 5 分ごとに再試行する。
6. 書けたら synced_version＝読んだ時点の row_version、synced_values を書いた列だけ更新する。

行・列・タブの挿入・削除・並べ替え・作り直しは出さない（SheetsApi と safety が拒む）。
"""
from __future__ import annotations

import logging
from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import func, select, update
from sqlalchemy.orm import Session

from app import sheet_queue, timeutil
from app.jsutil import normalize_gas
from app.models import LeaveRequest, SheetSyncState, SyncRun
from app.safety import UnsafeWriteError
from app.services import sheet_codec as codec
from app.services import write_gate
from app.services.sheets_api import (
    SheetsApi,
    SheetsError,
    column_letter,
    quote_sheet,
    sheet_id_by_title,
    spreadsheet_timezone,
)

log = logging.getLogger(__name__)

RETRY_DELAYS = (10, 30, 60, 120, 300)  # 秒。以後は 300 秒ごと
RECHECK_SECONDS = 300
GATE_RECHECK_SECONDS = 30
TZ_CHECK_SECONDS = 3600
BACKOFF_SECONDS = 30
REF_WINDOW = 20  # 表示形式を写すために読む、最後のデータ行からの行数
REF_CACHE_SECONDS = 600
LATE_SECONDS = 60  # 統合業務メールの 17:00・18:00 に間に合う遅れの上限
GRID_FIELDS = "sheets(data(startRow,startColumn,rowData(values(userEnteredValue,userEnteredFormat(numberFormat)))))"
RUN_KIND = "row_sync"
TABLE = codec.TABLE
# 表示形式を「同じ列で多い形」で写す列（06 2-5: APPROVED_AT2 は M/d/yyyy の行が多い）
MAJORITY_FORMAT_HEADERS = ("APPROVED_AT2",)


class LaneHalt(RuntimeError):
    """レーンを止める理由（見出しの違い・時間帯の違い・タブが無い）。再試行しない。"""


class GateClosed(LaneHalt):
    """書き込みの関所（write_gate）で止めた。"""


@dataclass
class _Lane:
    halted: str = ""
    recheck_at: datetime | None = None
    tz_checked_at: datetime | None = None
    refs: tuple[datetime, dict] | None = None
    backoff_until: datetime | None = None


_lane = _Lane()


def reset_lane() -> None:
    """試験用・手での再開用。"""
    global _lane
    _lane = _Lane()


def lane_state() -> dict:
    return {"halted": _lane.halted, "recheckAt": _lane.recheck_at.isoformat() if _lane.recheck_at else None}


def _default_api(session_factory: Callable[[], Session], app_settings: Any) -> SheetsApi:
    with session_factory() as db:
        return SheetsApi.for_db(db, app_settings)


# 試験で差し替える（偽の Sheets API を返す関数）
api_factory: Callable[[Callable[[], Session], Any], SheetsApi] = _default_api


def _settings(app_settings: Any):
    if app_settings is not None:
        return app_settings
    from app.config import get_settings

    return get_settings()


# ===== タブの今の形 =====


@dataclass
class TabContext:
    sheet_id: int
    col: dict[str, int]  # 見出し → 0 始まりの列
    id_rows: dict[int, str]  # 1 始まりの行 → REQ_ID（normalize_）
    last_id_row: int
    last_data_row: int
    row_count: int
    col_count: int
    refs: dict[str, dict] | None = None

    def rows_with_id(self, key: str) -> list[int]:
        return sorted(r for r, v in self.id_rows.items() if v == key)


def check_timezone(api: SheetsApi, sid: str, settings) -> dict:
    meta = api.get_spreadsheet(sid)
    tz = spreadsheet_timezone(meta)
    if tz != settings.db_sheet_timezone:
        raise LaneHalt(f"スプレッドシートの時間帯（{tz or '空'}）が DB_SHEET_TIMEZONE（{settings.db_sheet_timezone}）と"
                       "違うため、行の反映を止めました。")
    return meta


def _grid(meta: dict, tab: str) -> dict:
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == tab:
            return s["properties"].get("gridProperties") or {}
    return {}


def load_tab(api: SheetsApi, sid: str, meta: dict, *, need_tail: bool) -> TabContext:
    """見出しを確かめ、REQ_ID の列と最後のデータ行を読む。見出しが欠けていれば LaneHalt。"""
    tab = codec.TAB
    sheet_id = sheet_id_by_title(meta, tab)
    if sheet_id is None:
        raise LaneHalt(f"タブ「{tab}」がありません（タブは作りません）。行の反映を止めました。")
    grid = _grid(meta, tab)
    row_count = int(grid.get("rowCount") or 0)
    col_count = int(grid.get("columnCount") or 26)
    q = quote_sheet(tab)
    header_rows = api.get_values(sid, f"{q}!{codec.HEADER_ROW}:{codec.HEADER_ROW}", "UNFORMATTED_VALUE")
    header = header_rows[0] if header_rows else []
    missing = codec.missing_headers(header)
    if missing:
        raise LaneHalt(f"{tab}: 決まりの見出し {missing} がありません（列は足しません）。行の反映を止めました。")
    col = codec.rule_columns(header)
    start = codec.DATA_START_ROW
    id_col = column_letter(col[codec.ID_RULE.header])
    values = api.get_values(sid, f"{q}!{id_col}{start}:{id_col}", "UNFORMATTED_VALUE")
    id_rows: dict[int, str] = {}
    last_id_row = codec.HEADER_ROW
    for i, row in enumerate(values):
        key = normalize_gas(row[0] if row else "")
        id_rows[start + i] = key
        if key:
            last_id_row = start + i
    last_data_row = last_id_row
    if need_tail and row_count > last_id_row:
        # REQ_ID が空でも、どれかの列に値がある行の上には書かない（GAS の getLastRow と同じ）
        last_col = column_letter(max(col_count, max(col.values()) + 1) - 1)
        tail = api.get_values(sid, f"{q}!A{last_id_row + 1}:{last_col}{row_count}", "UNFORMATTED_VALUE")
        for i, row in enumerate(tail):
            if any(not codec.is_blank(v) for v in row):
                last_data_row = last_id_row + 1 + i
    return TabContext(sheet_id=sheet_id, col=col, id_rows=id_rows, last_id_row=last_id_row,
                      last_data_row=last_data_row, row_count=row_count, col_count=col_count)


def load_refs(api: SheetsApi, sid: str, ctx: TabContext, now: datetime) -> dict[str, dict]:
    """最後のデータ行あたりの、日付・日時の列の表示形式（新しく書くセルに写す）。"""
    if _lane.refs and (now - _lane.refs[0]).total_seconds() < REF_CACHE_SECONDS:
        ctx.refs = _lane.refs[1]
        return ctx.refs
    refs: dict[str, dict] = {}
    if ctx.last_id_row >= codec.DATA_START_ROW:
        lo = max(codec.DATA_START_ROW, ctx.last_id_row - REF_WINDOW + 1)
        last_col = column_letter(max(ctx.col.values()))
        meta = api.get_spreadsheet_row_metadata(
            sid, [f"{quote_sheet(codec.TAB)}!A{lo}:{last_col}{ctx.last_id_row}"], GRID_FIELDS)
        rows = []
        for s in meta.get("sheets", []):
            for d in s.get("data", []):
                rows.extend(r.get("values", []) for r in d.get("rowData", []))
        for r in codec.COLUMN_RULES:
            if not r.date_like:
                continue
            c = ctx.col[r.header]
            found: list[tuple[str, str]] = []
            for vals in reversed(rows):
                cd = vals[c] if c < len(vals) else {}
                uv = cd.get("userEnteredValue") or {}
                nf = (cd.get("userEnteredFormat") or {}).get("numberFormat") or {}
                if "numberValue" in uv and nf.get("type") in codec.DATE_FORMAT_TYPES:
                    found.append((nf.get("type"), nf.get("pattern", "")))
            if not found:
                continue
            if r.header in MAJORITY_FORMAT_HEADERS:
                counts = Counter(found)
                best = max(counts.values())
                pick = next(f for f in found if counts[f] == best)  # 同じ数なら新しい行の形
            else:
                pick = found[0]
            refs[r.header] = {"format": {"type": pick[0], "pattern": pick[1]}}
    _lane.refs = (now, refs)
    ctx.refs = refs
    return refs


# ===== 書く要求 =====


@dataclass
class RowWork:
    pk: str
    version: int
    values: dict  # 見出し → 論理の値（DB の今）
    kinds: dict  # 行の cell_kinds（属性名 → 型）
    synced: dict  # synced_values（KINDS_KEY を含む）
    sheet_row: int | None
    seq: int
    first_dirty_at: datetime | None
    attempts: int
    target_row: int | None = None
    adopted: bool = False
    deferred: bool = False
    written: dict | None = None
    headers: list[str] | None = None
    error: str = ""
    blocked: str = ""

    @property
    def is_new(self) -> bool:
        return self.sheet_row is None and not self.adopted

    @property
    def key(self) -> str:
        return normalize_gas(self.values.get(codec.ID_RULE.header))


def changed_headers(work: RowWork) -> list[str]:
    """書く見出し。新しい行は全列（G・N 列は空で書く）、今ある行は synced_values と違う V 列だけ。"""
    if work.is_new:
        return list(codec.HEADERS)
    synced_kinds = work.synced.get(codec.KINDS_KEY) or {}
    out = []
    for r in codec.V_RULES:
        cur = codec.canonical(r, work.values.get(r.header))
        if (work.adopted or r.header not in work.synced
                or codec.canonical_stored(r, work.synced.get(r.header)) != cur):
            out.append(r.header)
            continue
        if r.date_like and cur != "":
            want, have = work.kinds.get(r.attr), synced_kinds.get(r.attr)
            if want in (codec.SERIAL, codec.STRING) and have in (codec.SERIAL, codec.STRING) and want != have:
                out.append(r.header)
    return out


def cell_requests(ctx: TabContext, row: int, headers: list[str], values: dict, kinds: dict, tz) -> tuple[list, dict]:
    """1 行の書く列 → updateCells の要求（セルごと）と、書いた型（属性名 → 型。空にした列は None）。"""
    reqs, written_kinds = [], {}
    for h in headers:
        r = codec.rule(h)
        value = values.get(h) if r.cls == "V" else ""
        fmt = ((ctx.refs or {}).get(h) or {}).get("format")
        enc = codec.encode_cell(r, value, kinds.get(r.attr), tz, fmt)
        written_kinds[r.attr] = enc["kind"]
        reqs.append({"updateCells": {
            "start": {"sheetId": ctx.sheet_id, "rowIndex": row - 1, "columnIndex": ctx.col[h]},
            "rows": [{"values": [enc["cell"]]}],
            "fields": enc["fields"],
        }})
    return reqs, written_kinds


def _needs_refs(works: list[RowWork]) -> bool:
    for w in works:
        for h in (w.headers or []):
            r = codec.rule(h)
            if r.date_like and r.cls == "V" and not codec.is_blank(w.values.get(h)):
                return True
    return False


# ===== 1 回分 =====


def select_work(db: Session, limit: int, now: datetime) -> tuple[list[RowWork], list[tuple[str, str]]]:
    states = list(db.scalars(select(SheetSyncState).where(
        SheetSyncState.table_name == TABLE,
        SheetSyncState.desired_version > SheetSyncState.synced_version,
        SheetSyncState.state != "blocked")))
    missing: list[tuple[str, str]] = []
    new: list[RowWork] = []
    existing: list[RowWork] = []
    ready: dict[str, bool] = {}
    for st in states:
        row = db.get(LeaveRequest, st.pk)
        if row is None:
            missing.append((st.pk, "DB に行がありません"))
            continue
        work = RowWork(pk=st.pk, version=int(row.row_version), values=codec.db_values(row),
                       kinds=dict(row.cell_kinds or {}), synced=dict(st.synced_values or {}), sheet_row=row.sheet_row,
                       seq=int(st.seq or 0), first_dirty_at=st.first_dirty_at, attempts=int(st.attempts or 0))
        ready[st.pk] = st.next_try_at is None or st.next_try_at <= now
        (new if row.sheet_row is None else existing).append(work)
    chosen: list[RowWork] = []
    for w in sorted(new, key=lambda w: w.seq):
        if not ready[w.pk]:
            break  # 前の新しい行が待ちなら、後ろも待つ（作った順を守る）
        chosen.append(w)
    far = datetime.max.replace(tzinfo=timeutil.UTC)
    existing = [w for w in existing if ready[w.pk]]
    existing.sort(key=lambda w: (w.first_dirty_at or far, w.seq))
    chosen.extend(existing)
    return chosen[:limit], missing


def _claimed_rows(db: Session) -> dict[int, tuple[str, str]]:
    """DB で sheet_row を持つ行（行番号 → (pk, REQ_ID)）。"""
    out: dict[int, tuple[str, str]] = {}
    for req_id, sheet_row in db.execute(select(LeaveRequest.req_id, LeaveRequest.sheet_row)
                                        .where(LeaveRequest.sheet_row.is_not(None))):
        out[int(sheet_row)] = (str(req_id), normalize_gas(req_id))
    return out


def place_rows(ctx: TabContext, works: list[RowWork], claimed: dict[int, tuple[str, str]]) -> None:
    """今ある行の位置を REQ_ID で確かめ（違えば探し直す）、新しい行の位置を決める。"""
    next_row = max(ctx.last_id_row, ctx.last_data_row) + 1
    taken = dict(claimed)

    def owned_by_other(r: int, pk: str) -> bool:
        owner = taken.get(r)
        if owner is None or owner[0] == pk:
            return False
        return ctx.id_rows.get(r, "") == owner[1]

    for w in works:
        key = w.key
        if w.sheet_row is not None:
            if ctx.id_rows.get(w.sheet_row) == key:
                w.target_row = w.sheet_row
                continue
            found = [r for r in ctx.rows_with_id(key) if not owned_by_other(r, w.pk)]
            if not found:
                w.blocked = (f"{codec.TAB} の {w.sheet_row} 行目の REQ_ID が DB と違い、REQ_ID の列にも見つかりません"
                             f"（{key}）。書きません")
                continue
            w.target_row = found[0]
            taken[w.target_row] = (w.pk, key)
            continue
        # 新しい行: 書きかけで落ちた行（同じ REQ_ID で、どの DB の行のものでもない行）があれば、それを使う
        found = [r for r in ctx.rows_with_id(key) if r not in taken]
        if found:
            w.target_row = found[0]
            w.adopted = True
        else:
            w.target_row = next_row
            next_row += 1
        taken[w.target_row] = (w.pk, key)


def _retry_delay(attempts: int) -> int:
    return RETRY_DELAYS[min(attempts, len(RETRY_DELAYS)) - 1] if attempts >= 1 else RETRY_DELAYS[0]


def _record_halt(session_factory, reason: str, recheck_seconds: int = RECHECK_SECONDS) -> None:
    now = timeutil.now_utc()
    first = _lane.halted != reason
    _lane.halted = reason
    _lane.recheck_at = now + timedelta(seconds=recheck_seconds)
    log.error("row_sync halted: %s", reason)
    if first:
        with session_factory() as db:
            db.add(SyncRun(kind=RUN_KIND, started_at=now, finished_at=now, ok=False,
                           summary_json={"halted": True, "reason": reason}))
            db.commit()


def _record_resume(session_factory) -> None:
    now = timeutil.now_utc()
    reason = _lane.halted
    _lane.halted = ""
    _lane.recheck_at = None
    with session_factory() as db:
        db.add(SyncRun(kind=RUN_KIND, started_at=now, finished_at=now, ok=True,
                       summary_json={"resumed": True, "previousReason": reason}))
        db.commit()


def write_rows(api: SheetsApi, sid: str, ctx: TabContext, works: list[RowWork], tz, *, purpose: str) -> None:
    """works（target_row と headers が決まった行）を書く。失敗した行は work.error に理由。UnsafeWriteError は投げる。"""
    plans: list[tuple[RowWork, list[dict], dict]] = []
    for w in works:
        if not w.headers:
            plans.append((w, [], {}))
            continue
        reqs, kinds = cell_requests(ctx, w.target_row, w.headers, w.values, w.kinds, tz)
        plans.append((w, reqs, kinds))
    need = max([w.target_row for w, r, _ in plans if r] + [0])
    if need > ctx.row_count:
        api.append_rows(sid, codec.TAB, need - ctx.row_count, purpose="row_sync")
        ctx.row_count = need
    all_reqs = [r for _, reqs, _ in plans for r in reqs]
    ok_all = True
    if all_reqs:
        try:
            api.batch_update(sid, all_reqs, tab=codec.TAB, purpose=purpose)
        except UnsafeWriteError:
            raise
        except SheetsError as error:
            ok_all = False
            log.warning("row_sync batch failed, retrying row by row: %s", error)
    new_failed = False
    for w, reqs, kinds in plans:
        if not ok_all and reqs:
            if w.is_new and new_failed:
                w.deferred = True
                continue
            try:
                api.batch_update(sid, reqs, tab=codec.TAB, purpose=purpose)
            except UnsafeWriteError:
                raise
            except SheetsError as error:
                w.error = str(error)[:1000]
                new_failed = new_failed or w.is_new
                continue
        snap = {}
        for h in (w.headers or []):
            r = codec.rule(h)
            snap[h] = codec.canonical(r, w.values.get(h) if r.cls == "V" else "")
        w.written = dict(snap)
        w.written[codec.KINDS_KEY] = kinds
        w.synced = merge_synced(w.synced, w.written)


def merge_synced(current: dict | None, written: dict) -> dict:
    """今の synced_values に、書いた列だけを重ねる。"""
    out = dict(current or {})
    kinds = dict(out.get(codec.KINDS_KEY) or {})
    for a, k in (written.get(codec.KINDS_KEY) or {}).items():
        if k is None:
            kinds.pop(a, None)
        else:
            kinds[a] = k
    out.update({h: v for h, v in written.items() if h != codec.KINDS_KEY})
    out[codec.KINDS_KEY] = kinds
    return out


def _save_results(session_factory, works: list[RowWork], missing: list[tuple[str, str]], now: datetime) -> dict:
    counts = {"written": 0, "failed": 0, "blocked": 0}
    with session_factory() as db:
        for pk, reason in missing:
            st = db.get(SheetSyncState, (TABLE, pk))
            if st is not None:
                st.state, st.last_error = "blocked", reason
                counts["blocked"] += 1
        for w in works:
            st = db.get(SheetSyncState, (TABLE, w.pk))
            if st is None or w.deferred:
                continue
            if w.blocked:
                st.state, st.last_error = "blocked", w.blocked[:1000]
                counts["blocked"] += 1
                continue
            if w.error:
                st.attempts = int(st.attempts or 0) + 1
                st.state = "retry"
                st.last_error = w.error
                st.next_try_at = now + timedelta(seconds=_retry_delay(st.attempts))
                counts["failed"] += 1
                continue
            if w.written is None and w.headers is None:
                continue  # 書く所まで進まなかった行（途中で止めた）
            st.synced_version = max(int(st.synced_version or 0), w.version)
            st.synced_values = merge_synced(st.synced_values, w.written or {})
            st.state, st.attempts, st.last_error = "ok", 0, ""
            st.last_synced_at = now
            if st.desired_version <= st.synced_version:
                st.first_dirty_at = None
                st.next_try_at = None
            else:
                st.next_try_at = now
            if w.target_row is not None and w.target_row != w.sheet_row:
                db.execute(update(LeaveRequest).where(LeaveRequest.req_id == w.pk).values(sheet_row=w.target_row))
            counts["written"] += 1
        db.commit()
    return counts


def run_once(session_factory: Callable[[], Session], limit: int = 50, *, api: SheetsApi | None = None,
             app_settings: Any = None) -> dict:
    """行の反映を 1 回。戻り値 {ok, written, failed, blocked, halted?, maxLatencySeconds}。"""
    settings = _settings(app_settings)
    now = timeutil.now_utc()
    if _lane.halted and _lane.recheck_at and now < _lane.recheck_at:
        return {"ok": False, "halted": _lane.halted, "written": 0}
    if _lane.backoff_until and now < _lane.backoff_until:
        return {"ok": False, "error": "一時的な失敗の後の待ち", "written": 0}
    was_halted = bool(_lane.halted)
    with session_factory() as db:
        works, missing = select_work(db, limit, now)
    need_tz = _lane.tz_checked_at is None or (now - _lane.tz_checked_at).total_seconds() >= TZ_CHECK_SECONDS
    if not works and not missing and not need_tz and not was_halted:
        return {"ok": True, "written": 0, "failed": 0, "blocked": 0}
    sid = settings.db_spreadsheet_id
    tz = settings.db_sheet_tz
    result: dict = {"ok": True, "written": 0, "failed": 0, "blocked": 0}
    try:
        if works:
            with session_factory() as db:
                allowed, reason = write_gate.check(db, settings)
            if not allowed:
                raise GateClosed("書き込みの関所: " + reason)
        if api is None:
            api = api_factory(session_factory, settings)
        meta = check_timezone(api, sid, settings)
        _lane.tz_checked_at = now
        if works:
            ctx = load_tab(api, sid, meta, need_tail=any(w.sheet_row is None for w in works))
            with session_factory() as db:
                claimed = _claimed_rows(db)
            place_rows(ctx, works, claimed)
            todo = [w for w in works if not w.blocked]
            for w in todo:
                w.headers = changed_headers(w)
            if _needs_refs(todo):
                load_refs(api, sid, ctx, now)
            write_rows(api, sid, ctx, todo, tz, purpose="row_sync")
    except (LaneHalt, UnsafeWriteError) as halt:
        _record_halt(session_factory, str(halt),
                     GATE_RECHECK_SECONDS if isinstance(halt, GateClosed) else RECHECK_SECONDS)
        return {"ok": False, "halted": str(halt), "written": 0}
    except SheetsError as error:
        log.warning("row_sync: Sheets API failed: %s", error)
        _lane.backoff_until = now + timedelta(seconds=BACKOFF_SECONDS)
        return {"ok": False, "error": str(error)[:500], "written": 0}
    except Exception as error:  # noqa: BLE001 - 委任が無いなど。次の回にもう一度
        log.exception("row_sync failed")
        _lane.backoff_until = now + timedelta(seconds=BACKOFF_SECONDS)
        return {"ok": False, "error": f"{type(error).__name__}: {error}"[:500], "written": 0}
    if was_halted:
        _record_resume(session_factory)
    counts = _save_results(session_factory, works, missing, timeutil.now_utc())
    result.update(counts)
    late = [(now - w.first_dirty_at).total_seconds() for w in works
            if w.first_dirty_at and not (w.error or w.blocked or w.deferred)]
    result["maxLatencySeconds"] = max(late) if late else 0
    return result


# ===== 管理者の欄（ops_status）と手での操作 =====


def status(db: Session) -> dict:
    """反映の待ち・遅れ・止めた行・レーンの停止（管理者の欄）。"""
    now = timeutil.now_utc()
    dirty = SheetSyncState.desired_version > SheetSyncState.synced_version
    pending = int(db.scalar(select(func.count()).select_from(SheetSyncState).where(dirty)) or 0)
    oldest = db.scalar(select(func.min(SheetSyncState.first_dirty_at)).where(dirty))
    blocked = [{"table": s.table_name, "pk": s.pk, "error": s.last_error}
               for s in db.scalars(select(SheetSyncState).where(SheetSyncState.state == "blocked").limit(50))]
    retry = int(db.scalar(select(func.count()).select_from(SheetSyncState).where(SheetSyncState.state == "retry")) or 0)
    last = db.scalars(select(SyncRun).where(SyncRun.kind == RUN_KIND).order_by(SyncRun.id.desc()).limit(1)).first()
    halted = bool(last and not last.ok and (last.summary_json or {}).get("halted"))
    if oldest is not None and oldest.tzinfo is None:
        oldest = oldest.replace(tzinfo=timeutil.UTC)
    lag = (now - oldest).total_seconds() if oldest else 0
    return {"pending": pending, "retry": retry, "blocked": blocked, "oldestDirtySeconds": lag,
            "late": lag > LATE_SECONDS, "halted": halted,
            "haltReason": (last.summary_json or {}).get("reason", "") if halted else ""}


def unblock(db: Session, pk: str) -> bool:
    """blocked の行をもう一度試す（管理者が行を直した後。commit は呼ぶ側）。"""
    st = db.get(SheetSyncState, (TABLE, str(pk)))
    if st is None or st.state != "blocked":
        return False
    st.state, st.attempts, st.last_error = "ok", 0, ""
    st.next_try_at = timeutil.now_utc()
    return True


def reappend_blocked(db: Session, pk: str) -> bool:
    """blocked の行（人がシートの行を消したなど）を、末尾に新しい行として書き直す印にする（commit は呼ぶ側）。"""
    st = db.get(SheetSyncState, (TABLE, str(pk)))
    row = db.get(LeaveRequest, str(pk))
    if st is None or st.state != "blocked" or row is None:
        return False
    row.sheet_row = None
    st.state, st.attempts, st.last_error = "ok", 0, ""
    st.synced_values = {}
    st.next_try_at = timeutil.now_utc()
    db.flush()
    sheet_queue.mark_dirty(db, TABLE, str(pk))
    return True


def main(argv: list[str] | None = None) -> int:  # pragma: no cover - VPS で動かす
    """docker compose exec worker python -m app.services.row_sync status|unblock <REQ_ID>|reappend <REQ_ID>"""
    import argparse
    import json

    from app.db import get_sessionmaker

    parser = argparse.ArgumentParser(prog="python -m app.services.row_sync")
    parser.add_argument("op", choices=["status", "unblock", "reappend"])
    parser.add_argument("pk", nargs="?")
    args = parser.parse_args(argv)
    with get_sessionmaker()() as db:
        if args.op == "status":
            print(json.dumps(status(db), ensure_ascii=False, indent=1, default=str))
            return 0
        if not args.pk:
            parser.error("REQ_ID が要ります")
        done = (unblock if args.op == "unblock" else reappend_blocked)(db, args.pk)
        db.commit()
    print("ok" if done else "blocked の行が見つかりません")
    return 0 if done else 1


def register_worker() -> None:
    """行の反映のレーンは worker が run_once を直接呼ぶ（app.worker._resolve_row_sync）。登録するものは無い。"""


if __name__ == "__main__":  # pragma: no cover
    import sys

    sys.exit(main())


__all__ = ["run_once", "status", "unblock", "reappend_blocked"]
