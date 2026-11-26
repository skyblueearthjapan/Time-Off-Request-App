"""GAS の後処理の仕事（決定 B・実装仕様 7-3・7-4。担当 M）。

W の業務関数は、業務の書き込み（と sheet_queue.mark_dirty）と同じトランザクションで enqueue_* を呼ぶ（commit は
呼び出し側）。worker の GAS のレーン（全体で同時に 1 本。app/jobs.py）が、積んだ順に 1 件ずつ GAS の内部用の入口
（services/gas_client。G の src/Internal.gs）を呼ぶ。

仕事の種類（jobs.kind は GAS の action と同じ名前）:

| kind（7-4）              | 積む関数                  | GAS がすること（7-2）                         | guard で飛ばすか |
|--------------------------|---------------------------|-----------------------------------------------|------------------|
| internal.afterApprove1   | enqueue_after_approve1    | サイン → 承認完了メール → 台帳                | 承認済・approvedAt |
| internal.afterApprove2   | enqueue_after_approve2    | PDF →（trash）→ 台帳                          | 承認済・approvedAt2 |
| internal.afterUnapprove  | enqueue_after_unapprove   | サインの列を空 →（trash）→ 台帳の行を消す     | 飛ばさない |
| internal.sendWarnMail    | enqueue_send_warn_mail    | 集計タブを VPS が書いた後、有給警告メール     | 飛ばさない |

動かす前（7-3）:
  1. requires（積んだ時点の版。その申請の行と、同じ作業員の反映待ちの行）は jobs.JobRunner が待つ。ここでも今の版まで
     反映が済んだかを確かめ、まだなら RetryLater（試行回数は増やさない）。24 時間を過ぎても済まなければ failed。
  2. guard（積んだ時の状態・承認日時）と DB の今の値を比べ、違えば skipped（afterApprove1 は sign_png を消す）。
  3. expect と args は DB の今の値で作る（signBase64 は sign_png から。PNG は jobs.args_json に入れない）。
  4. 最後の gas_ping が 10 分以内に合格していなければ動かさない（1 分後に）。
結果（record_gas_result）: GAS の返事の G 列（サイン画像の URL・PDF の URL とファイル ID）を DB に入れる。G 列は
行の反映の対象外なので mark_dirty も synced_values も触らない（VPS がシートの G 列を上書きしない）。
再試行: retryable:true と通信の失敗・時間切れは jobs の間隔で、同じ jobId（guard_json.gasJobId）で送り直す
（GAS が jobId ごとに済んだ段を覚えていて二重にしない）。入口に入れない失敗（GasAuthError）と retryable:false は
送り直さず failed（管理者の欄。jobs.retry_failed でやり直せる）。failed の仕事が同じ申請の後の仕事を止めるのは、
後の仕事がその段に頼るときだけ（failed_blocks）。
一部の失敗: ok:true の返事の sign・mail・trash・ledger の status:'error' は、仕事を done にしたまま result_json の
partialErrors に残し、運用の欄（ops_status の gasJobs.partial）に出す。
"""
from __future__ import annotations

import base64
import json
import logging
import re
import uuid
from collections.abc import Callable, Mapping
from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import and_, select
from sqlalchemy.orm import Session

from app import jobs, models, safety, sheet_queue, store, timeutil
from app.jobs import JobContext, PermanentJobError, RetryLater, SkipJob

log = logging.getLogger(__name__)

# ===== 定数 =====

KIND_AFTER_APPROVE1 = "internal.afterApprove1"
KIND_AFTER_APPROVE2 = "internal.afterApprove2"
KIND_AFTER_UNAPPROVE = "internal.afterUnapprove"
KIND_SEND_WARN_MAIL = "internal.sendWarnMail"
KINDS = (KIND_AFTER_APPROVE1, KIND_AFTER_APPROVE2, KIND_AFTER_UNAPPROVE, KIND_SEND_WARN_MAIL)
UNAPPROVE_REASONS = ("edit", "cancel")

TABLE = "leave_requests"
STATUS_APPROVED = "承認済"
GAS_PING_KIND = "gas_ping"  # sync_runs.kind（write_gate も読む）
EXHAUSTED_KIND = "gas_job_failed"  # sync_runs.kind（管理者の欄に出す）
GAS_PING_MAX_AGE = timedelta(minutes=10)
GAS_PING_WAIT_SECONDS = 60.0
SYNC_WAIT_SECONDS = 5.0
SYNC_WAIT_DEADLINE = timedelta(hours=24)
DEFAULT_PDF_CLEANUP_MODE = "asis"
_PDF_ID_RE = re.compile(r"/d/([A-Za-z0-9_-]+)")


class GasJobError(RuntimeError):
    """GAS が ok:false・retryable:true を返した。jobs の間隔で同じ jobId で送り直す。"""


# ===== 差し替えの口（試験で偽の GasClient を入れる） =====


def _default_client_factory(settings: Any) -> Any:
    from app.services.gas_client import GasClient

    return GasClient.from_settings(settings)


_client_factory: Callable[[Any], Any] = _default_client_factory


def set_client_factory(factory: Callable[[Any], Any] | None) -> Callable[[Any], Any]:
    """GasClient を作る関数を差し替える（試験）。None で元に戻す。前の関数を返す。"""
    global _client_factory
    previous = _client_factory
    _client_factory = factory or _default_client_factory
    return previous


def _settings(settings: Any = None) -> Any:
    if settings is not None:
        return settings
    from app.config import get_settings

    return get_settings()


def call_gas(settings: Any, action: str, *, job_id: str, request_id: str = "", expect: dict | None = None,
             args: dict | None = None) -> dict:
    """GAS を 1 回呼ぶ（10-2 の GasClient.call。通信の失敗は GasTransportError のまま上げる）。"""
    client = _client_factory(_settings(settings))
    return client.call(action, job_id=job_id, request_id=request_id, expect=expect, args=args)


def _is_auth_error(exc: BaseException) -> bool:
    from app.services.gas_client import GasAuthError

    return isinstance(exc, GasAuthError)


# ===== 小道具 =====


def _dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, default=str, separators=(",", ":"))


def _iso(value: datetime | None) -> str:
    return value.astimezone(timeutil.UTC).isoformat() if value is not None else ""


def _jst(value: datetime | None) -> str:
    return timeutil.to_jst_text(value, "yyyy-MM-dd HH:mm:ss") if value is not None else ""


def _ymd(req: models.LeaveRequest) -> str:
    if req.leave_date is not None:
        return timeutil.to_jst_text(req.leave_date, "yyyy-MM-dd")
    return str(getattr(req, "leave_date_text", "") or "").strip()


def _pdf_cleanup_mode() -> str:
    """quirks.PDF_CLEANUP_MODE（W。読むだけ）。まだ無ければ既定 'asis'。"""
    try:
        from app.domain import quirks
    except ImportError:
        return DEFAULT_PDF_CLEANUP_MODE
    mode = str(getattr(quirks, "PDF_CLEANUP_MODE", DEFAULT_PDF_CLEANUP_MODE) or DEFAULT_PDF_CLEANUP_MODE)
    return mode if mode in ("asis", "trash") else DEFAULT_PDF_CLEANUP_MODE


def _fiscal_year_today() -> str:
    """JST の今日の年度（domain.fiscal.fiscal_year。まだ無ければ 3/16 始まりの年）。"""
    today = timeutil.today_jst()
    try:
        from app.domain import fiscal

        return str(fiscal.fiscal_year(today))
    except (ImportError, AttributeError):
        return str(timeutil.fy_start_0316(today).year)


def pdf_file_id_from_url(url: str) -> str:
    match = _PDF_ID_RE.search(url or "")
    return match.group(1) if match else ""


def _scrub(value: Any) -> Any:
    """jobs.result_json に入れる前に、PNG（base64）を落とす。"""
    if isinstance(value, Mapping):
        return {k: _scrub(v) for k, v in value.items() if k not in ("signBase64", "signPng", "png")}
    if isinstance(value, list):
        return [_scrub(v) for v in value]
    if isinstance(value, str) and value.startswith("data:image/"):
        return "(省略)"
    return value


# ===== 行の反映の待ち（7-3 の 1） =====


def _dirty_rows_same_worker(db: Session, req: models.LeaveRequest) -> list[str]:
    """同じ作業員 ID の、反映待ち（desired > synced）の行の REQ_ID（台帳の有給累計が全行を数えるため）。"""
    worker_id = str(req.worker_id or "").strip()
    if not worker_id:
        return []
    L, S = models.LeaveRequest, models.SheetSyncState
    q = (select(L.req_id).join(S, and_(S.table_name == TABLE, S.pk == L.req_id))
         .where(L.worker_id == req.worker_id, L.req_id != req.req_id, S.desired_version > S.synced_version)
         .order_by(L.req_id))
    return list(db.scalars(q))


def requires_for_request(db: Session, req: models.LeaveRequest) -> list[dict]:
    """requires（10-2 の形）: その申請の行（印がある＝シートへ書く行のとき）と、同じ作業員の反映待ちの行。"""
    rows: list[tuple[str, str]] = []
    if db.get(models.SheetSyncState, (TABLE, req.req_id)) is not None:
        rows.append((TABLE, req.req_id))
    rows += [(TABLE, pk) for pk in _dirty_rows_same_worker(db, req)]
    return sheet_queue.requires_for(db, *rows)


def unsynced_rows(db: Session, req: models.LeaveRequest) -> list[str]:
    """今の版まで反映が済んでいない行（その申請と同じ作業員の行）。空なら済んでいる。"""
    pending = []
    state = db.get(models.SheetSyncState, (TABLE, req.req_id))
    if state is not None and int(state.synced_version) < int(req.row_version or 0):
        pending.append(req.req_id)
    return pending + _dirty_rows_same_worker(db, req)


def _require_synced(db: Session, req: models.LeaveRequest, ctx: JobContext) -> None:
    pending = unsynced_rows(db, req)
    if not pending:
        return
    reason = "row_not_synced（行の反映待ち）: " + ", ".join(pending[:20])
    if ctx.created_at is not None and ctx.now - ctx.created_at > SYNC_WAIT_DEADLINE:
        raise PermanentJobError(reason + "（24 時間を過ぎました）")
    raise RetryLater(SYNC_WAIT_SECONDS, reason)


# ===== gas_ping の確かめ（7-3 の 4） =====


def last_gas_ping(db: Session) -> models.SyncRun | None:
    R = models.SyncRun
    return db.scalars(select(R).where(R.kind == GAS_PING_KIND)
                      .order_by(R.started_at.desc(), R.id.desc()).limit(1)).first()


def gas_ping_ok(db: Session, now: datetime | None = None) -> tuple[bool, str]:
    """最後の gas_ping が 10 分以内に合格しているか。"""
    now = now or timeutil.now_utc()
    row = last_gas_ping(db)
    if row is None:
        return False, "gas_ping がまだありません"
    if not row.ok:
        problems = (row.summary_json or {}).get("problems") or []
        return False, "gas_ping が不合格です: " + " / ".join(str(p) for p in problems)[:500]
    at = row.finished_at or row.started_at
    if now - at > GAS_PING_MAX_AGE:
        return False, "gas_ping が 10 分以上ありません"
    return True, ""


def _require_gas_ping(db: Session, now: datetime) -> None:
    ok, reason = gas_ping_ok(db, now)
    if not ok:
        raise RetryLater(GAS_PING_WAIT_SECONDS, reason)


# ===== 結果の記録（7-3） =====


def _later_sign_job_pending(db: Session, req_id: str) -> bool:
    """同じ申請に、まだ動いていない afterApprove1 がある（承認のやり直しで新しい sign_png が入っている）。"""
    J = models.Job
    return db.scalar(select(J.id).where(J.kind == KIND_AFTER_APPROVE1, J.request_id == req_id,
                                        J.state == "pending").limit(1)) is not None


def _clear_sign_png(db: Session, req: models.LeaveRequest) -> None:
    if req.sign_png is not None and not _later_sign_job_pending(db, req.req_id):
        req.sign_png = None


def record_gas_result(db: Session, req_id: str, resp: dict) -> None:
    """GAS の返事の G 列を DB に入れる（同じトランザクション。commit しない。10-2）。

    - sign.signImageUrl があれば（空でも）sign_image_url に入れ、sign_png を消す（後に別の afterApprove1 が
      待っているときは、その仕事が使う新しい PNG なので消さない）。
    - pdf.status が created／already なら pdf_url・pdf_file_id を入れる（ID が無ければ URL の /d/<id>/ から）。
    - trash が trashed で afterUnapprove なら pdf_url・pdf_file_id を空に（GAS がシートの列も空にした）。
    - G 列は行の反映の対象外。mark_dirty も synced_values も触らない（VPS がシートの G 列を上書きしない）。
    """
    if not isinstance(resp, Mapping) or not req_id:
        return
    req = store.lock_request(db, req_id)
    if req is None:
        return
    sign = resp.get("sign")
    if isinstance(sign, Mapping) and "signImageUrl" in sign:
        req.sign_image_url = str(sign.get("signImageUrl") or "")
        _clear_sign_png(db, req)
    pdf = resp.get("pdf")
    if isinstance(pdf, Mapping) and pdf.get("status") in ("created", "already"):
        url = str(pdf.get("pdfUrl") or "")
        if url:
            req.pdf_url = url
            req.pdf_file_id = str(pdf.get("pdfFileId") or "") or pdf_file_id_from_url(url)
    trash = resp.get("trash")
    if (isinstance(trash, Mapping) and trash.get("status") == "trashed"
            and resp.get("action") == KIND_AFTER_UNAPPROVE):
        req.pdf_url = ""
        req.pdf_file_id = ""
    db.flush()


# ===== 実行 =====


def _guard_matches(kind: str, req: models.LeaveRequest, guard: Mapping[str, Any]) -> bool:
    if kind == KIND_AFTER_APPROVE1:
        return req.status == STATUS_APPROVED and _iso(req.approved_at) == str(guard.get("approvedAt") or "")
    if kind == KIND_AFTER_APPROVE2:
        return req.status == STATUS_APPROVED and _iso(req.approved_at2) == str(guard.get("approvedAt2") or "")
    return True  # afterUnapprove は飛ばさない


def build_call(kind: str, req: models.LeaveRequest | None, args: Mapping[str, Any]) -> tuple[dict | None, dict]:
    """GAS へ送る expect と args（7-2。DB の今の値で作る。signBase64 はここで初めて作る）。"""
    if kind == KIND_SEND_WARN_MAIL or req is None:
        return None, {"fiscalYear": str(args.get("fiscalYear") or _fiscal_year_today())}
    if kind == KIND_AFTER_APPROVE1:
        png = req.sign_png
        sign = "data:image/png;base64," + base64.b64encode(png).decode("ascii") if png else ""
        return ({"status": req.status, "approvedAt": _jst(req.approved_at), "approvedBy1Email": req.approved_by1_email},
                {"signBase64": sign, "leaveDate": _ymd(req)})
    if kind == KIND_AFTER_APPROVE2:
        return ({"status": req.status, "approvedAt2": _jst(req.approved_at2),
                 "approvedBy2Email": req.approved_by2_email},
                {"approver1Email": req.approved_by1_email, "approver2Email": req.approved_by2_email,
                 "oldPdfUrl": str(args.get("oldPdfUrl") or ""), "oldPdfFileId": str(args.get("oldPdfFileId") or ""),
                 "cleanupMode": _pdf_cleanup_mode()})
    # afterUnapprove: 状態は今の DB の値（編集の直後に承認し直したなら承認済。台帳の行を消した後で afterApprove1 が書く）
    return ({"status": req.status},
            {"reason": str(args.get("reason") or ""), "oldLeaveDate": str(args.get("oldLeaveDate") or ""),
             "clearSign": bool(args.get("clearSign")), "oldPdfUrl": str(args.get("oldPdfUrl") or ""),
             "oldPdfFileId": str(args.get("oldPdfFileId") or ""), "cleanupMode": _pdf_cleanup_mode()})


def _gas_job_id(ctx: JobContext) -> str:
    gid = str((ctx.guard or {}).get("gasJobId") or "")
    return gid or f"kyuka-job-{ctx.job_id}"


def _store_job_result(db: Session, job_id: int, result: Mapping[str, Any]) -> None:
    row = db.get(models.Job, job_id)
    if row is not None:
        row.result_json = _dumps(_scrub(dict(result)))


def _error_text(resp: Mapping[str, Any]) -> str:
    code = str(resp.get("error") or "error")
    message = str(resp.get("message") or "")
    return f"{code}: {message}" if message and message != code else code


def _record_ping_failure(db: Session, message: str) -> None:
    """入口に入れない（認証）: gas_ping を不合格にして GAS のレーン全体を止める（次の合格まで待つ）。"""
    now = timeutil.now_utc()
    db.add(models.SyncRun(kind=GAS_PING_KIND, started_at=now, finished_at=now, ok=False,
                          summary_json={"problems": [message[:500]], "source": "job"}))


def _write_summary_sheet(db: Session) -> dict:
    """有給警告メールの前に集計タブを同期で書く（D の summary_sheet.write）。"""
    from app.services import summary_sheet

    result = summary_sheet.write(db)
    if isinstance(result, Mapping) and result.get("ok") is False:
        raise RuntimeError(f"集計タブを書けませんでした: {result.get('error') or result.get('reason') or ''}")
    return dict(result) if isinstance(result, Mapping) else {}


PARTIAL_STEPS = ("sign", "mail", "trash", "ledger")


def partial_errors(resp: Mapping[str, Any]) -> list[dict]:
    """ok:true の返事の中で失敗した段（sign・mail・trash・ledger の status:'error'）。

    GAS は台帳・メール・サイン・PDF の捨てだけの失敗を、全体を ok:true のまま段の status:'error' で返す
    （今の GAS が握りつぶして続けるのと同じ）。仕事は done にし、運用の欄（「一部が失敗した仕事」）に出す。
    """
    out = []
    for step in PARTIAL_STEPS:
        part = resp.get(step)
        if isinstance(part, Mapping) and part.get("status") == "error":
            out.append({"step": step, "error": str(part.get("error") or "")[:300],
                        "message": str(part.get("message") or "")[:500]})
    return out


def failed_blocks(db: Session, failed: models.Job, later: models.Job) -> bool:
    """同じ申請の先の仕事が failed のとき、後の仕事を止めるか（jobs.register_kind の failed_blocks）。

    止めるのは、後の仕事が先の仕事の段に頼るときだけ（ほかは進める。やり直さない失敗で後の仕事を永久に止めない）:
    - 先が afterUnapprove: 止める。台帳の行を消す・サインの列を空にする段の後に、承認し直しの段が書く順番が要る。
      afterUnapprove には guard が無いので、後の仕事を先に進めてから再実行すると新しいサイン・台帳の行を消してしまう。
    - 先が afterApprove1・後が afterApprove2: サインの段が済んでいない（sign_png が残っている）なら止める（PDF に
      手書きサインが入らない）。サインの段が済んだ（record_gas_result が sign_png を消した）・電子印なら進める。
    - ほか（先が afterApprove1・afterApprove2）は進める。先の仕事を後で再実行しても、guard（状態・承認日時）が
      今の値と違えば skipped になり、後の仕事の結果を上書きしない。行の反映は後の仕事の requires が待つ。
    failed の仕事は運用の欄（gasJobs.failed）に出て、ops_retryJob でやり直せる。
    """
    if failed.kind == KIND_AFTER_UNAPPROVE:
        return True
    if failed.kind == KIND_AFTER_APPROVE1 and later.kind == KIND_AFTER_APPROVE2:
        req = db.get(models.LeaveRequest, later.request_id)
        return req is not None and req.sign_png is not None
    return False


def handle_job(ctx: JobContext) -> dict:
    """GAS の仕事 1 件（jobs.register_kind のハンドラー）。"""
    db = ctx.db
    req: models.LeaveRequest | None = None
    if ctx.kind != KIND_SEND_WARN_MAIL:
        req = db.get(models.LeaveRequest, ctx.request_id)
        if req is None:
            raise SkipJob("request_not_found")
        if not _guard_matches(ctx.kind, req, ctx.guard):
            if ctx.kind == KIND_AFTER_APPROVE1:
                _clear_sign_png(db, req)
                ctx.commit()  # SkipJob で戻されないよう先に確定する
            raise SkipJob("guard_mismatch（承認の後に状態・承認日時が変わった）")
        _require_synced(db, req, ctx)
    _require_gas_ping(db, ctx.now)
    if ctx.kind == KIND_SEND_WARN_MAIL:
        _write_summary_sheet(db)
        ctx.commit()
    expect, gas_args = build_call(ctx.kind, req, ctx.args)
    try:
        resp = call_gas(ctx.app_settings, ctx.kind, job_id=_gas_job_id(ctx), request_id=ctx.request_id,
                        expect=expect, args=gas_args)
    except Exception as exc:
        if _is_auth_error(exc):
            db.rollback()
            _record_ping_failure(db, f"GAS の入口に入れません: {exc}")
            ctx.commit()
            raise PermanentJobError(f"auth: {exc}") from exc
        raise  # 結果が分からない（時間切れ・通信）: 同じ jobId で送り直す
    resp = dict(resp) if isinstance(resp, Mapping) else {"ok": False, "error": "bad_response", "retryable": True}
    if req is not None:
        record_for_db = dict(resp)
        record_for_db.setdefault("action", ctx.kind)
        if (ctx.kind == KIND_AFTER_UNAPPROVE and gas_args.get("clearSign") and resp.get("ok") is True
                and not isinstance(resp.get("sign"), Mapping)):
            record_for_db["sign"] = {"status": "cleared", "signImageUrl": ""}
        record_gas_result(db, req.req_id, record_for_db)
    if resp.get("ok") is True:
        result = _scrub(resp)
        partial = partial_errors(resp)
        if partial:  # 仕事は done（後の仕事は止めない）。result_json に残し、運用の欄に出す
            result["partialErrors"] = partial
            log.warning("GAS の仕事の一部の段が失敗しました: job=%s kind=%s request=%s %s", ctx.job_id, ctx.kind,
                        ctx.request_id, partial)
        return result
    _store_job_result(db, ctx.job_id, resp)
    ctx.commit()  # 済んだ段の結果と返事を残してから失敗にする
    if resp.get("retryable") is True:
        raise GasJobError(_error_text(resp))
    raise PermanentJobError(_error_text(resp))


def on_exhausted(ctx: JobContext, message: str) -> None:
    """failed になった仕事を管理者の欄に出す（sync_runs kind='gas_job_failed'）。"""
    now = timeutil.now_utc()
    ctx.db.add(models.SyncRun(kind=EXHAUSTED_KIND, started_at=now, finished_at=now, ok=False, summary_json={
        "jobId": ctx.job_id, "kind": ctx.kind, "requestId": ctx.request_id, "attempts": ctx.attempt,
        "error": message[:1000]}))
    log.error("GAS の仕事が失敗のまま止まりました: job=%s kind=%s request=%s %s", ctx.job_id, ctx.kind,
              ctx.request_id, message)


# ===== 積む（10-2。W が業務のトランザクションの中で呼ぶ。commit しない） =====


def new_gas_job_id() -> str:
    """GAS へ送る jobId（積んだときに決め、送り直しでも同じ）。"""
    return str(uuid.uuid4())


def _request_or_error(db: Session, req_id: str) -> models.LeaveRequest:
    req = db.get(models.LeaveRequest, req_id)
    if req is None:
        raise ValueError(f"申請がありません: {req_id}")
    return req


def _guard(req: models.LeaveRequest) -> dict:
    return {"gasJobId": new_gas_job_id(), "status": req.status, "approvedAt": _iso(req.approved_at),
            "approvedAt2": _iso(req.approved_at2)}


def _enqueue(db: Session, kind: str, req: models.LeaveRequest, args: dict) -> int:
    return jobs.enqueue(db, kind, args, lane="gas", request_id=req.req_id,
                        requires=requires_for_request(db, req), guard=_guard(req))


def enqueue_after_approve1(db: Session, req_id: str) -> int:
    """1 次承認の後（サイン → メール → 台帳）。W は先に状態・承認日時・sign_png を書き、mark_dirty しておく。"""
    return _enqueue(db, KIND_AFTER_APPROVE1, _request_or_error(db, req_id), {})


def enqueue_after_approve2(db: Session, req_id: str) -> int:
    """2 次承認の後（PDF →（trash）→ 台帳）。前の PDF（oldPdfUrl・oldPdfFileId）は積んだ時の DB の値。"""
    req = _request_or_error(db, req_id)
    return _enqueue(db, KIND_AFTER_APPROVE2, req, {"oldPdfUrl": req.pdf_url or "", "oldPdfFileId": req.pdf_file_id or ""})


def enqueue_after_unapprove(db: Session, req_id: str, *, reason: str, old_leave_date: str, old_pdf_url: str,
                            old_pdf_file_id: str, clear_sign: bool) -> int:
    """編集で承認が外れた後（reason='edit'）・取消の後（'cancel'）。old_* は操作の前の値。"""
    if reason not in UNAPPROVE_REASONS:
        raise ValueError(f"reason は {UNAPPROVE_REASONS} のどれかです: {reason}")
    req = _request_or_error(db, req_id)
    args = {"reason": reason, "oldLeaveDate": old_leave_date or "", "oldPdfUrl": old_pdf_url or "",
            "oldPdfFileId": old_pdf_file_id or "", "clearSign": bool(clear_sign)}
    return _enqueue(db, KIND_AFTER_UNAPPROVE, req, args)


def enqueue_send_warn_mail(db: Session) -> int:
    """有給警告メール（運用の欄）。待ちのものがあればそれを返す（二重押しで 2 通にしない）。"""
    J = models.Job
    existing = db.scalar(select(J.id).where(J.kind == KIND_SEND_WARN_MAIL, J.state.in_(("pending", "running")))
                         .order_by(J.id).limit(1))
    if existing is not None:
        return int(existing)
    return jobs.enqueue(db, KIND_SEND_WARN_MAIL, {}, lane="gas", guard={"gasJobId": new_gas_job_id()})


# ===== gas_ping（5 分ごと・起動時。7-3・9-1） =====


def run_gas_ping(session_factory: Callable[[], Session], settings: Any = None) -> dict:
    """internal.ping を呼び、safety.verify_gas_ping の結果を sync_runs（kind='gas_ping'）に残す。

    write_gate（シートへ書く前の関所）はここで残した最後の合格の行の migrationMode を読む。
    """
    s = _settings(settings)
    started = timeutil.now_utc()
    summary: dict[str, Any]
    try:
        ping = call_gas(s, "internal.ping", job_id="ping-" + uuid.uuid4().hex)
        problems = safety.verify_gas_ping(ping, s)
        gas_settings = ping.get("settings") if isinstance(ping.get("settings"), Mapping) else {}
        migrated = ping.get("migrated", ping.get("migrationMode"))
        warnings = []
        if int(ping.get("ledgerKeyRows") or 0) > 1:
            warnings.append("M_SYSTEM_SETTING に LEDGER_SSID の行が 2 つ以上あります（下の行が使われます）")
        if ping.get("ledgerOpenable") is False:
            warnings.append("GAS から台帳を開けません")
        if ping.get("pdfFolderOpenable") is False:
            warnings.append("GAS から PDF フォルダを開けません")
        summary = {"problems": problems, "warnings": warnings, "migrated": migrated, "migrationMode": migrated,
                   "apiVersion": ping.get("apiVersion"), "auth": ping.get("auth") or "",
                   "spreadsheetTimeZone": ping.get("spreadsheetTimeZone") or "",
                   "mailToSet": gas_settings.get("MAIL_TO_SET")}
    except Exception as exc:  # noqa: BLE001 - 届かない・設定の誤りも不合格として残す
        problems = [f"GAS に届きませんでした: {type(exc).__name__}: {exc}"[:500]]
        summary = {"problems": problems}
    ok = not problems
    with session_factory() as db:
        db.add(models.SyncRun(kind=GAS_PING_KIND, started_at=started, finished_at=timeutil.now_utc(), ok=ok,
                              summary_json=summary))
        db.commit()
    if not ok:
        log.error("gas_ping 不合格（GAS のレーンを止めます）: %s", problems)
    return {"ok": ok, **summary}


# ===== 登録（worker.load_registrations が呼ぶ） =====


def register_worker() -> None:
    from app import worker

    for kind in KINDS:
        jobs.register_kind(kind, handle_job, lane="gas", on_exhausted=on_exhausted, failed_blocks=failed_blocks)
    worker.register_schedule("gas_ping", task=run_gas_ping, every_minutes=5, run_at_startup=True)


__all__ = [
    "GasJobError", "KINDS", "build_call", "enqueue_after_approve1", "enqueue_after_approve2",
    "enqueue_after_unapprove", "enqueue_send_warn_mail", "gas_ping_ok", "handle_job", "last_gas_ping",
    "record_gas_result", "register_worker", "requires_for_request", "run_gas_ping", "set_client_factory",
]
