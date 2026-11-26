"""管理者の「運用」欄の中身（実装仕様 3-11・4-2 の ops_getStatus。担当 M）。

get_ops_status(db) は JSON にできる dict を返す（日時は JST の 'yyyy-MM-dd HH:mm:ss'）。W の ops_getStatus がそのまま返す。
- rowSync: 行の反映の待ち件数・最も古い待ちの経過秒・5 分を超えた遅れの件数・止まっている行（blocked）
- gasJobs: GAS の仕事の状態ごとの件数、失敗の一覧（ops_retryJob で再実行できる id）、前の仕事の失敗で止まっている仕事、
  理由のある待ち（行の反映待ち・gas_ping 待ち・送り直し待ち）、一部の段が失敗した仕事（partial。仕事は done）
- gasPing: 最後の gas_ping の時刻・結果・合わない点・警告・移転モード。lanePaused が真なら GAS のレーンは止まっている
- sheetWriteGate: シートへ書く前の関所（write_gate.status）
- lastRuns: 予定の最後の結果（マスタ同期・設定の取り込み・毎晩の照合・集計タブ・purge）
- exhausted: 再試行を使い切った仕事の記録（新しい順に最大 20）
秘密（token・PNG・メールの宛先）は出さない。
"""
from __future__ import annotations

import json
from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app import jobs, models, timeutil
from app.services import gas_jobs, write_gate

ROW_SYNC_DELAY_ALERT = timedelta(minutes=5)
LAST_RUN_KINDS = ("sync_masters", "sync_sheet_config", "sheet_reconcile", "summary_sheet", "purge")


def _t(value: datetime | None) -> str:
    return timeutil.to_jst_text(value, "yyyy-MM-dd HH:mm:ss") if value else ""


def _plain(value: Any) -> Any:
    if isinstance(value, datetime):
        return _t(value)
    if isinstance(value, dict):
        return {k: _plain(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_plain(v) for v in value]
    return value


def _row_sync(db: Session, now: datetime) -> dict:
    S = models.SheetSyncState
    dirty = S.desired_version > S.synced_version
    pending = int(db.scalar(select(func.count()).select_from(S).where(dirty)) or 0)
    oldest = db.scalar(select(func.min(S.first_dirty_at)).where(dirty))
    oldest_seconds = int((now - oldest).total_seconds()) if oldest else 0
    delayed = int(db.scalar(select(func.count()).select_from(S).where(
        dirty, S.first_dirty_at < now - ROW_SYNC_DELAY_ALERT)) or 0)
    blocked = [{"table": r.table_name, "pk": r.pk, "error": r.last_error or "", "attempts": r.attempts,
                "since": _t(r.first_dirty_at)}
               for r in db.scalars(select(S).where(S.state == "blocked").order_by(S.first_dirty_at).limit(100))]
    out = {"pending": pending, "oldestPendingSeconds": oldest_seconds, "delayed": delayed, "blocked": blocked}
    try:  # D の row_sync.status（見出し・時間帯の違いでレーンが止まっているかなど）があれば添える
        from app.services import row_sync
    except ImportError:
        return out
    status_fn = getattr(row_sync, "status", None)
    if callable(status_fn):
        out["lane"] = _plain(status_fn(db))
    return out


def _job_row(j: dict) -> dict:
    return {"id": j["id"], "reqId": j["reqId"], "action": j["kind"], "reason": j["error"],
            "attempts": j["attempts"], "state": j["state"], "createdAt": _t(j["createdAt"]),
            "runAfter": _t(j["runAfter"])}


def _gas_jobs(db: Session) -> dict:
    J = models.Job
    waiting = [
        {"id": r.id, "reqId": r.request_id, "action": r.kind, "reason": r.last_error, "attempts": r.attempts,
         "runAfter": _t(r.run_after)}
        for r in db.scalars(select(J).where(J.lane == "gas", J.state == "pending", J.last_error != "")
                            .order_by(J.id).limit(100))
    ]
    return {
        "counts": jobs.counts_by_state(db, lane="gas"),
        "failed": [_job_row(j) for j in jobs.list_failed(db, lane="gas")],
        "blocked": [_job_row(j) for j in jobs.list_blocked(db)],
        "waiting": waiting,
        "partial": _partial(db),
    }


def _partial(db: Session, limit: int = 100) -> list[dict]:
    """一部の段（台帳・メール・サイン・PDF の捨て）が失敗した仕事（仕事は done。新しい順）。"""
    J = models.Job
    rows = db.scalars(select(J).where(J.lane == "gas", J.state == "done", J.result_json.like('%"partialErrors"%'))
                      .order_by(J.id.desc()).limit(limit))
    out = []
    for r in rows:
        try:
            errors = json.loads(r.result_json or "{}").get("partialErrors") or []
        except (ValueError, AttributeError):
            errors = []
        out.append({"id": r.id, "reqId": r.request_id, "action": r.kind, "finishedAt": _t(r.finished_at),
                    "errors": [{"step": e.get("step", ""), "error": e.get("error", ""), "message": e.get("message", "")}
                               for e in errors if isinstance(e, dict)]})
    return out


def _gas_ping(db: Session, now: datetime) -> dict:
    row = gas_jobs.last_gas_ping(db)
    ok, reason = gas_jobs.gas_ping_ok(db, now)
    if row is None:
        return {"at": "", "ok": False, "fresh": False, "problems": [reason], "warnings": [], "lanePaused": True,
                "reason": reason, "migrated": None}
    summary = row.summary_json or {}
    return {"at": _t(row.finished_at or row.started_at), "ok": bool(row.ok), "fresh": ok,
            "problems": list(summary.get("problems") or []), "warnings": list(summary.get("warnings") or []),
            "lanePaused": not ok, "reason": reason, "migrated": summary.get("migrated"),
            "auth": summary.get("auth", ""), "apiVersion": summary.get("apiVersion")}


def _last_run(db: Session, kind: str) -> dict | None:
    R = models.SyncRun
    row = db.scalars(select(R).where(R.kind.in_((kind, "schedule:" + kind)))
                     .order_by(R.started_at.desc(), R.id.desc()).limit(1)).first()
    if row is None:
        return None
    return {"kind": row.kind, "startedAt": _t(row.started_at), "finishedAt": _t(row.finished_at), "ok": bool(row.ok),
            "summary": _plain(row.summary_json or {})}


def _exhausted(db: Session) -> list[dict]:
    R = models.SyncRun
    rows = db.scalars(select(R).where(R.kind == gas_jobs.EXHAUSTED_KIND)
                      .order_by(R.started_at.desc(), R.id.desc()).limit(20))
    return [{"at": _t(r.started_at), **_plain(r.summary_json or {})} for r in rows]


def get_ops_status(db: Session) -> dict:
    from app.config import get_settings

    now = timeutil.now_utc()
    return {
        "now": _t(now),
        "rowSync": _row_sync(db, now),
        "gasJobs": _gas_jobs(db),
        "gasPing": _gas_ping(db, now),
        "sheetWriteGate": write_gate.status(db, get_settings(), now),
        "lastRuns": {kind: _last_run(db, kind) for kind in LAST_RUN_KINDS},
        "exhausted": _exhausted(db),
    }
