"""受け入れテスト（ステージングの API を HTTP で呼んで主な流れを確かめる。実装仕様 12 章・docs/acceptance-checklist.md）。

項目ごとに PASS／FAIL／SKIP を表（項目・期待・実際）で出し、FAIL が 1 つでもあれば終了コード 1。
**ステージングでだけ動く**: 最初に /healthz を読み、env が staging でなければ何もせずに終了コード 2（本番では動かない）。

動かし方は 2 通り:

1) api コンテナと同じネットワークの中（おすすめ。DB・シート・GAS の仕事も確かめる）

   sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app \
       -v /opt/kyuka/tools:/app/tools:ro api \
       python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session [--write]"

   --mint-session: staging_session.py で BOOTSTRAP_ADMINS の 1 人目のセッションと、権限の無い架空のメールの
   セッション（権限の試験用）を作り、最後に必ず消す（staging でだけ作れる）。DB とシート（委任トークンで読むだけ）
   を見る probe も付く。

2) 外から（クッキーを渡す。DB・シート・GAS の仕事の項目は SKIP）

   python tools/qa/acceptance.py --base-url https://leave.lineworks-local.net --cookie "<値>" \
       [--outsider-cookie "<値>"] [--write]

   クッキーの値は staging_session.py mint（--outsider）で作り、終わったら revoke する。値は表示しない。

書き込みの試験（--write）の流れ:
  境目の日付（何も書かない 200 {ok:false}）→ 申請 A 全日・有給（同じ冪等キーで 2 回送って 1 件）、B 半日・特別、
  C 振替、D 過去申請（氏名に XSS の文字）→ プレビュー → 1 次承認（A は手書きサインの小さな PNG、B は電子印）
  → 承認済みをもう一度承認（K-4 の文で拒む）→ D を取消して承認（拒む）→ 2 次承認 A → もう一度（拒む）
  → B を編集して承認が外れる → 全部取消。
  各段で、DB の状態・シートの行（統合業務メールと同じ条件: STATUS の文字、休暇日が日付型、APPROVED_AT2 が空か）・
  GAS の仕事（サイン画像の URL・PDF の URL・台帳・承認完了メール）が済むまで待って確かめる。
  作る申請の作業員 ID は QA-<実行 ID>、作業員名と有給詳細は「【受入テスト】<実行 ID>」で始まる。最後に全部を取消し、
  申請 ID の一覧を出す。承認完了メールはステージングの MAIL_TO（オーナーの 1 件）にだけ届き、1 次承認の 2 回だけ出る
  （件数を報告に出す）。有給警告メール（ops_sendWarnMail）は呼ばない（JEV で確かめる）。

書き込みをしないとき（既定）:
  読む関数の形、存在しない reqId、権限（総務でないアカウントの 2 次承認・ops_* の拒否）、境目の入力の誤りだけを確かめる。
  書く関数は「存在しない申請 ID」や「入力の誤り」で呼び、GAS と同じ文で止まることを見る（申請の行は作らない）。
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import statistics
import struct
import sys
import time
import uuid
import zlib
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable, Iterable

PASS, FAIL, SKIP = "PASS", "FAIL", "SKIP"
TEST_MARK = "【受入テスト】"
JST = timezone(timedelta(hours=9))
DEFAULT_ORIGIN = "https://leave.lineworks-local.net"
SESSION_COOKIE = "__Host-kyuka_session"
USER_AGENT = "kyuka-acceptance/1"
BOGUS_ID = "LV-00000000-999"
XSS_TEXT = "<img src=x onerror=alert(1)>&\"'</script>"

# 閉じたリスト（app/rpc.py と同じ。test_qa_tools.py が食い違いを止める）
SCREEN_FUNCTIONS: tuple[str, ...] = (
    "api_isDeptApprover", "api_isAdmin", "api_isSomu", "api_getDeptList", "api_getWorkerInfo", "api_getLookupData",
    "api_getWorkingDays", "api_getWorkersByDept", "api_getUpcomingLeaves", "api_getHolidayDays",
    "api_getPastWorkingDays", "api_submitLeaveRequest", "api_getAllPendingRequests", "api_editLeaveRequest",
    "api_deleteLeaveRequest", "api_getLeaveRequestById", "api_getApproverList", "api_approveLeaveRequest",
    "api_getCurrentApproverInfo", "api_getPending2ndApprovals", "api_approveLeaveRequest2", "api_getSomuAdminView",
    "api_getApproverProfiles", "api_getDeptAdminView", "api_getAdminSomuView", "api_getAdminBossView",
    "api_getLeaveStatusSearch",
)
OPS_FUNCTIONS: tuple[str, ...] = (
    "ops_getStatus", "ops_retryJob", "ops_syncSheetConfigNow", "ops_runMasterSync", "ops_rebuildSummaryNow",
    "ops_sendWarnMail", "ops_revokeSessions",
)
RPC_FUNCTIONS: tuple[str, ...] = SCREEN_FUNCTIONS + OPS_FUNCTIONS
# 書く関数（app/domain/register.py の WRITES と同じ。冪等キーを付ける）
WRITE_FUNCTIONS = frozenset({
    "api_submitLeaveRequest", "api_editLeaveRequest", "api_deleteLeaveRequest", "api_approveLeaveRequest",
    "api_approveLeaveRequest2", "ops_retryJob", "ops_syncSheetConfigNow", "ops_runMasterSync",
    "ops_rebuildSummaryNow", "ops_sendWarnMail", "ops_revokeSessions",
})

# GAS と同じ文（app/domain の原文。test_qa_tools.py が domain のソースにあることを確かめる）
MSG = {
    "already_approved": "この申請はすでに承認済みです。",
    "canceled_approve": "この申請は取消済みです。",
    "already_approved2": "この申請はすでに2次承認済みです。",
    "not_approved1": "1次承認されていない申請です。",
    "already_canceled": "既に取消済みです。",
    "not_found": "申請が見つかりません。",
    "not_found_prefix": "申請が見つかりません: ",
    "no_req_id": "REQ_IDが指定されていません。",
    "somu_only": "権限がありません（総務または管理者のみ）",
    "admin_page_only": "権限がありません",
    "dept_approver_only": "部署承認者の権限がありません。",
    "bad_format": "休暇日の形式が不正です（yyyy-MM-dd）。",
    "not_working_day": "指定された休暇日は勤務日ではありません。過去日の場合は「過去申請」から選択してください。",
    "retro_out_of_fy": "過去日申請は今年度内（年度開始の3/16から本日まで）の日付のみ指定できます。",
    "substitute_missing": "振替休暇の場合は振替元出勤日を選択してください。",
    "substitute_period": "振替休暇は同一月次期間（前月16日〜当月15日）内で消化する必要があります。",
    "bad_job_id": "仕事の番号が正しくありません。",
    "no_email": "メールアドレスを指定してください。",
}
# rpc.py の入口の文
RPC_MSG = {
    "login": "ログインが必要です。",
    "shape": "呼び出しの形が正しくありません。",
    "origin": "この画面からは操作できません。",
    "unknown": "その操作はありません。",
    "not_operator": "管理者権限がありません。",
}

# 状態・区分（app/constants.py）
ST_SUBMITTED, ST_APPROVED, ST_CANCELED = "申請中", "承認済", "取消"
DAY_FULL, DAY_HALF, HALF_AM, HALF_PM = "全日休", "半日休", "午前", "午後"
KIND_PAID, KIND_SUBSTITUTE, KIND_SPECIAL = "有給消化（私用）", "振替休暇", "特別休暇"

# GAS の仕事（services/gas_jobs.py の jobs.kind）と、GAS の返事の段の状態
K_APPROVE1 = "internal.afterApprove1"
K_APPROVE2 = "internal.afterApprove2"
K_UNAPPROVE = "internal.afterUnapprove"
JOB_FINISHED = ("done", "skipped", "failed")

# 画面の題名（templates/pages）と、権限の無い人に出す文（app/pages.py の NO_AUTH）
PAGE_TITLES = {
    "top": "休暇届 申請TOP", "preview": "休暇届 プレビュー", "sign": "休暇届 サイン", "deptAdmin": "休暇届 部署管理",
    "somuAdmin": "休暇届 総務管理", "admin": "休暇届 管理", "leaveStatus": "休暇届 取得状況",
}
NO_AUTH_TITLES = {"admin": "アクセス権限エラー", "deptAdmin": "アクセス権限なし", "somuAdmin": "アクセス権限なし"}
SOMU_NO_AUTH_TEXT = "このページは総務部・管理者のみアクセスできます。"

# nginx のヘッダー（web/nginx.conf）
CSP_EXPECTED = ("default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; "
                "connect-src 'self'; frame-ancestors 'self'; base-uri 'self'; form-action 'self'")
HEADERS_EXPECTED = {
    "x-content-type-options": "nosniff",
    "x-frame-options": "SAMEORIGIN",
    "referrer-policy": "strict-origin-when-cross-origin",
}
HEADERS_PRESENT = ("strict-transport-security", "permissions-policy")

YMD_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
SLASH_YMD_RE = re.compile(r"^\d{4}/\d{2}/\d{2}$")
REQ_ID_RE = re.compile(r"^LV-\d{8}-\d{3,}$")
INLINE_SCRIPT_RE = re.compile(r"<script(?![^>]*\bsrc=)(?![^>]*application/json)[^>]*>", re.I)
ON_ATTR_RE = re.compile(r"<[^>]+\son[a-z]+\s*=", re.I)

UPCOMING_KEYS = ("reqId", "deptName", "workerName", "leaveKubun", "halfDayType", "leaveDate", "leaveType", "status",
                 "isRetroactive", "daysAgo")
PENDING_KEYS = UPCOMING_KEYS + ("submittedAt",)
PENDING2_KEYS = ("reqId", "deptName", "workerName", "leaveDate", "leaveKubun", "leaveType", "approvedAt",
                 "isRetroactive", "daysAgo")
WORKER_VIEW_KEYS = ("deptName", "workerId", "workerName", "paidCount", "paidHalfCount", "substituteCount",
                    "specialCount", "totalCount", "lastPaidDate", "warnLevel", "quarters", "currentShortage")
SUMMARY_VIEW_KEYS = ("deptId", "deptName", "workerId", "workerName", "paidCount", "paidHalfCount", "substituteCount",
                     "specialCount", "totalCount", "lastPaidDate", "warnLevel")
REQUEST_KEYS = ("reqId", "fiscalYear", "deptId", "deptName", "workerId", "workerName", "leaveKubun", "halfDayType",
                "leaveDate", "leaveType", "substituteDate", "specialReason", "paidDetail", "additionalDetail",
                "submittedAt", "approvedAt", "status", "signImageUrl", "pdfUrl", "approvedBy1Email", "approvedBy2",
                "approvedBy2Email", "approvedAt2", "rowNo")


# ===== 日付の小道具（JST） =====


def today_jst(now: datetime | None = None) -> date:
    return (now or datetime.now(JST)).astimezone(JST).date()


def fiscal_year(d: date) -> int:
    """年度（3/16 始まり）。"""
    return d.year if (d.month, d.day) >= (3, 16) else d.year - 1


def fy_start(d: date) -> date:
    return date(fiscal_year(d), 3, 16)


def monthly_period(d: date) -> tuple[date, date]:
    """月次期間（前月 16 日〜当月 15 日）の (始め, 終わり)。"""
    if d.day >= 16:
        start = date(d.year, d.month, 16)
        end = date(d.year + (d.month == 12), d.month % 12 + 1, 15)
    else:
        prev_y, prev_m = (d.year - 1, 12) if d.month == 1 else (d.year, d.month - 1)
        start = date(prev_y, prev_m, 16)
        end = date(d.year, d.month, 15)
    return start, end


def ymd(d: date) -> str:
    return d.isoformat()


def parse_ymd(s: Any) -> date | None:
    try:
        return date.fromisoformat(str(s)) if YMD_RE.match(str(s)) else None
    except ValueError:
        return None


def pick_substitute_pair(working_days: Iterable[str], holidays: Iterable[dict]) -> tuple[str, str] | None:
    """振替の (休暇日, 振替元) を選ぶ。休暇日は勤務日、振替元は同じ月次期間の休日の候補。"""
    hol = [parse_ymd(h.get("date")) for h in holidays if isinstance(h, dict)]
    hol = [h for h in hol if h is not None]
    for text in list(working_days)[:90]:
        leave = parse_ymd(text)
        if leave is None:
            continue
        period = monthly_period(leave)
        for h in hol:
            if period[0] <= h <= period[1]:
                return ymd(leave), ymd(h)
    return None


def tiny_png_data_url() -> str:
    """手書きサインの代わりの 4×4 の黒い PNG（data URL）。"""
    w = h = 4

    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + b"\x00\x00\x00\xff" * w for _ in range(h))
    png = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))
    return "data:image/png;base64," + base64.b64encode(png).decode("ascii")


# ===== HTTP =====


@dataclass
class HttpResponse:
    status: int
    headers: dict[str, str]  # 小文字のキー
    text: str

    def json(self) -> Any:
        return json.loads(self.text)


Transport = Callable[[str, str, dict, "bytes | None"], HttpResponse]


class UrllibTransport:
    """標準ライブラリだけの送り手（リダイレクトは追わない）。base_url + path に送る。"""

    def __init__(self, base_url: str, timeout: float = 120.0):
        import urllib.request

        self.base_url = base_url.rstrip("/")
        self.timeout = timeout

        class _NoRedirect(urllib.request.HTTPRedirectHandler):
            def redirect_request(self, *a, **k):  # noqa: ANN002
                return None

        self._opener = urllib.request.build_opener(_NoRedirect)

    def __call__(self, method: str, path: str, headers: dict, body: bytes | None) -> HttpResponse:
        import urllib.error
        import urllib.request

        req = urllib.request.Request(self.base_url + path, data=body, method=method,
                                     headers={"User-Agent": USER_AGENT, **headers})
        try:
            with self._opener.open(req, timeout=self.timeout) as r:
                return HttpResponse(r.status, {k.lower(): v for k, v in r.headers.items()},
                                    r.read().decode("utf-8", "replace"))
        except urllib.error.HTTPError as e:
            return HttpResponse(e.code, {k.lower(): v for k, v in (e.headers or {}).items()},
                                e.read().decode("utf-8", "replace") if e.fp else "")


# ===== RPC の返事 =====


@dataclass
class RpcReply:
    status: int
    ok: bool | None
    result: Any = None
    error: str | None = None
    raw: str = ""
    ms: float = 0.0

    def brief(self) -> str:
        if self.status == 200 and self.ok:
            return "200 " + _short(self.result)
        return f"{self.status} {self.error if self.error is not None else _short(self.raw)}"

    def is_error(self, status: int, message: str) -> bool:
        return self.status == status and self.ok is False and self.error == message


def parse_rpc(resp: HttpResponse, ms: float = 0.0) -> RpcReply:
    """{"ok":true,"result":R} / {"ok":false,"error":msg} を読む。JSON でなければ ok=None。"""
    try:
        body = resp.json()
    except ValueError:
        return RpcReply(resp.status, None, error="JSON ではありません: " + _short(resp.text), raw=resp.text, ms=ms)
    if not isinstance(body, dict) or "ok" not in body:
        return RpcReply(resp.status, None, error="形が違います: " + _short(resp.text), raw=resp.text, ms=ms)
    if body["ok"] is True:
        return RpcReply(resp.status, True, result=body.get("result"), raw=resp.text, ms=ms)
    return RpcReply(resp.status, False, error=str(body.get("error", "")), raw=resp.text, ms=ms)


def _short(value: Any, n: int = 240) -> str:
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False, default=str)
    return text if len(text) <= n else text[:n] + "…"


# ===== 表 =====


@dataclass
class Item:
    id: str
    title: str
    status: str
    expected: str = ""
    actual: str = ""


@dataclass
class Report:
    items: list[Item] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def add(self, item_id: str, title: str, ok: bool | None, actual: Any = "", expected: str = "") -> Item:
        """ok=True→PASS、False→FAIL、None→SKIP（actual に理由）。"""
        status = SKIP if ok is None else (PASS if ok else FAIL)
        text = actual if isinstance(actual, str) else _short(actual)
        item = Item(item_id, title, status, expected, text)
        self.items.append(item)
        return item

    def problems(self, item_id: str, title: str, problems: list[str], expected: str = "",
                 skip_reason: str | None = None) -> Item:
        if skip_reason is not None:
            return self.add(item_id, title, None, skip_reason, expected)
        return self.add(item_id, title, not problems, "; ".join(problems) or "期待どおり", expected)

    def count(self, status: str) -> int:
        return sum(1 for i in self.items if i.status == status)

    @property
    def failed(self) -> list[Item]:
        return [i for i in self.items if i.status == FAIL]

    def exit_code(self) -> int:
        return 1 if self.failed else 0

    def table(self) -> str:
        def cell(s: str) -> str:
            return s.replace("|", "\\|").replace("\n", " ")

        lines = ["| # | 結果 | 項目 | 期待 | 実際 |", "|---|---|---|---|---|"]
        for i in self.items:
            lines.append(f"| {i.id} | {i.status} | {cell(i.title)} | {cell(i.expected)} | {cell(i.actual)} |")
        lines.append("")
        lines.append(f"PASS {self.count(PASS)} / FAIL {self.count(FAIL)} / SKIP {self.count(SKIP)}"
                     f"（全 {len(self.items)} 項目）")
        if self.failed:
            lines.append("FAIL: " + ", ".join(i.id for i in self.failed))
        return "\n".join(lines)

    def to_json(self) -> dict:
        return {"items": [i.__dict__ for i in self.items], "notes": self.notes,
                "counts": {s: self.count(s) for s in (PASS, FAIL, SKIP)}, "exitCode": self.exit_code()}


# ===== 返す形の確かめ（GAS の形。実装仕様 4-3・4-4） =====


def missing_keys(obj: Any, keys: Iterable[str]) -> list[str]:
    if not isinstance(obj, dict):
        return [f"dict ではありません: {_short(obj, 80)}"]
    return [f"キー {k} がありません" for k in keys if k not in obj]


def _list_of_dicts(value: Any, keys: Iterable[str], label: str) -> list[str]:
    if not isinstance(value, list):
        return [f"{label}が配列ではありません"]
    for it in value:
        p = missing_keys(it, keys)
        if p:
            return [f"{label}: " + p[0]]
    return []


def shape_problems(fn: str, result: Any, ctx: dict) -> list[str]:
    """読む関数の返事が GAS の形か。ctx: today・me・dept。"""
    p: list[str] = []
    if fn in ("api_isAdmin", "api_isSomu", "api_isDeptApprover"):
        if not isinstance(result, bool):
            p.append("true/false ではありません")
    elif fn == "api_getDeptList":
        if not isinstance(result, list) or not result:
            p.append("部署の配列（1 件以上）ではありません")
        else:
            p += _list_of_dicts(result, ("deptId", "deptName"), "部署")
            if not p and any(d["deptId"] != d["deptName"] for d in result):
                p.append("deptId と deptName が違う行があります（4-4-a は同じ値）")
    elif fn == "api_getWorkersByDept":
        p += _list_of_dicts(result, ("workerId", "workerName", "deptId"), "作業員")
        if not p and ctx.get("dept") and not result:
            p.append("部署の作業員が 0 人です")
    elif fn == "api_getWorkerInfo":
        if result is not None:
            p += missing_keys(result, ("workerId", "workerName", "deptId", "deptName", "email"))
    elif fn == "api_getLookupData":
        p += missing_keys(result, ("leaveKubun", "halfDayType", "leaveTypeFull", "leaveTypeHalf", "specialReasons"))
        if not p:
            if result["leaveKubun"] != [DAY_FULL, DAY_HALF] or result["halfDayType"] != [HALF_AM, HALF_PM]:
                p.append("leaveKubun・halfDayType が GAS の値と違います")
            if result["leaveTypeFull"] != [KIND_PAID, KIND_SUBSTITUTE, KIND_SPECIAL]:
                p.append("leaveTypeFull が GAS の値と違います")
            if not isinstance(result["specialReasons"], list):
                p.append("specialReasons が配列ではありません")
    elif fn == "api_getWorkingDays":
        if not isinstance(result, list) or not all(YMD_RE.match(str(d)) for d in result):
            p.append("yyyy-MM-dd の配列ではありません")
        elif result and (result != sorted(result) or result[0] < ctx.get("today", "")):
            p.append("今日から先の昇順ではありません")
    elif fn == "api_getHolidayDays":
        p += _list_of_dicts(result, ("date", "label"), "休日")
    elif fn == "api_getPastWorkingDays":
        p += _list_of_dicts(result, ("date", "dow", "label"), "過去の勤務日")
        if not p and any(str(d["date"]) > ctx.get("today", "9") for d in result):
            p.append("今日より後の日があります")
    elif fn == "api_getUpcomingLeaves":
        p += _list_of_dicts(result, UPCOMING_KEYS, "予定")
    elif fn == "api_getAllPendingRequests":
        p += _list_of_dicts(result, PENDING_KEYS, "承認待ち")
        if not p and any(it["status"] != ST_SUBMITTED for it in result):
            p.append("申請中でない行があります")
    elif fn == "api_getPending2ndApprovals":
        p += _list_of_dicts(result, PENDING2_KEYS, "2 次承認待ち")
        if not p and any("halfDayType" in it for it in result):
            p.append("halfDayType が入っています（GAS は入れない）")
    elif fn == "api_getApproverList":
        p += _list_of_dicts(result, ("email", "name"), "電子印")
    elif fn == "api_getCurrentApproverInfo":
        p += missing_keys(result, ("email", "name"))
        if not p and ctx.get("me") and result["email"] != ctx["me"]:
            p.append("email がログインのメールと違います")
    elif fn == "api_getSomuAdminView":
        p += _list_of_dicts(result, WORKER_VIEW_KEYS, "総務管理")
    elif fn == "api_getApproverProfiles":
        p += missing_keys(result, ("isAdmin", "currentEmail", "profiles"))
        if not p and ctx.get("me") and result["currentEmail"] != ctx["me"]:
            p.append("currentEmail がログインのメールと違います")
    elif fn == "api_getDeptAdminView":
        p += missing_keys(result, ("workers", "deptNames"))
        if not p:
            p += _list_of_dicts(result["workers"], WORKER_VIEW_KEYS, "部署管理")
    elif fn in ("api_getAdminSomuView", "api_getAdminBossView"):
        p += _list_of_dicts(result, SUMMARY_VIEW_KEYS, "管理")
    elif fn == "api_getLeaveStatusSearch":
        p += missing_keys(result, ("mode", "fyYear"))
        if not p:
            if result["mode"] == "list":
                p += missing_keys(result, ("deptFilter", "workers"))
            elif result["mode"] == "personal":
                p += missing_keys(result, ("worker", "milestones", "upcoming"))
            else:
                p.append("mode が list か personal ではありません")
    elif fn == "api_getLeaveRequestById":
        if result is not None:
            p += missing_keys(result, REQUEST_KEYS)
    elif fn == "ops_getStatus":
        p += missing_keys(result, ("rowSync", "gasJobs", "gasPing", "sheetWriteGate", "lastRuns"))
    else:
        p.append("形の確かめがありません: " + fn)
    return p


def check_headers(headers: dict[str, str]) -> list[str]:
    p = []
    csp = headers.get("content-security-policy")
    if csp != CSP_EXPECTED:
        p.append("CSP が違います: " + str(csp))
    for k, v in HEADERS_EXPECTED.items():
        if headers.get(k) != v:
            p.append(f"{k} が {v} ではありません: {headers.get(k)}")
    for k in HEADERS_PRESENT:
        if not headers.get(k):
            p.append(f"{k} がありません")
    return p


def check_page_html(text: str, title: str) -> list[str]:
    p = []
    if f"<title>{title}</title>" not in text:
        p.append(f"題名が「{title}」ではありません")
    if INLINE_SCRIPT_RE.search(text):
        p.append("インラインの script があります")
    if ON_ATTR_RE.search(text):
        p.append("on 属性があります")
    return p


def page_title(text: str) -> str:
    m = re.search(r"<title>([^<]*)</title>", text)
    return m.group(1) if m else ""


# ===== シートの行（統合業務メールと同じ条件）と GAS の仕事 =====


def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def sheet_row_problems(cells: dict | None, *, status: str, at2_empty: bool, sign: str | None = None,
                       pdf: str | None = None, worker_name: str | None = None) -> list[str]:
    """T_LEAVE_REQUEST の 1 行（見出し→UNFORMATTED の値）を、統合業務メール（readerLeave.gs）と同じ条件で確かめる。

    - STATUS が状態の文字と同じ
    - LEAVE_DATE が日付のセル（シリアル値の数）。文字や空なら統合業務メールの通知から黙って消える
    - APPROVED_AT2 が空（2 次承認待ち）か、日時のセル（2 次承認済み）
    sign・pdf: 'set'（空でない）／'empty'（空）／None（見ない）。worker_name: 文字のまま同じか。
    """
    if cells is None:
        return ["シートの行が見つかりません"]
    p: list[str] = []
    if cells.get("STATUS") != status:
        p.append(f"STATUS が {status} ではありません: {cells.get('STATUS')!r}")
    if not _is_number(cells.get("LEAVE_DATE")):
        p.append(f"LEAVE_DATE が日付のセルではありません: {cells.get('LEAVE_DATE')!r}")
    at2 = cells.get("APPROVED_AT2", "")
    if at2_empty and at2 not in ("", None):
        p.append(f"APPROVED_AT2 が空ではありません: {at2!r}")
    if not at2_empty and not _is_number(at2):
        p.append(f"APPROVED_AT2 が日時のセルではありません: {at2!r}")
    for col, want in (("SIGN_IMAGE_URL", sign), ("PDF_URL", pdf)):
        if want is None:
            continue
        has = str(cells.get(col) or "").strip() != ""
        if want == "set" and not has:
            p.append(f"{col} が空です")
        if want == "empty" and has:
            p.append(f"{col} が空ではありません")
    if worker_name is not None and cells.get("WORKER_NAME") != worker_name:
        p.append(f"WORKER_NAME が入れた文字と違います: {cells.get('WORKER_NAME')!r}")
    return p


def job_problems(jobs: list[dict], expected: list[tuple[str, dict[str, set[str]]]]) -> list[str]:
    """申請の GAS の仕事（id の順）が expected の並びで、すべて done、段の状態が期待どおりか。

    expected: [(kind, {段: 受け入れる状態の集合}), …]。段は GAS の返事の sign・mail・pdf・trash・ledger。
    """
    ordered = sorted(jobs, key=lambda j: j["id"])
    kinds = [j["kind"] for j in ordered]
    want = [k for k, _ in expected]
    if kinds != want:
        return [f"仕事の並びが違います: {' → '.join(kinds) or '（なし）'}（期待: {' → '.join(want)}）"]
    p: list[str] = []
    for job, (kind, stages) in zip(ordered, expected):
        if job["state"] != "done":
            p.append(f"{kind}（#{job['id']}）が done ではありません: {job['state']} {job.get('error', '')}".strip())
            continue
        result = job.get("result") or {}
        for stage, allowed in stages.items():
            got = (result.get(stage) or {}).get("status") if isinstance(result.get(stage), dict) else None
            if got not in allowed:
                p.append(f"{kind}（#{job['id']}）の {stage} が {'/'.join(sorted(allowed))} ではありません: {got}")
    return p


def mails_sent(jobs: Iterable[dict]) -> int:
    n = 0
    for j in jobs:
        mail = (j.get("result") or {}).get("mail")
        if isinstance(mail, dict) and mail.get("status") == "sent":
            n += 1
    return n


def cookie_header(value: str) -> str:
    value = value.strip()
    if value.startswith(SESSION_COOKIE + "="):
        return value
    if value.lower().startswith("cookie:"):
        return value.split(":", 1)[1].strip()
    return f"{SESSION_COOKIE}={value}"


# ===== DB とシートを見る（コンテナの中だけ。読むだけ） =====


class DbProbe:
    """leave_requests・jobs・sheet_sync_state を読み、ステージングの T_LEAVE_REQUEST の行を Sheets API で読む。"""

    def __init__(self, session_factory, settings, sheets_factory: Callable[[Any], Any] | None = None):
        self.session_factory = session_factory
        self.settings = settings
        self.sheets_factory = sheets_factory
        self._header: list[str] | None = None

    @classmethod
    def from_app(cls) -> "DbProbe":
        from app.config import get_settings
        from app.db import get_sessionmaker

        return cls(get_sessionmaker(), get_settings())

    def request(self, req_id: str) -> dict | None:
        from app.models import LeaveRequest

        with self.session_factory() as s:
            r = s.get(LeaveRequest, req_id)
            if r is None:
                return None
            return {"status": r.status, "approvedAt": r.approved_at is not None, "approvedAt2": r.approved_at2 is not None,
                    "signImageUrl": r.sign_image_url or "", "pdfUrl": r.pdf_url or "", "sheetRow": r.sheet_row,
                    "signPng": r.sign_png is not None, "workerName": r.worker_name}

    def jobs_for(self, req_id: str) -> list[dict]:
        from sqlalchemy import select

        from app.models import Job

        out = []
        with self.session_factory() as s:
            for r in s.scalars(select(Job).where(Job.request_id == req_id).order_by(Job.id)):
                try:
                    result = json.loads(r.result_json) if r.result_json else {}
                except ValueError:
                    result = {}
                out.append({"id": r.id, "kind": r.kind, "state": r.state, "error": (r.last_error or "")[:200],
                            "result": result if isinstance(result, dict) else {}})
        return out

    def synced(self, req_id: str) -> bool:
        from app.models import SheetSyncState

        with self.session_factory() as s:
            st = s.get(SheetSyncState, ("leave_requests", req_id))
            return st is not None and st.synced_version >= st.desired_version and st.state != "blocked"

    def _sheets(self, db):
        if self.sheets_factory is not None:
            return self.sheets_factory(db)
        from app.services.sheets_api import SheetsApi

        return SheetsApi.for_db(db, self.settings)

    def sheet_row(self, req_id: str) -> dict | None:
        """シートの行（見出し → UNFORMATTED の値）。ステージングの DB スプレッドシートだけを読む。"""
        if not getattr(self.settings, "is_staging", False):
            raise RuntimeError("ステージングでだけシートを読みます")
        info = self.request(req_id)
        if not info or not info["sheetRow"]:
            return None
        sid = self.settings.staging_spreadsheet_id
        n = int(info["sheetRow"])
        with self.session_factory() as s:
            api = self._sheets(s)
            header_rows, rows = api.batch_get_values(sid, ["'T_LEAVE_REQUEST'!1:1", f"'T_LEAVE_REQUEST'!{n}:{n}"])
        header = [str(h).strip() for h in (header_rows[0] if header_rows else [])]
        row = list(rows[0]) if rows else []
        row += [""] * (len(header) - len(row))
        cells = dict(zip(header, row))
        if cells.get("REQ_ID") != req_id:
            return None
        return cells


# ===== 本体 =====


class Refused(RuntimeError):
    pass


class Runner:
    def __init__(self, api: Transport, *, web: Transport | None = None, origin: str = DEFAULT_ORIGIN,
                 cookie: str | None = None, outsider_cookie: str | None = None, probe: DbProbe | None = None,
                 write: bool = False, job_wait: float = 300.0, poll_interval: float = 3.0, dept: str | None = None,
                 run_id: str | None = None, sleep: Callable[[float], None] = time.sleep,
                 clock: Callable[[], float] = time.monotonic, now: datetime | None = None):
        self.api = api
        self.web = web
        self.origin = origin.rstrip("/")
        self.cookie = cookie_header(cookie) if cookie else None
        self.outsider_cookie = cookie_header(outsider_cookie) if outsider_cookie else None
        self.probe = probe
        self.write = write
        self.job_wait = job_wait
        self.poll_interval = poll_interval
        self.dept = dept
        self.now = now or datetime.now(JST)
        self.run_id = run_id or self.now.astimezone(JST).strftime("%Y%m%d-%H%M%S")
        self.sleep = sleep
        self.clock = clock
        self.report = Report()
        self.times: dict[str, list[float]] = {}
        self.health: dict | None = None
        self.env: str | None = None
        self.me: str = ""
        self.today: date = today_jst(self.now)
        self.fy: int = fiscal_year(self.today)
        self.created: list[str] = []
        self.canceled: set[str] = set()

    # ---- 送る ----

    def http(self, method: str, path: str, headers: dict | None = None, body: bytes | None = None,
             cookie: str | None = None, transport: Transport | None = None) -> HttpResponse:
        h = dict(headers or {})
        if cookie:
            h["Cookie"] = cookie
        return (transport or self.api)(method, path, h, body)

    def rpc(self, fn: str, *args: Any, key: str | None = "auto", origin: str | None = None,
            content_type: str = "application/json", cookie: str | None = "main", raw_body: bytes | None = None
            ) -> RpcReply:
        headers = {"Content-Type": content_type, "Origin": origin or self.origin}
        if fn in WRITE_FUNCTIONS and key:
            headers["X-Idempotency-Key"] = str(uuid.uuid4()) if key == "auto" else key
        body = raw_body if raw_body is not None else json.dumps(list(args), ensure_ascii=False).encode("utf-8")
        jar = {"main": self.cookie, "outsider": self.outsider_cookie}.get(cookie, cookie) if cookie else None
        t = time.perf_counter()
        resp = self.http("POST", "/api/rpc/" + fn, headers, body, cookie=jar)
        ms = (time.perf_counter() - t) * 1000
        self.times.setdefault(fn, []).append(ms)
        return parse_rpc(resp, ms)

    # ---- 入口 ----

    def check_health(self) -> str | None:
        r = self.http("GET", "/healthz")
        try:
            body = r.json()
        except ValueError:
            body = None
        self.health = body if isinstance(body, dict) else None
        self.env = (self.health or {}).get("env")
        self.report.add("H-1", "/healthz が 200 {ok:true, env:staging}", r.status == 200 and isinstance(body, dict) and
                        body.get("ok") is True and self.env == "staging", f"{r.status} {_short(r.text)}",
                        '{"ok":true,"env":"staging"}')
        if isinstance(body, dict):
            extra = sorted(set(body) - {"ok", "env"})
            self.report.add("H-2", "ログインなしの /healthz は ok と env だけ", not extra, f"余分なキー: {extra}",
                            "ok と env だけ")
        return self.env

    def run(self) -> Report:
        if self.health is None:
            self.check_health()
        if self.env != "staging":
            raise Refused(f"/healthz の env が staging ではありません（{self.env}）。受け入れテストはステージングでだけ動かします。")
        self.check_entry()
        self.check_web_headers()
        if self.cookie:
            self.check_session()
            self.check_pages()
            self.check_reads()
            self.check_not_found()
            self.check_boundaries()
        else:
            self.report.add("R", "ログインしての確かめ（画面・34 関数・書き込み）", None,
                            "セッションがありません（--mint-session か --cookie）")
        if self.outsider_cookie:
            self.check_permissions()
        else:
            self.report.add("PM", "権限（総務でないアカウントの 2 次承認・ops_*）", None,
                            "権限の無いセッションがありません（--mint-session か --outsider-cookie）")
        if self.write and self.cookie:
            self.write_scenario()
        elif self.write:
            self.report.add("W", "書き込みの試験", False, "セッションがありません")
        self.report.notes.append("速さ（中央値 ms）: " + json.dumps(
            {k: round(statistics.median(v), 1) for k, v in sorted(self.times.items())}, ensure_ascii=False))
        return self.report

    def check_entry(self) -> None:
        rep = self.report
        r = self.http("GET", "/")
        loc = r.headers.get("location", "")
        rep.add("E-1", "/ はログインなしで 302 /auth/login?next=…", r.status in (302, 303, 307) and
                loc.startswith("/auth/login?next="), f"{r.status} location={loc}", "302 /auth/login?next=")
        r = self.http("GET", "/?page=preview&reqId=LV-20261126-001")
        loc = r.headers.get("location", "")
        rep.add("E-2", "ログインへの next が page と reqId を保つ", r.status in (302, 303, 307) and
                "preview" in loc and "LV-20261126-001" in loc, f"{r.status} location={loc}", "next に preview と reqId")
        reply = self.rpc("api_getDeptList", cookie=None)
        want = {"ok": False, "error": RPC_MSG["login"], "login": "/auth/login"}
        try:
            body = json.loads(reply.raw)
        except ValueError:
            body = None
        rep.add("E-3", "ログインなしの RPC は 401", reply.status == 401 and body == want, reply.brief(),
                "401 " + json.dumps(want, ensure_ascii=False))
        r = self.http("GET", "/auth/me")
        rep.add("E-4", "ログインなしの /auth/me は 401", r.status == 401, f"{r.status}", "401")
        if not self.cookie:
            for i, t in (("E-5", "JSON でない本文は 415"), ("E-6", "別サイトの Origin は 403"),
                         ("E-7", "知らない関数は 404"), ("E-8", "出さない GAS の関数（api_canAccessAdmin）は 404"),
                         ("E-9", "本文が配列でなければ 400")):
                rep.add(i, t, None, "セッションがありません（401 が先に返る）")
            return
        r = self.rpc("api_getDeptList", content_type="text/plain")
        rep.add("E-5", "JSON でない本文は 415", r.is_error(415, RPC_MSG["shape"]), r.brief(), "415 " + RPC_MSG["shape"])
        r = self.rpc("api_getDeptList", origin="https://evil.example")
        rep.add("E-6", "別サイトの Origin は 403", r.is_error(403, RPC_MSG["origin"]), r.brief(), "403 " + RPC_MSG["origin"])
        r = self.rpc("api_noSuchFunction")
        rep.add("E-7", "知らない関数は 404", r.is_error(404, RPC_MSG["unknown"]), r.brief(), "404 " + RPC_MSG["unknown"])
        r = self.rpc("api_canAccessAdmin")
        rep.add("E-8", "出さない GAS の関数（api_canAccessAdmin）は 404", r.status == 404, r.brief(), "404")
        r = self.rpc("api_getDeptList", raw_body=b'{"a":1}')
        rep.add("E-9", "本文が配列でなければ 400", r.is_error(400, RPC_MSG["shape"]), r.brief(), "400 " + RPC_MSG["shape"])

    def check_web_headers(self) -> None:
        if self.web is None:
            self.report.add("E-10", "CSP などのセキュリティヘッダー（nginx）", None, "--web-url がありません")
            return
        r = self.http("GET", "/healthz", transport=self.web)
        self.report.problems("E-10", "CSP などのセキュリティヘッダー（nginx）", check_headers(r.headers),
                             "nginx.conf と同じ CSP・nosniff・SAMEORIGIN ほか")
        r = self.http("GET", "/static/js/gas-shim.js", transport=self.web)
        self.report.add("E-11", "/static/js/gas-shim.js を nginx が返す", r.status == 200 and "google" in r.text,
                        f"{r.status}", "200")

    def check_session(self) -> None:
        r = self.http("GET", "/auth/me", cookie=self.cookie)
        try:
            body = r.json()
        except ValueError:
            body = {}
        self.me = str(body.get("email", "")).lower() if isinstance(body, dict) else ""
        self.report.add("S-0", "セッションで /auth/me が 200 {ok, email}", r.status == 200 and bool(self.me),
                        f"{r.status} {_short(r.text)}", "200 {ok, email}")
        admin, somu = self.rpc("api_isAdmin"), self.rpc("api_isSomu")
        self.report.add("S-1", "受け入れのアカウントが総務か管理者（2 次承認の流れに要る）",
                        admin.result is True or somu.result is True, f"isAdmin={admin.result} isSomu={somu.result}",
                        "どちらかが true（M_SYSTEM_SETTING の ADMIN_EMAILS か SOMU_EMAILS）")

    def check_pages(self) -> None:
        for i, page in enumerate(PAGE_TITLES, start=1):
            r = self.http("GET", "/?page=" + page, cookie=self.cookie)
            title = PAGE_TITLES[page]
            got = page_title(r.text) if r.status == 200 else ""
            if r.status == 200 and page in NO_AUTH_TITLES and got == NO_AUTH_TITLES[page]:
                self.report.add(f"P-{i}", f"画面 ?page={page} が開く", None,
                                f"このアカウントには権限が無く「{got}」が出ました（JEV で権限のあるアカウントで確かめる）",
                                f"200・題名「{title}」")
                continue
            problems = [] if r.status == 200 else [f"HTTP {r.status}"]
            if r.status == 200:
                problems += check_page_html(r.text, title)
                if r.headers.get("cache-control") != "no-store":
                    problems.append("Cache-Control が no-store ではありません")
            self.report.problems(f"P-{i}", f"画面 ?page={page} が開く（題名・インライン script なし）", problems,
                                 f"200・題名「{title}」")

    # ---- 読む関数 ----

    def check_reads(self) -> None:
        rep = self.report
        depts = self.rpc("api_getDeptList")
        dept = self.dept or (depts.result[0]["deptName"] if depts.ok and isinstance(depts.result, list) and depts.result
                             and isinstance(depts.result[0], dict) else "")
        self._dept = dept
        workers = self.rpc("api_getWorkersByDept", dept)
        worker = (workers.result[0]["workerId"] if workers.ok and isinstance(workers.result, list) and workers.result
                  and isinstance(workers.result[0], dict) else "")
        ym = ymd(self.today)[:7]
        ctx = {"today": ymd(self.today), "me": self.me, "dept": dept}
        fy = self.fy
        calls: list[tuple[str, tuple]] = [
            ("api_isAdmin", ()), ("api_isSomu", ()), ("api_isDeptApprover", ()), ("api_getDeptList", ()),
            ("api_getWorkersByDept", (dept,)), ("api_getWorkerInfo", ()), ("api_getLookupData", ()),
            ("api_getWorkingDays", ()), ("api_getHolidayDays", ("",)), ("api_getPastWorkingDays", (ym,)),
            ("api_getUpcomingLeaves", ("", "")), ("api_getAllPendingRequests", ()), ("api_getApproverList", ()),
            ("api_getCurrentApproverInfo", ()), ("api_getPending2ndApprovals", ()),
            ("api_getSomuAdminView", (fy, "ALL")), ("api_getApproverProfiles", ()),
            ("api_getDeptAdminView", (fy, [dept])), ("api_getAdminSomuView", (fy, "")),
            ("api_getAdminBossView", ("", fy)), ("api_getLeaveStatusSearch", (fy, dept, "")),
            ("api_getLeaveRequestById", ("",)), ("ops_getStatus", ()),
        ]
        if worker:
            calls.append(("api_getLeaveStatusSearch", (fy, dept, worker)))
        for fn, args in calls:
            r = self.rpc(fn, *args)
            title = f"{fn}{_args_text(args)} が GAS の形を返す"
            if fn in ("api_getAdminSomuView", "api_getAdminBossView") and r.is_error(400, MSG["admin_page_only"]):
                rep.add(f"F-{fn}", title, None, "このアカウントは ALLOW_ADMIN_PAGE_EMAILS に入っていません（権限の拒否は PM で確かめる）")
                continue
            if fn == "api_getDeptAdminView" and r.is_error(400, MSG["dept_approver_only"]):
                rep.add(f"F-{fn}", title, None, "このアカウントは管理者でも部署承認者でもありません")
                continue
            problems = [] if (r.status == 200 and r.ok) else [r.brief()]
            if not problems:
                problems = shape_problems(fn, r.result, ctx)
            key = f"F-{fn}" + ("-personal" if fn == "api_getLeaveStatusSearch" and len(args) == 3 and args[2] else "")
            rep.problems(key, title, problems, "200・実装仕様 4-3／4-4 の形")
            if fn == "ops_getStatus" and r.ok and isinstance(r.result, dict):
                ping = r.result.get("gasPing") or {}
                rep.add("O-1", "GAS の内部用の入口の照合（gas_ping）が通っている",
                        ping.get("ok") is True and not ping.get("problems"), _short(ping), "ok:true・problems なし")
                failed = ((r.result.get("gasJobs") or {}).get("counts") or {}).get("failed", 0)
                rep.add("O-2", "失敗した GAS の仕事が 0 件", not failed, f"failed={failed}", "0")
                blocked = (r.result.get("rowSync") or {}).get("blocked") or []
                rep.add("O-3", "止まっている行の反映が 0 件", not blocked, _short(blocked), "0")
                gate = r.result.get("sheetWriteGate") or {}
                rep.add("O-4", "シートへ書く前の関所が開いている（staging）", gate.get("allowed") is True, _short(gate),
                        "allowed:true")
        r = self.rpc("api_getWorkersByDept", "")
        rep.add("F-api_getWorkersByDept-empty", "api_getWorkersByDept('') は []", r.ok and r.result == [], r.brief(), "[]")

    def check_not_found(self) -> None:
        rep = self.report
        r = self.rpc("api_getLeaveRequestById", BOGUS_ID)
        rep.add("N-1", "存在しない reqId のプレビューはすぐ null（GAS の約 6 秒の待ちをしない）",
                r.status == 200 and r.ok is True and r.result is None and r.ms < 3000, f"{r.brief()} {r.ms:.0f}ms",
                "200 null・3 秒未満")
        probes: list[tuple[str, str, tuple, int, str]] = [
            ("N-2", "api_approveLeaveRequest", (BOGUS_ID, "", self.me), 400, MSG["not_found_prefix"] + BOGUS_ID),
            ("N-3", "api_editLeaveRequest", (BOGUS_ID, {}), 400, MSG["not_found"]),
            ("N-4", "api_deleteLeaveRequest", (BOGUS_ID,), 400, MSG["not_found"]),
            ("N-5", "api_approveLeaveRequest2", (BOGUS_ID, self.me), 400, MSG["not_found_prefix"] + BOGUS_ID),
            ("N-6", "api_approveLeaveRequest", ("", "", self.me), 400, MSG["no_req_id"]),
            ("N-7", "ops_retryJob", ("x",), 400, MSG["bad_job_id"]),
            ("N-8", "ops_revokeSessions", ("",), 400, MSG["no_email"]),
        ]
        for item_id, fn, args, status, message in probes:
            r = self.rpc(fn, *args)
            rep.add(item_id, f"{fn}{_args_text(args)} は何も書かずに GAS と同じ文で止まる", r.is_error(status, message),
                    r.brief(), f"{status} {message}")

    def _submit_base(self, worker_id: str, worker_name: str, leave_date: str) -> dict:
        return {"deptId": getattr(self, "_dept", ""), "deptName": getattr(self, "_dept", ""), "workerId": worker_id,
                "workerName": worker_name, "leaveKubun": DAY_FULL, "halfDayType": "", "leaveType": KIND_PAID,
                "leaveDate": leave_date, "substituteDate": "", "specialReason": "",
                "paidDetail": f"{TEST_MARK}{self.run_id} 受け入れテストの申請", "isRetroactive": False}

    @property
    def worker_id(self) -> str:
        return f"QA-{self.run_id}"

    @property
    def worker_name(self) -> str:
        return f"{TEST_MARK}{self.run_id}"

    def check_boundaries(self) -> None:
        """境目の日付と入力の誤り（200 {ok:false}。行は作らない）。"""
        rep = self.report
        today = self.today
        wd = self.rpc("api_getWorkingDays")
        working = [d for d in (wd.result or []) if isinstance(d, str)] if wd.ok else []
        leave = working[0] if working else ymd(today + timedelta(days=1))
        hol = self.rpc("api_getHolidayDays", "")
        future_holidays = [h["date"] for h in (hol.result or []) if isinstance(h, dict) and
                           str(h.get("date", "")) >= ymd(today) and h.get("date") not in working] if hol.ok else []
        non_working = future_holidays[0] if future_holidays else "2000-01-03"
        start = fy_start(today)
        sub_out = ymd(monthly_period(parse_ymd(leave) or today)[0] - timedelta(days=1))
        cases: list[tuple[str, str, dict, str, bool]] = [
            ("B-1", "形の違う休暇日（2026/01/05）", {"leaveDate": "2026/01/05"}, MSG["bad_format"], False),
            ("B-2", f"勤務日でない日（{non_working}）を通常の申請", {"leaveDate": non_working}, MSG["not_working_day"], False),
            ("B-3", f"過去申請で年度の始まりの前日（{ymd(start - timedelta(days=1))}）",
             {"leaveDate": ymd(start - timedelta(days=1)), "isRetroactive": True}, MSG["retro_out_of_fy"], False),
            ("B-4", f"過去申請で明日（{ymd(today + timedelta(days=1))}）",
             {"leaveDate": ymd(today + timedelta(days=1)), "isRetroactive": True}, MSG["retro_out_of_fy"], False),
            ("B-5", "振替で振替元なし", {"leaveDate": leave, "leaveType": KIND_SUBSTITUTE}, MSG["substitute_missing"], False),
            ("B-6", f"振替で振替元が前の月次期間の最後の日（{sub_out}）",
             {"leaveDate": leave, "leaveType": KIND_SUBSTITUTE, "substituteDate": sub_out}, MSG["substitute_period"], True),
        ]
        for item_id, title, patch, message, prefix in cases:
            data = dict(self._submit_base(self.worker_id, self.worker_name, leave), **patch)
            r = self.rpc("api_submitLeaveRequest", data)
            res = r.result if r.ok and isinstance(r.result, dict) else {}
            got = str(res.get("error", ""))
            ok = r.status == 200 and res.get("ok") is False and (got.startswith(message) if prefix else got == message)
            rep.add(item_id, f"境目: {title} は 200 {{ok:false}} で止まる（書かない）", ok, r.brief(),
                    "200 {ok:false, error:" + message + ("…" if prefix else "") + "}")
        up = self.rpc("api_getUpcomingLeaves", "", self.worker_id)
        mine = [it for it in (up.result or []) if isinstance(it, dict)] if up.ok else None
        rep.add("B-7", "境目の確かめで申請が 1 件もできていない", mine == [], _short(mine), "試験の作業員の予定が 0 件")

    # ---- 権限 ----

    def check_permissions(self) -> None:
        rep = self.report
        c = "outsider"
        r = self.http("GET", "/auth/me", cookie=self.outsider_cookie)
        try:
            who = str(r.json().get("email", ""))
        except (ValueError, AttributeError):
            who = ""
        a, s = self.rpc("api_isAdmin", cookie=c), self.rpc("api_isSomu", cookie=c)
        rep.add("PM-0", "権限の無いアカウントでログインしている", r.status == 200 and a.result is False and s.result is False,
                f"{who} isAdmin={a.result} isSomu={s.result}", "isAdmin=false・isSomu=false")
        cases: list[tuple[str, str, tuple, int, str]] = [
            ("PM-1", "api_approveLeaveRequest2", (BOGUS_ID, who), 400, MSG["somu_only"]),
            ("PM-2", "api_getPending2ndApprovals", (), 400, MSG["somu_only"]),
            ("PM-3", "api_getSomuAdminView", (self.fy, "ALL"), 400, MSG["somu_only"]),
            ("PM-4", "api_getDeptAdminView", (self.fy, []), 400, MSG["dept_approver_only"]),
            ("PM-5", "api_getAdminBossView", ("", self.fy), 400, MSG["admin_page_only"]),
            ("PM-6", "api_getAdminSomuView", (self.fy, ""), 400, MSG["admin_page_only"]),
        ]
        cases += [(f"PM-{7 + i}", fn, args, 403, RPC_MSG["not_operator"]) for i, (fn, args) in enumerate((
            ("ops_getStatus", ()), ("ops_retryJob", ("1",)), ("ops_syncSheetConfigNow", ()), ("ops_runMasterSync", ()),
            ("ops_rebuildSummaryNow", ()), ("ops_sendWarnMail", ()), ("ops_revokeSessions", ("x@example.com",))))]
        for item_id, fn, args, status, message in cases:
            r = self.rpc(fn, *args, cookie=c)
            rep.add(item_id, f"権限の無いアカウントの {fn} は拒まれる", r.is_error(status, message), r.brief(),
                    f"{status} {message}")
        r = self.http("GET", "/?page=somuAdmin", cookie=self.outsider_cookie)
        rep.add("PM-14", "権限の無いアカウントの ?page=somuAdmin は権限エラーの画面", r.status == 200 and
                SOMU_NO_AUTH_TEXT in r.text, f"{r.status} 題名「{page_title(r.text)}」", SOMU_NO_AUTH_TEXT)

    # ---- 書き込みの試験（staging だけ） ----

    def wait_for(self, predicate: Callable[[], Any]) -> Any:
        deadline = self.clock() + self.job_wait
        while True:
            value = predicate()
            if value or self.clock() >= deadline:
                return value
            self.sleep(self.poll_interval)

    def wait_jobs(self, rid: str, count: int) -> list[dict]:
        """その申請の GAS の仕事が count 件以上になり、すべて終わる（done・skipped・failed）まで待つ。"""
        def finished() -> tuple[list[dict]] | None:
            js = self.probe.jobs_for(rid)
            return (js,) if len(js) >= count and all(j["state"] in JOB_FINISHED for j in js) else None

        done = self.wait_for(finished)
        return done[0] if done else self.probe.jobs_for(rid)

    def stage(self, item_id: str, title: str, rid: str | None, *, status: str, at2_empty: bool,
              jobs: list[tuple[str, dict[str, set[str]]]] | None = None, sign: str | None = None, pdf: str | None = None,
              worker_name: str | None = None) -> None:
        """1 つの段の確かめ: GAS の仕事 → DB の状態 → 行の反映 → シートの行（統合業務メールの条件）。"""
        expected = f"DB と STATUS={status}・休暇日が日付・APPROVED_AT2 {'空' if at2_empty else 'あり'}" + (
            f"・仕事 {' → '.join(k.split('.')[-1] for k, _ in jobs)}" if jobs else "")
        if rid is None:
            self.report.add(item_id, title, None, "申請がありません", expected)
            return
        if self.probe is None:
            self.report.add(item_id, title, None, "DB を見られません（--mint-session で動かす）", expected)
            return
        p: list[str] = []
        if jobs is not None:
            p += job_problems(self.wait_jobs(rid, len(jobs)), jobs)
        info = self.probe.request(rid) or {}
        if info.get("status") != status:
            p.append(f"DB の状態が {status} ではありません: {info.get('status')}")
        if bool(info.get("approvedAt2")) == at2_empty:
            p.append("DB の 2 次承認日時が期待と違います")
        if info.get("signPng"):
            p.append("DB に手書きサインの PNG が残っています（afterApprove1 の後は消える）")
        if not self.wait_for(lambda: self.probe.synced(rid)):
            p.append("行の反映が済みません")
        try:
            cells = self.probe.sheet_row(rid)
        except Exception as exc:  # noqa: BLE001 - シートの読み取りの失敗は不合格として残す
            cells = None
            p.append(f"シートを読めません: {type(exc).__name__}: {_short(str(exc), 120)}")
        if cells is not None or not p:
            p += sheet_row_problems(cells, status=status, at2_empty=at2_empty, sign=sign, pdf=pdf,
                                    worker_name=worker_name)
        self.report.problems(item_id, title, p, expected)

    def submit(self, item_id: str, title: str, data: dict) -> str | None:
        r = self.rpc("api_submitLeaveRequest", data)
        res = r.result if r.ok and isinstance(r.result, dict) else {}
        rid = res.get("reqId") if res.get("ok") is True else None
        if rid:
            self.created.append(rid)
        self.report.add(item_id, title, bool(rid) and bool(REQ_ID_RE.match(str(rid))), r.brief(),
                        "200 {ok:true, reqId:LV-yyyyMMdd-NNN}")
        return rid

    def write_scenario(self) -> None:
        rep = self.report
        today = self.today
        try:
            wd = self.rpc("api_getWorkingDays")
            working = [d for d in (wd.result or []) if isinstance(d, str)] if wd.ok else []
            lookup = self.rpc("api_getLookupData")
            reasons = (lookup.result or {}).get("specialReasons") or [] if lookup.ok else []
            hol = self.rpc("api_getHolidayDays", self.worker_id)
            holidays = [h for h in (hol.result or []) if isinstance(h, dict)] if hol.ok else []
            past: list[str] = []
            for ym in (ymd(today)[:7], ymd(today.replace(day=1) - timedelta(days=1))[:7]):
                pw = self.rpc("api_getPastWorkingDays", ym)
                past = [d["date"] for d in (pw.result or []) if isinstance(d, dict) and str(d.get("date")) < ymd(today)]
                if past:
                    break
            if not getattr(self, "_dept", "") or not working:
                rep.add("W-0", "書き込みの試験の準備（部署と勤務日）", False,
                        f"dept={getattr(self, '_dept', '')} 勤務日={len(working)}", "部署 1 つ・勤務日 1 日以上")
                return
            leave = working[0]
            base = self._submit_base(self.worker_id, self.worker_name, leave)
            rep.notes.append(f"試験の作業員: {self.worker_id}（部署 {self._dept}）。休暇日 {leave}")

            # 申請 A: 全日・有給。同じ冪等キーで 2 回送って 1 件
            key = str(uuid.uuid4())
            a1 = self.rpc("api_submitLeaveRequest", base, key=key)
            a2 = self.rpc("api_submitLeaveRequest", base, key=key)
            rid_a = (a1.result or {}).get("reqId") if a1.ok and isinstance(a1.result, dict) else None
            if rid_a:
                self.created.append(rid_a)
            up = self.rpc("api_getUpcomingLeaves", "", self.worker_id)
            n_up = len(up.result) if up.ok and isinstance(up.result, list) else None
            rep.add("W-1", "申請 A（全日・有給）: 同じ冪等キーで 2 回送っても 1 件（同じ reqId）",
                    bool(rid_a) and a1.result == a2.result and n_up == 1, f"1={a1.brief()} 2={a2.brief()} 予定 {n_up} 件",
                    "2 回とも同じ {ok:true, reqId}・予定 1 件")
            # 申請 B: 半日（午後）・特別
            rid_b = self.submit("W-2", "申請 B（半日・午後・特別休暇）", dict(
                base, leaveKubun=DAY_HALF, halfDayType=HALF_PM, leaveType=KIND_SPECIAL,
                specialReason=(reasons[0] if reasons else f"{TEST_MARK}{self.run_id}")))
            # 申請 C: 振替（同じ月次期間の休日を振替元に）
            pair = pick_substitute_pair(working, holidays)
            rid_c = None
            if pair:
                rid_c = self.submit("W-3", f"申請 C（振替。休暇日 {pair[0]}・振替元 {pair[1]}）",
                                    dict(base, leaveDate=pair[0], leaveType=KIND_SUBSTITUTE, substituteDate=pair[1]))
            else:
                rep.add("W-3", "申請 C（振替）", None, "同じ月次期間に振替元の候補がありません")
            # 申請 D: 過去申請。氏名に XSS の文字
            rid_d = None
            xss_name = self.worker_name + XSS_TEXT
            if past:
                rid_d = self.submit("W-4", f"申請 D（過去申請 {past[-1]}・氏名に XSS の文字）",
                                    dict(base, leaveDate=past[-1], isRetroactive=True, workerName=xss_name))
            else:
                rep.add("W-4", "申請 D（過去申請）", None, "今年度の過去の勤務日がありません")
            if not rid_a:
                return

            pending = self.rpc("api_getAllPendingRequests")
            mine = {it["reqId"]: it for it in (pending.result or []) if isinstance(it, dict) and
                    it.get("reqId") in self.created} if pending.ok else {}
            want = [x for x in (rid_a, rid_b, rid_c, rid_d) if x]
            rep.add("W-5", "作った申請が承認待ち一覧に出る（過去申請は isRetroactive:true）",
                    set(mine) == set(want) and (not rid_d or mine.get(rid_d, {}).get("isRetroactive") is True),
                    f"{sorted(mine)}", f"{sorted(want)}")
            for rid, label in ((rid_a, "A"), (rid_b, "B"), (rid_c, "C")):
                self.stage(f"W-6{label}", f"申請 {label} の行がシートに出る（申請中）", rid, status=ST_SUBMITTED,
                           at2_empty=True, jobs=[], worker_name=self.worker_name)
            self.stage("W-6D", "申請 D の行がシートに出る（氏名の XSS の文字がそのまま）", rid_d, status=ST_SUBMITTED,
                       at2_empty=True, jobs=[], worker_name=xss_name)

            # プレビュー
            pv = self.rpc("api_getLeaveRequestById", rid_a)
            res = pv.result if pv.ok and isinstance(pv.result, dict) else {}
            p = missing_keys(res, REQUEST_KEYS) if res else ["null が返りました"]
            if res and not p:
                for k, v in (("status", ST_SUBMITTED), ("leaveDate", leave), ("leaveKubun", DAY_FULL),
                             ("leaveType", KIND_PAID), ("workerName", self.worker_name), ("approvedAt", ""),
                             ("approvedAt2", "")):
                    if res.get(k) != v:
                        p.append(f"{k}={res.get(k)!r}（期待 {v!r}）")
            page = self.http("GET", f"/?page=preview&reqId={rid_a}", cookie=self.cookie)
            if page.status != 200 or rid_a not in page.text:
                p.append(f"?page=preview&reqId が開きません: {page.status}")
            rep.problems("W-7", "プレビュー: api_getLeaveRequestById と ?page=preview の画面", p,
                         "申請中・休暇日・区分・氏名が入れたとおり")
            if rid_d:
                pd = self.rpc("api_getLeaveRequestById", rid_d)
                got = (pd.result or {}).get("workerName") if pd.ok and isinstance(pd.result, dict) else None
                rep.add("W-8", "XSS の文字を入れた氏名が文字のまま返る", got == xss_name, _short(got), _short(xss_name))
            else:
                rep.add("W-8", "XSS の文字を入れた氏名が文字のまま返る", None, "申請 D がありません")

            r = self.rpc("api_approveLeaveRequest2", rid_a, self.me)
            rep.add("W-9", "申請中の 2 次承認は拒まれる", r.is_error(400, MSG["not_approved1"]), r.brief(),
                    "400 " + MSG["not_approved1"])

            # 1 次承認: A は手書き（小さな PNG）、B は電子印。承認済みをもう一度承認すると拒む（K-4）
            t1 = self.rpc("api_approveLeaveRequest", rid_a, tiny_png_data_url(), self.me)
            t2 = self.rpc("api_approveLeaveRequest", rid_a, tiny_png_data_url(), self.me)
            rep.add("W-10", "1 次承認 A（手書き）。2 回目は「すでに承認済み」（K-4）",
                    t1.ok is True and t1.result == {"ok": True, "reqId": rid_a} and
                    t2.is_error(400, MSG["already_approved"]), f"1={t1.brief()} 2={t2.brief()}",
                    "1={ok:true, reqId} 2=400 " + MSG["already_approved"])
            self.stage("W-11", "1 次承認 A の後: サイン画像の URL・承認完了メール・台帳、シートは承認済で 2 次待ち",
                       rid_a, status=ST_APPROVED, at2_empty=True, sign="set",
                       jobs=[(K_APPROVE1, {"sign": {"saved"}, "mail": {"sent"}, "ledger": {"upserted"}})])
            if rid_b:
                b1 = self.rpc("api_approveLeaveRequest", rid_b, "", self.me)
                rep.add("W-12", "1 次承認 B（電子印）", b1.ok is True and b1.result == {"ok": True, "reqId": rid_b},
                        b1.brief(), "{ok:true, reqId}")
                self.stage("W-13", "1 次承認 B の後: サインの列は空・承認完了メール・台帳", rid_b, status=ST_APPROVED,
                           at2_empty=True, sign="empty",
                           jobs=[(K_APPROVE1, {"sign": {"cleared"}, "mail": {"sent"}, "ledger": {"upserted"}})])
            p2 = self.rpc("api_getPending2ndApprovals")
            ids = [it.get("reqId") for it in (p2.result or []) if isinstance(it, dict)] if p2.ok else []
            rep.add("W-14", "1 次承認した A が 2 次承認待ちに出る", rid_a in ids, f"{len(ids)} 件", f"{rid_a} を含む")

            # 取消済みは承認できない
            victim = rid_d or rid_c
            if victim:
                d1 = self.rpc("api_deleteLeaveRequest", victim)
                d2 = self.rpc("api_deleteLeaveRequest", victim)
                ap = self.rpc("api_approveLeaveRequest", victim, "", self.me)
                if d1.ok:
                    self.canceled.add(victim)
                rep.add("W-15", f"取消（{victim}）→ もう一度の取消は「既に取消済み」→ 承認は「取消済み」（K-4）",
                        d1.result == {"ok": True} and d2.is_error(400, MSG["already_canceled"]) and
                        ap.is_error(400, MSG["canceled_approve"]), f"取消={d1.brief()} 再={d2.brief()} 承認={ap.brief()}",
                        "{ok:true} → 400 " + MSG["already_canceled"] + " → 400 " + MSG["canceled_approve"])
                self.stage("W-16", "取消した申請: シートは取消、台帳の行を消す仕事", victim, status=ST_CANCELED,
                           at2_empty=True, jobs=[(K_UNAPPROVE, {"ledger": {"removed"}})])
            else:
                rep.add("W-15", "取消済みは承認できない", None, "取消に使う申請（C か D）がありません")

            # 2 次承認 A。もう一度は拒む
            s1 = self.rpc("api_approveLeaveRequest2", rid_a, self.me)
            s2 = self.rpc("api_approveLeaveRequest2", rid_a, self.me)
            rep.add("W-17", "2 次承認 A。2 回目は「すでに2次承認済み」",
                    s1.ok is True and s1.result == {"ok": True, "reqId": rid_a} and
                    s2.is_error(400, MSG["already_approved2"]), f"1={s1.brief()} 2={s2.brief()}",
                    "1={ok:true, reqId} 2=400 " + MSG["already_approved2"])
            p2 = self.rpc("api_getPending2ndApprovals")
            ids = [it.get("reqId") for it in (p2.result or []) if isinstance(it, dict)] if p2.ok else []
            rep.add("W-18", "2 次承認した A は 2 次承認待ちから消える", rid_a not in ids, f"{len(ids)} 件", f"{rid_a} を含まない")
            self.stage("W-19", "2 次承認 A の後: PDF の URL・台帳、シートの APPROVED_AT2 が日時", rid_a, status=ST_APPROVED,
                       at2_empty=False, sign="set", pdf="set",
                       jobs=[(K_APPROVE1, {"sign": {"saved"}, "mail": {"sent"}, "ledger": {"upserted"}}),
                             (K_APPROVE2, {"pdf": {"created"}, "ledger": {"upserted"}})])

            # 編集で承認が外れる（B）
            if rid_b:
                e1 = self.rpc("api_editLeaveRequest", rid_b, {"leaveKubun": DAY_HALF, "halfDayType": HALF_AM})
                eb = self.rpc("api_getLeaveRequestById", rid_b)
                rb = eb.result if eb.ok and isinstance(eb.result, dict) else {}
                rep.add("W-20", "承認済みの B を編集すると申請中に戻り、承認の列が空になる",
                        e1.ok is True and e1.result == {"ok": True, "newStatus": ST_SUBMITTED} and
                        rb.get("status") == ST_SUBMITTED and rb.get("approvedAt") == "" and
                        rb.get("approvedBy1Email") == "" and rb.get("halfDayType") == HALF_AM,
                        f"編集={e1.brief()} 後={_short({k: rb.get(k) for k in ('status', 'approvedAt', 'halfDayType')})}",
                        "{ok:true, newStatus:申請中}・approvedAt 空・半日区分 午前")
                self.stage("W-21", "編集の後: サインの列を空・台帳の行を消す、シートは申請中", rid_b, status=ST_SUBMITTED,
                           at2_empty=True, sign="empty",
                           jobs=[(K_APPROVE1, {"sign": {"cleared"}, "mail": {"sent"}, "ledger": {"upserted"}}),
                                 (K_UNAPPROVE, {"sign": {"cleared"}, "ledger": {"removed"}})])

            # 取消（残り全部）
            results = {}
            for rid in (rid_a, rid_b, rid_c):
                if rid and rid not in self.canceled:
                    r = self.rpc("api_deleteLeaveRequest", rid)
                    results[rid] = r.brief()
                    if r.ok and r.result == {"ok": True}:
                        self.canceled.add(rid)
            rep.add("W-22", "残りの申請を取消（{ok:true}）", all(x in self.canceled for x in (rid_a, rid_b, rid_c) if x),
                    _short(results), "すべて {ok:true}")
            self.stage("W-23", "取消した A: シートは取消、PDF_URL は残る（asis）、台帳の行を消す", rid_a, status=ST_CANCELED,
                       at2_empty=False, pdf="set",
                       jobs=[(K_APPROVE1, {"sign": {"saved"}, "mail": {"sent"}, "ledger": {"upserted"}}),
                             (K_APPROVE2, {"pdf": {"created"}, "ledger": {"upserted"}}),
                             (K_UNAPPROVE, {"ledger": {"removed"}})])
            if rid_b:
                self.stage("W-24", "取消した B: シートは取消", rid_b, status=ST_CANCELED, at2_empty=True,
                           jobs=[(K_APPROVE1, {"sign": {"cleared"}, "mail": {"sent"}, "ledger": {"upserted"}}),
                                 (K_UNAPPROVE, {"sign": {"cleared"}, "ledger": {"removed"}}),
                                 (K_UNAPPROVE, {"ledger": {"removed"}})])
            if rid_c and rid_c != victim:
                self.stage("W-25", "取消した C: シートは取消", rid_c, status=ST_CANCELED, at2_empty=True,
                           jobs=[(K_UNAPPROVE, {"ledger": {"removed"}})])

            pending = self.rpc("api_getAllPendingRequests")
            left = [it.get("reqId") for it in (pending.result or []) if isinstance(it, dict) and
                    it.get("reqId") in self.created] if pending.ok else ["（読めません）"]
            up = self.rpc("api_getUpcomingLeaves", "", self.worker_id)
            up_left = [it.get("reqId") for it in (up.result or []) if isinstance(it, dict)] if up.ok else ["（読めません）"]
            rep.add("W-26", "取消した申請は承認待ち・予定の一覧から消える", not left and not up_left,
                    f"承認待ち {left} 予定 {up_left}", "どちらも 0 件")
            if self.probe is not None:
                sent = mails_sent(j for rid in self.created for j in self.probe.jobs_for(rid))
                approvals = 1 + (1 if rid_b else 0)
                rep.add("W-27", "承認完了メールは 1 次承認の回数だけ（ステージングの MAIL_TO に届く）", sent == approvals,
                        f"{sent} 通", f"{approvals} 通")
                rep.notes.append(f"承認完了メール: {sent} 通（ステージングの MAIL_TO＝オーナーの 1 件に届く）")
            else:
                rep.add("W-27", "承認完了メールの件数", None, "DB を見られません")
            self.write_ops()
        finally:
            self.cleanup()

    def write_ops(self) -> None:
        rep = self.report
        r = self.rpc("ops_syncSheetConfigNow")
        rep.add("F-ops_syncSheetConfigNow", "ops_syncSheetConfigNow() が {ok:true, …}",
                r.ok is True and isinstance(r.result, dict) and r.result.get("ok") is True, r.brief(), "{ok:true}")
        r = self.rpc("ops_runMasterSync")
        rep.add("F-ops_runMasterSync", "ops_runMasterSync() が {ok:true, tabs}（ステージングの 4 タブ）",
                r.ok is True and isinstance(r.result, dict) and r.result.get("ok") is True and "tabs" in r.result,
                r.brief(), "{ok:true, tabs}")
        r = self.rpc("ops_rebuildSummaryNow")
        rep.add("F-ops_rebuildSummaryNow", "ops_rebuildSummaryNow() が {ok:true, rows}",
                r.ok is True and isinstance(r.result, dict) and r.result.get("ok") is True, r.brief(), "{ok:true, rows}")

    def cleanup(self) -> None:
        """途中で止まっても、作った試験の申請を取消しておく（ステージングのデータだけ）。"""
        for rid in self.created:
            if rid in self.canceled:
                continue
            r = self.rpc("api_deleteLeaveRequest", rid)
            self.report.notes.append(f"後片付け: {rid} を取消 → {r.brief()}")
            if (r.ok and r.result == {"ok": True}) or r.is_error(400, MSG["already_canceled"]):
                self.canceled.add(rid)
        if self.created:
            self.report.notes.append("試験の申請（すべて取消済み）: " + ", ".join(self.created))

    def coverage(self) -> None:
        """34 関数をすべて呼んだか（ops_sendWarnMail は PM の拒否だけ。送信は JEV）。"""
        missing = [fn for fn in RPC_FUNCTIONS if fn not in self.times]
        both = bool(self.cookie and self.outsider_cookie)
        self.report.add("F-all", f"{len(RPC_FUNCTIONS)} 関数をすべて確かめの対象にした", (not missing) if both else None,
                        f"漏れ: {missing}" if both else "セッションが 2 つそろっていません", "漏れなし")


def _args_text(args: tuple) -> str:
    return "(" + ", ".join(json.dumps(a, ensure_ascii=False) for a in args) + ")"


# ===== 入口 =====


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="休暇届 VPS 版の受け入れテスト（ステージングだけ）")
    p.add_argument("--base-url", default="http://localhost:8000", help="api の URL（コンテナの中は http://api:8000）")
    p.add_argument("--web-url", default=None, help="nginx の URL（ヘッダーを見る。外からなら --base-url と同じ）")
    p.add_argument("--origin", default=os.environ.get("PUBLIC_ORIGIN", DEFAULT_ORIGIN))
    p.add_argument("--cookie", default=os.environ.get("ACCEPTANCE_COOKIE"), help="セッションのクッキーの値")
    p.add_argument("--outsider-cookie", default=os.environ.get("ACCEPTANCE_OUTSIDER_COOKIE"),
                   help="権限の無いアカウントのクッキーの値（staging_session.py mint --outsider）")
    p.add_argument("--mint-session", action="store_true", help="staging でセッションを 2 つ作り、最後に消す（コンテナの中）")
    p.add_argument("--email", default=None, help="--mint-session のメール（BOOTSTRAP_ADMINS の中）")
    p.add_argument("--write", action="store_true", help="書き込みの試験（staging だけ）")
    p.add_argument("--job-wait", type=float, default=300.0, help="GAS の仕事・行の反映を待つ秒")
    p.add_argument("--dept", default=None, help="試験の申請の部署（既定は部署の一覧の 1 つ目）")
    p.add_argument("--json-out", default=None, help="結果の JSON の書き出し先")
    p.add_argument("--timeout", type=float, default=120.0)
    return p


def main(argv: list[str] | None = None, *, api: Transport | None = None, web: Transport | None = None,
         probe: DbProbe | None = None, minter: Callable[[bool], tuple[str, Callable[[], None]]] | None = None,
         runner_kw: dict | None = None) -> int:
    """minter（試験用）: outsider を受けて (クッキーの値, 消す関数) を返す。"""
    args = build_parser().parse_args(argv)
    api = api or UrllibTransport(args.base_url, args.timeout)
    if web is None and args.web_url:
        web = UrllibTransport(args.web_url, args.timeout)
    runner = Runner(api, web=web, origin=args.origin, cookie=args.cookie, outsider_cookie=args.outsider_cookie,
                    probe=probe, write=args.write, job_wait=args.job_wait, dept=args.dept, **(runner_kw or {}))
    env = runner.check_health()
    if env != "staging":
        print(f"REFUSED: /healthz の env が staging ではありません（{env}）。受け入れテストはステージングでだけ動かします。",
              file=sys.stderr)
        return 2
    cleanups: list[Callable[[], None]] = []
    try:
        if args.mint_session:
            if minter is None:
                try:
                    minter, in_container_probe = _container_minter(args.email)
                except RuntimeError as exc:
                    print("REFUSED: " + str(exc), file=sys.stderr)
                    return 2
                runner.probe = runner.probe or in_container_probe
            try:
                main_cookie, undo = minter(False)
                cleanups.append(undo)
                outsider_cookie, undo = minter(True)
                cleanups.append(undo)
            except RuntimeError as exc:
                print("REFUSED: " + str(exc), file=sys.stderr)
                return 2
            runner.cookie = cookie_header(main_cookie)
            runner.outsider_cookie = cookie_header(outsider_cookie)
        try:
            report = runner.run()
        except Refused as exc:
            print("REFUSED: " + str(exc), file=sys.stderr)
            return 2
        runner.coverage()
        print(report.table())
        for note in report.notes:
            print(note)
        if args.json_out:
            Path(args.json_out).write_text(json.dumps(report.to_json(), ensure_ascii=False, indent=1), "utf-8")
        return report.exit_code()
    finally:
        for undo in reversed(cleanups):
            try:
                undo()
            except Exception as exc:  # noqa: BLE001 - 後片付けの失敗は出して続ける
                print(f"session cleanup failed: {type(exc).__name__}", file=sys.stderr)


def _container_minter(email: str | None):
    """コンテナの中で staging_session.py を使ってセッションを作る関数と、DB の probe を返す。"""
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import staging_session
    from app.config import get_settings
    from app.db import get_sessionmaker

    settings = get_settings()
    if settings.app_env != "staging":
        raise RuntimeError(f"APP_ENV が staging ではありません（{settings.app_env}）")
    sf = get_sessionmaker()

    def minter(outsider: bool):
        with sf() as db:
            try:
                m = staging_session.mint(db, settings, email, hours=2, outsider=outsider)
            except staging_session.Refused as exc:
                raise RuntimeError(str(exc)) from exc

        def undo() -> None:
            with sf() as db2:
                n = staging_session.revoke(db2, m.token_hash)
            print(f"session removed ({'outsider' if outsider else 'main'}): {n == 1}")

        return m.cookie_value, undo

    return minter, DbProbe(sf, settings)


if __name__ == "__main__":
    sys.exit(main())
