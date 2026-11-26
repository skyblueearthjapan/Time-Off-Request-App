"""受け入れの道具（vps-app/tools/qa/）の試験。偽の HTTP（FakeLeaveServer）で流れが通ること、本番では止まること。

FakeLeaveServer は実装仕様 4-3・8-5〜8-8 と K-4 の決まり（申請・承認・2 次承認・編集・取消・権限・入口の確かめ）を
小さく真似る。返す文は道具の MSG を使い、MSG が domain のソースにそのままあることは別の試験で確かめる。
"""
from __future__ import annotations

import json
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import parse_qs, quote, urlsplit

import pytest

QA_DIR = Path(__file__).resolve().parents[2] / "tools" / "qa"
if str(QA_DIR) not in sys.path:
    sys.path.insert(0, str(QA_DIR))

import acceptance as acc  # noqa: E402
import compare_views as cv  # noqa: E402
import staging_session  # noqa: E402

from app import rpc  # noqa: E402
from app.domain import register as domain_register  # noqa: E402
from conftest import make_admin, make_settings  # noqa: E402

JST = timezone(timedelta(hours=9))
NOW = datetime(2026, 11, 26, 10, 0, tzinfo=JST)
TODAY = date(2026, 11, 26)
WORKING = ["2026-11-26", "2026-11-27", "2026-11-30", "2026-12-01"]
HOLIDAYS = [{"date": "2026-11-21", "label": "土曜"}, {"date": "2026-11-22", "label": "日曜"},
            {"date": "2026-11-28", "label": "土曜"}]
PAST = {"2026-11": [{"date": "2026-11-24", "dow": 2, "label": "2026-11-24（火）"},
                    {"date": "2026-11-25", "dow": 3, "label": "2026-11-25（水）"}]}
MAIN, OUTSIDER = "owner@lineworks-local.info", "acceptance-outsider@lineworks-local.info"
DEPT = "製造部"


class RpcErr(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(message)
        self.status = status
        self.message = message


def _json(status: int, body, headers=None) -> acc.HttpResponse:
    return acc.HttpResponse(status, {"content-type": "application/json", **(headers or {})},
                            json.dumps(body, ensure_ascii=False))


class FakeLeaveServer:
    """道具が呼ぶ HTTP の入口と 34 関数を、状態つきで小さく真似る。"""

    def __init__(self, env: str = "staging", seed: bool = False, break_k4: bool = False):
        self.env = env
        self.break_k4 = break_k4
        self.calls: list[str] = []
        self.reqs: dict[str, dict] = {}
        self.jobs: dict[str, list[dict]] = {}
        self.idem: dict[tuple, str] = {}
        self.seq = 0
        self.job_seq = 0
        if seed:
            self._new({"leaveDate": "2026-11-27", "workerName": "山田", "workerId": "101", "deptName": DEPT,
                       "leaveType": acc.KIND_PAID, "leaveKubun": acc.DAY_FULL}, status=acc.ST_SUBMITTED)
            self._new({"leaveDate": "2026-11-30", "workerName": "佐藤", "workerId": "102", "deptName": DEPT,
                       "leaveType": acc.KIND_PAID, "leaveKubun": acc.DAY_FULL}, status=acc.ST_APPROVED)

    # ---- HTTP ----

    def __call__(self, method: str, path: str, headers: dict, body: bytes | None) -> acc.HttpResponse:
        self.calls.append(f"{method} {path}")
        url = urlsplit(path)
        user = {"__Host-kyuka_session=main": MAIN, "__Host-kyuka_session=outsider": OUTSIDER}.get(
            headers.get("Cookie", ""))
        if url.path == "/healthz":
            return _json(200, {"ok": True, "env": self.env})
        if url.path == "/auth/me":
            return _json(200, {"ok": True, "email": user, "name": ""}) if user else _json(401, {"ok": False})
        if url.path == "/" and method == "GET":
            if not user:
                return acc.HttpResponse(302, {"location": "/auth/login?next=" + quote(path, safe="")}, "")
            q = parse_qs(url.query)
            page = q.get("page", ["top"])[0]
            if page == "somuAdmin" and user == OUTSIDER:
                html = "<title>アクセス権限なし</title>このページは総務部・管理者のみアクセスできます。"
            else:
                html = (f"<html><head><title>{acc.PAGE_TITLES.get(page, '休暇届 申請TOP')}</title></head><body>"
                        f"<script src=\"/static/js/top.js\"></script>{q.get('reqId', [''])[0]}</body></html>")
            return acc.HttpResponse(200, {"cache-control": "no-store"}, html)
        if url.path.startswith("/api/rpc/"):
            return self.rpc(url.path[len("/api/rpc/"):], headers, body, user)
        return _json(404, {"ok": False})

    def rpc(self, fn: str, headers: dict, body: bytes | None, user: str | None) -> acc.HttpResponse:
        if not user:
            return _json(401, {"ok": False, "error": acc.RPC_MSG["login"], "login": "/auth/login"})
        if headers.get("Origin") != acc.DEFAULT_ORIGIN:
            return _json(403, {"ok": False, "error": acc.RPC_MSG["origin"]})
        if headers.get("Content-Type") != "application/json":
            return _json(415, {"ok": False, "error": acc.RPC_MSG["shape"]})
        if fn not in acc.RPC_FUNCTIONS:
            return _json(404, {"ok": False, "error": acc.RPC_MSG["unknown"]})
        args = json.loads(body or b"[]")
        if not isinstance(args, list):
            return _json(400, {"ok": False, "error": acc.RPC_MSG["shape"]})
        key = headers.get("X-Idempotency-Key")
        if key and (user, key) in self.idem:
            return acc.HttpResponse(200, {}, self.idem[(user, key)])
        try:
            if fn.startswith("ops_") and user != MAIN:
                raise RpcErr(403, acc.RPC_MSG["not_operator"])
            result = getattr(self, fn)(user, *args)
        except RpcErr as e:
            return _json(e.status, {"ok": False, "error": e.message})
        resp = _json(200, {"ok": True, "result": result})
        if key:
            self.idem[(user, key)] = resp.text
        return resp

    # ---- 状態 ----

    def _new(self, data: dict, status: str = acc.ST_SUBMITTED) -> str:
        self.seq += 1
        rid = f"LV-20261126-{self.seq:03d}"
        self.reqs[rid] = {"reqId": rid, "status": status, "deptName": data.get("deptName", ""),
                          "workerId": data.get("workerId", ""), "workerName": data.get("workerName", ""),
                          "leaveKubun": data.get("leaveKubun", ""), "halfDayType": data.get("halfDayType", ""),
                          "leaveType": data.get("leaveType", ""), "leaveDate": data.get("leaveDate", ""),
                          "isRetroactive": data.get("isRetroactive") is True, "approvedAt": "",
                          "approvedBy1Email": "", "approvedAt2": "", "sign": "", "pdf": "", "row": 300 + self.seq,
                          "fy": "2026"}
        self.jobs[rid] = []
        return rid

    def _job(self, rid: str, kind: str, result: dict) -> None:
        self.job_seq += 1
        self.jobs[rid].append({"id": self.job_seq, "kind": kind, "state": "done", "error": "", "result": result})

    def _find(self, rid) -> dict:
        r = self.reqs.get(rid)
        if r is None:
            raise RpcErr(400, acc.MSG["not_found_prefix"] + str(rid))
        return r

    def _brief(self, r: dict, keys) -> dict:
        out = {k: r.get(k, "") for k in keys}
        out["leaveDate"] = r["leaveDate"].replace("-", "/")
        out["daysAgo"] = 0
        out["submittedAt"] = "11/26 10:00"
        return {k: out[k] for k in keys}

    # ---- 読む関数 ----

    def api_isAdmin(self, user):
        return user == MAIN

    def api_isSomu(self, user):
        return user == MAIN

    def api_isDeptApprover(self, user):
        return False

    def api_getDeptList(self, user):
        return [{"deptId": DEPT, "deptName": DEPT}]

    def api_getWorkersByDept(self, user, dept=None):
        return [] if not dept else [{"workerId": "101", "workerName": "山田", "deptId": dept},
                                    {"workerId": "102", "workerName": "佐藤", "deptId": dept}]

    def api_getWorkerInfo(self, user):
        return None

    def api_getLookupData(self, user):
        return {"leaveKubun": ["全日休", "半日休"], "halfDayType": ["午前", "午後"],
                "leaveTypeFull": ["有給消化（私用）", "振替休暇", "特別休暇"],
                "leaveTypeHalf": ["有給消化（私用）", "特別休暇"], "specialReasons": ["結婚休暇"]}

    def api_getWorkingDays(self, user):
        return list(WORKING)

    def api_getHolidayDays(self, user, worker=None):
        return list(HOLIDAYS)

    def api_getPastWorkingDays(self, user, ym=None):
        return list(PAST.get(ym, []))

    def api_getUpcomingLeaves(self, user, dept=None, worker=None):
        return [self._brief(r, acc.UPCOMING_KEYS) for r in self.reqs.values()
                if r["status"] != acc.ST_CANCELED and r["leaveDate"] >= TODAY.isoformat()
                and (not dept or r["deptName"] == dept) and (not worker or r["workerId"] == worker)]

    def api_getAllPendingRequests(self, user):
        return [self._brief(r, acc.PENDING_KEYS) for r in self.reqs.values() if r["status"] == acc.ST_SUBMITTED]

    def api_getApproverList(self, user):
        return [{"email": MAIN, "name": "今泉"}]

    def api_getCurrentApproverInfo(self, user):
        return {"email": user, "name": "今泉"}

    def _somu(self, user):
        if user != MAIN:
            raise RpcErr(400, acc.MSG["somu_only"])

    def api_getPending2ndApprovals(self, user):
        self._somu(user)
        return [{k: v for k, v in self._brief(r, acc.UPCOMING_KEYS).items() if k != "halfDayType"} | {"approvedAt": ""}
                for r in self.reqs.values() if r["status"] == acc.ST_APPROVED and not r["approvedAt2"]]

    def api_getSomuAdminView(self, user, fy=None, dept=None):
        self._somu(user)
        return []

    def api_getApproverProfiles(self, user):
        return {"isAdmin": user == MAIN, "currentEmail": user, "profiles": []}

    def api_getDeptAdminView(self, user, fy=None, depts=None):
        if user != MAIN:
            raise RpcErr(400, acc.MSG["dept_approver_only"])
        return {"workers": [], "deptNames": depts or []}

    def api_getAdminBossView(self, user, dept=None, fy=None):
        if user != MAIN:
            raise RpcErr(400, acc.MSG["admin_page_only"])
        rows = [r for r in self.reqs.values() if r["status"] == acc.ST_APPROVED and str(fy) == r["fy"]]
        if not rows:
            return []
        paid = sum(1 for r in rows if r["leaveType"] == acc.KIND_PAID)
        return [{"deptId": DEPT, "deptName": DEPT, "workerId": "102", "workerName": "佐藤", "paidCount": paid,
                 "paidHalfCount": 0, "substituteCount": 0, "specialCount": 0, "totalCount": len(rows),
                 "lastPaidDate": "", "warnLevel": ""}]

    def api_getAdminSomuView(self, user, fy=None, dept=None):
        return self.api_getAdminBossView(user, dept, fy)

    def api_getLeaveStatusSearch(self, user, fy=None, dept=None, worker=None):
        if worker:
            return {"mode": "personal", "fyYear": fy, "worker": None, "milestones": [], "upcoming": []}
        return {"mode": "list", "fyYear": fy, "deptFilter": dept or "", "workers": []}

    def api_getLeaveRequestById(self, user, rid=None):
        r = self.reqs.get(rid) if rid else None
        if r is None:
            return None
        out = {k: "" for k in acc.REQUEST_KEYS}
        out.update({k: r[k] for k in ("reqId", "status", "workerName", "leaveKubun", "halfDayType", "leaveType",
                                      "leaveDate", "approvedAt", "approvedBy1Email", "approvedAt2")})
        out.update({"signImageUrl": r["sign"], "pdfUrl": r["pdf"], "rowNo": r["row"]})
        return out

    def ops_getStatus(self, user):
        return {"rowSync": {"blocked": []}, "gasJobs": {"counts": {}}, "gasPing": {"ok": True, "problems": []},
                "sheetWriteGate": {"allowed": True}, "lastRuns": {}}

    # ---- 書く関数 ----

    def api_submitLeaveRequest(self, user, data=None):
        data = data or {}
        leave = str(data.get("leaveDate") or "")
        d = acc.parse_ymd(leave)
        if d is None:
            return {"ok": False, "error": acc.MSG["bad_format"]}
        if data.get("isRetroactive") is True:
            if not (acc.fy_start(TODAY) <= d <= TODAY):
                return {"ok": False, "error": acc.MSG["retro_out_of_fy"]}
        elif leave not in WORKING:
            return {"ok": False, "error": acc.MSG["not_working_day"]}
        if data.get("leaveType") == acc.KIND_SUBSTITUTE:
            sub = acc.parse_ymd(data.get("substituteDate") or "")
            if sub is None:
                return {"ok": False, "error": acc.MSG["substitute_missing"]}
            if acc.monthly_period(sub) != acc.monthly_period(d):
                return {"ok": False, "error": acc.MSG["substitute_period"] + "\n休暇日: …"}
        return {"ok": True, "reqId": self._new(data)}

    def api_approveLeaveRequest(self, user, rid=None, sign=None, email=None):
        if not rid:
            raise RpcErr(400, acc.MSG["no_req_id"])
        r = self._find(rid)
        if r["status"] == acc.ST_APPROVED and not self.break_k4:
            raise RpcErr(400, acc.MSG["already_approved"])
        if r["status"] == acc.ST_CANCELED:
            raise RpcErr(400, acc.MSG["canceled_approve"])
        r.update(status=acc.ST_APPROVED, approvedAt="2026-11-26 10:00:00", approvedBy1Email=email or "",
                 sign="https://drive.google.com/sign" if sign else "")
        self._job(rid, acc.K_APPROVE1, {"sign": {"status": "saved" if sign else "cleared"},
                                        "mail": {"status": "sent"}, "ledger": {"status": "upserted"}})
        return {"ok": True, "reqId": rid}

    def api_approveLeaveRequest2(self, user, rid=None, email=None):
        if not rid:
            raise RpcErr(400, acc.MSG["no_req_id"])
        self._somu(user)
        r = self._find(rid)
        if r["status"] != acc.ST_APPROVED:
            raise RpcErr(400, acc.MSG["not_approved1"])
        if r["approvedAt2"]:
            raise RpcErr(400, acc.MSG["already_approved2"])
        r.update(approvedAt2="2026-11-26 10:05:00", pdf="https://drive.google.com/pdf")
        self._job(rid, acc.K_APPROVE2, {"pdf": {"status": "created"}, "ledger": {"status": "upserted"}})
        return {"ok": True, "reqId": rid}

    def api_editLeaveRequest(self, user, rid=None, patch=None):
        r = self.reqs.get(rid)
        if r is None:
            raise RpcErr(400, acc.MSG["not_found"])
        patch = patch or {}
        if patch.get("leaveKubun"):
            r["leaveKubun"] = patch["leaveKubun"]
            r["halfDayType"] = patch.get("halfDayType", "")
        if r["status"] == acc.ST_APPROVED:
            r.update(status=acc.ST_SUBMITTED, approvedAt="", approvedBy1Email="", approvedAt2="", sign="")
            self._job(rid, acc.K_UNAPPROVE, {"sign": {"status": "cleared"}, "ledger": {"status": "removed"}})
        return {"ok": True, "newStatus": r["status"]}

    def api_deleteLeaveRequest(self, user, rid=None):
        r = self.reqs.get(rid)
        if r is None:
            raise RpcErr(400, acc.MSG["not_found"])
        if r["status"] == acc.ST_CANCELED:
            raise RpcErr(400, acc.MSG["already_canceled"])
        r["status"] = acc.ST_CANCELED
        self._job(rid, acc.K_UNAPPROVE, {"sign": {"status": "skipped"}, "ledger": {"status": "removed"}})
        return {"ok": True}

    def ops_retryJob(self, user, job_id=None):
        raise RpcErr(400, acc.MSG["bad_job_id"])

    def ops_revokeSessions(self, user, email=None):
        if not email:
            raise RpcErr(400, acc.MSG["no_email"])
        return {"ok": True, "revoked": 0}

    def ops_syncSheetConfigNow(self, user):
        return {"ok": True, "counts": {}}

    def ops_runMasterSync(self, user):
        return {"ok": True, "tabs": {}}

    def ops_rebuildSummaryNow(self, user):
        return {"ok": True, "rows": 1}

    def ops_sendWarnMail(self, user):  # pragma: no cover - 道具は運用者として呼ばない
        raise AssertionError("ops_sendWarnMail を運用者として呼びました（メールが出る）")


class FakeProbe:
    """DbProbe の代わり（FakeLeaveServer の状態を DB とシートとして読む）。"""

    def __init__(self, server: FakeLeaveServer):
        self.server = server

    def request(self, rid):
        r = self.server.reqs.get(rid)
        if r is None:
            return None
        return {"status": r["status"], "approvedAt": bool(r["approvedAt"]), "approvedAt2": bool(r["approvedAt2"]),
                "signImageUrl": r["sign"], "pdfUrl": r["pdf"], "sheetRow": r["row"], "signPng": False,
                "workerName": r["workerName"]}

    def jobs_for(self, rid):
        return [dict(j) for j in self.server.jobs.get(rid, [])]

    def synced(self, rid):
        return rid in self.server.reqs

    def sheet_row(self, rid):
        r = self.server.reqs[rid]
        return {"REQ_ID": rid, "STATUS": r["status"], "LEAVE_DATE": 46352.625,
                "APPROVED_AT2": 46352.4 if r["approvedAt2"] else "", "SIGN_IMAGE_URL": r["sign"], "PDF_URL": r["pdf"],
                "WORKER_NAME": r["workerName"]}


def _minter(log: list):
    def mint(outsider: bool):
        name = "outsider" if outsider else "main"
        log.append("mint " + name)
        return name, lambda: log.append("revoke " + name)
    return mint


def _run(server, argv, probe="auto", log=None):
    log = [] if log is None else log
    code = acc.main(argv, api=server, probe=FakeProbe(server) if probe == "auto" else probe, minter=_minter(log),
                    runner_kw={"now": NOW, "sleep": lambda s: None})
    return code, log


# ===== 閉じたリスト・文 =====


def test_function_lists_match_app():
    assert acc.SCREEN_FUNCTIONS == rpc.SCREEN_FUNCTIONS
    assert acc.OPS_FUNCTIONS == rpc.OPS_FUNCTIONS
    assert acc.WRITE_FUNCTIONS == frozenset(domain_register.WRITES)


def test_messages_are_the_domain_originals():
    api_dir = Path(__file__).resolve().parents[1] / "app"
    domain = "".join(p.read_text("utf-8") for p in (api_dir / "domain").glob("*.py"))
    for name, text in acc.MSG.items():
        assert text in domain, f"MSG[{name!r}] が app/domain にありません: {text}"
    rpc_src = (api_dir / "rpc.py").read_text("utf-8")
    for name, text in acc.RPC_MSG.items():
        assert text in rpc_src, f"RPC_MSG[{name!r}] が app/rpc.py にありません: {text}"
    pages = (api_dir / "pages.py").read_text("utf-8")
    assert acc.SOMU_NO_AUTH_TEXT in pages
    for page, title in acc.NO_AUTH_TITLES.items():
        assert f'"{page}": ("{title}"' in pages
    templates = "".join(p.read_text("utf-8") for p in (api_dir / "templates" / "pages").glob("*.html"))
    for title in acc.PAGE_TITLES.values():
        assert f"<title>{title}</title>" in templates


# ===== 本番では止まる =====


def test_acceptance_refuses_production_without_calls():
    server = FakeLeaveServer(env="production")
    code, log = _run(server, ["--mint-session", "--write"])
    assert code == 2
    assert server.calls == ["GET /healthz"]  # healthz のほかは何も呼ばない
    assert log == []  # セッションも作らない


def test_compare_views_refuses_production():
    server = FakeLeaveServer(env="production")
    assert cv.main(["--cookie", "main"], api=server, now=NOW) == 2
    assert server.calls == ["GET /healthz"]


def test_staging_session_refuses_production(db):
    prod = SimpleNamespace(app_env="production", bootstrap_admin_list=[MAIN], allowed_domain="lineworks-local.info",
                           session_max_age_days=90)
    with pytest.raises(staging_session.Refused):
        staging_session.mint(db, prod)


def test_staging_session_mint_and_outsider(db):
    from app.models import UserSession

    s = make_settings()
    m = staging_session.mint(db, s, hours=1)
    assert m.email == MAIN
    o = staging_session.mint(db, s, hours=1, outsider=True)
    assert o.email == OUTSIDER
    with pytest.raises(staging_session.Refused):
        staging_session.mint(db, s, email="someone@lineworks-local.info")
    assert db.query(UserSession).count() == 2
    assert staging_session.revoke(db, m.token_hash) == 1
    assert staging_session.purge(db) == 1


def test_staging_session_outsider_refused_when_listed(db):
    make_admin(db, OUTSIDER)
    with pytest.raises(staging_session.Refused):
        staging_session.mint(db, make_settings(), outsider=True)


# ===== 流れ（偽の HTTP） =====


def test_full_flow_passes_and_cleans_up(capsys):
    server = FakeLeaveServer()
    code, log = _run(server, ["--mint-session", "--write", "--job-wait", "0.01"])
    out = capsys.readouterr().out
    assert code == 0, out
    assert "FAIL 0" in out and "PASS" in out
    assert log == ["mint main", "mint outsider", "revoke outsider", "revoke main"]
    created = list(server.reqs)
    assert len(created) == 4  # A・B・C・D（境目の確かめと二度押しでは増えない）
    assert all(r["status"] == acc.ST_CANCELED for r in server.reqs.values())
    assert all(r["workerName"].startswith(acc.TEST_MARK) for r in server.reqs.values())
    assert "承認完了メール: 2 通" in out
    assert "試験の申請（すべて取消済み）" in out
    for item_id in ("W-1", "W-10", "W-11", "W-15", "W-17", "W-19", "W-20", "W-21", "W-23", "W-27", "PM-1", "PM-7",
                    "B-6", "N-1", "F-all"):
        assert f"| {item_id} | PASS |" in out, item_id


def test_read_only_run_writes_nothing(capsys):
    server = FakeLeaveServer()
    code, _ = _run(server, ["--mint-session"])
    out = capsys.readouterr().out
    assert code == 0, out
    assert server.reqs == {}
    assert "| W-" not in out


def test_k4_regression_is_reported_as_fail(capsys):
    server = FakeLeaveServer(break_k4=True)  # 承認済みをもう一度承認できてしまう
    code, log = _run(server, ["--mint-session", "--write", "--job-wait", "0.01"])
    out = capsys.readouterr().out
    assert code == 1
    assert "| W-10 | FAIL |" in out
    assert log[-2:] == ["revoke outsider", "revoke main"]  # 不合格でもセッションは消す
    assert all(r["status"] == acc.ST_CANCELED for r in server.reqs.values())  # 後片付けで取消


def test_without_probe_db_items_are_skipped(capsys):
    server = FakeLeaveServer()
    code = acc.main(["--cookie", "main", "--outsider-cookie", "outsider", "--write", "--job-wait", "0.01"],
                    api=server, runner_kw={"now": NOW, "sleep": lambda s: None})
    out = capsys.readouterr().out
    assert code == 0, out
    assert "| W-11 | SKIP |" in out and "| W-27 | SKIP |" in out


# ===== 小さな決まり =====


def test_monthly_period_and_substitute_pair():
    assert acc.monthly_period(date(2026, 11, 15)) == (date(2026, 10, 16), date(2026, 11, 15))
    assert acc.monthly_period(date(2026, 11, 16)) == (date(2026, 11, 16), date(2026, 12, 15))
    assert acc.monthly_period(date(2026, 12, 20)) == (date(2026, 12, 16), date(2027, 1, 15))
    assert acc.monthly_period(date(2027, 1, 10)) == (date(2026, 12, 16), date(2027, 1, 15))
    assert acc.fiscal_year(date(2027, 3, 15)) == 2026 and acc.fiscal_year(date(2027, 3, 16)) == 2027
    assert acc.pick_substitute_pair(["2026-11-16"], [{"date": "2026-11-15"}, {"date": "2026-11-21"}]) == (
        "2026-11-16", "2026-11-21")
    assert acc.pick_substitute_pair(["2026-11-16"], [{"date": "2026-11-15"}]) is None


def test_sheet_row_problems_matches_unified_mail_conditions():
    ok = {"REQ_ID": "LV-1", "STATUS": "承認済", "LEAVE_DATE": 46352.625, "APPROVED_AT2": "", "SIGN_IMAGE_URL": "u"}
    assert acc.sheet_row_problems(ok, status="承認済", at2_empty=True, sign="set") == []
    assert acc.sheet_row_problems(dict(ok, LEAVE_DATE="2026-11-27"), status="承認済", at2_empty=True)
    assert acc.sheet_row_problems(dict(ok, APPROVED_AT2=46352.5), status="承認済", at2_empty=True)
    assert acc.sheet_row_problems(ok, status="承認済", at2_empty=False)
    assert acc.sheet_row_problems(ok, status="申請中", at2_empty=True)
    assert acc.sheet_row_problems(None, status="申請中", at2_empty=True) == ["シートの行が見つかりません"]


def test_job_problems():
    jobs = [{"id": 2, "kind": acc.K_APPROVE2, "state": "done", "result": {"pdf": {"status": "created"}}},
            {"id": 1, "kind": acc.K_APPROVE1, "state": "done", "result": {"mail": {"status": "sent"}}}]
    want = [(acc.K_APPROVE1, {"mail": {"sent"}}), (acc.K_APPROVE2, {"pdf": {"created"}})]
    assert acc.job_problems(jobs, want) == []
    assert acc.job_problems(jobs[:1], want)
    jobs[0]["state"] = "failed"
    assert acc.job_problems(jobs, want)
    assert acc.mails_sent(jobs) == 1


def test_tiny_png_is_accepted_by_domain():
    from app.domain.approval import decode_sign

    data = decode_sign(acc.tiny_png_data_url())
    assert data is not None and data.startswith(b"\x89PNG")


# ===== compare_views =====


class FakeCounts:
    def __init__(self, server: FakeLeaveServer):
        self.server = server

    def req_ids(self):
        return sorted(self.server.reqs)

    def fiscal_years(self):
        return [2026]

    def counts(self, today):
        rs = list(self.server.reqs.values())
        return {"byStatus": {}, "pending": sum(r["status"] == acc.ST_SUBMITTED for r in rs),
                "pending2": sum(r["status"] == acc.ST_APPROVED and not r["approvedAt2"] for r in rs),
                "upcoming": sum(r["status"] != acc.ST_CANCELED and r["leaveDate"] >= today.isoformat() for r in rs),
                "perFy": {"2026": {"approved": sum(r["status"] == acc.ST_APPROVED for r in rs), "paid": 1.0}}}


def test_compare_views_plan_saves_and_matches_gas(tmp_path, capsys):
    server = FakeLeaveServer(seed=True)
    out, plan_out = tmp_path / "vps.json", tmp_path / "plan.json"
    code = cv.main(["--cookie", "main", "--out", str(out), "--plan-out", str(plan_out)], api=server,
                   counts_probe=FakeCounts(server), now=NOW)
    text = capsys.readouterr().out
    assert code == 0, text
    plan = json.loads(plan_out.read_text("utf-8"))
    keys = [c["key"] for c in plan]
    assert len(keys) == len(set(keys))
    assert cv.call_key("api_getLeaveRequestById", ("LV-20261126-001",)) in keys
    assert cv.call_key("api_getLeaveStatusSearch", (2026, DEPT, "102")) in keys
    assert cv.call_key("api_getPastWorkingDays", ("2026-03",)) in keys
    for item in ("C-1", "C-2", "C-3", "C-4-2026"):
        assert f"| {item} | PASS |" in text, item
    assert all(not c["fn"].startswith("ops_") for c in plan)

    # GAS 版の返事（同じ値。承認待ちは並びが違っても同じとみなす＝K-10）
    vps = json.loads(out.read_text("utf-8"))["results"]
    gas = {"results": json.loads(json.dumps(vps))}
    pend = cv.call_key("api_getAllPendingRequests", ())
    gas["results"][pend]["result"] = list(reversed(gas["results"][pend]["result"]))
    gas_file = tmp_path / "gas.json"
    gas_file.write_text(json.dumps(gas, ensure_ascii=False), "utf-8")
    assert cv.main(["--cookie", "main", "--gas-file", str(gas_file)], api=server,
                   counts_probe=FakeCounts(server), now=NOW) == 0
    assert "| G-1 | PASS |" in capsys.readouterr().out

    # 1 か所違えば不合格
    gas["results"][cv.call_key("api_getLeaveRequestById", ("LV-20261126-001",))]["result"]["status"] = "承認済"
    gas_file.write_text(json.dumps(gas, ensure_ascii=False), "utf-8")
    assert cv.main(["--cookie", "main", "--gas-file", str(gas_file)], api=server,
                   counts_probe=FakeCounts(server), now=NOW) == 1
    text = capsys.readouterr().out
    assert "| G-1 | FAIL |" in text and "$.status" in text


def test_compare_views_counts_mismatch_fails(capsys):
    server = FakeLeaveServer(seed=True)

    class Wrong(FakeCounts):
        def counts(self, today):
            c = super().counts(today)
            c["pending"] += 1
            return c

    assert cv.main(["--cookie", "main"], api=server, counts_probe=Wrong(server), now=NOW) == 1
    assert "| C-1 | FAIL |" in capsys.readouterr().out


def test_diff_and_sample():
    assert cv.diff({"a": 1, "b": [1, 2]}, {"b": [1.0, 2], "a": 1.0}) == []
    assert cv.diff({"a": True}, {"a": 1})
    assert cv.diff([1], [1, 2]) == ["$: 件数 1 / 2"]
    assert cv.sample(list(range(10)), 3) == [0, 4, 9] or cv.sample(list(range(10)), 3) == [0, 5, 9]
    assert cv.sample([1, 2], 3) == [1, 2]
    assert cv.sample(list(range(5)), 0) == list(range(5))
    assert cv.months_of_fy(TODAY)[0] == "2026-03" and cv.months_of_fy(TODAY)[-1] == "2026-11"
