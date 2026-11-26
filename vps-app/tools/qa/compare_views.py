"""読むだけの関数の返事を総当たりで呼んで保存し、GAS 版の返事か DB から独立に数えた値と比べる（実装仕様 12-3 A-3・A-10）。

ステージングでだけ動く（/healthz の env が staging でなければ終了コード 2）。何も書かない（読む関数だけを呼ぶ）。

呼ぶもの（plan）:
  - いつも: api_isAdmin・api_isSomu・api_isDeptApprover・api_getDeptList・api_getWorkerInfo・api_getLookupData・
    api_getWorkingDays・api_getApproverList・api_getCurrentApproverInfo・api_getApproverProfiles・
    api_getAllPendingRequests（承認待ち）・api_getPending2ndApprovals（2 次承認待ち）・api_getHolidayDays('')・
    api_getUpcomingLeaves('', '')（予定）
  - 今年度の月ごと: api_getPastWorkingDays(yyyy-MM)
  - 部署ごと: api_getWorkersByDept・api_getUpcomingLeaves(部署, '')
  - 年度ごと（全部署と部署ごと）: api_getSomuAdminView・api_getDeptAdminView・api_getAdminSomuView・
    api_getAdminBossView・api_getLeaveStatusSearch（取得状況の一覧）
  - 作業員の抜き取り（部署ごとに --workers-per-dept 人。0 で全員）: api_getHolidayDays(作業員)（振替元の候補）・
    api_getUpcomingLeaves(部署, 作業員)・年度ごとの api_getLeaveStatusSearch(年度, 部署, 作業員)（取得状況の個人）
  - DB を見られるとき（--mint-session）: 全 REQ_ID の api_getLeaveRequestById（preview）
  年度は DB の FY_START_YEAR の値すべて（DB を見られないときは --years。既定は去年度と今年度）。

動かし方（VPS。api コンテナの中。tools はイメージに入っていないので付けて動かす＝罠 5）:

  sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro \
      -v /opt/kyuka/cutover:/out api python tools/qa/compare_views.py --base-url http://api:8000 --mint-session \
      --out /out/compare_vps_$(date +%Y%m%d_%H%M).json --plan-out /out/compare_plan.json [--gas-file /out/gas.json]"

  - --out: VPS 版の返事（plan と結果）を JSON で保存する。
  - --plan-out: GAS 版で同じ呼び出しをするための plan（[{key, fn, args}]）だけを保存する。
  - --gas-file: GAS 版の返事を保存したファイル（下の取り方）。あれば key ごとに突き合わせ、違いを表に出す。
  - DB を見られるとき（--mint-session）は、DB から独立に数えた値（状態別の件数・承認待ち・2 次承認待ち・予定・
    年度ごとの承認済みの件数と有給の回数）を返事の件数・合計と比べる。
  終了コード: 0＝違いなし、1＝違いあり（GAS と違う・独立に数えた値と違う・呼び出しが失敗した）、2＝止めた。

GAS 版の返事の取り方（ステージングの写しの GAS。移転モードでないこと。同じ日・同じデータのうちに）:
  1. 上のコマンドで --plan-out を作り、手元に持ってくる（scp か cat）。
  2. 写しの GAS（scriptId 1T3bWvTY…）のエディタで、一時ファイル `CompareExport.gs` を作り、下を貼る。
     PLAN に plan の JSON（配列）を貼る。6 分で止まるので、REQ_ID の preview が多いときは FROM・TO で分けて
     何回か実行する（ファイルが回数分できる。--gas-file は複数を受ける）。
  3. 関数 exportViewsForCompare を実行する（オーナーのアカウント。VPS のセッションと同じ BOOTSTRAP_ADMINS の人）。
     ログに出たファイル ID の JSON をドライブから落とし、--gas-file に渡す。
  4. 終わったら `CompareExport.gs` を消す（本番の GAS には入れない）。
  注意: GAS の画面の関数のうち api_getSomuAdminView・api_getDeptAdminView・api_getAdminSomuView・
  api_getAdminBossView・api_getLeaveStatusSearch は、呼ぶたびに V_LEAVE_SUMMARY を書き直す（今の GAS の動き）。
  ステージングの写しのタブだけが書き直される。

      function exportViewsForCompare() {
        var PLAN = [];            // ← plan の JSON（配列）を貼る
        var FROM = 0, TO = PLAN.length;
        var out = { generatedAt: new Date().toISOString(), results: {} };
        PLAN.slice(FROM, TO).forEach(function (c) {
          try {
            var r = globalThis[c.fn].apply(null, c.args);
            out.results[c.key] = { ok: true, result: JSON.parse(JSON.stringify(r === undefined ? null : r)) };
          } catch (e) {
            out.results[c.key] = { ok: false, error: String((e && e.message) || e) };
          }
        });
        var name = 'compare_views_gas_' + FROM + '_' + TO + '_' +
          Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm') + '.json';
        Logger.log(DriveApp.createFile(name, JSON.stringify(out), 'application/json').getId());
      }

比べ方:
  - key（fn と引数の JSON）ごとに、成功か失敗か・返事（失敗なら文）を比べる。辞書のキーの並びは見ない。
    数は 1 と 1.0 を同じとする。配列の並びは見る（api_getAllPendingRequests だけは K-10 で並びを直したので、
    reqId の順に並べ直して比べる）。
  - GAS 版に無い key は「GAS に無い」として数える（plan の範囲を分けて取ったときは、全部のファイルを渡す）。
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable

sys.path.insert(0, str(Path(__file__).resolve().parent))

import acceptance as acc  # noqa: E402

# K-10 で並びを直した（年をまたいでも新しい申請が上）。中身だけを比べる
UNORDERED = {"api_getAllPendingRequests": "reqId"}
MAX_DIFFS_PER_KEY = 5


# ===== plan =====


def call_key(fn: str, args: Iterable[Any]) -> str:
    return fn + json.dumps(list(args), ensure_ascii=False, separators=(",", ":"))


def sample(items: list, n: int) -> list:
    """先頭・真ん中・最後から n 件（0 なら全部）。並びは元のまま。"""
    if n <= 0 or len(items) <= n:
        return list(items)
    idx = sorted({round(i * (len(items) - 1) / (n - 1)) for i in range(n)}) if n > 1 else [0]
    return [items[i] for i in idx]


def months_of_fy(today: date) -> list[str]:
    """今年度（3/16 始まり）の 3 月から今月までの yyyy-MM。"""
    start = acc.fy_start(today)
    y, m = start.year, start.month
    out = []
    while (y, m) <= (today.year, today.month):
        out.append(f"{y:04d}-{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def build_plan(*, depts: list[str], workers: dict[str, list[str]], years: list[int], today: date,
               req_ids: list[str], workers_per_dept: int) -> list[dict]:
    calls: list[tuple[str, tuple]] = [
        ("api_isAdmin", ()), ("api_isSomu", ()), ("api_isDeptApprover", ()), ("api_getDeptList", ()),
        ("api_getWorkerInfo", ()), ("api_getLookupData", ()), ("api_getWorkingDays", ()),
        ("api_getApproverList", ()), ("api_getCurrentApproverInfo", ()), ("api_getApproverProfiles", ()),
        ("api_getAllPendingRequests", ()), ("api_getPending2ndApprovals", ()), ("api_getHolidayDays", ("",)),
        ("api_getUpcomingLeaves", ("", "")),
    ]
    calls += [("api_getPastWorkingDays", (ym,)) for ym in months_of_fy(today)]
    for d in depts:
        calls += [("api_getWorkersByDept", (d,)), ("api_getUpcomingLeaves", (d, ""))]
    for fy in years:
        calls += [("api_getSomuAdminView", (fy, "ALL")), ("api_getAdminSomuView", (fy, "")),
                  ("api_getAdminBossView", ("", fy)), ("api_getLeaveStatusSearch", (fy, "", ""))]
        for d in depts:
            calls += [("api_getSomuAdminView", (fy, d)), ("api_getDeptAdminView", (fy, [d])),
                      ("api_getAdminBossView", (d, fy)), ("api_getLeaveStatusSearch", (fy, d, ""))]
    for d in depts:
        for w in sample(workers.get(d, []), workers_per_dept):
            calls += [("api_getHolidayDays", (w,)), ("api_getUpcomingLeaves", (d, w))]
            calls += [("api_getLeaveStatusSearch", (fy, d, w)) for fy in years]
    calls += [("api_getLeaveRequestById", (rid,)) for rid in req_ids]
    plan, seen = [], set()
    for fn, args in calls:
        key = call_key(fn, args)
        if key not in seen:
            seen.add(key)
            plan.append({"key": key, "fn": fn, "args": list(args)})
    return plan


# ===== 比べる =====


def _norm(fn: str, value: Any) -> Any:
    key = UNORDERED.get(fn)
    if key and isinstance(value, list):
        return sorted(value, key=lambda it: str(it.get(key)) if isinstance(it, dict) else str(it))
    return value


def diff(a: Any, b: Any, path: str = "$", out: list[str] | None = None, limit: int = MAX_DIFFS_PER_KEY) -> list[str]:
    """a（VPS）と b（GAS）の違い（path: VPS の値 / GAS の値）。辞書の並びは見ない。数は値で比べる。"""
    out = [] if out is None else out
    if len(out) >= limit:
        return out
    num = (int, float)
    if isinstance(a, bool) or isinstance(b, bool):
        if a is not b:
            out.append(f"{path}: {acc._short(a, 80)} / {acc._short(b, 80)}")
    elif isinstance(a, num) and isinstance(b, num):
        if a != b:
            out.append(f"{path}: {a} / {b}")
    elif isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b)):
            if k not in a:
                out.append(f"{path}.{k}: （無い） / {acc._short(b[k], 80)}")
            elif k not in b:
                out.append(f"{path}.{k}: {acc._short(a[k], 80)} / （無い）")
            else:
                diff(a[k], b[k], f"{path}.{k}", out, limit)
            if len(out) >= limit:
                break
    elif isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            out.append(f"{path}: 件数 {len(a)} / {len(b)}")
        for i, (x, y) in enumerate(zip(a, b)):
            diff(x, y, f"{path}[{i}]", out, limit)
            if len(out) >= limit:
                break
    elif a != b:
        out.append(f"{path}: {acc._short(a, 80)} / {acc._short(b, 80)}")
    return out


def compare_with_gas(plan: list[dict], vps: dict[str, dict], gas: dict[str, dict]) -> list[tuple[str, list[str]]]:
    """key ごとの違い（違いのある key だけ）。"""
    out = []
    for call in plan:
        key, fn = call["key"], call["fn"]
        v, g = vps.get(key), gas.get(key)
        if g is None:
            out.append((key, ["GAS の返事がありません"]))
            continue
        if v is None:
            out.append((key, ["VPS の返事がありません"]))
            continue
        if bool(v.get("ok")) != bool(g.get("ok")):
            out.append((key, [f"成否: VPS {'成功' if v.get('ok') else '失敗 ' + str(v.get('error'))} / "
                              f"GAS {'成功' if g.get('ok') else '失敗 ' + str(g.get('error'))}"]))
            continue
        if not v.get("ok"):
            if str(v.get("error")) != str(g.get("error")):
                out.append((key, [f"エラー文: {v.get('error')} / {g.get('error')}"]))
            continue
        d = diff(_norm(fn, v.get("result")), _norm(fn, g.get("result")))
        if d:
            out.append((key, d))
    return out


def load_gas_files(paths: Iterable[str]) -> dict[str, dict]:
    merged: dict[str, dict] = {}
    for p in paths:
        data = json.loads(Path(p).read_text("utf-8"))
        merged.update(data.get("results") or {})
    return merged


# ===== DB から独立に数える（コンテナの中だけ） =====


class CountsProbe:
    """leave_requests を直接数える（domain の関数を通さない）。"""

    def __init__(self, session_factory):
        self.session_factory = session_factory

    def _rows(self) -> list[Any]:
        from sqlalchemy import select

        from app.models import LeaveRequest

        with self.session_factory() as s:
            return list(s.scalars(select(LeaveRequest)))

    def req_ids(self) -> list[str]:
        return sorted(r.req_id for r in self._rows() if (r.req_id or "").strip())

    def fiscal_years(self) -> list[int]:
        out = set()
        for r in self._rows():
            try:
                out.add(int(float(str(r.fiscal_year).strip())))
            except ValueError:
                continue
        return sorted(out)

    def counts(self, today: date) -> dict:
        rows = [r for r in self._rows() if (r.req_id or "").strip()]
        by_status: dict[str, int] = {}
        for r in rows:
            by_status[r.status] = by_status.get(r.status, 0) + 1
        pending = sum(1 for r in rows if r.status == acc.ST_SUBMITTED)
        pending2 = sum(1 for r in rows if r.status == acc.ST_APPROVED and r.approved_at2 is None and
                       not (r.approved_at2_text or "").strip())
        upcoming = sum(1 for r in rows if r.status != acc.ST_CANCELED and (r.leave_date is None or r.leave_date >= today))
        per_fy: dict[str, dict] = {}
        for r in rows:
            if r.status != acc.ST_APPROVED:
                continue
            try:
                fy = str(int(float(str(r.fiscal_year).strip())))
            except ValueError:
                continue
            s = per_fy.setdefault(fy, {"approved": 0, "paid": 0.0})
            s["approved"] += 1
            if (r.leave_type or "").strip() == acc.KIND_PAID:
                s["paid"] += 0.5 if (r.leave_kubun or "").strip() == acc.DAY_HALF else 1
        return {"byStatus": by_status, "pending": pending, "pending2": pending2, "upcoming": upcoming, "perFy": per_fy}


def independent_problems(counts: dict, vps: dict[str, dict]) -> list[tuple[str, str, bool | None, str]]:
    """(id, 項目, ok, 実際)。ok=None は比べられない（呼び出しが無い・権限が無い）。"""
    def length(key: str) -> int | None:
        v = vps.get(key) or {}
        return len(v["result"]) if v.get("ok") and isinstance(v.get("result"), list) else None

    out = []
    for item_id, title, key, want in (
        ("C-1", "承認待ち（申請中）の件数", call_key("api_getAllPendingRequests", ()), counts["pending"]),
        ("C-2", "2 次承認待ち（承認済で 2 次承認日時が空）の件数", call_key("api_getPending2ndApprovals", ()),
         counts["pending2"]),
        ("C-3", "予定（取消でなく、休暇日が今日以降か日付でない）の件数", call_key("api_getUpcomingLeaves", ("", "")),
         counts["upcoming"]),
    ):
        got = length(key)
        out.append((item_id, title, None if got is None else got == want, f"返事 {got} / DB {want}"))
    for fy, s in sorted(counts["perFy"].items()):
        v = vps.get(call_key("api_getAdminBossView", ("", int(fy)))) or {}
        if not v.get("ok") or not isinstance(v.get("result"), list):
            out.append((f"C-4-{fy}", f"{fy} 年度の承認済みの件数と有給の回数（古い管理画面の合計）", None,
                        "api_getAdminBossView を読めません（ALLOW_ADMIN_PAGE_EMAILS）"))
            continue
        total = sum(float(it.get("totalCount") or 0) for it in v["result"] if isinstance(it, dict))
        paid = sum(float(it.get("paidCount") or 0) for it in v["result"] if isinstance(it, dict))
        out.append((f"C-4-{fy}", f"{fy} 年度の承認済みの件数と有給の回数（古い管理画面の合計）",
                    total == s["approved"] and paid == s["paid"],
                    f"返事 件数 {total:g}・有給 {paid:g} / DB 件数 {s['approved']}・有給 {s['paid']:g}"))
    return out


# ===== 本体 =====


def run_plan(runner: acc.Runner, plan: list[dict]) -> dict[str, dict]:
    out = {}
    for call in plan:
        r = runner.rpc(call["fn"], *call["args"])
        if r.status == 200 and r.ok:
            out[call["key"]] = {"ok": True, "result": r.result}
        else:
            out[call["key"]] = {"ok": False, "status": r.status, "error": r.error}
    return out


def gather_plan(runner: acc.Runner, *, years: list[int], req_ids: list[str], workers_per_dept: int) -> list[dict]:
    depts_r = runner.rpc("api_getDeptList")
    depts = [d["deptName"] for d in (depts_r.result or []) if isinstance(d, dict)] if depts_r.ok else []
    workers: dict[str, list[str]] = {}
    for d in depts:
        w = runner.rpc("api_getWorkersByDept", d)
        workers[d] = [x["workerId"] for x in (w.result or []) if isinstance(x, dict)] if w.ok else []
    return build_plan(depts=depts, workers=workers, years=years, today=runner.today, req_ids=req_ids,
                      workers_per_dept=workers_per_dept)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="読むだけの関数の総当たり比較（ステージングだけ）")
    p.add_argument("--base-url", default="http://localhost:8000")
    p.add_argument("--origin", default=acc.DEFAULT_ORIGIN)
    p.add_argument("--cookie", default=None, help="セッションのクッキーの値（外から動かすとき）")
    p.add_argument("--mint-session", action="store_true", help="staging でセッションを作り、最後に消す（コンテナの中）")
    p.add_argument("--email", default=None)
    p.add_argument("--years", default="", help="年度（カンマ区切り）。DB を見られないときに使う。既定は去年度と今年度")
    p.add_argument("--workers-per-dept", type=int, default=3, help="部署ごとに抜き取る作業員の数（0 で全員）")
    p.add_argument("--out", default=None, help="VPS 版の返事の保存先（JSON）")
    p.add_argument("--plan-out", default=None, help="GAS 版で同じ呼び出しをする plan の保存先（JSON）")
    p.add_argument("--gas-file", action="append", default=[], help="GAS 版の返事のファイル（何回でも）")
    p.add_argument("--timeout", type=float, default=120.0)
    return p


def main(argv: list[str] | None = None, *, api: acc.Transport | None = None, counts_probe: CountsProbe | None = None,
         minter: Callable[[bool], tuple[str, Callable[[], None]]] | None = None, now: datetime | None = None) -> int:
    args = build_parser().parse_args(argv)
    api = api or acc.UrllibTransport(args.base_url, args.timeout)
    runner = acc.Runner(api, origin=args.origin, cookie=args.cookie, now=now)
    env = runner.check_health()
    if env != "staging":
        print(f"REFUSED: /healthz の env が staging ではありません（{env}）。総当たりはステージングでだけ動かします。",
              file=sys.stderr)
        return 2
    cleanups: list[Callable[[], None]] = []
    try:
        if args.mint_session:
            if minter is None:
                try:
                    minter, probe = acc._container_minter(args.email)
                except RuntimeError as exc:
                    print("REFUSED: " + str(exc), file=sys.stderr)
                    return 2
                counts_probe = counts_probe or CountsProbe(probe.session_factory)
            try:
                cookie, undo = minter(False)
            except RuntimeError as exc:
                print("REFUSED: " + str(exc), file=sys.stderr)
                return 2
            cleanups.append(undo)
            runner.cookie = acc.cookie_header(cookie)
        if not runner.cookie:
            print("REFUSED: セッションがありません（--mint-session か --cookie）", file=sys.stderr)
            return 2

        if args.years:
            years = [int(y) for y in args.years.split(",") if y.strip()]
        elif counts_probe is not None:
            years = counts_probe.fiscal_years() or [runner.fy]
        else:
            years = [runner.fy - 1, runner.fy]
        req_ids = counts_probe.req_ids() if counts_probe is not None else []
        plan = gather_plan(runner, years=years, req_ids=req_ids, workers_per_dept=args.workers_per_dept)
        if args.plan_out:
            Path(args.plan_out).write_text(json.dumps(plan, ensure_ascii=False, indent=0), "utf-8")
        vps = run_plan(runner, plan)
        rep = runner.report
        failed_calls = [k for k, v in vps.items() if not v.get("ok") and v.get("status") != 400]
        rep.add("V-1", f"読む関数 {len(plan)} 回の呼び出しが 200 か業務エラー（400）で返る", not failed_calls,
                f"失敗: {failed_calls[:5]}" if failed_calls else f"{len(plan)} 回", "500・501・401 なし")
        if counts_probe is not None:
            counts = counts_probe.counts(runner.today)
            rep.notes.append("DB の状態別の件数: " + json.dumps(counts["byStatus"], ensure_ascii=False))
            for item_id, title, ok, actual in independent_problems(counts, vps):
                rep.add(item_id, title, ok, actual, "DB から独立に数えた値と同じ")
        else:
            rep.add("C", "DB から独立に数えた値との比べ", None, "DB を見られません（--mint-session で動かす）")
        if args.gas_file:
            gas = load_gas_files(args.gas_file)
            diffs = compare_with_gas(plan, vps, gas)
            rep.add("G-1", f"GAS 版の返事と同じ（{len(plan)} 回）", not diffs, f"違い {len(diffs)} 件", "違い 0 件")
            for key, d in diffs[:50]:
                rep.notes.append(f"違い {key}: " + " | ".join(d))
        else:
            rep.add("G-1", "GAS 版の返事と同じ", None, "GAS 版のファイルがありません（--gas-file）")
        if args.out:
            Path(args.out).write_text(json.dumps({
                "meta": {"generatedAt": datetime.now(acc.JST).isoformat(), "today": acc.ymd(runner.today),
                         "years": years, "calls": len(plan)},
                "plan": plan, "results": vps}, ensure_ascii=False, indent=0, default=str), "utf-8")
        print(rep.table())
        for note in rep.notes:
            print(note)
        return rep.exit_code()
    finally:
        for undo in reversed(cleanups):
            try:
                undo()
            except Exception as exc:  # noqa: BLE001
                print(f"session cleanup failed: {type(exc).__name__}", file=sys.stderr)


if __name__ == "__main__":
    sys.exit(main())
