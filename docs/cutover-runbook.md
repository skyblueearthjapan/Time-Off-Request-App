# 本番切り替え手順書（休暇届 → leave.lineworks-local.net）

作成: 2026-11-26（担当 C）。もと: `docs/vps-migration/06-day1-findings-and-decisions.md`（K-0〜K-11）、計画書 `docs/vps-migration-plan.md` 6 章・8 章、実装仕様 `docs/vps-implementation-spec.md` 7・9・12・14 章、`docs/vps-migration/05-risks-and-checks.md` 4 節、残業の `docs/cutover-runbook.md` と当日の記録 `docs/handoff-2026-11-26.md`、共通スキル `gas-to-vps-migration` の 05〜07。

読み手は次のディレクター（AI）と今泉オーナーです。**この順番で行い、各手順の「確かめ」が通らなければ先へ進みません。** 戻すと決めるのはオーナーです（C 章）。

〇〇 はまだ決まっていない値で、当日に控えます。【要確認】は値はあるが実物で確かめていないものです。どちらも最後の「確かめること」に一覧があります。

## 値

| 項目 | 値 |
|---|---|
| VPS | `srv1508169`（`ssh lineworks-vps-user`）。毎回 `hostname` を確かめる。`srv1300618` には触れない |
| VPS の置き場 | `/opt/kyuka`（`vps-app/` の中身を直置き。git のフォルダではない。`.env` も直下、権限 600） |
| コンテナ | `kyuka-db`・`kyuka-api`・`kyuka-worker`・`kyuka-web`・`kyuka-cloudflared`・`kyuka-backup`（compose のサービス名は `db`・`api`・`worker`・`web`・`cloudflared`・`backup`） |
| DB | PostgreSQL 16、DB 名 `kyuka`、ユーザー `kyuka` |
| 新しい URL | `https://leave.lineworks-local.net`（ステージングと同じホスト名。切り替えで `APP_ENV` だけ変える） |
| GAS scriptId（本番） | `15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i` |
| 本番の画面のデプロイ（2 つ） | **@106** `AKfycbx7sB9x7iLigEqZX9LuErK95R2ZC0cMOyU9efr6pIbpnqAnnDC3o5_DGFbtwgSh64j2`（`APP_URL` とポータルが指す）、**@104** `AKfycbwHI1uz6tcer59sRC4dYEACDaevSn9a_eJa3J0E8LZgjdWCGPJVaqr_Qa9L-HXVKI8k`（ブックマークだけ）。どちらも `access: DOMAIN`・`executeAs: USER_DEPLOYING`。**戻す版は 106 と 104** |
| 移転モード入りの版 | @〇〇（B-3 で作る） |
| 本番の内部用のデプロイ | `AKfyc〇〇`（B-4 で作る。`access: DOMAIN` のまま）。`GAS_INTERNAL_URL`＝`https://script.google.com/macros/s/<この ID>/exec`、`PROD_GAS_INTERNAL_DEPLOYMENT_ID`＝この ID |
| GAS の呼び方 | `GAS_INTERNAL_AUTH=bearer`（スパイク S-1 で決定。委任のトークンを `Authorization` に付ける） |
| GAS のコード | ブランチ `gas-cutover` の `src/`（`src/Cutover.gs`・`src/Internal.gs` と 12 ファイルの小さな変更）。ディレクターが `master` に取り込んでいれば `master` |
| 本番 DB スプレッドシート | `16v1tQNCFkDTk2mnAZ4E2eEdHeUQ2Bepdcsn5vBAmllw`「休暇届けapp」（時間帯 `Etc/GMT`、地域 `en_US`） |
| 台帳 | `1B7tnsT4lr80o3pFMBxB2zEhsx2UW2PCL48YTQc9i1eU`（VPS は書かない。GAS が書く） |
| 本番 PDF フォルダ | `1MQibR08wWsE_OGurO9RocEEDGgDOwo-w`（VPS は書かない） |
| 読むだけの元 | マスタ元 `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ`、残業 DB `1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ`（**絶対に書かない**） |
| ステージング（安全装置の確かめに使う） | DB `1SG4FV9raYuFu8jaWzCVpXvKqo1LPEO7aR44KifWrYqw`、台帳 `1FpeButODbprm7Bri2WQWlDgwzYDwiwLb8FvWu1XO2YM`、PDF フォルダ `1T6ZZ8t-G7fq8BxMUoB7yw2I8i2VXqRCm`、GAS `1T3bWvTY3bv8R3UeWolurVSRtq1-20NESypGegxh-iP7PPo5xcJqzXyA3`（デプロイ `AKfycbz6KZ0t…`） |
| 空にする表 | 移行の道具の `--truncate` が空にする `leave_requests, settings, lookup, workers, calendar, stamps, dept_approvers, sheet_sync_state, jobs, idempotency_keys, sync_runs, sessions` と、手で `audit_log` |
| 残す表 | `google_credentials`（委任）、`alembic_version` |
| 統合業務メール | 別のスクリプト（`1lb7bmF9…`）。7:00・17:00・18:00 にこのシートを読む。**トリガーもコードも触らない** |
| 社内ポータル | 「社内ダッシュボード」`1DdhbJuz-zyS_2pKcM8y6Xf2pHofuzwNDgoW9eWcB0yg` の `APP_MASTER`、key `leave_request`（値の後ろに空白。trim して探す）、url 列（今は @106 の `/exec`） |
| 手元の秘密の置き場 | `C:\Users\imaizumi.LINEWORKS-NET\.secrets\kyuka\`。本番の内部用トークンは当日 `gas-internal-token.prod` に置く |
| 切り替えの控えの置き場（VPS） | `/opt/kyuka/cutover/`（`backups/` は自動で消えるので使わない） |
| 戻せる期間 | 切り替えから 5 営業日 |

### コマンドの書き方

VPS のコマンドは、手元（Git Bash）から次の形で流します。毎回、先頭で `hostname` を確かめ、違えば何もしません。docker は `sg docker -c "…"` で動かします（lineworks ユーザーは sudo できません）。

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST: $(hostname)"; exit 9; }
cd /opt/kyuka
# ここに手順のコマンド
EOF
```

- この here-doc の中で `docker compose exec -T`・`docker compose run -T` を使うときは、**入力を渡さない行に必ず `< /dev/null` を付けます**。付けないと here-doc の残りの行を読み込んでしまい、後の手順が動きません（罠 77）。入力を `<<'PY'`・`<<'SQL'` で渡す行は付けません。
- DB へは次の関数で SQL を流します（上の形の中で定義してから使います）。

  ```sh
  psql_q() { sg docker -c "docker compose exec -T db psql -U kyuka -d kyuka -v ON_ERROR_STOP=1 -X -P footer=off"; }
  ```

- 道具（`tools/`）はイメージに入っていないので、`-e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro` を付けて動かします（罠 5）。
- 時間のかかるもの（移行・控え）は `timeout 590 ssh …` で包みます（罠 2）。
- GAS の関数は、GAS エディタ `https://script.google.com/u/1/home/projects/15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i/edit` を開き、右上のアカウントが `imaizumi@lineworks-local.info` であることを確かめてから、関数を選んで「実行」します（運用関数は持ち主がエディタから実行したときだけ動きます。罠 48）。
- clasp は手元で `CLASP="npx -y @google/clasp@3.4.1"` として使います。最初に `$CLASP show-authorized-user` で `imaizumi@lineworks-local.info` を確かめます（切れていればオーナーに `! npx clasp login` を頼む。罠 28）。

### 統合業務メールと同じ条件でシートを数える（B-1・B-17・D 章で使う）

本番シートの `T_LEAVE_REQUEST` と `V_LEAVE_SUMMARY` を、統合業務メール（`readerLeave.gs`）と同じ条件で読んで数えます。読むだけです。api コンテナの中から委任トークンで読むので、`.env` がステージングのままでも本番シートを読めます（読むだけは安全装置の対象外）。

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF' | tee -a ~/kyuka-unified-counts.log
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app api python -" <<'PY'
import json
from datetime import datetime, timedelta, timezone
from app import config as C
from app.db import get_sessionmaker
from app.services.sheets_api import SheetsApi
JST = timezone(timedelta(hours=9))
def jst(serial):  # シートの時間帯（Etc/GMT）のシリアル値 → JST の日付
    return (datetime(1899, 12, 30, tzinfo=timezone.utc) + timedelta(days=serial)).astimezone(JST).date()
def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)
with get_sessionmaker()() as db:
    t, v = SheetsApi.for_db(db).batch_get_values(C.PROD_SPREADSHEET_ID_FIXED, ["'T_LEAVE_REQUEST'!A1:Z", "'V_LEAVE_SUMMARY'!A1:L"])
ix = {str(h).strip(): i for i, h in enumerate(t[0])}
def c(r, h):
    i = ix[h]
    return r[i] if i < len(r) else ""
today = datetime.now(JST).date()
fy = today.year if (today.month, today.day) >= (3, 16) else today.year - 1
rows = [r for r in t[1:] if str(c(r, "REQ_ID")).strip()]
st = {}
for r in rows:
    st[c(r, "STATUS")] = st.get(c(r, "STATUS"), 0) + 1
applied_today = [r for r in rows if num(c(r, "APPLIED_AT")) and jst(c(r, "APPLIED_AT")) == today]
out = {
    "at": datetime.now(JST).strftime("%Y-%m-%d %H:%M"), "rows": len(rows), "byStatus": st,
    "leaveDateNotDate": sum(1 for r in rows if not num(c(r, "LEAVE_DATE"))),
    "todayLeaves(7:00)": sum(1 for r in rows if c(r, "STATUS") == "承認済" and num(c(r, "LEAVE_DATE")) and jst(c(r, "LEAVE_DATE")) == today),
    "pending1(17:00)": sum(1 for r in rows if c(r, "STATUS") == "申請中" and num(c(r, "LEAVE_DATE")) and jst(c(r, "LEAVE_DATE")) >= today),
    "pending2(18:00)": sum(1 for r in rows if c(r, "STATUS") == "承認済" and str(c(r, "APPROVED_AT2")).strip() == ""),
    "appliedToday": len(applied_today),
    "appliedTodayPast": sum(1 for r in applied_today if num(c(r, "LEAVE_DATE")) and jst(c(r, "LEAVE_DATE")) < today),
}
vh = {str(h).strip(): i for i, h in enumerate(v[0])}
vrows = [r for r in v[1:] if r]
def vc(r, h):
    i = vh[h]
    return r[i] if i < len(r) else ""
out["summaryRows"] = len(vrows)
out["warnWorkers(今年度)"] = sum(1 for r in vrows if num(vc(r, "年度")) and int(vc(r, "年度")) == fy and str(vc(r, "警告レベル")).strip())
print(json.dumps(out, ensure_ascii=False))
PY
EOF
```

出る値: 全件数・状態別・休暇日が日付でない行の数（0 であること）・本日の休暇者（7:00）・1 次未承認（17:00）・2 次承認待ち（18:00）・今日の申請と、そのうちの過去申請（統合業務メールは「今日提出の過去申請」も載せる）・集計タブの行数・今年度の有給未達者（警告レベルあり）。結果は手元の `~/kyuka-unified-counts.log` に足していきます。

---

## A. 前もって済ませること（前日まで）

- [ ] **A-1 受け入れ**: `docs/acceptance-checklist.md` の行がすべて「済」（自動・JEV ハンズの実機・オーナーの実機を分けて書く）。ステージングで `tools/qa/acceptance.py` を読むだけと `--write` の両方で流し、FAIL が 0。手元の自動テストが全件通る（`cd vps-app/api && .venv/Scripts/python -m pytest -q` の最後の行。罠 44）。レビューの指摘が片付いている。

  ```sh
  ssh lineworks-vps-user 'bash -s' <<'EOF'
  set -eu
  test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
  cd /opt/kyuka
  sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session" < /dev/null
  sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session --write" < /dev/null
  EOF
  ```

  確かめ: 両方とも `PASS n / FAIL 0`、最後に `session removed (main): True` と `session removed (outsider): True`。`--write` は承認完了メールを 2 通（オーナーの 1 件だけに）出し、作った申請 4 件は【受入テスト】の印つきで、最後に取消されています（報告の「試験の申請（すべて取消済み）」）。

- [ ] **A-2 総当たり比較**: `tools/qa/compare_views.py` を `--mint-session --plan-out` で流し、写しの GAS で同じ plan の返事を取り（道具の docstring の手順）、`--gas-file` で突き合わせて `G-1` が PASS、`C-*`（DB から独立に数えた値）が PASS。
- [ ] **A-3 リハーサル**: ステージングのまま、B-9〜B-16 の移行・照合・api だけ起動・worker 起動を 2 回以上なぞり（読むだけのリハーサルを 1 回含む）、`verifyProblems: []`、照合 SQL（B-13）の数を控え、移行直後の `sheet_reconcile --check` で差 0。戻しの練習（C 章）をステージングの写しの GAS で 1 回。
- [ ] **A-4 GAS のコードの下見**（読むだけ。本番には入れない＝罠 61）: `gas-cutover` のコードが本番の今のコードと「この移行の変更」以外で同じことを確かめます。

  ```sh
  cd /d/projects/kyuka-todoke-app
  CLASP="npx -y @google/clasp@3.4.1"
  rm -rf /tmp/kyuka-pull && mkdir /tmp/kyuka-pull && cd /tmp/kyuka-pull
  echo '{"scriptId":"15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i"}' > .clasp.json && $CLASP pull
  cd /d/projects/kyuka-todoke-app/repo
  git diff --stat master gas-cutover -- src     # Cutover.gs・Internal.gs の追加と 12 ファイルの小さな変更だけ
  for f in /tmp/kyuka-pull/*; do b=$(basename "$f"); cmp -s <(tr -d '\r' < "$f") <(git show master:src/"$b" | tr -d '\r') || echo "DIFF $b"; done
  ```

  確かめ: `DIFF` が 1 行も出ない（本番の HEAD ＝ `master` の `src/`。CRLF の違いは見ない）。`appsscript.json` が `master` と `gas-cutover` で同じ（`git diff master gas-cutover -- src/appsscript.json` が空。`webapp` は `executeAs: USER_DEPLOYING`・`access: DOMAIN`。罠 45）。
- [ ] **A-5 トリガーの実物**: GAS の「トリガー」画面で「すべての人のトリガー」を表示し、関数名・持ち主・時刻を控える（スパイク S-7。`syncAllMasters` のほか、昔の `checkAndSendWarnMails_`・`notifyTodayLeaves_`・`notifyPendingApproval1_`・`notifyPendingApproval2_` が残っていないか）。imaizumi@ 以外の持ち物があれば、消してもらう段取りを決める。
- [ ] **A-6 控えておくもの**: DB の時間帯 `Etc/GMT`・地域 `en_US`（確認済み）、台帳のタブと見出し 20 列（S-9）、M_SYSTEM_SETTING に同じキーが無いこと（確認済み）、M_SYSTEM_SETTING の `APP_URL` の今の値（B-19 で書き戻すときに使う）、ポータルの url 列の今の値。
- [ ] **A-7 部署の共有アカウント**: 部署の共有アカウントが新しい URL（ステージング）でログインできること。翌朝に初めてのログインが集中しないよう、前日の業務中に各部署の端末で 1 回開いてもらうかを、オーナーと決める（ログインの持ち時間は 90 日＝K-6）。
- [ ] **A-8 `.env` の下書き**: 本番で変えるキーは `APP_ENV`・`GAS_INTERNAL_URL`・`GAS_INTERNAL_TOKEN`・`PROD_GAS_INTERNAL_DEPLOYMENT_ID`・`APP_IMAGE_TAG`（最新のイメージのタグ）・`SKIP_DAILY_ON_STARTUP` だけ。ほかは同じ（`DB_SHEET_TIMEZONE=Etc/GMT`、`BOOTSTRAP_ADMINS`、`PROD_*` の ID は今のまま）。今のタグを控える: `ssh lineworks-vps-user 'grep -E "^APP_IMAGE_TAG=" /opt/kyuka/.env; sg docker -c "docker image ls kyuka-app"'`。
- [ ] **A-9 オーナーへの声かけ**（下の「オーナーへの連絡」の 1）。当日の日付を決め、本番 GAS の変更（B-3〜B-7）・ポータルの書き換え（B-20）・`APP_URL` の書き換え（B-19）の許可をもらう。

---

## B. 当日（夜。18:00 の統合業務メールの後に始め、翌朝 6 時台のマスタ同期の前に終える＝K-8）

**始める前にオーナーの合図をもらいます**（「切り替えを始めます」。下の「オーナーへの連絡」の 2）。18:00 の統合業務メール（2 次承認待ち）が届いたことを確かめてから始めます。

| # | 手順 | 確かめ | 戻す範囲 |
|---|---|---|---|
| B-0 | ホスト名とオーナーの合図 | `srv1508169`、合図あり、18:00 のメールが届いた | — |
| B-1 | 進行中の件数を数える（本番シート。統合業務メールの条件） | 申請中・2 次承認待ち・承認済で PDF なしを控える（2026-11-26 午前は申請中 1・2 次待ち 0） | — |
| B-2 | VPS の api・worker を止める（罠 75） | `kyuka-api`・`kyuka-worker` が Exited | `docker compose start api worker` |
| B-3 | GAS: `.clasp.json` とマニフェストを確かめて push → 版 → **@106 と @104 の両方**に当てる | 両方の versionNumber が新しい版（`list-deployments --json`）。`/exec` は今までどおり開く | 両方を 106・104 に戻す |
| B-4 | GAS: 本番の内部用のデプロイを作る | deploymentId を控える。画面のデプロイ 2 つは変わらない | 作ったデプロイを消す（任意） |
| B-5 | GAS: `setInternalToken` | ログに「64 字」、値を手元の秘密の置き場へ | プロパティ `VPS_INTERNAL_TOKEN` を消す |
| B-6 | GAS: `enterMigrationModeForKyuka` | **@106 と @104 の両方**の `/exec` が移転ページ。`?page=preview&reqId=` のボタンが最上位の窓で開く | `exitMigrationMode` |
| B-7 | GAS: `removeAllTriggersForCutover` | 消したトリガーが A-5 の控えと同じ。「すべての人のトリガー」が空。統合業務メールは触らない | `restoreTriggersAfterRollback` |
| B-8 | 流しきり: 1〜2 分待ち、`showMigrationStatus` と「実行数」 | 動いている実行なし、トリガーなし、移転モードあり | — |
| B-9 | VPS の DB を `pg_dump -Fc`（ステージングのデータ） | ダンプが 1KB 超 | — |
| B-10 | `.env` を控えて本番の値に | 値と権限 600、イメージのタグが最新 | 控えを戻す |
| B-11 | 控え: 本番 DB と台帳を Drive API で複製（`cutover_backup.py`） | 2 冊とも `ok: true`、控えの ID を控える | — |
| B-12 | 表を空にして移行（`--dry-run` → `--truncate`） | 終了コード 0、`verifyProblems: []`、`droppedColumns: {}` | ダンプから `pg_restore` |
| B-13 | 照合 SQL と `sheet_reconcile --check` | 件数が B-1 とリハーサルと同じ、差 0 | B-12 からやり直すか C へ |
| B-14 | **api だけ**起動し、安全装置と `gas_ping` を確かめる | `NG` 0 行、healthz が production、ping が本番の DB と移転モード | ログを見て直すか C へ |
| B-15 | マスタ同期と設定の取り込みを手で 1 回 | 作業員 50・カレンダー 130 前後、控えとの差 0 | 控えから 4 タブを写し戻す |
| B-16 | worker を起動（`SKIP_DAILY_ON_STARTUP=1`）。ping・行の反映・集計タブ・照合 | 反映待ち 0、jobs 0、`gas_ping` 成功、集計タブと T_LEAVE_REQUEST の控えとの差 0 | worker を止め、控えから写し戻して C へ |
| B-17 | 統合業務メールの条件で本番シートを数える | B-1 と同じ数 | C へ |
| B-18 | 画面の確かめ（書き込みはしない）。古いリンクが移転ページから新しい画面の同じ申請を開く | 期待どおり | — |
| B-19 | M_SYSTEM_SETTING の `APP_URL` を、元の値を控えてから `https://leave.lineworks-local.net/` に（K-11） | 値が変わり、5 分以内に VPS の取り込みにも出る | 控えた元の値に書き戻す |
| B-20 | 社内ポータル APP_MASTER（key `leave_request`）の url を差し替える（前に声をかける） | ポータルから新しい URL が開く | 控えた元の値に書き戻す |
| B-21 | 最終確認 | 全コンテナ up、jobs の pending・failed 0、外から healthz 200 | — |
| B-22 | 記録と報告 | — | — |

### B-0 ホスト名と合図

```sh
ssh lineworks-vps-user hostname    # srv1508169 でなければ中止
```

オーナーの合図をもらい、18:00 の統合業務メール（2 次承認待ち）がオーナーに届いたことを確かめます。ここから翌朝 5:30 までに B-22 まで終える見込みで進めます（目安 60〜90 分）。

### B-1 進行中の件数

上の「統合業務メールと同じ条件でシートを数える」を流し、結果を控えます。あわせて、本番シートで「承認済で `PDF_URL` が空」の件数を見ます（移行の `--dry-run` の報告でも数えられます）。

確かめ: `leaveDateNotDate` が 0。申請中・2 次承認待ちの件数を控える（2026-11-26 午前は申請中 1 件 `LV-20261126-001`・2 次承認待ち 0 件）。**送り直すメールはありません**（承認完了メールは 1 次承認のときに GAS から出ている。申請中の案件は VPS で 1 次承認されたときに、今と同じく GAS から出る）。

### B-2 VPS を止める

移転ページのボタンの先（新しい URL）は、この時点ではまだステージングです。利用者がステージングのデータに書き込まないよう、移転モードより先に止めます（罠 75）。

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
sg docker -c "docker compose stop worker api"
sg docker -c "docker compose ps -a"
EOF
```

確かめ: `kyuka-api`・`kyuka-worker` が `Exited`（`kyuka-web` は動いたまま。新しい URL は 502 になる）。戻す: `sg docker -c "docker compose start api worker"`。

### B-3 GAS: push・版・両方のデプロイ

GAS のコードを本番に入れるのは今日だけです（本番のトリガーは HEAD のコードで動くため。罠 61）。`clasp create --parentId` は使いません（指定を無視して新しいスプレッドシートを作る。06 の 7 節）。

1. 押す前の確かめ（手元）。

   ```sh
   cd /d/projects/kyuka-todoke-app/repo
   git worktree add /d/projects/kyuka-todoke-app/gas-push gas-cutover 2>/dev/null || true   # master に取り込み済みなら master
   cd /d/projects/kyuka-todoke-app/gas-push
   git log --oneline -1                      # 4842727 以降の gas-cutover（取り込み済みなら master の HEAD）
   cat .clasp.json                           # scriptId が 15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i、rootDir が src
   CLASP="npx -y @google/clasp@3.4.1"
   $CLASP show-authorized-user               # imaizumi@lineworks-local.info
   $CLASP list-deployments --json > /tmp/kyuka-deployments-before.json
   rm -rf /tmp/kyuka-pull && mkdir /tmp/kyuka-pull && (cd /tmp/kyuka-pull && echo '{"scriptId":"15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i"}' > .clasp.json && $CLASP pull)
   diff <(tr -d '\r' < /tmp/kyuka-pull/appsscript.json) <(tr -d '\r' < src/appsscript.json) && echo "MANIFEST SAME"
   ```

   確かめ: `.clasp.json` の scriptId が本番、`MANIFEST SAME`（スコープ・`webapp` が今と同じ。違えば押さずに止める）、`kyuka-deployments-before.json` に @106・@104 の 2 つ（versionNumber 106・104）があること。A-4 の `DIFF` が 0 だったこと。
2. push と版。説明文は英数字にします（罠 27）。

   ```sh
   $CLASP push --force
   $CLASP create-version "kyuka migration mode and vps internal endpoint"   # 出てきた版の番号を 〇〇 として控える
   $CLASP update-deployment AKfycbx7sB9x7iLigEqZX9LuErK95R2ZC0cMOyU9efr6pIbpnqAnnDC3o5_DGFbtwgSh64j2 --versionNumber 〇〇 --description "migration mode"
   $CLASP update-deployment AKfycbwHI1uz6tcer59sRC4dYEACDaevSn9a_eJa3J0E8LZgjdWCGPJVaqr_Qa9L-HXVKI8k --versionNumber 〇〇 --description "migration mode"
   $CLASP list-deployments --json | grep -A6 -E 'AKfycbx7sB9x|AKfycbwHI1uz'   # 両方の versionNumber が 〇〇
   ```

   **clasp の `update-deployment` は「Redeployed … @N」と出ても版が変わらないことがあります**（罠 60。clasp 3.4.1 の不具合）。`versionNumber` が変わっていないデプロイには、Apps Script API で当てます（clasp の認証ファイルのトークンを使い、トークンは画面に出しません）。

   ```sh
   for DEP in AKfycbx7sB9x7iLigEqZX9LuErK95R2ZC0cMOyU9efr6pIbpnqAnnDC3o5_DGFbtwgSh64j2 AKfycbwHI1uz6tcer59sRC4dYEACDaevSn9a_eJa3J0E8LZgjdWCGPJVaqr_Qa9L-HXVKI8k; do
   DEPLOYMENT_ID=$DEP VERSION=〇〇 python - <<'PY'
   import json, os, urllib.parse, urllib.request
   SCRIPT_ID = "15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i"
   dep, ver = os.environ["DEPLOYMENT_ID"], int(os.environ["VERSION"])
   t = json.load(open(os.path.expanduser("~/.clasprc.json")))["tokens"]["default"]
   form = urllib.parse.urlencode({"client_id": t["client_id"], "client_secret": t["client_secret"],
                                  "refresh_token": t["refresh_token"], "grant_type": "refresh_token"}).encode()
   token = json.load(urllib.request.urlopen("https://oauth2.googleapis.com/token", form))["access_token"]
   body = {"deploymentConfig": {"scriptId": SCRIPT_ID, "versionNumber": ver, "manifestFileName": "appsscript",
                                "description": "migration mode"}}
   req = urllib.request.Request(f"https://script.googleapis.com/v1/projects/{SCRIPT_ID}/deployments/{dep}",
                                data=json.dumps(body).encode(), method="PUT",
                                headers={"Authorization": "Bearer " + token, "Content-Type": "application/json"})
   print(dep[:14], "versionNumber =", json.load(urllib.request.urlopen(req))["deploymentConfig"]["versionNumber"])
   PY
   done
   ```

3. 確かめ: `list-deployments --json` で @106 と @104 の**両方**の `versionNumber` が 〇〇。`https://script.google.com/macros/s/AKfycbx7sB9x…/exec` と `…/AKfycbwHI1uz…/exec` が**今までどおりの画面**で開く（`MIGRATED_TO_URL` がまだ無いので今と同じ動き。オーナーのアカウントのブラウザで見る。書き込みはしない）。GAS エディタで `showMigrationStatus` を 1 回実行し、「移転モード: なし」「内部用の token: なし」「トリガー:（A-5 の控えと同じ）」を見る（権限の許可が出たら許可する）。

戻す: 両方のデプロイを版 106・104 に戻す（C-5 のコマンド）。HEAD は移転モード入りのコードのままでも、`MIGRATED_TO_URL` が無ければ今と同じ動きです。元の HEAD に戻すときは `master` の `src/` を `clasp push --force`。

### B-4 GAS: 本番の内部用のデプロイ

```sh
$CLASP create-deployment --versionNumber 〇〇 --description "vps internal"
$CLASP list-deployments --json > /tmp/kyuka-deployments-after.json
```

出てきた deploymentId（`AKfyc…`）を控えます（〇〇）。画面のデプロイ（@106・@104）とは別の ID であること、`access: DOMAIN` のまま（マニフェストを変えていない）であることを確かめます。`GAS_INTERNAL_AUTH=bearer` なので、VPS は委任のトークンを付けて呼びます（スパイク S-1）。戻す: 何もしなくてよい（token が無い・違う呼び出しには何もしない）。消すときは `$CLASP undeploy <ID>`。

### B-5 GAS: 本番の内部用トークン

GAS エディタで `setInternalToken` を（引数なしで）実行します。64 字の乱数をスクリプトのプロパティ `VPS_INTERNAL_TOKEN` に保存し、ログには長さだけが出ます。

確かめ: ログ「VPS_INTERNAL_TOKEN を保存しました（64 字）」。「プロジェクトの設定 → スクリプト プロパティ」で値をコピーし、手元の `C:\Users\imaizumi.LINEWORKS-NET\.secrets\kyuka\gas-internal-token.prod` に保存します（改行なし。画面・チャット・ログに貼らない）。**`MIGRATED_TO_URL` は手で入れません**（B-6 の関数が入れます）。戻す: プロパティ `VPS_INTERNAL_TOKEN` を消す。

### B-6 GAS: 移転モード

GAS エディタで `enterMigrationModeForKyuka` を実行します（`MIGRATED_TO_URL` に `https://leave.lineworks-local.net` を入れます）。

確かめ:
- **@106 と @104 の両方**の `/exec` が「休暇届は新しいアドレスへ移りました」のページになる（どちらも B-3 で新しい版を当てたため）。
- `…/exec?page=preview&reqId=<B-1 の申請中の REQ_ID>` のボタンで、`https://leave.lineworks-local.net/?page=preview&reqId=<同じ ID>` が**最上位の窓で**開く（GAS の枠の中で開かない。自動の移動は「動けばよい」＝罠 29）。この時点では VPS が止まっているので、開いた先は 502 でかまいません（行き先の URL だけを見る）。
- `showMigrationStatus` で「移転モード: https://leave.lineworks-local.net」「内部用の token: あり」。

戻す: `exitMigrationMode`。

### B-7 GAS: トリガーを全部控えて消す

GAS エディタで `removeAllTriggersForCutover` を実行します（`removeAllTriggersForMigration` と同じ処理。消す前に関数名・種類・ID をプロパティ `MIGRATION_REMOVED_TRIGGERS` に控えます）。

確かめ: ログ「削除したトリガー: syncAllMasters (…)」と、A-5 で控えたほかのトリガー（昔の通知・警告のトリガーが残っていれば、それも）。「残っているトリガー: なし」。トリガーの画面で「すべての人のトリガー」も空。**統合業務メール（別のスクリプト）のトリガーは触りません。** 移転モードでも `syncAllMasters`・`checkAndSendWarnMails_`・`notify*_` の関数は先頭で何もせずに戻るので、消し忘れても書き込みは起きません。

戻す: `restoreTriggersAfterRollback`（`syncAllMasters` の毎日 6 時台の 1 本だけ。昔のトリガーも戻すとオーナーが決めたときだけ `restoreAllTriggersAfterRollback`）。

### B-8 流しきり

1〜2 分待ちます（入口を通り過ぎた直後の書き込みと、2 次承認の PDF・台帳が書き終わるのを待つ。PDF は 1 件 15 秒ほど＝罠 30）。GAS の「実行数」に実行中の行が無いこと、`showMigrationStatus` で「トリガー: なし」・移転モードありを確かめます。**控え（B-11）はこの後に取ります**（GAS が最後に書いた後の状態を控えるため）。

### B-9 VPS の DB を控える

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
mkdir -p cutover
F=cutover/kyuka_before_cutover_$(date +%Y%m%d_%H%M).dump
sg docker -c "docker compose exec -T db pg_dump -U kyuka -Fc kyuka" < /dev/null > "$F"
ls -l "$F"
test "$(stat -c %s "$F")" -gt 1024 && echo "DUMP OK: $F" || echo "DUMP TOO SMALL: $F"
EOF
```

確かめ: `DUMP OK`（ステージングのデータの控え）。順番は「ダンプ → `.env` を本番に → 控え」です。控えと移行の道具は `.env` の環境を対象にするためです（罠 76）。

### B-10 `.env` を本番の値に

先に手元から本番の内部用トークンを送ってから書き換えます（トークンは画面に出しません）。

```sh
ssh lineworks-vps-user 'test "$(hostname)" = srv1508169 || exit 9; umask 077; cat > /opt/kyuka/cutover/gas_internal_token_prod.txt' < "$HOME/.secrets/kyuka/gas-internal-token.prod"
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
cp -p .env cutover/.env.before-cutover
DEP=〇〇                                   # B-4 の deploymentId
TOK=$(tr -d '\r\n' < cutover/gas_internal_token_prod.txt)
test ${#TOK} -ge 32 || { echo "TOKEN TOO SHORT"; exit 8; }
sed -i "s|^APP_ENV=.*|APP_ENV=production|" .env
sed -i "s|^GAS_INTERNAL_URL=.*|GAS_INTERNAL_URL=https://script.google.com/macros/s/$DEP/exec|" .env
sed -i "s|^GAS_INTERNAL_TOKEN=.*|GAS_INTERNAL_TOKEN=$TOK|" .env
grep -q '^PROD_GAS_INTERNAL_DEPLOYMENT_ID=' .env && sed -i "s|^PROD_GAS_INTERNAL_DEPLOYMENT_ID=.*|PROD_GAS_INTERNAL_DEPLOYMENT_ID=$DEP|" .env || echo "PROD_GAS_INTERNAL_DEPLOYMENT_ID=$DEP" >> .env
rm -f cutover/gas_internal_token_prod.txt
chmod 600 .env
grep -E '^(APP_ENV|APP_IMAGE_TAG|GAS_INTERNAL_URL|GAS_INTERNAL_AUTH|PROD_GAS_INTERNAL_DEPLOYMENT_ID|PUBLIC_ORIGIN|DB_SHEET_TIMEZONE|PROD_SPREADSHEET_ID)=' .env
awk -F= '$1=="GAS_INTERNAL_TOKEN"{print "GAS_INTERNAL_TOKEN length", length(substr($0, index($0,"=")+1))}' .env
stat -c '%a %n' .env
sg docker -c "docker image ls kyuka-app"
EOF
```

確かめ: `APP_ENV=production`、`GAS_INTERNAL_URL` が B-4 のデプロイ、`PROD_GAS_INTERNAL_DEPLOYMENT_ID` が同じ ID、`GAS_INTERNAL_AUTH=bearer`、`DB_SHEET_TIMEZONE=Etc/GMT`、トークンの長さ 64、`600 .env`、`APP_IMAGE_TAG` が最新のイメージのタグ（〇〇。古いタグだと `up -d` で古いコードに戻る＝罠 24e）。ステージングの `.env` は `cutover/.env.before-cutover` に残ります。戻す: `cp -p cutover/.env.before-cutover .env`。

### B-11 控え（本番 DB と台帳の複製）

`tools/cutover_backup.py` で、本番の DB スプレッドシートと台帳を Drive API の `files.copy`（委任トークン）で複製します。元は許可リストの 2 つだけで（ID は引数で受け取らない）、複製はマイドライブの直下に置き、共有しません。

```sh
timeout 590 ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
R=cutover/backup_report_$(date +%Y%m%d_%H%M).json
if sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api python tools/cutover_backup.py" < /dev/null > "$R"; then echo "exit=0"; else echo "exit=$?"; fi
grep -E '"(appEnv|label|copyId|ok)"' "$R"
EOF
```

確かめ: 終了コード 0（1 は行数が合わない冊がある、2 は止めた）、`appEnv: production`、2 冊とも `"ok": true`（すべてのタブの A 列の行数が元と控えで同じ）。**DB の控えの `copyId` を控えます**（〇〇。B-15・B-16・C 章で `tools/restore_tabs_from_copy.py --from-copy` に渡す）。`ok: false` ならもう一度流します。

### B-12 表を空にして移行

先に `--dry-run`（読むだけ）で数字を見て、よければ `--truncate` で本実行します（`--truncate` は業務の表と `jobs`・`idempotency_keys`・`sync_runs`・`sessions` を空にしてから入れる。`jobs` を残すとステージングで積んだ仕事が本番の GAS へ流れ、`sessions` を残すと受け入れテストのセッションが本番で生きる＝罠 24b）。`audit_log` は手で空にします。

```sh
timeout 590 ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
S=$(date +%H%M)
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api python tools/migrate_from_sheets.py --source sheets --dry-run" < /dev/null | tee cutover/migrate_dryrun_$S.json
echo "exit=${PIPESTATUS[0]}"
EOF
```

`appEnv: production`、`sheetTimeZone: Etc/GMT`、`droppedColumns: {}`、`missingColumns: []`、`summary` の件数がリハーサル以上（2026-11-26 は 279 件。承認済 260・取消 18・申請中 1）、`skippedBlankIdRows` が既知の 52 行、`settingsForEnv` の `PDF_FOLDER_ID`・`LEDGER_SSID` が本番の ID、`stagingPdfAndSignBlanked: null`（本番なので PDF とサインの URL を消さない）を見て、よければ本実行します。

```sh
timeout 590 ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
psql_q() { sg docker -c "docker compose exec -T db psql -U kyuka -d kyuka -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
TRUNCATE audit_log RESTART IDENTITY;
SQL
S=$(date +%H%M)
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api python tools/migrate_from_sheets.py --source sheets --truncate" < /dev/null | tee cutover/migrate_run_$S.json
echo "exit=${PIPESTATUS[0]}"
EOF
```

確かめ: 終了コード 0（2 は止めた、1 は入れた後の照合が合わない）、`verifyProblems: []`（件数・状態別・休暇日と振替元の JST の日付・申請日時と承認日時・作業員 ID・年度、移した直後の照合の差 0、DB から作った集計と本番の `V_LEAVE_SUMMARY` の 12 列が同じ）。戻す: ダンプから `pg_restore`（`sg docker -c "docker compose exec -T db pg_restore -U kyuka -d kyuka --clean --if-exists" < cutover/kyuka_before_cutover_<日時>.dump`）して B-12 をやり直す。

### B-13 照合

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
psql_q() { sg docker -c "docker compose exec -T db psql -U kyuka -d kyuka -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT count(*) total FROM leave_requests;
SELECT status, count(*) FROM leave_requests GROUP BY 1 ORDER BY 1;
SELECT leave_kubun, count(*) FROM leave_requests GROUP BY 1 ORDER BY 1;
SELECT leave_type, count(*) FROM leave_requests GROUP BY 1 ORDER BY 1;
SELECT count(*) FILTER (WHERE status = '申請中') submitted,
       count(*) FILTER (WHERE status = '承認済' AND approved_at2 IS NULL AND btrim(approved_at2_text) = '') waiting_2nd,
       count(*) FILTER (WHERE status = '承認済' AND pdf_url = '') approved_no_pdf,
       count(*) FILTER (WHERE approved_at IS NOT NULL) with_approved_at,
       count(*) FILTER (WHERE sign_image_url <> '') with_sign,
       count(*) FILTER (WHERE pdf_url <> '') with_pdf,
       count(*) FILTER (WHERE leave_date IS NULL) leave_not_date,
       count(*) FILTER (WHERE sheet_row IS NULL) no_sheet_row
FROM leave_requests;
SELECT fiscal_year, count(*) FROM leave_requests GROUP BY 1 ORDER BY 1;
SELECT key, value FROM settings WHERE key IN ('PDF_FOLDER_ID', 'LEDGER_SSID', 'MASTER_SOURCE_SSID', 'CALENDAR_SOURCE_SSID', 'APP_URL', 'MASTER_SOURCE_WORKER_SHEET', 'CALENDAR_SOURCE_SHEET') ORDER BY key;
SELECT (SELECT count(*) FROM jobs) jobs, (SELECT count(*) FROM sessions) sessions, (SELECT count(*) FROM google_credentials) delegate;
SQL
grep -E '^(PROD_SPREADSHEET_ID|PROD_PDF_FOLDER_ID|PROD_LEDGER_SPREADSHEET_ID|MASTER_SOURCE_SPREADSHEET_ID|CALENDAR_SOURCE_SPREADSHEET_ID|DB_SHEET_TIMEZONE)=' .env
sg docker -c "docker compose run --rm --no-deps -T api python -m app.worker run sheet_reconcile --check" < /dev/null
EOF
```

確かめ:
- 件数・状態別・区分別・種類別が、dry-run の報告と B-1 の数（シート）と同じ。進行中（`submitted`・`waiting_2nd`・`approved_no_pdf`）が B-1 と同じ。`leave_not_date` 0、`no_sheet_row` 0。
- Settings の `PDF_FOLDER_ID`＝`PROD_PDF_FOLDER_ID`、`LEDGER_SSID`＝`PROD_LEDGER_SPREADSHEET_ID`、`MASTER_SOURCE_SSID`＝`MASTER_SOURCE_SPREADSHEET_ID`、`CALENDAR_SOURCE_SSID`＝`CALENDAR_SOURCE_SPREADSHEET_ID`。`APP_URL` は @106 の `/exec`（B-19 で書き換える前の値）。
- `jobs 0`・`sessions 0`・`delegate 1`。
- `sheet_reconcile: ok=True` で差 0（`--check` は書きません）。

戻す: 合わなければ B-12 からやり直すか、C へ。

### B-14 api だけ起動して、安全装置と `gas_ping` を確かめる

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
sg docker -c "docker compose up -d api"
sleep 20
sg docker -c "docker compose ps"
sg docker -c "docker compose exec -T web wget -qO- http://127.0.0.1/healthz" < /dev/null; echo
sg docker -c "docker compose exec -T api env PYTHONPATH=/app python -" <<'PY'
from app import safety
from app import config as C
from app.config import get_settings
S = get_settings()
print("env", S.app_env, "/ writable", safety.writable_spreadsheet_id(S), "(本番 DB 16v1tQNC… であること)")
print("internal deployment", C.gas_deployment_id(S.gas_internal_url))
def refuse(label, sid, tab, purpose):
    try:
        safety.assert_writable(sid, tab, purpose, S); print("NG: allowed", label)
    except safety.UnsafeWriteError:
        print("OK: refused", label)
def allow(label, sid, tab, purpose):
    try:
        safety.assert_writable(sid, tab, purpose, S); print("OK: allowed", label)
    except Exception as e:
        print("NG: refused", label, e)
refuse("staging DB", C.STAGING_SPREADSHEET_ID_KNOWN, "T_LEAVE_REQUEST", "row_sync")
refuse("staging ledger", C.STAGING_LEDGER_SPREADSHEET_ID_KNOWN, "T_LEAVE_REQUEST", "row_sync")
refuse("staging pdf folder", C.STAGING_PDF_FOLDER_ID_KNOWN, "T_LEAVE_REQUEST", "row_sync")
refuse("master source", C.MASTER_SOURCE_SPREADSHEET_ID_FIXED, "作業員マスタ", "master_sync")
refuse("calendar source (残業 DB)", C.CALENDAR_SOURCE_SPREADSHEET_ID_FIXED, "社内カレンダーマスタ", "master_sync")
refuse("portal", "1DdhbJuz-zyS_2pKcM8y6Xf2pHofuzwNDgoW9eWcB0yg", "APP_MASTER", "row_sync")
refuse("prod ledger", C.PROD_LEDGER_SPREADSHEET_ID_FIXED, "T_LEAVE_REQUEST", "row_sync")
refuse("prod pdf folder", C.PROD_PDF_FOLDER_ID_FIXED, "T_LEAVE_REQUEST", "row_sync")
refuse("prod DB M_SYSTEM_SETTING", C.PROD_SPREADSHEET_ID_FIXED, "M_SYSTEM_SETTING", "row_sync")
refuse("prod DB M_LOOKUP", C.PROD_SPREADSHEET_ID_FIXED, "M_LOOKUP", "master_sync")
refuse("prod DB T_LEAVE_REQUEST by master_sync", C.PROD_SPREADSHEET_ID_FIXED, "T_LEAVE_REQUEST", "master_sync")
refuse("prod DB V_LEAVE_SUMMARY by row_sync", C.PROD_SPREADSHEET_ID_FIXED, "V_LEAVE_SUMMARY", "row_sync")
allow("prod DB T_LEAVE_REQUEST", C.PROD_SPREADSHEET_ID_FIXED, "T_LEAVE_REQUEST", "row_sync")
allow("prod DB T_LEAVE_REQUEST reconcile", C.PROD_SPREADSHEET_ID_FIXED, "T_LEAVE_REQUEST", "reconcile")
allow("prod DB V_LEAVE_SUMMARY", C.PROD_SPREADSHEET_ID_FIXED, "V_LEAVE_SUMMARY", "summary_sheet")
for tab in ("作業員マスタ", "M_CALENDAR", "M_STAMP", "M_DEPT_APPROVERS"):
    allow("prod DB " + tab, C.PROD_SPREADSHEET_ID_FIXED, tab, "master_sync")
PY
sg docker -c "docker compose exec -T api python -m app.worker run gas_ping" < /dev/null
EOF
curl -s https://leave.lineworks-local.net/healthz; echo
```

確かめ: healthz が `{"ok":true,"env":"production"}`（中からも外からも）。`NG` が 0 行（1 行でもあれば先へ進まない）。`writable` が `16v1tQNC…`。`internal deployment` が B-4 の ID。`gas_ping: ok=True` で、`problems: []`・`migrationMode: True`・`auth: bearer`（`spreadsheetTimeZone` は GMT のため空でよい＝罠 68）。戻す: api のログ（`sg docker -c "docker compose logs --tail 100 api"`）を見て直すか、C へ。

### B-15 マスタ同期と設定の取り込み（worker はまだ起動しない）

```sh
timeout 590 ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
sg docker -c "docker compose exec -T api python -m app.worker run sync_masters" < /dev/null
sg docker -c "docker compose exec -T api python -m app.worker run sync_sheet_config" < /dev/null
psql_q() { sg docker -c "docker compose exec -T db psql -U kyuka -d kyuka -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT (SELECT count(*) FROM workers) workers, (SELECT count(*) FROM calendar) calendar,
       (SELECT count(*) FROM stamps) stamps, (SELECT count(*) FROM dept_approvers) dept_approvers,
       (SELECT count(*) FROM settings) settings, (SELECT count(*) FROM lookup) lookup;
SQL
COPY=〇〇   # B-11 の DB の控えの copyId
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api python tools/restore_tabs_from_copy.py --from-copy $COPY --tabs 作業員マスタ,M_CALENDAR,M_STAMP,M_DEPT_APPROVERS" < /dev/null
EOF
```

確かめ: `sync_masters: ok=True`（4 タブとも `ok`、作業員 50・カレンダー 130 前後・電子印と部署承認者が B-11 の控えと同じ行数）、`sync_sheet_config: ok=True`（Settings 14・M_LOOKUP 6 前後）。写し戻しの道具の確かめ（`--apply` なし＝書かない）で、4 タブとも控えと違うセルが 0（元のマスタが控えの後に変わっていなければ 0）。マスタ同期はシートへ書く前の関所を通ります（B-14 の `gas_ping` が 10 分以内で `migrationMode: True` のときだけ書く）。`writeBlocked` が出たら、`gas_ping` をもう一度流してから B-15 をやり直します。

戻す: 控えから 4 タブの値を写し戻す（上の最後の行に `--apply` を付ける。先に差の数を見る）。

### B-16 worker を起動して、ping・行の反映・集計タブ・照合を確かめる

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
grep -q '^SKIP_DAILY_ON_STARTUP=1$' .env || echo 'SKIP_DAILY_ON_STARTUP=1' >> .env
sg docker -c "docker compose up -d"
sleep 60
sg docker -c "docker compose ps"
sg docker -c "docker compose exec -T worker python -m app.worker run summary_sheet" < /dev/null
sg docker -c "docker compose exec -T worker python -m app.worker run sheet_reconcile --check" < /dev/null
psql_q() { sg docker -c "docker compose exec -T db psql -U kyuka -d kyuka -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT count(*) FILTER (WHERE desired_version > synced_version) row_sync_pending, count(*) FILTER (WHERE state = 'blocked') blocked FROM sheet_sync_state;
SELECT lane, state, count(*) FROM jobs GROUP BY 1, 2 ORDER BY 1, 2;
SELECT kind, ok, started_at, summary_json->'problems' problems, summary_json->'migrationMode' migrated FROM sync_runs ORDER BY id DESC LIMIT 10;
SQL
COPY=〇〇   # B-11 の DB の控えの copyId
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api python tools/restore_tabs_from_copy.py --from-copy $COPY --tabs T_LEAVE_REQUEST,V_LEAVE_SUMMARY" < /dev/null
sg docker -c "docker compose logs --since 2m worker" | grep -iE 'error|exception|unsafe' || echo "worker: エラーなし"
EOF
```

確かめ:
- 6 コンテナが `Up`。起動ログの `schedules=` に `sync_masters`・`summary_sheet`・`sheet_reconcile`・`sync_sheet_config`・`gas_ping`・`purge` がある（罠 73）。
- `summary_sheet: ok=True`（V_LEAVE_SUMMARY を書き直す）。`sheet_reconcile: ok=True` で差 0。
- `row_sync_pending 0`・`blocked 0`、jobs は 0 件（まだ誰も操作していない）、最新の `gas_ping` が `ok = t`・`migrated = true`。
- 写し戻しの道具の確かめ（書かない）で、`T_LEAVE_REQUEST`（V 列）と `V_LEAVE_SUMMARY` が控えと違うセル 0。`V_LEAVE_SUMMARY` に違いが出たら、どのセルか（最終有給日の型・警告レベル）を見て、オーナーに聞くまで先へ進みません。
- 管理者の画面の「運用」欄（`ops_getStatus`）の `sheetWriteGate.allowed` が `true`。
- worker の起動は `SKIP_DAILY_ON_STARTUP=1` なので、2:10・3:30・6:00・6:50 の枠の中で起動しても、その回の毎日の予定は動きません（`sync_runs` に `skipped`）。確かめが済んだら `.env` の `SKIP_DAILY_ON_STARTUP=1` の行を消します（次に起動し直すまで効き目は変わりません）。

戻す: `sg docker -c "docker compose stop worker"`、控えから写し戻して（`restore_tabs_from_copy.py --from-copy <控え> --tabs all --apply`）C へ。

### B-17 統合業務メールと同じ条件で、本番シートの件数を確かめる

「統合業務メールと同じ条件でシートを数える」をもう一度流します。

確かめ: `rows`・`byStatus`・`todayLeaves(7:00)`・`pending1(17:00)`・`pending2(18:00)`・`summaryRows`・`warnWorkers(今年度)` が B-1 と同じ。`leaveDateNotDate` 0。違えば C へ（シートの値と型が変わっている）。

### B-18 画面の確かめ（書き込みはしない）

オーナーのアカウントと、部署の共有アカウント 1 つで、JEV ハンズで画面を開いて見ます。

- TOP（承認待ちの一覧に B-1 の申請中が出る）、申請中 1 件のプレビュー、総務管理（2 次承認待ちの件数が B-1 と同じ。運用の欄はオーナーにだけ出る）、部署管理、取得状況、古い管理画面（`?page=admin`）。
- 古いリンク: `https://script.google.com/macros/s/AKfycbx7sB9x…/exec?page=preview&reqId=<申請中の REQ_ID>` → 移転ページ → ボタン → （セッションを空にしたので）ログイン → 新しいアプリの**同じ申請**のプレビュー。@104（`AKfycbwHI1uz…`）でも同じ。統合業務メールの本文のリンク（`?page=somuAdmin`）も移転ページから総務管理へ届く。
- 外から `curl -s https://leave.lineworks-local.net/healthz` が `{"ok":true,"env":"production"}`。受け入れの道具は本番では動かない（REFUSED で終わる）ので流しません。

**ここでは書き込みの操作（申請・承認・取消）をしません。** 最初の書き込みは翌朝の利用者の操作です（D 章で確かめる）。

### B-19 M_SYSTEM_SETTING の `APP_URL` を書き換える（K-11）

統合業務メールの本文のリンク（`APP_URL`）を新しいアドレスにします。移転モードの GAS は `APP_URL` を画面の移動に使わない（移転ページは `MIGRATED_TO_URL` を使う）ので、書き換えても害はありません。アプリの安全装置は M_SYSTEM_SETTING への書き込みを拒むので、api コンテナの中から委任トークンで直接書きます。既定は何も書かずに見るだけで、`APPLY=1` のときだけ、今の値が @106 の `/exec` のときに書きます。

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
sg docker -c "docker compose exec -T -e APPLY=0 api env PYTHONPATH=/app python -" <<'PY'
import json, os, urllib.parse, urllib.request
from app import config as C
from app.google_delegate import get_access_token
SID = C.PROD_SPREADSHEET_ID_FIXED
NEW = "https://leave.lineworks-local.net/"
OLD_MARK = "AKfycbx7sB9x"                 # 今の値は @106 の /exec のはず
H = {"Authorization": "Bearer " + get_access_token(), "Content-Type": "application/json"}
def call(method, url, body=None):
    req = urllib.request.Request(url, data=None if body is None else json.dumps(body).encode(), method=method, headers=H)
    return json.load(urllib.request.urlopen(req))
rows = call("GET", f"https://sheets.googleapis.com/v4/spreadsheets/{SID}/values/{urllib.parse.quote('M_SYSTEM_SETTING!A1:B200')}")["values"]
hit = [(i + 1, r) for i, r in enumerate(rows) if i > 0 and r and str(r[0]).strip() == "APP_URL"]
assert len(hit) == 1, f"APP_URL の行が {len(hit)} 行です（下の行が勝つ。人が確かめる）"
n, row = hit[0]
cur = row[1] if len(row) > 1 else ""
print(f"行 {n} 列 B: 今の APP_URL = {cur}")
if os.environ.get("APPLY") == "1":
    assert OLD_MARK in cur, "APP_URL が想定の古い値（@106 の /exec）ではないので書きません"
    call("PUT", f"https://sheets.googleapis.com/v4/spreadsheets/{SID}/values/{urllib.parse.quote(f'M_SYSTEM_SETTING!B{n}')}?valueInputOption=RAW", {"values": [[NEW]]})
    print("書き換えました →", NEW)
PY
EOF
```

1 回目（`APPLY=0`）で行と**今の値を控え**（〇〇。戻すときに使う）、同じコマンドを `APPLY=1` にしてもう一度流します。確かめ: もう一度 `APPLY=0` で流して `https://leave.lineworks-local.net/`。5 分以内に VPS の取り込み（`sync_sheet_config`）でも DB の `settings` の `APP_URL` が新しい値になる（すぐ見たいときは `sg docker -c "docker compose exec -T api python -m app.worker run sync_sheet_config" < /dev/null`）。戻す: 同じ形で `NEW` を控えた元の値に、`OLD_MARK` を `leave.lineworks-local.net` にして流す。

### B-20 社内ポータルの行き先

**前にオーナーへ声をかけます**（下の「オーナーへの連絡」の 3）。アプリの安全装置はポータルのシートへの書き込みを拒むので、api コンテナの中から委任トークンで直接書きます。key は値の後ろに空白があるので trim して探します。既定は見るだけで、`APPLY=1` のときだけ、url 列が想定の古い値のときに書きます。

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
sg docker -c "docker compose exec -T -e APPLY=0 api env PYTHONPATH=/app python -" <<'PY'
import json, os, urllib.parse, urllib.request
from app.google_delegate import get_access_token
PORTAL = "1DdhbJuz-zyS_2pKcM8y6Xf2pHofuzwNDgoW9eWcB0yg"
KEY = "leave_request"
NEW = "https://leave.lineworks-local.net/"
OLD_MARK = "AKfycbx7sB9x"                 # 今の url は @106 の /exec のはず
H = {"Authorization": "Bearer " + get_access_token(), "Content-Type": "application/json"}
def call(method, url, body=None):
    req = urllib.request.Request(url, data=None if body is None else json.dumps(body).encode(), method=method, headers=H)
    return json.load(urllib.request.urlopen(req))
rows = call("GET", f"https://sheets.googleapis.com/v4/spreadsheets/{PORTAL}/values/{urllib.parse.quote('APP_MASTER!A1:Z500')}")["values"]
head = [h.strip().lower() for h in rows[0]]
kc, uc = head.index("key"), head.index("url")
hit = [(i + 1, r) for i, r in enumerate(rows) if i > 0 and len(r) > kc and str(r[kc]).strip() == KEY]
assert len(hit) == 1, f"key {KEY} の行が {len(hit)} 行です"
n, row = hit[0]
cur = row[uc] if len(row) > uc else ""
col = chr(ord("A") + uc)
print(f"行 {n} 列 {col}: 今の url = {cur}")
if os.environ.get("APPLY") == "1":
    assert OLD_MARK in cur, "url が想定の古い値（@106 の /exec）ではないので書きません"
    call("PUT", f"https://sheets.googleapis.com/v4/spreadsheets/{PORTAL}/values/{urllib.parse.quote(f'APP_MASTER!{col}{n}')}?valueInputOption=RAW", {"values": [[NEW]]})
    print("書き換えました →", NEW)
PY
EOF
```

1 回目（`APPLY=0`）で行と**今の値を控え**（〇〇）、同じコマンドを `APPLY=1` にしてもう一度流します。確かめ: ポータルを開き直して、「休暇届け申請」のボタンの行き先が新しい URL（open_mode `same_tab` のまま）。戻す: 同じ形で `NEW` を控えた元の値に、`OLD_MARK` を `leave.lineworks-local.net` にして流す。

### B-21 最終確認

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
sg docker -c "docker compose ps"
psql_q() { sg docker -c "docker compose exec -T db psql -U kyuka -d kyuka -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT lane, state, count(*) FROM jobs GROUP BY 1, 2 ORDER BY 1, 2;
SELECT count(*) FILTER (WHERE desired_version > synced_version) row_sync_pending FROM sheet_sync_state;
SELECT kind, ok, started_at FROM sync_runs WHERE kind = 'gas_ping' ORDER BY id DESC LIMIT 1;
SQL
sg docker -c "docker compose logs --since 30m api worker" | grep -iE ' error|exception|traceback' || echo "エラーなし"
sed -i '/^SKIP_DAILY_ON_STARTUP=1$/d' .env && grep -c '^SKIP_DAILY_ON_STARTUP' .env || true
ls -l backups/ | tail -3
EOF
curl -s -o /dev/null -w '%{http_code}\n' https://leave.lineworks-local.net/healthz
```

確かめ: 6 コンテナが `Up`、jobs の pending・failed が 0、反映待ち 0、`gas_ping` が `t`、エラーなし、外から 200、`SKIP_DAILY_ON_STARTUP` の行が消えた（`0`）。GAS の「実行数」に、切り替え後のトリガーの実行が無い。`backups/` が空や 1KB 未満なら、起動直後の空の DB のダンプが捨てられただけ（罠 52）なので、`sg docker -c "docker compose exec -T backup /usr/local/bin/run_backup.sh" < /dev/null` で 1 回取り、`[backup] ok … bytes` を確かめます。

### B-22 記録と報告

引継ぎメモ（`docs/handoff-<日付>.md`）・メモリ・スキルの下書きを更新し、オーナーへ報告します。報告に入れるもの:

- 移した件数と状態別の内訳、シートとの一致、進行中の件数（申請中・2 次承認待ち）
- 本番 GAS: 移転モード入りの版 〇〇（@106・@104 の両方）、内部用のデプロイ ID、戻す版 106・104、消したトリガーの一覧
- 控えの場所: DB と台帳の複製の ID、`cutover/` のダンプ、`.env.before-cutover`、`APP_URL` とポータルの url の元の値
- 利用者から見た変化（ポータル、古いリンク、ログイン画面、統合業務メールのリンク）
- 判断して変えた点と理由、オーナーにお願いすること（翌朝の最初の PDF を見る、スマホ・タブレットでの表示とサインの写真など）

---

## C. 元に戻す（戻すと決めるのはオーナー。戻せる期間は 5 営業日）

決定 A でシートは行単位で常に最新、決定 B で PDF・サイン画像・台帳・メールは初めから GAS が作っています。そのため、書き出しを待つ・写し戻す作業は要らず、「VPS を止めて、GAS の移転モードを外す」が中心です。戻した GAS は、シートの最新の状態からそのまま動きます。

1. **C-1 VPS の api を止める**（新しい書き込みを止める。worker は流しきりのため動かしたまま）。

   ```sh
   ssh lineworks-vps-user 'test "$(hostname)" = srv1508169 || exit 9; cd /opt/kyuka && sg docker -c "docker compose stop api" && sg docker -c "docker compose ps"'
   ```

2. **C-2 worker の待ちを流しきって止める**。行の反映の待ち 0、GAS の仕事の待ち 0（GAS は移転モードのまま内部用の入口が動くので数分）になるまで、次を繰り返します。`failed` や流しきれない GAS の仕事は、申請 ID と action を控えます（C-8 で使う）。blocked の行は流しきれないので、`reappend` で末尾に書き直してから流します。照合を `--check` で 1 回流し、差 0 を確かめます。**最後に worker を止め、コンテナを消します**（`docker compose rm -sf worker`。止めるだけだと、後の `docker compose up -d` で勝手に戻ります。本番の worker は GAS が移転モードでない間はシートへ書きませんが、念のため消します）。

   ```sh
   ssh lineworks-vps-user 'bash -s' <<'EOF'
   set -eu
   test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
   cd /opt/kyuka
   psql_q() { sg docker -c "docker compose exec -T db psql -U kyuka -d kyuka -v ON_ERROR_STOP=1 -X -P footer=off"; }
   psql_q <<'SQL'
   SELECT count(*) FILTER (WHERE desired_version > synced_version) row_sync_pending, count(*) FILTER (WHERE state = 'blocked') blocked FROM sheet_sync_state;
   SELECT id, request_id, kind, state, attempts, left(last_error, 120) FROM jobs WHERE state NOT IN ('done', 'skipped') ORDER BY id;
   SQL
   EOF
   # blocked が 1 以上なら（人がシートの行を消したなど）:
   #   sg docker -c "docker compose exec -T worker python -m app.services.row_sync status" < /dev/null
   #   sg docker -c "docker compose exec -T worker python -m app.services.row_sync reappend <REQ_ID>" < /dev/null
   # 待ちと blocked が 0 になったら:
   ssh lineworks-vps-user 'test "$(hostname)" = srv1508169 || exit 9; cd /opt/kyuka && sg docker -c "docker compose exec -T worker python -m app.worker run sheet_reconcile --check" < /dev/null && sg docker -c "docker compose rm -sf worker" && sg docker -c "docker compose ps -a"'
   ```

3. **C-3 GAS: 移転モードを外す**。GAS エディタで `exitMigrationMode`。確かめ: `showMigrationStatus` で「移転モード: なし」。内部用の入口は token が無い呼び出しには何もしないので、残してかまいません。
4. **C-4 GAS: トリガーを戻す**。`restoreTriggersAfterRollback`（`syncAllMasters` の毎日 6 時台の 1 本。同じ関数のトリガーが既にあれば作りません）。B-7 で昔の通知・警告のトリガーも消していて、それも戻すとオーナーが決めたときだけ `restoreAllTriggersAfterRollback`（統合業務メールと二重に届くおそれがあるので、既定は戻さない）。確かめ: トリガーの画面に `syncAllMasters` の 1 本。
5. **C-5 GAS: 両方の画面のデプロイを 106・104 に戻す**。

   ```sh
   CLASP="npx -y @google/clasp@3.4.1"
   $CLASP update-deployment AKfycbx7sB9x7iLigEqZX9LuErK95R2ZC0cMOyU9efr6pIbpnqAnnDC3o5_DGFbtwgSh64j2 --versionNumber 106 --description "rollback"
   $CLASP update-deployment AKfycbwHI1uz6tcer59sRC4dYEACDaevSn9a_eJa3J0E8LZgjdWCGPJVaqr_Qa9L-HXVKI8k --versionNumber 104 --description "rollback"
   $CLASP list-deployments --json | grep -A6 -E 'AKfycbx7sB9x|AKfycbwHI1uz'
   ```

   確かめ: `versionNumber` が 106 と 104。変わっていなければ B-3 の Apps Script API のコマンドで、`VERSION=106`（@106）と `VERSION=104`（@104）を 1 つずつ当てます（罠 60）。HEAD は移転モード入りのコードのままでも、`MIGRATED_TO_URL` が無ければ今と同じ動きです（元の HEAD に戻すなら `master` の `src/` を `clasp push --force`）。
6. **C-6 `APP_URL` とポータルの url を元に戻す**。B-19・B-20 のコマンドで、`NEW` を控えた元の値（@106 の `/exec`）に、`OLD_MARK` を `leave.lineworks-local.net` にして `APPLY=1` で流します。確かめ: `APPLY=0` で元の値、ポータルのボタンの行き先が @106 の `/exec`。
7. **C-7 マスタ同期と集計を手で 1 回**（必要なら）: GAS エディタで `syncAllMasters` を実行し、ログの `OK` を確かめます（VPS が書いていたマスタ 4 タブに、GAS の同期が今の形で書けること）。集計タブはメニューの「サマリー再構築」（または画面を開くと GAS が書き直す）。
8. **C-8 控えた仕事を GAS で処理する**: C-2 で控えた申請ごとに、オーナーと相談して、1 次承認のサイン・メール・台帳（`updateLedger_`）や 2 次承認の PDF（`generateLeavePdf_`）を GAS エディタから手で行います。
9. **C-9 確かめ**: 本番 `/exec`（@106）で一覧が出て、VPS で最後に作られた申請が見えること。@104 も今の画面が開くこと。
10. **C-10 `.env`**: 必要なら `cp -p cutover/.env.before-cutover .env`。api と worker は、次に切り替えると決めるまで止めたままにします。戻した後は VPS の DB を使いません。もう一度切り替えるときは B-9 からやり直します。
11. **C-11 利用者への案内**: オーナーと相談します（「オーナーへの連絡」の 5）。

### 当日の途中で戻すとき（B の「戻す範囲」）

- B-3〜B-8 の途中: 実行した GAS の関数を逆の順に戻します（`restoreTriggersAfterRollback` → `exitMigrationMode` → 両方のデプロイを 106・104 に）。VPS は止めただけなので `sg docker -c "docker compose start api worker"`（ステージングのまま）。
- B-9〜B-13 の途中: 表はダンプから `pg_restore`、`.env` は控えを戻し、`docker compose start api worker`。GAS は上と同じに戻します。シートはまだ VPS から書かれていません（B-12 の移行は読むだけ）。
- B-14 以降: C 章の手順で戻します。B-15・B-16 でシートを書いた後に戻すときは、控えとの差を `restore_tabs_from_copy.py`（`--apply` なし）で先に見ます。

---

## D. 見守り（戻せる期間の 5 営業日）

翌朝（切り替えの翌日）に確かめること:

- 6:00 の VPS のマスタ同期（`sync_masters`）が `ok`。6:50 の集計タブの書き直し（`summary_sheet`）が `ok`。2:10 の毎晩の照合（`sheet_reconcile`）が差 0。3:00 のバックアップが 1KB 超。
- 7:00 の統合業務メール「本日の休暇者」がいつもどおり届き、件数が「統合業務メールと同じ条件でシートを数える」の `todayLeaves(7:00)` と VPS の画面（TOP の予定）と同じ（オーナーの受信箱）。
- 最初の申請が 1 分以内にシートの `T_LEAVE_REQUEST` の末尾（REQ_ID のある最後の行の次）に出て、休暇日が日付のセル（`leaveDateNotDate` 0）。
- 最初の 1 次承認でサイン画像（手書きのとき）と承認完了メール、最初の 2 次承認で PDF が GAS の内部用の入口で作られ、DB とシートの G 列（`SIGN_IMAGE_URL`・`PDF_URL`）が同じ。本番で最初に作られた PDF 1 件をオーナーが見る。
- 17:00「1 次未承認」・18:00「2 次承認待ち」の件数が、`pending1(17:00)`・`pending2(18:00)` と VPS の画面（TOP の承認待ち・総務管理の 2 次承認待ち）と同じ。

毎日（朝 7:00 の後・夕方 17:00 と 18:00 の後）見るもの:

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
psql_q() { sg docker -c "docker compose exec -T db psql -U kyuka -d kyuka -v ON_ERROR_STOP=1 -X -P footer=off"; }
psql_q <<'SQL'
SELECT lane, state, count(*) FROM jobs GROUP BY 1, 2 ORDER BY 1, 2;
SELECT id, request_id, kind, attempts, left(last_error, 120) FROM jobs WHERE state = 'failed' ORDER BY id;
SELECT count(*) FILTER (WHERE desired_version > synced_version) pending,
       count(*) FILTER (WHERE state = 'blocked') blocked,
       coalesce(extract(epoch FROM now() - min(first_dirty_at) FILTER (WHERE desired_version > synced_version)), 0)::int oldest_seconds
FROM sheet_sync_state;
SELECT kind, ok, started_at FROM sync_runs WHERE kind NOT IN ('gas_ping') ORDER BY id DESC LIMIT 12;
SELECT kind, ok, started_at, summary_json->'problems' problems FROM sync_runs WHERE kind = 'gas_ping' ORDER BY id DESC LIMIT 1;
SELECT kind, started_at, summary_json FROM sync_runs WHERE kind = 'gas_job_failed' ORDER BY id DESC LIMIT 5;
SQL
ls -l backups/ | tail -3
sg docker -c "docker compose logs --since 24h api worker" | grep -ciE ' error|exception' || true
EOF
```

- 失敗した仕事（`failed`・`gas_job_failed`）が 0。あれば管理者の画面の「運用」欄（`ops_retryJob`）でやり直す。行の反映の待ち・blocked が 0、最も古い待ちが 60 秒以内（5 分を超えたら遅れ）。
- マスタ同期（6:00）・集計タブ（6:50）・毎晩の照合（2:10）・設定の取り込み（5 分ごと）・`gas_ping`（5 分ごと）が `ok`。バックアップ（3:00）が 1KB 超。
- 7:00・17:00・18:00 の統合業務メールの件数を、上の「シートを数える」の結果と比べる（オーナーの受信箱。見守りの初日は 3 回とも）。
- 管理者の画面の「運用」欄でも同じものが見えます。
- 期間が過ぎたら: ステージングの資源（DB・台帳・PDF フォルダの複製、写しの GAS）の片付けを、オーナーに確かめてから行います。VPS の外へのバックアップと死活監視は、後日 4 アプリまとめて決めます（K-0）。

---

## オーナーへの連絡

1. **前日まで**: 当日の日付（夜。K-8）、利用者への案内を出すか（前例は「静かに」で出していない）、部署の共有アカウントで前日に 1 回ログインしてもらうか（A-7）、本番 GAS の変更（push・版・移転モード・トリガーの削除）と `APP_URL`・ポータルの書き換えの許可。
2. **当日の開始前**: 「切り替えを始めます」と声をかけ、合図をもらう（K-3・K-8）。
3. **ポータルの書き換えの前**（B-20）: 「ポータルの行き先を新しいアドレスに変えます」と声をかける。
4. **終わったら**: B-22 の報告。お願いすること（翌朝最初の PDF を見る、スマホ・現場のタブレットでの表示と手書きサインの写真など）を最後に分けて書く。
5. **戻すとき**: 戻すかはオーナーが決める。戻した後の利用者への案内（「元のアドレスに戻しました」）もオーナーと相談する。

## 確かめること

- 移転モード入りの版の番号 〇〇（B-3）と、@106・@104 の両方に当たったこと（`list-deployments --json`）。
- 本番の内部用のデプロイ ID 〇〇（B-4。`.env` の `GAS_INTERNAL_URL`・`PROD_GAS_INTERNAL_DEPLOYMENT_ID`）。ステージングの `.env` にも `PROD_GAS_INTERNAL_DEPLOYMENT_ID` を書くかは、ステージングを片付けるまでに決める（書くとステージングの取り違えの歯止めになる）。
- 本番の最新のイメージのタグ 〇〇（B-10 の `APP_IMAGE_TAG`）。
- B-11 の DB の控えの `copyId` 〇〇（B-15・B-16・C 章の写し戻しで使う）。
- `APP_URL` とポータルの url の元の値 〇〇（B-19・B-20。C-6 で書き戻す）。
- 本番のトリガーの実物（A-5。`syncAllMasters` のほかに残っていないか）【要確認】。
- 社内ポータルのスプレッドシート ID `1DdhbJuz…` が「社内ダッシュボード」であること【要確認】（`docs/vps-migration/01-dependencies.md` の見込み。B-20 のコマンドは key の行が 1 行でなければ止まる）。
- リハーサルで控えた件数（B-12・B-13 の比べる相手）。
- `M_STAMP`・`M_DEPT_APPROVERS` の行数（B-15。gviz では 6・13、Sheets API ではタブの行数 9・14＝見出しを含む）。
