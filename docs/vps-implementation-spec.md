# 休暇届 VPS 版 実装仕様（エージェント共通の取り決め）

作成: 2026-11-26（段階 2）。対象: 本番 GAS「休暇届けapp」（scriptId `15aGUL5i…`、画面のデプロイ @106・@104）を VPS（FastAPI＋PostgreSQL＋worker＋nginx＋Cloudflare Tunnel）へ移す。
読む順: この文書 → `docs/vps-migration/06-day1-findings-and-decisions.md`（オーナーの決定 K-0〜K-7 と本番の実物）→ 棚卸し 4 本（`inv-1-workflow.md`・`inv-2-screens.md`・`inv-3-mail-pdf-ledger.md`・`inv-4-data-ops-history.md`）の自分の担当に関わる節 → 自分の担当の GAS コード（`src/`）。

**方針: 今の GAS 版の機能と使い勝手をまったく残す。変えるのは 11 章「変えるもの」だけ。** 迷ったら GAS のコードの動きに合わせ、報告に書く。資料が食い違えば、`06`（本番の実物とオーナーの決定）が勝ち、次に GAS のコード、次にこの文書、次に棚卸し。棚卸しの記述と GAS のコードが違えば、コードが正。

> 書き方: 「必須」は外せない約束。「既定」はオーナーが未回答の癖（15 章）にディレクターが決めた値で、`app/domain/quirks.py`（W）の定数 1 か所か、小さな差分で変えられるように作る。〇〇 は段階 2b のスパイク（14 章）で決まる値。「推測」はコードで確かめていないこと。秘密の値はこの文書にも git にも書かない。

参照実装（読むだけ。書き換えない）:

- 姉妹アプリ「残業・休日出勤申請」: `D:\projects\zangyo-kyujitsu-app\github-repo\`（`vps-app/` 全体、`src/Internal.gs`・`src/Cutover.gs`、`docs/vps-implementation-spec.md`・`docs/vps-migration-plan.md`）。**土台・行の反映・GAS の仕事・画面の移し方は、ここを写して直す。**
- 共通スキル: `C:\Users\imaizumi.LINEWORKS-NET\.claude\skills\gas-to-vps-migration\`（`references/02-architecture.md` 18 節、`04-data-migration.md` 10 節、`07-pitfalls.md` の罠 1〜77）

---

## 目次

1. 確定値
2. リポジトリの形と、写すもの・捨てるもの
3. 画面（7 ページ＋権限エラー）
4. RPC（画面の 27 関数・管理者の ops_*）
5. DB の表
6. シートへの書き込み（T_LEAVE_REQUEST の行単位の反映・V_LEAVE_SUMMARY・マスタ・設定）
7. GAS の内部用の入口と、きっかけごとの呼び順
8. 同じにしなければならない業務の決まり
9. 安全装置
10. 担当と境目（担当をまたぐ関数の表、作る順）
11. 変えるもの／変えないもの
12. テストと受け入れ表
13. 意図して残す固有の文字列
14. スパイクで決めること
15. オーナーに聞くこと

---

## 1. 確定値

| 値 | このアプリ | 出典 |
|---|---|---|
| アプリ名（フォルダ・DB・接頭辞） | `kyuka` | 06 K-1 |
| 公開 URL | `https://leave.lineworks-local.net`（staging も同じホスト名。切り替えで `APP_ENV` だけ変える） | 06 K-1 |
| VPS | `srv1508169`（`ssh lineworks-vps-user`。作業前に `hostname`）。`srv1300618` には触れない | スキル |
| VPS の置き場 | `/opt/kyuka`（`vps-app/` の中身を `git archive` で直置き。`.env` も直下、権限 600） | 06 K-1・2-1 |
| Tunnel 名 | `leave-vps`。転送先 `http://web:80` | 06 K-1 |
| docker のサブネット | `10.94.0.0/24`。web の固定 IP `10.94.0.10` | 06 2-1（空きを確認済み） |
| compose のプロジェクト名 | `kyuka` | — |
| コンテナ名 | `kyuka-db`・`kyuka-api`・`kyuka-worker`・`kyuka-web`・`kyuka-cloudflared`・`kyuka-backup` | 00 3 節 |
| ボリューム・ネットワーク | `kyuka-pgdata`・`kyuka-net` | — |
| DB | PostgreSQL 16。DB 名 `kyuka`、ユーザー `kyuka`。結合テスト用は `kyuka_test`（名前に `test` が無ければ conftest が拒む） | — |
| クッキー | `__Host-kyuka_session`・`__Host-kyuka_login`・`__Host-kyuka_delegate` | — |
| ログインの持ち時間 | 90 日（無操作でも 90 日） | 06 K-6 |
| バックアップのダンプ名 | `kyuka_*.dump`（毎日 3 時、1KB 未満は FAILED）。切り替えの控えは `/opt/kyuka/cutover/` | 罠 25・26 |
| 手元の秘密の置き場 | `C:\Users\imaizumi.LINEWORKS-NET\.secrets\kyuka\` | 00 3 節 |
| OAuth クライアント | GCP `sebastian01` の「Time Off Request (VPS)」。リダイレクト URI は `/auth/callback` と `/auth/google-delegate/callback` | 06 2-6 |
| 本番 DB スプレッドシート | `16v1tQNCFkDTk2mnAZ4E2eEdHeUQ2Bepdcsn5vBAmllw`「休暇届けapp」（GAS が紐づく） | 00 2 節 |
| マスタ元（作業員） | `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ` のタブ `作業員マスタ`（読むだけ） | 06 2-5 |
| カレンダー・印・部署承認者の元 | 残業 DB `1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ` のタブ `社内カレンダーマスタ`・`StampMap`・`DeptApprovers`（**読むだけ。絶対に書かない**） | 06 2-5・01 1 節 |
| 台帳 | `1B7tnsT4lr80o3pFMBxB2zEhsx2UW2PCL48YTQc9i1eU`「休暇届台帳（総務確認用）」（VPS は書かない。GAS が書く） | 06 2-5 |
| 本番 PDF フォルダ | `1MQibR08wWsE_OGurO9RocEEDGgDOwo-w`（VPS は書かない。GAS が書く） | 06 2-5 |
| GAS scriptId（本番） | `15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i` | 00 2 節 |
| GAS の画面のデプロイ | @106 `AKfycbx7sB9x7iLigEqZX9LuErK95R2ZC0cMOyU9efr6pIbpnqAnnDC3o5_DGFbtwgSh64j2`（本番。Settings `APP_URL` とポータルが指す）、@104 `AKfycbwHI1uz6tcer59sRC4dYEACDaevSn9a_eJa3J0E8LZgjdWCGPJVaqr_Qa9L-HXVKI8k`（ブックマークだけ）。どちらも `access: DOMAIN`・`executeAs: USER_DEPLOYING` | 06 2-2 |
| GAS の内部用のデプロイ | 〇〇（スパイクで作る。7-1） | 14 章 |
| 社内ポータル | 「社内ダッシュボード」APP_MASTER の key `leave_request`（値の後ろに空白。探すときは trim） | 06 2-3 |
| 委任の持ち主・オーナー | `imaizumi@lineworks-local.info`。ドメイン `lineworks-local.info` | 00 3 節 |
| 統合業務メール（外の読み手） | GAS `1lb7bmF9…`。T_LEAVE_REQUEST・V_LEAVE_SUMMARY・M_SYSTEM_SETTING を 7:00・17:00・18:00 に読む。**変えない** | 01 2 節 |
| Sebastian | このシートに書いていない（確認済み）。`sheet_intake` は作らない | 06 2-4 |

### 1-1. 環境変数（`.env`、`vps-app/.env.example` に全キーを書く）

読むのは `api/app/config.py` の 1 か所だけ。値が欠ければ api も worker も起動しない（fail-closed）。「秘密」は `.env.example` に値を書かず `CHANGE_ME` だけ置く。形は残業の `.env.example` を写し、**`CALENDAR_SOURCE_SPREADSHEET_ID`・`GAS_INTERNAL_AUTH` を足し、`PORTAL_URL`・`SEBASTIAN_*`・`SHEET_INTAKE_*` を外す**（F は着手済み）。

| キー | 秘密 | 中身・検証 |
|---|---|---|
| `APP_ENV` | — | `staging`／`production` |
| `DATABASE_URL` | 秘密 | `postgresql+psycopg://kyuka:<pw>@db:5432/kyuka` |
| `POSTGRES_USER`／`POSTGRES_DB` | — | `kyuka`／`kyuka` |
| `POSTGRES_PASSWORD` | 秘密 | |
| `PUBLIC_ORIGIN` | — | `https://leave.lineworks-local.net`（`https://` 必須、末尾 `/` なし） |
| `GOOGLE_OAUTH_CLIENT_ID` | — | 「Time Off Request (VPS)」 |
| `GOOGLE_OAUTH_CLIENT_SECRET` | 秘密 | |
| `SESSION_SECRET` | 秘密 | 32 文字以上 |
| `SESSION_MAX_AGE_DAYS` | — | **`90`** |
| `SESSION_IDLE_HOURS` | — | **`2160`**（90 日） |
| `ALLOWED_DOMAIN` | — | `lineworks-local.info` |
| `BOOTSTRAP_ADMINS` | — | `imaizumi@lineworks-local.info`（空なら起動しない。**運用の関数 `ops_*` にだけ効く**。8-9） |
| `FERNET_KEY` | 秘密 | 委任トークンの暗号化 |
| `DELEGATE_EMAIL` | — | `imaizumi@lineworks-local.info` |
| `PROD_SPREADSHEET_ID` | — | 本番 DB。定数と違えば起動しない |
| `STAGING_SPREADSHEET_ID` | — | staging で必須。本番 DB の「コピーを作成」。本番と同じなら起動しない |
| `MASTER_SOURCE_SPREADSHEET_ID` | — | `1iu5Hoak…`。定数と違えば起動しない |
| `CALENDAR_SOURCE_SPREADSHEET_ID` | — | `1Knx_kaQ…`。定数と違えば起動しない（staging も本番の残業 DB を読むだけ） |
| `PROD_PDF_FOLDER_ID`／`STAGING_PDF_FOLDER_ID` | — | VPS は書かない。`internal.ping` の照合だけ。staging では両方必須で、同じなら起動しない |
| `PROD_LEDGER_SPREADSHEET_ID`／`STAGING_LEDGER_SPREADSHEET_ID` | — | 同上（台帳。staging は台帳のコピー） |
| `GAS_INTERNAL_URL` | — | GAS の内部用デプロイの `/exec`。`^https://script\.google\.com/macros/s/[A-Za-z0-9_-]+/exec$`。画面のデプロイ（@106・@104）の ID なら起動しない |
| `GAS_INTERNAL_TOKEN` | 秘密 | GAS のスクリプトのプロパティ `VPS_INTERNAL_TOKEN` と同じ値（32 バイト以上の乱数） |
| `GAS_INTERNAL_AUTH` | — | `bearer`（第 1 案）／`anonymous`（第 2 案）。7-1。〇〇（スパイクで決める。`.env.example` は `bearer`） |
| `GAS_CALL_TIMEOUT_SECONDS` | — | 既定 `330` |
| `DB_SHEET_TIMEZONE` | — | DB スプレッドシートの時間帯（IANA 名）。〇〇（`Etc/GMT` の見込み。Sheets API の `properties.timeZone` で決める）。シートと違えば行の反映を止める |
| `SHEET_CONFIG_SYNC_MINUTES` | — | 既定 `5` |
| `FORWARDED_ALLOW_IPS` | — | `10.94.0.10` |
| `KYUKA_NET_SUBNET`／`KYUKA_WEB_IP` | — | `10.94.0.0/24`／`10.94.0.10`（compose だけ） |
| `APP_IMAGE_TAG` | — | compose。切り替えの前に確かめる（罠 24e） |
| `TUNNEL_TOKEN` | 秘密 | compose だけ |
| `BACKUPS_HOST_PATH`／`BACKUP_KEEP_DAYS` | — | compose だけ |

- 使わないキー: `PORTAL_URL`（ポータルの URL は Settings の `PORTAL_URL` を画面に出す＝GAS と同じ）、`MAIL_REDIRECT_TO`・`MAIL_ALLOWED_DOMAINS`（VPS はメールを送らない）、`SEBASTIAN_*`、`EXPORT_HOURLY_UNTIL`。
- 委任のスコープ（最初に決め、後から足さない＝罠 7）: `openid email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive`。**Gmail は持たない**（メールは GAS が送る。K-5）。`drive` は ① `GAS_INTERNAL_AUTH=bearer` のときの GAS の呼び出し（7-1）② 切り替えの控えとステージングの準備の道具、だけが使う。api・worker のコードは Drive の API を呼ばない。
- コードに定数で持つ ID は `config.py` だけ: `PROD_SPREADSHEET_ID_FIXED`・`MASTER_SOURCE_SPREADSHEET_ID_FIXED`・`CALENDAR_SOURCE_SPREADSHEET_ID_FIXED`・`PROD_GAS_SCRIPT_ID_FIXED`・`PROD_PDF_FOLDER_ID_FIXED`・`PROD_LEDGER_SPREADSHEET_ID_FIXED`・`PROD_GAS_SCREEN_DEPLOYMENT_IDS`（@106・@104）。ほかのモジュールに ID を直書きしない（`test_config_safety.py` が拒む）。

---

## 2. リポジトリの形と、写すもの・捨てるもの

### 2-1. 置き場所

- 新しいフォルダ `repo/vps-app/`（このリポジトリの直下）。土台は残業の `vps-app/` を写して作る。
- **GAS のファイル（`src/`）は G 担当だけが変える。** このリポジトリに GitHub Actions は無い（00 2 節）が、本番のトリガーは HEAD のコードで動く（罠 61）。`clasp push` は切り替えの合図までしない。G は自分のブランチにコミットし、push しない。
- 担当エージェントは VPS・Cloudflare・Google・clasp・GitHub に触れない（ssh・デプロイ・外部 API・`git push` をしない）。手元のテストだけ。

```
vps-app/
  docker-compose.yml  .env.example
  api/
    Dockerfile  requirements.txt  requirements-dev.txt  pytest.ini  alembic.ini
    alembic/  env.py  versions/0001_initial_schema.py
    app/
      main.py  config.py  db.py  models.py  timeutil.py  jsutil.py  constants.py
      safety.py  auth.py  google_delegate.py  permissions.py  rpc.py  serialize.py
      sheet_queue.py  audit.py  jobs.py  worker.py                        # F
      pages.py                                                            # U
      templates/pages/top.html preview.html sign.html deptAdmin.html somuAdmin.html
                      leaveStatus.html admin.html no_auth.html            # U
      domain/   quirks.py fiscal.py calendar.py requests.py approval.py
                summaries.py masters.py views.py ops.py                   # W
      services/ sheets_api.py gas_client.py write_gate.py                 # F
                sheet_codec.py row_sync.py reconcile.py summary_sheet.py
                config_sync.py master_sync.py                             # D
                gas_jobs.py ops_status.py                                 # M
    tests/
  web/
    nginx.conf                                                            # F
    static/js/gas-shim.js top.js preview.js sign.js deptAdmin.js somuAdmin.js
              leaveStatus.js admin.js                                     # U
    static/css/mobile.css                                                 # U
    dev/port-pages.js mock-server.js shim.test.js pages.test.js           # U
  backup/backup_entrypoint.sh  cloudflare/setup_tunnel.py                 # F
  tools/migrate_from_sheets.py tools/cutover_backup.py tools/restore_tabs_from_copy.py   # D
  tools/spike_gas_auth.py                                                 # M
  tools/qa/acceptance.py tools/qa/compare_views.py                        # C
  docs/RUNBOOK.md                                                         # F
docs/cutover-runbook.md  docs/acceptance-checklist.md                     # C（リポジトリの docs/）
gas-tests/            # G。GAS の関数をスタブで動かす node の試験（src の外なので clasp は送らない）
src/Internal.gs src/Cutover.gs   # G（新）。ほかは 7-6 の決まった行だけ直す
```

### 2-2. 残業から写すもの（F・D・M・U）

| 写すもの | 直すこと |
|---|---|
| `docker-compose.yml`・`api/Dockerfile`・`requirements*.txt` | 名前・接頭辞・サブネット・固定 IP を 1 章に。playwright・chromium・segno を入れない |
| `web/nginx.conf` | 3-7。`client_max_body_size 4m`（サインの PNG） |
| `app/db.py`・`timeutil.py`・`jsutil.py`・`main.py`・`audit.py`・`serialize.py`・`alembic/env.py` | ほぼそのまま。`/api/sebastian` の経路を外す |
| `app/auth.py`・`google_delegate.py` | クッキー名、期限 90 日、`prompt` を付けない、`/auth/me`。スコープは 1-1。`get_access_token` は残す（GAS の呼び出しと道具が使う） |
| `app/rpc.py` | 閉じたリストを 4 章の 34 関数に。冪等キー・Origin 照合・`commit` を rpc が行う形はそのまま。本文の上限は `api_approveLeaveRequest` だけ 3 MB、ほかは 256 KB |
| `app/safety.py`・`services/write_gate.py` | 9 章の許可リスト |
| `app/jobs.py`・`worker.py`・`sheet_queue.py` | そのまま（罠 73 の直しが入っていること）。予定は 10-3 |
| `services/sheets_api.py`・`gas_client.py` | `gas_client` に `GAS_INTERNAL_AUTH` の切り替え（7-1）と action の一覧（7-2）を足す |
| `services/sheet_codec.py`・`row_sync.py`・`reconcile.py`・`config_sync.py`・`master_sync.py` | 対象のタブと列の決まりを 6 章に |
| `services/gas_jobs.py`・`ops_status.py` | 仕事の種類を 7-2・7-4 に。朝バッチ・Sebastian の待ちを外す |
| `web/static/js/gas-shim.js`・`web/dev/port-pages.js`・`mock-server.js`・`shim.test.js` | 下書きの退避の対象（3-5）、on 属性の置き換えの例外（3-6） |
| `tools/migrate_from_sheets.py`・`cutover_backup.py`・`restore_tabs_from_copy.py` | 6-6 |
| `src/Internal.gs`・`src/Cutover.gs` | 7-1〜7-6（このアプリには `doPost` が無いので新しく足す） |

### 2-3. 捨てるもの

残業の業務（`domain/` の残業・休日・休憩・WorkLogs・Sebastian・二次承認一覧・スタッフ集計）、`services/sheet_intake.py`・`morning_batch.py`、`/api/sebastian`、BatchLogs・WorkLogs・Requests の表と決まり、旧フォームの処理、残業の画面 5 つ、`tools/resend_after_cutover.py`。

### 2-4. 段階 3 の完了条件（全担当）

次の grep が、13 章に理由を書いたものを除いて 0 件。

```sh
grep -rnE "overtime|Overtime|OVERTIME|1GPNBR9U|1Tkl|1Bs1FvgR|\bRequests\b|WorkLogs|BatchLogs|requestId|approve2|staffSummary|sebastian|Sebastian|sheet_intake|morning_batch|worklog|breaks|10\.93\.|supplies|chozouhin|\bparts\b|1VKbB13|kiosk" vps-app/api vps-app/tools vps-app/web/nginx.conf vps-app/web/static/js/gas-shim.js vps-app/docker-compose.yml vps-app/.env.example gas-tests/ src/Internal.gs src/Cutover.gs --exclude-dir=.venv --exclude-dir=node_modules --exclude-dir=__pycache__
grep -rnE "残業|貯蔵品|工番" vps-app/api/app --include=*.py
```

`16v1tQNC`・`1iu5Hoak`・`1Knx_kaQ`・`1B7tnsT4`・`1MQibR08`・`15aGUL5i`・`AKfycbx7sB9x`・`AKfycbwHI1uz`・`imaizumi@` はこのアプリの正しい値なので、`config.py`・`.env.example`・`docs/RUNBOOK.md`・`src/Cutover.gs`（持ち主の確かめ）にだけ現れてよい。画面のテンプレート（`templates/pages/*`・`web/static/js/<page>.js`）は GAS の HTML の写しなので、日本語の語の grep の対象から外す（`-- 休日出勤日を選択` などの文言が正しく入っている）。

---

## 3. 画面（7 ページ＋権限エラー）

### 3-1. 方針

- `src/top.html`・`preview.html`・`sign.html`・`deptAdmin.html`・`somuAdmin.html`・`leaveStatus.html`・`admin.html`・`no_auth.html` を、**HTML・CSS・文言・JS の処理をそのまま**にして api のテンプレート（Jinja2、自動エスケープ有効）で返す。PC での見た目は 1 px も変えない（K-7）。
- 変えてよいのは次だけ（必須の直しを含む）: テンプレートの差し込み（3-3）、`google.script.run` のシム（3-4）、401 のときの書きかけの退避（3-5）、CSP のための JS の置き場所と on 属性の置き換え（3-6）、XSS の直し（3-8。必須）、スマホの CSS（3-9）、総務管理の運用の欄（3-11）。
- 「今日」をブラウザの時計で決める箇所（inv-2 8 節）はそのまま。

### 3-2. 入口 `GET /`（`app/pages.py`、U）

| 項目 | 決まり（`menu.gs:3-89` を写す） |
|---|---|
| ログイン | 全ページで必須（K-0）。未ログインは `302 /auth/login?next=<今のパス＋クエリ>`（`?page=…&reqId=…` を保つ）。`next` は `auth.safe_next` を通す |
| `page` | 無指定・一覧に無い値は `top`。一覧は `top`・`preview`・`sign`・`deptAdmin`・`somuAdmin`・`leaveStatus`・`admin`。大文字小文字を区別する（GAS と同じ） |
| `admin` | `permissions.can_access_admin(db, email)` が偽なら `no_auth`。`message`＝`このページは管理者・総務のみアクセスできます。`、題名 `アクセス権限エラー` |
| `deptAdmin` | `is_admin` か `is_dept_approver` が真なら表示。偽なら `no_auth`、`このページは部署承認者または管理者のみアクセスできます。`、題名 `アクセス権限なし` |
| `somuAdmin` | `is_somu` か `is_admin` が真なら表示。偽なら `no_auth`、`このページは総務部・管理者のみアクセスできます。`、題名 `アクセス権限なし` |
| `top`・`preview`・`sign`・`leaveStatus` | 判定なし（誰でも） |
| `reqId` | `preview`・`sign` だけが使う。URL の値をそのまま（形を確かめない＝GAS と同じ）。JS へは 3-6 の JSON の塊でだけ渡す |
| `approved` | 読まない（GAS も読まない） |
| 題名 | `<title>` を GAS の `setTitle` の値にする: `休暇届 申請TOP`・`休暇届 プレビュー`・`休暇届 サイン`・`休暇届 部署管理`・`休暇届 総務管理`・`休暇届 取得状況`・`休暇届 管理`。leaveStatus の HTML にある `<title>` も同じ値に置き換える |
| 応答ヘッダー | `Cache-Control: no-store`。CSP などは nginx（3-7） |

### 3-3. テンプレート変数（inv-2 2 節の 28 か所）

GAS の `<?= X ?>` は、HTML の属性・本文なら `{{ X }}`、JS の文字の中なら 3-6 の JSON の塊へ移す。`<?!= ?>` は無い。

| 変数 | 値 | 使い方 |
|---|---|---|
| `APP_URL` | `PUBLIC_ORIGIN + "/"`（例 `https://leave.lineworks-local.net/`）。リンクは `{{ APP_URL }}?page=top` のまま。Settings の `APP_URL`（GAS の `/exec`）は**使わない**（統合業務メールのリンク用に GAS のまま残す。01 2-2 の 6） | 全ページのヘッダー、no_auth の「トップ画面へ戻る」、JS の `APP_URL` |
| `PORTAL_URL` | DB の `settings` の `PORTAL_URL` を `normalize_gas` した文字。空ならページの今の処理どおりボタンを出さない | top・somuAdmin・deptAdmin・admin のヘッダー、JS の `PORTAL_URL` |
| `REQ_ID` | 3-2 | preview・sign の JS |
| `message` | 3-2 の文 | no_auth |
| `SHOW_OPS` | `permissions.is_operator(db, email)`（8-9） | somuAdmin の運用の欄（3-11） |

### 3-4. `google.script.run` の代わり（`web/static/js/gas-shim.js`、U）

残業の `gas-shim.js` を写す。各ページの呼び出しの行（37 か所。inv-2 4 節）は 1 字も変えずに動くこと。

- 形: `google.script.run.withSuccessHandler(f).withFailureHandler(g).<fn>(...args)`。片方だけ・どちらも無し・順が逆、も動く。呼ぶたびに新しい runner を返す。
- 送り方: `fetch('/api/rpc/' + fn, {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(args), credentials:'same-origin'})`。**省略された引数は送らない**。
- 返事: 200 `{"ok":true,"result":R}` → `f(R)`（`R` が `null`・`{ok:false,error}`・`[]` でもそのまま成功ハンドラへ。`api_submitLeaveRequest` の `{ok:false}` を失敗にしない）。400 `{"ok":false,"error":msg}` → `g(err)`、`err` は `new Error(msg)` で `toString` を `msg` に上書き（`err.message` と `(err.message || err)` の両方で文が出る）。409 → 1 秒おきに最大 10 回、同じ冪等キーで送り直す。401 → 3-5。5xx・通信エラー → `g(Error('サーバーでエラーが起きました。時間をおいてもう一度お試しください。'))`。
- 冪等キー: 4-3 の「書く」関数と `ops_*` の書く関数にだけ `X-Idempotency-Key`。キーは「関数名＋引数の JSON の SHA-256」ごとに 1 つ、200・400 で捨てる。同じ関数・同じ引数の呼び出しが返事待ちの間に来たら、送らずに前の返事を両方のハンドラに渡す（二度押しで 2 件できない）。
- `failure なし` の呼び出し（inv-2 4 節で「失敗なし」の 12 か所）は、失敗しても何もしない（GAS と同じく「読込中」のまま）。

### 3-5. 操作中のログイン切れ（罠 20、R-15）

- 401 を受けたら、失敗ハンドラは呼ばない。
- 書きかけを `sessionStorage` に退避（キー `kyuka-draft:<page>[:<reqId>]`、24 時間、同じメールのときだけ戻す）。対象:
  - top の申請モーダル（部署・作業員・区分・休暇日（選択肢と過去申請の選択）・半日区分・種類・振替元・特別理由・有給の理由・今の手順）と編集モーダル（対象の reqId・休暇日・区分・半日区分）。
  - **sign の承認者の選択と、手書きサインの画（`canvas.toDataURL('image/png')`）**。戻したら、手書きを選んだ状態にして `resizeCanvas()` の後に画像を枠いっぱいに描き、`hasDrawn=true` にする（サインを捨てない。必須）。
  - somuAdmin の選択中のチェック。
- 手順: `beforeunload` があれば外し、`alert('ログインが切れました。ログイン後にもう一度押してください。')`、`location.href = '/auth/login?next=' + encodeURIComponent(location.pathname + location.search)`。書き込みは自動で送り直さない。冪等キーは捨てない。
- 先に確かめる: top の申請モーダルを開くとき、sign を開いたとき・「承認する」を押す前、`visibilitychange` で画面に戻ったときに `GET /auth/me` を呼び、401 なら上と同じ手順（サインを書き始める前に気づかせる）。

### 3-6. CSP と JS の置き場所

残業 3-6 の案 1（`script-src 'self'`。インラインの script と on 属性を禁止）を写す。

- 各ページの `<script>` の中身を `web/static/js/<page>.js` に移す（中身は書き換えない。関数はグローバルのまま）。読み込みは `<script src="/static/js/gas-shim.js"></script><script src="/static/js/<page>.js"></script>` を今の script の位置に置く。
- JS の中の `<?= ?>`（12 か所）は `<script type="application/json" id="gas-template-data">{{ data_json }}</script>` で渡す。`data_json` は `json.dumps` の後に `<`・`>`・`&` を `\u003c`・`\u003e`・`\u0026` に置き換えた文字（`pages.json_for_script`）。`<page>.js` の先頭で `var __T = JSON.parse(document.getElementById('gas-template-data').textContent);`、今の変数名で受ける（`var REQ_ID = __T.REQ_ID; var APP_URL = __T.APP_URL; var PORTAL_URL = __T.PORTAL_URL;`）。
- on 属性 **42 個**（HTML 37・JS の文字 5。inv-2 5 節）を `data-on-click`・`data-on-change`＋`data-args='<JSON 配列>'` に置き換え、`gas-shim.js` の 1 つの受け口（`document` への委譲）で呼ぶ。`this` は `"$this"`。JS の文字の 5 個のうち top の 3 個（787・788・972）は `port-pages.js` の決まった形、somuAdmin の 259（`onchange="updateSelectedCount()"`）と 266（`singleApprove2(' + i + ', this)` → `data-args` に `[i, "$this"]`）は個別に置き換える。deptAdmin の `.onclick = 関数`（246・256）は属性ではないのでそのまま。
- `style` 属性と `<style>` は書き換えない（`style-src 'unsafe-inline'`）。
- sign の承認後の `<meta http-equiv="refresh">` を JS で足す動き（sign:264-267）はそのまま（CSP は止めない。A-14 で確かめる）。

### 3-7. nginx（F）

- `location /static/` だけ nginx が返す（`expires -1`）。`/`・`/api/`・`/auth/`・`= /healthz` は api へ。
- CSP: `default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'self'; base-uri 'self'; form-action 'self'`（`img-src data:` は leaveStatus の選択肢の矢印と sign の下書きの復元。ポータルは `same_tab` で開くので iframe は無い＝06 2-3）。HSTS・nosniff・Referrer-Policy・Permissions-Policy は残業と同じ、server レベルに 1 回だけ。
- `client_max_body_size 4m`。`limit_req` は残業と同じ値。利用者の IP は `CF-Connecting-IP`。

### 3-8. XSS の直し（必須。inv-2 11 節・inv-1 S9 の全箇所）

- 逃がす関数: 各ページの今の `escHtml`／`esc` に `'` → `&#39;` を足す（`if (!s) return ''` の癖は残す）。**数や 0 になりうる値には今の `escHtml` を使わない**（`escHtml(0)` が空になり表示が変わる）。新しく `__escText(v)`（`v == null` なら `''`、それ以外は `String(v)` を `& < > " '` で逃がす）を `gas-shim.js` に置き、数の値に使う。
- 直す箇所（逃がす。見た目は同じ）:

| 画面:行 | 値 |
|---|---|
| top:760-764・794-800 | 本日の休暇者・今後の予定の `deptName・workerName・leaveKubun・halfDayType・leaveType・status・leaveDate` |
| top:787-791・1173-1175・1233 | `reqId`（`data-args` の JSON と `href` の `encodeURIComponent`） |
| top:1210-1233 | 承認待ちの全キー（`daysAgo` は `__escText`） |
| top:613-619・647・651 | 振替の検証の枠・選択肢の日付と期間 |
| sign:180-181 | `err.message`（`textContent` か `escHtml`） |
| somuAdmin:255・471・477-480 | `daysAgo`・回数（`__escText`）・`lastPaidDate` |
| somuAdmin:274・399 | `[data-reqid="…"]` の `reqId` は `CSS.escape` を通す |
| admin:207・214・252-254・260-264 | `err.message`・`warnLevel`・回数（`__escText`）・`lastPaidDate` |
| leaveStatus:320・333-336・379・381-384・417-420・435 | 数は `__escText`、`m.state` を class に入れる所は `done・risk・warn・note・future` の一覧に無ければ空 |
| deptAdmin・preview | 既に逃がしている。`'` の追加だけ |

- 試験（12 章 U）: 氏名が `x');alert(1)//` と `"><img src=x onerror=alert(1)>` の作業員で申請を作り、top・preview・somuAdmin・deptAdmin・admin・leaveStatus で何も実行されず、文字のまま見える。

### 3-9. スマホ・タブレット（K-7。PC の見た目は変えない）

- 足す CSS は `web/static/css/mobile.css` の 1 ファイルに集め、各ページの最後の `<link>` で読む。**すべての規則を `@media (max-width: 820px)` 以下の中に書く**（`test_web_config.py` が確かめる）。
- 必ず入れる: top の表（7 列）と admin の表を包む要素に `overflow-x:auto`（inv-2 10 節）。`input[type=date]` の `-webkit-appearance:none; appearance:none; display:block; min-width:0; max-width:100%; box-sizing:border-box; text-align:left` と `::-webkit-date-and-time-value { text-align:left }`（top の編集モーダル）。押せる大きさ 40px 以上。モーダル（540px・420px・480px）は画面幅に収める。
- sign の viewport（`maximum-scale=1, user-scalable=no`）と canvas の大きさの決め方は変えない。リサイズで消える癖は今のまま（15 章 Q14）。
- Safari の日付: preview の `new Date('yyyy-MM-dd HH:mm:ss')`（空白区切り）は古い Safari で `Invalid Date` になるおそれ（inv-2 8 節）。**JS は変えない**。代わりに iPhone の実機で A-12 を確かめ、出なければ 15 章で聞く。U は 7 ページの `new Date(` を全部見て、文字を渡している箇所の一覧を報告に書く。

### 3-10. サイン（sign、U。今のまま移す）

- canvas・pointer イベント・`devicePixelRatio`・クリア・`toDataURL('image/png')`（背景は透明）・承認後の 1 秒の meta refresh は、inv-2 9 節のとおり変えない。
- 送る値は今のまま: 電子印なら `api_approveLeaveRequest(REQ_ID, '', email)`、手書きなら `api_approveLeaveRequest(REQ_ID, dataUrl, '')`。
- 3-5 の退避と復元だけを足す。

### 3-11. 運用の欄（管理者だけ。U が somuAdmin.html の末尾に足す）

somuAdmin の一番下に、`SHOW_OPS` が真のときだけ折りたたみの「運用（管理者）」欄を出す。今ある部分の見た目は変えない。中身は `ops_getStatus()` の結果（4-2）と次のボタン: ［設定タブを今すぐ取り込む］［マスタ同期］［集計タブを今すぐ書き直す］［失敗した仕事の再実行］（行ごと）［有給警告メールを送る］［委任の許可をやり直す］（`/auth/google-delegate/start` へのリンク）［セッションの取り消し（メール）］。押す前に confirm を出す。有給警告メールの confirm は `有給警告メールを送信しますか？\n（警告送信日（7/1・9/1・1/1・2/16）以外は送信されません）`。

---

## 4. RPC

### 4-1. 取り決め（F が作り、W が関数を登録する）

- `POST /api/rpc/{fn}`、本文は引数の JSON 配列。Origin が `PUBLIC_ORIGIN`（無ければ `Sec-Fetch-Site: same-origin`）、`Content-Type: application/json`、`fn` が閉じたリストにある、本文が配列、を入口で確かめる。
- 応答: 200 `{"ok":true,"result":…}`／業務エラー 400 `{"ok":false,"error":"GAS と同じ文"}`／401 `{"ok":false,"error":"ログインが必要です。","login":"/auth/login"}`／403（`ops_*` を運用者以外が呼んだ）`{"ok":false,"error":"管理者権限がありません。"}`／409（冪等キーが処理中）／413（本文が大きすぎる）／501（未登録）／500（詳細はログだけ）。
- **GAS で `throw new Error(文)` だった所は 400 の `RpcError(文)`、`return {ok:false, error}` だった所は 200 の戻り値**（申請の入力エラーだけ。inv-1 1 節）。権限の拒否も GAS では例外なので 400（403 ではない）。
- 業務関数の形: `def api_x(ctx: RpcContext, *args)`。引数は GAS と同じ並びの位置引数。省略できる引数は既定値を持つ（GAS は足りない引数を `undefined` とする。足りない引数はすべて `None` として受ける）。`ctx.user.email`（小文字）・`ctx.db`・`ctx.now`（1 回の呼び出しで同じ時刻、UTC）・`ctx.settings`。**業務関数は commit しない**。
- 書き込みの関数（「書く」）は `audit.record(...)` を同じトランザクションで残す（F が rpc の中で呼ぶ。`request_ids` は引数の reqId か戻り値の `reqId`）。
- 返す JSON の作り（GAS の直列化の癖を写す）: 日付・日時はすべて JST の文字（表の形）。GAS で `undefined` のキーは出さない。数は GAS と同じ（有給回数は `0.5` 刻みの float、ほかは int）。キーの並び順は気にしない。
- **シートの型の扱い（必須）**: GAS は「セルが Date か」で分けて返す（`instanceof Date`）。VPS は行の `cell_kinds`（5 章）で同じに分ける。`kind` が `serial`（日付・日時のセル）なら Date、`string` なら文字のセル（GAS が `String(値)` を返す所は DB の `*_text` を返す）、空なら空。新しい行・VPS が書いた値は `serial`。以下の表の「Date なら」はこの意味。
- 閉じたリスト `RPC_FUNCTIONS` = 4-3 の 27 関数＋4-2 の 7 関数（計 34）。出さないもの: `api_canAccessAdmin`・`api_getDeptApproverDepts`・`api_approveLeaveRequest2Batch`（画面は呼ばない）、`setupAllSheets`・`setupAllTriggers`・`syncAllMasters`・`syncCalendarMaster`・`syncStampMaster`・`syncDeptApprovers`・`rebuildLeaveSummary`・`checkAndSendWarnMails`・`debug_checkLeaveRequestSheet`・`test_generatePdf`・`test_stampLookup`・`migrateLeaveRequestHeaders_`（K-0・inv-1 S5。保守は 4-2 で行う＝既定。オーナーに確認中）。

### 4-2. 運用者だけの関数（W が `domain/ops.py` に作る。すべて `permissions.is_operator` が真のときだけ。偽なら 403）

| 関数（引数） | 返す形 | 中身 |
|---|---|---|
| `ops_getStatus()` | `ops_status.get_ops_status(db)` の dict（M） | 行の反映の待ち件数・最も古い待ちの経過秒・止まっている行、集計タブの最後の書き直し、GAS の仕事の待ち・実行中・失敗の件数と失敗の一覧（id・REQ_ID・action・理由・回数）、最後の `gas_ping` の時刻と結果（照合の違い）、マスタ同期・設定の取り込み・毎晩の照合の最後の結果 |
| `ops_retryJob(jobId)` | `{ok:true}` | `failed` の仕事を `pending` に戻す（書く）。`jobId` は bool でない整数か数字だけの文字 |
| `ops_syncSheetConfigNow()` | `config_sync.sync_sheet_config(db)` の結果 | 設定とマスタのタブを今すぐ取り込む（6-5） |
| `ops_runMasterSync()` | `master_sync.sync_masters(db)` の結果 | マスタ同期（6-4。今のメニュー「マスタ同期（手動実行）」の代わり） |
| `ops_rebuildSummaryNow()` | `summary_sheet.rebuild_now(db)` の結果 `{ok, rows}` | V_LEAVE_SUMMARY を今すぐ書き直す（6-3。メニュー「サマリー再構築」の代わり） |
| `ops_sendWarnMail()` | `{ok:true, jobId}` | `gas_jobs.enqueue_send_warn_mail(db)`（7-2 `internal.sendWarnMail`。メニュー「有給警告メール送信（手動）」の代わり。既定） |
| `ops_revokeSessions(email)` | `{ok:true, revoked:N}` | `auth.revoke_sessions` |

### 4-3. 画面の 27 関数（すべて W。シムと画面は U）

「書く」は冪等キーと監査の対象。「積む」は 7 章の GAS の仕事（M の `gas_jobs.enqueue_*`）と 6 章の行の反映（`sheet_queue.mark_dirty`）と集計タブの書き直し（`summary_sheet.mark_dirty`）。エラー文は GAS の原文で 1 字も変えない。行の探し方は「DB の `req_id`（シートの REQ_ID を `normalize_gas` した値）と引数を、引数を正規化せずに完全一致」（`leaveRequest.gs:287・488・637` と同じ）。

| # | 関数（引数） | 返す形 | 権限 | 書く・積む | GAS |
|---|---|---|---|---|---|
| 1 | `api_isDeptApprover()` | bool | なし | 読む | `auth.gs:116` |
| 2 | `api_isAdmin()` | bool | なし | 読む | `auth.gs:108` |
| 3 | `api_isSomu()` | bool | なし | 読む | `auth.gs:112` |
| 4 | `api_getDeptList()` | `[{deptId, deptName}]`（同じ値＝部署名。4-4-a） | なし | 読む | `masters.gs:190` |
| 5 | `api_getWorkerInfo()` | `null` か `{workerId, workerName, deptId, deptName, email}`（4-4-a。本番の作業員マスタに「Googleアカウント」の列が無いので、いつも `null`） | なし | 読む | `masters.gs:77` |
| 6 | `api_getLookupData()` | `{leaveKubun:['全日休','半日休'], halfDayType:['午前','午後'], leaveTypeFull:['有給消化（私用）','振替休暇','特別休暇'], leaveTypeHalf:['有給消化（私用）','特別休暇'], specialReasons:[…]}`。`specialReasons` は M_LOOKUP のカテゴリ `特別休暇理由` の値（`normalize_gas`）を SORT（`Number(値 or 0)`。SORT 列が無ければ行の順）の昇順、同じなら元の順 | なし | 読む | `masters.gs:202`・`config.gs:233` |
| 7 | `api_getWorkingDays()` | `['yyyy-MM-dd', …]`（8-3） | なし | 読む | `masters.gs:419` |
| 8 | `api_getWorkersByDept(deptId)` | `deptId` が偽なら `[]`。それ以外は `[{workerId, workerName, deptId}]`（4-4-a） | なし | 読む | `masters.gs:194` |
| 9 | `api_getUpcomingLeaves(deptId, workerId)` | 4-4-b | なし | 読む | `leaveRequest.gs:356` |
| 10 | `api_getHolidayDays(workerId)` | `[{date:'yyyy-MM-dd', label}]`（8-3） | なし | 読む | `masters.gs:220` |
| 11 | `api_getPastWorkingDays(yearMonth)` | `[{date:'yyyy-MM-dd', dow:0-6, label:'yyyy-MM-dd（曜）'}]`（8-3） | なし | 読む | `masters.gs:480` |
| 12 | `api_submitLeaveRequest(data)` | 成功 `{ok:true, reqId:'LV-yyyyMMdd-NNN'}`、入力エラー `{ok:false, error}` を **200 で**（8-5） | なし（誰の名義でも） | **書く**: `leave_requests` 1 行（新しい行）。積む: 行の反映だけ（集計・GAS の仕事は積まない） | `leaveRequest.gs:131` |
| 13 | `api_getAllPendingRequests()` | 4-4-c | なし | 読む | `leaveRequest.gs:419` |
| 14 | `api_editLeaveRequest(reqId, patch)` | `{ok:true, newStatus:'申請中'｜元の状態}`。エラーは 400（8-6） | なし（誰でも、どの申請でも＝今のまま） | **書く**: 8-6。承認済みなら積む: `internal.afterUnapprove`（`reason:'edit'`）。いつも積む: 行の反映・集計 | `leaveRequest.gs:852` |
| 15 | `api_deleteLeaveRequest(reqId)` | `{ok:true}`。`申請が見つかりません。`・`既に取消済みです。` は 400 | なし（今のまま） | **書く**: 状態＝`取消`（ほかの列は残す）。積む: `internal.afterUnapprove`（`reason:'cancel'`）・行の反映・集計 | `leaveRequest.gs:970` |
| 16 | `api_getLeaveRequestById(reqId)` | `null` か 4-4-d。**見つからなければすぐ `null`**（GAS の約 6 秒の待ちをしない＝変えるもの） | なし | 読む | `leaveRequest.gs:239` |
| 17 | `api_getApproverList()` | `[{email, name}]`: `stamps` の行の順に、A 列を `trim`（`normalize_` しない）して空でない行の `{email, name: C 列の trim}`。例外のときは `[]` | なし | 読む | `masters.gs:163` |
| 18 | `api_approveLeaveRequest(reqId, signBase64, approverEmail)` | `{ok:true, reqId}`（`reqId` は引数のまま） | なし（今のまま） | **書く**: 8-7。積む: `internal.afterApprove1`・行の反映・集計 | `leaveRequest.gs:471` |
| 19 | `api_getCurrentApproverInfo()` | `{email, name}`: `email`＝ログインのメール（小文字）、`name`＝`stamps` で A 列を `normalize_gas`＋小文字にして一致した最初の行の C 列（`normalize_gas`）、空か無ければ `email` | なし | 読む | `auth.gs:187` |
| 20 | `api_getPending2ndApprovals()` | 4-4-e | 総務か管理者。違えば 400 `権限がありません（総務または管理者のみ）` | 読む | `leaveRequest.gs:565` |
| 21 | `api_approveLeaveRequest2(reqId, approver2Email)` | `{ok:true, reqId}` | 先に `reqId` が空なら `REQ_IDが指定されていません。`、次に総務か管理者（違えば `権限がありません（総務または管理者のみ）`）。`approver2Email` がログインの人かは見ない（今のまま。Q2） | **書く**: 8-8。積む: `internal.afterApprove2`・行の反映 | `leaveRequest.gs:619` |
| 22 | `api_getSomuAdminView(fyYear, deptFilter)` | `[WorkerView]`（4-4-f） | 総務か管理者。違えば `権限がありません（総務または管理者のみ）` | 読む | `leaveSummary.gs:455` |
| 23 | `api_getApproverProfiles()` | `{isAdmin, currentEmail, profiles:[{email, name, depts:[…]}]}`（4-4-g） | なし（管理者は全員分、ほかは自分の分だけ） | 読む | `auth.gs:130` |
| 24 | `api_getDeptAdminView(fyYear, selectedDepts)` | `{workers:[WorkerView], deptNames:[…]}`（4-4-f） | 管理者か部署承認者。違えば `部署承認者の権限がありません。` | 読む | `leaveSummary.gs:428` |
| 25 | `api_getAdminSomuView(fyYear, deptFilter)` | `[SummaryView]`（4-4-h） | `ALLOW_ADMIN_PAGE_EMAILS`。違えば `権限がありません` | 読む | `leaveSummary.gs:203` |
| 26 | `api_getAdminBossView(deptId, fyYear)` | `[SummaryView]`（4-4-h） | 同上 | 読む | `leaveSummary.gs:155` |
| 27 | `api_getLeaveStatusSearch(fyYear, deptId, workerId)` | 4-4-i | なし（誰でも・どの人でも＝今のまま） | 読む | `leaveStatusApi.gs:12` |

**GAS が「画面を開くたびに集計タブを書き直す」動き（22・24〜27）は写さない。** VPS はシートを書かずに DB から同じ計算で作って返す（6-3。変えるもの）。

### 4-4. 返す形の細かい決まり

**a. 部署と作業員（`masters.gs:8-114`）**

- 作業員の表（`workers`）を 1 行目の見出しで読む: `作業員コード`（無ければ `作業員ID`）・`氏名`・`部署`（無ければ `部署ID`）・`在籍フラグ`・「`Googleアカウント` を含む」列。コード・氏名・部署のどれかが無ければ 400 `作業員マスタのヘッダが想定と違います（作業員コード/氏名/部署）。実際: ` ＋見出しを `, ` でつないだ文字。
- 行は部署・コード・氏名を `normalize_gas` し、どれかが空なら捨てる。`在籍フラグ` の列があれば、空・`1`・`true`（小文字にして）以外の行を捨てる。
- 部署の一覧: 部署の重なりを除き、JS の既定の `sort()`（UTF-16 の符号順＝`jsutil.js_sort_key`）。
- 部署ごとの作業員: 元の行の順に入れ、`workerId.localeCompare(…)`（言語の指定なし＝ICU の既定の照合。8-10）で並べる。`api_getWorkersByDept` は引数を正規化せずに部署名と完全一致。
- `api_getWorkerInfo`: ログインのメールが空なら `null`。「Googleアカウント」を含む列が無ければ `null`。在籍の条件を通った行で、その列を `normalize_gas`＋小文字にしてログインのメールと一致した最初の行。

**b. `api_getUpcomingLeaves(deptId, workerId)`**

- 対象: `req_id` が空でない、状態が `取消` でない（`却下` は出す）、`deptId` が真なら部署 ID が完全一致、`workerId` が真なら作業員 ID が完全一致（行の値は `normalize_gas`、引数はそのまま）。休暇日が Date で JST の今日より前なら除く。**休暇日が Date でない行は除かずに載せる**。
- 1 件: `{reqId, deptName, workerName, leaveKubun, halfDayType, leaveDate, leaveType, status, isRetroactive, daysAgo}`。`leaveDate` は Date なら `yyyy/MM/dd`、ほかは `''`。`isRetroactive`・`daysAgo` は 8-2 の式（休暇日が Date でなければ `false`・`0`）。
- 並び: `leaveDate` の `localeCompare`（言語の指定なし。8-10）の昇順。

**c. `api_getAllPendingRequests()`**

- 対象: `req_id` が空でなく、状態が `申請中`。
- 1 件: `{reqId, deptName, workerName, leaveKubun, halfDayType, leaveDate, leaveType, submittedAt, status, isRetroactive, daysAgo}`。`leaveDate` は Date なら `yyyy/MM/dd`、ほかは `''`。`submittedAt` は申請日時が Date なら `MM/dd HH:mm`、ほかは `''`。
- 並び: `submittedAt` の `localeCompare` の**降順**（年をまたぐと 12 月が 1 月より上に来る癖は今のまま＝既定。Q8）。

**d. `api_getLeaveRequestById(reqId)`**（preview・GAS の内部と同じ形）

`reqId` が偽なら `null`。見つかれば:

| キー | 値 |
|---|---|
| `reqId` | 行の REQ_ID |
| `fiscalYear` | 年度のセルを JS の `String(値 \|\| '')` にした文字（数の `2026` は `'2026'`、`0`・空は `''`） |
| `deptId`・`deptName`・`workerId`・`workerName`・`leaveKubun`・`halfDayType`・`leaveType`・`specialReason`・`paidDetail`・`additionalDetail`・`status`・`signImageUrl`・`pdfUrl`・`approvedBy1Email`・`approvedBy2`・`approvedBy2Email` | 各列を `normalize_gas`（`additionalDetail`＝`DETAIL_FREE`、`specialReason`＝`SPECIAL_REASON_TEXT`） |
| `leaveDate`・`substituteDate` | Date なら `yyyy-MM-dd`、ほかは元の文字（空は `''`） |
| `submittedAt`・`approvedAt` | Date なら `yyyy-MM-dd HH:mm:ss`、ほかは元の文字 |
| `approvedAt2` | Date なら `yyyy-MM-dd HH:mm:ss`、**ほかは `''`**（文字のセルでも `''`） |
| `rowNo` | `sheet_row`（まだシートに書いていない新しい行は `null`。画面は使わない） |

`isRetroactive`・`daysAgo` は返さない（preview の過去申請のバッジは今も出ない＝今のまま）。

**e. `api_getPending2ndApprovals()`**

- 対象: `req_id` が空でなく、状態が `承認済` で、2 次承認日時が空（セルが空、か文字を `trim` して空）。
- 1 件: `{reqId, deptName, workerName, leaveDate, leaveKubun, leaveType, approvedAt, isRetroactive, daysAgo}`（`halfDayType` は**入れない**）。`leaveDate` は Date なら `yyyy/MM/dd`、ほかは元の文字。`approvedAt` は Date なら `yyyy/MM/dd HH:mm`、ほかは `''`。
- 並び: `leaveDate` の `localeCompare` の昇順。

**f. `WorkerView`（`buildAdminViewData_`、`leaveSummary.gs:284-419`）**

- 22: `fyYear` が偽なら JST の今日の年度。`deptFilter` が偽か、`normalize_gas` して `ALL` か空なら全部署（4-4-a の部署の一覧）、それ以外は `[deptFilter]`。
- 24: 管理者で `selectedDepts` が 1 件以上ならそれ。管理者でなく 1 件以上なら、自分の担当部署（8-9）に含まれるものだけ。0 件なら自分の担当部署。結果が空なら `{workers:[], deptNames:[]}`。年度は `fyYear || 今日の年度`。`deptNames` は使った部署の配列。
- 作り方（年度 `fy = Number(fyYear)`）:
  1. 集計（6-3 の `summary_rows`）のうち `Number(年度) === fy` の行を作業員 ID で引く表（同じ作業員が 2 行あればキーの順で後の行が勝つ）。
  2. `leave_requests` の `req_id` が空でなく、状態 `承認済`、`Number(年度) === fy`、種類 `有給消化（私用）`、休暇日が Date の行の休暇日を、作業員 ID ごとに集める。
  3. 作業員の表を部署ごとにまとめた順（部署が最初に出た順、部署の中は 4-4-a の並び）で回し、対象の部署（`normalize_gas` で比べる）の人ごとに 1 件: `{deptName, workerId, workerName, paidCount, paidHalfCount, substituteCount, specialCount, totalCount, lastPaidDate, warnLevel, quarters, currentShortage}`。集計に無い人は回数 0・`lastPaidDate:''`。`warnLevel` は `computeWarnLevel_(paidCount, fy)` で計算し直す（8-4）。`quarters` は 8-4 の 5 件。`currentShortage` は今か過去の期間の `shortage` の最大。
  4. 並び: `paidCount` の昇順 → `currentShortage` の降順 → `deptName` の `localeCompare(…,'ja')` → `workerName` の同じ比べ（安定な並べ替え）。
- `lastPaidDate` は「GAS が集計タブのセルを読み直した文字」と同じにする（D7・K5）。6-3 の `last_paid_date_display()`（既定 `quirks.LAST_PAID_DATE_STYLE = 'gas'`。文字の形はスパイクで決める）。

**g. `api_getApproverProfiles()`**: `dept_approvers` を 2 行目から、A（部署）・B（メールのカンマ区切り）・C（名前）を `normalize_gas`（部署はさらに `trim`）。部署かメールが空の行は飛ばす。メールを `trim`＋小文字に分け、メールごとに最初の行で `{email, name: 名前 || email, depts:[]}` を作り、部署を重ならないように足す。名前がメールのままで後の行に名前があれば上書き。メールが最初に出た順に並べた後、`name.localeCompare(…,'ja')` で並べる。管理者なら `{isAdmin:true, currentEmail, profiles}`、ほかは `{isAdmin:false, currentEmail, profiles: 自分があれば [自分] なければ []}`。

**h. `SummaryView`（`api_getAdminBossView`、`leaveSummary.gs:155-198`）**

- 26: `fyYear` が偽なら今日の年度。25 は `api_getAdminBossView(deptFilter || '', fyYear)` と同じ。
- 集計の行（6-3）のうち `Number(年度) === Number(fyYear)`、`deptId` が真なら部署 ID が完全一致、の行ごとに `{deptId, deptName, workerId, workerName, paidCount, paidHalfCount, substituteCount, specialCount, totalCount, lastPaidDate, warnLevel}`。回数は `Number(値 || 0)`、`warnLevel` は集計の行の値（書き直しの時点で計算した値＝GAS が開くたびに書き直して読んだ値と同じ）、`lastPaidDate` は 4-4-f と同じ。**承認済みが 1 件も無い人は出ない**（今のまま）。
- 並び: `(deptName + workerName).localeCompare(…,'ja')`。

**i. `api_getLeaveStatusSearch(fyYear, deptId, workerId)`**

- `fy = fyYear ? Number(fyYear) : 今日の年度`。`deptId` が真なら `[deptId]`、偽なら全部署。`workers = WorkerView の配列`（4-4-f の作り方）。
- `workerId` が真: `workers` から `workerId` が完全一致の最初の人（`me`）。`{mode:'personal', fyYear: fy, worker: me か null, milestones: me ? 8-4 の 5 件 : [], upcoming: (me かつ fy === 今日の年度) ? api_getUpcomingLeaves(me.deptName, workerId) : []}`。
- 偽: `{mode:'list', fyYear: fy, deptFilter: deptId || '', workers}`。

---

## 5. DB の表（F が `models.py` と `0001_initial_schema` に全部作る）

共通: 文字の列は `NOT NULL DEFAULT ''`。日時は `UTCDateTime`（timestamptz）。業務の表は `row_version int NOT NULL DEFAULT 1`、`sheet_row int NULL`、`cell_kinds jsonb NOT NULL DEFAULT '{}'`（型が混ざる列の行ごとの型。`serial`／`string`／`number`）。**シートの値は GAS が読んだ形で入れ、正規化は業務の側で `jsutil.normalize_gas()` を通す**。F が列名を先に作っていれば、この表に合わせて直す（この表が取り決め）。

### 5-1. `leave_requests`（シート `T_LEAVE_REQUEST`。見出し 1 行目、データ 2 行目から。**見出しの文字で列を引く**）

本番の見出しは英語 26 列（06 2-5）。列の位置は使わない。`HEADER_ALIAS_`（`config.gs:70-95`）の日本語の見出しも同じ列として受ける。

| シートの見出し（別名） | 列 | 型 | 書き手（6-1） |
|---|---|---|---|
| `REQ_ID` | `req_id` | text 主キー | V |
| `FY_START_YEAR`（年度） | `fiscal_year` | text（JS の `String(値)`。`cell_kinds.fiscal_year` が `number`／`string`）。**計算し直さずに移す**（inv-4 0 節 #4） | V |
| `STATUS`（承認状態） | `status` | text（`申請中`・`承認済`・`取消`。`却下` は移行で受けるが書く処理は無い） | V |
| `APPLIED_AT`（申請日時） | `applied_at` | timestamptz NULL＋`applied_at_text` | V |
| `DEPT_ID`・`DEPT_NAME`・`WORKER_ID`・`WORKER_NAME` | `dept_id`・`dept_name`・`worker_id`・`worker_name` | text（`worker_id` は `cell_kinds` で数か文字かを持つ） | V |
| `LEAVE_DATE`（休暇日） | `leave_date` date NULL＋`leave_date_text` | `cell_kinds.leave_date` | V |
| `DAY_TYPE`（休暇区分）・`HALF_TYPE`（半日区分）・`LEAVE_KIND`（休暇種類） | `leave_kubun`・`half_day_type`・`leave_type` | text | V |
| `SUBSTITUTE_FOR_WORK_DATE`（振替元出勤日） | `substitute_date` date NULL＋`substitute_date_text` | `cell_kinds` | V |
| `SPECIAL_REASON_REF` | `special_reason_ref` | text | N |
| `SPECIAL_REASON_TEXT`（特別理由） | `special_reason` | text | V |
| `PAIDLEAVE_DETAIL`（有給詳細） | `paid_detail` | text | V |
| `DETAIL_FREE`（追加詳細） | `additional_detail` | text | N |
| `APPROVED_AT`（承認日時） | `approved_at` | timestamptz NULL＋`approved_at_text` | V |
| `SIGN_IMAGE_URL`（サイン画像URL） | `sign_image_url` | text | **G** |
| `PDF_URL` | `pdf_url` | text | **G** |
| `CREATED_BY_EMAIL`（作成者メール） | `created_by_email` | text | V |
| `UPDATED_AT`（更新日時） | `updated_at_text` | text（全行が空。持つだけ） | N |
| `APPROVED_BY1_EMAIL`（1次承認者メール） | `approved_by1_email` | text | V |
| `APPROVED_BY2`（2次承認者名） | `approved_by2` | text | V |
| `APPROVED_BY2_EMAIL`（2次承認者メール） | `approved_by2_email` | text | V |
| `APPROVED_AT2`（2次承認日時） | `approved_at2` | timestamptz NULL＋`approved_at2_text` | V |
| （シートに無い） | `pdf_file_id` | text。GAS の返事の `pdfFileId`。移行では `PDF_URL` の `/d/<id>/` から取り出す。**本番に `PDF_FILE_ID` 列は無い**（06 2-5）。シートに書かない | VPS だけ |
| （シートに無い） | `sign_png` | bytea NULL。手書きサインの PNG。1 次承認で入れ、`internal.afterApprove1` が済んだら（成功・skipped とも）NULL に戻す。管理者の欄・ログに出さない | VPS だけ |

- 本番の見出しに無い列があれば、移行の dry-run で `missingColumns` に出して止める。シートにあって決まりに無い見出しは `droppedColumns` に出す（触らない）。将来 `PDF_FILE_ID` 列が足されたら G 区分として扱う（GAS が書く）。
- `isRetroactive`・`daysAgo` は列を作らない（毎回求める）。
- 採番の一意: `req_id` の主キー。採番は 8-1。

### 5-2. 設定とマスタの表（シートを正とする。`config_sync` が DB のタブから取り込む＝6-5）

| 表 | DB のタブ | 持ち方 |
|---|---|---|
| `settings` | `M_SYSTEM_SETTING`（1 行目 KEY・VALUE・MEMO、2 行目から A=キー、B=値） | `key`（`trim`）主キー、`value`（生の文字。数・日付のセルは GAS が読む形の文字）、`sheet_row`。**同じキーは下の行が勝ち、値が空でも勝つ**（`config.gs:107-119`）。重なったキーは `sync_runs` に出す |
| `lookup` | `M_LOOKUP`（1 行目 `カテゴリ`・`値`・`SORT`） | `sheet_row` 主キー、`category`・`value`（生）、`sort_text` |
| `workers` | `作業員マスタ` | `sort_order` 主キー、`cells` jsonb（見出し付きの 1 行、GAS が読む形）。4-4-a は見出しの文字で読む |
| `calendar` | `M_CALENDAR`（見出し `日付`・`区分`・`曜日`・`備考`） | `sort_order` 主キー、`date` date NULL（日付のセルを JST の日付にした値。読めない行は NULL）、`kubun`（`normalize_gas`）、`cells` jsonb |
| `stamps` | `M_STAMP`（A メール・B stampFileId・C 備考＝名前。**位置で読む**） | `sort_order` 主キー（2 行目から）、`col_a`・`col_b`・`col_c`（生の文字）、`cells` jsonb |
| `dept_approvers` | `M_DEPT_APPROVERS`（A dept・B approverEmails・C ` approverName`・D 役職・E 通知メール。**位置で読む**） | `sort_order` 主キー（2 行目から）、`col_a`〜`col_e`（生の文字） |

### 5-3. 土台の表（残業と同じ）

`sessions`・`google_credentials`・`idempotency_keys`・`jobs`（`kind`・`lane`・`priority`・`request_id`・`args_json`・`requires_json`・`guard_json`・`state`・`attempts`・`run_after`・`leased_until`・`last_error`・`result_json`・`created_at`・`updated_at`）・`sheet_sync_state`（主キー（`table_name`, `pk`）、`desired_version`・`synced_version`・`synced_values`・`state`・`attempts`・`next_try_at`・`last_error`・`first_dirty_at`・`last_synced_at`・`seq`）・`sync_runs`・`audit_log`。`mail_log` は作らない。`jobs.args_json` に PNG を入れない（PNG は `leave_requests.sign_png`）。

alembic は `0001_initial_schema` 1 本。`test_migrations.py`（SQLite で upgrade → `compare_metadata == []` → downgrade）を通す。

### 5-4. 行ロック

- 書き込みの関数は、申請 1 件ごとに `leave_requests` の行を `SELECT … FOR UPDATE`（`store.lock_request` の 1 か所。SQLite では無視）してから状態を読み直して決める（GAS のスクリプトのロックの代わり。別の申請は待たない）。
- 採番は `pg_advisory_xact_lock(739410001)` の中で最大を数えて入れる（8-1）。
- 振替元の使い回しの確かめ（8-5）は、同じ作業員の行を読む前に `pg_advisory_xact_lock(739410002, hashtext(worker_id))` を取る（2 つの申請が同じ振替元を同時に通らない）。

---

## 6. シートへの書き込み（決定 A。D が作る）

### 6-1. どのタブを書くか・列の区分

| 区分 | タブ・列 | VPS の扱い |
|---|---|---|
| A（VPS が行単位で書く） | `T_LEAVE_REQUEST` の V 列（5-1） | 書くたびに、その行の変わった列だけを同じ位置へ。新しい行は 6-2。行もタブも消さない |
| G（GAS だけが書く） | `T_LEAVE_REQUEST` の `SIGN_IMAGE_URL`・`PDF_URL`（と、もしあれば `PDF_FILE_ID`） | **VPS は書かない（空にするのも GAS）**。行の反映・照合・戻しの道具の対象から外す。GAS の返事の値を DB に入れる（7-3） |
| N（誰も書かない） | `SPECIAL_REASON_REF`・`DETAIL_FREE`・`UPDATED_AT`・決まりに無い列 | 移行で値を持つだけ |
| S（VPS が全行を書き直す） | `V_LEAVE_SUMMARY` | 6-3 |
| M（マスタ同期だけが書く） | `作業員マスタ`・`M_CALENDAR`・`M_STAMP`・`M_DEPT_APPROVERS` | 6-4 |
| B（読むだけ） | `M_SYSTEM_SETTING`・`M_LOOKUP` | 書かない。`config_sync` が数分ごとに読む |
| E（触らない） | `BatchLogs`・雛形由来のタブ（README・CONFIG・LOOKUP・M_DEPT・M_WORKER・PDF_TEMPLATE・PDF_MAP）・ほか | 読みもしない |

### 6-2. T_LEAVE_REQUEST の行単位の反映（`row_sync`。残業の `row_sync.py` を写す）

- 流れ: W が行を書いたら同じトランザクションで `sheet_queue.mark_dirty(db, 'leave_requests', req_id)`。worker の行の反映のレーン（全体で 1 本、2 秒ごと）が最大 50 件取り、DB の今の値と `synced_values` を比べて違う V 列だけを書く。遅れは 1 分以内、ふだんは 10 秒以内（統合業務メールが 17:00・18:00 に読む）。
- 書き方: `spreadsheets.batchUpdate` の `updateCells`（`fields: "userEnteredValue,userEnteredFormat.numberFormat"`）。`formulaValue` は使わない。
- 列ごとの型（`sheet_codec.COLUMN_RULES`。スパイクで本番の型の表を取って直す）:

| 種類 | 列 | 書き方 |
|---|---|---|
| 日付のセル | `LEAVE_DATE`・`SUBSTITUTE_FOR_WORK_DATE` | 値は「その日の JST の 0 時」という時刻（GAS の `new Date('yyyy-MM-ddT00:00:00+09:00')` と同じ瞬間）を `DB_SHEET_TIMEZONE` の壁時計にしたシリアル値（GMT のシートでは前日 15:00 になる。これが今の GAS と同じ値）。行の `cell_kinds` が `string` の古い行は、値を変えるときだけ `serial` に変える（GAS の編集も Date で書く） |
| 日時のセル | `APPLIED_AT`・`APPROVED_AT`・`APPROVED_AT2` | その時刻を `DB_SHEET_TIMEZONE` の壁時計にしたシリアル値（ミリ秒まで） |
| 空にする日時 | 編集で承認が外れたときの `APPROVED_AT`・`APPROVED_AT2` | `userEnteredValue` を持たないセル（GAS の `setValue('')` と同じ空） |
| 数 | `FY_START_YEAR`（`cell_kinds` が `number`） | 整数の `numberValue` |
| 文字（数に見えるものは数） | そのほかの V 列 | GAS の `setValues` と同じく、`^[+-]?(\d+\.?\d*\|\.\d+)$` に合う文字は `numberValue`（DB には `String(Number(値))` を入れ、`cell_kinds` を `number`）、ほかは `stringValue`。日付に見える文字・`=` で始まる文字も文字のセル（GAS は日付・数式に変えていた。受け入れる差＝11 章） |

- 表示形式: 新しく書く日付・日時のセルは、同じ列の最後のデータ行の `numberFormat` を写す（そのセルが文字なら日付は `yyyy-mm-dd`、日時は `yyyy-mm-dd hh:mm`。06 2-5 の表示に合わせる。`APPROVED_AT2` は `M/d/yyyy` の行が多ければそれを写す）。
- **新しい行の位置（必須）**: 「どれかの列に値がある最後の行」の次の行（GAS の `getLastRow()+1` と同じ）。**`values.append` は使わない**（本番には REQ_ID が空で 1 セルだけ値がある行がある＝06 2-5。`values.append` は A 列の切れ目を表の終わりと見て、その行を上書きするおそれ）。外からこのタブに行を足す仕組みは無い（Sebastian も書かない）。書いた位置を `sheet_row` に入れる。行数が足りなければ `appendDimension`（ROWS、末尾だけ）。
- 今ある行は `sheet_row` に書く。書く前に REQ_ID の列を `batchGet` で読んで同じか確かめ、違えば REQ_ID の列を全部読んで探し直す。見つからなければ書かずに `state='blocked'`、管理者の欄に出す。
- 見出しの確かめ: 書く前に 1 行目を読み、`COLUMN_RULES` の見出し（別名を含む）が 1 つでも無ければレーン全体を止める（列を足さない）。
- 時間帯の確かめ: 起動時と 1 時間ごとに `properties.timeZone` が `DB_SHEET_TIMEZONE` と同じか。違えば止める。
- 書き込みの関所（`write_gate`）: production は最新の `gas_ping` が 10 分以内で `migrated=true` のときだけ書く（切り替えの途中と戻した後に GAS と取り合わない）。staging はいつでも書く。
- 失敗のとき: 1 行の失敗は 10 秒・30 秒・1 分・2 分・5 分、以後 5 分ごと。安全装置の拒否・見出しの違い・時間帯の違いはレーンを止める。

### 6-3. V_LEAVE_SUMMARY（`services/summary_sheet.py`、D。計算は W の `domain/summaries.py`）

**数え方は今のまま（`rebuildLeaveSummary_`、`leaveSummary.gs:15-116`）。1 つも変えない。**

- `summaries.summary_rows(db, now) -> list[dict]`（W）: `leave_requests` の `req_id` が空でなく、状態（`normalize_gas`）が `承認済` の行（2 次承認の有無は見ない）を DB の並び（`sheet_row` の昇順、まだ書いていない新しい行は採番の順で後ろ）で回す。キー＝`fiscal_year`（JS の `String(値)`）＋`'_'`＋作業員 ID（`normalize_gas`）。部署 ID・部署名・氏名はそのキーで最初に出た行の値。合計＋1。種類が `有給消化（私用）` なら区分が `半日休` で有給半日＋1・有給＋0.5、ほかは有給＋1、休暇日が Date なら最終有給日を `yyyy-MM-dd` の文字で大きい方に。`振替休暇` は振替＋1、`特別休暇` は特別＋1。並びはキーの JS の既定の `sort()`。警告レベルは `computeWarnLevel_(有給回数, 年度の値, now)`（8-4）。
- 書く（`summary_sheet.write(db)`、D）: 見出し 12 列 `年度・部署ID・部署名・作業員ID・作業員名・有給回数・有給半日回数・振替回数・特別回数・合計回数・最終有給日・警告レベル` と全行を、`values.update`（`A1:L<n+1>`、`valueInputOption=USER_ENTERED`）で書き、その後で余りを `values.batchClear`（`A<n+2>:ZZ` と `M1:ZZ<n+1>`）。年度は元のセルの型（数なら数）、回数は数、作業員 ID は USER_ENTERED（数字だけなら数になる＝GAS の `setValues` と同じ）、最終有給日は `yyyy-MM-dd` の文字を USER_ENTERED（シートが日付に変える＝GAS と同じ）。`=`・`+`・`-`・`@` で始まり数でない文字は先頭に `'` を付ける（数式にしない）。
  - GAS は「全部消してから書く」ので一瞬空になった。VPS は「書いてから余りを消す」ので空の瞬間が無い（受け入れる差。良い方向）。タブが無ければ作らずに止めて知らせる。
  - 歯止め: `leave_requests` が 0 行なら書かない。
- いつ書くか: ① 書き込みの後（1 次承認・編集・取消。GAS が書き直していた場面）に `summary_sheet.mark_dirty(db)` で予約し、worker が最後の予約から 10 秒後（最大 60 秒遅れ）に 1 回書く ② 毎朝 6:50 JST（統合業務メールの 7:00 の前。期限の日付が変わったときの警告レベルを反映）③ `ops_rebuildSummaryNow` ④ `internal.sendWarnMail` の前（7-2）。
- 書き込みの関所: 6-2 と同じ（production は移転モードのときだけ）。
- `last_paid_date_display(value: str, style) -> str`（W）: 画面に返す「最終有給日」の文字。GAS は集計タブのセルを `normalize_(セル)` で読むので、シートが日付に変えたセルは `String(Date)` の長い英語の文字になる見込み（D7・K5。推測）。`quirks.LAST_PAID_DATE_STYLE` が `'gas'`（既定）ならスパイクで本番の画面から写した形（例 `Fri Nov 27 2026 09:00:00 GMT+0900 (日本標準時)` 〇〇）を、`DB_SHEET_TIMEZONE` の 0 時の瞬間を JST で表した文字として作る。`'ymd'` なら `yyyy-MM-dd` のまま。スパイクで本番の画面が `yyyy-MM-dd` と分かれば既定を `'ymd'` にする。

### 6-4. マスタ同期（決定 D。`master_sync.sync_masters`、毎日 6:00 JST と手動）

今の `syncAllMasters`（`syncMasters.gs`）と同じ規則で、元を丸ごと DB のタブに写す。

| 書き先（DB のタブ） | 元 | 元のタブ名 | 見出しの確かめ（違えば止める） |
|---|---|---|---|
| `作業員マスタ` | `MASTER_SOURCE_SPREADSHEET_ID` | Settings `MASTER_SOURCE_WORKER_SHEET`（空なら `作業員マスタ`） | `作業員コード`（か `作業員ID`）・`氏名`・`部署`（か `部署ID`）がある |
| `M_CALENDAR` | `CALENDAR_SOURCE_SPREADSHEET_ID` | Settings `CALENDAR_SOURCE_SHEET`（空なら `社内カレンダーマスタ`） | `日付`・`区分` がある |
| `M_STAMP` | 同上 | `StampMap`（固定） | 1〜3 列目の見出しが今の DB のタブの 1〜3 列目と同じ（位置で読むため。06 2-5 では `メール`・`stampFileId`・`備考`。〇〇 スパイクで 1 字ずつ確かめる） |
| `M_DEPT_APPROVERS` | 同上 | `DeptApprovers`（固定） | 1〜5 列目が今の DB のタブと同じ（`dept`・`approverEmails`・` approverName`・`役職`・`通知メール`） |

- 読む: 1 行目の値がある最後の列まで × 最後の行まで。UNFORMATTED（SERIAL_NUMBER）と FORMATTED の両方。読み取りの失敗は 2 秒おきに 2 回まで再試行（GAS と同じ回数）。
- 1 列目が空（`trim` して空）の行を捨て、見出しは残す（`compactRows_`）。
- 書く: 1 行目から `values.batchUpdate`。日付のセルは元のシリアル値（元と書き先の時間帯が違えば瞬間を保つように直す）を数で書き、元の表示形式を写す。数・真偽はそのまま。文字は USER_ENTERED（GAS の `setValues` と同じ。数式に見える文字は `'` を付ける）。余った行・列は値だけ消す（`values.batchClear`）。**行・列・タブを消したり作り直したりしない**（GAS は余りが 6 行以上なら行を消し、失敗するとタブを作り直していた＝変えるもの）。
- 守り（変えるもの）: 元が読めない・見出しだけ・有効な行 0・見出しの確かめが合わない、なら**そのタブは上書きしない**（GAS は見出しだけに置き換えていた）。タブごとに失敗を握り、ほかのタブは続ける。結果を `sync_runs(kind='sync_masters')` に `{ok, tabs:{名前:{ok, rows, skipped_reason}}}` で残す。
- Settings の `MASTER_SOURCE_SSID`・`CALENDAR_SOURCE_SSID` が `.env` の値と違えば、読むのは `.env` の ID のまま、`sync_runs` と管理者の欄に「食い違い」を出す。
- 書いた後、同じデータから 5-2 の `workers`・`calendar`・`stamps`・`dept_approvers` を全置換する。

### 6-5. 設定とマスタのタブの取り込み（`config_sync.sync_sheet_config`、`SHEET_CONFIG_SYNC_MINUTES` ごとと手動）

- DB スプレッドシートの `M_SYSTEM_SETTING`・`M_LOOKUP`・`作業員マスタ`・`M_CALENDAR`・`M_STAMP`・`M_DEPT_APPROVERS` を読み、5-2 の表を全置換する（GAS は画面を開くたびにタブを読んでいた。人が DB のタブを直した場合も最大 5 分で効く＝変えるもの。［今すぐ取り込む］あり）。
- タブごとに、読めない・空・有効な行 0 なら上書きしない。
- Settings の `PDF_FOLDER_ID`・`LEDGER_SSID`（同じキーは下の行）・`MASTER_SOURCE_SSID`・`CALENDAR_SOURCE_SSID` が `.env` の値と違えば「食い違い」を出す。`LEDGER_SSID` の行が 2 つ以上なら警告（GAS が台帳を勝手に作った跡。inv-3 4-1）。

### 6-6. 移行の道具（`tools/migrate_from_sheets.py`、D）

残業の骨組み（`--source sheets`・`--dry-run`・`--snapshot-out`・空でなければ止める・1 トランザクション・`verify_in_db`・終了コード 2/1）を写す。

- **`PROD_SPREADSHEET_ID` 以外を拒む**（staging への移行も本番を読む＝読むだけ）。
- T_LEAVE_REQUEST を UNFORMATTED（SERIAL_NUMBER）と FORMATTED で全行読み、`sheet_codec.decode_row` で行ごとに読む（`cell_kinds` を行ごとに持つ）。
- **REQ_ID が空の行は飛ばし、行番号と値のあるセルを一覧に出す**（06 2-5 の半端な行）。REQ_ID の重なり・`LV-\d{8}-\d{3,}` でない REQ_ID・決まりに無い状態は止める。
- **古い行の年度・日付の型は計算し直さず保つ**（JST 9 時の休暇日、文字の休暇日、4/1 区切りの年度。inv-4 0 節 #3・#4）。年度と `computeFiscalYear_(休暇日)` が違う行の数は報告に出すだけ。
- `sheet_row`＝その行、`row_version=1`、`sheet_sync_state` は反映済み（`desired_version=synced_version=1`、`synced_values`＝今シートにある値）。
- `pdf_file_id` は `PDF_URL` の `/d/([A-Za-z0-9_-]+)` から取る（無ければ空）。
- **staging への移行では `PDF_URL`・`SIGN_IMAGE_URL`・`pdf_file_id` を空にして入れる**（罠 72。ステージングの GAS が本番のファイルに触らないため）。
- 5-2 の表も同じ実行で DB のタブから入れる。V_LEAVE_SUMMARY は移さない（DB から作る）。
- 照合（ここが一致しなければ先へ進まない）: 件数、状態別の件数、2 次承認済みの件数、**JST の日付**で比べた休暇日・振替元の一致（9 時の行も同じ日）、申請日時・承認日時を秒で、作業員 ID は GAS が読む形（`normalize_gas`）で比べる。移した直後の `python -m app.worker run sheet_reconcile --check` が差 0。DB から作った集計（6-3）と本番の V_LEAVE_SUMMARY が 12 列とも同じ（最終有給日は表示の文字で）。

### 6-7. 毎晩の照合（`sheet_reconcile`、2:10 JST）

- T_LEAVE_REQUEST を UNFORMATTED と FORMATTED で全行読み、V 列を DB と 1 セルずつ比べる。日付は JST の日付、日時は秒で比べる。G 列・N 列は比べるが書かない（差は報告だけ）。反映待ちの行は比べない。
- 差があれば DB の値でその行の V 列を書き直し、直した行・列・前後の値を `sync_runs.summary_json` と管理者の欄に残す。シートにあって DB に無い行（REQ_ID が空の行を含む）は消さず書かずに報告する。
- 歯止め（罠 49）: DB の `leave_requests` が 0 行か、シートより DB の行が少ないときは 1 行も書かずに止める（REQ_ID が空の行は数えない）。
- V_LEAVE_SUMMARY も DB から作った値と比べ、違えば 6-3 で書き直す。
- `--check` は書かずに数えるだけ（切り替えと戻しの確かめに使う）。

---

## 7. GAS の内部用の入口と、きっかけごとの呼び順（決定 B・K-5）

PDF・サイン画像の保存・台帳・承認完了メール・有給警告メールは、今の GAS の関数で行う。VPS の worker が、行の反映が済んでから、GAS の内部用の `doPost` を 1 件ずつ呼ぶ。

### 7-1. 呼び方と取り決め（G が `src/Internal.gs` に作り、F の `gas_client` が呼ぶ）

| 項目 | 決め |
|---|---|
| 入口 | 新しい `doPost(e)`（このアプリには今 `doPost` が無い）。本文の JSON が `internal === true` でなければ `{"ok":false,"error":"bad_request","retryable":false}` |
| 呼び方（`GAS_INTERNAL_AUTH`） | **`bearer`（第 1 案）**: `gas_client` が `google_delegate.get_access_token()`（委任のトークン。スコープに `drive` を含む）を `Authorization: Bearer` に付けて `/exec` を POST（本文に token）。GAS の Web アプリは `access: DOMAIN` のまま。内部用のデプロイは画面のデプロイと同じ版でも別の版でもよい（〇〇）。**`anonymous`（第 2 案）**: 内部用のデプロイだけ、マニフェストの `webapp.access` を `ANYONE_ANONYMOUS` にした別の版を当てる。画面のデプロイ（@106・@104、切り替えの後は移転モードの版）は `DOMAIN` の版のまま。本文の token だけで守る。どちらもスパイク（14 章 S-1）で決める |
| 第 2 案の守り | 匿名の版では誰でも `/exec` を GET できる。G は `doGet` の先頭で `Session.getActiveUser().getEmail()` が空なら画面を出さず、移転モードなら移転ページ、そうでなければ `このアドレスでは画面を開けません。` の 1 行を返す（**この分岐は第 2 案の版だけに効くよう、スクリプトのプロパティ `INTERNAL_ONLY_DEPLOYMENT_ID` と `ScriptApp.getService().getUrl()` の照合で入れる。照合の可否はスパイクで確かめる**。DOMAIN の版では今と同じ動き） |
| 302 の扱い | `/exec` への POST は `script.googleusercontent.com` へ 302 で送られ、GET で JSON を読む（httpx は別のホストへの転送で `Authorization` を外す。転送先が認証なしで読めるかをスパイクで確かめる） |
| 認証（どちらの案でも必須） | 本文の `token` をスクリプトのプロパティ `VPS_INTERNAL_TOKEN` と長さ一定の比較で比べる。無い・違うときは何もせず `{"ok":false,"error":"unauthorized","retryable":false}` |
| 本文 | `{"internal":true,"token":"…","action":"internal.<名前>","jobId":"<uuid>","reqId":"LV-…","expect":{…},"args":{…}}` |
| 返事 | いつも HTTP 200 の JSON。成功 `{"ok":true,"action","jobId","reqId","sign"?,"mail"?,"pdf"?,"trash"?,"ledger"?,"durationMs","replayed"?}`。失敗 `{"ok":false,"error":"<コード>","message","retryable":true｜false, 済んだ段の結果}`（失敗でも済んだ段（サインの URL・PDF）は入れる。VPS は先に記録する） |
| エラーのコード | `unauthorized`・`bad_request`（false）・`row_not_found`（シートに REQ_ID が無い。true）・`row_not_synced`（`expect` と合わない。true）・`busy`（ロックが取れない。true）・`ledger_unavailable`（台帳が開けない。true）・`ledger_header_mismatch`（false）・`ledger_unconfirmed`（書いた・消した後の読み直しで確かめられない。true）・`internal`（そのほかの例外。true） |
| ロック | **入口がスクリプトのロックを持ったまま最後まで動かす**（`LockService.getScriptLock().tryLock(20000)`、取れなければ `busy`。今の `generateLeavePdf_`・`updateLedger_` はロックを取らず、呼び出し側が持つ前提＝inv-3 5-2）。中で `api_*` を呼ばない（`api_getLeaveRequestById` だけは例外: 今の関数が中で使う。行は先に入口が確かめるので待ちは起きない） |
| 段の記録（冪等） | `jobId` ごとに、済んだ段（`sign`・`mail`・`pdf`・`trash`・`ledger`）の結果をスクリプトのプロパティ `ij:<jobId>:<段>` に、段を終えた直後に書く（値は小さい JSON。PNG・本文を入れない）。同じ `jobId` の再試行では、記録のある段を飛ばして記録の結果を使う（**メール 2 通・サイン画像 2 つ・PDF 2 つを防ぐ**。罠 22・32）。全段が済んだ返事は `CacheService.getScriptCache()` に 6 時間（`ij:<jobId>`）置き、同じ `jobId` には処理せず `"replayed":true` を足して返す。プロパティの記録は 7 日より古いものを `internal.ping` が消す。キャッシュ・プロパティの例外は握って続ける（罠 36） |
| `expect` | VPS が送った `{status, approvedAt, approvedBy1Email, approvedAt2, approvedBy2Email, leaveDate}` のうちあるものを、シートの行（入口が REQ_ID で 1 回読む。6 秒待ちをしない）と比べる。日時は JST の `yyyy-MM-dd HH:mm:ss`、日付は `yyyy-MM-dd`、ほかは `normalize_` した文字。違えば何もせず `row_not_synced` |
| 実行の印 | token を確かめた実行の中だけ、グローバルの `__KYUKA_INTERNAL_OK__ = true`（`finally` で戻す）。移転モードの `getDb_()` の止めはこの印があれば通す（罠 47） |
| 時間 | 1 回の呼び出しで 1 件だけ。VPS の待ちは `GAS_CALL_TIMEOUT_SECONDS`。時間切れは「結果が分からない」として同じ `jobId` で送り直す |
| 台帳の前の確かめ | 台帳を書く・消す段の前に、Settings の `LEDGER_SSID`（同じキーは下の行）を `SpreadsheetApp.openById` で開けるか確かめ、開けなければ書かずに `ledger_unavailable`（**今の `getLedgerSs_` の「新しく作る」を動かさない**＝変えるもの）。年度タブ（`^\d{4}年度$`。年度は休暇日から `computeFiscalYear_`）が既にあれば 1 行目が `ledger.gs:62-68` の 20 列と 1 字ずつ同じか確かめ、違えば `ledger_header_mismatch`。タブが無ければ通す（今の関数が作る） |
| 台帳の後の確かめ | `updateLedger_`・`removeLedgerRow_` は失敗を黙る（inv-3 T5）。呼んだ後に年度タブの A 列を読み直し、upsert なら REQ_ID の行がある、削除ならその年度タブに無い、を確かめる。違えば `ledger_unconfirmed` |

### 7-2. action（中身は今の関数を、今と同じ順序・条件で呼ぶ）

| action | `expect` | `args` | GAS がすること（この順。各段は 7-1 の段の記録をする） |
|---|---|---|---|
| `internal.ping` | — | — | 何も書かない（古い段の記録の掃除だけ）。`{ok:true, scriptId, spreadsheetId, spreadsheetTimeZone, scriptTimeZone, migrated, settings:{PDF_FOLDER_ID, LEDGER_SSID, MASTER_SOURCE_SSID, CALENDAR_SOURCE_SSID, MAIL_TO_SET: bool}, ledgerKeyRows, ledgerOpenable, pdfFolderOpenable, auth:'bearer'｜'anonymous'（自分の版のマニフェストから）, apiVersion:1}`。`MAIL_TO` の値は返さない |
| `internal.afterApprove1` | `status:'承認済'`・`approvedAt`・`approvedBy1Email` | `signBase64`（手書きのときの `data:image/png;base64,…`。電子印は `''`）、`leaveDate`（`yyyy-MM-dd`） | `api_approveLeaveRequest`（`leaveRequest.gs:505-551`）の裏の処理を写す: ① **sign**: `signBase64` があれば `try { url = saveSignImage_(reqId, signBase64, new Date(leaveDate + 'T00:00:00+09:00')) } catch { url = '' }`。無ければ `url = ''`。シートの `サイン画像URL` の列に `url` を書く（**電子印でも空で上書き**＝今のまま）。返事 `sign:{status:'saved'｜'cleared'｜'error', signImageUrl}` ② **mail**: `try { sendApprovalNotification_(reqId) } catch { 記録して続ける }`。`MAIL_TO` が空なら `skipped`。返事 `mail:{status:'sent'｜'skipped'｜'error', error?}` ③ **ledger**: 前の確かめ → `try { updateLedger_(reqId) }` → 後の確かめ。返事 `ledger:{status:'upserted'｜'error', sheetName}`。集計（`rebuildLeaveSummary_`）は呼ばない（VPS が書く） |
| `internal.afterApprove2` | `status:'承認済'`・`approvedAt2`・`approvedBy2Email` | `approver1Email`（DB の値）、`approver2Email`、`oldPdfUrl`、`oldPdfFileId`、`cleanupMode`（`'asis'`｜`'trash'`） | `api_approveLeaveRequest2`（`leaveRequest.gs:678-720`）の裏の処理を写す: ① **pdf**: 「もうある」の判定: シートの `PDF_URL` が空でなく `oldPdfUrl` と違えば作らずに `already`（2 つのタブ・時間切れの再送で PDF を 2 つ作らない）。そうでなければ `try { r = generateLeavePdf_(reqId, approver1Email, approver2Email) }`。成功なら `PDF_URL`（と、列があれば `PDF_FILE_ID`）に書く。返事 `pdf:{status:'created'｜'already'｜'error', pdfFileId, pdfUrl, pdfName, folderId, error?}` ② **trash**: `cleanupMode==='trash'` で `oldPdfFileId` があり、①が `created` のときだけ、前の PDF をゴミ箱へ（**新しい PDF ができてから捨てる**。失敗は記録して続ける）。`'asis'`（既定）なら何もしない（**本番に `PDF_FILE_ID` 列が無いので、今の GAS は前の PDF を捨てていない**）。返事 `trash:{status:'trashed'｜'skipped'｜'error'}` ③ **ledger**: PDF の成否に関係なく、前の確かめ → `updateLedger_(reqId)` → 後の確かめ |
| `internal.afterUnapprove` | `status`（編集なら `'申請中'`、取消なら `'取消'`） | `reason`（`'edit'`｜`'cancel'`）、`oldLeaveDate`（**編集の前の休暇日** `yyyy-MM-dd`。取消は取消のときの休暇日）、`clearSign`（編集で承認が外れたとき true）、`oldPdfUrl`、`oldPdfFileId`、`cleanupMode` | `api_editLeaveRequest`（`:937-954`）・`api_deleteLeaveRequest`（`:986-997`）の裏の処理を写す。**シートから休暇日・PDF を読まず、引数を使う**（VPS の行の反映が先に新しい値を書くため。inv-3 T4）: ① **sign**: `clearSign` なら `サイン画像URL` を空に（ファイルは消さない＝今のまま） ② **trash**: `cleanupMode==='trash'` で `oldPdfFileId` があれば、シートの `PDF_URL` が `oldPdfUrl` と同じときだけゴミ箱へ移し、`PDF_URL`（と `PDF_FILE_ID`）を空に。`'asis'` なら何もしない（今の GAS は `PDF_FILE_ID` 列が無いので `PDF_URL` も残している） ③ **ledger**: 前の確かめ → `removeLedgerRow_(reqId, oldLeaveDate)` → 後の確かめ（申請中の取消でも呼ぶ＝今のまま） |
| `internal.sendWarnMail` | — | `fiscalYear`（JST の今日の年度） | `checkAndSendWarnMails_`（`warnMail.gs:11-57`）を、`rebuildLeaveSummary_()` の行（33 行）を除いて写した新しい関数 `vpsSendWarnMail_()` で行う（日付の確かめ・未達者の抽出・`MAIL_TO`・件名・本文はそのまま。集計タブは VPS が直前に書いたものを読む）。**mail** の段の記録をする。返事 `mail:{status:'sent'｜'not_warn_day'｜'no_targets'｜'mail_to_empty'｜'error', count}` |

- どの action も、呼ぶ前に入口が REQ_ID でシートの行を探し（無ければ `row_not_found`）、`expect` と照合する。
- 一時スプレッドシートの片付け・印の探し方・ファイル名・フォルダ・令和の計算・`（承認印取得不可）` は今の関数がするとおり（変えない）。
- `api_approveLeaveRequest2Batch` は写さない（画面は 1 件ずつ呼ぶ）。

### 7-3. VPS 側の仕事（M。`services/gas_jobs.py`）

- 仕事は `jobs` の `lane='gas'`。worker の GAS のレーンは**全体で同時に 1 本**（日付フォルダ・台帳の取り合いを防ぐ）。`id` 順。**同じ申請の仕事は積んだ順で、前の仕事が `done`／`skipped` になるまで次は動かさない**。前が `failed` なら同じ申請の次も止めて管理者の欄に出す。別の申請は先へ進める。
- 動かす前に:
  1. `requires_json` の各行で `sheet_sync_state.synced_version >= version`。まだなら `run_after = now + 5 秒`（試行回数は増やさない）。`requires` は、その申請の行と、**同じ作業員 ID の反映待ちの行すべて**（台帳の「有給累計取得回数」が T_LEAVE_REQUEST の同じ作業員の全行を数えるため＝inv-3 7-3 の 3）。
  2. DB の申請を読み直し、`guard_json`（積んだときの状態など）と比べる。`afterApprove1`: 状態が `承認済` で `approved_at` が積んだときと同じでなければ `skipped`（`sign_png` を NULL に）。`afterApprove2`: 状態が `承認済` で `approved_at2` が積んだときと同じでなければ `skipped`。`afterUnapprove`・`sendWarnMail` は飛ばさない。
  3. `expect` と `args` を作る（DB の今の値。`signBase64` は `sign_png` を `data:image/png;base64,` 付きの文字にする）。`cleanupMode` は `quirks.PDF_CLEANUP_MODE`（既定 `'asis'`）。
  4. 最後の `gas_ping` の照合（9-1）が 10 分以内に成功していなければ動かさない（`run_after` を 1 分後に）。
- 結果の記録（同じトランザクション）: `record_gas_result(db, req_id, resp)`。`sign.signImageUrl` があれば（空でも）`leave_requests.sign_image_url` に入れ、`sign_png` を NULL に。`pdf.status` が `created`／`already` なら `pdf_url`・`pdf_file_id` を入れる。`trash` が `trashed` で、afterUnapprove なら `pdf_url`・`pdf_file_id` を空に。**G 列は行の反映の対象外なので `synced_values` は触らない。GAS が返した G 列の値は状態に関係なくいつも DB に入れる**（DB の G 列をシートと同じに保つ）。`jobs.result_json` に返事を入れる（PNG は入れない）。
- 再試行: `retryable:true`・通信エラー・時間切れは 1・2・4・8・16・30 分、その後は 30 分ごとに 24 時間まで。`retryable:false` と 24 時間を過ぎたものは `failed`（管理者の欄）。利用者の操作は止めない。
- 有給警告メール（`enqueue_send_warn_mail`）: 動かす前に `summary_sheet.write(db)` を同期で行い、済んでから `internal.sendWarnMail` を呼ぶ。
- `gas_ping`（5 分ごと・起動時）: `internal.ping` を呼び、`safety.verify_gas_ping` の結果を `sync_runs(kind='gas_ping')` に残す。

### 7-4. きっかけごとの呼び順（今と同じ順序・条件）

| きっかけ | VPS の操作のトランザクション（W） | 積む仕事（M の関数） | 集計タブ |
|---|---|---|---|
| 申請 | 新しい行 | なし | なし |
| 1 次承認 | 8-7 | `enqueue_after_approve1`（サイン → メール → 台帳） | 予約 |
| 2 次承認 | 8-8 | `enqueue_after_approve2`（PDF →（trash）→ 台帳） | なし（GAS も書き直さない） |
| 編集（承認済みから） | 8-6 | `enqueue_after_unapprove(reason='edit', clear_sign=True)`（サインの列を空 →（trash）→ 台帳の行を消す） | 予約 |
| 編集（申請中のまま） | 8-6 | なし | 予約 |
| 取消 | 状態＝取消 | `enqueue_after_unapprove(reason='cancel', clear_sign=False)`（（trash）→ 台帳の行を消す） | 予約 |
| 有給警告メール（運用の欄） | — | `enqueue_send_warn_mail`（集計タブを書く → GAS が送る） | 直前に同期で書く |

### 7-5. GAS 側の変更（G。決まった所だけ。止める場所は inv-3 8 節の全部）

| ファイル・場所 | 変更 |
|---|---|
| `src/Internal.gs`（新） | 7-1・7-2 のすべて（`doPost`・token・ロック・段の記録・`expect`・台帳の前後の確かめ・`vpsSendWarnMail_`）。`INTERNAL_API_VERSION = 1` |
| `src/Cutover.gs`（新） | 残業の `src/Cutover.gs` を写して直す。`MIGRATION_TRIGGER_FUNCTIONS = ['syncAllMasters','checkAndSendWarnMails_','notifyTodayLeaves_','notifyPendingApproval1_','notifyPendingApproval2_']`（実物のトリガーに残っている恐れ＝inv-3 T3）。`renderMigratedPage_(url, page, reqId)`: `page` は 7 ページの名前だけ、`reqId` は `^LV-\d{8}-\d{3,}$` だけを引き継いだ `https://leave.lineworks-local.net/?page=…&reqId=…` へのボタン。題名と文言は「休暇届」。`<base target="_top">` とボタンの `target="_top"` は必須、1 秒後の `window.top.location` は「動けばよい」（罠 29）。運用関数: `enterMigrationMode(url)`（URL は引数必須。`^https://leave\.lineworks-local\.net/?$` 以外を拒む）、`enterMigrationModeForKyuka()`（確定の URL で包む）、`exitMigrationMode()`、`setInternalToken()`（乱数を作って保存し、長さだけ出す）、`removeAllTriggersForMigration()`（時刻のトリガーを全部消して一覧を記録）、`restoreTriggersAfterRollback()`（`syncAllMasters` の毎日 6 時台だけ。同じ関数のトリガーがあれば作らない）、`showMigrationStatus()`（読むだけ。版・移転モード・トリガー。`getDb_()` の止めを通らない）。運用関数は `Session.getActiveUser().getEmail() === Session.getEffectiveUser().getEmail()`（持ち主）のときだけ（罠 48）。**既定の URL を持たない**（`enterMigrationMode()` を引数なしで呼ぶと止まる） |
| `config.gs` の `getDb_()`（103 行） | 先頭に `assertDbOpenAllowed_()`（移転モードなら止める。`__KYUKA_INTERNAL_OK__` が立っていれば通す） |
| `menu.gs` の `doGet`（3 行） | **`getSettings_()`（5 行）より前**で、移転モードなら `renderMigratedPage_(url, e.parameter.page, e.parameter.reqId)` を返す（inv-3 T8） |
| `menu.gs` の `onOpen` | 移転モードなら「移転済み（新しいアドレス）」の 1 項目だけ |
| `menu.gs` の `setupAllSheets`（`getActive()` を直接開く）・`setupAllTriggers`・`setupSyncTrigger_`・`setupWarnMailTrigger_`・`setupNotifyTriggers_`・`migrateLeaveRequestHeaders_` | 先頭で `assertNotMigrated_()` |
| トリガーの 5 関数 | 先頭で `if (skipBecauseMigrated_('<名前>')) return;` |
| `syncCalendarMaster`・`syncStampMaster`・`syncDeptApprovers`・`rebuildLeaveSummary`・`checkAndSendWarnMails`（包み）・`debug_checkLeaveRequestSheet`・`test_generatePdf`・`test_stampLookup` | 先頭で `assertNotMigrated_()` |
| `api_*` の 30 関数（画面の 27＋`api_canAccessAdmin`・`api_getDeptApproverDepts`・`api_approveLeaveRequest2Batch`） | 先頭で `assertNotMigrated_()`（`api_getApproverList` は例外を握って `[]` を返すので、止めないと「承認者 0 人」に見える） |
| 内部用の入口が呼ぶ関数（`saveSignImage_`・`generateLeavePdf_`・`updateLedger_`・`removeLedgerRow_`・`sendApprovalNotification_`・`api_getLeaveRequestById`・`getUnderCountWorkers_`・`buildWarnMailBody_`） | **止めを置かない**（移転モードでも内部用の入口は動く） |

- ほかの関数（PDF・台帳・メール・業務）は 1 行も変えない。`appsscript.json` は第 2 案に決まったときだけ、内部用の版を作る間だけ `ANYONE_ANONYMOUS` にする（手順は C の手順書。HEAD は `DOMAIN` に戻す）。
- 移転モードでないとき（`MIGRATED_TO_URL` が無い）は、内部用の入口が増えた以外、今とまったく同じ動き（G の試験で確かめる）。
- 切り替えでは画面のデプロイ **@106 と @104 の両方**に移転モードの版を当てる（K-2。罠 60: `list-deployments --json` で版を確かめる）。

---

## 8. 同じにしなければならない業務の決まり（W）

### 8-0. 時刻（決定 C。必須）

- サーバーの「今」は `timeutil.now_utc()` だけ。「今日」は `timeutil.today_jst()`。`date.today()`・`datetime.now()`・`datetime.utcnow()` を使わない（F の試験が `app/` を grep して拒む）。各モジュールは `from app import timeutil` として呼ぶ。
- 年度・月次期間・祝日・勤務日・期限・期間・過去申請・日数は、すべて Asia/Tokyo の日付と時刻で行う。DB の日時は UTC、JSON は JST の文字。
- 1 回の呼び出しの中は同じ `ctx.now` を使う。

### 8-1. 採番（`generateReqId_`、`leaveRequest.gs:6-29`）

- `LV-` ＋**申請した日**（JST の今日）の `yyyyMMdd` ＋ `-` ＋ 連番。連番は `leave_requests` の `req_id` のうちその日の接頭辞で始まるものの、接頭辞の後ろの先頭の数字（JS の `parseInt(…,10)`。数字で始まらなければ数えない）の最大＋1。3 桁に満たなければ 0 で埋める（1000 件目は `1000`）。
- `pg_advisory_xact_lock` の中で数えて入れる（5-4）。取消の行も数える。

### 8-2. 年度・月次期間・過去申請の印

| 関数（`domain/fiscal.py`） | 決まり | GAS |
|---|---|---|
| `fiscal_year(d) -> int` | `m < 3 or (m == 3 and day < 16)` なら `year - 1`、ほかは `year`（**3/16 始まり**） | `leaveRequest.gs:34` |
| `fy_start(d) -> date` | `date(fiscal_year(d), 3, 16)` | `config.gs:184` |
| `monthly_period(d) -> (start, end)` | 日が 16 以上なら（その月 16 日, 翌月 15 日）、15 以下なら（前月 16 日, その月 15 日） | `leaveRequest.gs:50` |
| `is_within_current_fy(s) -> bool` | `s` が `^\d{4}-\d{2}-\d{2}$` に合い、`fy_start(今日)` の文字 ≦ `s` ≦ 今日の文字（**文字で比べる**。`2026-02-30` も通る＝今のまま） | `config.gs:193` |
| `is_retroactive(applied_at, leave_date_iso) -> bool` | どちらかが空・申請日時が読めなければ `false`。申請日時の JST の日付の文字 ＞ 休暇日の文字の先頭 10 字 | `config.gs:206` |
| `days_ago(leave_date_iso, today) -> int` | 休暇日の文字の先頭 10 字が日付でなければ 0。`(今日 − 休暇日)` の日数、0 以下は 0 | `config.gs:219` |

### 8-3. 勤務日・休日・祝日（`domain/calendar.py`。`masters.gs:220-560` を 1 行ずつ写す）

- 祝日（`japanese_holidays(today) -> dict[str, str]`）: **今日の年度** `fy` と `fy+1` の 2 年だけ計算する。固定（1/1 元日、2/11 建国記念の日、2/23 天皇誕生日、4/29 昭和の日、5/3 憲法記念日、5/4 みどりの日、5/5 こどもの日、8/11 山の日、11/3 文化の日、11/23 勤労感謝の日）、ハッピーマンデー（1 月第 2・7 月第 3・9 月第 3・10 月第 2 の月曜）、春分 `floor(20.8431 + 0.242194*(y-1980) - floor((y-1980)/4))`、秋分 `floor(23.2488 + …)`（同じ式の定数違い）。**振替休日**: その時点の祝日の日付を文字で並べ、日曜の祝日ごとに翌日から祝日でない日まで進め `振替休日`。**国民の休日**: 振替休日を足す**前**の並びで、隣どうしの差がちょうど 2 日、間の日が（その時点で）祝日でなく土日でもなければ `国民の休日`。
- `api_getWorkingDays`: `calendar` の表が空（GAS: タブが無い・2 行目が無い・`日付`／`区分` の見出しが無い）なら `[]`。今日 0 時から**年度末 3/15** まで 1 日ずつ: 区分が `休日`・`祝日` なら除く、`出勤土曜` なら入れる、ほかは月〜金で祝日（計算）でなければ入れる。
- `api_getHolidayDays(workerId)`: 年度開始 3/16 から**今日＋7 日**まで（年度末を越えても止めない）。カレンダーにある日: 区分が `休日`・`祝日` なら `{date, label: 区分}`、`出勤土曜` なら除く、ほかで土日なら `日曜`／`土曜`。カレンダーに無い日: 祝日なら `祝日（<名前>）`、日曜 `日曜`、土曜 `土曜`。`workerId` が真なら、同じ作業員（`normalize_gas` で比べる）で種類が `振替休暇`、状態が `取消` でない（**却下も使用済み**）行の振替元出勤日（Date か、文字を日付として読めるもの）を除く。
- `api_getPastWorkingDays(yearMonth)`: `normalize_gas(String(…))` が `^\d{4}-\d{2}$` か `^\d{6}$` ならその月、ほかは今月。範囲は「月初」と「年度開始」の遅い方から「月末」と「今日」の早い方、今日より後は打ち切る。勤務日の判定は `api_getWorkingDays` と同じ。`label` は `yyyy-MM-dd（<日〜土>）`。
- カレンダーの日付の読み方: セルが Date ならその瞬間の JST の日付、文字なら JS の `new Date(文字)` と同じに読める形（`yyyy-MM-dd` は UTC の 0 時＝JST で同じ日、`M/d/yyyy` は JST）。読めない行は飛ばす。同じ日が 2 行あれば下の行が勝つ。
- 境目の試験（必須）: 0:30 JST・8:59 JST の「今日」、3/15 と 3/16、15 日と 16 日、祝日が日曜（振替休日）、5/3〜5/6、9 月の国民の休日がある年（2026 年 9 月 22 日など）、年度末の直前の勤務日の上限。

### 8-4. 有給の期限・期間（`leaveSummary.gs`・`leaveStatusApi.gs`。今のまま）

- `computeWarnLevel_(paid, fy, now)`: `PAID_LEAVE_WARN`（6/30=1・8/31=2・12/31=3・2/15=4・3/15=5。3/16 以降の月日は `fy` 年、ほかは `fy+1` 年、期限は JST の **23:59:59**）を後ろから見て、`now < 期限` なら飛ばし、最初に過ぎた期限だけで判定: `paid < 必要数` なら不足 3 以上 `危険`、2 以上 `警告`、ほか `注意`。足りていれば `''`。過ぎた期限が無ければ `''`。`fy` は `Number(年度の値)`。
- `computeQuarterStatus_(dates, fy, now)`: `PAID_LEAVE_QUARTERS` の 5 件（春期 3/16〜6/30・夏期 7/1〜8/31・秋期 9/1〜**11/30**・冬期 12/1〜2/15・年度末 2/16〜3/15。開始は 0:00:00、終わりは 23:59:59 JST）。期間に入る休暇日の件数（**半日も 1**）、累計、`shortage = max(0, target − 累計)`、`isCurrent`・`isPast`・`isFuture`、`level`（今か過去だけ: 不足 2 以上 `danger`、1 `warn`、0 `ok`、未来は `''`）、`taken`、`overdue`。キーは `{id, label, targetCount, periodCount, cumulativeCount, shortage, level, taken, overdue, isCurrent, isPast, isFuture}`、`label` は `config.gs:53-57` の文字そのまま。
- `buildMilestones_(paid, fy, now)`: 5 件 `{label, month, day, dateText:'M/D', required, met, passed, deficit, remaining, state}`。`passed = now >= 期限`、`met = paid >= 必要数`、`deficit = remaining = max(0, 必要数 − paid)`、`state`: met なら `done`、過ぎていれば不足 3 以上 `risk`・2 以上 `warn`・ほか `note`、ほか `future`。
- 「過ぎたか」の比べ方の 3 通りの違い（`<`・`>=`・メールの `>`）は今のまま（Q12）。

### 8-5. 申請（`api_submitLeaveRequest`、`leaveRequest.gs:131-234`。文は 1 字も変えない）

`data` が `None` なら `{}` として扱う。順に（1〜6 は `{ok:false, error}` を 200 で返す）:

1. 休暇日 `normalize_gas(data.leaveDate)` が空か `^\d{4}-\d{2}-\d{2}$` でない → `休暇日の形式が不正です（yyyy-MM-dd）。`
2. `data.isRetroactive is True`（**真偽値の true だけ**）なら `is_within_current_fy(休暇日)` でなければ → `過去日申請は今年度内（年度開始の3/16から本日まで）の日付のみ指定できます。`。そうでなければ `api_getWorkingDays()` の一覧に無ければ → `指定された休暇日は勤務日ではありません。過去日の場合は「過去申請」から選択してください。`
3. `data.leaveType == '振替休暇'`（正規化しない完全一致）なら: 振替元 `normalize_gas(data.substituteDate)` が空か形が違えば → `振替休暇の場合は振替元出勤日を選択してください。`
4. 同じ月次期間でなければ → `振替休暇は同一月次期間（前月16日〜当月15日）内で消化する必要があります。\n休暇日: <yyyy-MM-dd>（期間: <start>〜<end>）\n振替元: <yyyy-MM-dd>（期間: <start>〜<end>）`
5. 使い回し（`validateSubstituteUniqueness_`）: `data.workerId` と振替元が真なら、作業員 ID（`normalize_gas` で比べる）が同じ・種類 `振替休暇`・状態が `取消` でない（却下も含む）行の振替元の日付の文字が同じなら → `指定された振替元出勤日（<yyyy-MM-dd>）は既に別の振替休暇申請で使用されています。`
6. 通れば採番して 1 行入れる: REQ_ID、年度＝`fiscal_year(休暇日)`（数）、部署 ID・部署名・作業員 ID・作業員名・休暇区分・半日区分・休暇種類・特別理由・有給詳細＝`data` の値そのまま（`値 || ''`。**正規化しない**）、休暇日、振替元＝`data.substituteDate` が真なら**種類に関係なく**その日付（形が違えば 400 `振替元出勤日の形式が不正です: <値>`＝GAS は不正な日付を書いて失敗する見込み。受け入れる差）、申請日時＝`ctx.now`、状態＝`申請中`、作成者メール＝ログインのメール。ほかの列は空。
- 半日区分・振替元・特別理由・有給詳細が種類と合わなくても書く（D-6。今のまま＝Q10）。区分・種類の値、部署・作業員がマスタにあるか、同じ人・同じ日の重複は確かめない（今のまま）。
- 文字の項目が 50,000 字を超えると 400 `<項目>が長すぎます（50000字まで）。`（シートのセルの上限。GAS も書き込みで失敗する。受け入れる差）。

### 8-6. 編集（`api_editLeaveRequest`、`leaveRequest.gs:852-966`）

- `patch` が `None` なら `{}`。申請が無ければ `申請が見つかりません。`。状態が `取消`・`却下` なら `取消/却下済みの申請は編集できません。`。
- `leaveDate`・`leaveKubun`・`halfDayType`・`substituteDate` を `normalize_gas`。順に: 休暇日が空でなく形が違えば `休暇日の形式が不正です: <値>`、区分が空でなく `全日休`・`半日休` でなければ `休暇区分が不正です: <値>`、区分が `半日休` で半日区分が `午前`・`午後` でなければ `半日区分を選択してください。`、振替元が空でなく形が違えば `振替元出勤日の形式が不正です: <値>`。
- 申請の種類が `振替休暇` なら: 使う休暇日＝新しい値か元の値（4-4-d の `leaveDate`）、使う振替元＝同じく。振替元が空なら `振替休暇には振替元出勤日が必須です。編集画面から振替元を指定してください。`、休暇日が空なら `振替休暇には休暇日が必須です。`、月次期間（8-5 の 4 と同じ文）、振替元が送られて元と違うときだけ使い回し（自分の REQ_ID を除く）。
- 書く: 休暇日が送られれば休暇日（`serial`）と年度（`fiscal_year`、数）。区分が送られれば区分、`全日休` なら半日区分を空、`半日休` で半日区分が送られればそれ（空なら書かない）。振替元は種類が `振替休暇` で送られたときだけ（画面は送らない）。**勤務日・年度内・過去かは確かめない**（今のまま）。
- 状態が `承認済` なら（**変更が何も無くても**）: 状態＝`申請中`、承認日時・1 次承認者メール・2 次承認者名・2 次承認者メール・2 次承認日時を空。DB の `sign_image_url` も空、`sign_png` を NULL（シートのサインの列は GAS が空にする）。`pdf_url` は `quirks.PDF_CLEANUP_MODE == 'asis'` なら残す（今のまま）。積む: `internal.afterUnapprove(reason='edit', old_leave_date=編集前の休暇日, clear_sign=True, …)`。
- 返す `{ok:true, newStatus}`。

### 8-7. 1 次承認（`api_approveLeaveRequest`、`leaveRequest.gs:471-557`＋K-4）

- `reqId` が偽なら `REQ_IDが指定されていません。`。`leave_requests` が 0 行なら `申請データがありません。`。無ければ `申請が見つかりません: <reqId>`。
- 行ロックの後で状態を見る（**K-4。変えるもの**）: `申請中` なら通す。`承認済` なら 400 `この申請はすでに承認済みです。`、`取消`・`却下`・そのほかなら 400 `この申請は取消済みです。`。2 つの画面からの二重承認は、後の方がこの文になる（同じ画面の二度押しは冪等キーで同じ返事）。
- `approverEmail` が M_STAMP にあるか、だれの印かは確かめない（今のまま）。
- 書く: 状態＝`承認済`、承認日時＝`ctx.now`、1 次承認者メール＝`approverEmail || ''`、DB の `sign_image_url`＝`''`（シートは GAS が URL か空を書く）。`signBase64` が真なら、先頭の `^data:image/\w+;base64,` を外して base64 を読み、読めれば `sign_png` に入れる。読めなければ入れない（GAS も保存に失敗して URL 空で続ける）。`signBase64` の文字が 3,000,000 字を超えれば 400 `サイン画像が大きすぎます。`（受け入れる差）。
- 積む: `enqueue_after_approve1`・行の反映・集計の予約。返す `{ok:true, reqId}`。

### 8-8. 2 次承認（`api_approveLeaveRequest2`、`leaveRequest.gs:619-726`）

- 順に: `reqId` が偽 → `REQ_IDが指定されていません。`、総務か管理者でない → `権限がありません（総務または管理者のみ）`、0 行 → `申請データがありません。`、無い → `申請が見つかりません: <reqId>`、状態が `承認済` でない → `1次承認されていない申請です。`。
- **2 次承認済み（2 次承認日時が空でない）なら 400 `この申請はすでに2次承認済みです。`（変えるもの。K-4 と同じ趣旨の直し。やり直しで PDF を作り直し、前の PDF を捨てる穴をふさぐ。オーナーへの報告事項）**。
- 書く: 2 次承認者名＝`stamps` の A 列（`normalize_gas`＋小文字）が `approver2Email` の小文字と一致した最初の行の C 列（`normalize_gas`）、空か無ければ `approver2Email`（`approver2Email` が空なら `''`）。2 次承認者メール＝`approver2Email`（そのまま）、2 次承認日時＝`ctx.now`。
- 積む: `enqueue_after_approve2`・行の反映（集計は予約しない）。返す `{ok:true, reqId}`。PDF は数十秒後にできる（画面は PDF を見ない）。

### 8-9. 権限（F の `permissions.py`。GAS と同じ）

| 関数 | 決まり | GAS |
|---|---|---|
| `is_admin(db, email)` | Settings `ADMIN_EMAILS` を `normalize_gas`、カンマで分けて `trim`＋小文字、空を除いた一覧にログインのメール（小文字）がある。空の一覧は誰も当たらない。**`BOOTSTRAP_ADMINS` は足さない**（画面の振る舞いを変えないため） | `auth.gs:18` |
| `is_somu(db, email)` | 同じく `SOMU_EMAILS` | `auth.gs:30` |
| `can_access_admin(db, email)` | 同じく `ALLOW_ADMIN_PAGE_EMAILS` | `auth.gs:6` |
| `dept_approver_depts(db, email) -> list[str]` | `dept_approvers` の行ごとに A（`normalize_gas`＋`trim`）・B（`normalize_gas`）。どちらか空なら飛ばす。B をカンマで分けて `trim`＋小文字、`normalize_gas(email).lower()` があればその部署。例外は `[]` | `auth.gs:76` |
| `is_dept_approver(db, email)` | 上が 1 件以上 | `auth.gs:45` |
| `is_operator(db, email)` | `is_admin` か `BOOTSTRAP_ADMINS` に含まれる。**`ops_*` と運用の欄だけに使う** | 新 |

- 1 次承認・編集・取消・申請・プレビュー・取得状況は誰でも（共有アカウントを信頼する作り。今のまま＝Q1）。

### 8-10. 並べ方（`jsutil`）

- JS の既定の `sort()`: `jsutil.js_sort_key`（UTF-16 の符号順）。
- `localeCompare(x)`（言語の指定なし）: ICU の既定の照合（`PyICU` の `Collator.createInstance(Locale('en_US'))`。GAS の既定の言語は推測。スパイクで部署名・REQ_ID・`yyyy/MM/dd`・`MM/dd HH:mm` の並びを GAS と突き合わせる）。
- `localeCompare(x, 'ja')`: `Collator.createInstance(Locale('ja'))`。
- PyICU が使えない環境の試験では、同じ並びになる表を試験で確かめる。合わない並びは受け入れる差として報告。

---

## 9. 安全装置（F が `safety.py`、最重要）

### 9-1. 環境ごとの許可リスト

| 対象 | production | staging |
|---|---|---|
| 書いてよいスプレッドシート | `PROD_SPREADSHEET_ID` だけ | `STAGING_SPREADSHEET_ID` だけ |
| 書いてよいタブ | `T_LEAVE_REQUEST`（`row_sync`・`reconcile` だけ。G・N 列は書かない）、`V_LEAVE_SUMMARY`（`summary_sheet` だけ）、`作業員マスタ`・`M_CALENDAR`・`M_STAMP`・`M_DEPT_APPROVERS`（`master_sync` だけ） | 同じ |
| 書いてはいけない（どの環境も） | マスタ元 `1iu5Hoak…`・**残業 DB `1Knx_kaQ…`**・台帳（両方の環境の ID）・PDF フォルダ・ほかのスプレッドシートすべて。`M_SYSTEM_SETTING`・`M_LOOKUP`・`BatchLogs`・雛形のタブ | |
| 使ってよい書き方 | `updateCells`、`values.update`／`values.batchUpdate`（`V_LEAVE_SUMMARY` とマスタ）、`values.batchClear`（`V_LEAVE_SUMMARY` とマスタの余りだけ）、`appendDimension`（ROWS、`T_LEAVE_REQUEST` の末尾だけ） | 同じ |
| 使わない書き方 | `values.append`・`deleteDimension`・`insertDimension`・`deleteSheet`・`addSheet`・`duplicateSheet`・`sortRange`・`moveDimension`・`T_LEAVE_REQUEST` の `clear` | |
| GAS の内部用の入口 | `GAS_INTERNAL_URL` のデプロイ ID が画面のデプロイ（@106・@104）でない。`internal.ping` の `spreadsheetId == PROD_SPREADSHEET_ID`、`scriptId == PROD_GAS_SCRIPT_ID_FIXED`、Settings の `PDF_FOLDER_ID == PROD_PDF_FOLDER_ID`・`LEDGER_SSID == PROD_LEDGER_SPREADSHEET_ID`、時間帯が `DB_SHEET_TIMEZONE`（GAS の `getSpreadsheetTimeZone()` は GMT のとき空を返す＝罠 68。空と GMT 系を同じに扱う）。`ledgerKeyRows > 1` は警告だけ | `spreadsheetId == STAGING_SPREADSHEET_ID`、`scriptId != PROD_GAS_SCRIPT_ID_FIXED`、Settings の 2 つの ID が `STAGING_*` と同じで本番と違う、時間帯が同じ、`MAIL_TO_SET` は真でよい（写しの `MAIL_TO` はオーナーだけ＝05 3 節） |

- 確かめる時: 起動時（合わなければ起動しない）と、**書き込みの直前**（`SheetsApi` の書く関数の中で `safety.assert_writable(spreadsheet_id, tab, purpose)`）。
- GAS の照合が合わなければ `sync_runs(kind='gas_ping', ok=false)` に理由を残し、GAS のレーンを止める。管理者の欄と `/healthz?detail=1` に出す。
- `/healthz` は外には `{"ok":true,"env":"staging|production"}` だけ。

### 9-2. ほかの守り

- 本番 GAS（`src/`）を勝手に push しない（2-1）。
- `tools/qa/acceptance.py` は `APP_ENV=staging` でしか動かない。作る申請の部署名に【受け入れテスト】の印。作ったセッションは最後に消す。
- 移行・控えの道具は `.env` の環境を対象にする。動かす前に `APP_ENV` と ID を確かめる（罠 76）。
- ステージングの準備（ディレクター）: 本番 DB の「コピーを作成」（GAS の写しもできる）→ 写しの `M_SYSTEM_SETTING` の**一番下の行**の `PDF_FOLDER_ID`・`LEDGER_SSID`（台帳のコピー。空にしない）・`MAIL_TO`（オーナーだけ）・`APP_URL`（写しのデプロイ）を書き換える（05 3 節）→ 写しの T_LEAVE_REQUEST の `PDF_URL`・`SIGN_IMAGE_URL` を空に → 写しの GAS に G のコードを入れ、内部用のデプロイを作り、`VPS_INTERNAL_TOKEN` を置く。写しのトリガーは作らない。

---

## 10. 担当と境目

### 10-1. 担当（それぞれ git worktree の自分のブランチ。ほかの担当のファイルを書き換えない。push しない）

| 担当 | 範囲（書いてよいファイル） | 始める条件 |
|---|---|---|
| **F 土台** | `docker-compose.yml`・`.env.example`・`api/Dockerfile`・`requirements*.txt`・`pytest.ini`・`alembic*`（全表）・`app/main.py`・`config.py`・`db.py`・`models.py`（全表）・`timeutil.py`・`jsutil.py`・`constants.py`・`safety.py`・`auth.py`・`google_delegate.py`・`permissions.py`・`rpc.py`（34 関数の閉じたリスト、未登録は 501）・`serialize.py`・`sheet_queue.py`・`audit.py`・`jobs.py`・`worker.py`・`services/sheets_api.py`・`gas_client.py`・`write_gate.py`・`web/nginx.conf`・`backup/`・`cloudflare/`・`docs/RUNBOOK.md`・`tests/conftest.py` と土台の試験 | 着手済み |
| **U 画面** | `app/pages.py`・`app/templates/pages/*`・`web/static/**`・`web/dev/**`・`tests/test_pages.py`・`tests/test_static_assets.py` | F の後すぐ（10-2 の形を信じ、mock-server で単独で進める） |
| **G GAS** | `src/Internal.gs`・`src/Cutover.gs`・7-5 の決まった行・`gas-tests/**` | F の後すぐ（F を待たなくてもよい） |
| **D データ** | `services/sheet_codec.py`・`row_sync.py`・`reconcile.py`・`summary_sheet.py`・`config_sync.py`・`master_sync.py`・`tools/migrate_from_sheets.py`・`tools/cutover_backup.py`（DB と台帳の複製）・`tools/restore_tabs_from_copy.py`（控えから `T_LEAVE_REQUEST` の V 列・`V_LEAVE_SUMMARY`・マスタ 4 タブを写し戻す。G 列は戻さない）・対応する試験 | F の後すぐ |
| **M GAS の仕事** | `services/gas_jobs.py`・`ops_status.py`・`tools/spike_gas_auth.py`・対応する試験 | F の後すぐ |
| **W 業務** | `app/domain/**`（`quirks.py`・`ops.py` を含む）・`tests/test_domain_*.py`・`tests/test_calendar.py`・`tests/test_fiscal_jst.py`・`tests/test_summaries.py`・`tests/test_concurrency.py` | F を main に取り込んだ後 |
| **C 受け入れと切り替え** | `tools/qa/acceptance.py`・`tools/qa/compare_views.py`（GAS と VPS の返事の総当たり比較）・`docs/cutover-runbook.md`・`docs/acceptance-checklist.md` | 最後（U・G・W・D・M の取り込みの後） |

作る順: **F → U・G・D・M・W を並行（W は F を main に取り込んでから）→ 順に取り込み、全試験 → つなぎ目の確かめ → C → レビュー 2 本（正しさ・安全）→ 修正。**

### 10-2. 担当をまたぐ関数（名前と引数をここで固定する。呼ぶ側は相手の完成を待たず、この形を信じて書く。試験では偽物を入れてよい）

| 関数 | 持ち主 | 呼ぶ側 |
|---|---|---|
| `config.get_settings() -> Settings`（`gas_internal_auth` を含む） | F | 全員 |
| `db.get_db()`・`db.get_sessionmaker()` | F | 全員 |
| `timeutil.now_utc()`・`now_jst()`・`today_jst() -> date`・`to_jst_text(dt, pattern) -> str`（`'yyyy-MM-dd'`・`'yyyy/MM/dd'`・`'yyyy-MM-dd HH:mm:ss'`・`'yyyy/MM/dd HH:mm'`・`'MM/dd HH:mm'`・`'yyyyMMdd'`） | F | 全員 |
| `jsutil.normalize_gas(s) -> str`・`js_sort_key(s)`・`js_string(v) -> str`（JS の `String(値)`）・`collate_default(a, b)`・`collate_ja(a, b)`・`sheet_readback(text) -> (value, kind)`（6-2 の「数に見える文字は数」） | F | W・D |
| `permissions.is_admin(db, email)`・`is_somu`・`can_access_admin`・`dept_approver_depts(db, email) -> list[str]`・`is_dept_approver`・`is_operator` | F | U（pages）・W |
| `rpc.register(name)`・`RpcContext`・`RpcError(message)` | F | W |
| `auth.require_user`・`auth.get_optional_user`・`auth.revoke_sessions(db, email) -> int` | F | U・W |
| `store.lock_request(db, req_id) -> LeaveRequest | None`（`FOR UPDATE`） | F（`sheet_queue.py` の隣に置く） | W |
| `sheet_queue.mark_dirty(db, table_name, pk) -> int`・`current_version(db, table_name, pk) -> int`・`requires_for(db, *rows) -> list[dict]` | F | W・M・D |
| `jobs.enqueue(db, kind, args, *, lane="gas", request_id="", priority=0, requires=(), guard=None, batch_id="") -> int`・`jobs.register_kind(kind, handler, *, lane, on_exhausted=None)`・`JobContext`・`SkipJob`・`RetryLater(seconds)` | F | M・D |
| `worker.register_schedule(name, *, at=(h, m) | every_minutes=n, task)` | F | M・D |
| `services.sheets_api.SheetsApi.for_db(db)`（`get_spreadsheet`・`batch_get_values`・`batch_update(spreadsheet_id, requests, *, tab, purpose)`・`values_update`・`values_batch_update`・`values_batch_clear`・`append_rows(spreadsheet_id, tab, n)`。書く関数はすべて `safety.assert_writable` を通す） | F | D |
| `services.gas_client.GasClient.from_settings(settings).call(action, *, job_id, request_id="", expect=None, args=None, timeout=None) -> dict`（`GAS_INTERNAL_AUTH=bearer` なら `google_delegate.get_access_token()` を `Authorization` に付ける。通信の失敗・時間切れは `GasTransportError`。本文の `requestId` の代わりに `reqId` を送る） | F | M |
| `google_delegate.get_access_token(scopes=None) -> str`・`get_delegate_credentials(db)` | F | F（gas_client）・D・tools |
| `safety.assert_writable(spreadsheet_id, tab, purpose)`・`safety.verify_gas_ping(ping, settings) -> list[str]` | F | D・M |
| `write_gate.check(db, settings) -> (bool, str)` | F | D |
| `domain.quirks`（定数: `PDF_CLEANUP_MODE='asis'`・`LAST_PAID_DATE_STYLE='gas'`・ほか 15 章の既定） | W | M・D（読むだけ） |
| `domain.fiscal.fiscal_year(d)`・`fy_start(d)`・`monthly_period(d)`・`is_retroactive(...)`・`days_ago(...)` | W | W・D（照合の報告） |
| `domain.summaries.summary_rows(db, now) -> list[dict]`（キー `fiscalYear` は元のセルの値と型、`deptId`・`deptName`・`workerId`・`workerName`・`paidCount`・`paidHalfCount`・`substituteCount`・`specialCount`・`totalCount`・`lastPaidDate`（`yyyy-MM-dd`）・`warnLevel`）・`last_paid_date_display(value, style) -> str` | W | D（`summary_sheet`）・M |
| `domain.ops.*`（4-2 の 7 関数） | W | U（RPC で） |
| `services.summary_sheet.mark_dirty(db) -> None`・`write(db) -> dict`・`rebuild_now(db) -> dict` | D | W・M |
| `services.config_sync.sync_sheet_config(db) -> dict`（`{ok, counts, overlaps, mismatches}`） | D | W・worker |
| `services.master_sync.sync_masters(db) -> dict`（`{ok, tabs:{名前:{ok, rows, skipped_reason}}, mismatches}`） | D | W・worker |
| `services.row_sync.run_once(session_factory, limit=50) -> dict`・`status(db) -> dict` | D | worker・M（ops_status） |
| `services.reconcile.run(session_factory, *, check_only: bool) -> dict` | D | worker・tools |
| `services.sheet_codec.COLUMN_RULES`・`encode_cell`・`decode_cell`・`decode_row`・`synced_values_for` | D | D・tools |
| `services.gas_jobs.enqueue_after_approve1(db, req_id) -> int` | M | W |
| `services.gas_jobs.enqueue_after_approve2(db, req_id) -> int` | M | W |
| `services.gas_jobs.enqueue_after_unapprove(db, req_id, *, reason: str, old_leave_date: str, old_pdf_url: str, old_pdf_file_id: str, clear_sign: bool) -> int` | M | W |
| `services.gas_jobs.enqueue_send_warn_mail(db) -> int` | M | W |
| `services.gas_jobs.record_gas_result(db, req_id, resp: dict) -> None`（7-3） | M | M（D は試験で使う） |
| `services.ops_status.get_ops_status(db) -> dict` | M | W |

`guard` の形は `{"status":"承認済","approvedAt":"<ISO UTC>","approvedAt2":"<ISO UTC>|"}`。`requires` の形は `[{"table":"leave_requests","pk":"LV-…","version":7}, …]`（M が `sheet_queue.requires_for` で作る）。

### 10-3. worker の形（F が骨組み、D・M が中身を登録）

- `python -m app.worker` は 3 本のスレッド: ①行の反映のレーン（`row_sync.run_once` を 2 秒ごと）、②GAS のレーン（`lane='gas'` を 1 本ずつ。`pg_try_advisory_lock` で全体 1 本）、③予定と `lane='default'` の仕事（集計タブの書き直し）。`if __name__ == "__main__": from app import worker; worker.main()`（罠 73）。起動ログに `schedules=` の一覧を出す。
- 予定（`sync_runs` の `schedule:<名前>` で枠を判定。起動の直後に走らせ直さない）:

| 名前 | 時刻 | 持ち主 |
|---|---|---|
| `sync_masters` | 毎日 6:00（GAS のトリガーの 6 時台） | D |
| `summary_sheet` | 毎日 6:50（＋書き込みの後の予約を 10 秒ごとに見る） | D |
| `sheet_reconcile` | 毎日 2:10 | D |
| `sync_sheet_config` | `SHEET_CONFIG_SYNC_MINUTES` ごと | D |
| `gas_ping` | 5 分ごと（起動時にも 1 回） | M |
| `purge` | 毎日 3:30（冪等キー 24 時間、期限切れのセッション、`done` の仕事 90 日、`audit_log` は消さない） | F |

- `python -m app.worker run <名前>`（例 `run sync_masters`、`run sheet_reconcile --check`、`run summary_sheet`）で手で動かせる。

---

## 11. 変えるもの／変えないもの

### 11-1. 変えるもの（これ以外は変えない）

| 変えること | 根拠 |
|---|---|
| 全画面で Google ログイン（共有アカウントのまま）。持ち時間 90 日 | K-0・K-6 |
| 1 次承認は「申請中」だけ（違えば `この申請はすでに承認済みです。`／`この申請は取消済みです。`）。2 つの画面からの二重承認・申請の二重押しを防ぐ（行ロック・状態の照合・冪等キー） | K-4 |
| **2 次承認は「承認済で 2 次承認がまだ」のときだけ（違えば `この申請はすでに2次承認済みです。`）。やり直しで PDF を作り直す穴をふさぐ（K-4 と同じ趣旨の直し。オーナーへの報告事項）** | ディレクター方針 7 |
| 保守の関数を画面から呼べなくし、運用者だけの `ops_*` にする。有給警告メールの手動送信は総務管理の運用の欄から | 既定（確認中） |
| 存在しない REQ_ID の約 6 秒の待ちをせず、すぐ返す | ディレクター方針 9 |
| PDF・サイン画像の保存・台帳・承認完了メールは GAS が裏で行う（数秒〜数十秒後）。失敗は再試行する（今は失敗すると黙って抜ける） | 決定 B・K-5 |
| 台帳が開けないとき新しい台帳を作らず、失敗として知らせる。書く前に見出しを、書いた後に行を確かめる | inv-3 T5 |
| 画面を開くたびに V_LEAVE_SUMMARY を書き直さない（書き込みの後と毎朝 6:50 に VPS が書く）。書き直しで空になる瞬間が無い | ディレクター方針 4 |
| マスタ同期で、元が読めない・空・見出しが違うなら上書きしない。行・列・タブを消さない | 決定 D |
| 設定・マスタのタブを人が直してから VPS の判定に効くまで最大 5 分（［今すぐ取り込む］あり）。GAS の PDF の印はすぐ効く | 6-5 |
| 業務データは書くたびに行単位で反映（行の位置・型・表示形式を保つ） | 決定 A |
| XSS の直し（3-8）。シートへ書く文字を数式にしない | 必須 |
| スマホ向けの配置（PC は今と同じ） | K-7 |
| 401 のとき書きかけ（手書きサインを含む）を退避して戻す | ディレクター方針 9 |
| 受け入れる小さな差: 自由入力の日付に見える文字が文字のまま、取消・編集が数秒以内に続くと承認メールを送らないことがある（7-3 の 2）、文字が 5 万字を超えたときの文、サイン画像の上限 | 6-2・7-3・8-5・8-7 |

### 11-2. 変えないもの（主なもの）

1 次承認・編集・取消・取得状況は誰でも、電子印は誰の印でも選べる、2 次承認者のメールを引数で受ける、編集を保存するだけで承認が外れる、編集の休暇日は何日でも選べる、取消は 2 次承認済みでもできて承認の列を残す、申請は入力の残り（半日区分・振替元・理由）をそのまま書く、同じ人・同じ日の重複を止めない、年度末の直前に翌年度を申請できない、振替の使い回しは却下も使用済み、出勤土曜・祝日の扱い（8-3）、半日の有給は集計で 0.5・期間別と台帳で 1、年度は集計で列の値・台帳で休暇日から、警告 12/31 と期間 11/30、「過ぎたか」の比べ方 3 通り、古い管理画面は承認済みが 0 件の人を出さない、取得状況の予定は部署を移る前の申請を出さない、承認待ちの年またぎの並び、2 次承認待ちに `halfDayType` が無い、preview の過去申請のバッジが出ない、`api_getLookupData` の結果を使わない、一括 2 次承認の戻り値を見ない、サインは画面の大きさが変わると消える、年度を変えても読み直さない画面、PDF の様式・令和・印・ファイル名・フォルダ、台帳の 20 列と癖（「私用の為」・縞）、承認完了メールの宛先・件名・本文・差出人、有給警告メールの送信日の確かめ、前の PDF を捨てない（`PDF_FILE_ID` 列が無いため）、画面の文言・alert・confirm・エラー文のすべて。

---

## 12. テストと受け入れ表

### 12-1. 動かし方

- Python: `cd vps-app/api && .venv/Scripts/python -m pytest -q`。**最後の行（passed/failed の数）まで読み、1 件でも落ちていたら報告に書いて先へ進まない**（罠 44）。SQLite（メモリー）で動くこと。外の API は呼ばない。
- PostgreSQL の結合試験（ディレクターが VPS で）: `docker build --target test`、DB `kyuka_test`、`@pytest.mark.postgres`（行ロック・advisory lock・jsonb・bytea）。
- JS（U）: `node --test vps-app/web/dev/`。画面は `node vps-app/web/dev/mock-server.js`（nginx と同じ CSP、`/__calls`、`/__fail?status=401|500&count=1`）。
- GAS（G）: `node --test gas-tests/`。`src/*.gs` を `vm` で読み、`SpreadsheetApp`・`PropertiesService`・`CacheService`・`LockService`・`DriveApp`・`GmailApp`・`ContentService`・`UrlFetchApp`・`Utilities`・`Session`・`ScriptApp`・`HtmlService`・`Logger` の偽物で動かす。
- 受け入れ（ディレクターが staging で）: `docker compose exec -T api env PYTHONPATH=/app python - < tools/qa/acceptance.py`（ssh の here-doc の中では動かさない＝罠 77）。

### 12-2. 担当ごとの必須の試験

| 担当 | 試験 |
|---|---|
| F | `test_config_safety.py`（必須値・定数との照合・staging と本番の取り違え・`GAS_INTERNAL_URL` が画面のデプロイなら拒む・タブの許可リスト・`values.append` と使わない書き方の拒否・残業 DB とマスタ元への書き込みの拒否・`verify_gas_ping` の全条件と GMT の空・ID の直書きを拒む grep）、`test_rpc.py`（Origin・415・413・404・400・501・401・403・冪等キーの再生と 409・業務エラーで鍵が残らない・34 関数の閉じたリスト）、`test_auth.py`（90 日・`prompt` を付けない・`/auth/me`・`safe_next`）、`test_gas_client.py`（`bearer` で `Authorization` が付く・`anonymous` で付かない・token をログに出さない・302 の後の GET）、`test_jobs.py`、`test_worker_schedule.py`、`test_timeutil.py`（`date.today`・`datetime.now`・`round(` を grep で拒む）、`test_jsutil.py`（`normalize_gas` の全角・`js_string(2026.0)=='2026'`・`sheet_readback('0042')==('42','number')`）、`test_permissions.py`（4 つの一覧・空でも全員にしない・`BOOTSTRAP_ADMINS` は `is_admin` に入らず `is_operator` に入る・部署承認者）、`test_migrations.py`、`test_web_config.py`（CSP・インラインの script と on 属性が 0・`mobile.css` の規則がすべて 820px 以下・`client_max_body_size 4m`） |
| U | `test_pages.py`（ログインへの 302 と `next` の `reqId`・page の振り分け・不明は top・admin/deptAdmin/somuAdmin の `no_auth` の文と題名 2 種・7 つの題名・`REQ_ID` に `</script>` を含む値・`APP_URL` が `PUBLIC_ORIGIN/`・`PORTAL_URL` が空なら出ない・`SHOW_OPS`）、`test_static_assets.py`（8 ページに on 属性 0・`escHtml`／`esc` が `'` を逃がす・`__escText(0)=='0'`・`new Date(` に空白入りの文字を渡す箇所の一覧が報告どおり）、`web/dev/shim.test.js`（ハンドラの付け方 4 通り・省略引数・200 の `{ok:false}` が成功に届く・400 で `err.message`・401 で失敗ハンドラを呼ばずログインへ・**サインの下書きの退避と復元**・二重送信が 1 回・409 の送り直し・冪等キーの保持と破棄）、`pages.test.js`（42 個の置き換え後に全ボタンが呼ぶ関数と引数が元と同じ。somuAdmin の 259・266） |
| G | `gas-tests/internal.test.js`（token 違い・無し・`internal` が無い本文・`jobId` の再生・**段の記録で再試行してもメール 1 通・サイン 1 つ・PDF 1 つ**・`expect` の不一致・行が無い・ロックが取れない・各 action の呼び順と条件（7-2 の表の行ごと）・電子印でサインの列を空で上書き・PDF が例外でも afterApprove2 は台帳を書く・`PDF_URL` が `oldPdfUrl` と違えば `already`・`cleanupMode='asis'` で何も捨てない・`'trash'` で新しい PDF の後に捨てる・afterUnapprove が引数の休暇日の年度タブから消す・台帳が開けないとき作らない（`SpreadsheetApp.create` が呼ばれない）・見出しの不一致・後の確かめ・`sendWarnMail` が `rebuildLeaveSummary_` を呼ばない）、`cutover.test.js`（移転モードでないとき今と同じ・`doGet` が `getSettings_` より前に移転ページ・`page`・`reqId` と `target="_top"`・不正な `page`／`reqId` を捨てる・`getDb_()` の止めと内部用の印・トリガー 5 関数が何もしない・7-5 の入口が全部止まる・`onOpen` が 1 項目・運用関数が持ち主以外で拒む・`enterMigrationMode()` が引数なしで止まる） |
| W | `test_domain_requests.py`・`approval.py`・`views.py`・`masters.py`・`ops.py`（27 関数それぞれ: 返す形のキーと日付の形・エラー文の 1 字ずつ・積む仕事と `mark_dirty`・権限・Date でないセルの扱い）、`test_calendar.py`（8-3 の境目のすべて・祝日の 2 年だけ・振替休日・国民の休日・カレンダーが空で `[]`・出勤土曜）、`test_fiscal_jst.py`（**0:30 JST の申請の REQ_ID が今日の日付**、**3/16 0:30 JST の年度**、**15 日 23:30 と 16 日 0:30 の月次期間**、期限の 23:59:59 と翌 0:00、過去申請の両端、`days_ago`）、`test_summaries.py`（半日 0.5・キーの年度の値・最初の行の部署・最終有給日・警告レベル・期間別・マイルストーン・並び。本番の写しで GAS の V_LEAVE_SUMMARY と同じ）、`test_concurrency.py`（**2 つの画面で同じ申請を 1 次承認しても仕事は 1 つ、2 回目は「すでに承認済み」**、**2 次承認も同じ**、**同じ冪等キーの二重送信で申請は 1 件**、同じ日の同時の申請で REQ_ID が重ならない、同じ振替元の同時の申請で 1 件だけ通る＝postgres） |
| D | `test_sheet_codec.py`（GMT のシートで休暇日が前日 15:00 のシリアル値・日時のミリ秒・`cell_kinds`・数に見える文字・表示形式の写し）、`test_row_sync.py`（差分だけ書く・**G 列と N 列を書かない**・新しい行は最後の行の次（A 列が空で別の列に値がある行を上書きしない）・ID の確かめと探し直し・見出しの違い・時間帯の違い・関所）、`test_summary_sheet.py`（12 列・並び・型・余りを消す・書いてから消す・0 行で書かない・予約の間引き）、`test_reconcile.py`（JST の日付・秒・G 列は書かない・DB が少なければ書かない・`--check`）、`test_config_sync.py`（下の行が勝つ・空でも勝つ・空なら上書きしない・6 タブ）、`test_master_sync.py`（1 列目が空の行を捨てる・見出しの確かめ・位置で読むタブの見出しの違いで止まる・USER_ENTERED・日付のシリアル値・余りは値だけ消す・空なら上書きしない・タブごとの失敗）、`test_migrate_from_sheets.py`（本番 ID 以外を拒む・REQ_ID が空の行を飛ばして一覧に出す・重なりで止まる・年度と古い型を保つ・staging で PDF とサインの URL を空・`pdf_file_id` の取り出し・移した直後の照合が差 0） |
| M | `test_gas_jobs.py`（7-4 の表の行ごとに積む action と args・`requires` に同じ作業員の反映待ちの行が入る・`guard` が変われば skipped と `sign_png` を消す・`record_gas_result` で G 列が入る・同じ申請の順番・前が failed なら止まる・再試行の間隔と 24 時間・`retryable:false` で failed・`gas_ping` が合わないと止まる・有給警告メールの前に集計タブを書く・`jobs.result_json` に PNG が入らない）、`test_ops_status.py`、`tools/spike_gas_auth.py` の試験（偽の GAS で両方の案） |

### 12-3. 受け入れ表（C が `docs/acceptance-checklist.md` に写す）

「自動」は 12-2 か `acceptance.py`。「JEV」は JEV ハンズでオーナーが確かめる実機の操作（報告は「実機で押した／自動だけ／まだ」に分ける）。「SB」はステージング（DB のコピーと写しの GAS）。

| # | 項目 | 確かめ方 |
|---|---|---|
| A-1 | 7 ページ＋権限エラー、`?page=`、不明は top、題名 | 自動 U。JEV: 各ページを開き、ヘッダーのリンクを全部押す（inv-2 15 節 A-1） |
| A-2 | 入口の判定（admin・deptAdmin・somuAdmin） | 自動 U。JEV: 権限の無い共有アカウントで 3 つの文と題名（A-28） |
| A-3 | 27 関数の形・エラー文 | 自動 W。SB: `compare_views.py` で同じ入力の GAS と VPS の返事を突き合わせる（全部署×全年度、全作業員の取得状況、全 REQ_ID の preview） |
| A-4 | 申請（全日有給・半日特別・振替の期間の中と外・過去申請） | 自動 W。JEV: inv-2 A-3〜A-8 |
| A-5 | 二重押し | 自動 F・U・W。JEV: 「申請する」を素早く 2 回で 1 件（A-9） |
| A-6 | 編集・取消（申請中・承認済・2 次承認済） | 自動 W・G。JEV: A-10・A-11。SB: 台帳の行が消え、サインの列が空、PDF_URL は残る（`asis`） |
| A-7 | 1 次承認（電子印・手書き）とサイン | 自動 W・G。JEV: A-14〜A-17（**手書きはスマホ・iPad の指とペン**）。SB: サインの PNG が休暇日のフォルダに 1 つ、SIGN_IMAGE_URL、承認完了メールがオーナーに 1 通、台帳の行 |
| A-8 | 1 次承認の状態の確かめ（K-4） | 自動 W。JEV: 取消した申請の sign を開いて承認 →「この申請は取消済みです。」。2 つのタブで同じ申請を承認 → 2 回目が「すでに承認済みです。」、メール 1 通 |
| A-9 | 2 次承認（1 件・一括）と PDF | 自動 W・G・M。JEV: A-18〜A-20。SB: PDF の見た目を本番の PDF と見比べる（印・令和・時間帯・理由）、台帳の 2 次承認日・2 次承認者・PDF のリンク、2 回目は「すでに2次承認済みです。」 |
| A-10 | 集計・取得状況（somuAdmin・deptAdmin・admin・leaveStatus） | 自動 W `test_summaries.py`。SB: `compare_views.py` の総当たり。JEV: A-21〜A-27 |
| A-11 | V_LEAVE_SUMMARY と統合業務メール | SB: VPS が書いたタブを GAS の `rebuildLeaveSummary_` の結果と 1 セルずつ比べる（値と型）。統合業務メールの `readerLeave.gs` と同じ条件（承認状態の文字・休暇日が Date・`APPROVED_AT2` が空）で写しを読み、件数が同じ |
| A-12 | Safari の日付 | iPhone の実機で preview の届出日・承認日・曜日（オーナーの写真） |
| A-13 | 行の反映の型 | SB: VPS が足した行・直した行を写しの GAS の `getValues()` で読み、休暇日・振替元・申請日時・承認日時・2 次承認日時が `Date`、`api_getLeaveRequestById` が元の値と同じ |
| A-14 | 行の位置・半端な行 | 自動 D。SB: 新しい行が最後の行の次に入り、REQ_ID が空の半端な行が変わらない |
| A-15 | マスタ同期と DB の 4 タブ | 自動 D。SB: GAS の `syncAllMasters` の後と VPS の後でタブを 1 セルずつ比べる（値と表示） |
| A-16 | 設定の取り込み・権限 | 自動 F・D。SB: `SOMU_EMAILS` を変えて 5 分以内（［今すぐ取り込む］ですぐ）に somuAdmin の入口が変わる |
| A-17 | GAS が止まったとき | SB: 写しの GAS の token を外して承認し、画面は動き、仕事が再試行され、管理者の欄に出て、戻すと流れる（メールは 1 通） |
| A-18 | 移転モード | 自動 G。リハーサル: @106・@104 の両方の `/exec?page=preview&reqId=LV-…` から新しいアプリの同じ申請が最上位の窓で開く。メールの `?page=admin` も |
| A-19 | 内部用の入口の呼び方 | スパイク S-1 の合格の条件（14 章） |
| A-20 | **0:30 JST の申請** | 自動 W。SB: 0 時台に 1 件出し、REQ_ID の日付・年度・シートの休暇日 |
| A-21 | CSP で全ボタンが動く | 自動 F・U。JEV: 8 ページの全ボタン・選択・サインの描画 |
| A-22 | スマホ・タブレット | 自動 F。iPhone・現場の Android・タブレットで申請・サイン・2 次承認。PC（1366px）で今の画面と見比べて差が無い |
| A-23 | XSS | 自動 U。JEV: 3-8 の氏名で 6 画面 |
| A-24 | 401 と書きかけ | 自動 U。mock-server の `status=401` で申請モーダル・編集・**手書きサイン**が戻る |
| A-25 | ステージングから本番に書かない | 自動 F。staging で本番の ID・台帳・残業 DB・マスタ元への書き込みが拒まれる |
| A-26 | 運用の欄 | JEV: 運用者で somuAdmin の欄が出る、ほかでは出ない。有給警告メール（送信日以外は「送信日ではありません」相当の結果） |
| A-27 | 移行の照合 | 自動 D と切り替えの `--check` 差 0 |
| A-28 | 速さ | acceptance で中央値（目安: 一覧 100ms 以下、承認 100ms 以下） |
| A-29 | 手順書のリハーサルと戻しの練習 | ディレクターが SB と staging で 2 回 |

---

## 13. 意図して残す固有の文字列

（各担当が、残した理由と一緒にここへ足す。）

- `imaizumi@lineworks-local.info`: `.env.example` の `BOOTSTRAP_ADMINS`・`DELEGATE_EMAIL` の例と `docs/RUNBOOK.md`。コードと試験には書かない（試験は `owner@lineworks-local.info`）。
- 1 章の ID とデプロイ ID: `app/config.py` の定数と `.env.example` の見本だけ。`1Knx_kaQ…` は「読むだけの元」としてだけ現れる。
- `https://leave.lineworks-local.net`: `src/Cutover.gs`（移転先の確かめ）と `.env.example`・`docs/RUNBOOK.md`。
- `C:\Users\imaizumi.LINEWORKS-NET\.secrets\kyuka\`: `docs/RUNBOOK.md`。
- 画面のテンプレートの文言（`休日出勤日` など）: GAS の HTML の写し。

---

## 14. スパイクで決めること（段階 2b。ステージングと本番の読むだけで）

| # | 決める値 | やり方 | 合格の条件・決め方 |
|---|---|---|---|
| S-1 | `GAS_INTERNAL_AUTH`（`bearer`／`anonymous`）と内部用のデプロイ ID | ① 写しの GAS に G の `Internal.gs` を入れ、`access: DOMAIN` のまま内部用のデプロイを作る ② staging の api コンテナで `python tools/spike_gas_auth.py --mode bearer`（`internal.ping`・token 違い・`Authorization` なし・`internal.afterApprove1` の偽の REQ_ID）③ だめなら内部用の版だけ `ANYONE_ANONYMOUS` にして `--mode anonymous` | 合格: `bearer` で 200 の JSON `{ok:true}` が返り、`scriptId` が写しの ID、`spreadsheetId` が staging の ID、token 違いは `unauthorized`、`Authorization` なしは JSON でない（`GasTransportError`）、302 の後の GET が読める、10 回続けて成功、1 回の所要が 5 秒以内。`bearer` が通れば `bearer` に決める（画面の版を変えずに済む）。通らなければ `anonymous` で同じ条件と、匿名の GET で画面が出ないこと（7-1 の守り） |
| S-2 | `DB_SHEET_TIMEZONE`・地域 | Sheets API `spreadsheets.get?fields=properties(timeZone,locale)` | 値を `.env` に（`Etc/GMT`・en_US の見込み） |
| S-3 | T_LEAVE_REQUEST の型の表と行の位置 | UNFORMATTED と FORMATTED で全行。`LEAVE_DATE` の小数（0.625＝JST 0 時か、0.375 以下＝古い 9 時の行か、文字か）、`APPROVED_AT2` の型、`WORKER_ID`・`FY_START_YEAR` が数か文字か、REQ_ID が空の行の位置と中身、空の行が先頭に約 500 行あるか（inv-4 と 06 の食い違い） | `COLUMN_RULES` と移行の道具を直す |
| S-4 | 最終有給日の画面の文字（K5） | 本番の somuAdmin・deptAdmin・admin・leaveStatus を開いて見る（書かない）。V_LEAVE_SUMMARY の `最終有給日` の型 | `quirks.LAST_PAID_DATE_STYLE` の文字の形（長い英語の日付なら写す、`yyyy-MM-dd` なら `'ymd'`） |
| S-5 | マスタのタブの見出し | 作業員マスタ・M_CALENDAR・M_STAMP・M_DEPT_APPROVERS と元のタブの 1 行目を 1 字ずつ | 6-4 の見出しの確かめの値 |
| S-6 | PDF 1 件の所要時間と失敗の頻度 | 写しで 3 件。本番の「実行数」で `api_approveLeaveRequest2` の秒数 | `GAS_CALL_TIMEOUT_SECONDS`（PDF の最大の 3 倍以上） |
| S-7 | 実物のトリガー | 本番 GAS の「トリガー」画面（読むだけ） | `syncAllMasters` 以外があれば切り替えで消す一覧に（inv-3 T3） |
| S-8 | 承認完了メールが届いているか | 本番の「実行数」の `メール通知エラー` のログ、オーナーの送信済み（inv-3 T2） | 届いていなければ 15 章で聞く（今のまま＝GAS の同じ関数を呼ぶので結果も同じ） |
| S-9 | 台帳 | `LEDGER_SSID` の行の数、年度タブの一覧と 1 行目の 20 列 | 7-1 の見出しの確かめの値 |
| S-10 | `localeCompare` の既定の言語 | 写しの GAS で部署名・作業員コード・日付の文字の並びを出し、PyICU と比べる | 8-10 の照合の言語 |
| S-11 | サイン画像の大きさ | 本番の Drive の `sign_*.png` 数件と、タブレットで描いた data URL の長さ | nginx と rpc の上限（4m・3MB）で足りるか |

---

## 15. オーナーに聞くこと（直す／今のまま。おすすめを先に。社外に公開すると重くなるものを上に）

どれも既定は「今のまま」（Q4・Q5 だけは既定で直して報告）。答えで変えるときは `app/domain/quirks.py`（W）か、示した小さな差分だけを直す。

| # | 癖 | おすすめ | 既定 | 変える所 |
|---|---|---|---|---|
| Q1 | 1 次承認・編集・取消を誰でもでき、電子印は誰の印でも選べる（inv-1 S1・S3） | 今のまま（全画面がログイン必須になり、社外の人は触れない。共有アカウントの運用と合う） | 今のまま | `quirks.APPROVE1_REQUIRES_STAMP_OWNER` など |
| Q2 | 2 次承認者のメールを画面から引数で受ける（総務・管理者なら別の人の印で 2 次承認できる。S4） | **直す**（ログインのメールを使う。画面の見た目は同じ） | 今のまま | `quirks.SECOND_APPROVER_FROM_LOGIN = False` |
| Q3 | preview・承認者の一覧で、メール・サインの URL・PDF の URL が誰にでも見える（S5） | 今のまま（ログイン必須。社内の画面のため） | 今のまま | — |
| Q4 | 保守の関数を画面から呼べなくし、運用者だけの `ops_*` にする | このまま進める（確認中の既定） | 直す | 4-2 |
| Q5 | 2 次承認のやり直しを止める（`この申請はすでに2次承認済みです。`） | このまま進める（報告事項。K-4 と同じ趣旨） | 直す | 8-8 |
| Q6 | 前の PDF を捨てない（本番に `PDF_FILE_ID` 列が無く、今の GAS は 2 次承認のやり直し・編集・取消で前の PDF を捨てず、編集・取消で `PDF_URL` も残す。再承認すると台帳の PDF のリンクが前の PDF を指す） | 今のまま（移行で Drive のファイルを捨てる動きを新しく増やさない。直すなら移行の後に `'trash'`＝新しい PDF ができてから捨てる順） | 今のまま | `quirks.PDF_CLEANUP_MODE = 'asis'` |
| Q7 | 最終有給日が長い英語の日付で出ている疑い（K5。スパイク S-4 の結果を添える） | 出ていれば**直す**（`yyyy-MM-dd`。統合業務メールは別に読むので影響なし） | 今のまま | `quirks.LAST_PAID_DATE_STYLE` |
| Q8 | 承認待ち一覧の年またぎの並び（12 月が 1 月より上） | **直す**（年を含めて並べ、表示は `MM/dd HH:mm` のまま） | 今のまま | `quirks.PENDING_SORT_WITH_YEAR` |
| Q9 | 編集を保存するだけで承認が外れる（変更が無くても） | 今のまま | 今のまま | — |
| Q10 | 半日→全日に戻しても半日区分・振替元・理由がそのまま送られ保存される（D-6・M16） | 今のまま（画面の入力の作りの話。直すなら別途） | 今のまま | — |
| Q11 | 有給警告メールの手動送信を運用の欄に置く（送信日以外は送らない今の確かめのまま） | 置く | 置く | 3-11 |
| Q12 | 半日有給の数え方（集計 0.5・期間別と台帳 1）、年度の決め方の違い、警告 12/31 と期間 11/30、「過ぎたか」の 3 通り（D1〜D3） | 今のまま（総務の見え方が変わるため） | 今のまま | — |
| Q13 | 同じ人・同じ日の重複申請、年度末の直前に翌年度を申請できない（inv-1 の Q9・Q10） | 今のまま | 今のまま | — |
| Q14 | サインは画面を回す・大きさが変わると消える（D-3） | 今のまま（タブレットの実機で困れば直す） | 今のまま | `sign.js` の `resize` の 1 か所 |
| Q15 | 統合業務メールのリンク（Settings `APP_URL`）を新しい URL に書き換える時期（Q-B7） | 見守りが終わるまで今のまま（GAS の `/exec` → 移転ページ → 新しいアプリ） | 今のまま | 運用 |
| Q16 | 承認完了メールが届いていない疑い（S-8 の結果次第） | 届いていなければ、移行と別に直す（スコープを足すのは GAS 側） | 今のまま | — |

## 確かめること（まとめ）

- 14 章の S-1〜S-11（特に S-1 の呼び方、S-3 の行の位置、S-4 の最終有給日）。
- 本番の申請中・2 次承認待ちの件数を切り替えの直前にもう一度数える（06 2-5）。
- 統合業務メールのトリガーが動いているか（読むだけ。止めない）。
