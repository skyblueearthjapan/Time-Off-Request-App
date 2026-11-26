# 写して使うもの・直すもの・新しく作るもの

作成: 2026-11-26。写す元は残業の `D:\projects\zangyo-kyujitsu-app\github-repo\vps-app\`（master `1f6d5f7`、本番で動いている版）と `src/Cutover.gs`・`src/Internal.gs`。
残業は貯蔵品・工番部品の土台を写して作ったので、**3 件の中で一番新しく、罠 60〜77 の手当てが入っている**。このアプリはこれを写す。

> 写したら、残業・貯蔵品・工番部品の固有の文字列が残っていないことを grep で確かめる（スキル 02 の 17 節。`overtime`・`Overtime`・`1Knx_kaQ` を DB の ID として使う箇所・`1GPNBR9U`・`Requests`・`WorkLogs`・`approve2`・`supplies`・`parts`）。ただし `1Knx_kaQ` は**このアプリでも読み先のマスタとして正しく使う**ので、読み先の設定にだけ残す。

## 1. ファイルごとの扱い

印: **そのまま**＝名前の置き換えだけ、**直す**＝形を残して中身をこのアプリに合わせる、**新しく**＝前例に無い、**使わない**。

### 1-1. 土台（F）

| 残業のファイル | 扱い | 直す点 |
|---|---|---|
| `docker-compose.yml`・`.env.example` | 直す | 接頭辞 `kyuka-`、サブネット `10.94.0.0/24`、web の固定 IP。`.env` の ID 一式（`PROD_SPREADSHEET_ID=16v1tQNC…`、`PROD_PDF_FOLDER_ID=1MQibR08…`、`PROD_LEDGER_SPREADSHEET_ID=1B7tnsT4…`、`PROD_GAS_SCRIPT_ID_FIXED=15aGUL5i…`、読み先 `MASTER_SOURCE_SSID`・`CALENDAR_SOURCE_SSID`）。Sebastian の項目を外す |
| `api/Dockerfile`・`requirements*.txt`・`pytest.ini`・`alembic.ini` | そのまま | `python:3.12-slim-bookworm`（罠 1）。PDF を VPS で作らないので playwright は入れない |
| `app/main.py`・`db.py`・`audit.py`・`serialize.py`・`jsutil.py`・`timeutil.py` | そのまま | `timeutil` の JST の関数と、`date.today` を拒む試験も写す |
| `app/config.py`・`safety.py` | 直す | 許可するタブ（A: `T_LEAVE_REQUEST`・`V_LEAVE_SUMMARY`、C: マスタ 4 つ）、本番の ID の照合、ステージングの ID。Sebastian のトークン検査を外す |
| `app/auth.py`・`google_delegate.py` | そのまま | セッションの持ち時間（残業 #24 は 90 日。04 の Q-A3）。スコープは Sheets（と控えの道具の Drive）だけ。**Gmail・Drive への書き込みは VPS からしない**（GAS に残すため） |
| `app/permissions.py` | 直す | 4 つの一覧（`ADMIN_EMAILS`・`SOMU_EMAILS`・`ALLOW_ADMIN_PAGE_EMAILS`・M_DEPT_APPROVERS の B 列）と `BOOTSTRAP_ADMINS`。空でも全員にしない |
| `app/rpc.py` | 直す | 閉じたリストを 27 関数＋管理用に。未登録は 501。冪等キー・`expectedStatus` の照合（K1・K2 を直すと決めた場合） |
| `app/jobs.py`・`worker.py`・`sheet_queue.py` | そのまま | 罠 73 の直し（`__main__` から `app.worker.main()`）が入っていることを確かめる。レーン: 行の反映／GAS の仕事／定期 |
| `services/sheets_api.py`・`gas_client.py`・`write_gate.py` | そのまま | 移転モードの ping と書き込みの関所はそのまま使える |
| `web/nginx.conf` | 直す | `client_max_body_size`（サインの base64。実物は 10〜14KB だが、高解像度の端末で 1MB 程度まで見込む）。Sebastian の経路を外す |
| `backup/`・`cloudflare/setup_tunnel.py` | 直す | 接頭辞・Tunnel 名 `leave-vps`。`.out/` を持ち込まない |
| `alembic/versions/0001`・`0002` | 新しく（形は写す） | 表は 3 節 |
| `docs/RUNBOOK.md` | 直す | パス・コンテナ名・タブ名 |

### 1-2. 画面（U）

| 残業のファイル | 扱い | 直す点 |
|---|---|---|
| `web/static/js/gas-shim.js` | そのまま | `google.script.run` の置き換え。成功の中の `{ok:false, error}`（`api_submitLeaveRequest`）を成功ハンドラに渡す形を守る |
| `web/dev/port-pages.js`・`rpc-spec.js`・`mock-server.js`・`shim.test.js`・`pages.test.js` | 直す | 対象の HTML 7 つ＋`no_auth`。on 属性 50 個を `data-action` に置き換える |
| `app/pages.py` | 直す | `page` の振り分け（top・preview・sign・admin・deptAdmin・somuAdmin・leaveStatus、不明は top）、入口の判定 3 つ（`00` の 1-2）、題名、`REQ_ID` を `json.dumps` で JS の文字に、`APP_URL` は自分の公開 URL、`PORTAL_URL` は Settings |
| `templates/pages/*.html`・`web/static/js/*.js` | 新しく（作り方は写す） | このアプリの 8 画面。**サイン（canvas・pointer イベント・`user-scalable=no`・meta refresh）** と **取得状況（2026-07 の新しい画面）** は前例に無い |
| `web/static/css/mobile.css` | 直す | このアプリは viewport の meta があるのに GAS の iframe で効いていなかった（罠 16）。スマホの配置にするかは 04 の Q-A5 |

### 1-3. 業務（W）

| 中身 | 扱い | 元の GAS |
|---|---|---|
| 申請・編集・削除・取得（`domain/requests.py` の形） | 直す | `leaveRequest.gs`（振替の月次期間・使い回しの確かめ、過去申請の範囲、REQ_ID の採番） |
| 1 次承認・2 次承認（`domain/approval.py`・`secondary.py` の形） | 直す | 1 次は権限なし・印は選んだ人（K1 を直すなら状態の照合を足す）。2 次はログイン中のアカウント |
| 勤務日・休日・過去の勤務日・祝日の計算 | **新しく** | `masters.gs:220-560`（日本の祝日の近似式・振替休日・国民の休日をそのまま写す。JST） |
| 集計（V_LEAVE_SUMMARY）・警告レベル・期間別・マイルストーン | 直す（`domain/summaries.py`・`dashboards.py` の形） | `leaveSummary.gs`・`leaveStatusApi.gs`（K3 の数え方を今のまま） |
| 部署・作業員・承認者の一覧（`domain/masters.py`） | 直す | `masters.gs`・`auth.gs` |
| `domain/fiscal.py` | 直す | 年度 3/16、月次期間 16 日（残業の締め月と同じ考え方） |
| `domain/sebastian.py`・`worklog.py`・`breaks.py` | 使わない | — |
| `tools/qa/acceptance.py`・`staging_session.py` | 直す | このアプリの流れ（申請→1 次→2 次→PDF→編集→取消）と時刻の境目 |

### 1-4. データ（D）

| 残業のファイル | 扱い | 直す点 |
|---|---|---|
| `services/sheet_codec.py` | そのまま | シートの時間帯でシリアル値。日付型・日時型の列の表（休暇日・振替元出勤日・申請日時・承認日時・APPROVED_AT2） |
| `services/row_sync.py` | 直す | 対象は T_LEAVE_REQUEST だけ（主キー REQ_ID、見出しの文字で列を引く、英語・日本語の見出しの両方）。GAS が書く列（サイン画像URL・PDF_URL・PDF_FILE_ID）は上書きしない |
| （新）集計タブの書き直し | **新しく** | V_LEAVE_SUMMARY を DB から作り、見出しと並び（年度_作業員ID の文字順）を今と同じに全行書く。毎朝 6:50 と書き込みの後（間引き） |
| `services/reconcile.py` | 直す | 毎晩の全行照合（T_LEAVE_REQUEST）。DB が少なければ書かない（罠 49） |
| `services/config_sync.py` | 直す | M_SYSTEM_SETTING（2 行目から、下の行が勝つ）と M_LOOKUP の取り込み |
| `services/master_sync.py` | 直す | 4 つの同期（作業員マスタ・M_CALENDAR・M_STAMP・M_DEPT_APPROVERS）。元を丸ごと写す形（今の GAS と同じ）。数式を書かない、空なら上書きしない、タブを作り直さない |
| `services/sheet_intake.py` | 条件つき | Sebastian の inout-request がこのシートに書くときだけ（`01` の 3 節） |
| `tools/migrate_from_sheets.py` | 直す | 本番 ID 以外を拒む。表は 3 節。ステージングへは PDF の ID・サイン画像 URL を空にして入れる（罠 72） |
| `tools/cutover_backup.py`・`restore_tabs_from_copy.py` | 直す | 控えの対象は DB と台帳 |

### 1-5. GAS の仕事（M）と GAS 側（G）

| 残業のファイル | 扱い | 直す点 |
|---|---|---|
| `services/gas_jobs.py`・`ops_status.py` | 直す | 仕事の種類（下の表）、`requires`（行の反映が済んでから）、PDF の仕事は全体で 1 本ずつ |
| `services/morning_batch.py` | 使わない | このアプリに朝バッチは無い |
| `src/Internal.gs`（888 行） | 直す | 入口は `doPost`（このアプリには今 doPost が無いので新しく足す）。token 認証・`jobId` の再生・`expect` の照合・`ping` はそのまま。action は下の表 |
| `src/Cutover.gs`（515 行） | 直す | 移転ページは `page`・`reqId` を引き継ぐ（残業は `page`・`date`）。止める場所は `getDb_()`（`config.gs:103`。全関数が通る。罠 47）と `doGet`。トリガー 1 本（`syncAllMasters`）を消す・戻す。旧フォームの処理は外す。Sebastian の転送は外す。運用関数は持ち主だけ（罠 48）。既定の URL を持たない（罠 24d） |

内部用の入口の action（案。今の関数をそのまま呼ぶ）:

| action | VPS が呼ぶとき | GAS で動く今の関数 |
|---|---|---|
| `internal.ping` | 1 分ごと・切り替えのとき | 新（scriptId・紐づくシート・Settings の `PDF_FOLDER_ID`・`LEDGER_SSID`・時間帯・移転モードを返す） |
| `internal.afterApprove1` | 1 次承認の後（行の反映が済んでから） | `saveSignImage_`（サインがあれば）→ サイン URL を行に書く → `sendApprovalNotification_` → `updateLedger_` |
| `internal.afterApprove2` | 2 次承認の後 | 前の PDF をゴミ箱 → `generateLeavePdf_` → PDF の 2 列を書く → `updateLedger_` |
| `internal.afterUnapprove` | 編集で承認が外れた後・取消の後 | PDF をゴミ箱 → PDF の 2 列を空に → `removeLedgerRow_` |
| `internal.ledgerUpsert` | 台帳だけ直したいとき（再試行・照合） | `updateLedger_` |

集計（`rebuildLeaveSummary_`）は VPS が DB から作ってタブに書く（GAS に頼まない）。

## 2. このアプリで新しく作るもの（前例に無い）

1. **サインの canvas の移植**: pointer イベント、`devicePixelRatio`、リサイズで消える動き、`toDataURL('image/png')` の base64 を API へ（本文の上限）。承認後の `meta refresh`。タブレットの実機確認が要る。
2. **集計タブの全書き直し**（V_LEAVE_SUMMARY）: 統合業務メールの読み先なので、型（年度は数値、最終有給日の書き方＝K5 の扱い）を今と同じにする。
3. **祝日・勤務日の計算**: 春分・秋分の近似式、振替休日、国民の休日、M_CALENDAR の上書き。0 時台・3/16・16 日の境目の試験。
4. **内部用の入口の action**（上の表）と、サイン画像の base64 を GAS に渡す取り決め（`doPost` の本文の大きさ）。

## 3. DB の表（案）

| 表 | 元 | 主な列 |
|---|---|---|
| `leave_requests` | T_LEAVE_REQUEST | req_id（主キー）・fiscal_year・dept_id・dept_name・worker_id・worker_name・leave_kubun・half_day_type・leave_date（date）・leave_type・substitute_date（date）・special_reason・paid_detail・additional_detail・applied_at・approved_at・status・sign_image_url・pdf_url・pdf_file_id・created_by_email・approved_by1_email・approved_by2・approved_by2_email・approved_at2（timestamptz）・`sheet_row`・`version` |
| `settings`・`lookup` | M_SYSTEM_SETTING・M_LOOKUP | key・value（下の行が勝つ） |
| `workers`・`calendar`・`stamps`・`dept_approvers` | 4 つのマスタ | 元の列を持つ（位置の意味も） |
| 土台 | — | `sessions`・`google_credentials`・`jobs`・`idempotency_keys`・`sync_runs`・`sheet_sync_state`（`mail_log` は VPS から送らないので作らない） |

## 4. 工数の見込み（担当ごと）

残業の実績は、コミット時刻から読んだ担当ごとの作業時間（推測を含む。始まりは前の担当の取り込みか仕様書のコミット）。このアプリの見込みは推測。

| 担当 | 残業の実績 | このアプリの見込み | 理由 |
|---|---|---|---|
| 計画（棚卸し 4 人・計画書・抜け調査 5 観点） | 〜05:54 にまとめてコミット（所要は記録なし） | 2〜3 時間 | この下調べ（00〜05）が下敷きになる |
| 仕様書 | 06:17 にコミット（計画書の後 20 分ほど） | 30〜60 分 | 関数 27、内部用の入口の action 5 |
| **F 土台** | 約 25 分（06:17→06:41） | 20〜30 分 | 写して名前を置き換える。Sebastian を外す |
| **U 画面** | 約 20 分（06:17→06:38、6 画面） | 40〜60 分 | 8 画面、on 属性 50、サインの canvas と取得状況が新しい |
| **G GAS** | 本体 20 分（06:13）＋直し 4 回（〜08:33） | 60〜90 分 | doPost を新しく足す。action 5。移転ページの引数が `reqId` |
| **W 業務** | 約 25 分（06:41→07:06、29 関数＋管理 9） | 40〜60 分 | 祝日・勤務日・集計・期間別の計算が多い |
| **D データ** | 約 27 分（06:41→07:08）＋D2 取り込み 約 50 分 | 30〜45 分（取り込みが要れば＋45 分） | 対象のタブが 1 つ＋集計タブ＋マスタ 4 |
| **M GAS の仕事** | 約 16 分（06:41→06:57） | 15〜25 分 | 朝バッチが無い。仕事の種類 4 |
| **C 受け入れ道具・手順書・受け入れ表** | 約 20 分（07:06→07:27） | 30〜45 分 | 手順書は残業の B 章を写して直す |
| R レビュー 2 本と修正 | 約 50 分（07:09→07:59）＋直し | 45〜60 分 | 1 次承認に権限が無い作りの扱い（外に公開するので安全のレビューで出る） |
| 置く・ステージング・受け入れ・JEV | 約 1 時間（07:59→08:56） | 1.5〜2 時間 | サインの実機（タブレット）、PDF の見比べ |
| 切り替え | 16 分（09:30→09:46） | 20〜40 分 | フォーム・Sebastian が無いぶん短い。デプロイが 2 つなら手順が増える |

合計の見込み: 作業 7〜10 時間（推測）。前もって決めることが早く決まれば 1 日で切り替えまで届く。見守りは 5 営業日。
