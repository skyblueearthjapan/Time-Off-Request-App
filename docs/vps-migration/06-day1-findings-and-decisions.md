# 初日の確かめの結果と、オーナーの決定（2026-11-26）

作成: 2026-11-26 午前。ディレクター（本体）が、読むだけの確かめ（`05` の C-1〜C-12）とオーナーへの質問（`04`）の結果をまとめた。
計画書（`docs/vps-migration-plan.md`）の 4-0・9 章の元になる。**下準備資料（00〜05）と食い違う点は、この文書が正**（実物を読んだため）。

## 1. オーナーの決定（AskUserQuestion の回答。2026-11-26）

| # | 項目 | 決定 |
|---|---|---|
| K-0 | 残業・休日出勤申請の共通の決定（`04` の Q-0 の表。残業の 9-1 #1・#3・#4・#5・#6・#7・#8・#9・#10・#12・#17・#18〜#21（決定 A〜D）） | **すべて同じでよい**。VPS の外へのバックアップと死活監視は未決のまま、後日 4 アプリまとめて決める |
| K-1 | 公開 URL | **`https://leave.lineworks-local.net`**（Tunnel 名 `leave-vps`、アプリ名 `kyuka`、`/opt/kyuka`、サブネット `10.94.0.0/24`） |
| K-2 | 本番のデプロイ（@106・@104） | AI が調べる。切り替えでは**両方**に移転モードを当てる（結果は 2 節） |
| K-3 | 前もって任せる範囲 | 残業と同じ（`/opt/kyuka` への配置、Cloudflare の Tunnel・DNS、OAuth クライアント、GitHub への push、本番シートを読む）＋ **Sebastian を読むだけ可**。本番の切り替え・本番 GAS の変更・ポータルの書き換え・利用者へのメールはその都度聞く |
| K-4 | 1 次承認（サイン）の状態の確かめ（`04` の Q-B1・Q-B2） | **直す**。「申請中」だけ承認できる。違えば「この申請はすでに承認済み／取消済みです」。2 つの画面からの二重承認と、申請の二重押し（冪等キー）も防ぐ。ふだんの操作と見た目は変えない |
| K-5 | 承認完了メール（Q-B4） | **今のまま、今の GAS から送る**（宛先 `MAIL_TO`・本文・差出人は同じ） |
| K-6 | ログインの持ち時間（Q-A3） | **90 日** |
| K-7 | スマホ・タブレット（Q-A5） | **スマホ向けの配置を足す**（PC の見た目は変えない）。手書きサインは現場のタブレットでオーナーが実機確認 |

まだ聞いていないもの（`04` の Q-B3・Q-B5〜Q-B10、棚卸しで出た癖）は、棚卸しの結果をまとめてから 1 回で聞く。それまでの既定は「今のまま」。

## 2. 確かめた事実（読むだけ。確認済み）

### 2-1. VPS（`ssh lineworks-vps-user`、hostname `srv1508169`）

- `/opt/kyuka` は無い（使える）。`/opt` にあるのは chozouhin・overtime・parts ほか 20 個余り。
- docker のサブネット: 10.91（chozouhin）・10.92（parts）・10.93（overtime）・10.77〜10.79・10.84・10.89.2 が使用中。**`10.94.0.0/24` は空き**。
- コンテナ名に `kyuka-` で始まるものは無い。

### 2-2. GAS のデプロイ（`clasp list-deployments`）

- @HEAD `AKfycbw_GAlP…`、**@104** `AKfycbwHI1uz…`（振替元プルダウン月次絞込）、**@106** `AKfycbx7sB9x…`。下準備と同じ。
- **本番は @106**: M_SYSTEM_SETTING の `APP_URL` と、社内ポータル APP_MASTER の url が、どちらも `AKfycbx7sB9x…/exec` を指す。@104 はブックマークからだけ届く。切り替えでは両方に移転モードを当てる（K-2）。

### 2-3. 社内ポータル

- APP_MASTER の **key `leave_request`**（sort 3、label「休暇届け申請」、open_mode `same_tab`、status OK）。url は @106 の `/exec`。key と open_mode の値の後ろに空白が付いている（探すときは trim する）。

### 2-4. Sebastian（VPS `/root/.hermes/skills/internal/inout-request`。読むだけ）

- **休暇届のシートには書いていない**。スクリプトが触るのは外出記入板（`1w_AXXnn…`）と残業 DB（`1Knx_kaQ…`）だけ。SKILL.md は休暇届を「社内ポータルへ案内」としている。`/root/.hermes` 全体でも、休暇届の DB・台帳・デプロイの ID は 0 件。
- → **`sheet_intake` は要らない**（`04` の Q-B11 は聞かない）。

### 2-5. 本番シートの実物（gviz の CSV を Edge で取得。置き場は `~/.secrets/kyuka/prod-read-20261126/`、git の外）

**T_LEAVE_REQUEST の見出しは英語で、列の並びが `setupAllSheets` の形（下準備 02 の 3 節）と違う。26 列。`PDF_FILE_ID` 列は無い。`UPDATED_AT` 列がある（全行が空）。**

| 列 | 見出し | 表示の形（gviz） |
|---|---|---|
| A | REQ_ID | `LV-yyyyMMdd-NNN`（279 件すべてこの形。重複 0） |
| B | FY_START_YEAR | `2026` |
| C | STATUS | `承認済` 260・`取消` 18・`申請中` 1 |
| D | APPLIED_AT | `yyyy-MM-dd HH:mm`（**シートの時間帯での表示。今日 10:12 JST の申請が `01:12` と出る → シートの時間帯は GMT の見込み**。Sheets API で確かめる） |
| E〜H | DEPT_ID・DEPT_NAME・WORKER_ID・WORKER_NAME | |
| I | LEAVE_DATE | `yyyy-MM-dd` |
| J | DAY_TYPE | `全日休` 223・`半日休` 56 |
| K | HALF_TYPE | 空 224・`午後` 46・`午前` 10（全日休なのに値が残る行がある＝画面の癖） |
| L | LEAVE_KIND | `有給消化（私用）` 198・`振替休暇` 79・`特別休暇` 2 |
| M | SUBSTITUTE_FOR_WORK_DATE | `yyyy-MM-dd`（79 件） |
| N〜Q | SPECIAL_REASON_REF・SPECIAL_REASON_TEXT・PAIDLEAVE_DETAIL・DETAIL_FREE | 文字（改行を含むものがある） |
| R | APPROVED_AT | `yyyy-MM-dd HH:mm`（273 件） |
| S・T | SIGN_IMAGE_URL（21 件）・PDF_URL（262 件） | Drive の URL |
| U | CREATED_BY_EMAIL | 全件 `@lineworks-local.info` |
| V | UPDATED_AT | 全行が空 |
| W〜Y | APPROVED_BY1_EMAIL・APPROVED_BY2（名前）・APPROVED_BY2_EMAIL | |
| Z | APPROVED_AT2 | `M/d/yyyy`（en_US の日付の表示。時刻は表示されないだけの見込み。UNFORMATTED で確かめる） |

- 行数 280（うち 1 行は REQ_ID が空で、1 セルだけ値がある半端な行＝シートの 11 行目あたり。移行の道具は REQ_ID が空の行を飛ばし、一覧に残す）。
- **進行中: 申請中 1 件（`LV-20261126-001`、今日の申請）、2 次承認待ち 0 件。** 切り替えの直前にもう一度数える。
- 申請は日中に入り続けている（今日 10:12 にも 1 件）。

**M_SYSTEM_SETTING のキー（値は伏せる。種類だけ）**

`APP_URL`（@106 の /exec）・`PORTAL_URL`（別の GAS の /exec）・`ADMIN_EMAILS`（2 件）・`SOMU_EMAILS`（2 件）・`ALLOW_ADMIN_PAGE_EMAILS`（2 件）・**`MAIL_TO`（2 件。空ではない → 承認完了メールは今も送られている）**・`PDF_FOLDER_ID`（`1MQibR08…`＝見込みどおり）・`PDF_TEMPLATE_SSID`（DB 自身の ID。使っていない）・`MASTER_SOURCE_SSID`（`1iu5Hoak…`）・`MASTER_SOURCE_DEPT_SHEET`（空。下準備に無かったキー）・`MASTER_SOURCE_WORKER_SHEET`（`作業員マスタ`）・`CALENDAR_SOURCE_SSID`（`1Knx_kaQ…`）・`CALENDAR_SOURCE_SHEET`（`社内カレンダーマスタ`）・`LEDGER_SSID`（`1B7tnsT4…`＝見込みどおり）。同じキーの重複は無い。メールのドメインはすべて `lineworks-local.info`。

**ほかのタブ**

| タブ | 行数 | 見出し |
|---|---|---|
| V_LEAVE_SUMMARY | 37 | 年度・部署ID・部署名・作業員ID・作業員名・有給回数（`0.5` 刻みの数）・有給半日回数・振替回数・特別回数・合計回数・最終有給日（表示は `yyyy-MM-dd`。セルの型は Sheets API で確かめる）・警告レベル（6 行に値） |
| M_STAMP | 6 | メール・stampFileId・備考 |
| M_DEPT_APPROVERS | 13 | dept・approverEmails・` approverName`（先頭に空白）・役職・通知メール（5 列） |
| 作業員マスタ | 50 | 作業員コード・氏名・部署・担当業務・拠点・スタッフ種類（**Googleアカウント・在籍フラグの列は無い** → `api_getWorkerInfo` はいつも null。K6 は確定） |
| M_CALENDAR | 130 | 日付（表示 `M/d/yyyy`）・区分（休日 93・祝日 19・出勤土曜 18）・曜日（`Sat` など英語）・備考 |
| M_LOOKUP | 6 | カテゴリ・値・SORT（特別休暇理由: 結婚休暇・忌引休暇・出産休暇・生理休暇・裁判員休暇・その他） |

日付の表示（`M/d/yyyy`、曜日が英語）から、シートの地域は en_US の見込み（残業 DB と同じ）。

### 2-6. Google の準備

- OAuth クライアント **「Time Off Request (VPS)」** を GCP `sebastian01`（imaizumi@lineworks-local.info）に作った。クライアント ID の頭は `253269323289-nafikdg7…`。リダイレクト URI は `https://leave.lineworks-local.net/auth/callback` と `https://leave.lineworks-local.net/auth/google-delegate/callback`。JSON は `~/.secrets/kyuka/google-oauth-client.json`。

## 3. まだ確かめていないこと

- DB のファイルの時間帯・地域（Sheets API の `properties.timeZone`・`locale`）。委任の許可の後に api コンテナから読む。それまでは「GMT・en_US の見込み」。
- T_LEAVE_REQUEST・V_LEAVE_SUMMARY のセルの型（UNFORMATTED_VALUE）。特に APPROVED_AT2 と 最終有給日。
- GAS の「トリガー」「実行数」の画面（このアプリと統合業務メール）。登録されているトリガーの実物と、PDF 1 件の秒数。
- 本番の画面での表示（最終有給日の文字、K5）。

## 4. 追加の決定（2026-11-26 午後。AskUserQuestion の回答。仕様書・計画書と食い違えばここが正）

| # | 項目 | 決定 |
|---|---|---|
| K-8 | 切り替えの日時 | **準備が済んだ日の夜**（18:00 の統合業務メールの後に始め、翌朝 6 時台のマスタ同期の前に終える）。始める前に合図する |
| K-9 | 保守の機能（マスタ同期・集計の作り直し・PDF の作り直し・警告メール・セッションの取り消しなど）を使える人 | **総務部と今泉オーナーだけ**（オーナーの言葉:「アカウントが部署ごとに共通なので、上席者と一般作業員での分け方ができない。総務部と今泉だけが入れるように」）。実装: `permissions.is_operator` ＝ `SOMU_EMAILS` か `ADMIN_EMAILS` か `BOOTSTRAP_ADMINS`（今泉オーナー）のどれかに含まれる。運用の欄は総務管理（somuAdmin）の一番下。※ `ADMIN_EMAILS` を含めたのはディレクターの解釈（今も総務管理を開ける人たちのため） |
| K-10 | 承認待ち一覧の年またぎの並び | **直す**（いつでも新しい申請が上。表示の形 `MM/dd HH:mm` は今のまま） |
| K-11 | 統合業務メールの本文のリンク（M_SYSTEM_SETTING の `APP_URL`） | **切り替えと同時に新しいアドレス `https://leave.lineworks-local.net/` に書き換える**（オーナー:「期限が切れた際の書き換えを忘れそうなので」）。切り替えの手順書に「`APP_URL` の元の値を控える → 書き換える」を入れ、戻す手順に「元の値に書き戻す」を入れる。移転モードの GAS は `APP_URL` を画面の移動に使わない（移転ページは `MIGRATED_TO_URL` を使う）ので、切り替え後の害は無い |

仕様書 15 章のほかの項目（Q1・Q3・Q6・Q9〜Q14）は既定「今のまま」で進める。Q2（2 次承認者はログイン中のメールを使う）と Q5（2 次承認のやり直しを止める）はディレクター判断で「直す」。Q7（最終有給日）と Q16（承認完了メールが届いているか）は実物を見てから。

## 5. Sheets API で読んだ実物（2026-11-26 10:50。委任の許可の後、api コンテナから読むだけ。**確認済み**。3 節の「まだ確かめていないこと」のうち時間帯と型はこれで確定）

- **時間帯 `Etc/GMT`、地域 `en_US`**（`properties.timeZone`・`locale`）。`.env` の `DB_SHEET_TIMEZONE=Etc/GMT` で確定。
- タブ: `SHEET_MAP`・`M_CALENDAR`（131 行）・`作業員マスタ`（51）・`M_SYSTEM_SETTING`・`M_DEPT_APPROVERS`（14）・`M_LOOKUP`・`BatchLogs`・`PDF_CELL_MAP`・`README`・`CONFIG`・`LOOKUP`・`M_DEPT`・`M_WORKER`・`T_LEAVE_REQUEST`（グリッド 948 行 × 26 列、1 行目固定）・`V_LEAVE_SUMMARY`・`PDF_TEMPLATE`・`M_STAMP`（9）・`PDF_MAP`。雛形由来のタブ（SHEET_MAP・PDF_CELL_MAP・README・CONFIG・LOOKUP・M_DEPT・M_WORKER・PDF_TEMPLATE・PDF_MAP・BatchLogs）は触らない。
- **T_LEAVE_REQUEST の行の位置**: 見出しは 1 行目、データは 2 行目から 332 行目まで。**REQ_ID が空の行が途中に 52 行ある**（11 行目から続く）。REQ_ID のある行は 279。先頭に空の行は無い（inv-4 の「先頭に約 500 行」は当たらない）。新しい行は「REQ_ID の列で値のある最後の行（332）の次」に書く。
- **セルの型（UNFORMATTED_VALUE、279 行）**

| 列 | 型 |
|---|---|
| REQ_ID・STATUS・DEPT_ID・DEPT_NAME・WORKER_NAME・DAY_TYPE・LEAVE_KIND・CREATED_BY_EMAIL | 文字 |
| FY_START_YEAR・**WORKER_ID** | **数（int）**。WORKER_ID は文字ではなく数のセル |
| APPLIED_AT・APPROVED_AT・APPROVED_AT2 | 日時のシリアル値（小数つき）。GMT の壁時計（JST − 9 時間）。APPROVED_AT2 も時刻を持つ（表示の形が日付だけ） |
| **LEAVE_DATE** | **270 行が `.625`（＝GMT 15:00＝JST 翌日 0:00）、9 行が整数（＝GMT 0:00＝JST 9:00）**。例: 最後の行は `46352.625` で、表示は `2026-11-26` だが、**休暇日は JST の 2026-11-27**。シートの表示・gviz の日付は 1 日早く見える（2-5 節の gviz の日付はこの意味で読む） |
| SUBSTITUTE_FOR_WORK_DATE | 77 行が `.625`、2 行が整数、200 行が空 |
| HALF_TYPE・SPECIAL_REASON_TEXT・PAIDLEAVE_DETAIL・SIGN_IMAGE_URL・PDF_URL・APPROVED_BY1_EMAIL・APPROVED_BY2・APPROVED_BY2_EMAIL | 文字か空 |
| SPECIAL_REASON_REF・DETAIL_FREE・UPDATED_AT | 全行が空 |

- 文字で書かれた日付の行は無い（inv-4 の「3/28〜4/11 の編集は文字」は、今の実物には残っていない）。
- **V_LEAVE_SUMMARY**: 年度・作業員ID・回数は数（有給回数は 0.5 刻み）。**最終有給日は日付のシリアル値（数のセル）**。GAS はこれを Date で読み `String(Date)` にするので、画面には長い英語の日付で出ている見込み（K5。画面で確かめる）。警告レベルは文字か空。
- M_SYSTEM_SETTING は 15 行（見出し＋14 キー）。重複するキーは無い。

## 6. ステージングの資源（2026-11-26 作成。imaizumi@ 所有。見守りの後に片付ける）

| もの | ID | 中身 |
|---|---|---|
| フォルダ「【ステージング】休暇届 PDF・台帳（消してよい）」 | `1T6ZZ8t-G7fq8BxMUoB7yw2I8i2VXqRCm` | 試験用の PDF・サイン画像の置き場（`STAGING_PDF_FOLDER_ID`）。下の 2 冊もこの中 |
| 「【ステージング】休暇届けapp（消してよい）」 | `1SG4FV9raYuFu8jaWzCVpXvKqo1LPEO7aR44KifWrYqw` | 本番 DB の 2026-11-26 10:55 ごろの複製（`STAGING_SPREADSHEET_ID`）。時間帯 `Etc/GMT`・`en_US`。M_SYSTEM_SETTING の `PDF_FOLDER_ID`・`LEDGER_SSID`・`MAIL_TO`（オーナーの 1 件だけ）を書き換え、T_LEAVE_REQUEST の `SIGN_IMAGE_URL`・`PDF_URL` を空にした。`APP_URL` は写しの GAS をデプロイした後で書き換える |
| 「【ステージング】休暇届台帳（消してよい）」 | `1FpeButODbprm7Bri2WQWlDgwzYDwiwLb8FvWu1XO2YM` | 台帳の複製（`STAGING_LEDGER_SPREADSHEET_ID`） |

写しの GAS（複製に付いたスクリプト）の scriptId とデプロイは、G 担当のコードを入れるときに控える。
