# 一次棚卸し（関数・画面・シート・設定・メール・PDF・権限・時刻・癖）

作成: 2026-11-26。コードは master `216db54`（＝GAS HEAD＝@106）。全行を読んだのはディレクター 1 人で、**計画書の前に 4 人の棚卸しで全行を読み直す**前提の下書き（スキル 01 の 6-2）。
印: **確認済み**＝コードで確かめた、**推測**＝実物（シート・本番画面）で未確認。

## 1. 公開関数（`google.script.run` で呼べるもの 43）

### 1-1. 画面から呼ぶ 27 関数（確認済み）

| 関数（定義） | 引数 | 返すもの | 呼ぶ画面（行） | 権限の確かめ | 書き込み |
|---|---|---|---|---|---|
| `api_getDeptList`（`masters.gs:190`） | — | `[{deptId, deptName}]`（作業員マスタの部署を重なりなしで並べ替え。deptId＝部署名） | top 499・admin 177・somuAdmin 194・leaveStatus 268 | なし | — |
| `api_getWorkersByDept`（`masters.gs:194`） | deptId | `[{workerId, workerName, deptId}]`（コード順） | top 719・leaveStatus 287 | なし | — |
| `api_getWorkerInfo`（`masters.gs:77`） | — | ログイン中のメールと一致する作業員、無ければ `null` | top 494・leaveStatus 490 | なし | — |
| `api_getLookupData`（`masters.gs:202`） | — | 区分・種類の固定値＋M_LOOKUP の `特別休暇理由` | top 507 | なし | — |
| `api_getWorkingDays`（`masters.gs:419`） | — | 今日〜年度末（3/15）の勤務日 `['yyyy-MM-dd']` | top 528（申請のときにもサーバーが呼ぶ `leaveRequest.gs:149`） | なし | — |
| `api_getHolidayDays`（`masters.gs:220`） | workerId（省略可） | 振替元の候補 `[{date, label}]`（3/16〜今日＋7 日の休日。使用済みを除く） | top 680 | なし | — |
| `api_getPastWorkingDays`（`masters.gs:480`） | `'yyyy-MM'` か `'yyyyMM'` | 過去の勤務日 `[{date, dow, label}]` | top 936 | なし | — |
| `api_getUpcomingLeaves`（`leaveRequest.gs:356`） | deptId, workerId | 今日以降の申請（取消以外）`[{reqId, deptName, workerName, leaveKubun, halfDayType, leaveDate:'yyyy/MM/dd', leaveType, status, isRetroactive, daysAgo}]` | top 809 | なし | — |
| `api_getAllPendingRequests`（`leaveRequest.gs:419`） | — | 全員の申請中（`submittedAt:'MM/dd HH:mm'` の新しい順） | top 1244 | なし | — |
| `api_submitLeaveRequest`（`leaveRequest.gs:131`） | data（deptId, deptName, workerId, workerName, leaveKubun, halfDayType, leaveDate, leaveType, substituteDate, specialReason, paidDetail, isRetroactive） | `{ok, reqId}` か `{ok:false, error}` | top 1191 | なし | T_LEAVE_REQUEST に 1 行（ScriptLock 15 秒） |
| `api_editLeaveRequest`（`leaveRequest.gs:852`） | reqId, patch（leaveDate, leaveKubun, halfDayType, substituteDate） | `{ok, newStatus}`。不正は例外 | top 1338 | なし | 行の更新、承認済みなら申請中へ戻す＋PDF ゴミ箱＋台帳の行削除＋集計 |
| `api_deleteLeaveRequest`（`leaveRequest.gs:970`） | reqId | `{ok}`。取消済みは例外 | top 1360 | なし | 取消、PDF ゴミ箱、台帳の行削除、集計 |
| `api_getLeaveRequestById`（`leaveRequest.gs:239`） | reqId | 申請 1 件（日時は `yyyy-MM-dd HH:mm:ss`、`rowNo` を含む）か `null` | preview 193 | なし | —（見つからないと 1.5 秒×4 回待つ） |
| `api_getApproverList`（`masters.gs:163`） | — | M_STAMP の `[{email, name}]` | sign 183 | なし | — |
| `api_approveLeaveRequest`（`leaveRequest.gs:471`） | reqId, signBase64, approverEmail | `{ok, reqId}`。見つからないと例外 | sign 284 | **なし**（誰でも・どの印でも） | 承認済・承認日時・サイン URL・APPROVED_BY1_EMAIL、サイン画像を Drive、集計、承認メール、台帳（ScriptLock 60 秒） |
| `api_isAdmin`（`auth.gs:108`） | — | bool | top 464・somuAdmin 170 | — | — |
| `api_isSomu`（`auth.gs:112`） | — | bool | top 469・admin 180・deptAdmin 182 | — | — |
| `api_isDeptApprover`（`auth.gs:116`） | — | bool | top 461・somuAdmin 167 | — | — |
| `api_getApproverProfiles`（`auth.gs:130`） | — | `{isAdmin, currentEmail, profiles:[{email, name, depts}]}`（管理者は全員分） | deptAdmin 226 | 管理者以外は自分の分だけ | — |
| `api_getDeptAdminView`（`leaveSummary.gs:428`） | fyYear, selectedDepts | `{workers, deptNames}` | deptAdmin 286 | 管理者か部署承認者。承認者は自分の部署だけに絞る | **集計タブを書き直す** |
| `api_getCurrentApproverInfo`（`auth.gs:187`） | — | `{email, name}`（名前は M_STAMP の C 列） | somuAdmin 218 | なし | — |
| `api_getPending2ndApprovals`（`leaveRequest.gs:565`） | — | 承認済で `APPROVED_AT2` が空の一覧（休暇日順） | somuAdmin 232 | 総務か管理者 | — |
| `api_approveLeaveRequest2`（`leaveRequest.gs:619`） | reqId, approver2Email（画面はログイン中のメールを渡す） | `{ok, reqId}` | somuAdmin 328・403（選んだ分は画面が 1 件ずつ順に呼ぶ） | 総務か管理者 | 2次承認の 3 列、前の PDF をゴミ箱、PDF 作成、PDF の 2 列、台帳（ScriptLock 60 秒の中で PDF まで） |
| `api_getSomuAdminView`（`leaveSummary.gs:455`） | fyYear, deptFilter | 作業員ごとの取得状況（期間別つき） | somuAdmin 421 | 総務か管理者 | **集計タブを書き直す** |
| `api_getAdminBossView`（`leaveSummary.gs:155`） | deptId, fyYear | 集計タブの行 | admin 215 | `ALLOW_ADMIN_PAGE_EMAILS` | **集計タブを書き直す** |
| `api_getAdminSomuView`（`leaveSummary.gs:203`） | fyYear, deptFilter | 同上 | admin 208 | 同上 | 同上 |
| `api_getLeaveStatusSearch`（`leaveStatusApi.gs:12`） | fyYear, deptId, workerId | `{mode:'personal', …}` か `{mode:'list', …}` | leaveStatus 313 | **なし（誰でも）** | **集計タブを書き直す** |

### 1-2. 画面から呼ばれないが、誰でも呼べる 16 関数（確認済み）

末尾が `_` でない関数は、ドメインの誰でも `google.script.run` から呼べる（罠 48）。

| 関数 | 何をするか | 危なさ |
|---|---|---|
| `api_approveLeaveRequest2Batch`（`leaveRequest.gs:735`） | 最大 20 件の 2 次承認（ロックの外で PDF） | 画面は使っていない。総務・管理者の確かめはある |
| `api_canAccessAdmin`（`auth.gs:104`）・`api_getDeptApproverDepts`（`auth.gs:120`） | 判定を返す | 低 |
| `doGet`・`onOpen` | 入口・メニュー | — |
| `setupAllSheets`（`menu.gs:117`） | タブと設定のキーを足す | 中（設定の空欄に既定値を書く） |
| `setupAllTriggers`（`menu.gs:345`） | 同期トリガーを作る | 中 |
| `syncAllMasters`・`syncCalendarMaster`・`syncStampMaster`・`syncDeptApprovers`（`syncMasters.gs`） | マスタの全置き換え | 中 |
| `rebuildLeaveSummary`（`leaveSummary.gs:4`） | 集計の書き直し | 低 |
| `checkAndSendWarnMails`（`warnMail.gs:4`） | **送信日の確かめを通り、送信日なら `MAIL_TO` へ警告メール**（コメントは「強制送信」だが、実際は日付で止まる） | 中 |
| `debug_checkLeaveRequestSheet`（`leaveRequest.gs:328`） | 全 REQ_ID を返す | 低（情報が出る） |
| `test_generatePdf`・`test_stampLookup`（`pdfExport.gs:7,24`） | 固定の REQ_ID で PDF を作る・印を調べる | 中（PDF フォルダに書く） |

VPS では閉じたリストにして、これらは出さないか管理者だけにする（04 の Q-B9）。

## 2. 入口（確認済み）

| 入口 | 中身 |
|---|---|
| `doGet`（`menu.gs:3-89`） | 1 節の 7 ページ＋`no_auth`。`page` が無い・不明なら top。`reqId` を preview・sign のテンプレート変数 `REQ_ID` に入れる（GAS の `<?= ?>` が JS の文字として逃がす。VPS では `json.dumps` 相当が要る） |
| `doPost` | 無い |
| 時間主導トリガー | `syncAllMasters` 毎日 6 時台（`menu.gs:381-390`）。ほかはコードからは作られない。実物は【要確認】 |
| シンプルトリガー | `onOpen`（メニュー 8 項目。`migrateLeaveRequestHeaders_` は末尾 `_` なのでメニューから動かない見込み＝推測） |
| ライブラリ | 無い |

## 3. シートとタブ（DB `16v1tQNC…`）

コードが使うタブ（確認済み・`config.gs:4-14`）。実物のタブの一覧・見出し・行数は【要確認】（Sheets API で読む。`05` の C-2）。

| タブ | 見出し（`setupAllSheets` が作る形） | 書くもの | 読むもの | 移行の区分（提案） |
|---|---|---|---|---|
| `T_LEAVE_REQUEST` | REQ_ID・年度・部署ID・部署名・作業員ID・作業員名・休暇区分・半日区分・休暇日・休暇種類・振替元出勤日・特別理由・有給詳細・追加詳細・申請日時・承認日時・承認状態・サイン画像URL・PDF_URL・PDF_FILE_ID・CREATED_BY_EMAIL・APPROVED_BY1_EMAIL・APPROVED_BY2・APPROVED_BY2_EMAIL・APPROVED_AT2（25 列） | 申請・承認・編集・削除 | 全画面・統合業務メール | A（VPS が正。行単位で反映） |
| `V_LEAVE_SUMMARY` | 年度・部署ID・部署名・作業員ID・作業員名・有給回数・有給半日回数・振替回数・特別回数・合計回数・最終有給日・警告レベル | 集計の全書き直し | 管理画面・統合業務メール | A'（VPS が DB から書き直す。行の位置に意味は無い） |
| `M_SYSTEM_SETTING` | KEY・VALUE・MEMO（2 行目から。**同じキーは下の行が勝つ**＝`config.gs:113-117`） | 人が直す。台帳を作ったときだけ GAS が行を足す | 全体・統合業務メール | B（シートが正。VPS は取り込むだけ） |
| `M_LOOKUP` | カテゴリ・値・SORT | 人 | 申請モーダル | B |
| `作業員マスタ` | マスタ元と同じ（作業員コード・氏名・部署・担当業務…） | 同期 | 部署・作業員の一覧 | C（VPS の同期が書く） |
| `M_CALENDAR` | 日付・区分・曜日・備考（元と同じ） | 同期 | 勤務日・休日の計算 | C |
| `M_STAMP` | メール・stampFileId・備考（元と同じ） | 同期 | 承認者の一覧・印・2次承認者の名前 | C |
| `M_DEPT_APPROVERS` | dept・approverEmails・approverName（元の 5 列が来る見込み） | 同期 | 部署承認者の判定 | C |
| `BatchLogs` | 実行日時・バッチ名… | **どこからも書かない**（確認済み・grep） | — | E（触らない） |
| 雛形由来のタブ（README・CONFIG・LOOKUP・M_DEPT・M_WORKER・PDF_TEMPLATE・PDF_MAP など） | DB マスターガイドの形 | — | — | E（残っているかは【要確認】） |

型（確認済み・コード）: 休暇日・振替元出勤日は `new Date('yyyy-MM-ddT00:00:00+09:00')`、申請日時・承認日時・APPROVED_AT2 は `new Date()` を `setValue(s)` で書く＝**日付時刻型のセル**。年度は数値。承認状態・区分は日本語の文字。REQ_ID は `LV-yyyyMMdd-NNN`（その日の最大の連番＋1、`leaveRequest.gs:6-29`）。

見出しは英語でも日本語でも読める（`HEADER_ALIAS_`、`config.gs:70-95`）。本番の見出しがどちらかは【要確認】。統合業務メールのエイリアスは一部の列だけ（`APPROVED_BY1_EMAIL` などは無い）なので、見出しの文字は変えない。

## 4. Settings のキー（M_SYSTEM_SETTING。確認済み・コード、値は【要確認】）

| キー | 使う場所 | 既定 |
|---|---|---|
| `APP_URL` | 全画面のリンク（`menu.gs:6`）、メール本文 | 空（デプロイ後に人が書く） |
| `PORTAL_URL` | ポータルへのボタン（空なら出さない） | 空 |
| `ADMIN_EMAILS` | 管理者 | 空 |
| `SOMU_EMAILS` | 総務（総務管理・2 次承認）。統合業務メールも読む | 空 |
| `ALLOW_ADMIN_PAGE_EMAILS` | 古い管理画面 | 空 |
| `MAIL_TO` | 承認完了通知・有給警告の宛先 | 空 |
| `PDF_FOLDER_ID` | PDF・サイン画像の置き場、台帳を作る場所 | 空 |
| `PDF_TEMPLATE_SSID` | **どこからも使っていない**（PDF はひな形なし） | 空 |
| `MASTER_SOURCE_SSID`・`MASTER_SOURCE_WORKER_SHEET` | 作業員マスタの同期元 | `1iu5Hoak…`・`作業員マスタ` |
| `CALENDAR_SOURCE_SSID`・`CALENDAR_SOURCE_SHEET` | カレンダー・StampMap・DeptApprovers の同期元 | `1Knx_kaQ…`・`社内カレンダーマスタ` |
| `LEDGER_SSID` | 台帳 | 空なら自動で作って行を足す |
| `GENERAL_EMAILS`（設計決定事項に記載） | コードでは使っていない | — |

## 5. メール（確認済み）

`01-dependencies.md` 2 節と `00` の 1-4 を参照。このアプリが今送り得るのは「承認完了通知」（1 次承認のたび、`MAIL_TO` が空でなければ）と「有給警告」（メニューから手で、送信日だけ）の 2 種。件名・本文は `mail.gs:20-68`、`warnMail.gs:52-163`。本文のリンクは `APP_URL?page=admin`。差出人は imaizumi@。

## 6. PDF・サイン・台帳（確認済み）

| 項目 | 中身 |
|---|---|
| 作る場面 | 2 次承認のとき（1 件ずつ、一括でも）。1 次承認・申請では作らない |
| 作り方 | `SpreadsheetApp.create('TMP_PDF_<REQ_ID>_<日時>')` → `buildLeavePdfSheet_`（A1:J30、列幅・行高・結合・色 `PDF_COLORS_`・Noto Sans JP）→ 印を `insertImage`（サムネイル `drive.google.com/thumbnail?id=…&sz=w300` か `lh3.googleusercontent.com/d/…=s300` を `UrlFetchApp` で、OAuth トークン付き）→ `Utilities.sleep(4000)`＋`sleep(1000)` → `docs.google.com/spreadsheets/d/<id>/export?format=pdf&size=A4&portrait=true&fitw=true&…` → 日付フォルダへ保存 → 一時シートをゴミ箱（`pdfExport.gs:124-180, 187-542, 663-698, 732-757`） |
| 中身 | 届出日（令和）・部署・氏名・休暇日（令和と曜日）・時間（終日 8:05〜17:10、午前 8:05〜12:15、午後 13:00〜17:10）・区分・種類・振替元・理由（特別理由／有給詳細か「私用のため」／「振替休暇」）・承認印（所属長＝1 次、部長＝2 次）・承認日（令和の年/月/日の短い形） |
| 前の PDF | 2 次承認・編集（承認済みから戻すとき）・取消で `PDF_FILE_ID` のファイルをゴミ箱へ |
| サイン | canvas の PNG（`toDataURL`）を base64 で送る。電子印を選んだときは送らない。`sign_<REQ_ID>.png` を休暇日のフォルダへ（`pdfExport.gs:89-104`） |
| 台帳 | `1B7tnsT4…` の「<年度>年度」タブ（無ければ先頭に作る）。20 列（管理番号・作業員ID・部署・氏名・休暇日・曜日・休暇区分・半日区分・勤務時間帯・休暇種類・理由・振替元・申請日・1次承認日・1次承認者（**メール**）・2次承認日・2次承認者（名前）・承認状態・有給累計取得回数・PDF（HYPERLINK 式））。偶数行の背景色。行の削除あり（`ledger.gs`） |

## 7. 権限（確認済み）

`00` の 1-3 を参照。追加の事実:
- 1 次承認（サイン）は**入口も関数も権限の確かめが無い**。承認者の選択肢は M_STAMP の全員で、誰の印でも選べる（要件「承認者の特定は不要」どおり）。
- 2 次承認者は画面がログイン中のメールを渡すが、サーバーは引数をそのまま使う（`approver2Email` を書き換えれば別の人の印になる。総務・管理者に限られる）。
- 編集・削除は誰でも、どの申請でも（申請者の確かめなし）。
- 取得状況（leaveStatus）は誰でも、どの人でも見られる（実装手順書の決め）。
- Google アカウントは部署で共有の見込み（推測）。`api_getWorkerInfo` が人を自動で選ぶのは、作業員マスタに「Googleアカウント」列があるときだけ。マスタ元には無い（スキル 00 の注意）ので、**いつも null を返し、自動選択は働いていない見込み**（推測。初日に見出しを見る）。

## 8. 日付と時間帯（確認済み・コード）

- `appsscript.json` の `timeZone: Asia/Tokyo`。`fmtDate_` は Asia/Tokyo。`new Date(y, m, d)`・`getMonth()`・`setHours(0,0,0,0)` もスクリプトの時間帯（JST）で動く。VPS ではすべて JST の関数に置き換える（残業の決定 C）。
- 年度: 3/16〜翌 3/15（`computeFiscalYear_`）。月次期間: 前月 16 日〜当月 15 日。有給の期限 5 つ（6/30・8/31・12/31・2/15・3/15）と期間 5 つ（3/16-6/30 など）。警告メールの送信日 7/1・9/1・1/1・2/16。
- 祝日: M_CALENDAR に載っている日はそれに従い、載っていない日は日本の祝日を計算（春分・秋分は近似式、振替休日、国民の休日。`masters.gs:320-413`）。
- 「今日」はサーバー側（GAS の JST）とブラウザの両方で決める箇所がある（年度の選択肢はブラウザ。`somuAdmin.html:173-176`）。
- シートのファイルの時間帯は【要確認】。残業 DB は `Etc/GMT` だった（罠 68）。日付型のセルの壁時計がずれて見えても GAS は正しく読む。VPS が値を書くときはシートの時間帯でシリアル値にする。

## 9. 癖・不具合の候補（直すか今のままかはオーナーに聞く）

| # | 中身 | 根拠 | 確かさ | 確かめ方 |
|---|---|---|---|---|
| K1 | 1 次承認が状態を見ない。取消・却下・承認済みの申請もサインで承認済にできる。2 次承認済みをもう一度 1 次承認すると `APPROVED_AT2` が残ったまま承認日時と印が変わり、PDF は作り直されない | `leaveRequest.gs:471-557`（状態の確かめなし）、`sign.html` も見ない | 確認済み（コード） | ステージングで試す（本番では試さない） |
| K2 | 2 つのタブで同じ申請を同時に承認・2 次承認できる（段の照合なし）。二重押しで申請が 2 件できる（冪等キーなし。ボタンを止めるだけ） | `leaveRequest.gs:131`、`top.html:1145-1146` | 確認済み（コード） | — |
| K3 | 集計（V_LEAVE_SUMMARY）は 1 次承認だけの申請も数える。半日の有給は集計で 0.5、期間別（quarters）では 1、台帳の「有給累計取得回数」でも 1 | `leaveSummary.gs:66-72`、`217-276`、`ledger.gs:275-308` | 確認済み（コード） | 計画書で「今のまま」を確かめる |
| K4 | 画面を開くだけで集計タブを書き直す（取得状況は誰でも開ける）。同時に開くと書き込みが重なる | `leaveSummary.gs:162, 288` | 確認済み（コード） | — |
| K5 | 「最終有給日」が長い日付の文字（`Fri Nov 27 2026 00:00:00 GMT+0900 …`）で出るおそれ。集計タブに `'yyyy-MM-dd'` の文字を書くとシートが日付に変え、読むときに `String(Date)` になる | `leaveSummary.gs:75, 108, 188, 311` | **推測** | 本番の総務管理・部署管理・取得状況の「最終有給日」を見る（`05` の C-10） |
| K6 | `api_getWorkerInfo` が人を自動で選ばない（作業員マスタに Googleアカウント列が無い） | `masters.gs:89-94` | 推測 | 作業員マスタの見出しを見る |
| K7 | 存在しない reqId のプレビューは約 6 秒待ってから「見つかりません」 | `leaveRequest.gs:243-248` | 確認済み（コード） | VPS ではすぐ返す（画面の文言は同じ） |
| K8 | `checkAndSendWarnMails` のコメントは「日付チェックをスキップして強制送信」だが、実際は送信日以外は何もしない | `warnMail.gs:3-25` | 確認済み（コード） | — |
| K9 | 台帳の「1次承認者」はメール、「2次承認者」は名前 | `ledger.gs:184, 192` | 確認済み（コード） | 今のまま |
| K10 | 申請の「追加詳細」は画面にも保存にも無い（列だけある） | `leaveRequest.gs:206-221`、`top.html:1148-1161` | 確認済み（コード） | 今のまま |
| K11 | 古い管理画面（`?page=admin`）は「最終有給日」だけエスケープせずに表に入れる（部署名・氏名は `esc` で逃がしている） | `admin.html:264` | 確認済み（コード） | VPS では逃がす（見た目は同じ） |
| K12 | 2 次承認の一括は、画面が 1 件ずつ順に呼ぶ（1 件 15 秒ほど）。`api_approveLeaveRequest2Batch` は使っていない | `somuAdmin.html:349-403` | 確認済み（コード） | — |
| K13 | 台帳が開けないと、黙って新しい台帳を作り Settings に `LEDGER_SSID` の行を足す（下の行が勝つ） | `ledger.gs:7-45` | 確認済み（コード） | 内部用の入口では止める |
| K14 | 2 次承認の PDF で、承認日の「年」が令和の数だけ（`8/11/27` の形） | `pdfExport.gs:432, 527` | 確認済み（コード） | 今のまま |

## 10. GAS 特有の仕組み（移すときに置き換えるもの）

- `HtmlService.createTemplateFromFile` とテンプレート変数 `APP_URL`・`PORTAL_URL`・`REQ_ID`・`message`。`setTitle`。`ALLOWALL`。`<base target="_top">`、`window.top.location.href`。
- `google.script.run.withSuccessHandler(…).withFailureHandler(…).api_xxx(…)`（画面ごとの on 属性は合計 50 個: top 27・somuAdmin 7・admin 5・sign 5・deptAdmin 3・preview・leaveStatus・no_auth 各 1）。
- サインのページ: `meta refresh` で 1 秒後にプレビューへ。canvas は pointer イベント、`user-scalable=no`。
- `LockService.getScriptLock` → DB の行ロック。`SpreadsheetApp.flush` と読み直しの待ち → 不要。
- `DriveApp`・`UrlFetchApp`（サムネイル・export）・`SpreadsheetApp.create`・`GmailApp` → 今の GAS に残す（決定 B 相当）。
