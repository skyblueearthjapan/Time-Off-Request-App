# 棚卸し 4: データ・マスタ・運用と、これまでの経緯

作成: 2026-11-26。担当: 棚卸し 4（データ・マスタ・運用・経緯）。読むだけで作った（本番・GAS・外部には触れていない）。
対象: `repo/src/`（master `216db54`）、`git log`（104 コミット）、`repo/docs/` の全資料、`gas-pulled/`、残業の前例（`row_sync.py` ほか）、罠の一覧。

> 印: **確認済み**＝コード・git で確かめた。**推測**＝実物（本番シート・トリガー画面）で未確認。値が分からないものは 〇〇、確かでないものは【要確認】。
> 行番号は `src/` のファイル。`gs` は省いて `leaveRequest:131` のように書く。

## 0. 先に読む要点

| # | 要点 | 確かさ | 影響 |
|---|---|---|---|
| 1 | 本番の T_LEAVE_REQUEST は**雛形 xlsx 由来の英語の見出し**で始まった（`769a2bd` の本文「実際のヘッダが英語(STATUS, LEAVE_DATE等)」）。雛形の 22 列には `PDF_FILE_ID` も 2 次承認の 4 列も無い。2 次承認の 4 列は日本語名で足された見込み（`aad3ec9`「日本語ヘッダー名しかない場合」） | 確認済み（git）／今の見出しは推測 | 下書きの「25 列（setupAllSheets の形）」は実物と違う見込みが高い。`PDF_FILE_ID` が無ければ、古い PDF は一度もゴミ箱へ移っていない |
| 2 | 雛形の T_LEAVE_REQUEST は 501 行あり、`c04fdd6`「空行500件表示」を直した。**本物の行は 502 行目以降にある見込み**（新しい行は `getLastRow()+1` に書く。`leaveRequest:225`） | 確認済み（git）／行番号は推測 | 移行の道具は REQ_ID が空の行を飛ばし、行番号（sheet_row）は元のまま持つ |
| 3 | 休暇日・振替元出勤日の**時刻が時期で違う**。休暇日は 4/10 18:35 より前、振替元は 4/11 10:57 より前に `new Date('yyyy-MM-dd')`（UTC 0 時＝JST 9 時）で書いた。3/28〜4/11 の編集は休暇日を**文字**で `setValue` した | 確認済み（git `2c22068`・`3b8a1f7`・`1238fba`） | 型の表は行ごとに持つ。照合で「9 時」と「0 時」を同じ日として扱う |
| 4 | 年度の境目は 3/11 19:24 に 4/1→3/16 へ変わったが、**それまでに書いた行の `年度` は直していない**。集計・部署管理は行の `年度` をそのまま使う | 確認済み（`c85813d`、`leaveSummary:37,334`） | 移行では `年度` を計算し直さず、そのまま移す。食い違いは報告に出す |
| 5 | マスタ 4 つは「消してから書く」。元が**見出しだけ**だと、DB のタブも見出しだけになる。M_STAMP・M_DEPT_APPROVERS は**列の位置**（A・B・C）で読む | 確認済み（`syncMasters:300-323`、`auth:85-86`） | VPS の同期は「空なら上書きしない」に変える（前例どおり）。列の順の約束を残業側と共有する |
| 6 | 3/4〜3/11 の間は `setupAllTriggers` が有給警告（8 時）と通知 3 本も作っていた。消す処理は無い。**古いトリガーが残っている恐れ** | 確認済み（git `0b1a657`〜`5ce6954`）／登録の有無は推測 | 初日の C-9 で、`syncAllMasters` 以外のトリガーも探す。切り替えで全部止める |

## 1. タブごとの棚卸し

コードが使うタブは `config:4-14` の 9 つ。どのタブも見出しは 1 行目、データは 2 行目から。

### 1-1. T_LEAVE_REQUEST（申請。A 区分）

見出しの読み方（確認済み・`config:135-176`）:
- 1 行目の全列を読み、**値がある最後の列まで**を見出しにする（`getMaxColumns` まで読んで後ろの空を落とす）。各見出しは `normalize_`（前後の空白を取り、全角英数を半角に。`config:128-133`）を通す。
- `idx[見出し]` に加え、英語→日本語（`HEADER_ALIAS_`）と日本語→英語（逆引き。`aad3ec9` で追加）の両方を登録する。同じ名前が先にあれば上書きしない。
- 読むときは `Math.max(見出しの数, getLastColumn())` 列を読む（`leaveRequest:273` ほか）。

`HEADER_ALIAS_` の全部（`config:70-95`）と、コードが使う名前・書き込みの有無:

| 英語の見出し | 日本語の別名 | コードが使う名前 | 書くか |
|---|---|---|---|
| （なし） | （なし） | `REQ_ID` | 書く |
| `FY_START_YEAR` | `年度` | 年度 | 書く |
| `STATUS` | `承認状態` | 承認状態 | 書く |
| `APPLIED_AT` | `申請日時` | 申請日時 | 書く |
| `DEPT_ID` | `部署ID` | 部署ID | 書く |
| `DEPT_NAME` | `部署名` | 部署名 | 書く |
| `WORKER_ID` | `作業員ID` | 作業員ID | 書く |
| `WORKER_NAME` | `作業員名` | 作業員名 | 書く |
| `LEAVE_DATE` | `休暇日` | 休暇日 | 書く |
| `DAY_TYPE` | `休暇区分` | 休暇区分 | 書く |
| `HALF_TYPE` | `半日区分` | 半日区分 | 書く |
| `LEAVE_KIND` | `休暇種類` | 休暇種類 | 書く |
| `SUBSTITUTE_FOR_WORK_DATE` | `振替元出勤日` | 振替元出勤日 | 書く |
| `SPECIAL_REASON_REF` | `特別理由参照` | （使わない） | **書かない** |
| `SPECIAL_REASON_TEXT` | `特別理由` | 特別理由 | 書く |
| `PAIDLEAVE_DETAIL` | `有給詳細` | 有給詳細 | 書く |
| `DETAIL_FREE` | `追加詳細` | 追加詳細 | **今は書かない**（3/7 `cbb80bf` まで書いていた）。読むだけ（`leaveRequest:303`、画面では使わない） |
| `APPROVED_AT` | `承認日時` | 承認日時 | 書く |
| `SIGN_IMAGE_URL` | `サイン画像URL` | サイン画像URL | 書く |
| （なし） | （なし） | `PDF_URL` | 書く |
| （なし） | （なし） | `PDF_FILE_ID` | 書く（列があるときだけ） |
| `CREATED_BY_EMAIL` | `作成者メール` | 作成者メール | 書く |
| `UPDATED_AT` | `更新日時` | （使わない） | **書かない** |
| `APPROVED_BY1_EMAIL` | `1次承認者メール` | APPROVED_BY1_EMAIL | 書く |
| `APPROVED_BY2` | `2次承認者名` | APPROVED_BY2 | 書く |
| `APPROVED_BY2_EMAIL` | `2次承認者メール` | APPROVED_BY2_EMAIL | 書く |
| `APPROVED_AT2` | `2次承認日時` | APPROVED_AT2 | 書く |

見出しの形は 3 通りあり得ます（どれが本番かは【要確認】。初日の C-2）:

| 形 | 列 | 根拠 |
|---|---|---|
| ア. 雛形（英語 22 列）＋後から足した列 | A `REQ_ID`・B `FY_START_YEAR`・C `STATUS`・D `APPLIED_AT`・E `DEPT_ID`・F `DEPT_NAME`・G `WORKER_ID`・H `WORKER_NAME`・I `LEAVE_DATE`・J `DAY_TYPE`・K `HALF_TYPE`・L `LEAVE_KIND`・M `SUBSTITUTE_FOR_WORK_DATE`・N `SPECIAL_REASON_REF`・O `SPECIAL_REASON_TEXT`・P `PAIDLEAVE_DETAIL`・Q `DETAIL_FREE`・R `APPROVED_AT`・S `SIGN_IMAGE_URL`・T `PDF_URL`・U `CREATED_BY_EMAIL`・V `UPDATED_AT`、W 以降に 2 次承認の 4 列（英語名か日本語名）と、もしかすると `PDF_FILE_ID` | `docs/休暇届アプリ_最強DBマスターガイド.md` 3 節、`769a2bd`・`aad3ec9` の本文。**本番はこの形の見込みが高い**（推測） |
| イ. `setupAllSheets` が作る日本語 25 列 | `menu:259-266` | 空のタブのときだけ作る。本番は雛形を取り込んだので当たらない見込み |
| ウ. 申請時の緊急生成 25 列 | `leaveRequest:183-190`（イと同じ並び） | 見出しが空のときだけ |

位置で読む箇所（確認済み）:

| 箇所 | 読み方 | 前提 |
|---|---|---|
| `generateReqId_`（`leaveRequest:16`） | **A 列**の 2 行目以降 | REQ_ID が A 列 |
| `rebuildLeaveSummary_`・`buildAdminViewData_`（`leaveSummary:31,328`） | REQ_ID の列が無ければ **A 列**で空行を判定 | 同上 |
| `api_getLeaveRequestById`（`leaveRequest:260-267`） | REQ_ID が無ければ見出しに `REQ` を含む最初の列 | — |
| `migrateLeaveRequestHeaders_`（`menu:356-379`） | 最後の列の右に英語名で 4 列足す | 英語名が無いとき（日本語名があっても足す＝重なる恐れ） |
| ほかはすべて | 見出しの文字（別名つき） | — |

### 1-2. ほかのタブ

| タブ | 見出し（コードが前提にするもの） | 読み方 | 誰が書く | 誰が読む |
|---|---|---|---|---|
| `V_LEAVE_SUMMARY` | `年度・部署ID・部署名・作業員ID・作業員名・有給回数・有給半日回数・振替回数・特別回数・合計回数・最終有給日・警告レベル`（12 列。書き直すたびに**この見出しを書く**＝雛形の英語 11 列は残っていない。確認済み `leaveSummary:94-97`） | 見出しの文字 | `rebuildLeaveSummary_` だけ（全消し→全書き） | 画面 4 つ（admin・deptAdmin・somuAdmin・leaveStatus）、`warnMail`、統合業務メール |
| `M_SYSTEM_SETTING` | `KEY・VALUE・MEMO`。A＝キー、B＝値（位置で読む。`config:113-116`） | 位置 | 人。GAS は `setupAllSheets`（キーを足す・空に既定を書く）と台帳の自動作成（行を足す） | 全体、統合業務メール |
| `M_LOOKUP` | `カテゴリ・値・SORT` | 見出しの文字（`config:236-239`） | 人（`setupAllSheets` が空のときだけ 6 行を入れる） | 申請モーダル（`特別休暇理由` だけ） |
| `作業員マスタ` | 元と同じ。コードが探すのは `作業員コード`（無ければ `作業員ID`）・`氏名`・`部署`（無ければ `部署ID`）・`在籍フラグ`・`Googleアカウント` を含む列 | 見出しの文字（`masters:23-33,84-93`） | `syncAllMasters` | 部署・作業員の一覧、`api_getWorkerInfo` |
| `M_CALENDAR` | 元（残業 DB の `社内カレンダーマスタ`）と同じ。コードは `日付`・`区分` だけ | 見出しの文字（`masters:227-229,424-426,513-515`） | `syncCalendarMaster` | 勤務日・休日・過去勤務日 |
| `M_STAMP` | 元（`StampMap`）と同じ。A メール・B stampFileId・C 名前（備考） | **位置**（`masters:145-146,175-176`、`auth:195-196`、`leaveRequest:657-658,748-749`） | `syncStampMaster` | 承認者の一覧・電子印・2 次承認者の名前 |
| `M_DEPT_APPROVERS` | 元（`DeptApprovers`）と同じ。残業では 5 列（A dept・B approverEmails・C ` approverName`（先頭に空白）・D 役職・E 通知メール）なので、同じ 5 列が写っている見込み（推測。残業の計画書 4-2） | **位置**（`auth:85-86,140-142`、`notifyMail:24-26`） | `syncDeptApprovers` | 部署承認者の判定・部署管理 |
| `BatchLogs` | `実行日時・バッチ名・対象日・成功・スキップ・失敗・エラー詳細` | — | **どこからも書かない**（`SHEET.BATCH_LOGS` は定義だけ） | なし |
| 雛形由来（README・CONFIG・LOOKUP・M_DEPT・M_WORKER・PDF_TEMPLATE・PDF_MAP） | ガイドの形 | — | なし | なし（`28aa79d` で一度 PDF_TEMPLATE を名前で読んだが、`c874e11` で使わなくなった） |

### 1-3. 列の型（コードが書く値）

| 列 | 書く値 | 空の表し方 | 補足 |
|---|---|---|---|
| REQ_ID | 文字 `LV-yyyyMMdd-NNN`（申請した日の JST、その日の最大＋1、3 桁ゼロ埋め） | — | 取消の行も数える。1000 件目は 4 桁 |
| 年度 | 数値（`computeFiscalYear_` の戻り） | — | 3/11 より前は 4/1 区切りで計算した値 |
| 部署ID・部署名 | 文字（どちらも部署名。`masters:11-13`） | `''` | — |
| 作業員ID | 文字のつもり（マスタの `作業員コード` を `normalize_`） | `''` | `setValues` は数字だけの文字を数にする。先頭 0 は消える（推測。7 節） |
| 作業員名 | 文字 | `''` | — |
| 休暇区分・半日区分・休暇種類・承認状態 | 決まった日本語（`config:16-40`） | 半日区分は `''` | 承認状態は `申請中`・`承認済`・`取消`。`却下` は定義だけで書かない |
| 休暇日・振替元出勤日 | Date（今は `yyyy-MM-ddT00:00:00+09:00`） | 振替元は `''` | 古い行は JST 9 時、文字の行もあり得る（0 節 #3） |
| 申請日時・承認日時・APPROVED_AT2 | Date（`new Date()`） | `''` | 承認日時・APPROVED_AT2 は編集で `''` に戻る |
| 特別理由・有給詳細 | 自由入力の文字 | `''` | `1/2` のような文字は日付になる恐れ（罠 11） |
| サイン画像URL・PDF_URL | Drive の URL 文字 | `''` | — |
| PDF_FILE_ID | Drive のファイル ID 文字 | `''` | — |
| 作成者メール・APPROVED_BY1_EMAIL・APPROVED_BY2_EMAIL | メール文字 | `''` | 手書きサインの承認は APPROVED_BY1_EMAIL が `''`（`sign.html:205,284`） |
| APPROVED_BY2 | M_STAMP C 列の名前、無ければメール | — | — |

V_LEAVE_SUMMARY の型（確認済み・`leaveSummary:99-111`）: 年度は T_LEAVE_REQUEST の値をそのまま（数値の見込み）。回数は数値（有給回数は半日で 0.5 刻み）。作業員ID は文字を書くが数字だけなら数になる（推測）。最終有給日は `yyyy-MM-dd` の文字を書くので、シートが日付に変える見込み（推測。下書きの K5）。行の並びは `年度_作業員ID` の**文字の順**。

## 2. T_LEAVE_REQUEST の列ごとの書き手と、VPS との列の分け方

### 2-1. 書く関数（確認済み）

| 列 | 申請 `api_submitLeaveRequest`（131） | 1 次承認 `api_approveLeaveRequest`（471） | 2 次承認 `…2`（619）・一括（735） | 編集 `api_editLeaveRequest`（852） | 取消 `api_deleteLeaveRequest`（970） |
|---|---|---|---|---|---|
| REQ_ID | 書く（206） | — | — | — | — |
| 年度 | 書く（207） | — | — | 休暇日を変えたとき再計算（911-912） | — |
| 部署ID・部署名・作業員ID・作業員名 | 書く（208-211） | — | — | — | — |
| 休暇区分 | 書く（212） | — | — | 変えたとき（917） | — |
| 半日区分 | 書く（213。全日休でも画面の選択が残れば入る恐れ＝推測） | — | — | 全日休なら `''`、半日で値があれば上書き（918-924） | — |
| 休暇日 | 書く（214） | — | — | 変えたとき（910） | — |
| 休暇種類 | 書く（215） | — | — | — | — |
| 振替元出勤日 | 値が来れば**種類に関係なく**書く（216） | — | — | 振替休暇で値が来たとき（928-930） | — |
| 特別理由・有給詳細 | 書く（217-218。種類に関係なく画面の値） | — | — | — | — |
| 申請日時 | 書く（219） | — | — | — | — |
| 承認状態 | `申請中`（220） | `承認済`（517。**今の状態を見ない**） | 見るだけ（`承認済` でなければ止める） | 承認済なら `申請中`（936） | `取消`（984） |
| 作成者メール | 書く（221） | — | — | — | — |
| 承認日時 | — | `now`（518） | — | 承認済なら `''`（938） | — |
| サイン画像URL | — | URL か `''`（519。いつも上書き） | — | 承認済なら `''`（940） | — |
| APPROVED_BY1_EMAIL | — | 選んだメールか `''`（521-524） | 読むだけ | 承認済なら `''`（939） | — |
| APPROVED_BY2・BY2_EMAIL・AT2 | — | — | 書く（673-675、791-793） | 承認済なら `''`（941-943） | — |
| PDF_URL・PDF_FILE_ID | — | （4/10 まで: 書いた） | PDF の後に書く（706-707、830-831） | 承認済で PDF_FILE_ID があれば `''`（949-950） | PDF_FILE_ID があれば `''`（990-991） |

後から上書きされる場面のまとめ:
- 1 次承認は何度でも上書きする（取消・2 次承認済みの行も `承認済` に戻り、承認日時・サイン・1 次承認者が変わる。APPROVED_AT2 と PDF は残る）。
- 編集で承認が外れると、承認の 7 列と PDF の 2 列を空にする。取消は PDF の 2 列だけ空にし、承認の列は残す。
- 2 次承認は「前の PDF をゴミ箱→作る→2 列を書く」。PDF が失敗しても 2 次承認の 3 列は残る（`leaveRequest:709-711`）。

### 2-2. 列の分け方の案（提案）

前例（残業 `row_sync.py:8-9,550-551`）は「最後にシートへ書いた値（synced_values）と違う列だけを書く」ことで、GAS が書く列を古い値で上書きしない作りです。このアプリでは次のように分けるのが安全です。

| 区分 | 列 | 書き手 | VPS の扱い |
|---|---|---|---|
| V（VPS が正） | REQ_ID・年度・部署ID・部署名・作業員ID・作業員名・休暇区分・半日区分・休暇日・休暇種類・振替元出勤日・特別理由・有給詳細・申請日時・承認日時・承認状態・作成者メール・APPROVED_BY1_EMAIL・APPROVED_BY2・APPROVED_BY2_EMAIL・APPROVED_AT2 | VPS の行の反映 | 差分の列だけ書く |
| G（GAS が正） | サイン画像URL・PDF_URL・PDF_FILE_ID | GAS の内部用の入口だけ | **行の反映の対象から外す**。GAS の返事の値を DB に入れ、synced_values も同じ値にそろえる（書き戻さない） |
| N（誰も書かない） | SPECIAL_REASON_REF・DETAIL_FREE・UPDATED_AT・決まりに無い列 | なし | 移行で値を持つだけ。シートに書かない |

G 列を空にする場面も GAS に任せるのがよいと考えます。理由は次のとおりです。
- 編集で承認が外れたとき・取消のときは、PDF をゴミ箱へ移す必要があり、それは GAS でしかできません。
- 内部用の入口に「PDF を捨てて 2 列を空にし、台帳の行を消す」仕事（例: `internal.afterUnapprove`・`internal.afterCancel`）を置けば、G 列の書き手が GAS の 1 か所に限られ、取り合いが起きません。
- サイン画像URL を空にするのも同じ仕事の中で行います。

順序の守り（確認済みの依存）:
- `generateLeavePdf_`（`pdfExport:125`）・`updateLedger_`（`ledger:127`）・`sendApprovalNotification_`（`mail:14`）は、**シートの行**を `api_getLeaveRequestById` で読み直して使います。GAS の仕事は、その行の反映が済んでから呼ぶ必要があります（前例の `requires`）。
- 台帳の「有給累計取得回数」（`ledger:275-308`）は**同じ作業員のほかの行**も数えます。ほかの申請の反映が遅れていると 1 つずれます。台帳の仕事の前に「待ちの行が 0 件」を確かめるか、ずれを受け入れるかを仕様で決めます。
- 1 次承認の GAS の仕事（サイン画像の保存）が終わる前に編集で承認が外れた場合、後から URL を書いてしまいます。同じ申請の仕事は積んだ順に 1 本ずつ動かします（前例と同じ）。

## 3. M_SYSTEM_SETTING のキー

### 3-1. キーの全部（確認済み）

| キー | 使う場所 | 既定値（`setupAllSheets`） | コードの予備の値 |
|---|---|---|---|
| `APP_URL` | 画面のリンク（`menu:6`）、メール本文（`mail:37`、`warnMail:103`、`notifyMail:89,187,284`） | 空 | — |
| `PORTAL_URL` | ポータルのボタン（`menu:7`） | 空 | — |
| `ADMIN_EMAILS` | `isAdmin_`（`auth:21`） | 空 | — |
| `SOMU_EMAILS` | `isSomu_`（`auth:33`）、`getSomuEmails_`（`notifyMail:9`。今は呼ばれない） | 空 | — |
| `ALLOW_ADMIN_PAGE_EMAILS` | `canAccessAdmin_`（`auth:9`）＝`?page=admin` と `api_getAdminBossView`・`…SomuView` | 空 | — |
| `MAIL_TO` | 承認完了通知（`mail:8`）、有給警告（`warnMail:46`） | 空 | — |
| `PDF_FOLDER_ID` | PDF（`pdfExport:129`）、サイン（`pdfExport:91`）、台帳の作り場所（`ledger:21`）、通知本文（`notifyMail:285`） | 空 | — |
| `PDF_TEMPLATE_SSID` | **使っていない**（`c874e11` でひな形をやめた） | 空 | — |
| `MASTER_SOURCE_SSID` | 作業員マスタの元（`syncMasters:19`） | `1iu5Hoak…` | 空なら `1iu5Hoak…`（`syncMasters:21-24`） |
| `MASTER_SOURCE_WORKER_SHEET` | 同上のタブ名（`syncMasters:34`） | `作業員マスタ` | 空なら `作業員マスタ` |
| `CALENDAR_SOURCE_SSID` | カレンダー・**StampMap・DeptApprovers も**この ID（`syncMasters:127,173,215`） | `1Knx_kaQ…` | 空なら `1Knx_kaQ…` |
| `CALENDAR_SOURCE_SHEET` | カレンダーのタブ名（`syncMasters:132`） | `社内カレンダーマスタ` | 空なら `社内カレンダーマスタ`。StampMap・DeptApprovers のタブ名は固定 |
| `LEDGER_SSID` | 台帳（`ledger:9`） | 空 | 開けなければ新しく作って行を足す |

docs にあってコードに無いキー: `GENERAL_EMAILS`（要件定義・設計決定事項）、`MASTER_SOURCE_DEPT_SHEET`（設計決定事項）、`MAIL_FROM_NAME`・`FY_START_MONTH`・`FY_START_DAY`・`WARN_1..5_*`・`WARN_MAIL_TIMING`（DB マスターガイドの CONFIG）。どれも読まれません（確認済み・grep）。

### 3-2. 「同じキーは下の行が勝つ」の実装（確認済み・`config:107-119`）

- 2 行目から下へ順に読み、`map[key] = val` で上書きします。**値が空の行でも上書きします**（`if (key)` だけを見る）。下に空の行があると、上の値が消えます。
- キーは `String(...).trim()` だけ（全角→半角はしない）。値は加工しません。使う側がたいてい `normalize_` を通すので、値の全角英数は半角になります。
- `setupAllSheets`（`menu:150-171`）は既存のキーを最後の行で覚え、キーが無ければ末尾に足し、「既定値があり、今の値が空」なら値と MEMO を上書きします。既定値があるのは同期元の 4 キーだけです。
- 台帳が開けないと `getLedgerSs_`（`ledger:35-38`）が末尾に `LEDGER_SSID` の行を足します。下の行が勝つので新しい台帳が使われます。**`LEDGER_SSID` の行が 2 つ以上ある恐れ**があります（推測。`setupAllSheets` が先に空の行を作った場合など）。
- VPS の取り込みも同じ規則（空の値でも下が勝つ）にし、重なったキーは報告に出します（前例 `config_sync.py`）。ステージングで書き換えるときは**一番下の行**を書き換えます。

## 4. マスタ同期 4 つ

| 項目 | 作業員マスタ | M_CALENDAR | M_STAMP | M_DEPT_APPROVERS |
|---|---|---|---|---|
| 関数 | `syncAllMasters` の中のループ（`syncMasters:30-82`） | `syncCalendarMaster`（121-165） | `syncStampMaster`（171-207） | `syncDeptApprovers`（213-249） |
| 元 | `MASTER_SOURCE_SSID` の `MASTER_SOURCE_WORKER_SHEET` | `CALENDAR_SOURCE_SSID` の `CALENDAR_SOURCE_SHEET` | `CALENDAR_SOURCE_SSID` の `StampMap`（固定） | `CALENDAR_SOURCE_SSID` の `DeptApprovers`（固定） |
| ロック | `syncAllMasters` が 30 秒 | 自分で 30 秒（`syncAllMasters` の中からも呼ぶ） | なし | なし |
| 読む範囲 | 1 行目の値がある最後の列まで × `getLastRow` 行（`syncMasters:252-270`）。失敗したら 2 秒待って再読、さらに 2 秒待って 50 行ずつ | 同じ | 同じ | 同じ |
| 行の捨て方 | A 列が空の行を捨てる（`0` も空扱い。見出しは残す。`syncMasters:326-335`） | 同じ | 同じ | 同じ |
| 書き方 | `clearContents`（書式は残る）→ 行・列が足りなければ足す → 余りが 6 行以上なら行を消す、3 列以上なら列を消す → 1 行目から `setValues`（`syncMasters:300-323`） | 同じ | 同じ | 同じ |
| 書き方の失敗 | タブを消して作り直し、値だけ書く（書式・列幅・タブの位置が変わる。`syncMasters:68-75`） | 同じ（148-156） | 同じ（194-201） | 同じ（236-243） |
| 元のタブが無い | `SKIP`（何もしない） | 例外（何もしない） | 例外 | 例外 |
| 書き先のタブが無い | `SKIP`（**作らない**） | 例外（`requireSheet_`） | 作る | 作る |
| 元が 0 行 | `SKIP` | 例外 | 例外 | 例外 |
| 元が見出しだけ | **見出しだけに置き換える** | 同じ | 同じ | 同じ |
| 元の列が増えた | 増えた列もそのまま写る（見出しの最後の値の列まで） | 同じ | 同じ。位置で読むので、**途中に挿入されると壊れる** | 同じ（位置で読む） |
| 値の型 | `getValues` の値を `setValues`。Date は Date のまま、数字だけの文字は数になる見込み（推測） | 日付は Date の瞬間のまま写る | — | — |

運用の事実（確認済み）:
- `syncAllMasters` は 4 つを順に動かし、タブごとに失敗を握って続けます。結果はログだけで、どこにも残しません（BatchLogs にも書かない）。
- `syncAllMasters` は自分のロックを持ったまま `syncCalendarMaster` を呼び、その中でもう一度 `waitLock` と `releaseLock` をします（`syncMasters:86,122-123,163`）。スクリプトのロックが同じ実行の中で入れ子にできるなら、内側の `releaseLock` で外側のロックも外れ、StampMap・DeptApprovers はロックなしで動きます（推測。実行数の画面で `ERROR M_CALENDAR` が出ていないかを見ると、入れ子が通るかが分かります）。`pdfExport:132-133` のコメントは「再入不可」と書いています。
- 同期の間（数秒）は書き先のタブが空です。その間に申請すると `api_getWorkingDays` が空を返し、「勤務日ではありません」になる恐れがあります（推測。6 時台なので影響は小さい）。
- M_CALENDAR の元は残業 DB（時間帯は `Etc/GMT`）。Date の瞬間のまま写すので、このアプリの DB の時間帯が JST なら、日付のセルに 9 時の時刻が付いて見える見込みです（推測）。GAS は JST で日付にするので正しく読めます。

VPS で変える点（前例と同じ提案）: 元が読めない・見出しだけ・有効な行 0 なら上書きしない。タブを消して作り直さない。M_STAMP・M_DEPT_APPROVERS の見出しを照合し、違えば止めて管理者に出す。作業員マスタは「書き先のタブが無ければ作らない」を残すかを仕様で決める。

## 5. トリガー・メニュー・保守用の関数

| 関数 | 入口 | 今も作られるか・動くか | 確かさ |
|---|---|---|---|
| `doGet` | Web アプリ | 動く | 確認済み |
| `onOpen` | シンプルトリガー（DB を開いたとき） | 動く。メニュー「休暇届アプリ」8 項目（`menu:97-113`） | 確認済み |
| `syncAllMasters` | 時間主導・毎日 6 時台（`setupSyncTrigger_`、`menu:381-390`）。メニュー | `setupAllTriggers` で作る（同名が無いときだけ） | コード確認済み／登録は【要確認】（DB の最終更新 06:03 と合う＝00 の 2 節） |
| `checkAndSendWarnMails_` | 時間主導・毎日 8 時台（`setupWarnMailTrigger_`、`menu:392-401`） | 今は作らない。**3/4〜3/11 は `setupAllTriggers` が作っていた** | 登録の有無は推測 |
| `notifyTodayLeaves_`（7 時台）・`notifyPendingApproval1_`（17 時台）・`notifyPendingApproval2_`（18 時台） | `setupNotifyTriggers_`（`menu:404-423`） | 今は作らない。**3/11 16:03〜20:09 の間は作っていた** | 登録の有無は推測 |
| `setupAllSheets` | メニュー | 手で押すと、タブと設定キーを足す | 確認済み |
| `syncCalendarMaster`・`syncStampMaster`・`syncDeptApprovers` | メニュー | 手で | 確認済み |
| `rebuildLeaveSummary` | メニュー | 手で | 確認済み |
| `migrateLeaveRequestHeaders_` | メニュー | 末尾が `_` なのでメニューから動かない見込み（`0b1a657` で「メニュー呼出関数を public 化」したが、この関数は後で足された） | 推測 |
| `checkAndSendWarnMails` | メニュー | 送信日（7/1・9/1・1/1・2/16）以外は何もしない（コメントの「強制送信」は誤り） | 確認済み |
| `setupAllTriggers` | メニュー | 同期トリガーだけ作る | 確認済み |
| `debug_checkLeaveRequestSheet`・`test_generatePdf`（REQ_ID `LV-20260307-001` 固定）・`test_stampLookup` | エディタ（誰でも `google.script.run` からも呼べる） | 手で | 確認済み |
| `api_approveLeaveRequest2Batch` | `google.script.run` | 画面からは呼ばない（`fb22b3a` で 1 件ずつに変えた） | 確認済み |
| `getResizedStampBlob_` | — | どこからも呼ばない（`SlidesApp` を使うがスコープも無い） | 確認済み |

末尾 `_` の関数を時間主導トリガーの宛先にできるかは、確かな資料がありません（推測）。古いトリガーが残っていれば、統合業務メールと同じ通知が二重に届いているはずなので、オーナーの受信箱でも確かめられます。切り替えでは、トリガーの画面に出た**全部**（`syncAllMasters` 以外も）を消す一覧に入れます。

## 6. ロックと待ち

### 6-1. `LockService.getScriptLock` の全箇所（確認済み）

| 関数 | 待ち時間 | ロックの中でしていること | 注意 |
|---|---|---|---|
| `api_submitLeaveRequest`（`leaveRequest:132-133`） | 15 秒 | 検証（勤務日の計算・振替の重なり）→ 採番 → 1 行書く → flush | 採番の重なりを防ぐ要 |
| `api_approveLeaveRequest`（474-475） | 60 秒 | サイン画像を Drive へ → 4 列 → flush → 集計の全書き直し → 承認メール → 台帳 | 重い処理をすべてロックの中で行う |
| `api_approveLeaveRequest2`（623-624） | 60 秒 | 3 列 → flush → 前の PDF をゴミ箱 → PDF（約 15 秒）→ 2 列 → 台帳 | 画面が 1 件ずつ呼ぶので、総務の一括は直列 |
| `api_approveLeaveRequest2Batch`（762-817） | 60 秒 | 3 列だけ。PDF・台帳はロックの外 | 画面からは使わない |
| `api_editLeaveRequest`（853-854） | 20 秒 | 読み直し（最大 6 秒待つ）→ 書く → PDF ゴミ箱 → 台帳の行削除 → 集計 | — |
| `api_deleteLeaveRequest`（971-972） | 20 秒 | 同じ | — |
| `syncAllMasters`（`syncMasters:14-15`） | 30 秒 | 4 つの同期 | 入れ子（4 節） |
| `syncCalendarMaster`（122-123） | 30 秒 | カレンダーの同期 | — |

ロックが無い書き込み: `rebuildLeaveSummary_`（管理画面・取得状況を開くたび。同時に開くと消去と書き込みが重なる）、`syncStampMaster`・`syncDeptApprovers` を単独で押したとき、一括 2 次承認の PDF と台帳、`setupAllSheets`。台帳が開けないときの自動作成がロックの外で 2 本同時に走ると、台帳が 2 冊できる恐れがあります（推測）。

`waitLock` が時間切れになると例外になり、画面には失敗の文が出ます。VPS では行ロック（`SELECT … FOR UPDATE`）と採番の一意制約に置き換えます。

### 6-2. `flush` と読み直しの待ち（確認済み）

| 箇所 | 待ち | 目的 |
|---|---|---|
| `api_getLeaveRequestById`（`leaveRequest:243-248`） | 見つからなければ flush＋1.5 秒を最大 4 回（約 6 秒） | 書いた直後に読めない対策（`69976d9`）。編集・取消・台帳・メール・PDF も通る |
| `generateLeavePdf_`（`pdfExport:144-148`） | flush＋4 秒＋flush＋1 秒 | 画像の反映待ち |
| 電子印の差し込み（`pdfExport:467-468,511-512`） | flush＋0.5 秒（1 次・2 次それぞれ） | 同上 |
| `readSheetData_`（`syncMasters:276,284`） | 失敗したら 2 秒を 2 回 | 大きなタブの読み取り失敗 |
| 書いた後の flush | 申請・承認・編集・取消・集計・台帳・同期・セットアップ | 次の読み取りに見せるため |

VPS の DB では読み直しの待ちは要りません。ただし GAS の内部用の入口は今もシートを読み直すので、「行の反映が済んでから呼ぶ」ことで待ちを不要にします（2-2）。

## 7. 移行の道具が気をつける型の化け

実物はまだ読めていません。コードと git から「あり得るもの」を挙げ、確かめ方を添えます。gviz の基本の URL は `https://docs.google.com/spreadsheets/d/16v1tQNCFkDTk2mnAZ4E2eEdHeUQ2Bepdcsn5vBAmllw/gviz/tq?tqx=out:html&headers=1&sheet=T_LEAVE_REQUEST&tq=<クエリ>` です（`05` の書き方）。列の文字は**先に見出しを見て決めます**。下の例は 1-1 の「ア. 雛形」の並び（A REQ_ID・B 年度・C 承認状態・D 申請日時・G 作業員ID・I 休暇日・J 休暇区分・K 半日区分・L 休暇種類・M 振替元）を仮に使っています。

gviz の性質: 1 つの列に型が混ざると、少ない方の型のセルは null になります（罠 11）。この性質を使うと「日付の列にある文字」を null として拾えます。型を正しく確かめるのは Sheets API（`valueRenderOption=UNFORMATTED_VALUE`＋`FORMATTED_VALUE`、`dateTimeRenderOption=SERIAL_NUMBER`）が正です。

| # | あり得る化け | 根拠 | 確かめ方 |
|---|---|---|---|
| T1 | 見出しが英語・日本語・混在。`PDF_FILE_ID` が無い。2 次承認の列が英語名と日本語名の両方ある | 1-1 | `tq=select * limit 0`（見出しだけ出る）。Sheets API で 1 行目を読む |
| T2 | REQ_ID が空の行が 500 行ほど先頭にある（雛形の名残） | `c04fdd6` | `select count(B) where A is null`、`select count(A) where A is not null` |
| T3 | REQ_ID の形が違う行（3/4 の試し・位置ずれ・人が書いた行） | 採番は初日から同じ形（`6e2c72a`）。3/4 13:28〜14:54 は日本語の並びのまま英語の見出しへ `appendRow` していた（`6e2c72a`→`769a2bd`） | `select A, C, D where A is not null and not A matches 'LV-[0-9]{8}-[0-9]{3,}'`。3/4 の行は列がずれている恐れがあるので、`select * where A matches 'LV-20260304-.*'` で中身を見る |
| T4 | REQ_ID の重なり | 採番はロックの中なので通常は起きない。人の複製はあり得る | `select A, count(B) group by A having count(B) > 1`（gviz に having は無いので、`select A, count(B) group by A order by count(B) desc limit 5`） |
| T5 | 休暇日・振替元出勤日が文字 | 3/28〜4/11 の編集が文字で書いた（`1238fba`→`3b8a1f7`）。表示形式が `@` なら文字のまま残る | `select A, I where A is not null and I is null`（文字のセルが null で出る）。`M` も同じ（ただし振替以外は空が正常） |
| T6 | 休暇日の時刻が JST 9 時（4/10 18:35 より前の申請）、振替元が 9 時（4/11 10:57 より前） | `new Date('yyyy-MM-dd')` は UTC 0 時 | Sheets API の UNFORMATTED でシリアル値の小数を見る（シートの時間帯で 0.375 か 0）。gviz では分からない |
| T7 | 年度が 4/1 区切りの行（休暇日が 2026-03-16〜03-31 で、3/11 19:24 より前に出した申請は 2025 のまま） | `c85813d` は既存の行を直していない | `select A, B, I where I >= date '2026-03-16' and I <= date '2026-03-31'`。Sheets API で全行の `年度` と `computeFiscalYear_(休暇日)` を比べる |
| T8 | 年度が数ではなく文字 | 書くのは数。人が直すと文字もあり得る | `select B, count(A) group by B`（文字の行は null の組に入る） |
| T9 | 作業員ID の先頭 0 が消えて数になっている。マスタ側と形が違う | `setValues` は数字だけの文字を数にする（前例の `master_sync.py` の注記「'0042' は '42'」と同じ仕組み） | `select G, count(A) group by G`、作業員マスタで `select A limit 20`。両方を Sheets API の UNFORMATTED で読み、数か文字かを見る |
| T10 | 文字の前後・中の空白（全角の空白を含む）、全角英数 | GAS は読むとき `normalize_` で前後の空白と全角英数を直す。書くときは直さない | `select A, F, H where H matches '.*[ 　].*' or F matches '.*[ 　].*'`。DB には `normalize_` 後の値を入れ、元の文字は照合用に持つか、仕様で決める |
| T11 | 承認状態の想定外の値（空・`却下`・前後の空白） | `却下` は書かないが定義はある | `select C, count(A) where A is not null group by C` |
| T12 | 種類と合わない値（全日休なのに半日区分、振替以外に振替元） | 申請は画面の値を種類に関係なく書く（`leaveRequest:213,216-218`） | `select A, J, K where J = '全日休' and K is not null`、`select A, L, M where M is not null and L != '振替休暇'` |
| T13 | 2 次承認済みなのに承認状態が `申請中` か `取消` | 編集は 2 次の列も消すが、取消は消さない。1 次承認は状態を見ない | 2 次承認日時の列を X として `select A, C, X where X is not null and C != '承認済'` |
| T14 | 4/10 より前の 1 次承認だけの行に PDF の URL がある（1 次の印だけの PDF） | `2a37cbe` まで 1 次承認で PDF を作った | `select A, C, T where T is not null and X is null`（X＝2 次承認日時） |
| T15 | 手書きサインの承認で APPROVED_BY1_EMAIL が空 | `sign.html:205` | `select count(A) where C = '承認済' and <1次メールの列> is null` |
| T16 | 作成者メールが空 | `Session.getActiveUser()` が空を返す端末 | `select U, count(A) group by U` |
| T17 | 追加詳細（DETAIL_FREE）に値がある行 | 3/4〜3/7 は書いていた | `select A, Q where Q is not null` |

移行の道具で決めること（提案）:
- 年度・休暇日・振替元は**計算し直さずに移す**。T6・T7 の行は数えて報告に出し、照合では「JST の日付が同じなら同じ」とみなす。行の型（文字か日付か、時刻の小数）は行ごとに持ち、書き戻すときに元の型を変えない（前例の `cell_kinds`）。
- 作業員ID は「GAS が読む形」（`normalize_` 後の文字）で照合し、シートへは元の型で書き戻す。
- REQ_ID は空の行を飛ばし、形の違う行は止めて報告する（テスト行を外すかはオーナーに聞く）。VPS の採番は、**シートに残る全行**（外した行も）を数に入れる。外した行と同じ番号を出すと、行の反映で取り合うためです。
- 承認状態は `申請中`・`承認済`・`取消` だけを許し、ほかは止める（罠 13）。

## 8. これまでの経緯

### 8-1. 年表（git。日付は JST）

| 日 | コミット | できごと | 今のコードへの名残 |
|---|---|---|---|
| 3/4 13:28 | `6e2c72a` | 初期実装（16 ファイル）。年度 4/1、PDF はひな形を写す、`appendRow` で日本語の並び | — |
| 3/4 13:50 | `a04a90a` | 実データに合わせ M_DEPT をやめ、作業員マスタの「部署」列から部署を作る | `SYNC_MAP` の `src:null`、完了メッセージ「M_DEPT / M_WORKER のサンプル」（`menu:311`） |
| 3/4 14:54 | `769a2bd` | **本番の見出しが英語だった**ので別名を足し、見出しの位置で書くように | `HEADER_ALIAS_` |
| 3/4 14:59 | `c04fdd6` | 空行 500 件が一覧に出たのを直す（REQ_ID が空なら飛ばす） | 各所の `if (!rowReqId) continue` |
| 3/4 15:07 | `69976d9` | `appendRow`→`setValues`、書いた直後の読み直しを 5 回まで | 6-2 の 6 秒待ち |
| 3/7 | `4b4ee3b`〜`2721ff7` | M_CALENDAR の同期、勤務日だけの休暇日、振替元の候補を休日から | — |
| 3/7 | `c874e11` | PDF をひな形なしの動的な組み立てに | `PDF_TEMPLATE_SSID`（使わない）、要件定義の「テンプレをコピー」 |
| 3/7 | `cbb80bf` | 追加詳細（ステップ 8）をやめる | `DETAIL_FREE`／`追加詳細` の列と読み取り |
| 3/7〜3/11 | `a75e1ab`〜`96014d2` | 電子印の 2MB 制限で 6 回作り直し（サムネイル→makeCopy→IMAGE 式→insertImage） | `getResizedStampBlob_`（使わない）、空の説明だけのコメント（`pdfExport:653-662,700-706`） |
| 3/11 12:56 | `40b67a9` | 部署管理・総務管理を追加 | — |
| 3/11 16:03 | `e0a4b7e` | 2 次承認、通知メール 3 種、2 次承認の 4 列、トリガー 3 本 | `notifyMail.gs`、`setupNotifyTriggers_`、`migrateLeaveRequestHeaders_` |
| 3/11 16:52 | `aad3ec9` | 日本語の見出ししか無いと 2 次承認が書かれない不具合 → 逆引きの別名 | 本番の 2 次承認の列は日本語名の見込み |
| 3/11 19:24〜19:33 | `c85813d`・`819476e`・`1c67f05` | **年度 4/1→3/16**、期間の表記 | 既存の行の年度は直していない（T7） |
| 3/11 20:09 | `5ce6954` | **メールのトリガーを統合業務メールへ移す**（作る処理をコメントにした） | `notifyMail.gs`・`warnMail.gs` の関数、古いトリガーの恐れ |
| 3/12 | `11fb2a5` | 台帳を自動で作る | 3/12 より前の承認は台帳に無い見込み（再承認しない限り） |
| 3/28 | `1238fba` | 編集・削除（取消）を追加 | 3/28〜4/11 の編集は休暇日を文字で書いた（T5） |
| 4/10 18:35 | `2c22068` | 過去日の遡及申請。休暇日を JST の 0 時に | 過去申請は列を持たず、申請日時と休暇日から計算 |
| 4/10 18:35 | `2a37cbe` | **PDF を 2 次承認のときだけに**（両方の印を付けて 1 回） | 4/10 より前の行の PDF（T14） |
| 4/10 | `4b87572` | `CREATED_BY_EMAIL` を見出しの定義に、`computeDaysAgo_` を JST に | — |
| 4/11 | `0a9f53b`→`e03a072` | 残業アプリの Requests と振替を連携 → 同日に撤回（現場は作業報告書で処理するため） | 月次期間と使い回しの禁止は残る |
| 4/11 10:57 | `3b8a1f7` | 振替元を JST の 0 時に、編集で休暇日を Date に | T5・T6 の境目 |
| 4/11 | `7935afb`・`f1012ac` | 振替の月次期間を画面でも確かめる | `gas-pulled` はこの時点の写し |
| 7/7 | `d39585e`・`216db54` | スタッフ向けの取得状況ページ（誰でも見られる） | 開くたびに集計タブを書き直す |

`gas-pulled\` は、`.clasp.json` が同じ scriptId を指す GAS の写しです。中身は `f1012ac`（4/11）と一致します（改行と `appsscript.json` の末尾の改行だけ違う。確認済み）。`leaveStatus` の 2 ファイルが無く、`menu.gs`・`top.html` だけが今と違うので、**デプロイ @104 の中身と同じ**見込みです（推測。00 の 2 節「@104 は menu.gs・top.html が違い、leaveStatus が無い」と合う）。ファイルの日付は 5/26 です。

### 8-2. 仕様の変わり目と、残っている名残

| 変わったこと | 今のコード | 名残 |
|---|---|---|
| 年度 4/1 → 3/16 | `computeFiscalYear_`（`leaveRequest:34-44`） | 関数の説明「FY_START_MONTH(デフォルト4)基準」（`leaveRequest:32`）、古い行の年度 |
| PDF を 1 次承認 → 2 次承認だけに | `leaveRequest:527-528` | `pdfExport:132-133` のコメント「呼び出し元（api_approveLeaveRequest）が保持」（今は 2 次承認） |
| 通知を統合業務メールへ | `menu:345-351` | `notifyMail.gs` 全体、`setupWarnMailTrigger_`・`setupNotifyTriggers_`、`getSomuEmails_` |
| PDF のひな形をやめた | `pdfExport.gs` | `PDF_TEMPLATE_SSID`、雛形のタブ（PDF_TEMPLATE・PDF_MAP） |
| 追加詳細をやめた | 申請で書かない | 列と `additionalDetail` の読み取り |
| 2 次承認の一括をやめ 1 件ずつに | `somuAdmin.html` | `api_approveLeaveRequest2Batch` |
| 却下の運用はしない | 書かない | `STATUS.REJECTED`、編集の「取消/却下」の文 |
| BatchLogs | 書かない | タブと `SHEET.BATCH_LOGS` |

### 8-3. docs の記述で今のコードと食い違うもの

| 資料 | 記述 | 今のコード |
|---|---|---|
| 要件定義 8 節 | 年度は 4/1〜3/31 | 3/16〜翌 3/15 |
| 要件定義 6・7 節 | 承認のときに PDF を作る、テンプレをコピー | 2 次承認のときだけ、ひな形なし |
| 要件定義 2 節 | 承認者のアカウントは保存しない | 1 次・2 次の承認者メールと名前を保存 |
| 要件定義 5 節 ⑧ | 詳細追加（自由記述） | やめた |
| 要件定義 3 節・設計決定事項 1 節 | `GENERAL_EMAILS` | 読まない |
| 要件定義 8 節 | 警告メールは時間トリガー | トリガーを作らない（統合業務メールが送る） |
| 要件定義 10 節 | T_LEAVE_REQUEST 19 項目 | 22〜27 列 |
| 設計決定事項 1 節 | 認可の真実は `ALLOW_ADMIN_PAGE_EMAILS`、`SOMU_EMAILS` は認可に使わない | 総務管理・2 次承認は `SOMU_EMAILS`・`ADMIN_EMAILS`。`ALLOW_…` は `?page=admin` だけ |
| 設計決定事項 2 節 | `取消` は将来用で画面に出さない | 3/28 から使う |
| 設計決定事項 3 節 | 管理ページは 1 画面 | admin・deptAdmin・somuAdmin の 3 画面 |
| 設計決定事項 4 節 | M_DEPT・M_WORKER、`MASTER_SOURCE_DEPT_SHEET` | 作業員マスタだけ |
| 設計決定事項 5 節 | サインの後は TOP へ | サインの後はプレビュー（`sign.html:263-267`） |
| DB マスターガイド | タブ名 CONFIG・LOOKUP・M_DEPT・M_WORKER、T_LEAVE_REQUEST 22 列（英語）、V_LEAVE_SUMMARY 11 列（英語）、年度 4/1 | タブ名は `SHEET`、集計は日本語 12 列で書き直す |
| `warnMail:3` のコメント | 日付の確かめを飛ばして強制送信 | 送信日だけ送る |
| `menu:311` の完了メッセージ | M_DEPT / M_WORKER のサンプルを置き換え | そのタブは無い |

## 9. 下書きとの違い（誤り・抜け）

### 9-1. `02-inventory-quick.md`

| 節 | 下書き | 正しくは・足すこと | 重さ |
|---|---|---|---|
| 3 節 T_LEAVE_REQUEST | 見出しは `setupAllSheets` の 25 列 | 本番は雛形の英語 22 列から始まり、2 次承認の列は日本語名で足された見込み。`PDF_FILE_ID` が無い恐れ（1-1） | 高 |
| 3 節 | （記述なし） | 先頭に REQ_ID が空の行が約 500 行ある見込み（`c04fdd6`） | 高 |
| 3 節の型の段落 | 休暇日・振替元は JST の 0 時 | 4/10・4/11 より前は JST 9 時、3/28〜4/11 の編集は文字（T5・T6） | 高 |
| 3 節の型の段落 | 年度は数値 | 3/11 より前の行は 4/1 区切りのまま（T7） | 中 |
| 3 節 M_SYSTEM_SETTING | 台帳を作ったときだけ GAS が行を足す | `setupAllSheets` も足す・空を既定で上書きする。下の行が勝つのは**空の値でも** | 中 |
| 3 節 マスタ | 「マスタ元と同じ」 | 元が見出しだけなら空になる。M_STAMP・M_DEPT_APPROVERS は位置で読む。作業員マスタは書き先が無いと作らない（4 節） | 中 |
| 3 節 M_LOOKUP | （中身は書いていない） | `setupAllSheets` が `生理休暇`・`裁判員休暇` を入れるが、`3eb98f2` で画面の参考例からは外した。実物の行は【要確認】 | 低 |
| 4 節 | 既定値の欄 | 予備の値は設定が空でもコードが使う（`syncMasters:21-24,36-39,128-136,174-176,216-218`）。StampMap・DeptApprovers のタブ名は固定 | 低 |
| 8 節 | 時間帯 | M_CALENDAR は GMT の残業 DB から Date の瞬間のまま写る（4 節） | 低 |
| 9 節に足す | — | `api_getAllPendingRequests` は `MM/dd HH:mm` の文字で並べるので、年をまたぐと 12 月の申請が 1 月より上に来る（`leaveRequest:450,458-460`。確認済み） | 低 |
| 9 節に足す | — | 申請は種類に関係なく振替元・特別理由・有給詳細・半日区分を書く（T12） | 低 |
| 9 節に足す | — | `syncAllMasters` のロックの入れ子（4 節） | 低 |
| 10 節 | flush と読み直しは不要 | GAS の内部用の入口はシートを読み直すので、行の反映を待つ順序が要る（2-2） | 中 |

### 9-2. `01-dependencies.md`

| 節 | 下書き | 正しくは・足すこと |
|---|---|---|
| 1 節の前文 | 「DB の同じ名前のタブを全部消して書く」 | 書き先の名前は違う（`M_CALENDAR`・`M_STAMP`・`M_DEPT_APPROVERS`）。消すのは値だけ（`clearContents`）で、タブを消すのは書き込みの失敗のときだけ |
| 1 節の表 DeptApprovers | A 部署・B メール・C 名前 | 元は 5 列（D 役職・E 通知メール）で、5 列とも写る見込み。C の見出しは ` approverName`（先頭に空白） |
| 1 節の位置で読む箇所 | `auth.gs:84-86`、`masters.gs:145-146` | ほかに `auth:140-142,195-196`、`leaveRequest:657-658,748-749`、`notifyMail:24-26`、`masters:175-176` |
| 2-2 の 2 | 日付は日付型のセル | 既存の行には 9 時の行・文字の行があり得る。行の反映は**変えていない列を書かない**ので、古い行の型はそのまま残る |

### 9-3. `05-risks-and-checks.md`

| 箇所 | 下書き | 足すこと |
|---|---|---|
| C-9 | `syncAllMasters` 1 本の見込み | 3/4〜3/11 に作られた有給警告（8 時）と通知 3 本が残っている恐れ（5 節）。あれば切り替えで消す |
| C-2 | 見出しと行数 | `PDF_FILE_ID` の列の有無、2 次承認の列の名前（英語・日本語）、REQ_ID が空の行の数 |
| C-3 の後に足す | — | 7 節の T3〜T17 のクエリ。型は Sheets API の UNFORMATTED で T5・T6・T9 を確かめる |
| 3 節 `LEDGER_SSID` | 台帳のコピーの ID を入れる | `LEDGER_SSID` の行が 2 つ以上あれば、**一番下の行**を書き換える（3-2） |
| R-12 | 列の位置 | M_STAMP の C 列（名前）も位置で読む。2 次承認者の名前と承認者の一覧に効く |
| 追加の危険 | — | `PDF_FILE_ID` の列が無いと、編集・取消・2 次承認で古い PDF をゴミ箱へ移していない。VPS の照合で「PDF の数」を比べるときに Drive の数と合わない（推測） |

## 10. 移行の道具と行の反映で決めるべき点

| # | 決めること | 提案 |
|---|---|---|
| D1 | 見出しの対応 | 英語・日本語の両方を `HEADER_ALIAS_` と同じ表で受ける。`normalize_` と同じ直し方で比べる。決まりに無い列は触らない |
| D2 | GAS が書く列 | サイン画像URL・PDF_URL・PDF_FILE_ID は行の反映から外し、空にするのも GAS の仕事にする（2-2） |
| D3 | 誰も書かない列 | SPECIAL_REASON_REF・DETAIL_FREE・UPDATED_AT は値を持つだけで書かない |
| D4 | 古い行の型 | 行ごとの型（文字・日付・時刻の小数）を持ち、変えていない列は書かない。照合は JST の日付で比べる |
| D5 | 年度 | 計算し直さない。食い違いは報告だけ |
| D6 | 作業員ID | GAS が読む形で照合、元の型で書き戻す |
| D7 | 空の行・行番号 | REQ_ID が空の行は移さないが、行番号は元のまま持つ。新しい行は最後のデータ行の次 |
| D8 | 採番 | シートの全行の REQ_ID を数に入れ、申請した日（JST）の最大＋1 |
| D9 | V_LEAVE_SUMMARY | 見出し 12 列、`年度_作業員ID` の文字順、年度は元の値、最終有給日の書き方は本番の画面を見てから（K5）。書き直しの後に統合業務メールの条件で読めることを確かめる |
| D10 | GAS の仕事の前提 | PDF・台帳・承認メールは、その行（台帳は同じ作業員の行も）の反映が済んでから呼ぶ |
| D11 | マスタ同期 | 空・見出しだけなら上書きしない、タブを作り直さない、位置で読むタブは見出しを照合 |
| D12 | Settings | 空の値でも下の行が勝つ。重なりは報告 |

## 11. 確かめること

- T_LEAVE_REQUEST の見出し（英語・日本語・混在）、`PDF_FILE_ID` の有無、REQ_ID が空の行の数（C-2）。
- 7 節の T3〜T17 の件数（gviz と Sheets API）。
- トリガー画面の全件（`syncAllMasters` 以外が残っていないか）。
- 実行数の画面で `syncAllMasters` のログに `ERROR M_CALENDAR`（ロックの入れ子）が出ていないか。
- M_SYSTEM_SETTING に同じキー（とくに `LEDGER_SSID`）が 2 行以上ないか。
- M_LOOKUP の実際の行、M_DEPT_APPROVERS・M_STAMP の列数と見出し。
- 作業員マスタの `作業員コード` が数か文字か（先頭 0 の有無）。
