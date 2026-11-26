# 棚卸し 1：業務ロジック（Workflow）

作成: 2026-11-26。対象は `src/` の `leaveRequest.gs`（1,006 行）・`masters.gs`（560 行）・`leaveSummary.gs`（474 行）・`leaveStatusApi.gs`（77 行）・`auth.gs`（203 行）・`config.gs`（252 行）の全行です。ここから呼ばれる `ledger.gs`・`mail.gs`・`warnMail.gs`・`pdfExport.gs`（`saveSignImage_`・`generateLeavePdf_` の入口）と、呼び出し元の画面（`top.html`・`sign.html`・`somuAdmin.html` の該当行）も追いました。

印の意味は次のとおりです。

- **確認済み**: コードで確かめたこと（行番号つき）
- **推測**: コードからそう読めるが、本番のシート・画面では確かめていないこと

この文書は VPS 版の仕様と受け入れテストの元にします。表の「受け入れ」列は、そのままテストの観点に使える形で書きました。

---

## 1. 画面から呼べる公開関数（業務ロジック分）

末尾が `_` でない `.gs` の関数は 43 個です（確認済み・grep）。下書き 02 の数と一致します。ここでは業務ロジックに関わる 27 個を、引数・戻り値・エラー・権限・ロック・書き込み・順番まで書きます。同期・メニュー・テスト用の 16 個は 02 の 1-2 節のとおりです（`api_approveLeaveRequest2Batch` だけ 1-12 に書きました）。

共通の決まり（確認済み）:

- ログイン中のメールは `Session.getActiveUser().getEmail()` を小文字にして使います（`auth.gs:7,19,31,47,62,131,188`）。
- 文字の比べ方は `normalize_`（前後の空白を除き、全角英数字を半角にする。`config.gs:128-133`）を通します。ただし引数の `reqId` 側は `normalize_` を通さずに比べる箇所が多いです（`leaveRequest.gs:287,488,637`）。
- 列は見出しの名前で引きます。英語・日本語のどちらの見出しでも引けます（`HEADER_ALIAS_`・`buildHeaderIndex_`、`config.gs:70-95,135-157`）。例外は REQ_ID の採番で、**A 列に決め打ち**です（`leaveRequest.gs:16`）。
- エラーの返し方は 2 通りあります。申請（`api_submitLeaveRequest`）の入力エラーだけが `{ok:false, error}` を返し、ほかはすべて `throw new Error(...)` です。画面は前者を `alert('申請に失敗しました：\n' + msg)`、後者を `'エラー: ' + err.message` で出します（`top.html:1180-1187`、`sign.html:279`）。VPS でもこの 2 通りを分けて残す必要があります。

### 1-1. 書き込みをする関数

#### `api_submitLeaveRequest(data)`（`leaveRequest.gs:131-234`）

| 項目 | 中身（確認済み） |
|---|---|
| 引数 | `data = {deptId, deptName, workerId, workerName, leaveKubun, halfDayType, leaveDate:'yyyy-MM-dd', leaveType, substituteDate:'yyyy-MM-dd' or '', specialReason, paidDetail, isRetroactive:boolean}`（画面 `top.html:1149-1162`）。`additionalDetail` は JSDoc にあるが、画面は送らず、サーバーも書きません |
| 戻り値 | 成功 `{ok:true, reqId:'LV-yyyyMMdd-NNN'}`。入力エラー `{ok:false, error:'…'}` |
| 権限 | なし（ドメインの誰でも） |
| ロック | `ScriptLock.waitLock(15000)`（133 行）。取れないと GAS 標準の例外（`{ok:false}` ではなく失敗ハンドラへ） |
| 検証の順番 | ① 休暇日の形式 → ② 過去申請なら年度内、通常なら勤務日 → ③ 振替休暇なら振替元の形式 → 月次期間 → 使い回し。詳しくは 3 節 |
| 書く列 | REQ_ID・年度（休暇日から `computeFiscalYear_`）・部署ID・部署名・作業員ID・作業員名・休暇区分・半日区分・休暇日（`new Date('…T00:00:00+09:00')`）・休暇種類・振替元出勤日（`data.substituteDate` が空でなければ**休暇種類に関係なく**日付で書く、216 行）・特別理由・有給詳細・申請日時（`new Date()`）・承認状態＝`申請中`・作成者メール（＝`CREATED_BY_EMAIL`、ログイン中のメール）。ほかの列は空文字 |
| 書き方 | 見出しが無ければ 25 列の見出しを先に書く（182-196 行）。`getLastRow()+1` の行へ `setValues`（225-226 行） |
| 裏の処理 | なし（集計・メール・台帳・PDF は動かない） |

エラー文言（原文）:

| 条件 | 文言 |
|---|---|
| 休暇日が空か `^\d{4}-\d{2}-\d{2}$` でない | `休暇日の形式が不正です（yyyy-MM-dd）。` |
| `isRetroactive===true` で年度内でない | `過去日申請は今年度内（年度開始の3/16から本日まで）の日付のみ指定できます。` |
| 通常申請で勤務日一覧に無い | `指定された休暇日は勤務日ではありません。過去日の場合は「過去申請」から選択してください。` |
| 振替休暇で振替元が空か形式違い | `振替休暇の場合は振替元出勤日を選択してください。` |
| 月次期間が違う | `振替休暇は同一月次期間（前月16日〜当月15日）内で消化する必要があります。\n休暇日: yyyy-MM-dd（期間: yyyy-MM-dd〜yyyy-MM-dd）\n振替元: yyyy-MM-dd（期間: yyyy-MM-dd〜yyyy-MM-dd）`（`leaveRequest.gs:76-78`） |
| 振替元を使用済み | `指定された振替元出勤日（yyyy-MM-dd）は既に別の振替休暇申請で使用されています。`（119 行） |

#### `api_approveLeaveRequest(reqId, signBase64, approverEmail)`（1 次承認、`leaveRequest.gs:471-557`）

| 項目 | 中身（確認済み） |
|---|---|
| 引数 | `reqId`、`signBase64`（手書きのときだけ `data:image/png;base64,…`。電子印のときは `''`）、`approverEmail`（電子印の持ち主のメール。手書きのときは `''`）（`sign.html:249-252,205,216,284`） |
| 戻り値 | `{ok:true, reqId}` |
| 権限 | **なし**。`approverEmail` が M_STAMP にあるかも確かめない |
| 状態の確かめ | **なし**（申請中・承認済・取消・却下のどれでも通る） |
| ロック | `waitLock(60000)`。裏の処理まで全部ロックの中 |
| エラー | `REQ_IDが指定されていません。`／`申請データがありません。`／`申請が見つかりません: <reqId>`／`T_LEAVE_REQUEST のヘッダに "<列名>" が見つかりません。ヘッダ: <見出しをカンマ区切り>`（必須列は REQ_ID・承認状態・承認日時・サイン画像URL） |
| 書く列 | 承認状態＝`承認済`、承認日時＝`new Date()`、サイン画像URL＝保存した URL か `''`（**電子印を選ぶと前の値を空で上書き**）、APPROVED_BY1_EMAIL＝`approverEmail` か `''` |
| 書かない列 | APPROVED_BY2・APPROVED_BY2_EMAIL・APPROVED_AT2・PDF_URL・PDF_FILE_ID はそのまま |
| 裏の処理の順番 | ① サイン画像を Drive の休暇日フォルダへ（失敗はログだけで続行、508-514 行）→ ② 4 列を書く → `flush` → ③ PDF は作らない（527 行）→ ④ `rebuildLeaveSummary_()`（失敗は続行）→ ⑤ `sendApprovalNotification_`（`MAIL_TO` が空なら送らない、失敗は続行）→ `flush` → ⑥ `updateLedger_`（台帳に upsert、失敗は続行） |

#### `api_approveLeaveRequest2(reqId, approver2Email)`（2 次承認、`leaveRequest.gs:619-726`）

| 項目 | 中身（確認済み） |
|---|---|
| 引数 | `reqId`、`approver2Email`（画面はログイン中のメールを `api_getCurrentApproverInfo` から取って渡す。`somuAdmin.html:212,328,403`） |
| 戻り値 | `{ok:true, reqId}` |
| 権限 | 総務（`SOMU_EMAILS`）か管理者（`ADMIN_EMAILS`）。**ただし REQ_ID の空チェックが先**（620-621 行）。`approver2Email` がログイン中の人と同じかは見ない |
| 状態の確かめ | 承認状態が `承認済` でなければ `1次承認されていない申請です。`。**2 次承認済み（APPROVED_AT2 あり）かは見ない**ので、もう一度通せる |
| ロック | `waitLock(60000)`。**PDF 作成（1 件 15 秒ほど）までロックの中** |
| エラー | `REQ_IDが指定されていません。`／`権限がありません（総務または管理者のみ）`／`申請データがありません。`／`申請が見つかりません: <reqId>`／`1次承認されていない申請です。` |
| 書く列 | APPROVED_BY2＝M_STAMP の C 列（備考＝名前。メールの一致は小文字で比べる）、無ければメール。APPROVED_BY2_EMAIL＝引数のまま。APPROVED_AT2＝`new Date()`。PDF ができたら PDF_URL・PDF_FILE_ID |
| 裏の処理の順番 | ① 3 列を書く → `flush` → ② APPROVED_BY1_EMAIL を読む（空なら同じセルを読み直す。空でも続行）→ ③ **前の PDF を先にゴミ箱へ** → ④ `generateLeavePdf_(reqId, 1次メール, 2次メール)` → 成功なら PDF の 2 列（PDF 失敗はログだけで続行）→ `flush` → ⑤ `updateLedger_` |

#### `api_editLeaveRequest(reqId, patch)`（`leaveRequest.gs:852-966`）

| 項目 | 中身（確認済み） |
|---|---|
| 引数 | `patch = {leaveDate:'yyyy-MM-dd', leaveKubun:'全日休'|'半日休', halfDayType:'午前'|'午後'|''}`（画面 `top.html:1307-1311`）。サーバーは `patch.substituteDate` も受けるが、**画面は送らない** |
| 戻り値 | `{ok:true, newStatus:'申請中' or 元の状態}` |
| 権限 | なし（誰でも、どの申請でも。申請者の確かめなし） |
| 状態の確かめ | `取消` と `却下` は `取消/却下済みの申請は編集できません。`。それ以外は通る |
| ロック | `waitLock(20000)` |
| 検証 | 3 節の表のとおり。**休暇日が勤務日か・年度内かは見ない**（形式だけ） |
| 書く列 | 休暇日（JST の 0 時）と年度（休暇日から計算）、休暇区分、全日休なら半日区分を空、半日休なら半日区分、振替休暇で `patch.substituteDate` があれば振替元出勤日 |
| 承認済みのとき | 承認状態＝`申請中`、承認日時・APPROVED_BY1_EMAIL・サイン画像URL・APPROVED_BY2・APPROVED_BY2_EMAIL・APPROVED_AT2 を空、PDF_FILE_ID があればそのファイルをゴミ箱＋PDF_URL・PDF_FILE_ID を空、台帳の行を消す（**変更が何も無くても、保存するだけで承認が外れる**） |
| 裏の処理の順番 | `api_getLeaveRequestById` で読む（見つからないと約 6 秒待つ）→ 検証 → 列を 1 セルずつ書く → 承認済みなら上の処理 → `flush` → `rebuildLeaveSummary_`（失敗は続行） |
| 書かない列 | 申請日時・更新日時（`UPDATED_AT` の別名はあるが、どこも書かない） |

エラー文言（原文）: `申請が見つかりません。`／`取消/却下済みの申請は編集できません。`／`休暇日の形式が不正です: <値>`／`休暇区分が不正です: <値>`／`半日区分を選択してください。`／`振替元出勤日の形式が不正です: <値>`／`振替休暇には振替元出勤日が必須です。編集画面から振替元を指定してください。`／`振替休暇には休暇日が必須です。`／月次期間・使い回しの文言（申請と同じ）。

#### `api_deleteLeaveRequest(reqId)`（取消、`leaveRequest.gs:970-1006`）

| 項目 | 中身（確認済み） |
|---|---|
| 戻り値 | `{ok:true}` |
| 権限 | なし（誰でも、どの申請でも） |
| 状態の確かめ | `取消` だけ `既に取消済みです。`。**承認済・2 次承認済・却下も取り消せる** |
| ロック | `waitLock(20000)` |
| エラー | `申請が見つかりません。`／`既に取消済みです。` |
| 書く列 | 承認状態＝`取消`。PDF_FILE_ID があればゴミ箱＋PDF の 2 列を空。**承認日時・承認者などの列は残す** |
| 裏の処理の順番 | 状態を書く → PDF → `flush` → 台帳の行を消す → `rebuildLeaveSummary_` |

#### `api_approveLeaveRequest2Batch(reqIds, approver2Email)`（`leaveRequest.gs:735-846`。画面からは呼ばれない）

| 項目 | 中身（確認済み） |
|---|---|
| 戻り値 | `{ok:true, results:[{reqId, ok:true} or {reqId, ok:false, error}]}`。PDF が失敗しても `ok:true` |
| 権限 | 総務か管理者（`権限がありません`、句点なし） |
| エラー | `申請IDが指定されていません`／`一度に承認できるのは20件までです`（21 件以上）／`申請データがありません。`／1 件ごとに `申請が見つかりません`・`1次承認されていない申請です`（句点なし） |
| ロック | 3 列の書き込みだけロックの中。**PDF の作成と PDF 列の書き込み・台帳はロックの外**（820-843 行） |

### 1-2. 読むだけの関数（ただし 4 つは集計タブを書き直す）

| 関数（定義） | 引数 | 戻り値の形（キー・型・日付の形） | 権限 | 裏で書くもの |
|---|---|---|---|---|
| `api_getLeaveRequestById`（`leaveRequest.gs:239-323`） | reqId | `null` か `{reqId, fiscalYear:文字, deptId, deptName, workerId, workerName, leaveKubun, halfDayType, leaveDate:'yyyy-MM-dd', leaveType, substituteDate:'yyyy-MM-dd' or '', specialReason, paidDetail, additionalDetail, submittedAt:'yyyy-MM-dd HH:mm:ss', approvedAt:'yyyy-MM-dd HH:mm:ss' or '', status, signImageUrl, pdfUrl, approvedBy1Email, approvedBy2, approvedBy2Email, approvedAt2:'yyyy-MM-dd HH:mm:ss' or '', rowNo:数}`。見つからないと 1.5 秒×4 回待ってから `null` | なし（誰でも全項目。メール・サイン URL・PDF URL も） | なし |
| `api_getUpcomingLeaves`（356-413） | deptId, workerId（空なら絞らない） | `[{reqId, deptName, workerName, leaveKubun, halfDayType, leaveDate:'yyyy/MM/dd', leaveType, status, isRetroactive:bool, daysAgo:数}]`。休暇日の昇順（文字で比べる） | なし | なし |
| `api_getAllPendingRequests`（419-463） | — | `[{reqId, deptName, workerName, leaveKubun, halfDayType, leaveDate:'yyyy/MM/dd', leaveType, submittedAt:'MM/dd HH:mm', status, isRetroactive, daysAgo}]`。`submittedAt` の文字の降順 | なし | なし |
| `api_getPending2ndApprovals`（565-611） | — | `[{reqId, deptName, workerName, leaveDate:'yyyy/MM/dd'（日付でなければ元の文字）, leaveKubun, leaveType, approvedAt:'yyyy/MM/dd HH:mm', isRetroactive, daysAgo}]`。休暇日の昇順。`halfDayType` は**入らない** | 総務か管理者（`権限がありません（総務または管理者のみ）`） | なし |
| `api_getWorkingDays`（`masters.gs:419-472`） | — | `['yyyy-MM-dd', …]`（今日〜年度末 3/15）。M_CALENDAR が無い・空・見出しが違うと `[]` | なし | なし |
| `api_getHolidayDays`（220-314） | workerId（`''` 可） | `[{date:'yyyy-MM-dd', label}]`。label は `休日`・`祝日`（M_CALENDAR の区分そのまま）、`土曜`・`日曜`、`祝日（<祝日名>）` | なし | なし |
| `api_getPastWorkingDays`（480-560） | `'yyyy-MM'` か `'yyyyMM'`（ほかは今月） | `[{date:'yyyy-MM-dd', dow:0-6, label:'yyyy-MM-dd（曜）'}]` | なし | なし |
| `api_getDeptList`（190） | — | `[{deptId, deptName}]`（同じ値＝部署名）。JS の既定の `sort()`（文字コード順） | なし | なし |
| `api_getWorkersByDept`（194） | deptId | `[{workerId, workerName, deptId}]`。`workerId.localeCompare` の順（数字も文字として比べる） | なし | なし |
| `api_getWorkerInfo`（77-114） | — | `null` か `{workerId, workerName, deptId, deptName, email}` | なし | なし |
| `api_getLookupData`（202-210） | — | `{leaveKubun:['全日休','半日休'], halfDayType:['午前','午後'], leaveTypeFull:['有給消化（私用）','振替休暇','特別休暇'], leaveTypeHalf:['有給消化（私用）','特別休暇'], specialReasons:[M_LOOKUP の値]}` | なし | なし |
| `api_getApproverList`（163-186） | — | `[{email, name}]`（M_STAMP の A 列・C 列、シートの順）。エラーは `[]` | なし | なし |
| `api_getCurrentApproverInfo`（`auth.gs:187-203`） | — | `{email, name}`（名前は M_STAMP の C 列、無ければメール） | なし | なし |
| `api_getApproverProfiles`（`auth.gs:130-180`） | — | `{isAdmin, currentEmail, profiles:[{email, name, depts:[]}]}`。名前の `localeCompare('ja')` 順 | 管理者は全員分、ほかは自分の分だけ | なし |
| `api_isAdmin`・`api_isSomu`・`api_isDeptApprover`（`auth.gs:108-118`） | — | bool | — | なし |
| `api_getDeptAdminView`（`leaveSummary.gs:428-447`） | fyYear, selectedDepts | `{workers:[4 節の形], deptNames:[]}` | 管理者か部署承認者（`部署承認者の権限がありません。`）。管理者は好きな部署、承認者は自分の部署だけ残す（残らなければ空） | **集計タブ** |
| `api_getSomuAdminView`（455-474） | fyYear, deptFilter（`''`・`'ALL'` で全部署） | `[4 節の形]` | 総務か管理者 | **集計タブ** |
| `api_getAdminBossView`（155-198） | deptId, fyYear | `[{deptId, deptName, workerId, workerName, paidCount, paidHalfCount, substituteCount, specialCount, totalCount, lastPaidDate, warnLevel}]`。部署名＋氏名の `localeCompare('ja')` 順。**承認済みが 1 件も無い人は出ない** | `ALLOW_ADMIN_PAGE_EMAILS`（`権限がありません`） | **集計タブ** |
| `api_getAdminSomuView`（203-206） | fyYear, deptFilter | 上と同じ | 同上 | **集計タブ** |
| `api_getLeaveStatusSearch`（`leaveStatusApi.gs:12-49`） | fyYear, deptId, workerId | workerId あり: `{mode:'personal', fyYear:数, worker:(4 節の形) or null, milestones:[], upcoming:[]}`。無し: `{mode:'list', fyYear, deptFilter, workers:[]}` | **なし（誰でも）** | **集計タブ** |

`buildAdminViewData_` が返す 1 人分の形（`leaveSummary.gs:381-406`）: `{deptName, workerId, workerName, paidCount, paidHalfCount, substituteCount, specialCount, totalCount, lastPaidDate, warnLevel, quarters:[{id, label, targetCount, periodCount, cumulativeCount, shortage, level:'danger'|'warn'|'ok'|'', taken, overdue, isCurrent, isPast, isFuture}×5], currentShortage}`。

`buildMilestones_` の 1 要素（`leaveStatusApi.gs:70-75`）: `{label, month, day, dateText:'M/D', required, met, passed, deficit, remaining, state:'done'|'risk'|'warn'|'note'|'future'}`。

---

## 2. 状態の遷移

状態の値は `申請中`・`承認済`・`取消`・`却下` の 4 つです（`config.gs:16-21`）。**2 次承認済みは状態の値ではありません**。`承認済` で APPROVED_AT2 が空でないことを指します（`leaveRequest.gs:590`）。`却下` に変える処理は**どこにもありません**（確認済み・grep）。

| 遷移 | 関数 | 条件（確認済み） |
|---|---|---|
| （新規）→ 申請中 | `api_submitLeaveRequest` | 3 節の検証を通ること |
| 申請中 → 承認済（1 次） | `api_approveLeaveRequest` | 条件なし |
| **取消・却下・承認済 → 承認済** | `api_approveLeaveRequest` | 条件なし（**状態を見ずに書く**） |
| 承認済 → 承認済＋2 次 | `api_approveLeaveRequest2`・`…Batch` | 状態が `承認済` |
| 承認済＋2 次 → 承認済＋2 次（やり直し） | 同上 | APPROVED_AT2 を見ないので通る。前の PDF を捨てて作り直す |
| 承認済（2 次の有無を問わず）→ 申請中 | `api_editLeaveRequest` | 取消・却下でないこと。変更が無くても戻る |
| 申請中 → 申請中 | `api_editLeaveRequest` | 同上 |
| 取消以外 → 取消 | `api_deleteLeaveRequest` | 取消でないこと |

状態を確かめずに書く箇所（確認済み）:

| 箇所 | 起きること |
|---|---|
| 1 次承認（`leaveRequest.gs:517-524`） | 取り消した申請が承認済に戻り、集計・台帳に入る。取消で空いた振替元を別の申請が使った後だと、**同じ振替元が 2 件**になる |
| 1 次承認を 2 次承認済みの申請にもう一度 | 承認日時・1 次承認者・サインが変わり、APPROVED_AT2 と PDF は古いまま。PDF の所属長印と台帳の 1 次承認者が食い違う。承認メールがもう一度出る |
| 2 次承認のやり直し | 前の PDF をゴミ箱へ入れてから作り直す。作り直しに失敗すると PDF_URL・PDF_FILE_ID は**ゴミ箱のファイルを指したまま**（`leaveRequest.gs:695-710`） |
| 一括 2 次承認の PDF 列（ロックの外） | その間に編集で申請中へ戻った申請に、新しい PDF の URL を書き戻す（推測。コードの並びからそう読める） |
| 取消（`api_deleteLeaveRequest`） | 2 次承認済みでも取り消せる。承認者の列は残る |

---

## 3. 入力の検証

### 3-1. 申請（`api_submitLeaveRequest`）

| 検証 | 中身・境目（確認済み） | 行 |
|---|---|---|
| 休暇日の形式 | `normalize_` の後、`^\d{4}-\d{2}-\d{2}$`。`2026-02-30` のような無い日付は形式では通る（勤務日一覧に無いので通常申請は落ちる。過去申請は文字の大小比べだけなので**通る**） | 136-139 |
| 過去申請の範囲 | `isRetroactive === true`（真偽値の true だけ）のとき、`年度開始 3/16 ≦ 休暇日 ≦ 今日`（**両端を含む**、文字で比べる）。**勤務日かは見ない**（画面のカレンダーが勤務日だけ出すので、画面からは勤務日になる） | 141-146、`config.gs:193-200` |
| 通常申請の範囲 | `api_getWorkingDays()` の一覧に入っていること＝`今日 ≦ 休暇日 ≦ 年度末 3/15`（両端を含む）かつ勤務日。**今日も申請できる**。年度末の直前は翌年度（3/16 以降）を申請できない | 148-153、`masters.gs:444-470` |
| 勤務日の判定 | 4-3 の表 | `masters.gs:456-467` |
| 振替元の必須・形式 | `leaveType === '振替休暇'` のときだけ。空か形式違いは落とす | 159-163 |
| 振替元の月次期間 | 休暇日と振替元の「月次期間の開始日」が同じこと（4-2） | 166-167 |
| 振替元の使い回し | 同じ作業員・休暇種類が振替休暇・状態が取消以外（**却下も使用済み扱い**）の行に、同じ振替元があれば落とす。比べるのは `yyyy-MM-dd` の文字。見出しの列が 1 つでも無いと**黙って通す** | `leaveRequest.gs:91-124` |
| 振替元が休日か | **見ない**（画面の候補が休日だけなので、画面からは休日になる） | — |
| 振替元が未来でないか | **見ない**（画面の候補は今日＋7 日まで） | — |
| 休暇区分・半日区分・休暇種類の値 | **見ない**（どんな文字でも保存する）。半日休の振替休暇も止めない（画面は半日休のとき振替休暇を出さない） | — |
| 部署・作業員がマスタにあるか | **見ない** | — |
| 特別休暇の理由の必須 | **見ない**（サーバーでは） | — |
| 同じ人・同じ日の重複 | **見ない**（同じ日に何件でも申請できる） | — |
| 振替休暇以外の振替元 | 送られてくれば**形式も見ずに**日付にして書く（216 行）。画面は種類を変えても振替元の選択を空に戻さないので、振替休暇→有給と選び直すと有給の行に振替元が残る（推測。`top.html:829,1158` の並びから） | 216 |

### 3-2. 編集（`api_editLeaveRequest`）

| 検証 | 中身・境目（確認済み） | 行 |
|---|---|---|
| 状態 | 取消・却下は不可 | 858-860 |
| 休暇日 | 空でなければ形式だけ。**勤務日・年度内・過去かは見ない**（画面は `type="date"` で何日でも選べる） | 868-870 |
| 休暇区分 | 空か `全日休`・`半日休` | 871-873 |
| 半日区分 | 休暇区分が `半日休` のとき `午前`・`午後` が必須。休暇区分が空なら半日区分は**書かない** | 874-876、916-925 |
| 振替元 | 空でなければ形式 | 877-879 |
| 振替休暇の月次期間 | 申請の休暇種類が振替休暇なら、（新しい休暇日か元の休暇日）と（新しい振替元か元の振替元）で月次期間を比べる | 883-895 |
| 振替休暇の使い回し | 振替元が送られ、元と違うときだけ。自分の REQ_ID は除く | 897-900 |
| 休暇種類の変更 | できない（patch に無い） | — |

---

## 4. 日付の計算

前提（確認済み）: `appsscript.json` の時間帯は `Asia/Tokyo`。`new Date(y, m, d)`・`getFullYear()`・`getMonth()`・`getDate()`・`getDay()`・`setHours(0,0,0,0)` はすべて JST で動きます。`fmtDate_` も JST です（`config.gs:98-101`）。

### 4-1. 年度

| 規則 | 中身（確認済み） | 行 |
|---|---|---|
| 年度 | 1/1〜3/15 は前年、3/16〜12/31 はその年（**3/16 は新年度**） | `leaveRequest.gs:34-44` |
| 年度の開始日 | `new Date(年度, 2, 16)`＝3/16 の 0 時 | `config.gs:184-187` |
| 年度末 | `new Date(年度+1, 2, 15)`＝3/15 の 0 時（勤務日の範囲の終わり） | `masters.gs:447` |
| 年度を保存する時期 | 申請と編集で、休暇日から計算して「年度」列へ | `leaveRequest.gs:175,911` |
| コメントの誤り | 32 行の「FY_START_MONTH(デフォルト4)基準」は古い。設定キーは無く、3/16 固定 | `leaveRequest.gs:32` |

### 4-2. 月次期間

| 規則 | 中身（確認済み） |
|---|---|
| 期間 | 日が 16 以上なら「その月 16 日〜翌月 15 日」、15 以下なら「前月 16 日〜その月 15 日」（`leaveRequest.gs:50-63`） |
| 同じ期間の判定 | 2 つの日の期間の開始日が同じ（`getTime()` で比べる、73 行） |
| 例 | 休暇日 4/15・振替元 3/16 → 同じ（3/16〜4/15）。休暇日 4/16・振替元 4/15 → 違う |

### 4-3. 勤務日・休日の判定（M_CALENDAR と計算した祝日）

M_CALENDAR は見出し `日付`・`区分` で読みます。区分で意味があるのは `休日`・`祝日`・`出勤土曜` の 3 つだけです（確認済み）。

| 日の種類 | `api_getWorkingDays`・`api_getPastWorkingDays`（勤務日か） | `api_getHolidayDays`（振替元の候補か） |
|---|---|---|
| M_CALENDAR の区分が `休日`・`祝日` | 勤務日でない | 候補（ラベルは区分そのまま） |
| M_CALENDAR の区分が `出勤土曜` | 勤務日 | 候補でない |
| M_CALENDAR にあるが区分がほかの値・平日 | **計算した祝日なら勤務日でない**、それ以外は勤務日 | **候補でない**（計算した祝日を見ない） |
| M_CALENDAR にあるが区分がほかの値・土日 | 勤務日でない | 候補（`土曜`・`日曜`） |
| M_CALENDAR に無い・平日 | 計算した祝日なら勤務日でない | 計算した祝日なら候補（`祝日（名前）`） |
| M_CALENDAR に無い・土日 | 勤務日でない | 候補（`土曜`・`日曜`） |
| M_CALENDAR のシートが無い・空 | `api_getWorkingDays` は**全部 `[]`**（通常申請がすべて落ちる）。`api_getPastWorkingDays` は計算だけで続ける | 計算だけで続ける |

この表の 3 行目は食い違いです。会社が出勤日にした祝日（区分が `休日`・`祝日`・`出勤土曜` 以外）は、勤務日にも振替元の候補にもなりません（確認済み・コード。M_CALENDAR に実際どんな区分があるかは【要確認】）。

計算した祝日（`masters.gs:320-413`、確認済み）:

| 規則 | 中身 |
|---|---|
| 計算する年 | **今日**の年度 `fy` と `fy+1` の 2 年だけ |
| 固定 | 1/1 元日、2/11 建国記念の日、2/23 天皇誕生日、4/29 昭和の日、5/3・5/4・5/5、8/11 山の日、11/3 文化の日、11/23 勤労感謝の日 |
| ハッピーマンデー | 1 月第 2 月曜 成人の日、7 月第 3 月曜 海の日、9 月第 3 月曜 敬老の日、10 月第 2 月曜 スポーツの日 |
| 春分 | `floor(20.8431 + 0.242194×(y−1980) − floor((y−1980)/4))` 日（3 月） |
| 秋分 | `floor(23.2488 + 0.242194×(y−1980) − floor((y−1980)/4))` 日（9 月） |
| 振替休日 | 祝日が日曜なら、翌日から数えて最初の「祝日でない日」。土曜の祝日は振り替えない |
| 国民の休日 | 祝日（振替休日を足す**前**の一覧）を日付順に並べ、隣どうしの差がちょうど 2 日で、間の日が祝日でなく土日でもなければ `国民の休日` |
| 並べ方 | `new Date('yyyy-MM-ddT00:00:00')`（時間帯なし＝JST として読む）で曜日を出す |

### 4-4. 範囲と「今日」

| 関数 | 範囲（確認済み） |
|---|---|
| `api_getWorkingDays` | 今日 0 時〜年度末 3/15（両端を含む） |
| `api_getHolidayDays` | 年度開始 3/16〜**今日＋7 日**（両端を含む。年度末を越えても止めない） |
| `api_getPastWorkingDays` | 「月初」と「年度開始 3/16」の遅い方〜「月末 23:59:59」と「今日 0 時」の早い方。さらに今日より後は打ち切る。前年度の月・未来の月は `[]` |
| `isWithinCurrentFiscalYear_` | 年度開始〜今日（両端を含む、文字で比べる） |
| `api_getUpcomingLeaves` | 休暇日 ≧ 今日 0 時（今日を含む）。休暇日が日付でない行は**除かずに**載せる（休暇日は `''`） |
| `computeIsRetroactive_` | 申請日（JST の日付）＞休暇日のとき true（同じ日は false） |
| `computeDaysAgo_` | `floor((今日0時 − 休暇日0時) / 1日)`。0 以下は 0 |

### 4-5. 有給の期限と期間

期限（`PAID_LEAVE_WARN`、`config.gs:43-49`）と期間（`PAID_LEAVE_QUARTERS`、`config.gs:52-58`）は、**3 番目の区切りが違います**（期限は 12/31、期間は 11/30）。

| # | 期限（警告レベル・マイルストーン） | 必要数 | 期間（期間別） | 目標（累計） |
|---|---|---|---|---|
| 1 | 6/30（6月末） | 1 | 春期 3/16〜6/30 | 1 |
| 2 | 8/31（8月末） | 2 | 夏期 7/1〜8/31 | 2 |
| 3 | **12/31（12月末）** | 3 | 秋期 9/1〜**11/30** | 3 |
| 4 | 2/15 | 4 | 冬期 12/1〜2/15 | 4 |
| 5 | 3/15（年度末） | 5 | 年度末 2/16〜3/15 | 5 |

年の決め方はどれも「3/16 以降の月日は年度の年、1/1〜3/15 は年度＋1」です。期限・期間の終わりは **23:59:59**、期間の始まりは 0:00:00 です。

「期限を過ぎたか」の比べ方が 3 か所で違います（確認済み）:

| 箇所 | 比べ方 |
|---|---|
| `computeWarnLevel_`（`leaveSummary.gs:137`） | `now < 期限` なら未到来（＝`now >= 期限` で過ぎた） |
| `buildMilestones_`（`leaveStatusApi.gs:63`） | `now >= 期限` で過ぎた |
| `buildWarnMailBody_`（`warnMail.gs:121`） | `now > 期限` で過ぎた |

差は 23:59:59 ちょうどの 1 秒だけです。警告メールの送信日（7/1・9/1・1/1・2/16、`config.gs:61-66`）は、期限 1〜4 の翌日です。

### 4-6. VPS で JST に置き換えるとき 1 日ずれそうな箇所

GAS では全部 JST で動くので、今はずれません。VPS のサーバーが UTC で動き、「今日」や日付を UTC のまま作ると、**JST の 0:00〜8:59 に 1 日前**になります。次の箇所はすべて JST の日付・時刻で作る必要があります（確認済み・コード）。

| 箇所 | 何がずれるか |
|---|---|
| `generateReqId_` の `yyyyMMdd`（`leaveRequest.gs:7-8`） | REQ_ID の日付が前日になり、前日の連番の続きになる |
| `computeFiscalYear_(new Date())` の既定（多数） | 3/16 の朝に前年度扱い |
| `api_getWorkingDays` の今日（`masters.gs:444-450`） | 今日の申請が「勤務日ではありません」になる、または前日を申請できる |
| `isWithinCurrentFiscalYear_`（`config.gs:195-199`） | 今日の過去申請が落ちる |
| `api_getUpcomingLeaves` の今日 0 時（`leaveRequest.gs:363-364`） | 昨日の休暇が予定に残る |
| `computeDaysAgo_`・`computeIsRetroactive_` | 日数・過去申請の印が 1 ずれる |
| `api_getHolidayDays`（今日＋7 日）・`api_getPastWorkingDays`（今日まで） | 候補が 1 日ずれる |
| 期限・期間の 23:59:59 と `now`（`leaveSummary.gs:134,229-230`、`leaveStatusApi.gs:62`） | 期限の翌日の朝まで警告が出ない |
| 警告メールの送信日の判定（`warnMail.gs:12-23`） | 7/1 の朝 9 時前に動かすと送らない |
| 休暇日・振替元を `'yyyy-MM-ddT00:00:00+09:00'` で日付にする箇所（`leaveRequest.gs:155,164,216,892-893,910-911,929`） | 時間帯を落とすと前日の 15:00 UTC。日付型（`date`）で持てば起きない |
| `new Date('yyyy-MM-dd')`（UTC の 0 時として読む）で日付にする箇所：`updateLedger_`（`ledger.gs:137,167`）、`removeLedgerRow_`（319）、`generateLeavePdf_`（`pdfExport.gs:151`）、`buildApprovalMailBody_`（`mail.gs:32,53`） | GAS では JST 9:00 になるので同じ日。移植で「UTC の 0 時」のまま JST で日付を取らずに使うと、ずれはしないが、時間帯の扱いを変えると 1 日前になる |
| `new Date('yyyy-MM-dd HH:mm:ss')`（時間帯なし＝JST）で読む箇所：`updateLedger_` の申請日・承認日（`ledger.gs:174,181,189`） | UTC として読むと、0:00〜8:59 の申請・承認の台帳の日付が前日になる |
| 祝日の計算の `new Date(… + 'T00:00:00')`（`masters.gs:339,352-353`） | 日付型で計算すれば起きない |
| `api_getAllPendingRequests` などの `MM/dd HH:mm`・`yyyy/MM/dd HH:mm` | 表示がずれる |

---

## 5. REQ_ID の採番

| 項目 | 中身（確認済み、`leaveRequest.gs:6-29`） |
|---|---|
| 形 | `LV-` + **申請した日**（休暇日ではない）の `yyyyMMdd`（JST）+ `-` + 連番 |
| 連番 | T_LEAVE_REQUEST の **A 列**（見出しで引かない）のうち、その日の接頭辞で始まるものの数字部分の最大＋1。3 桁に満たなければ 0 で埋める（1000 件目は `1000`） |
| 行を消したとき | 取消は行を消さないので番号は再利用されない。人が行を消すと、その日の最大番号なら再利用される |
| 排他 | 申請の `ScriptLock`（15 秒）の中で採番から書き込みまで。編集・取消・承認も同じ ScriptLock |

同時実行での穴（確認済み・コード。影響は推測）:

| 穴 | 中身 |
|---|---|
| 二重押し | 冪等キーが無い。画面はボタンを止めるだけ（`top.html:1146`）。通信の再送や 2 つのタブで 2 件できる |
| 同じ内容の重複 | 同じ人・同じ日の申請を止めない |
| A 列の決め打ち | 人が列を並べ替えると、採番だけ別の列を見て番号が重なる |
| ロック待ちの失敗 | 15 秒で取れないと例外になり、画面は「エラー: …」。承認の PDF 作成（60 秒のロックの中で 15 秒ほど）と重なると起き得る |
| 番号が推測できる | 連番なので、権限の無い API（プレビュー・承認・編集・取消）と組み合わせると他人の申請を順に触れる |

VPS では、日ごとの連番をトランザクションの中で採番し、（日付, 連番）に一意制約をかけることになります。形（`LV-yyyyMMdd-NNN`）は統合業務メール・台帳・PDF・URL で使うので変えません。

---

## 6. 集計の数え方

### 6-1. V_LEAVE_SUMMARY（`rebuildLeaveSummary_`、`leaveSummary.gs:15-116`）

| 規則 | 中身（確認済み） |
|---|---|
| 対象 | REQ_ID が空でなく、状態が `承認済` の行（**1 次承認だけでも数える**。2 次承認は見ない） |
| まとめ方 | キー `年度列の値 + '_' + 作業員ID`。年度は**列の値**（休暇日から計算し直さない） |
| 部署・氏名 | そのキーで**最初に出てきた行**の値（シートの行の順） |
| 合計回数 | 承認済の行数（休暇種類が何でも、半日も 1） |
| 有給回数 | 休暇種類が `有給消化（私用）` で、休暇区分が `半日休` なら +0.5（有給半日回数も +1）、それ以外は +1 |
| 振替回数・特別回数 | 行数（半日も 1） |
| 最終有給日 | 有給の行（半日も含む）の休暇日の最大を `'yyyy-MM-dd'` の文字で |
| 警告レベル | `computeWarnLevel_(有給回数, 年度)`（6-3） |
| 並べ方 | キーの文字の昇順 |
| 書き方 | タブが無ければ作る → `clearContents()`（書式は残す）→ 見出し 12 列 → 全行を `setValues`。**ロックなし**（承認・編集・取消からはロックの中、画面を開くときはロックの外） |
| 呼ばれる時 | 1 次承認・編集・取消のあと、`api_getAdminBossView`・`api_getAdminSomuView`・`api_getDeptAdminView`・`api_getSomuAdminView`・`api_getLeaveStatusSearch` を開くたび、警告メール、メニュー |
| 外の読み手 | 統合業務メール（別の GAS）が「警告レベル」などを読む（00 の 1-5） |

### 6-2. 期間別（`computeQuarterStatus_`、`leaveSummary.gs:217-276`）と画面用データ（`buildAdminViewData_`、284-419）

| 規則 | 中身（確認済み） |
|---|---|
| 対象の申請 | 状態 `承認済`・**年度列**が指定年度・休暇種類 `有給消化（私用）`・休暇日が日付型 |
| 期間の数 | 期間に入る休暇日の**件数**（**半日も 1**） |
| 累計 | 前の期間からの件数の足し上げ |
| 不足 | `max(0, 目標 − 累計)` |
| level | 今の期間か過ぎた期間だけ。不足 2 以上 `danger`、1 `warn`、0 `ok`。未来は `''` |
| overdue | 不足 > 0 かつ過ぎた期間 |
| 対象の人 | 作業員マスタの在籍者（在籍フラグ列があれば、空・`1`・`true`（大小無視）だけ）のうち、指定部署の人。**部署はマスタの部署**で決める（申請の部署ではない） |
| 回数 | V_LEAVE_SUMMARY から作業員 ID で引く（無ければ 0）。年度で絞ったあと作業員 ID だけで引くので、部署を移った人は新しい部署に全回数が出る |
| 警告レベル | タブの値ではなく `computeWarnLevel_` で計算し直す |
| currentShortage | 今か過去の期間の不足の最大 |
| 並べ方 | 有給回数の昇順 → currentShortage の降順 → 部署名（`localeCompare('ja')`）→ 氏名（同） |

### 6-3. 警告レベル・マイルストーン

| 関数 | 規則（確認済み） |
|---|---|
| `computeWarnLevel_`（`leaveSummary.gs:123-150`） | 期限を後ろ（年度末）から見て、**過ぎた期限のうち最も新しい 1 つだけ**で判定。有給回数（半日 0.5）が必要数未満なら、不足 3 以上 `危険`、2 以上 `警告`、それ未満（0.5 も）`注意`。足りていれば `''`。過ぎた期限が無ければ `''` |
| `buildMilestones_`（`leaveStatusApi.gs:55-77`） | 期限ごとに、有給回数（半日 0.5）≧ 必要数なら `done`。満たさず過ぎていれば不足 3 以上 `risk`、2 以上 `warn`、それ未満 `note`。過ぎていなければ `future` |
| 警告メールの対象（`warnMail.gs:62-96`） | V_LEAVE_SUMMARY の今年度の行で、警告レベルが空でない人。部署名＋氏名の `localeCompare('ja')` 順 |

### 6-4. 取得状況（`api_getLeaveStatusSearch`）

| 規則 | 中身（確認済み） |
|---|---|
| 年度 | 引数を `Number`、空なら今日の年度 |
| 部署 | 引数があればその 1 部署、無ければマスタの全部署 |
| 本人 | `workerId` と `===` で一致する人（`normalize_` しない）。いなければ `worker:null`・`milestones:[]` |
| upcoming | 今年度を見ているときだけ `api_getUpcomingLeaves(本人のマスタの部署, workerId)`。**申請の部署ID がマスタの部署と同じ行だけ**出る（部署を移る前の申請は出ない） |

### 6-5. 数え方が場所によって違う点

| # | 違い | 場所 |
|---|---|---|
| D1 | 半日の有給：集計・警告レベル・マイルストーンは **0.5**、期間別と台帳の「有給累計取得回数」は **1** | `leaveSummary.gs:67-72`、`233-239`、`ledger.gs:286-305` |
| D2 | 年度の決め方：集計と期間別は**年度列**、台帳の累計は**休暇日から計算し直す**（`computeFiscalYear_`） | `leaveSummary.gs:37,334`、`ledger.gs:302` |
| D3 | 3 番目の区切り：警告は 12/31 で 3 回、期間別は 11/30 で 3 回。12 月は期間別で「秋期 未達」、警告レベルは空、が同時に起きる | `config.gs:46,55` |
| D4 | 人の範囲：古い管理画面（`api_getAdminBossView`）は承認済みがある人だけ。総務・部署管理・取得状況はマスタの在籍者全員（0 回の人も） | `leaveSummary.gs:173-191`、`357-395` |
| D5 | 部署：集計タブは申請の部署、画面用データはマスタの部署 | `leaveSummary.gs:40`、`358` |
| D6 | 警告レベル：古い管理画面とメールはタブの値、ほかは計算し直し（同じ関数なので値は同じ） | `leaveSummary.gs:189,391` |
| D7 | 最終有給日：タブに `'yyyy-MM-dd'` の文字で書くが、シートが日付に変えると、読むとき `normalize_(Date)` で `Fri Nov 27 2026 00:00:00 GMT+0900 (日本標準時)` の形になる（**推測**。02 の K5 と同じ） | `leaveSummary.gs:75,108,188,311` |
| D8 | 状態：どこも「承認済」だけを数え、申請中は数えない（2 次承認の有無は見ない） | — |
| D9 | 有給の理由の既定：台帳は `私用の為`、画面の確認は `私用のため` | `ledger.gs:160`、`top.html:1119` |

---

## 7. 社外に公開すると重くなる点

GAS ではドメインのログインが前にあるので、社内の人しか呼べませんでした。VPS で社外から届く形にすると、次の穴が重くなります（確認済み・コード）。

| # | 穴 | 中身 | 重さ |
|---|---|---|---|
| S1 | 1 次承認に権限なし | 誰でも、どの申請でも、どの人の電子印でも承認できる（`approverEmail` は M_STAMP に無くても通る） | 高 |
| S2 | 1 次承認に状態の確かめなし | 取消・承認済みの申請も承認済にできる（2 節） | 高 |
| S3 | 編集・取消に権限なし | 誰でも他人の申請を取り消し、承認を外せる（2 次承認済みの PDF も捨てる） | 高 |
| S4 | 2 次承認者のメールを引数で受ける | 総務・管理者なら、別の人の電子印で 2 次承認できる | 中 |
| S5 | 情報が誰にでも見える | `api_getLeaveRequestById`（メール・サイン・PDF の URL）、`api_getAllPendingRequests`、`api_getUpcomingLeaves('', '')`（全員の予定）、`api_getLeaveStatusSearch`、`api_getApproverList`（印の持ち主のメール） | 中（要件で誰でも見られる画面もある） |
| S6 | 画面を開くだけで集計タブを書き直す | ロックの外で `clearContents` → 書き込み。同時に開くと、読み手が空のタブを読み 0 回で出る瞬間がある。取得状況は誰でも開けるので、連打で書き込みが続く | 中 |
| S7 | 段の照合が無い | 画面を開いた時点の状態を送らない。2 台で同時に承認・2 次承認・編集すると、後の操作が前を上書きする。更新日時の列も書いていない | 中 |
| S8 | 二重押し | 申請は冪等キーなし。承認の二重押しは承認メールが 2 通・サイン画像が 2 枚 | 中 |
| S9 | 入力をそのまま保存 | 部署名・氏名・種類などを検証せずに保存し、TOP の一覧が `innerHTML` に逃がさず入れる（`top.html:1228-1231`）。`=` で始まる値はシートで数式になる（`setValues`） | 高（保存型の XSS） |
| S10 | サイン画像の大きさに上限なし | 何でも PNG として Drive に保存する | 低 |
| S11 | 一括 2 次承認のロックの外の書き込み | PDF 列と台帳をロックの外で書く（画面は使っていないが呼べる） | 低 |
| S12 | 申請が見つからないと約 6 秒待つ | 存在しない REQ_ID を続けて送ると、そのぶん処理が埋まる | 低 |

---

## 8. 下書き 02（と 00）との違い

### 8-1. 誤り・足りない記述

| # | 下書きの記述 | 正しくは（確認済み） |
|---|---|---|
| E1 | 02 の 1-1 表 `api_editLeaveRequest` の引数に `substituteDate`、00 の 1-1「編集で振替元を変えられる」 | サーバーは受けるが、**画面は送らない**（`top.html:1307-1311`）。画面からは振替元を変えられない |
| E2 | 00 の 1-1「過去申請（遡及）: 今年度（3/16〜今日）の勤務日だけ」 | サーバーは**勤務日を確かめない**（範囲だけ）。勤務日に絞っているのは画面のカレンダー |
| E3 | 02 の 1-1 表 `api_deleteLeaveRequest`「取消済みは例外」だけ | 合ってはいるが、却下・2 次承認済みも取り消せる点、承認の列を残す点が抜けている |
| E4 | 02 の 3 節「REQ_ID は その日の最大の連番＋1」 | 「その日」は**申請した日**。連番は A 列決め打ちで探す |
| E5 | 02 の K3「半日の有給は … 期間別（quarters）では 1、台帳でも 1」 | 合っている。ただし台帳は年度を休暇日から計算し直すので、年度の決め方も違う（D2） |

### 8-2. 抜け

| # | 抜けていたこと | 場所 |
|---|---|---|
| M1 | 編集は変更が無くても承認を外す。2 次承認済みも外す | `leaveRequest.gs:933-955` |
| M2 | 編集は休暇日が勤務日か・過去か・年度内かを見ない | 868-870 |
| M3 | 2 次承認は 2 次承認済みかを見ない（やり直せる）。前の PDF を先に捨てるので、作り直しに失敗すると捨てたファイルを指す | 644-645、695-710 |
| M4 | 取消→1 次承認で生き返り、振替元が 2 件になり得る | 2 節 |
| M5 | 1 次承認で電子印を選ぶとサイン画像URL を空で上書き。サイン保存の失敗は黙って続行 | 508-519 |
| M6 | 申請の検証で見ないもの（区分・種類の値、半日の振替休暇、振替元が休日か、同じ人・同じ日の重複、マスタにあるか） | 3-1 |
| M7 | 振替元の使い回しは**却下も使用済み**、列が無いと黙って通す | `leaveRequest.gs:103-110` |
| M8 | 勤務日と振替元の候補で、M_CALENDAR にある平日の祝日の扱いが違う。M_CALENDAR が空だと通常申請が全部落ちる | 4-3 |
| M9 | 承認待ち一覧の並べ方が `MM/dd HH:mm` の文字なので、**年をまたぐと 12 月の申請が 1 月の申請より上**に出る | `leaveRequest.gs:450,458-460` |
| M10 | 警告の期限 12/31 と期間 11/30 の食い違い、「過ぎた」の比べ方が 3 通り | 4-5、D3 |
| M11 | 古い管理画面は承認済みが 0 件の人を出さない。取得状況の upcoming は部署を移る前の申請を出さない | D4、6-4 |
| M12 | 申請・編集・取消・承認のロックの秒数（15・20・20・60） | 1-1 |
| M13 | エラーの返し方が 2 通り（申請だけ `{ok:false}`） | 1 節の冒頭 |
| M14 | 2 次承認待ち一覧に `halfDayType` が無い。`leaveDate` は日付でなければ元の文字 | `leaveRequest.gs:595-605` |
| M15 | 有給の理由の既定が台帳だけ `私用の為` | D9 |
| M16 | 振替休暇以外の申請にも、送られた振替元を形式を見ずに書く | `leaveRequest.gs:216` |
| M17 | 保存型の XSS と数式の入り込み | S9 |

### 8-3. 行番号

02 と 00 が挙げた行番号（関数の定義行、K1〜K14 の範囲、`config.gs` の範囲）は、`leaveRequest.gs`・`masters.gs`・`leaveSummary.gs`・`leaveStatusApi.gs`・`auth.gs`・`config.gs` の分をすべて照らし、ずれはありませんでした。

---

## 9. オーナーに「直す／今のまま」を聞くべき癖

おすすめを先に書きます。どれも「今のまま」が移行の原則で、直すなら移行と別に扱う想定です。

| # | 癖 | おすすめ |
|---|---|---|
| Q1 | 1 次承認が状態を見ない（取消・承認済も承認できる） | 申請中だけ承認できるようにする（画面の見た目は同じ。S2・M4） |
| Q2 | 編集を保存するだけで承認が外れる（変更なしでも） | 今のまま、または変更があるときだけ外す |
| Q3 | 1 次承認・編集・取消を誰でもできる | 社内のログインを前に置く前提で今のまま。社外公開するなら要相談（S1・S3） |
| Q4 | 半日有給の数え方（0.5 と 1）と年度の決め方の不一致 | 今のまま（D1・D2。総務の見え方が変わるため） |
| Q5 | 警告 12/31 と期間 11/30 | 今のまま（D3） |
| Q6 | 承認待ち一覧の年またぎの並び順 | 直す（年を入れて並べ、表示は `MM/dd HH:mm` のまま。M9） |
| Q7 | 2 次承認の作り直しで前の PDF を先に捨てる | 直す（新しい PDF ができてから捨てる。M3） |
| Q8 | 出勤にした祝日が勤務日にも振替元にもならない | M_CALENDAR の区分の実物を見てから決める（M8） |
| Q9 | 年度末の直前に翌年度の申請ができない | 今のまま（運用を確かめる） |
| Q10 | 同じ人・同じ日の重複申請を止めない | 今のまま（要望があれば別途） |

---

## 10. 仕様を書く人が迷いそうな点

- **2 次承認済みは状態ではありません。** `承認済` かつ APPROVED_AT2 が空でない、で判定します。`却下` は値だけあり、どこからも書かれません。
- **日付の文字の形が API ごとに違います。** `yyyy-MM-dd`（プレビュー・勤務日）、`yyyy/MM/dd`（一覧）、`MM/dd HH:mm`（承認待ち）、`yyyy/MM/dd HH:mm`（2 次承認待ち）、`yyyy-MM-dd HH:mm:ss`（プレビューの日時）。画面はこの形を前提に曜日を出したり、`/` を `-` に置き換えたりしています（`top.html:1212-1216,1272`）。
- **「年度」は列の値と休暇日からの計算の 2 通り**があります（D2）。DB では列として持ち、台帳だけ計算し直す、を再現します。
- **V_LEAVE_SUMMARY は外の読み手がいます。** VPS が正になっても、統合業務メールのためにタブを同じ 12 列で書き続ける必要があります。最終有給日の形（D7）は本番のタブを見て合わせます。
- **並べ方の関数が混ざっています。** JS の既定 `sort()`（文字コード順）、`localeCompare()`（言語指定なし）、`localeCompare(…, 'ja')` の 3 種類です。Python の既定の並べ方とは、全角・半角やかなの並びで違う場合があります。
- **エラーの文言には句点の有無の揺れがあります**（1 件ずつの 2 次承認は `…です。`、一括は `…です`）。画面にそのまま出るので原文のまま写します。
- **`api_getLeaveRequestById` の約 6 秒の待ち**は、シートの書き込み直後の読み取り遅れへの対策です。DB では不要で、「見つかりません」をすぐ返して構いません（画面の文言は同じ）。
- **祝日は「今日」の年度とその翌年の 2 年分だけ**計算します。祝日の計算を年度に関係なく作り直すと、範囲の外の日の判定が変わります（今の API は範囲の外を聞かないので、結果は同じになる見込みです。推測）。
- **集計の書き直しは、承認などの処理の中では失敗しても続行**します。VPS でも、集計やメール・台帳の失敗で承認そのものを失敗にしない、を守ります。

## 確かめること

- M_CALENDAR の「区分」に、`休日`・`祝日`・`出勤土曜` 以外の値があるか（M8・Q8）
- V_LEAVE_SUMMARY の「最終有給日」の実際のセルの型と、画面での見え方（D7）
- T_LEAVE_REQUEST の見出しが英語か日本語か、A 列が REQ_ID か（5 節）
- 「年度」列と休暇日から計算した年度が食い違う行があるか（D2。年度を 4/1 始まりで入れた古い行が残っていないか）
- 取消→承認、2 次承認のやり直しが本番で起きた跡があるか（状態が承認済で取消の操作の跡がある行、PDF_FILE_ID がゴミ箱のファイルを指す行）
