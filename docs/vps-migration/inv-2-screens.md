# 棚卸し 2: 画面（HTML・JS・CSS と入口 doGet）

作成: 2026-11-26。対象は `src/` の全 `.html` 8 個（top 1364 行・somuAdmin 557・leaveStatus 495・deptAdmin 384・preview 371・admin 308・sign 288・no_auth 63）と `menu.gs:3-89` の `doGet`。全行を読みました。
印: **確認済み**＝コードで確かめた、**推測**＝実機・本番で未確認。行番号は `src/<ファイル>:<行>` です。
この表は、画面の移植の仕様と、実機で押して確かめる受け入れ表の元にします。下書き `02-inventory-quick.md` との違いは 13 節にまとめました。

## 1. 入口 `doGet` と画面の一覧（確認済み）

`page` が無い・一覧に無い値なら top です（`menu.gs:4, 81`）。全画面で `setXFrameOptionsMode(ALLOWALL)`。`include_` は定義だけで、どの HTML からも使っていません。

| page | ファイル | 入口の判定（`menu.gs`） | 題名（`setTitle`） | テンプレート変数 | 役割 |
|---|---|---|---|---|---|
| （無し・不明）／`top` | `top.html` | なし（誰でも） | 休暇届 申請TOP | APP_URL・PORTAL_URL | 申請モーダル、承認待ち一覧、本日の休暇者、今後の休暇予定（編集・削除） |
| `preview` | `preview.html` | なし | 休暇届 プレビュー | REQ_ID・APP_URL・PORTAL_URL（PORTAL_URL は HTML で使っていない） | 紙の様式で 1 件を表示、サイン画面へ |
| `sign` | `sign.html` | なし（**1 次承認に権限の確かめなし**） | 休暇届 サイン | REQ_ID・APP_URL（PORTAL_URL は渡さない） | 承認者を選び、電子印か手書きサインで 1 次承認 |
| `deptAdmin` | `deptAdmin.html` | `isAdmin_() \|\| isDeptApprover_()`。だめなら no_auth（`:45-51`） | 休暇届 部署管理 | APP_URL・PORTAL_URL | 部署ごとの有給取得状況 |
| `somuAdmin` | `somuAdmin.html` | `isSomu_() \|\| isAdmin_()`。だめなら no_auth（`:60-66`） | 休暇届 総務管理 | APP_URL・PORTAL_URL | 2 次承認（1 件・一括）と全社の有給取得状況、CSV |
| `leaveStatus` | `leaveStatus.html` | なし（誰でも） | 休暇届 取得状況（HTML にも `<title>` あり `:7`） | APP_URL・PORTAL_URL（PORTAL_URL は使っていない） | 本人ビュー（タイムライン）と一覧ビュー |
| `admin` | `admin.html` | `canAccessAdmin_()`。だめなら no_auth（`:11-18`） | 休暇届 管理 | APP_URL・PORTAL_URL | 古い管理画面（上長ビュー・総務ビュー・CSV）。TOP からのリンクは無い |
| （権限エラー） | `no_auth.html` | — | admin: 「アクセス権限エラー」、deptAdmin・somuAdmin: 「アクセス権限なし」（**2 種類ある**） | message・APP_URL | 鍵の絵文字、「アクセス権限がありません」、`message`、TOP へのボタン |

no_auth の `message` の原文（確認済み）:

| page | message |
|---|---|
| admin | このページは管理者・総務のみアクセスできます。 |
| deptAdmin | このページは部署承認者または管理者のみアクセスできます。 |
| somuAdmin | このページは総務部・管理者のみアクセスできます。 |

`APP_URL`・`PORTAL_URL` は Settings の値を `normalize_`（前後の空白を除き、全角英数を半角に）したものです（`menu.gs:6-7`、`config.gs:128`）。`REQ_ID` は `e.parameter.reqId` をそのまま入れます（形の確かめなし）。

## 2. テンプレート変数の全部（確認済み）

`<?!= ?>`（逃がさない形）は **1 つもありません**。全部 `<?= ?>`（GAS が置き場所に合わせて逃がす）です。

| ファイル:行 | 書き方 | 置き場所 | VPS での置き換え |
|---|---|---|---|
| top:182 / 183 / 184 | `<?= APP_URL ?>?page=leaveStatus` / `deptAdmin` / `somuAdmin` | `<a href>` | Jinja2 `{{ APP_URL }}`（自動エスケープ） |
| top:185 | `<?= PORTAL_URL ?>` | `<a href>` | `{{ PORTAL_URL }}` |
| top:422・423 | `var APP_URL = '<?= APP_URL ?>';`・`var PORTAL_URL = '<?= PORTAL_URL ?>';` | **JS の文字** | JSON の塊から読む（姉妹アプリの `__T`） |
| preview:168・169 | `var REQ_ID = '<?= REQ_ID ?>';`・`var APP_URL = …` | **JS の文字**（REQ_ID は URL から来る値そのもの） | JSON の塊。`REQ_ID` は `json_for_script` で逃がす |
| sign:108・109 | `var REQ_ID = '<?= REQ_ID ?>';`・`var APP_URL = …` | **JS の文字** | 同上 |
| somuAdmin:95 / 96 / 97 | APP_URL（top・deptAdmin）、PORTAL_URL | `<a href>` | `{{ }}` |
| somuAdmin:157・158 | APP_URL・PORTAL_URL | **JS の文字** | JSON |
| deptAdmin:121 / 123 / 124 / 135 | APP_URL（top・somuAdmin・モーダル内の top）、PORTAL_URL | `<a href>` | `{{ }}` |
| deptAdmin:170・171 | APP_URL・PORTAL_URL | **JS の文字** | JSON |
| admin:96 / 97 | APP_URL（top）、PORTAL_URL | `<a href>` | `{{ }}` |
| admin:140・141 | APP_URL・PORTAL_URL | **JS の文字** | JSON |
| leaveStatus:182 | `<?= APP_URL ?>?page=top` | `<a href>`（`target` なし。`<base target="_top">` が効く） | `{{ }}` |
| no_auth:59 | `<?= message ?>` | 本文 | `{{ message }}` |
| no_auth:60 | `<?= APP_URL ?>?page=top` | `<a href>` | `{{ }}` |

数: HTML の属性に 15 か所、JS の文字に 12 か所、本文に 1 か所、合計 28 か所。

リンクは全部 `APP_URL + '?page=…'` の形です。GAS の APP_URL は `…/exec` で終わる見込みなので、VPS では `https://<host>/` のように末尾を `/` にします（姉妹アプリ `pages.py:86-90` と同じ。推測: Settings の実際の値は【要確認】）。

## 3. URL の引数と画面の間の移動（確認済み）

URL の引数は `page`・`reqId`・`approved` の 3 つです。`approved=1`（sign:263）はどこも読みません。

| 元 | 先 | 方法 | 場所 |
|---|---|---|---|
| top ヘッダー | leaveStatus（常に）・deptAdmin（管理者か部署承認者のときだけ表示）・somuAdmin（総務のときだけ表示）・ポータル（PORTAL_URL が空でないとき） | `<a target="_top">` | top:182-185、表示の切り替え 456-469 |
| top 今後の休暇予定（状態が申請中の行） | preview?reqId= | JS で作る `<a target="_top">`「承認」 | top:791 |
| top 承認待ち一覧（全行） | preview?reqId= | 同「承認プレビュー」 | top:1233 |
| top 申請完了 | preview?reqId=（「プレビュー画面へ（承認依頼）」）／top（「TOPに戻る」） | 同 | top:1175-1176 |
| preview「TOPに戻る」 | top | `window.top.location.href` | preview:356-358 |
| preview「承認する（サイン画面へ）」 | sign?reqId= | `window.top.location.href` | preview:359-361 |
| sign「戻る」 | preview?reqId= | `window.top.location.href` | sign:225-227 |
| sign 承認成功 | preview?reqId=&approved=1 | **`meta refresh` 1 秒**（`<head>` に JS で足す）＋予備の `<a target="_top">`「自動で切り替わらない場合はこちら」 | sign:263-269 |
| somuAdmin ヘッダー | top・deptAdmin（管理者か部署承認者）・ポータル | `<a>` | somuAdmin:95-97、164-170 |
| deptAdmin ヘッダー | top・somuAdmin（総務）・ポータル、モーダル内の「申請TOPへ戻る」 | `<a>` | deptAdmin:121-135、179-182 |
| admin ヘッダー | top・ポータル | `<a>` | admin:96-97 |
| leaveStatus | top（「← 申請TOPへ戻る」） | `<a>` | leaveStatus:182 |
| no_auth | top（「トップ画面へ戻る」） | `<a>` | no_auth:60 |
| メール本文 | admin | `APP_URL?page=admin` | `mail.gs`（棚卸し 3 の範囲） |

全ページに `<base target="_top">` があります。VPS では iframe が無いので `window.top` は自分自身です（今のまま残してよい）。

## 4. `google.script.run` の呼び出し全部（確認済み）

画面から呼ぶ関数は **27 種類、呼び出し箇所は 37** です。どの呼び出しも `withUserObject` を使っていません。

| 画面 | 箇所 | 関数の種類 |
|---|---|---|
| top | 15 | 15 |
| somuAdmin | 7 | 6（`api_approveLeaveRequest2` が 2 か所） |
| admin | 4 | 4 |
| leaveStatus | 4 | 4 |
| deptAdmin | 3 | 3 |
| sign | 2 | 2 |
| preview | 1 | 1 |
| no_auth | 0 | 0 |

### 4-1. 1 件ずつ

「成功」「失敗」は withSuccessHandler・withFailureHandler の中身です。「失敗なし」は withFailureHandler が無く、失敗すると画面が「読込中」のまま止まる箇所です（今のまま移すなら、VPS でも何も出さない）。

| # | 画面:行 | 関数（引数） | いつ | 成功ですること・使うキー | 失敗 |
|---|---|---|---|---|---|
| 1 | top:461 | `api_isDeptApprover()` | 開いたとき | true なら「部署管理」を出す | なし |
| 2 | top:464 | `api_isAdmin()` | 同 | true なら「部署管理」を出す | なし |
| 3 | top:469 | `api_isSomu()` | 同 | true なら「総務管理」を出す | なし |
| 4 | top:499 | `api_getDeptList()` | 同 | 部署の選択肢（`deptId`・`deptName`、`data-name`）。続けて #5 | 「-- 取得エラー --」、#9 を呼ぶ |
| 5 | top:494 | `api_getWorkerInfo()`（#4 の成功の中） | 同 | 値があれば `deptId` を選び、`deptName` を覚え、#8 を `workerId` つきで呼ぶ。null なら #9 だけ | なし（予定一覧が「読み込み中...」のまま） |
| 6 | top:507 | `api_getLookupData()` | 同 | `lookupData` に入れるだけ。**どこからも読まない**（種類の選択肢は top:1078-1080 の固定値） | なし |
| 7 | top:528 | `api_getWorkingDays()` | 同 | 休暇日の選択肢。値 `yyyy-MM-dd`、表示 `yyyy-MM-dd（曜）`、土曜は青 `#0ea5e9` | 「-- 取得エラー --」 |
| 8 | top:719 | `api_getWorkersByDept(deptId)` | 部署を選んだとき・#5 から | 作業員の選択肢（表示 `workerId workerName`）。自動選択があれば申請ボタンを有効にし #10 を `workerId` で、無ければ #10 を `''` で。最後に #9 | なし（「読込中...」のまま） |
| 9 | top:809 | `api_getUpcomingLeaves(selectedDeptId, selectedWorkerId)` | 部署・作業員を変えたとき、編集・削除の後 | `leaveDate`（`yyyy/MM/dd`）が今日なら「本日の休暇者」、ほかは「今後の休暇予定」。使うキー `reqId・deptName・workerName・leaveKubun・halfDayType・leaveType・status`。`LEAVE_ITEMS` に全件を覚える | 「予定なし」「0件」 |
| 10 | top:680 | `api_getHolidayDays(workerId \|\| '')` | 開いたとき（`''`）・作業員を変えたとき | 振替元の候補を `ALL_HOLIDAY_CANDIDATES` に（`date`・`label`）、休暇日の月次期間で絞って描く | 「-- 取得エラー --」 |
| 11 | top:936 | `api_getPastWorkingDays('yyyy-MM')` | 過去申請カレンダーの月を描くとき（月ごとに画面でキャッシュ） | `date` だけ使う。応答の前に月が変わっていたら描かない | 空として描く（全部押せない） |
| 12 | top:1191 | `api_submitLeaveRequest(data)` | 「申請する」 | `ok` なら手順の中身を完了表示に差し替え（`reqId`）。`ok:false` なら alert（`error`）しボタンを戻す | alert「エラー: 」＋`err.message`、ボタンを戻す |
| 13 | top:1244 | `api_getAllPendingRequests()` | 開いたとき・編集・削除の後 | 承認待ち一覧。使うキー `leaveDate・isRetroactive・daysAgo・deptName・workerName・leaveKubun・halfDayType・leaveType・submittedAt・reqId` | 「取得エラー」、件数を空に |
| 14 | top:1338 | `api_editLeaveRequest(reqId, {leaveDate, leaveKubun, halfDayType})` | 編集モーダル「保存」 | モーダルを閉じ、`ok` なら alert「申請を更新しました。」→ #9・#13。違えば alert「更新に失敗しました。」 | ボタンを戻し alert「エラー: 」＋`(err.message \|\| err)` |
| 15 | top:1360 | `api_deleteLeaveRequest(reqId)` | 「削除」の確認の後 | `ok` なら alert「申請を削除しました。」→ #9・#13。違えば「削除に失敗しました。」 | alert「削除エラー: 」＋`(err.message \|\| err)` |
| 16 | preview:193 | `api_getLeaveRequestById(REQ_ID)` | 開いたとき（REQ_ID が空なら呼ばない） | null なら「申請データが見つかりません。（REQ_ID: …）」。あれば様式を描く。使うキー `leaveDate・submittedAt・leaveKubun・halfDayType・leaveType・substituteDate・specialReason・paidDetail・status・approvedAt・approvedAt2・approvedBy2・deptName・workerName・reqId・isRetroactive・daysAgo` | 「エラー: 」＋`err.message`（描く途中の例外は「表示エラー: 」） |
| 17 | sign:183 | `api_getApproverList()` | 開いたとき | 「-- 承認者を選択 --」「✍ 手書きサイン（電子印なし）」（値 `__handwrite__`）、続けて `name（email）`（値 `email`） | 選択肢 1 つ「取得エラー: 」＋`err.message`（**innerHTML**） |
| 18 | sign:284 | `api_approveLeaveRequest(REQ_ID, signBase64, selectedApprover)` | 「承認する」の確認の後 | `ok` なら 6 節の動き。違えば「承認に失敗しました。」、ボタンを戻す | 「エラー: 」＋`err.message`、ボタンを戻す |
| 19 | somuAdmin:167 | `api_isDeptApprover()` | 開いたとき | 「部署管理」を出す | なし |
| 20 | somuAdmin:170 | `api_isAdmin()` | 同 | 同 | なし |
| 21 | somuAdmin:194 | `api_getDeptList()` | 同 | 部署の絞り込みの選択肢（先頭は値 `ALL`「全部署」） | なし |
| 22 | somuAdmin:218 | `api_getCurrentApproverInfo()` | 同 | `email` を 2 次承認者として覚える。表示 `<strong>name</strong>（email）` | 「エラー: 」＋`err.message`（textContent） |
| 23 | somuAdmin:232 | `api_getPending2ndApprovals()` | 同・一括で失敗があった後 | 件数バッジと一覧。使うキー `reqId・deptName・workerName・isRetroactive・daysAgo・leaveDate・leaveType・leaveKubun・approvedAt` | 「エラー: 」＋`err.message`、バッジ「エラー」 |
| 24 | somuAdmin:328 | `api_approveLeaveRequest2(reqId, selectedApprover2)` | 行の「2次承認」 | `ok` ならボタン「✓」→行を消す（7 節）。違えばボタン「失敗」で押せる状態に | alert「エラー: 」＋`err.message`、ボタンを戻す |
| 25 | somuAdmin:403 | 同（一括。1 件ずつ順に） | 「一括2次承認」 | **戻り値を見ずに成功として数える**、行を消して次へ | 数える、行を赤 `#ffe0e0`、次へ |
| 26 | somuAdmin:421 | `api_getSomuAdminView(Number(fy), deptFilter)` | 開いたとき・部署を変えたとき・「更新」 | KPI と表。使うキー `paidCount・paidHalfCount・warnLevel・quarters[0-4].{cumulativeCount, shortage, isFuture}・substituteCount・specialCount・totalCount・lastPaidDate・deptName・workerName`。CSV は `quarters[].{taken, overdue}` | 「エラー: 」＋`err.message` |
| 27 | deptAdmin:182 | `api_isSomu()` | 開いたとき | 「総務管理」を出す | なし |
| 28 | deptAdmin:226 | `api_getApproverProfiles()` | 同 | `isAdmin` なら「承認者変更」ボタンを出しモーダル。`profiles` が 1 件なら自動で選ぶ。2 件以上ならモーダル。0 件なら「担当部署が登録されていません」「M_DEPT_APPROVERSに登録がありません。」 | 見出しの下に「エラー: 」＋`err.message` |
| 29 | deptAdmin:286 | `api_getDeptAdminView(fy, SELECTED_DEPTS)` | 承認者を選んだとき・「更新」 | `workers` で KPI と表（キーは #26 と同じ）。`deptNames` は使わない | 「エラー: 」＋`err.message` |
| 30 | admin:180 | `api_isSomu()` | 開いたとき | true なら総務ビューのタブ・部署の絞り込み・CSV を出し #31。最後に #32 か #33 | なし（表が「読み込み中...」のまま） |
| 31 | admin:177 | `api_getDeptList()`（#30 の中） | 同 | 部署の選択肢（先頭は値 `''`「全部署」） | なし |
| 32 | admin:208 | `api_getAdminSomuView(Number(fy), deptFilter)` | 総務ビューのとき | 表と KPI（`paidCount・paidHalfCount・substituteCount・specialCount・totalCount・lastPaidDate・warnLevel・deptName・workerName`） | 「エラー: 」＋`err.message`（**innerHTML**） |
| 33 | admin:215 | `api_getAdminBossView('', Number(fy))` | 上長ビューのとき（部署は常に `''`） | 同 | 同 |
| 34 | leaveStatus:268 | `api_getDeptList()` | 開いたとき | 「全部署」（値 `''`）＋部署 | 「全部署」だけ |
| 35 | leaveStatus:287 | `api_getWorkersByDept(state.dept)` | 部署を選んだとき・行を押したとき（全部署なら呼ばず「全員」だけ） | 「全員」＋`workerName`（値 `workerId`）。選択中の人がいなければ全員に戻す | 「全員」だけ |
| 36 | leaveStatus:313 | `api_getLeaveStatusSearch(state.fy, state.dept, state.staff)` | 年度・部署・スタッフを変えたとき・行を押したとき | `mode` が `personal` なら本人ビュー（`worker・fyYear・milestones[].{met,state,remaining,deficit,required,dateText,passed}・upcoming[]`）、ほかは一覧（`workers[]・deptFilter・fyYear`） | 「読み込みに失敗しました：」＋エラー |
| 37 | leaveStatus:490 | `api_getWorkerInfo()` | 開いたとき（最初） | 「ログイン：氏名（部署）」、`deptId`・`workerId` を初期の選択に。続けて #34・#35・#36 | 「—」、続けて #34・#35・#36 |

### 4-2. 送る値の形（確認済み）

| 関数 | 形 |
|---|---|
| `api_submitLeaveRequest` | `{deptId, deptName, workerId, workerName, leaveKubun, halfDayType, leaveDate:'yyyy-MM-dd', leaveType, substituteDate, specialReason, paidDetail, isRetroactive:bool}`（top:1149-1162）。**種類に関係なく全部の入力欄の値を送る**（14 節 D-6） |
| `api_editLeaveRequest` | `reqId`, `{leaveDate:'yyyy-MM-dd', leaveKubun, halfDayType}`（`substituteDate` は送らない。top:1307-1311） |
| `api_approveLeaveRequest` | `reqId`, `signBase64`（`data:image/png;base64,…` か `''`）, `approverEmail`（電子印ならメール、手書きなら `''`） |
| `api_approveLeaveRequest2` | `reqId`, ログイン中のメール（#22 の値） |
| `api_getDeptAdminView` | 年度（数）, 部署名の配列 |
| `api_getSomuAdminView`・`api_getAdminSomuView` | 年度（数）, 部署（somuAdmin の全部署は `'ALL'`、admin の全部署は `''`） |
| `api_getLeaveStatusSearch` | 年度（数）, 部署（`''`＝全部署）, 作業員（`''`＝全員） |
| `api_getPastWorkingDays` | `'yyyy-MM'` |

## 5. on 属性の全部（確認済み）: 42 個

HTML に直接書いたもの 37 個と、JS の文字で作るもの 5 個です。種類は `onclick` 38 個、`onchange` 4 個。preview・leaveStatus・no_auth には 1 つもありません。別に、deptAdmin は JS でボタンの `.onclick` に関数を入れています（246・256。属性ではないので CSP に引っかからない）。

| 画面 | 数 | 行と中身 |
|---|---|---|
| top | 26 | 202 `openModal()`／251 `closeModal()`／252 `nextStep(2)`／265 `toggleRetroSection()`／273 `changeRetroMonth(-1)`／275 `changeRetroMonth(1)`／283 `prevStep(1)`／284 `nextStep(3)`／298 `prevStep(2)`／299 `nextStep(4)`／310 `prevStepFromType()`／311 `nextStepFromType()`／324 **onchange** `validateSubstituteSelection()`／329 `prevStep(4)`／330 `nextStep(9)`／344 `prevStep(4)`／345 `nextStep(9)`／359 `prevStep(4)`／360 `nextStep(9)`／377 `prevStepFromConfirm()`／378 `submitRequest()`／415 `closeEditLeaveModal()`／416 `saveEditLeave()`／**JS の文字** 787 `openEditLeaveModal('<reqId>')`・788 `confirmDeleteLeave('<reqId>')`・972 `selectRetroDate('<yyyy-MM-dd>')` |
| somuAdmin | 6 | 117 **onchange** `toggleSelectAll()`／119 `batchApprove2()`／138 `loadData()`／139 `exportCsv()`／**JS の文字** 259 **onchange** `updateSelectedCount()`・266 `singleApprove2(<i>, this)` |
| admin | 4 | 113 `loadData()`／114 `exportCsv()`／120 `switchTab('boss')`／121 `switchTab('somu')` |
| sign | 4 | 84 **onchange** `onApproverChange()`／95 `clearCanvas()`／100 `goBack()`／101 `doApprove()` |
| deptAdmin | 2 | 122 `showApproverModal()`／148 `loadData()` |

姉妹アプリの `port-pages.js` で置き換えるときの注意（確認済み・正規表現を当てはめて確かめた）:
- 静的な 37 個は `parseHandler` でそのまま `data-on-click`・`data-on-change` にできます（引数は数か文字だけ）。
- top の JS の文字 3 個は `JS_ONCLICK` の形（`onclick="FN(\'' + expr + '\')"`）に合います（`replaceRegex(…, 3, …)`）。
- somuAdmin の 2 個は形が違うので、個別の置き換えが要ります。259 は姉妹アプリの approve2 の `row-cb` と同じ形、266 は `singleApprove2(' + i + ', this)`（数と `this`）で `__vpsArgs(i, '$this')` にします。

## 6. 確認・アラート・画面の文言（原文のまま。確認済み）

### 6-1. confirm（はい・いいえ）

| 場所 | 文言（`\n` は改行） |
|---|---|
| top:1141 | `過去申請として送信します。\n対象日: <yyyy/MM/dd>\n\n承認者および総務にすでにご連絡済みですか？\nよろしければ「OK」を押してください。` |
| top:1313-1316 | `申請内容を変更しますか？\n休暇日: <yyyy-MM-dd>\n区分: <区分>`＋（半日なら）`(<午前/午後>)`＋（承認済なら）`\n\n※ 承認済みのため、申請中に戻り再承認が必要になります。` |
| top:1344-1345 | `<氏名> の休暇届（<yyyy/MM/dd>）を削除しますか？\n\nこの操作は取り消せません。`（行が見つからないときは `選択した申請を削除しますか？…`） |
| sign:239 | `承認してよろしいですか？` |
| somuAdmin:309 | `<氏名>さんの休暇届（<yyyy/MM/dd>）を2次承認しますか？` |
| somuAdmin:361 | `<n>件を一括2次承認しますか？\n（1件あたり約15秒かかります）` |

### 6-2. alert

| 場所 | 文言 |
|---|---|
| top:814 | 作業員を選択してください |
| top:1012 | 休暇区分を選択してください |
| top:1017 | 休暇日を選択してください（過去申請カレンダーから選択することもできます） |
| top:1028 | 午前・午後を選択してください |
| top:1057 | 休暇種類を選択してください |
| top:1181 | `申請に失敗しました：\n` ＋サーバーの `error`（無ければ「申請に失敗しました。」） |
| top:1187 | `エラー: ` ＋ `err.message` |
| top:1261 | 申請データが見つかりません。画面を更新してください。 |
| top:1303 / 1304 / 1305 | 休暇日を入力してください。／休暇区分を選択してください。／半日区分を選択してください。 |
| top:1326 / 1330 / 1336 | 申請を更新しました。／更新に失敗しました。／`エラー: ` ＋ `(err.message \|\| err)` |
| top:1350 / 1354 / 1358 | 申請を削除しました。／削除に失敗しました。／`削除エラー: ` ＋ `(err.message \|\| err)` |
| sign:232 / 236 | 承認者を選択してください。／サインを書いてください。 |
| somuAdmin:305・351 | 2次承認者（電子印）を選択してください。（#22 が失敗したときだけ出る） |
| somuAdmin:324 | `エラー: ` ＋ `err.message` |
| somuAdmin:379 | `完了: <n>件成功, <n>件失敗`（失敗が 1 件以上のときだけ） |
| somuAdmin:501・admin:274 | データがありません |

サーバーの文言（`error` や例外）は画面にそのまま出ます。VPS の API も同じ文字を返す必要があります（例: `leaveRequest.gs:76, 119, 138, 145, 151, 162, 857-899, 975-976`）。

### 6-3. モーダル・画面の中の文言（主なもの）

| 画面 | 中身 |
|---|---|
| top 申請モーダル（top:237-382） | 題「休暇届 申請」。手順 1 区分（全日休・半日休）→ 2 休暇日（選択肢＋過去申請の折りたたみ「▶ 📅 過去日に申請する場合はこちら」、注意「⚠ 過去申請をされる場合は、必ず事前に承認者および総務へご連絡ください。」、「選択中: なし」／「選択中: yyyy/MM/dd（過去申請）」）→ 3 午前・午後（半日のときだけ）→ 4 種類（全日: 有給消化（私用）・振替休暇・特別休暇、半日: 有給消化（私用）・特別休暇）→ 5 振替元（説明 3 行、エラー枠・OK 枠）／6 特別休暇の理由（参考例：結婚休暇、忌引休暇、出産休暇 等）／7 有給の理由（任意。未入力は「私用のため」）→ 9 確認（部署・作業員・休暇区分・半日区分・休暇日・休暇種類・振替元出勤日・理由）。手順 8 は無い |
| top 申請完了（top:1169-1178） | ✔、「申請が完了しました」、「REQ_ID: …」、「プレビュー画面へ（承認依頼）」「TOPに戻る」。**閉じるボタンは無い**（どちらかのリンクで移るしかない） |
| top 振替の検証（top:600, 611-619） | 「先に休暇日を選択してください。」／「⚠ 月次期間違反」＋休暇日と振替元の期間 2 行＋「振替休暇は同一月次期間（前月16日〜当月15日）で消化する必要があります。」／「✅ 月次期間OK（start〜end）」 |
| top 振替元の選択肢（top:638, 647, 651, 657） | 「-- 先に休暇日を選択してください --」「-- 期間内(start〜end)に候補なし --」「-- 休日出勤日を選択（start〜end） --」、各行 `yyyy-MM-dd（曜・label）`（日曜は赤 `#ff5a5f`、土曜は青） |
| top 編集モーダル（top:385-419） | 題「休暇届 編集」、対象者「部署 ／ 氏名」、振替なら注意（同じ月次期間で消化、別期間は削除して新規作成）、休暇日（`<input type="date">`）、区分、半日区分（半日のときだけ） |
| preview（preview:251-332） | 「休 暇 届」、届出日「令和N年M月D日 届出」、所属部署・氏名、休暇日「令和N年M月D日（曜曜日）」、休暇時間（終日 8:05〜17:10／午前半休 8:05〜12:15／午後半休 13:00〜17:10）、区分・種類・振替元（令和・曜）・理由（有給で空なら「**私用の為**」＝確認画面の「私用のため」と字が違う）、「上記のとおり休暇を届け出いたします。」、届出者、承認欄（部長＝`approvedBy2` の名前の文字、所属長＝承認済なら赤丸「承認済」）、承認日（`令和年/月/日`）、「※後報の場合は、理由を詳しく記載すること。」、REQ_ID |
| sign（sign:79-103） | 「承認処理」、「承認者を選択してください」、バッジ「手書きサイン」／「電子印あり」、サイン枠の説明（初期文「電子印が未登録のため、下の枠内にサインを書いてください」→手書きを選ぶと「下の枠内にサインを書いてください」）、「クリア」「戻る」「承認する」、「承認処理中... しばらくお待ちください。」、「承認が完了しました。承認済プレビューを表示します...」 |
| deptAdmin モーダル（deptAdmin:129-138） | 「承認者を選択」「表示する承認者の部署を選んでください」。管理者には先頭に「全部署」（全員の部署を重なりなしで並べたもの）。**閉じるボタンは無い**（選ぶか「申請TOPへ戻る」） |

## 7. ボタンを止める動き・待ち表示・並び順・絞り込み・バッジ（確認済み）

| 画面 | 動き |
|---|---|
| top 申請ボタン | 作業員を選ぶまで `disabled`（top:202, 691, 713, 726） |
| top「申請する」 | 押すと `disabled`＋「送信中...」。失敗で戻す。成功で手順の中身ごと差し替え（top:1146-1147, 1169） |
| top 振替の「確認へ」 | 月次期間が違う・休暇日が無いと `disabled`（top:584-623） |
| top 編集「保存」 | 「保存中...」で止める。成功・`ok:false` ではモーダルを閉じるだけで戻さない（開き直すと戻る。top:1287-1289） |
| top 過去申請の前後ボタン | 年度の始まり（3/16）より前の月・今日より後の月には動けない（top:873-903） |
| sign「承認する」 | 承認者を選ぶまで `disabled`。押すと「処理中...」。成功で「承認する」「戻る」を隠す |
| somuAdmin 行の「2次承認」 | 押すと「処理中...」→ 成功「✓」、`ok:false`「失敗」 |
| somuAdmin 一括 | 1 件も選ばない・2 次承認者が無いと `disabled`。実行中は全部のチェックと「全選択」を止め、ボタンに「承認中... (i/n)」。終わると「全選択」だけ戻す（行のチェックは失敗があったときの再読込で戻る） |
| somuAdmin 行を消す動き | 背景を緑 `#e6ffed` → 0.4 秒後に右へ 40px ずらして薄く → 0.3 秒後に高さ 0 → 0.3 秒後に消し、件数を数え直す。0 件なら「2次承認待ちの申請はありません」（somuAdmin:273-301） |
| 待ち表示 | 「読み込み中...」「読込中...」（画面で字が違う。top は「読み込み中...」、somuAdmin・deptAdmin は「読込中...」、admin は「読み込み中...」） |

| 一覧 | 並び順 | 絞り込み・バッジ |
|---|---|---|
| top 承認待ち | サーバー: `submittedAt`（`MM/dd HH:mm` の文字）の新しい順（`leaveRequest.gs:458`）。年をまたぐと 12 月が 1 月より上に来る | 全員分（部署の選択と関係なし）。件数「n件」。過去申請は赤いバッジ「🔙 過去申請」＋「(n日前)」、行を薄黄 `#fffbeb` |
| top 本日の休暇者 | サーバーの順（休暇日順） | ブラウザの今日と `leaveDate` が同じ行。0 件ならカードごと隠す。件数「n名」。編集・削除・承認のボタンは無い |
| top 今後の休暇予定 | サーバー: 休暇日順（`leaveRequest.gs:408`） | 選んだ部署・作業員で絞る。件数「n件」。状態のバッジ（承認済は緑、ほかは青）。取消・却下以外に「編集」「削除」、申請中に「承認」 |
| somuAdmin 2 次承認待ち | サーバー: 休暇日順（`leaveRequest.gs:609`） | 件数バッジ「n件」、過去申請バッジ |
| somuAdmin・deptAdmin 有給 | サーバーの順 | 警告のバッジ（危険＝赤・警告＝橙・注意＝somuAdmin は青 `info`、deptAdmin は黄 `caution`）、期間別「残N」「達成」の色 |
| admin | サーバーの順 | 警告のバッジ。KPI の「有給未達者」は注意・警告・危険の数（somuAdmin・deptAdmin は空でない数） |
| leaveStatus 一覧 | **画面で並べ替え**: 危険→警告→注意→順調、同じなら有給の少ない順（leaveStatus:400-402） | 行を押すかキーボードの Enter・空白で本人ビューへ |

年度の選択肢（確認済み）:

| 画面 | 選択肢 | 変えたとき |
|---|---|---|
| somuAdmin・deptAdmin | 今年度＋2 〜 −2 の 5 つ（新しい順）、今年度を選ぶ | **読み直さない**。「更新」を押す必要がある（somuAdmin は部署を変えると読み直す） |
| admin | 今年度 〜 −2 の 3 つ（先頭が今年度） | 読み直さない。「更新」かタブ |
| leaveStatus | ‹ › で今年度 〜 −5 | すぐ読み直す |

## 8. ブラウザ側の検証と日付の計算（確認済み）

サーバーと二重に決めている箇所に「二重」と書きました。

| 場所 | 中身 | 「今日」の基準 | サーバー側 |
|---|---|---|---|
| top:441-447 `getFyStart_` | 過去申請の下限（今年度の 3/16） | **JST に直して計算** | 二重: `isWithinCurrentFiscalYear_`（`leaveRequest.gs:144`） |
| top:881-903, 953-963 | 過去申請カレンダーの上限（今日）と押せる日 | **ブラウザの時刻**（`new Date()`） | 二重: 同上 |
| top:739-740 | 「本日の休暇者」の今日 | ブラウザの時刻 | サーバーは今日以降を返す |
| top:846-849 | カレンダーを開いた月 | ブラウザの時刻 | — |
| top:553-574 `computeMonthlyPeriodJS` | 月次期間（16 日以降は当月 16 日〜翌月 15 日、15 日以前は前月 16 日〜当月 15 日）。振替元の候補を期間で絞り、違えば次へ進めない | — | 二重: `validateSubstituteMonthlyPeriod_`（`leaveRequest.gs:166, 894`） |
| top:1078-1080 | 種類の選択肢（半日は振替を出さない） | — | **サーバーは確かめない**（半日の振替も通る見込み。推測） |
| top:1008-1031, 1057 | 区分・休暇日・午前午後・種類の必須 | — | 休暇日の形と勤務日はサーバーも確かめる（`leaveRequest.gs:137-153`） |
| top:1303-1305 | 編集の必須（日・区分・半日区分） | — | 二重: `leaveRequest.gs:868-879`。**編集の休暇日は勤務日も過去も確かめない**（画面は `<input type="date">` で何日でも選べ、サーバーも勤務日を見ない） |
| top:518, 655 | `new Date('yyyy-MM-ddT00:00:00')` で曜日 | ブラウザの時間帯 | — |
| preview:197-248 | `new Date('yyyy-MM-dd')`（UTC の 0 時と読まれる）、`new Date('yyyy-MM-dd HH:mm:ss')` で令和・曜日。令和＝西暦−2018 | ブラウザの時間帯 | 推測: 日本の時間帯では正しく出る。古い Safari では空白区切りの日時が `Invalid Date` になるおそれ（実機で確かめる） |
| somuAdmin:173-183・deptAdmin:185-195・admin:150-159・leaveStatus:227-232 | 今年度（3/16 始まり） | **ブラウザの時刻** | 二重: `computeFiscalYear_`。3/15〜3/16 の切り替わりの夜だけずれ得る |
| leaveStatus:237-242 | タイムラインの「今日」の位置（3/16 からの日数÷365） | ブラウザの時刻 | — |
| sign:235 | 手書きなのに何も書いていないと止める | — | サーバーは空のサインも通す（推測） |

## 9. サインの画面（確認済み）

| 項目 | 中身 |
|---|---|
| 枠 | `<canvas id="signCanvas">`、CSS で幅 100%・高さ 260px、枠 2px、角丸 12px、背景（CSS）白、`touch-action:none`、`cursor:crosshair`（sign:50-54, 93） |
| 最初の表示 | サインの欄は隠れている（`display:none`）。承認者で「✍ 手書きサイン」を選んだときだけ出す |
| 大きさの合わせ方 `resizeCanvas` | 表示の幅・高さ × `devicePixelRatio` を canvas の画素にし、`setTransform(dpr…)` で CSS の座標で描く。線は `#0b2545`・太さ 2.5・端は丸（sign:121-131）。開いたとき（隠れているので 0×0）と、手書きを選んだときに呼ぶ |
| リサイズ | `window` の `resize` で **大きさを合わせ直し、書いたサインを消す**（sign:156）。スマホで画面を回す、アドレスバーが出入りする（推測）と消える |
| 書き方 | `pointerdown`（書き始め、`hasDrawn=true`）・`pointermove`・`pointerup`・`pointerleave`（終わり）。`setPointerCapture`・`pointercancel`・指 2 本の区別は無い。座標は `getBoundingClientRect` からの差（枠の 2px ぶんずれる） |
| クリア | 全部消して `hasDrawn=false` |
| 送る形 | `canvas.toDataURL('image/png')`＝`data:image/png;base64,…`。**背景は透明**（CSS の白は入らない）。サーバーは先頭を除いて PNG にする（`pdfExport.gs:97-98`） |
| 大きさ（推測） | PC（幅 452px・dpr 1〜2）で 452×260〜904×520 画素、スマホ（幅 287px 前後・dpr 3）で 861×780 画素。線だけの PNG なので base64 で数十 KB の見込み。VPS の API の受け取り上限は 1MB 以上にしておく |
| 電子印との分かれ方 | メールを選ぶと「電子印あり」、サインの欄を隠し、`signBase64=''`・`approverEmail=メール`。手書きを選ぶと `approverEmail=''`・`signBase64=PNG`。承認者の一覧は M_STAMP の全員（だれの印でも選べる） |
| 承認の後 | 「承認が完了しました。」、ボタン 2 つを隠し、`<meta http-equiv="refresh" content="1;url=…preview&reqId=…&approved=1">` を足して 1 秒後にプレビューへ。予備のリンクあり |
| 状態を見ない | プレビューは承認済のときだけ「承認する」を隠す。取消・却下でも出る。サインの画面は状態を見ない（`02` の K1） |

## 10. スマホ・タブレット（確認済み）

| 項目 | 中身 |
|---|---|
| viewport | 全画面 `width=device-width, initial-scale=1`。**sign だけ** `maximum-scale=1, user-scalable=no`（拡大できない） |
| 幅の切り替え | top・preview・somuAdmin・deptAdmin・admin は 640px 以下で、見出しを縦に並べ、選択肢を全幅、表の字を 12px。leaveStatus は 820px 以下と 520px 以下 |
| 横に流れる表 | somuAdmin・deptAdmin は `.table-scroll{overflow-x:auto}`、leaveStatus は表の最小幅 520px。**top の表（7 列）・admin の表は横の流れの枠が無い**（推測: 狭い画面ではみ出す） |
| 幅の固定 | top `.wrap` 最大 900px、somuAdmin・deptAdmin 1200px、admin 1100px、leaveStatus 1040px、preview の用紙 660px、sign のカード 500px、モーダル 540px（編集 420px）、deptAdmin のモーダル 480px |
| 印刷 | `@media print` は **1 つも無い**。プレビューを印刷する作りではない |
| 新しめの CSS | top:98 `label:has(input:checked)`（ラジオの枠の色）、leaveStatus:140 `color-mix()`。古いブラウザでは色が付かないだけ |
| 固定ヘッダー | top・somuAdmin・deptAdmin・admin は `position:sticky` と `backdrop-filter:blur(8px)` |

## 11. エスケープ漏れ（サーバーの値を逃がさず `innerHTML` に入れる箇所。確認済み）

preview と deptAdmin の表は逃がしています（`e()`・`escHtml()`）。leaveStatus も文字は逃がしています（`esc` は `'` を逃がさない）。値の多くは作業員マスタや固定の文字なので、今すぐの害は小さい見込みです（推測）。VPS では見た目を変えずに逃がします。

| 画面:行 | 入れる値 | 置き場所 |
|---|---|---|
| top:760-764 | `deptName・workerName・leaveKubun・halfDayType・leaveType・status` | 本日の休暇者の表 |
| top:787-791 | `reqId` | `onclick="…('reqId')"` と `href` |
| top:794-800 | `leaveDate・deptName・workerName・leaveKubun・halfDayType・leaveType・status` | 今後の休暇予定の表 |
| top:1173-1175 | `reqId` | 申請完了の文と `href` |
| top:1210-1233 | `leaveDate・daysAgo・deptName・workerName・leaveKubun・halfDayType・leaveType・submittedAt・reqId` | 承認待ちの表と `href` |
| top:613-619, 647, 651 | 休暇日・振替元・期間（サーバーの日付を選択肢から取ったもの） | 振替の検証の枠、選択肢 |
| sign:180-181 | `err.message` | 選択肢 |
| somuAdmin:255 | `daysAgo` | バッジ |
| somuAdmin:274・399 | `reqId` | `querySelector` の `[data-reqid="…"]`（`"` があると壊れる） |
| somuAdmin:471, 477-479, 480 | 回数と **`lastPaidDate`** | 有給の表 |
| admin:207・214 | `err.message` | エラーの表示 |
| admin:252-254 | `warnLevel` | バッジ |
| admin:260-264 | 回数と **`lastPaidDate`** | 表 |
| leaveStatus:320, 333-336, 379, 381-384, 417-420, 435 | 年度・回数・`m.state`（class 名）・`m.required` など数の値 | 本人ビュー・一覧 |
| 全画面の `escHtml`／`esc`（top 以外） | `'` を逃がさない。`escHtml(0)` は空になる（`if (!s) return ''`） | — |

## 12. 外部の読み込み（CSP に関係する。確認済み）

- 外部の CDN・Web フォント・画像は **1 つもありません**。フォントは端末の `system-ui`・`Noto Sans JP`・`Yu Gothic` を名前で指すだけです。
- 画像は leaveStatus:68 の選択肢の矢印（CSS の `url("data:image/svg+xml;…")`）だけ。CSP の `img-src` に `data:` が要ります。
- 絵文字（🔒 ✍ ✔ ⚠ ✅ 🔙 📅 🗓 🎉 💡 🔎）は文字です。
- `<script>` と `<style>` は各ページの中に直接書いてあります（`<script>` は 1 ページに 1 つ。no_auth は無し）。姉妹アプリと同じく、script は外の `.js` に出し、style と `style="…"` 属性（大量にあり、JS でも `el.style` を書く）は `style-src 'unsafe-inline'` で許す形になります。
- CSV は `Blob`＋`URL.createObjectURL`＋`<a download>` で端末に保存します（somuAdmin:500-534、admin:273-292）。BOM つき UTF-8、改行は `\n`。ファイル名は `有給取得管理_<年>年度.csv`（somuAdmin）と `有給取得状況_<年>.csv`（admin）。

## 13. 下書き `02-inventory-quick.md` との違い

| # | 下書きの記載 | 正しくは（確認済み） |
|---|---|---|
| 1 | 画面から呼ぶ関数は 27 | **合っています**（27 種類）。呼び出し箇所は 37 |
| 2 | on 属性は合計 50（top 27・somuAdmin 7・admin 5・sign 5・deptAdmin 3・preview・leaveStatus・no_auth 各 1） | **42 個**（top 26・somuAdmin 6・admin 4・sign 4・deptAdmin 2・preview・leaveStatus・no_auth 0）。うち 5 個は JS の文字で作るもの（5 節） |
| 3 | K11 エスケープ漏れは admin の「最終有給日」だけ | somuAdmin:480 の最終有給日、admin の `warnLevel`・`err.message`、**top の 3 つの一覧の全部**、sign の `err.message` も逃がしていない（11 節） |
| 4 | 「今日」をブラウザで決めるのは年度の選択肢（somuAdmin:173-176） | deptAdmin・admin・leaveStatus の年度、top の本日の休暇者・過去申請カレンダーの上限も同じ。top の過去申請の下限だけ JST に直している（8 節） |
| 5 | 2 節「REQ_ID を preview・sign に入れる」 | 合っています。加えて sign には PORTAL_URL を渡さない、leaveStatus は HTML にも `<title>` がある、no_auth の題が 2 種類ある |
| 6 | （記載なし） | preview の過去申請のバッジは **一度も出ない**。`api_getLeaveRequestById` が `isRetroactive`・`daysAgo` を返さない（`leaveRequest.gs:289-314`、preview:259） |
| 7 | （記載なし） | `api_getLookupData` は呼ぶが結果を使わない（top:505-507）。M_LOOKUP の「特別休暇理由」は画面に出ない |
| 8 | （記載なし） | 一括 2 次承認は戻り値の `ok` を見ない（somuAdmin:389）。失敗が 0 件なら終わりの知らせも再読込も無い |
| 9 | （記載なし） | 申請の入力が残ったまま送られる（14 節 D-6） |
| 10 | （記載なし） | somuAdmin・deptAdmin・admin は年度を変えても読み直さない（「更新」が要る） |
| 11 | （記載なし） | サインは画面の大きさが変わると消える（9 節） |
| 12 | （記載なし） | 承認待ち一覧の並びは `MM/dd HH:mm` の文字の比較で、年をまたぐと順が崩れる |
| 13 | （記載なし） | プレビューの理由の既定は「私用の為」、確認画面は「私用のため」（字が違う。今のまま） |

## 14. 移植で難しい箇所・オーナーに今のままか聞く癖

| # | 中身 | 場所 | 確かさ |
|---|---|---|---|
| D-1 | JS の文字に入る `REQ_ID`（URL から来る値）。GAS は文脈に合わせて逃がしているので、VPS は JSON の塊で渡す | preview:168・sign:108 | 確認済み |
| D-2 | on 属性 42 個の置き換え。somuAdmin の 2 個は `port-pages.js` の決まった形に合わない | 5 節 | 確認済み |
| D-3 | サインの canvas。`devicePixelRatio`、隠れた状態での大きさ、resize で消える、透明の PNG。今のまま移し、iPhone・iPad・Android で実機確認が要る | sign:115-156 | 確認済み（実機の動きは推測） |
| D-4 | `meta refresh` を JS で足す動き。VPS でも同じ動きになるが、CSP で止まらないか確かめる | sign:264-267 | 推測 |
| D-5 | top の画面の初期化は 8 本の呼び出しが同時に走り、#4→#5→#8→#10・#9 の順につながる。VPS で順番や速さが変わると、自動選択や件数の出方が変わり得る | top:455-549 | 確認済み |
| D-6 | 申請の入力が残る: 半日→午前を選んでから全日に戻すと `halfDayType='午前'` のまま送り、確認画面にも出る。振替元・特別理由・有給の理由も種類を変えても送る。サーバーは振替元をそのまま保存する（`leaveRequest.gs:213, 216-218`） | top:1149-1162 | 確認済み（直すかはオーナーに聞く） |
| D-7 | 申請完了の後、モーダルを閉じる手段が無い（リンクで移るだけ）。申請のモーダルの手順 9 の中身は書き換わるので、同じページで 2 件目は申請できない | top:1169 | 確認済み |
| D-8 | `resetModal` はページ中の全部のラジオを外す（編集モーダルの分も） | top:827 | 確認済み（害なし） |
| D-9 | 編集は休暇日を何日でも選べる（過去・休日も）。振替の半日も通る見込み | top:398、`leaveRequest.gs:862-901` | 確認済み（半日の振替は推測） |
| D-10 | 「今日」の基準がブラウザと JST で混ざる。日本の端末なら同じ | 8 節 | 確認済み |
| D-11 | `LEAVE_ITEMS` は今後の予定が 1 件以上のときだけ更新する | top:775-780 | 確認済み（害なし） |

## 15. 実機で必ず押す操作（受け入れ表の元）

「PC」はブラウザ、「スマホ」は iPhone と Android、「iPad」はタブレットです。

| # | 画面 | 操作 | 確かめること | 端末 |
|---|---|---|---|---|
| A-1 | top | 開く | 部署・承認待ち・予定が出る。権限に応じて「部署管理」「総務管理」、PORTAL_URL があればポータル | PC・スマホ |
| A-2 | top | 部署→作業員を選ぶ | 作業員の表示 `ID 氏名`、申請ボタンが押せる、予定が絞られる、本日の休暇者 | PC・スマホ |
| A-3 | top | 全日・有給（理由なし）で申請 | 確認画面の「私用のため」、完了表示、REQ_ID、2 つのリンク | PC・スマホ |
| A-4 | top | 半日・午前・特別休暇（理由あり）で申請 | 手順 3 が出る、種類に振替が無い | スマホ |
| A-5 | top | 振替休暇: 期間の中と外の振替元 | 候補が期間で絞られる、違反の赤い枠と「確認へ」が押せない、OK の緑の枠 | PC |
| A-6 | top | 過去申請: カレンダーを開き、前の月、日を選び、申請 | 年度の始まりより前・今日より後に行けない、選択の表示、確認ダイアログの文、承認待ちの「🔙 過去申請」 | PC・スマホ |
| A-7 | top | 選択肢と過去申請を交互に選ぶ | 片方を選ぶともう片方が消える | PC |
| A-8 | top | 入力を戻って変える（半日→全日など） | D-6 の動きが GAS と同じ | PC |
| A-9 | top | 「申請する」を素早く 2 回 | 1 件しかできない（ボタンが止まる） | PC |
| A-10 | top | 予定の「編集」: 日付と区分を変える（申請中・承認済の両方） | 確認の文、承認済の注意、更新後に一覧が変わる | PC・スマホ |
| A-11 | top | 予定の「削除」 | 確認の文、削除後に一覧から消える | PC |
| A-12 | preview | 承認待ちから開く | 様式、令和・曜日、時間帯、理由、REQ_ID、「承認する（サイン画面へ）」 | PC・スマホ・iPad |
| A-13 | preview | 存在しない reqId と reqId なし | 「申請データが見つかりません。（REQ_ID: …）」「REQ_IDが指定されていません。」 | PC |
| A-14 | sign | 電子印の人を選んで承認 | 「電子印あり」、確認、1 秒後にプレビュー、所属長の「承認済」と承認日 | PC |
| A-15 | sign | 手書きサインで承認 | 枠が出る、書ける、クリア、何も書かずに押すと止まる、承認後の画像 | **スマホ・iPad（指とペン）**・PC（マウス） |
| A-16 | sign | 書いた後に画面を回す・スクロールする | 消え方が GAS と同じ | スマホ・iPad |
| A-17 | sign | 「戻る」 | プレビューへ | スマホ |
| A-18 | somuAdmin | 開く（総務・管理者） | 2 次承認者の名前、待ち一覧、KPI、有給の表 | PC・iPad |
| A-19 | somuAdmin | 1 件を 2 次承認 | 確認の文、「処理中...」→「✓」→行が消える動き、件数、プレビューの部長の名前と承認日、PDF | PC |
| A-20 | somuAdmin | 全選択→一括 2 次承認（2〜3 件） | 確認の文、「承認中... (i/n)」、行が順に消える | PC |
| A-21 | somuAdmin | 年度を変える→「更新」、部署を変える | 年度だけでは変わらない、部署は変わる | PC |
| A-22 | somuAdmin・admin | CSV | ファイル名、Excel で文字化けしない、列 | PC |
| A-23 | deptAdmin | 管理者で開く | モーダル、全部署、承認者ごとのボタン、「承認者変更」 | PC |
| A-24 | deptAdmin | 部署承認者 1 人で開く | 自動で選ばれる、見出し「担当: 名前（部署）」 | PC・スマホ |
| A-25 | admin | 総務で開く・総務以外で開く | タブ、部署の絞り込み、CSV の出る・出ない | PC |
| A-26 | leaveStatus | 開く→部署→スタッフ→年度 ‹ › | 本人ビュー（KPI・タイムライン・今日の線・知らせ・予定・内訳）、一覧（並び順）、年度の端でボタンが止まる | PC・スマホ |
| A-27 | leaveStatus | 一覧の行を押す・Enter | その人の本人ビューへ、部署とスタッフの選択も変わる | PC |
| A-28 | no_auth | 権限の無い人で deptAdmin・somuAdmin・admin | 3 種類の文言、題、「トップ画面へ戻る」 | PC |
| A-29 | 全画面 | スマホの幅で見る | 640px／820px の切り替え、表がはみ出すか（top・admin）が GAS と同じ | スマホ |

## 確かめること

- Settings の `APP_URL`・`PORTAL_URL` の実際の値（末尾の形）。
- 古い Safari で preview の日時（`yyyy-MM-dd HH:mm:ss`）が正しく出るか（A-12 を iPhone で）。
- サインの画像の実際の大きさ（A-15 の後に Drive の `sign_<REQ_ID>.png` を見る）。
- D-6・D-9・一括 2 次承認の戻り値を見ない動きを、今のまま残すかどうか（オーナーに聞く）。
