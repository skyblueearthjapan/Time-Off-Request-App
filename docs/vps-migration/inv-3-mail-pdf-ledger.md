# 棚卸し 3: メール・PDF・サイン画像・台帳（内部用の入口の下調べ）

作成: 2026-11-26。担当: 棚卸し 3（メール・PDF・サイン・台帳）。読むだけで、コードもシートも変えていません。
対象: `src/mail.gs`・`warnMail.gs`・`notifyMail.gs`・`pdfExport.gs`・`ledger.gs`（全行）、`leaveRequest.gs` の承認・2 次承認・一括・編集・取消、関係する `config.gs`・`menu.gs`・`masters.gs:122-186`・`sign.html`・`somuAdmin.html`・`appsscript.json`。
前例: 残業の `src/Internal.gs`・`src/Cutover.gs`・`vps-app/api/app/services/gas_jobs.py`・`docs/vps-migration-plan.md` 4-6〜4-7、罠 `07-pitfalls.md`。

印: **確認済み**＝コードで確かめた。**推測**＝実物（本番のシート・Drive・実行数の画面）で未確認。行番号は `src/` からの `ファイル:行`。

---

## 0. 先に伝えたいこと（上位 8 件）

| # | 中身 | 確かさ |
|---|---|---|
| T1 | **このアプリの Web アプリは `access: DOMAIN`**（`appsscript.json:6-9`）。残業は `ANYONE_ANONYMOUS` だったので、VPS が token だけで `doPost` を呼べた。このアプリでは VPS から匿名で `/exec` を呼べない。マニフェストは版ごとに 1 つなので、内部用のデプロイだけ匿名にすることもできない（同じ版の画面用デプロイも匿名になり、`Session.getActiveUser()` が空になって権限の判定が壊れる見込み）。**内部用の入口の呼び方を決め直す必要があります**（7 節 D1） | 設定は確認済み。匿名にしたときの動きは推測 |
| T2 | **承認完了メールが今、本当に届いているか疑わしい**。`GmailApp.sendEmail` を使うが、マニフェストのスコープは `gmail.send` だけ（`appsscript.json:13`）。`GmailApp` は `https://mail.google.com/` を求める見込みで、失敗しても `leaveRequest.gs:538-542` が握りつぶす | 推測（実行数の画面で「メール通知エラー」のログ、imaizumi@ の送信済みを見る） |
| T3 | **今のトリガーが本当に `syncAllMasters` 1 本か**。通知 3 本（`notifyTodayLeaves_` ほか）と警告（`checkAndSendWarnMails_`）を作る関数は残っており（`menu.gs:392-423`）、呼び出しをコメントにしただけ（`menu.gs:347-349`）。昔作ったトリガーが残っていれば、今も総務へメールが出ている | 推測（トリガーの画面。05 の C-9） |
| T4 | 編集・取消の後処理（下書きの `afterUnapprove`）は、**編集の前の休暇日と前の `PDF_FILE_ID` を引数で渡す必要があります**。VPS の行の反映が先に休暇日を新しい日へ、PDF の列を空へ書き換えるので、GAS がシートから読むと台帳の別の年度タブを探し、捨てる PDF も分からなくなります | 確認済み（`ledger.gs:314-322`、`leaveRequest.gs:946-954`） |
| T5 | **台帳と行削除の関数は失敗を握りつぶします**（`ledger.gs:126/264-266`、`314/337-339`）。書けたかは呼び出し側で台帳を読み直さないと分かりません。台帳が開けないと新しい台帳を作り、Settings に行を足します（`ledger.gs:7-45`） | 確認済み |
| T6 | **再試行で二重になるもの**: 承認完了メール（2 通）、サイン画像（同名の PNG が 2 つ）、2 次承認の PDF（2 つ。前の PDF はシートの値で捨てるので、記録しそこねた分は残る）。今の関数には「もうある」の判定がありません（残業の `already` に当たるものが無い） | 確認済み（コード） |
| T7 | 2 次承認の PDF が失敗すると、**2 次承認済みのまま PDF の列が「ゴミ箱へ移した前の PDF」を指し続け**、台帳もその URL で書かれます（`leaveRequest.gs:694-720`）。2 次承認待ちの一覧からも消えるので、画面からはやり直せません | 確認済み（コード） |
| T8 | 移転モードで止める所は `getDb_()` だけでは足りません。`doGet` は先頭で `getSettings_()` を呼ぶので（`menu.gs:5`）、判定を `getSettings_` より前に置かないと移転ページの代わりにエラーが出ます。`setupAllSheets` は `getDb_()` を通らず `SpreadsheetApp.getActive()` を直接呼びます（`menu.gs:118`） | 確認済み |

---

## 1. メール

### 1-1. 種類の一覧

| 種類 | 送る関数 | きっかけ | 宛先 | 送らない条件 | 今動いているか |
|---|---|---|---|---|---|
| 承認完了通知 | `sendApprovalNotification_(reqId)`（`mail.gs:6-24`） | 1 次承認のたび（`leaveRequest.gs:538-542`）。2 次承認・編集・取消では送らない | Settings `MAIL_TO` の値をそのまま（カンマ区切りの複数も可） | `MAIL_TO` が空（`mail.gs:9-12`）、申請が読めない（`mail.gs:14-18`。読めないときは 1.5 秒×4 回待ってから） | 呼ばれてはいる（確認済み）。**届いているかは T2 のとおり疑わしい**（推測） |
| 有給警告 | `checkAndSendWarnMails_()`（`warnMail.gs:11-57`）。包み `checkAndSendWarnMails()`（`warnMail.gs:4-6`） | シートのメニュー「有給警告メール送信（手動）」（`menu.gs:109`）。誰でも `google.script.run` からも呼べる | `MAIL_TO` | 今日が 7/1・9/1・1/1・2/16 でない（`warnMail.gs:16-28`、`config.gs:61-66`）。未達者 0 人（`warnMail.gs:39-42`）。`MAIL_TO` が空（`warnMail.gs:45-50`） | メニューからだけ（確認済み）。トリガーを作る `setupWarnMailTrigger_` は呼ばれない（`menu.gs:348`）。実物のトリガーは推測 |
| 本日の休暇者 | `notifyTodayLeaves_()`（`notifyMail.gs:235-331`） | 毎日 7 時台（作る関数 `menu.gs:409-412`） | Settings `SOMU_EMAILS` | 該当 0 件、`SOMU_EMAILS` が空 | **作る関数が呼ばれない**（`menu.gs:349`）。末尾 `_` なので画面からも呼べない。実物のトリガーは推測（T3） |
| 1 次未承認 | `notifyPendingApproval1_()`（`notifyMail.gs:45-133`） | 毎日 17 時 15 分ごろ（`menu.gs:414-417`） | 部署ごとに M_DEPT_APPROVERS の B 列＋`SOMU_EMAILS`（重なりを除く） | 該当 0 件、宛先 0 件の部署は飛ばす | 同上 |
| 2 次未承認 | `notifyPendingApproval2_()`（`notifyMail.gs:141-228`） | 毎日 18 時ごろ（`menu.gs:419-422`） | `SOMU_EMAILS` | 該当 0 件、`SOMU_EMAILS` が空 | 同上 |

差出人はどれも GAS の実行者（Web アプリは `USER_DEPLOYING`＝デプロイした人。imaizumi@ の見込み＝推測）。送信はすべて `GmailApp.sendEmail(宛先, 件名, 本文)` の書式なしの文です。

### 1-2. 承認完了通知（`mail.gs`）

件名（`mail.gs:20`）: `'【休暇届 承認完了】' + 氏名 + ' ' + 休暇種類`

本文（`mail.gs:39-68`。改行は `\n`）:

```
【休暇届 承認完了通知】

以下の休暇届が承認されました。

■ 申請情報
  REQ_ID: <REQ_ID>
  部署: <部署名>
  氏名: <作業員名>
  休暇区分: <休暇区分>（<半日区分>）   ← 半日区分が空なら（）ごと付けない
  休暇日: <yyyy/MM/dd>                 ← 休暇日が空なら空
  休暇種類: <休暇種類>
  振替元出勤日: <yyyy/MM/dd>            ← 休暇種類が「振替休暇」で振替元があるときだけ
  特別理由: <特別理由>                  ← 特別理由が空でなければ（休暇種類は問わない）

管理画面: <APP_URL>?page=admin         ← APP_URL が空なら、この行と次の空行を出さない

※本メールは自動送信です。
```

- 中身は送る時点のシートの行（`api_getLeaveRequestById`）。休暇日は `'yyyy-MM-dd'` の文字を `new Date()` で読み直して JST で書く（確認済み。UTC 0 時＝JST 9 時なので日付はずれない）。
- リンクは Settings の `APP_URL`（GAS の `/exec` の見込み）。移転モードでは移転ページが `page=admin` を引き継いで VPS へ案内する形になる（推測。残業 4-11 と同じ考え方）。
- K1（承認済み・取消の申請をもう一度承認できる）のとき、そのたびにもう 1 通送る（確認済み）。

### 1-3. 有給警告（`warnMail.gs`）

件名（`warnMail.gs:52`）: `'【有給取得警告】有給取得義務未達者一覧 ' + yyyy/MM/dd`

本文（`warnMail.gs:105-162`）:

```
【有給取得義務 未達者一覧】

年度: <FY>年度（<FY>/3/16 - <FY+1>/3/15）
確認日: <yyyy/MM/dd>

現在の必要有給回数: <n>回（<基準の名前>基準）   ← 期限（6/30・8/31・12/31・2/15・3/15 の 23:59:59）を過ぎたもののうち最後

■ <部署名>                                     ← 部署名の文字順（sort()）
  <!!!|!!|!> <氏名>  有給: <n>回  （<警告レベル>）  最終: <最終有給日>   ← 危険=!!!、警告=!!、ほか=!。最終が空なら「最終:」ごと出さない

対象者数: <n>名

管理画面: <APP_URL>?page=admin                 ← APP_URL が空なら出さない

※本メールは自動送信です。
```

- 送る前に `rebuildLeaveSummary_()` で V_LEAVE_SUMMARY を書き直す（`warnMail.gs:33`）。**MAIL_TO が空でも書き直しは起きる**（確認済み。順番が 33 行→45 行）。
- 対象は V_LEAVE_SUMMARY の「年度＝今日の年度」で「警告レベル」が空でない行（`warnMail.gs:62-96`）。並びは「部署名＋氏名」の `localeCompare('ja')`。
- 「最終有給日」は `normalize_(セルの値)`。セルが日付型なら `String(Date)` の長い英語の文字になる（02 の K5 と同じ疑い。推測）。
- 包み `checkAndSendWarnMails` のコメントは「日付チェックをスキップして強制送信」だが、実際は日付で止まる（02 の K8。確認済み）。

### 1-4. 動いていない見込みの 3 通（`notifyMail.gs`）

| 関数 | 件名 | 本文の要点 | 本文のリンク |
|---|---|---|---|
| `notifyPendingApproval1_` | `'【休暇届 未承認通知】' + 部署 + ' ' + n + '件が未承認です'` | 見出し「【休暇届 1次承認未完了通知】」、案内 2 行、`■ 部署（n件）`、各行 `  - 氏名　yyyy/MM/dd　種類（区分）`。対象は申請中で休暇日が今日以降（日付型でないものも含む） | `承認はこちらから（申請TOP）:` の次の行に `APP_URL?page=top` |
| `notifyPendingApproval2_` | `'【休暇届 2次承認待ち】' + n + '件が2次承認待ちです'` | 見出し「【休暇届 2次承認未完了通知】」、部署ごと、各行の末尾に `　1次承認: yyyy/MM/dd HH:mm`、`合計: n件`。対象は承認済で `APPROVED_AT2` が空（休暇日は問わない） | `2次承認はこちらから（総務管理画面）:` ＋ `APP_URL?page=somuAdmin` |
| `notifyTodayLeaves_` | `'【本日の休暇者】' + yyyy/MM/dd + ' ' + n + '名'` | 見出し「【本日の休暇者一覧】」、`日付:`、部署ごと `  - 氏名　種類　区分（半日区分）`、`合計: n名`。対象は承認済で休暇日が今日（日付型のみ） | `PDFフォルダ:` ＋ `https://drive.google.com/drive/folders/<PDF_FOLDER_ID>`、`総務管理画面:` ＋ `APP_URL?page=somuAdmin` |

これらは統合業務メール（別の GAS）が同じ役目を担っている見込みです（`menu.gs:347` のコメント「メール送信は統合メール通知システムに移行済み」）。移行では作り直しません。ただし T3 のとおり、実物のトリガーが残っていないかを初日に見ます。

---

## 2. PDF

### 2-1. 作る場面・作らない場面（確認済み）

| 場面 | 作るか | 関数・行 | 前の PDF | PDF の後の台帳 |
|---|---|---|---|---|
| 申請 | 作らない | `api_submitLeaveRequest` | — | — |
| 1 次承認 | **作らない**（`leaveRequest.gs:527-528`） | `api_approveLeaveRequest` | 触らない | upsert（PDF 列は前の値のまま） |
| 2 次承認（1 件。画面の一括も 1 件ずつこれを呼ぶ） | 作る | `api_approveLeaveRequest2`（`leaveRequest.gs:678-711`）。ロックの中 | シートの `PDF_FILE_ID` をゴミ箱へ（`:694-701`。失敗してもログだけ） | **PDF の成否に関係なく** upsert（`:715-720`） |
| 2 次承認の一括 API | 作る | `api_approveLeaveRequest2Batch`（`:819-843`）。**ロックの外**。画面からは呼ばれない | 同上（`:824-826`） | 同上 |
| 編集（承認済みから申請中に戻すとき） | 作らない。前の PDF を捨てる | `api_editLeaveRequest`（`:945-951`） | ゴミ箱へ→PDF の 2 列を空に | 行を消す（`:954`） |
| 取消 | 作らない。前の PDF を捨てる | `api_deleteLeaveRequest`（`:986-992`） | 同上 | 行を消す（`:997`） |
| 手で試す | 作る | `test_generatePdf`（`pdfExport.gs:7-18`。固定の REQ_ID と承認者。誰でも呼べる） | 捨てない。シートにも書かない | 書かない |

### 2-2. 作り方の手順と待ち時間（`generateLeavePdf_`、`pdfExport.gs:124-180`）

| 順 | すること | 行 | 待ち・所要 |
|---|---|---|---|
| 1 | `api_getLeaveRequestById(reqId)` で行を読む。無ければ例外 | `:125-126`、`leaveRequest.gs:239-323` | 見つからないと 1.5 秒×4 回（約 6 秒）待ってから null |
| 2 | Settings の `PDF_FOLDER_ID`。空なら例外 | `:128-130` | — |
| 3 | 一時スプレッドシート `TMP_PDF_<REQ_ID>_<yyyyMMdd_HHmmss>` を新しく作り、1 枚目を「休暇届」に | `:138-140` | 数秒（推測） |
| 4 | `buildLeavePdfSheet_` でセルを組み、印を `insertImage` | `:143`、`:187-542` | 印 1 つにつき `flush`＋0.5 秒（`:467-468`、`:511-512`）、サムネイルの取得（`UrlFetchApp`） |
| 5 | `flush` → 4 秒 → `flush` → 1 秒（画像の反映待ち） | `:144-148` | 5 秒 |
| 6 | 休暇日のフォルダ `<PDF_FOLDER_ID>/yyyy.MM.dd` を探すか作る | `:151-152`、`:78-84` | — |
| 7 | ファイル名を決め、`/export?format=pdf&gid=…&portrait=true&size=A4&fitw=true&上下左右0.4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false` を OAuth トークン付きで取る。200 以外は例外 | `:155-161`、`:732-757` | 数秒（推測） |
| 8 | 日付フォルダに `createFile(pdf).setName(名前)` | `:164` | — |
| 9 | 返す `{ok:true, pdfFileId, pdfUrl, pdfName, folderId}` | `:166-172` | — |
| 10 | `finally` で一時スプレッドシートをゴミ箱へ（成功でも失敗でも） | `:173-179` | — |

- 1 件の所要は 15 秒ほど（02・05 の見込み。推測。05 の C-7 で実行数の画面から測る）。
- **シートに書かない**。PDF の 2 列（`PDF_URL`・`PDF_FILE_ID`）は呼び出し側が書く（`leaveRequest.gs:705-708`）。
- ロックを取らない（コメント `:132-133`）。呼び出し側（2 次承認）がスクリプトのロックを持っている前提。

### 2-3. 中身の項目と書式（`buildLeavePdfSheet_`、確認済み）

範囲は A1:J30（列幅 A=12・B,C=80・D〜I=68・J=12。31 行目から下と K 列から右は隠す `:534-540`）。全体のフォント `Noto Sans JP, sans-serif`、色は `PDF_COLORS_`（`:107-117`）。

| 行 | 項目 | 値の決め方 |
|---|---|---|
| 2 | 題 | 「休 暇 届」24pt 太字、紺地に白（`:215-220`） |
| 4 | 届出日（F4:I4 右寄せ 10pt） | 申請日時 → `令和<年-2018>年<月>月<日>日 届出`。申請日時が空なら**作った日**（`:231-235`） |
| 6 | 所属部署 | 部署名 |
| 7 | 氏名 | 作業員名 |
| 8 | 区切り | 「届 出 内 容」9pt |
| 9 | 休暇日 | `令和<n>年<月>月<日>日（<曜>曜日）`（`:272-276`。曜は 日〜土） |
| 10 | 休暇時間 | 半日休で午前 → `午前半休　8:05 ～ 12:15`、半日休でそれ以外（午後・空） → `午後半休　13:00 ～ 17:10`、全日休ほか → `終日　8:05 ～ 17:10`（`:284-293`。「～」は全角の波） |
| 11 | 休暇区分 | 区分＋（半日区分があれば `（午前）` など） |
| 12 | 休暇種類 | 休暇種類 |
| 13 | 振替元出勤日 | 振替休暇で振替元があれば `令和<n>年<月>月<日>日（<曜>）`（「曜日」でなく 1 字）、ほかは `―`（`:320-327`） |
| 14-15 | 理由・詳細 | `pdfBuildReason_`（`:570-581`）: 特別休暇で特別理由あり→特別理由／有給消化（私用）→有給詳細、空なら「私用のため」／振替休暇→「振替休暇」／ほか→休暇種類／種類も空→「―」。11pt、上寄せ、折り返し |
| 17 | 宣言 | 「上記のとおり休暇を届け出いたします。」 |
| 19-20 | 届出者 | 「氏名」＋作業員名（印なし） |
| 22 | 承認欄 | 「承 認 欄」 |
| 23 | 見出し | D23:F23「部　長」、G23:I23「所 属 長」 |
| 24-27 | 承認印 | 下の 2-4 |
| 28 | 承認日 | 2 次承認者のメールが無いとき: D28（部長の欄）に 1 次承認日、G28 は空。あるとき: G28（所属長）に 1 次承認日、D28（部長）に 2 次承認日（`APPROVED_AT2`、空なら**作った日**）（`:429-438`、`:519-528`）。形は `<年-2018>/<月>/<日>`（例 `8/11/27`。02 の K14） |
| 30 | 注記 | 「※後報の場合は、理由を詳しく記載すること。」8pt |

令和は `年 - 2018` の計算だけ（元年・平成の扱いなし）。`toReiwa_`（`:68-73`）は定義だけで使われていません（確認済み）。

### 2-4. 印の探し方と、無いときの表示

| 欄 | 探し方 | 無いとき |
|---|---|---|
| 所属長（G24、1 次） | 引数 `approverEmail`（2 次承認では `APPROVED_BY1_EMAIL` の列。空ならシートから読み直す `leaveRequest.gs:680-692`）で M_STAMP の A 列を小文字で完全一致 → B 列の stampFileId（`masters.gs:122-157`）→ `drive.google.com/thumbnail?id=…&sz=w300`、だめなら `lh3.googleusercontent.com/d/…=s300` を Bearer 付きで取り、1,000〜1,500,000 バイトなら使う（`pdfExport.gs:663-698`）→ `insertImage(blob, 7, 24, 40, 3)`、幅 120・高さ 70（`:469-471`） | ① 印が入らず、手書きサインの URL があれば、その Drive ファイルの元画像を同じ位置・大きさで入れる（`:484-491`、`:708-725`。サムネイルにしない） ② 印もサインも無ければ G24 に赤 8pt で「（承認印取得不可）」（`:493-498`） ③ 印が登録されていてサムネイルの取得に失敗し、サインも無いときも ② になる |
| 部長（D24、2 次） | 引数 `approver2Email`（画面はログイン中のメール。`somuAdmin.html:212`）で同じ探し方 → `insertImage(blob, 4, 24, 40, 3)`（`:503-517`） | 何も入れない（文字も出さない）。例外は握る（`:529-531`） |

- 1 次承認で「手書きサイン」を選ぶと `approverEmail` は空で送られ（`sign.html:203-205`、`:284`）、`APPROVED_BY1_EMAIL` も空になる。このとき PDF はサインの画像を使う。
- 承認者の一覧（`api_getApproverList`）に出る人は、stampFileId が空でも「電子印あり」と表示され、サインは送られない（`sign.html:216-219`）。このとき PDF は ② の「（承認印取得不可）」になる（確認済み・コード）。
- `getResizedStampBlob_`（`pdfExport.gs:587-650`。Slides を使う縮小）はどこからも呼ばれていません（確認済み）。

### 2-5. ファイル名・フォルダ・前の PDF

| 項目 | 中身 |
|---|---|
| ファイル名 | `<休暇日 yyyyMMdd>_<部署名>_<氏名>_休暇届.pdf`。`\ / : * ? " < > |` は `_` に（`:155-158`）。同じ人・同じ日の 2 件（午前と午後の半休など）は同じ名前の別ファイルになる |
| フォルダ | `PDF_FOLDER_ID` の下の `yyyy.MM.dd`（**休暇日**。承認日ではない）。同じ名前のフォルダが無ければ作る。同時に 2 つ動くと同じ名前のフォルダが 2 つできうる（今はロックで防いでいる） |
| 前の PDF | 2 次承認: 作る前に `PDF_FILE_ID` のファイルをゴミ箱へ。**列は空にしない**（成功すれば新しい値で上書き）。編集・取消: ゴミ箱へ移し、2 列を空にする |
| 一時スプレッドシート | 実行者のマイドライブに作られ、`finally` でゴミ箱へ。6 分の上限で実行ごと止まったときだけ残る（推測） |

### 2-6. 失敗したときに残る半端な状態（確認済み・コード）

| 失敗した所 | 残るもの |
|---|---|
| 2 次承認で PDF が例外 | 2 次承認の 3 列（`APPROVED_BY2`・`APPROVED_BY2_EMAIL`・`APPROVED_AT2`）は書き済み。前の PDF はゴミ箱へ移した後なのに、`PDF_URL`・`PDF_FILE_ID` は前の値のまま。台帳はその URL で upsert される。2 次承認待ちの一覧から消えるので、画面からやり直せない |
| 日付フォルダを作った後に export が失敗 | 空の日付フォルダ（害は無い） |
| `createFile` の後、シートへ書く前に止まる | Drive に PDF が 1 つ増え、シートは前の値（どこからも指されない PDF） |
| 前の PDF をゴミ箱へ移せない（持ち主が別など） | 前の PDF が残ったまま新しい PDF ができる |
| 1 次承認でサインの保存が例外 | サイン URL を空のまま承認済になる（`leaveRequest.gs:508-513`）。PDF は「（承認印取得不可）」になりうる |
| 1 次承認で状態を書く前に例外（列が無いなど） | サインの PNG だけが Drive に残る |

---

## 3. サイン画像

| 項目 | 中身 | 確かさ |
|---|---|---|
| 受け取る形 | `sign.html` の canvas（高さ 260px、幅は画面に合わせ、`devicePixelRatio` 倍。`sign.html:121-126`）を `toDataURL('image/png')` した `data:image/png;base64,…` の文字。`api_approveLeaveRequest(reqId, signBase64, approverEmail)` の 2 つ目の引数 | 確認済み |
| 送る条件 | 「手書きサイン」を選び、1 度でも描いたとき（`sign.html:249-252`）。電子印の人を選んだときは空文字 | 確認済み |
| 大きさ | 実物は 10〜14KB（03 の記録）。高解像度の端末では横 1,300px 程度の PNG になりうる（推測）。背景は透明（`clearRect`） | 推測 |
| 保存 | `saveSignImage_(reqId, base64, 休暇日)`（`pdfExport.gs:89-104`）: 頭の `data:image/…;base64,` を外して PNG の Blob `sign_<REQ_ID>.png` を作り、PDF と同じ `<PDF_FOLDER_ID>/<休暇日 yyyy.MM.dd>` に `createFile`。休暇日が日付型でなければ `new Date(文字)`、空なら今日 | 確認済み |
| 保存しない条件 | `PDF_FOLDER_ID` が空なら保存せず `''` を返す（承認は続く）。保存の例外も握って `''`（`leaveRequest.gs:508-513`） | 確認済み |
| URL を書く列 | `サイン画像URL`（英語の見出しなら `SIGN_IMAGE_URL`）。値は `file.getUrl()`（`https://drive.google.com/file/d/<id>/view?…` の形の見込み）。**1 次承認のたびに上書き**（電子印を選べば空で上書き）（`leaveRequest.gs:519`） | 確認済み（URL の形は推測） |
| 電子印を選んだとき | サインは送らず保存もしない。列は空で上書き。`APPROVED_BY1_EMAIL` に選んだ人のメール | 確認済み |
| 後始末 | 編集で承認が外れると列だけ空にする（`leaveRequest.gs:940`）。**ファイルは消さない**。取消では列もファイルも残す。もう一度承認すると同じ名前の PNG がもう 1 つできる | 確認済み |
| 使う場所 | 2 次承認の PDF で、1 次の印が入らなかったときの代わり（2-4）。画面（preview）には出ない（`preview.html` に `signImageUrl` の参照なし） | 確認済み |

---

## 4. 台帳（`ledger.gs`）

### 4-1. 開き方（`getLedgerSs_`、`ledger.gs:7-45`）

| 場合 | 動き |
|---|---|
| Settings の `LEDGER_SSID` があり開ける | それを使う（同じキーが 2 行あれば下の行。`config.gs:113-117`） |
| `LEDGER_SSID` が開けない（消えた・権限が無い・一時的な失敗も含む） | 警告を出して**新しく作る** |
| `LEDGER_SSID` が空 | 新しく作る |
| 新しく作る手順 | `PDF_FOLDER_ID` が空なら null を返して何もしない。あれば「休暇届台帳（総務確認用）」を作り、PDF フォルダへ入れてマイドライブの直下から外す。**M_SYSTEM_SETTING の末尾に `['LEDGER_SSID', 新ID, '総務確認用台帳スプレッドシートID（自動生成）']` を足す**（`:35-38`）。1 枚目を `_tmp` に改名 |

一時的に開けなかっただけでも新しい台帳ができ、以後は Settings の下の行（新しい方）が使われ続けます（確認済み・コード）。

### 4-2. 年度タブ（`getLedgerFySheet_`、`ledger.gs:53-117`）

- 名前は `<年度>年度`（例 `2026年度`）。年度は休暇日から `computeFiscalYear_`（3/16 始まり、`leaveRequest.gs:34-44`）。
- 無ければ**先頭に**作る（`insertSheet(名前, 0)`）。見出し 20 列を書き、太字・紺地 `#1a365d`・白字・10pt・中央・上下中央、1 行目を固定、高さ 32、列幅を設定。`_tmp` タブがあれば消す。

### 4-3. 20 列の中身と決め方（`updateLedger_`、`ledger.gs:125-267`）

| 列 | 見出し | 値（`api_getLeaveRequestById` の行から） |
|---|---|---|
| A | 管理番号 | REQ_ID |
| B | 作業員ID | 作業員ID |
| C | 部署 | 部署名 |
| D | 氏名 | 作業員名 |
| E | 休暇日 | `yyyy/MM/dd` の文字 |
| F | 曜日 | 日〜土の 1 字 |
| G | 休暇区分 | 区分 |
| H | 半日区分 | 半日区分。空なら `―` |
| I | 勤務時間帯 | 全日 `8:05 〜 17:10（終日）`、午前 `8:05 〜 12:15（午前半休）`、半日休でそれ以外 `13:00 〜 17:10（午後半休）`（「〜」は PDF と違う字＝U+301C） |
| J | 休暇種類 | 休暇種類 |
| K | 理由・詳細 | PDF とほぼ同じだが、有給の既定は「**私用の為**」（PDF は「私用のため」）、種類が空なら空（PDF は `―`） |
| L | 振替元出勤日 | 振替休暇で振替元があれば `yyyy/MM/dd`、ほかは `―` |
| M | 申請日 | 申請日時の `yyyy/MM/dd` |
| N | 1次承認日 | 承認日時の `yyyy/MM/dd` |
| O | 1次承認者 | `APPROVED_BY1_EMAIL`（**メール**。手書きサインなら空） |
| P | 2次承認日 | `APPROVED_AT2` の `yyyy/MM/dd` |
| Q | 2次承認者 | `APPROVED_BY2`（**名前**。M_STAMP の C 列、無ければメール） |
| R | 承認状態 | 承認状態の文字 |
| S | 有給累計取得回数 | `countPaidLeaves_(年度, 作業員ID)`（`ledger.gs:275-308`）: T_LEAVE_REQUEST の**全行**から、承認済・同じ作業員・「有給消化（私用）」・休暇日の年度が同じ、を数える（半日も 1）。書いた時点の値で、ほかの行の値は後から直さない |
| T | PDF | `PDF_URL`。空でなければ `=HYPERLINK("<URL>","PDF")` の式で上書き |

値は `setValues` で書く（`yyyy/MM/dd` の文字はシートが日付に変える見込み＝推測）。

### 4-4. upsert の探し方と書式

- 探す範囲は**休暇日の年度タブの A 列だけ**（2 行目から最終行）。REQ_ID と `normalize_` で完全一致した最初の行を上書き、無ければ最終行の次に書く（`:224-244`）。
- 書いた行だけ 10pt・上下中央、背景は行番号が偶数なら `#f8fafc`、奇数なら白（`:246-254`）。行を消すと下の行の縞がずれる（今のまま）。
- 最後に `flush`。**全体を `try/catch` で囲み、失敗はログだけ**（`:264-266`）。申請が読めないときも黙って戻る（`:127-131`）。

### 4-5. 行を消す場面（`removeLedgerRow_(reqId, leaveDate)`、`ledger.gs:314-340`）

| 場面 | 渡す休暇日 |
|---|---|
| 編集で承認済みから申請中へ戻すとき（`leaveRequest.gs:954`） | **編集前**の休暇日（`req.leaveDate`。シートを書き換える前に読んだ値） |
| 取消（`leaveRequest.gs:997`）。申請中の取消でも呼ぶ（行が無ければ何もしない） | 取消時の休暇日 |

- 休暇日の年度タブを下から探し、最初に見つけた 1 行だけ `deleteRow`。タブが無ければ何もしない。失敗はログだけ。
- 却下（`STATUS.REJECTED`）は定数だけで、設定する関数がありません（確認済み）。

---

## 5. 関数ごとの入出力と、内部用の action への切り出し

### 5-1. 一覧

| 関数 | 引数 | 読むもの | 書くもの | 返すもの | ロック | 内部用から |
|---|---|---|---|---|---|---|
| `saveSignImage_`（`pdfExport.gs:89`） | reqId, base64, leaveDate | Settings `PDF_FOLDER_ID` | Drive（日付フォルダ・PNG） | URL か `''` | なし | **そのまま呼べる**。URL を列に書くのは別に要る（今は `leaveRequest.gs:519`） |
| `sendApprovalNotification_`（`mail.gs:6`） | reqId | Settings `MAIL_TO`・`APP_URL`、T_LEAVE_REQUEST の行 | Gmail | なし（送らなかった理由も返さない） | なし | **そのまま呼べる**。送ったかどうかは呼び出し側で記録が要る |
| `generateLeavePdf_`（`pdfExport.gs:124`） | reqId, 1 次のメール, 2 次のメール | 行（全列）、Settings `PDF_FOLDER_ID`、M_STAMP の A・B 列、Drive の印 | Drive（一時シート・日付フォルダ・PDF） | `{ok, pdfFileId, pdfUrl, pdfName, folderId}`、失敗は例外 | 呼び出し側が持つ前提 | **そのまま呼べる**。前の PDF を捨てる・2 列を書くのは別（今は `leaveRequest.gs:694-708`） |
| `updateLedger_`（`ledger.gs:125`） | reqId | 行、T_LEAVE_REQUEST の全行（累計）、Settings `LEDGER_SSID`・`PDF_FOLDER_ID` | 台帳（行・年度タブ）、まれに新しい台帳と Settings の行 | なし（失敗も黙る） | なし | 呼べるが、**前に台帳を開けるかの確かめ**と、**後に行があるかの読み直し**が要る |
| `removeLedgerRow_`（`ledger.gs:314`） | reqId, leaveDate | Settings | 台帳の行を消す | なし（失敗も黙る） | なし | 同上。leaveDate は**引数で受ける** |
| `getLedgerSs_`（`ledger.gs:7`） | — | Settings | 新しい台帳・Settings（開けないとき） | Spreadsheet か null | なし | 直接は使わない（勝手に作るため） |
| `checkAndSendWarnMails_`（`warnMail.gs:11`） | — | V_LEAVE_SUMMARY、Settings | **V_LEAVE_SUMMARY を書き直す**、Gmail | なし | なし | **切り分けが要る**（集計の書き直しは VPS の役目と重なる） |
| `api_getLeaveRequestById`（`leaveRequest.gs:239`） | reqId | 行 | なし | 行の写し（`rowNo` つき）か null | なし | 呼べるが、無いと約 6 秒待つ。先に行の有無を自分で確かめる |
| `api_approveLeaveRequest`（`leaveRequest.gs:471`） | reqId, base64, メール | — | 状態・承認日時・サイン URL・1 次メール、集計、メール、台帳 | `{ok}` | 60 秒 | **呼ばない**。状態の書き込みが VPS と重なる |
| `api_approveLeaveRequest2`（`:619`） | reqId, メール | — | 2 次の 3 列、PDF、台帳 | `{ok}` | 60 秒 | **呼ばない**。`isSomu_()`/`isAdmin_()` がログイン中の人を見るので、内部用の実行では通らない見込み（推測） |
| `api_editLeaveRequest`（`:852`）・`api_deleteLeaveRequest`（`:970`） | reqId, … | — | 行の更新・PDF・台帳・集計 | `{ok}` | 20 秒 | **呼ばない**。後半（PDF・台帳）だけを内部用で再現する |
| `notify*_`（`notifyMail.gs`） | — | T_LEAVE_REQUEST、M_DEPT_APPROVERS、Settings | Gmail | なし | なし | 使わない（統合業務メールが担う） |

### 5-2. 内部用の入口の中で守ること（このアプリ特有）

- 今の関数はロックを取りません（`generateLeavePdf_` のコメントのとおり、呼び出し側が持つ前提）。残業の `Internal.gs` は「中の関数が自分でロックを取り直す」前提でしたが、このアプリでは**入口がスクリプトのロックを持ったまま最後まで動かす**必要があります（`tryLock(20000)`、取れなければ `busy`）。
- 中で `api_*` を呼ばないこと。権限の判定（ログイン中の人）、状態の書き込み、`waitLock` の取り直しが入るためです。
- `rebuildLeaveSummary_` は呼ばない（集計タブは VPS が書く。03 の 1-5 の決め）。

---

## 6. 再試行で二重になるものと、防ぎ方の案

| 二重になるもの | 起きる場面 | 今の関数の守り | 防ぎ方の案 |
|---|---|---|---|
| 承認完了メール 2 通 | 送った後に台帳で失敗して送り直す、時間切れで結果が分からず送り直す | なし | メールを**別の action**（または別の段）にし、`jobId` ごとに「送った」の印を `sendEmail` の直後にスクリプトのプロパティへ残す。印があれば送らずに `already` を返す。時間切れの送り直しでも印を見る（罠 22・32） |
| サイン画像 2 つ | 保存の後に失敗して送り直す | なし（同名でも新しいファイル） | 保存したら直ちに URL を列に書き、同じ `jobId` の送り直しでは「列に URL があり、その名前が `sign_<REQ_ID>.png`」なら保存しない。結果（URL）は返事に入れ、VPS が先に DB に入れる |
| 2 次承認の PDF 2 つ | PDF を作った後・列を書く前に止まる、時間切れの送り直し、2 つのタブで同じ申請を 2 次承認（K2） | なし（残業の `already` に当たる判定が無い） | 入口で「シートの `PDF_FILE_ID` が空でなく、VPS が渡した `oldPdfFileId` と違う」なら作らずに `already`（残業 `vpsTrashOldPdf_` の `sheet_has_newer_pdf` と同じ）。VPS は 2 次承認を `expectedStatus` で照合し、同じ申請の仕事は 1 本ずつ |
| 台帳の行 2 つ | ほぼ起きない（REQ_ID で上書き）。ただし行削除が黙って失敗した後に、別の年度へ日付を変えて承認すると、古い年度タブに行が残る | upsert は同じタブの中だけ | 行削除の後に読み直して消えたことを確かめ、残っていれば `retryable` で返す |
| 台帳そのもの 2 つ | 台帳を一時的に開けなかった | なし（作ってしまう） | 入口で `SpreadsheetApp.openById(LEDGER_SSID)` を先に試し、開けなければ `ledger_unavailable`（作らない）。Settings の `LEDGER_SSID` の行数を ping で返す |
| 日付フォルダ 2 つ | サインと PDF が同時に動く | 今はロック | VPS の GAS の仕事は全体で 1 本＋入口のロック |
| 一時スプレッドシートの残り | 6 分で実行ごと止まる | `finally`（止まると動かない） | 1 回 1 件。ping で `TMP_PDF_` の残りを数えるかは任意 |

---

## 7. 03-reuse-map.md の action の案への意見

### 7-1. 決めるべきこと（大きい順）

| # | 決めること | 案 |
|---|---|---|
| D1 | **呼び方**（T1）。`access: DOMAIN` のままでは匿名で呼べない | ① VPS の委任の資格（オーナー）で取ったトークンを `Authorization: Bearer` に付けて `/exec` を呼ぶ（スコープは drive を含む見込み。推測。段階 2 の最初に試す）② Apps Script API の `scripts.run`（API 実行用のデプロイと GCP の設定が要る）③ 切り替えの版だけ匿名にする（戻したときに権限の判定が壊れるおそれがあり、おすすめしない）。①を先に試し、だめなら②。token の照合は①②でも残す |
| D2 | 編集・取消の後処理の引数 | `oldPdfFileId`（VPS が列を空にする前の値）と `oldLeaveDate`（編集前の休暇日 `yyyy-MM-dd`）を本文で渡す。GAS はシートを読まずにこれで捨てる・消す |
| D3 | メールを独立させるか | `internal.afterApprove1` を「サイン保存＋URL を書く」「メール」「台帳」に分けるか、1 つの action の中で段ごとの印を残すか。おすすめは段ごとの印（今の順序を 1 回の呼び出しで守れる） |
| D4 | 有給警告メール（メニューの手動送信）を残すか | 移転モードでメニューは消える。残すなら `internal.sendWarnMail`（GAS は VPS が書いた V_LEAVE_SUMMARY を読み、**`rebuildLeaveSummary_` は呼ばない**）。統合業務メールが同じ役目なら作らない。オーナーに聞く |
| D5 | 2 次承認の PDF が失敗したときの再試行 | 今は「やり直せない」。VPS は再試行するので、後から PDF ができる（良い方向の差）。残業の Q20 と同じく「変える一覧」に載せる |

### 7-2. action ごとの意見

| action（03 の案） | 足りない点・直す点 |
|---|---|
| `internal.ping` | 足す: `LEDGER_SSID` の行の数（2 行以上なら下の行が使われる）、`LEDGER_SSID` を開けるか、`PDF_FOLDER_ID` を開けるか、`MAIL_TO` が空か（値は返さない）、DB の時間帯（GMT 系なら空になる。罠 68）、移転モード、版。書かない |
| `internal.afterApprove1` | 本文に `signBase64`（任意）、`approverEmail`、`expect`（`承認状態=承認済`・`承認日時`・`APPROVED_BY1_EMAIL`）。順序は ① 行が `expect` どおりか（違えば `row_not_synced`）② サインがあれば保存し URL を列に書く（**電子印のときは VPS が列を空で反映済みのはず**）③ メール ④ 台帳（読み直しで確かめる）。返事に `signImageUrl`・`mail.status`（sent/already/skipped:mail_to_empty/error）・`ledger.status`。VPS は `signImageUrl` を先に DB へ入れる（入れないと行の反映・毎晩の照合が空で上書きし、2 次承認の PDF からサインが消える） |
| `internal.afterApprove2` | 本文に `approver1Email`（DB の値）・`approver2Email`・`oldPdfFileId`・`expect`（`APPROVED_AT2`・`APPROVED_BY2`・`承認状態`）。**`APPROVED_AT2` が反映済みでないと D28 が「作った日」になる**ので `expect` は必須。台帳は今どおり PDF の成否に関係なく書く。返事に `pdfFileId`・`pdfUrl`・`pdfName`・`folderId`・`trash`・`ledger` |
| `internal.afterUnapprove` | D2 の 2 つの引数。編集と取消で中身は同じ（PDF を捨てる→台帳の行を消す）。**サイン画像のファイルは捨てない**（今どおり）。行削除は読み直しで確かめる。「シートの `PDF_FILE_ID` が `oldPdfFileId` と違う新しい値」なら捨てない |
| `internal.ledgerUpsert` | 今の経路では単独で使う場面が無い（承認と 2 次承認の中で呼ぶ）。再試行・照合用として残すのはよい。`expect` を付ける |
| 足りない action | `internal.sendWarnMail`（D4 の答え次第）。管理者が PDF を作り直す操作は今のアプリに無いので作らない（`test_generatePdf` は移さない） |

### 7-3. 順番の約束

1. 同じ申請の仕事は積んだ順に 1 本ずつ。GAS の仕事は全体で同時に 1 本（日付フォルダ・台帳の取り合いを防ぐ）。
2. 呼ぶ前に、その申請の行が「積んだ時点の版」までシートに反映済みであること（`requires`）。GAS 側でも `expect` で照合する。
3. **台帳の「有給累計取得回数」は T_LEAVE_REQUEST の全行を数える**。ほかの申請の行の反映が遅れていると数がずれる。厳密に今と同じにするなら「行の反映待ちが 0」を待つ。ずれを受け入れるなら変える一覧に書く。
4. 2 次承認は 1 次承認の仕事（サイン URL が列に入る）が済んでから。サイン URL が入る前に PDF を作ると、手書きサインの申請で「（承認印取得不可）」になる。
5. 呼ぶ前に DB の状態を読み直す。取消済みなら `afterApprove1`・`afterApprove2` は動かさない（今の GAS は状態を見ないが、K1 を直さない場合は今どおり動かす。オーナーの答え Q-B1 に合わせる）。
6. M_STAMP は GAS の印の探し方が読む。切り替え後も VPS のマスタ同期が M_STAMP のタブを書き続けること（止めると新しい承認者の印が出ない）。

### 7-4. 6 分の上限と本文の大きさ

- 1 回の呼び出しで 1 件だけ。PDF 1 件は 15 秒ほど（推測）、台帳の upsert は全行を 2 回読む。VPS の待ち時間は残業と同じ 330 秒、入口の `tryLock` は 20 秒でよい見込み。
- 行が見つからないと `api_getLeaveRequestById` が約 6 秒待つ。入口で先に行を探し、無ければすぐ `row_not_synced`。
- サインの base64 は数十 KB〜数百 KB（推測）。`doPost` の本文には十分入る。ただし ① `jobId` の結果をスクリプトのプロパティに残すとき、本文（base64）を入れない（1 件 9KB の上限。残業は結果だけを 3,000 字まで）② GAS のログに本文を出さない ③ VPS の jobs の表に base64 を持つので、仕事が済んだら消す・管理者の欄に出さない ④ VPS の nginx の `client_max_body_size` は 1MB 程度（03 の案どおり）。

---

## 8. 移転モード（`Cutover.gs`）を当てる場所

### 8-1. 当てる場所の候補

| 場所 | 行 | 当て方 | 理由 |
|---|---|---|---|
| `getDb_()` | `config.gs:103-105` | 先頭で `assertDbOpenAllowed_()`（内部用の実行だけ通す） | ほぼ全部の関数が通る（罠 47） |
| `doGet` | `menu.gs:3-5` | **`getSettings_()` より前**で移転ページを返す。`page`・`reqId`・`approved` を引き継ぐ | 5 行目の `getSettings_()` が `getDb_()` で止まり、エラー画面になるため |
| `onOpen` | `menu.gs:97-113` | 移転モードならメニューを「新しいアドレスを見る」1 項目だけにする（残業の `showMigratedAddress`） | メニューの 8 項目（同期・集計・ヘッダ追加・警告メール・トリガー作成） |
| `setupAllSheets` | `menu.gs:117-118` | 先頭で `assertNotMigrated_()` | `getDb_()` を通らず `SpreadsheetApp.getActive()` を直接開き、Settings・タブを書く |
| `setupAllTriggers`・`setupSyncTrigger_`・`setupWarnMailTrigger_`・`setupNotifyTriggers_` | `menu.gs:345-423` | 先頭で止める | DB を開かずにトリガーを作る |
| `syncAllMasters`（トリガー）・`syncCalendarMaster`・`syncStampMaster`・`syncDeptApprovers` | `syncMasters.gs:13/121/171/213` | トリガーの関数は `skipBecauseMigrated_` で黙って戻る。手で呼ぶ 3 つは `assertNotMigrated_` | `getSettings_()` で止まるが、トリガーが例外で終わると毎朝エラーのメールが持ち主に届く。切り替えでトリガーも消す |
| `rebuildLeaveSummary`（メニュー） | `leaveSummary.gs:4` | `assertNotMigrated_` | VPS が書く集計タブを上書きするため（`getDb_()` で止まるが明示する） |
| `checkAndSendWarnMails`・`checkAndSendWarnMails_` | `warnMail.gs:4/11` | 包みは `assertNotMigrated_`、本体（トリガーの可能性）は `skipBecauseMigrated_` | 同上 |
| `notifyTodayLeaves_`・`notifyPendingApproval1_`・`notifyPendingApproval2_` | `notifyMail.gs:45/141/235` | `skipBecauseMigrated_` | 実物のトリガーが残っていれば毎日例外になる（T3） |
| `api_*` 30（画面から呼ぶ 27 と、呼ばれない 3） | 02 の 1 節 | 先頭で `assertNotMigrated_()` | 開いたままの古いタブから呼ばれる。`api_getApproverList`（`masters.gs:163-185`）は例外を握って `[]` を返すので、`getDb_()` だけだと「承認者が 0 人」に見える |
| `test_generatePdf`・`test_stampLookup`・`debug_checkLeaveRequestSheet` | `pdfExport.gs:7/24`、`leaveRequest.gs:328` | `assertNotMigrated_` | 誰でも呼べる（罠 48） |
| 新しい `doPost` | （無い。新しく足す） | 内部用だけを受ける。token を確かめた後でだけ `_vpsInternalCall_ = true` にして `getDb_()` を通す | 移転モードでも動かす必要がある |

### 8-2. 止める入口の全部（チェック表）

- 公開関数（末尾が `_` でないもの、全 43。02 の数と同じ）: `api_*` 30（`auth.gs` 7、`leaveRequest.gs` 10、`leaveStatusApi.gs` 1、`leaveSummary.gs` 4、`masters.gs` 8）と、`doGet`・`onOpen`・`setupAllSheets`・`setupAllTriggers`・`syncAllMasters`・`syncCalendarMaster`・`syncStampMaster`・`syncDeptApprovers`・`rebuildLeaveSummary`・`checkAndSendWarnMails`・`debug_checkLeaveRequestSheet`・`test_generatePdf`・`test_stampLookup`（確認済み・grep `^function [^_]*[^_]\(`）。
- トリガー: コードが作るのは `syncAllMasters`（毎日 6 時台）だけ。作る関数が残っている `checkAndSendWarnMails_`（8 時）・`notifyTodayLeaves_`（7 時）・`notifyPendingApproval1_`（17 時台）・`notifyPendingApproval2_`（18 時）が実物に無いかを確かめる（推測。C-9）。消すのは残業の `removeAllTriggersForMigration` と同じく「時刻のトリガーを全部」。戻すときに作り直すのは `syncAllMasters` 1 本（今のコードどおり）。
- メニュー: 8 項目（`menu.gs:98-112`）。`migrateLeaveRequestHeaders_` は末尾 `_` なのでメニューからは動かない見込み（推測）。
- デプロイ: 画面用が 2 つ（@106・@104。05 の R-1）。両方を移転モードの版に当てる。内部用の入口を D1 の①で呼ぶなら、同じ版の既存のデプロイの `/exec` を使えるか、内部用のデプロイを足すかを決める。

### 8-3. 内部用の入口と移転モードの両立

- 残業と同じく、`_vpsInternalCall_` は token を確かめた実行の中だけ `true` にし、`finally` で `false` に戻す（`Internal.gs:194-266`）。
- 移転モードの印はスクリプトのプロパティ（`MIGRATED_TO_URL`）。ping は毎回読み直して返す（VPS は本番で書く前に確かめる）。
- 運用関数（移転モードの出し入れ・token の設定）は持ち主だけ（`assertCutoverOperator_`。罠 48）。既定の URL を持たない（罠 24d）。移転ページの文言は「休暇届」に直す（残業の文言が残らないよう grep。03 の頭の注意）。
- 内部用の入口が呼ぶ関数（`saveSignImage_`・`generateLeavePdf_`・`updateLedger_`・`removeLedgerRow_`・`sendApprovalNotification_`）には `assertNotMigrated_` を**置かない**（置くと移転モードで動かなくなる）。

---

## 9. 下書きとの違い（誤り・抜け）

| 下書き | 書いてあること | 違い・抜け | 確かさ |
|---|---|---|---|
| 02 の 5 節 | 送り得るのは承認完了と有給警告の 2 種 | コードの上ではそのとおり。ただし通知 3 本と警告を作る関数が残っており、**実物のトリガーの確かめが要る**（T3）。承認完了が実際に届いているかも疑わしい（T2） | 推測 |
| 02 の 5 節 | 差出人は imaizumi@ | `USER_DEPLOYING` なので「デプロイした人」。imaizumi@ かは推測 | 推測 |
| 02 の 6 節「中身」 | 理由は特別理由／有給詳細か「私用のため」／「振替休暇」 | 抜け: ほかの種類は休暇種類そのもの、空なら `―`。台帳は「私用の為」（字が違う）。時間帯の「～」も PDF と台帳で字が違う | 確認済み |
| 02 の 6 節「中身」 | 承認日（令和の短い形） | 抜け: 2 次承認者のメールが無いと 1 次承認日が**部長の欄**に出る。2 次承認日が空なら作った日。届出日も申請日時が空なら作った日 | 確認済み |
| 02 の 6 節 | 印の扱い（所属長＝1 次、部長＝2 次） | 抜け: 印が無いときのサインの代わり、「（承認印取得不可）」の表示、M_STAMP にいるが stampFileId が空の人でも「電子印あり」になりサインを送らない | 確認済み |
| 02 の 6 節「前の PDF」 | 2 次承認・編集・取消でゴミ箱へ | 抜け: 2 次承認では**列を空にしない**ので、PDF が失敗すると捨てた PDF を指し続け、台帳もその URL になる（T7） | 確認済み |
| 02 の 6 節「サイン」 | `sign_<REQ_ID>.png` を休暇日のフォルダへ | 抜け: `PDF_FOLDER_ID` が空なら保存しない。編集で列だけ空にしてファイルは残す。取消では列も残す。もう一度承認すると同名のファイルが増える | 確認済み |
| 02 の 6 節「台帳」 | 年度タブ・20 列・行の削除 | 抜け: 開けないと新しい台帳を作り Settings に行を足す（K13 にはある）、**失敗を黙る**、累計は全行を数え後から直らない、upsert は休暇日の年度タブの中だけ | 確認済み |
| 02 の 1-1 `api_approveLeaveRequest2` | 権限は総務か管理者 | 補足: 内部用の入口からこの関数を呼ぶと、ログイン中の人が居ないので権限で止まる見込み。中身だけを再現する | 推測 |
| 03 の 1-5 の action 表 | 内部用の入口は `doPost` を新しく足して token で守る | **`access: DOMAIN` のため匿名で呼べない**（T1、D1）。残業の作りをそのまま写せない | 設定は確認済み |
| 03 `internal.afterApprove1` | サイン保存 → URL を行に書く → メール → 台帳 | 抜け: 送り直しでメール・サインが二重になる（6 節）。VPS が返事の URL を DB に入れないと、行の反映が空で上書きする。`expect` が要る | 確認済み（二重の起き方はコード） |
| 03 `internal.afterApprove2` | 前の PDF → 作る → 2 列 → 台帳 | 抜け: `APPROVED_AT2` の反映待ち（D28 の日付）、1 次のメールと 2 次のメールを引数で渡す、「もうある」の判定が今の関数に無いので入口で作る、台帳は PDF の成否に関係なく書く | 確認済み |
| 03 `internal.afterUnapprove` | PDF をゴミ箱 → 2 列を空 → 台帳の行を消す | **誤りに近い抜け**: シートから読むと、休暇日は編集後の値、PDF の列は空（VPS が先に反映）。編集前の休暇日と前の PDF の ID を引数で渡す必要がある（T4） | 確認済み |
| 03 の 1-5 Cutover の行 | 止める場所は `getDb_()` と `doGet` | 抜け: `doGet` は `getSettings_()` より前で判定、`setupAllSheets` は `getActive()` を直接開く、`onOpen` のメニュー、例外を握る `api_getApproverList`、残っているかもしれない通知のトリガー（8 節） | 確認済み |
| 03 の 1-1 `google_delegate.py` | Gmail・Drive への書き込みは VPS からしない | 同意。ただし D1 の①を選ぶと、VPS が GAS を呼ぶためのトークン（スコープ）が要る | 推測 |
| 05 の R-6 | 再試行は「PDF_FILE_ID が空のときだけ」 | 補足: 「空か」ではなく「VPS が渡した前の値と同じか」で決める（2 つのタブの K2、作った後の時間切れ）。今の関数には判定が無いので入口に足す | 確認済み |
| 05 の R-8 | 台帳を作らない、行削除は REQ_ID で探し直し | 同意。足す: 削除と upsert の後に読み直して確かめる（関数が失敗を黙るため） | 確認済み |
| 05 の 3 節 | 写しの GAS のメールは `MAIL_TO` を imaizumi だけに | 同意。足す: 写しの GAS も `access: DOMAIN` なので、ステージングの VPS からの呼び方も D1 と同じ問題がある | 確認済み（設定） |
| 05 の C-7 | PDF の所要時間を実行数の画面で見る | 足す: 同じ画面で `api_approveLeaveRequest` のログに「メール通知エラー」が出ていないか（T2）、「台帳更新エラー（続行）」「2次承認PDF再生成エラー（続行）」の数 | — |
| 05 の C-9 | トリガーは `syncAllMasters` 1 本の見込み | 足す: 通知 3 本・警告 1 本が無いことを名前で確かめる（T3） | 推測 |

---

## 10. 初日に足す確かめ（読むだけ）

| # | 確かめること | やり方 |
|---|---|---|
| X-1 | 承認完了メールが届いているか（T2） | GAS の実行数で `api_approveLeaveRequest` のログ（「メール通知エラー」）、オーナーの送信済みに「【休暇届 承認完了】」があるか |
| X-2 | 実物のトリガー（T3） | トリガーの画面で関数名の一覧 |
| X-3 | 台帳の実物 | `LEDGER_SSID` の行の数（Settings）、台帳のタブ一覧（`_tmp` や同じ年度の重なりが無いか）、見出し 20 列が 4-3 と同じか |
| X-4 | PDF フォルダの実物 | 同じ日付の名前のフォルダが 2 つないか、`TMP_PDF_` の残り、サインの PNG の大きさ（数件） |
| X-5 | D1 の呼び方 | ステージングの写しで、委任のトークン付きで `/exec` に POST できるか（段階 2 の最初） |
