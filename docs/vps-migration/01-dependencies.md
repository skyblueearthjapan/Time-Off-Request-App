# 外の仕組みとのつながり（読む・書く）と移行への影響

作成: 2026-11-26。スキルの罠 65・66（「ほかの仕組みが DB を読み書きしていた」）を最初に潰すための一覧。
確かめ方: このアプリのコード、Apps Script API で読んだ統合業務メールのコード（読むだけ）、Drive のメタデータ、`D:\projects` の全リポジトリの `git grep`（DB・台帳・PDF フォルダの ID とデプロイ ID、`T_LEAVE_REQUEST`・`V_LEAVE_SUMMARY`）。

## 0. まとめ

| # | 相手 | 向き | 読む／書くもの | 確かさ | 移行への影響（提案） |
|---|---|---|---|---|---|
| 1 | 残業 DB `1Knx_kaQ…`（`社内カレンダーマスタ`・`StampMap`・`DeptApprovers`） | このアプリが**読む** | 3 タブを丸ごと写す | 確認済み（コード） | VPS の worker が同じ 3 タブを読む（決定 D 相当）。残業は VPS へ移った後もこの 3 タブを今の形で保つ（残業の決定 #21・4-2 B 区分） |
| 2 | マスタ元 `1iu5Hoak…`（`作業員マスタ`） | このアプリが**読む** | 1 タブを丸ごと写す | 確認済み（コード） | 同上 |
| 3 | **統合業務メール**（GAS `1lb7bmF9…`） | このアプリの DB を**読む** | `T_LEAVE_REQUEST`・`V_LEAVE_SUMMARY`・`M_SYSTEM_SETTING` | **確認済み（そのコード）** | 決定 A 相当（行単位の反映）が必須。日付は日付型のセル。V_LEAVE_SUMMARY を書き直し続ける。`APP_URL` の扱いを決める |
| 4 | Sebastian の `inout-request`（VPS `/root/.hermes/skills/internal/inout-request`） | 不明 | 不明 | 【要確認】 | このシートに行を足すなら `sheet_intake`（残業の決定 #26）が要る |
| 5 | 社内ポータル（「社内ダッシュボード」APP_MASTER） | このアプリへ**リンク** | url 列 | 推測（設計指示書） | 切り替えで url を差し替える。行の key は【要確認】 |
| 6 | このアプリの台帳 `1B7tnsT4…`・PDF フォルダ `1MQibR08…` | このアプリが**書く**。総務・統合業務メールのリンク先 | 年度タブ・PDF | 確認済み | 今の GAS が書き続ける（決定 B 相当）。ID は変えない |
| 7 | スタッフ向けダッシュボード | このアプリの中の `?page=leaveStatus` | — | 確認済み | 外の仕組みではない。画面の 1 つとして移す |
| 8 | ほかのリポジトリ（`D:\projects\*`） | — | — | 確認済み（`git grep` で 0 件） | 残業の `src/mail.gs` に「休暇届と年度をそろえた」というコメントがあるだけ |

## 1. このアプリが読む外のシート（確認済み・`syncMasters.gs`）

毎日 6 時台の `syncAllMasters`（トリガー）と、メニューの手動実行で動く。どれも「元のタブを丸ごと読み、DB の同じ名前のタブを全部消して書く」。

| 元 | 元のタブ | 写す先（このアプリの DB） | コード | 使う列 |
|---|---|---|---|---|
| `MASTER_SOURCE_SSID`（既定 `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ`） | `MASTER_SOURCE_WORKER_SHEET`（既定 `作業員マスタ`） | `作業員マスタ` | `syncMasters.gs:13-115` | 作業員コード・氏名・部署（＋「在籍フラグ」「Googleアカウント」があれば。実物には無い見込み＝スキル 00 の注意） |
| `CALENDAR_SOURCE_SSID`（既定 `1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ`） | `CALENDAR_SOURCE_SHEET`（既定 `社内カレンダーマスタ`） | `M_CALENDAR` | `syncMasters.gs:121-166` | 日付（日付型）・区分（`休日`・`祝日`・`出勤土曜`） |
| 同じ残業 DB | `StampMap`（固定名） | `M_STAMP` | `syncMasters.gs:171-208` | A メール・B stampFileId・C 名前（備考） |
| 同じ残業 DB | `DeptApprovers`（固定名） | `M_DEPT_APPROVERS` | `syncMasters.gs:213-249` | A 部署・B 承認者メール（カンマ区切り）・C 名前 |

移行への影響:
- 残業は 2026-11-26 に VPS へ移った。`社内カレンダーマスタ` は残業の VPS が今と同じ規則で書き（残業の決定 #21）、`StampMap`・`DeptApprovers` は人が直すタブとして VPS は読むだけ（残業 4-2 の B 区分）。**だからこのアプリの読み先は変わらない**（確認済み・残業の計画書）。
- 残業が戻し（ロールバック）をしても、タブは同じ場所に同じ形で残るので影響は無い。
- 気をつける点: 残業の VPS は「見出しの文字で列を読む」約束を貯蔵品のために守っている。このアプリは**位置**（A・B・C 列）で読む（`auth.gs:84-86`、`masters.gs:145-146`）。列の順番が変わると、このアプリの承認者判定と電子印が黙って壊れる。残業側で列を足すときは末尾に足す決まりを、両方の docs に書く。
- 同期の失敗時に `deleteSheet`→`insertSheet` でタブを作り直す（`syncMasters.gs:72-74` ほか）。VPS では作り直しを写さない（残業と同じ）。

## 2. 統合業務メール（確認済み。2026-11-26 に Apps Script API でコードを読んだ）

| 項目 | 値 |
|---|---|
| GAS | `1lb7bmF9loeftApfNA4_pTOL2x9RfdSZhpABMgf8ctHVEJ9ma1ucBPBvX` 題名「統合業務メール」（スタンドアロン、作成 2026-03-11、最終更新 2026-04-10） |
| デプロイ | @HEAD だけ（Web アプリとしては使っていない見込み）。動くのは時間主導トリガー |
| トリガー（コード上） | `sendMorningMail` 7:00、`sendEveningMail1` 17:00、`sendEveningMail2` 18:00（`trigger.gs` の `setupAllTriggers`）。実際の登録は【要確認】 |
| スコープ | spreadsheets・gmail.send・script.scriptapp |
| 読む ID（`config.gs` の定数） | `LEAVE_SS_ID = '16v1tQNCFkDTk2mnAZ4E2eEdHeUQ2Bepdcsn5vBAmllw'`（このアプリ）、`OT_SS_ID = '1Knx_kaQ…'`（残業）。本文のリンク用に PDF フォルダ `1MQibR08…` と台帳 `1B7tnsT4…` の URL |
| 宛先 | 残業 DB の Settings `HR_MAIL_TO`（夕方は `HR_MAIL_TO_EVENING_EXTRA` を足す） |
| 手元のコピー | この PC には無い（`C:\Users\imaizumi.LINEWORKS-NET` と `D:\` を名前で探して 0 件）。w2403c の `Documents\Integrated-Operations-Email` にあるとされるが未確認。**GAS 上のコードが正**なので、手元のコピーは要らない |

### 2-1. このアプリのシートから読むもの（確認済み・`readerLeave.gs`・`config.gs`）

| 関数 | タブ | 読む列（見出しの文字で探す。英語の見出しは日本語に読み替え） | 条件 | 使う場面 |
|---|---|---|---|---|
| `readTodayLeaves_` | `T_LEAVE_REQUEST` | 承認状態・休暇日・申請日時・部署名・作業員名・休暇種類・休暇区分・半日区分 | 承認状態＝`承認済`、**休暇日が Date 型**で今日（JST）、または今日提出の過去申請 | 朝 7:00「本日の休暇者」 |
| `readPendingApproval1_` | 同上 | 同上 | 承認状態＝`申請中`、休暇日が今日以降（Date 型）、または今日提出の過去申請 | 夕方 17:00「1 次未承認」 |
| `readPendingApproval2_` | 同上（＋承認日時・2次承認日時 `APPROVED_AT2`） | 同上 | 承認状態＝`承認済` で `APPROVED_AT2` が空 | 夕方 18:00「2 次承認待ち」 |
| `readWarnWorkers_` | `V_LEAVE_SUMMARY` | 年度・警告レベル・部署名・作業員名・有給回数・最終有給日 | 今年度で警告レベルが空でない | 朝 7:00（7/1・9/1・1/1・2/16 だけ）「有給未達者」 |
| `getLeaveSettings_` | `M_SYSTEM_SETTING`（2 行目から、A キー・B 値） | `SOMU_EMAILS`（関数はあるが使っていない見込み）、`APP_URL` | — | 本文のリンク「休暇届アプリ：<APP_URL>」「2次承認ページ：<APP_URL>?page=somuAdmin」 |

### 2-2. 移行への影響（提案）

1. **行単位の反映が要る**（決定 A 相当）。毎晩の書き出しだと、昼に出た申請が 17:00・18:00 の通知に出ない。書き込みのたびに数秒〜1 分で T_LEAVE_REQUEST の該当行を反映する。
2. **日付の列は日付型のセルで書く。** 統合業務メールは `leaveDate instanceof Date` でない行を黙って飛ばす（`readerLeave.gs` の `if (!(leaveDate instanceof Date)) continue;`）。文字や数値で書くと通知から消える。休暇日・振替元出勤日は「JST の 0 時」の時刻、申請日時・承認日時・2次承認日時は「その時刻」を、**シートの時間帯で**シリアル値にして書く（残業の `sheet_codec` と同じ。シートの時間帯は【要確認】。罠 68）。
3. **承認状態の文字はそのまま**（`申請中`・`承認済`・`取消`）。
4. **`APPROVED_AT2` の列を空のまま／日時で書く。** 2 次承認待ちは「空かどうか」で決まる。空を `''` 以外（`null` の文字など）で書かない。
5. **V_LEAVE_SUMMARY を書き直し続ける。** 今の GAS は承認・編集・削除・管理画面を開くたびに全部書き直している。VPS も DB の集計からこのタブを同じ見出し・同じ規則で書き直す（全行の置き換えは今の GAS と同じ動きなので、決定 A の「タブを消さない」の例外にしてよい。タブそのものは消さない）。朝 7:00 の前に最新になっていることを確かめる（警告の日だけでよいが、毎朝 6:50 に書き直すのが簡単）。
6. **`M_SYSTEM_SETTING` の `APP_URL`**: 統合業務メールの本文のリンクになる。残業の前例は「GAS の `/exec` のまま変えない。移転ページが `page` を引き継いで新しいアプリへ案内する」（残業 4-11）。このアプリでも同じにする案。ただし**このアプリの画面は `APP_URL` を画面の移動に使っている**（`menu.gs:6`、全テンプレート）ので、VPS の画面は Settings の `APP_URL` ではなく自分の公開 URL（`PUBLIC_ORIGIN`）を使う。戻せる期間の後に `APP_URL` を新しい URL に書き換えるかは 04 の Q-B7 で聞く。
7. 統合業務メールは**変えない**（オーナーの GAS だが、このアプリの移行の範囲の外）。

## 3. Sebastian の inout-request（【要確認】）

- 置き場: 同じ VPS の `/root/.hermes/skills/internal/inout-request`（こちらからは読めない。root の範囲で、移行の作業では触らない決まり）。
- 残業の移行のディレクターは、`inout-request/_common.py` に `append_request_row` があるのを見ている（タスクの説明より。残業の docs には書かれていない）。名前から「入出（外出・休暇など）の申請をシートに足す」仕組みの見込み（推測）。
- **このアプリのシート（`16v1tQNC…`）に書いているかは分からない。** 書いているなら、残業と同じく Sheets API の `values.append` で T_LEAVE_REQUEST に直接行を足している見込み（このアプリには doPost が無いので、GAS 経由の入口は無い。確認済み）。

確かめ方（初日。`05` の C-6）:
1. オーナーの了承を得て、VPS で読むだけ: `sudo grep -rl -e 16v1tQNC -e T_LEAVE_REQUEST -e 休暇 /root/.hermes/skills/internal/inout-request`（root は `ssh lineworks-vps`。`hostname` を先に確かめる）。
2. シート側の手がかり: T_LEAVE_REQUEST の `CREATED_BY_EMAIL` 列が空の行、`REQ_ID` が `LV-YYYYMMDD-NNN` の形でない行、`申請日時` が文字の行を数える（gviz の `select U, count(A) group by U` など）。
3. 書いていたら: VPS に `sheet_intake`（残業の決定 #26・#27）を足す。60 秒ごとに T_LEAVE_REQUEST の REQ_ID の列を読み、DB に無い行を移行の道具と同じ決まりで取り込む。時刻の 9 時間ずれ（罠 67）も同じく確かめる。書いていなければ不要。

## 4. 社内ポータル（推測）

- `docs/社内ポータル導線_共通設計指示書.md` の対象アプリの 3 番目が「休暇届け申請」。ポータルのスプレッドシート「社内ダッシュボード」`1DdhbJuz-zyS_2pKcM8y6Xf2pHofuzwNDgoW9eWcB0yg` の APP_MASTER に行がある見込み。key・行番号・今の url（@106 か @104 か）は【要確認】（`05` の C-1）。
- このアプリの「社内ポータルへ移動」ボタンは Settings の `PORTAL_URL` を使う（`menu.gs:7`）。値が空ならボタンを出さない。VPS も同じ値を使う。

## 5. このアプリが書く外のもの（確認済み）

| もの | ID | 書く処理 | 移行への影響（提案） |
|---|---|---|---|
| PDF フォルダ「休暇届け申請書／管理フォルダ」 | `1MQibR08wWsE_OGurO9RocEEDGgDOwo-w`（その下に `yyyy.MM.dd` のフォルダ 129） | PDF・サイン画像の保存、古い PDF をゴミ箱へ | 今の GAS が内部用の入口で書き続ける（決定 B 相当）。VPS の安全装置で、本番はこのフォルダの ID だけを許す |
| 台帳「休暇届台帳（総務確認用）」 | `1B7tnsT4lr80o3pFMBxB2zEhsx2UW2PCL48YTQc9i1eU`（PDF フォルダの中、年度タブ） | upsert・行の削除 | 同上。**台帳が開けないとき新しく作る動き（`ledger.gs:20-44`）は内部用の入口では止める**（残業の Internal.gs と同じ） |
| 一時スプレッドシート `TMP_PDF_…` | マイドライブの直下に作って、すぐゴミ箱へ | PDF 作成 | 今の GAS のまま |
| Gmail（imaizumi@ の送信済み） | — | 承認完了通知（`MAIL_TO` が空でなければ） | 今の GAS のまま内部用の入口から送る案（04 の Q-B4） |

## 6. 残っている確かめ

- 統合業務メールのトリガーが実際に登録されているか（GAS エディタの「トリガー」画面で 3 本。読むだけ）。
- Sebastian の inout-request のコード（オーナーの了承の上で読むだけ）。
- ポータルの APP_MASTER の行。
- w2403c にある統合業務メールの手元のコピーが GAS と同じか（違っても GAS 上のコードが正。参考まで）。
