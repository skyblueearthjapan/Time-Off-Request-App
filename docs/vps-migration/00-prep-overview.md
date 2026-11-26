# 休暇届アプリ VPS 移行 — 下調べのまとめ（ここから読む）

作成: 2026-11-26（段階 0 の下調べ。コード・本番には読むだけで触れた）。
対象: 「休暇届けapp」（GitHub `skyblueearthjapan/Time-Off-Request-App`、手元 `D:\projects\kyuka-todoke-app\repo`）。
姉妹アプリ「残業・休日出勤申請」（`D:\projects\zangyo-kyujitsu-app\github-repo`、2026-11-26 09:30〜09:46 に切り替え済み）を一番近い前例として写す。

> 書き方: 「確認済み」はコード・Google の API（読むだけ）で確かめた事実。「推測」は未確認。値が無いものは 〇〇、確かでない値は【要確認】。
> 行番号は `src/` のファイル（2026-11-26 時点の master `216db54`）。

## この資料群の読み方

| 資料 | 中身 | 誰が使うか |
|---|---|---|
| `00-prep-overview.md`（これ） | アプリの全体像、GAS の事実、確定値表の案、日ごとの進め方 | 今泉オーナー・次のディレクター |
| `01-dependencies.md` | このアプリのシートを読む・書く外の仕組みと、移行への影響 | ディレクター（計画書 4 章・8 章の元） |
| `02-inventory-quick.md` | 関数・画面・シート・設定・メール・PDF・権限・時刻・癖の一次棚卸し（行番号つき） | 棚卸しの 4 人・仕様を書く人 |
| `03-reuse-map.md` | 残業の `vps-app/` から写すもの・直すもの・新しく作るものと、担当ごとの工数 | ディレクター（担当の割り振り） |
| `04-owner-questions.md` | オーナーへの質問（おすすめを先頭）。残業の決定を「同じでよいか」にまとめた | ディレクター → オーナー |
| `05-risks-and-checks.md` | このアプリ固有の危険と、初日に行う確かめ | ディレクター |

スキル `gas-to-vps-migration`（`C:\Users\imaizumi.LINEWORKS-NET\.claude\skills\gas-to-vps-migration\`）の段階 0〜1 に当たる。計画書（`docs/vps-migration-plan.md`）と抜け調査はまだ作っていない。

## 1. アプリがしていること

### 1-1. 業務の流れ（確認済み・コード）

紙の休暇届をそのまま画面にしたアプリ。承認者を特定しない作り（要件定義「サイン＝承認」）。

```
TOP（部署・作業員を選ぶ → 申請モーダル）
  │ api_submitLeaveRequest → T_LEAVE_REQUEST に 1 行（承認状態=申請中）
  ▼
プレビュー（?page=preview&reqId=LV-…）  紙の様式を画面いっぱいに出す
  │ 「承認」ボタン（その場の上長が押す）
  ▼
サイン（?page=sign&reqId=…）  電子印の持ち主を選ぶ or 手書きサイン（canvas）
  │ api_approveLeaveRequest → 承認状態=承認済・承認日時・APPROVED_BY1_EMAIL・サイン画像
  │   裏で: V_LEAVE_SUMMARY 作り直し → 承認完了メール（MAIL_TO）→ 台帳に 1 行
  ▼
総務管理（?page=somuAdmin）の「2次承認待ち」
  │ api_approveLeaveRequest2（2次承認者＝ログイン中のアカウント）
  │   裏で: 前の PDF をゴミ箱 → PDF を作る（所属長印＋部長印）→ PDF_URL・PDF_FILE_ID → 台帳を更新
  ▼
完了（PDF は Drive「休暇届け申請書／管理フォルダ」の休暇日のフォルダ）
```

- **PDF は 2 次承認のときだけ作る**（1 次承認では作らない。`leaveRequest.gs:527`、コミット `2a37cbe`）。
- 編集（TOP の一覧から）: 休暇日・全日/半日・午前/午後・振替元を変えられる。承認済みなら申請中へ戻し、承認の列を空にし、PDF をゴミ箱へ、台帳の行を消す（`leaveRequest.gs:852-966`）。
- 削除（取消）: 承認状態=取消、PDF をゴミ箱へ、台帳の行を消す（`leaveRequest.gs:970-1006`）。
- 過去申請（遡及）: 今年度（3/16〜今日）の勤務日だけ。画面で確認ダイアログ（`config.gs:193`、`top.html:1141`）。
- 振替休暇: 振替元出勤日が必須。同じ月次期間（前月 16 日〜当月 15 日）の中、同じ人で使い回し不可（`leaveRequest.gs:50-124`）。
- 年度は 3/16 始まり（`leaveRequest.gs:34`。要件定義の 4/1 から変わっている）。

### 1-2. 画面（確認済み・`menu.gs:3-89`）

| `?page=` | テンプレート | 入口の判定 | 題名 | 中身 |
|---|---|---|---|---|
| `top`（既定・不明な値も） | `top.html`（1,364 行） | なし | 休暇届 申請TOP | 部署・作業員の選択、申請モーダル、休暇予定の一覧（編集・削除・承認へ）、全員の承認待ち一覧、過去申請カレンダー |
| `preview` | `preview.html` | なし | 休暇届 プレビュー | 紙の様式。承認済みなら赤丸「承認済」と部長印欄 |
| `sign` | `sign.html` | なし | 休暇届 サイン | 承認者（M_STAMP の一覧）か手書きサインを選んで承認 |
| `somuAdmin` | `somuAdmin.html` | `SOMU_EMAILS` か `ADMIN_EMAILS` | 休暇届 総務管理 | 2 次承認（1 件・選んだ分を順番に）、全部署の有給取得状況、CSV |
| `deptAdmin` | `deptAdmin.html` | `ADMIN_EMAILS` か M_DEPT_APPROVERS の承認者 | 休暇届 部署管理 | 担当部署の有給取得状況（管理者は承認者を選んで見られる） |
| `admin` | `admin.html` | `ALLOW_ADMIN_PAGE_EMAILS` | 休暇届 管理 | 古い管理画面（上長ビュー・総務ビュー）。TOP からのリンクは無いが、メール本文の `?page=admin` から開く |
| `leaveStatus` | `leaveStatus.html` | なし（誰でも・どの人でも検索できる） | 休暇届 取得状況 | スタッフ向けダッシュボード（2026-07-07 追加、`d39585e`・`216db54`） |
| — | `no_auth.html` | — | アクセス権限エラー／アクセス権限なし | 権限エラー |

URL の引数は `page`・`reqId`・`approved`（`approved=1` はサインの後に付くが、プレビューは使っていない）。全画面 `XFrameOptionsMode.ALLOWALL`、`<base target="_top">`。

### 1-3. 役割（確認済み・`auth.gs`）

| 役割 | 決め方 | できること |
|---|---|---|
| だれでも（ドメイン内） | Google ログインだけ | 申請・編集・削除・プレビュー・**1 次承認（サイン）**・取得状況 |
| 部署承認者 | M_DEPT_APPROVERS の B 列にメール（`auth.gs:76`） | 部署管理 |
| 総務 | M_SYSTEM_SETTING `SOMU_EMAILS` | 総務管理・2 次承認 |
| 管理者 | `ADMIN_EMAILS` | 部署管理（全承認者）・総務管理 |
| 古い管理画面 | `ALLOW_ADMIN_PAGE_EMAILS` | `?page=admin` |

一覧が空なら誰も当てはまらない（空で全員にはならない。罠 35 には当たらない）。Google アカウントは部署ごとの共有（残業の決定 #4）と同じ見込み（推測。04 で確かめる）。

### 1-4. メール（確認済み・コード、統合業務メールは API で読んだ）

| 種類 | 送り手 | 宛先 | いつ | 今の状態 |
|---|---|---|---|---|
| 承認完了通知「【休暇届 承認完了】氏名 種類」 | このアプリ `sendApprovalNotification_`（`mail.gs:6`） | M_SYSTEM_SETTING `MAIL_TO` | 1 次承認のたび | `MAIL_TO` が空なら送らない。値は【要確認】 |
| 有給取得警告 | このアプリ `checkAndSendWarnMails`（`warnMail.gs:4`） | `MAIL_TO` | メニューから手で | トリガーは作られない（`menu.gs:348`）。手で押せば送る |
| 1 次未承認・2 次未承認・本日の休暇者 | このアプリ `notifyMail.gs` | 部署承認者＋`SOMU_EMAILS` など | — | トリガーを作る処理を外した（2026-03-11、統合メールへ移した）。関数は残っている |
| 朝の報告（本日の休暇者、7/1・9/1・1/1・2/16 は有給未達者）、夕方①（1 次未承認）、夕方②（2 次承認待ち） | **統合業務メール**（別の GAS、`01-dependencies.md` 2 節） | 残業 DB の Settings `HR_MAIL_TO`（夕方は `HR_MAIL_TO_EVENING_EXTRA` も） | 7:00・17:00・18:00 | 動いている見込み（このアプリのシートを読む。確認済み・コード） |

差出人は GAS の持ち主 imaizumi@（`executeAs: USER_DEPLOYING`）。

### 1-5. PDF・サイン・台帳・集計（確認済み・コード）

- **PDF**: ひな形を使わない。`SpreadsheetApp.create` で一時スプレッドシートを作り、セルの結合・色・罫線で様式を組み、電子印をサムネイル（Drive の thumbnail URL）で `insertImage`、4 秒待って docs の export URL で A4 縦の PDF にし、一時シートを捨てる（`pdfExport.gs:124-180`、`732`）。置き場は `PDF_FOLDER_ID` の下の休暇日のフォルダ `yyyy.MM.dd`、名前 `yyyyMMdd_部署_氏名_休暇届.pdf`。1 件 15 秒ほど（画面の確認文「1件あたり約15秒」`somuAdmin.html`）。
- **印**: 所属長欄（G24）＝1 次承認者の電子印（M_STAMP の stampFileId）、無ければ手書きサインの PNG、どちらも無ければ「（承認印取得不可）」。部長欄（D24）＝2 次承認者の電子印。
- **サイン**: canvas の PNG（base64）をサーバーで Drive の同じ日付フォルダに `sign_<REQ_ID>.png` として保存し、URL を T_LEAVE_REQUEST に入れる（`pdfExport.gs:89`）。実物 22 件、1 件 10〜14KB（確認済み・Drive）。
- **台帳**: 別のスプレッドシート「休暇届台帳（総務確認用）」`1B7tnsT4lr80o3pFMBxB2zEhsx2UW2PCL48YTQc9i1eU`（PDF フォルダの中）。年度ごとのタブ「2026年度」、20 列。1 次承認・2 次承認で upsert、編集で承認が外れたとき・取消で**行を消す**（`ledger.gs`）。`LEDGER_SSID` が開けないと新しい台帳を作って設定に行を足す（`ledger.gs:7-45`）。
- **集計**: `V_LEAVE_SUMMARY` タブを、承認・編集・削除のたびと、管理画面・取得状況を開くたびに**全部消して書き直す**（`leaveSummary.gs:15-116`）。統合業務メールがこのタブの「警告レベル」を読む。

### 1-6. 実データの量（確認済み・Drive のメタデータ）

- PDF 274 件（2026-04-07〜）。月 26〜44 件、平日 1 日あたり 1〜2 件。
- 日付フォルダ 129、サイン画像 22。一時 PDF シート（`TMP_PDF_`）の残りは 0。
- T_LEAVE_REQUEST の行数・状態別の件数は【要確認】（Sheets を読む権限が clasp のトークンに無い。初日に gviz か委任で数える。`05` の C-2）。

## 2. GAS の事実（確認済み。2026-11-26 に Apps Script API・Drive API で読んだ）

| 項目 | 値 |
|---|---|
| scriptId | `15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i`（題名「休暇届けapp」、作成 2026-03-04、持ち主・最終更新 imaizumi@lineworks-local.info） |
| 紐づくスプレッドシート（parentId） | **`16v1tQNCFkDTk2mnAZ4E2eEdHeUQ2Bepdcsn5vBAmllw`「休暇届けapp」**（imaizumi@ 所有、最終更新 2026-11-25T21:03Z＝11/26 06:03 JST。毎朝 6 時のマスタ同期と合う） |
| 同じ名前の古いシート | 無い。名前に「休暇」を含むスプレッドシートは上の DB と台帳の 2 冊だけ（ゴミ箱にも無い） |
| HEAD | 2026-07-07T12:09:06Z 更新。GitHub master `216db54`（ローカル・origin とも同じ）と全 22 ファイル一致（改行の違いだけ） |
| 版 | 全 106 版。**@106 の中身＝HEAD**。@105 は `leaveStatus.html` だけ違う。@104 は `menu.gs`・`top.html` が違い、`leaveStatus` が無い |
| Web アプリのデプロイ | ① `AKfycbx7sB9x7iLigEqZX9LuErK95R2ZC0cMOyU9efr6pIbpnqAnnDC3o5_DGFbtwgSh64j2` **@106**（説明なし、2026-07-07 21:09 JST）② `AKfycbwHI1uz6tcer59sRC4dYEACDaevSn9a_eJa3J0E8LZgjdWCGPJVaqr_Qa9L-HXVKI8k` **@104**「振替元プルダウン月次絞込」（2026-04-11）③ @HEAD `AKfycbw_GAlPaXbETZTNpkK84k0YRwW5koclx1AHFpaqDL8`（開発用）。3 つとも `access: DOMAIN`・`executeAs: USER_DEPLOYING` |
| 本番の版 | **@106 の見込み**（取得状況ページの導線が @105 以降にしか無いため。推測）。ただし @104 も生きている。M_SYSTEM_SETTING の `APP_URL` と社内ポータルの url がどちらを指すかは【要確認】（`05` の C-1） |
| appsscript.json | `timeZone: Asia/Tokyo`、V8、`webapp: {executeAs: USER_DEPLOYING, access: DOMAIN}`（手元にも GAS にもある。罠 45 は当たらない）。スコープ: spreadsheets・drive・gmail.send・script.external_request・script.scriptapp |
| doPost | 無い（外から呼ばれる入口は doGet だけ） |
| トリガー（コード上） | `syncAllMasters` 毎日 6 時台（`menu.gs:381-390`）だけを `setupAllTriggers` が作る。有給警告（8 時）と通知 3 種（7:00・17:15・18:00）を作る関数は残っているが呼ばれない（`menu.gs:345-351`）。**実際に登録されているトリガーの一覧は【要確認】**（GAS エディタの「トリガー」画面。古いトリガーが消されずに残っている前例がある） |
| シンプルトリガー・メニュー | `onOpen` でメニュー「休暇届アプリ」8 項目（`menu.gs:97-113`） |
| GitHub Actions | 無い（`.github/` が無い。罠 62 は当たらない） |
| 公開関数（`google.script.run` で呼べる） | 43（画面から呼ぶもの 27、呼ばれないもの 16。`02-inventory-quick.md` 1 節） |
| ScriptProperties・CacheService | 使っていない。設定はすべて M_SYSTEM_SETTING タブ |
| LockService | `getScriptLock`（申請 15 秒、承認 60 秒、編集・削除 20 秒、同期 30 秒） |

## 3. 確定値表（案）

スキル `references/00-fixed-values.md` の「アプリごとに決める値」をこのアプリで埋めた。使用済みの名前（`supplies`・`parts`・`overtime`、サブネット 10.91〜10.93）と重ならない。

| 値 | このアプリ（案） | 状態 |
|---|---|---|
| アプリ名（フォルダ・コンテナ・DB 名） | `kyuka` | 案（リポジトリのフォルダ名に合わせた） |
| 公開 URL | `https://leave.lineworks-local.net` | 案。ほかの候補: `kyuka.`（日本語読み。社内で通じやすい）、`timeoff.`（GitHub の名前に合う）、`leave-request.`（長い）。オーナーが決める（04 の Q-A1） |
| Tunnel 名 | `leave-vps`（URL を `kyuka` にするなら `kyuka-vps`） | 案 |
| VPS | `srv1508169`（`ssh lineworks-vps-user`）。毎回 hostname を確かめる | 共通 |
| VPS の置き場 | `/opt/kyuka` | 案（未使用か初日に `ls /opt` で確かめる） |
| docker のサブネット | `10.94.0.0/24`（web の固定 IP `10.94.0.10`） | 案（初日に `docker network inspect` で空きを確かめる） |
| コンテナ・ボリューム | `kyuka-db`・`kyuka-api`・`kyuka-worker`・`kyuka-web`・`kyuka-cloudflared`・`kyuka-backup`、`kyuka-pgdata` | 案 |
| OAuth クライアント名 | `Time Off Request (VPS)`（GCP `sebastian01`、同意画面は共通名のまま） | 案 |
| リダイレクト URI | `https://<host>/auth/callback`、`https://<host>/auth/google-delegate/callback` | 共通の形 |
| 委任の持ち主 | `imaizumi@lineworks-local.info`（GAS の実行ユーザー） | 確認済み（deployments の executeAs と HEAD の更新者） |
| GAS scriptId | `15aGUL5ilIYawxnib_8-81vYvphg7DcekgDC-XRVRxr_vazRIiyIU-P4i` | 確認済み |
| GAS 本番 deploymentId・版 | `AKfycbx7sB9x…` @106（と `AKfycbwHI1uz…` @104） | 推測。【要確認】どちらが本番か・両方使われているか |
| 本番スプレッドシート（DB、コンテナ） | `16v1tQNCFkDTk2mnAZ4E2eEdHeUQ2Bepdcsn5vBAmllw`「休暇届けapp」 | 確認済み（parentId） |
| 使ってはいけない古いシート | なし（Drive 検索） | 確認済み。ただし docs の「DB マスターガイド」は別の雛形（xlsx、タブ名 CONFIG・M_WORKER など）を書いている。移行の道具は本番 ID 以外を拒む |
| 本番 PDF フォルダ（`PDF_FOLDER_ID`） | `1MQibR08wWsE_OGurO9RocEEDGgDOwo-w`「休暇届け申請書／管理フォルダ」 | 確認済み（PDF の親の親・統合業務メールの定数）。Settings の値と同じかは初日に照合 |
| 台帳（`LEDGER_SSID`） | `1B7tnsT4lr80o3pFMBxB2zEhsx2UW2PCL48YTQc9i1eU`「休暇届台帳（総務確認用）」 | 確認済み（Drive、統合業務メールの定数）。Settings の値は初日に照合 |
| 読むだけの外部シート | マスタ元 `1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ`（`作業員マスタ`）、残業 DB `1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ`（`社内カレンダーマスタ`・`StampMap`・`DeptApprovers`） | 確認済み（コードの既定値。Settings の値は初日に照合） |
| このシートを読む外の仕組み | 統合業務メール（GAS `1lb7bmF9…`）。Sebastian の inout-request は【要確認】 | `01-dependencies.md` |
| DB のファイルの時間帯・地域 | 〇〇（残業 DB は `Etc/GMT`・en_US だった） | 【要確認】初日に Sheets API で読む |
| 社内ポータルの行 | APP_MASTER の key 〇〇（label に「休暇」を含む行の見込み） | 【要確認】 |
| 手元の秘密の置き場 | `C:\Users\imaizumi.LINEWORKS-NET\.secrets\kyuka\`（`google-oauth-client.json`・`tunnel-token.env`・`vps.env`・`gas-internal-token.*`） | 案 |
| ステージング | DB の「コピーを作成」（紐づく GAS の写しができる）・台帳のコピー・試験用 PDF フォルダ | これから作る（`05` の C-8 の書き換え表を使う） |

## 4. 移行の形（提案。残業の決定 A〜D をそのまま当てはめた）

```
利用者 ──https──> Cloudflare Tunnel leave-vps ──> kyuka-web(nginx) ──> kyuka-api(FastAPI)
                                                                  │ DB の取引の中で outbox に積む
                                                   kyuka-db(PostgreSQL) <── kyuka-worker
                                                                  │①行の反映（T_LEAVE_REQUEST の該当行だけ）
                                                                  │②集計タブ V_LEAVE_SUMMARY の書き直し
                                                                  │③GAS の内部用の入口を呼ぶ（PDF・サイン画像・台帳・承認メール）
                                                                  │④マスタ同期・設定の取り込み
                                                                  ▼
                       今のスプレッドシート 16v1tQNC…（統合業務メール・戻した GAS が今と同じに読む）
                       今の GAS 15aGUL5i…（内部用の入口だけが動く。画面は移転ページ）
```

- **決定 A 相当（行単位の反映）**: 統合業務メールが T_LEAVE_REQUEST・V_LEAVE_SUMMARY・M_SYSTEM_SETTING を毎日読むので、毎晩のタブ置き換えではなく書き込みのたびに行を反映する。日付は**日付型のセル**で書く（統合業務メールは `instanceof Date` で読む。`01` の 2 節）。
- **決定 B 相当（GAS に残す処理）**: PDF（一時シート＋サムネイル＋export）、サイン画像の保存、台帳、承認完了メールは、今の GAS の関数をそのまま内部用の入口から呼ぶ。見た目と中身が今と同じになり、VPS は Drive・Gmail のスコープを持たずに済む。
- **決定 C 相当（時刻）**: サーバーの「今日」・年度（3/16）・月次期間（16 日）・祝日の計算は Asia/Tokyo。
- **決定 D 相当（マスタ）**: `作業員マスタ`・`M_CALENDAR`・`M_STAMP`・`M_DEPT_APPROVERS` を VPS の worker が今と同じ規則で同期し、DB の同じタブにも今の形で書く（戻した GAS がそのまま動くため）。
- 残業と違って **Sebastian の doPost・旧フォーム・朝バッチが無い**。その分は小さい。ただし inout-request がこのシートに書いているなら `sheet_intake` が要る（`01` の 3 節）。

## 5. 日ごとの進め方（残業の実績を下敷きにした案）

残業の実績: 棚卸しと計画書のコミット（11/26 05:54）から本番切り替えの完了（09:46）まで約 4 時間。担当ごとの作業はコミット時刻で 15〜60 分（`03-reuse-map.md` 3 節）。このアプリは公開関数 27（残業 29＋Sebastian）、Sebastian・フォーム・朝バッチが無いぶん小さく、PDF が一時シート方式・サインの canvas・集計タブの全書き直しの分が新しい。

| 日 | 時間の目安 | やること | 完了の条件 | 人の手が要るところ |
|---|---|---|---|---|
| 1 日目 午前 | 2〜3 時間 | 段階 0〜1: `05` の初日の確かめ（C-1〜C-12）→ 4 人の棚卸し（業務／画面／メール・PDF・台帳／データ・マスタ・運用と経緯）→ 計画書と抜け調査（5 観点）→ `04` の質問 | 計画書 9 章が埋まる（未決は切り替え日だけ）。確定値表 4-0 が埋まる | オーナー: 質問への回答、どちらのデプロイが本番か、Sebastian の inout-request を読んでよいか |
| 1 日目 午前（並行） | 1 時間 | Google・Cloudflare・VPS の準備: OAuth クライアント（GCP は本人確認が出るので先に。罠 70）、Tunnel・DNS、`/opt/kyuka`、ステージング（DB のコピー → Settings の書き換え（`05` の C-8）→ 写しの GAS） | 値を見せて声かけ済み。`/healthz` が `env: staging` | オーナー: GCP の再認証、委任の許可（`/auth/google-delegate/start`）、写しの GAS の「許可」 |
| 1 日目 昼 | 30〜60 分 | 段階 2: 仕様書（RPC の閉じたリスト、関数ごとの返す形、担当をまたぐ関数、GAS の内部用の入口の取り決め、型の表）。スパイク: 写しの GAS で PDF を 3 件作り本番の PDF と見比べ、所要時間を測る | `docs/vps-implementation-spec.md` | — |
| 1 日目 午後 | 3〜4 時間 | 段階 3: F・U・G を並行 → W・D・M を並行 → C → レビュー 2 本（正しさ・安全）→ R | 全試験が通る。残業・貯蔵品の固有の文字列が grep で 0 件 | — |
| 1 日目 夕方 | 2〜3 時間 | 段階 4〜5: VPS に置く、本番の写しを移す、`tools/qa/acceptance.py`、総当たり比較、JEV ハンズで画面を押す（申請→プレビュー→サイン（電子印と手書き）→2 次承認→PDF→編集→取消、管理画面 3 つ、取得状況）、手順書のリハーサル | 受け入れ表がすべて合格。「実機で押した／自動テストだけ／まだ」を分けて報告 | オーナー: 試験用でボタンを押してよいか（メールはオーナーにだけ届く）、スマホ・タブレット（サインの canvas）の実機 |
| 1 日目 夜（業務の後） | 30〜60 分 | 段階 6: 切り替え（残業の `cutover-runbook.md` B 章の型。VPS を止めてから移転モード → トリガーを消す → 控え → 移行と照合 → api だけ起動して安全装置 → worker → ポータル） | 照合一致、古いリンク（`?page=preview&reqId=…`）が新しいアプリの同じ申請を開く | オーナー: 始める合図、GAS 本番デプロイの更新、ポータルの書き換えの了承 |
| 2〜6 日目 | 毎日 15 分 | 段階 7: 見守り 5 営業日。統合業務メール（7:00・17:00・18:00）が今と同じ件数を出すか、PDF・台帳の仕事の失敗、行の反映の遅れ、マスタ同期（6 時台）、バックアップ | 期間を問題なく過ぎる。ステージングを片付ける | オーナー: 戻すかどうかの判断（問題があったときだけ） |

切り替えの日は、**次の有給警告の送信日（2027-01-01 の朝の報告）より前**なら、V_LEAVE_SUMMARY の反映を見守り期間中に一度確かめておく（`05` の R-5）。

## 6. 確かめること（この資料を作った時点で残っているもの）

- どちらのデプロイ（@106・@104）を利用者・ポータル・`APP_URL` が使っているか。
- M_SYSTEM_SETTING の全キーと値（`MAIL_TO`・`SOMU_EMAILS`・`ADMIN_EMAILS`・`ALLOW_ADMIN_PAGE_EMAILS`・`PDF_FOLDER_ID`・`LEDGER_SSID`・`APP_URL`・`PORTAL_URL`・同期元の ID）。
- DB のファイルの時間帯、タブの一覧、T_LEAVE_REQUEST の見出し（英語か日本語か）と件数・状態別の数。
- 登録されているトリガーの実物。
- Sebastian の inout-request がこのシートに書くか（`01` の 3 節）。
