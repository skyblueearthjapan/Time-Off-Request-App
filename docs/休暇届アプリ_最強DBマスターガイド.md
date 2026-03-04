# 休暇届アプリ：DB＋PDFテンプレート ブック全体マップ（最強DBマスターガイド）
このドキュメントは、GASコーディングエージェントが実装できるように、ブック内の全シート/列/役割/差し込みセルを網羅してまとめたものです。

## 1. ブック概要
- 対象ファイル: 休暇届アプリ_DB＋PDFテンプレート_雛形.xlsx
- シート数: 9
- シート一覧: README, CONFIG, LOOKUP, M_DEPT, M_WORKER, T_LEAVE_REQUEST, V_LEAVE_SUMMARY, PDF_TEMPLATE, PDF_MAP

## 2. シート別：用途・構造（列マップ）
### README
- 役割: 説明
- 最大行/列（Excel上）: rows=10, cols=11
- テーブル想定: header_row=1 / last_col=A(1) / last_data_row=10

| No | 列 | ヘッダ名 | 備考（用途/想定） |
|---:|:--:|---|---|
| 1 | A | 休暇届アプリ：DBスプレッドシート＋PDFテンプレート（Excel雛形） |  |

### CONFIG
- 役割: システム設定
- 最大行/列（Excel上）: rows=20, cols=13
- テーブル想定: header_row=1 / last_col=C(3) / last_data_row=20

| No | 列 | ヘッダ名 | 備考（用途/想定） |
|---:|:--:|---|---|
| 1 | A | KEY | KEYに対する設定値。GASはここを参照。 |
| 2 | B | VALUE | KEYに対する設定値。GASはここを参照。 |
| 3 | C | NOTE | KEYに対する設定値。GASはここを参照。 |

### LOOKUP
- 役割: 固定値/選択肢
- 最大行/列（Excel上）: rows=5, cols=19
- テーブル想定: header_row=1 / last_col=I(9) / last_data_row=5

| No | 列 | ヘッダ名 | 備考（用途/想定） |
|---:|:--:|---|---|
| 1 | A | 休暇区分 |  |
| 2 | C | 半日区分 |  |
| 3 | E | 休暇種類（全日） |  |
| 4 | G | 休暇種類（半日） |  |
| 5 | I | 特別休暇 参考例 |  |

### M_DEPT
- 役割: マスタ（同期/管理）
- 最大行/列（Excel上）: rows=400, cols=15
- テーブル想定: header_row=1 / last_col=E(5) / last_data_row=3

| No | 列 | ヘッダ名 | 備考（用途/想定） |
|---:|:--:|---|---|
| 1 | A | DEPT_ID |  |
| 2 | B | DEPT_NAME |  |
| 3 | C | SORT |  |
| 4 | D | IS_ACTIVE |  |
| 5 | E | UPDATED_AT |  |

### M_WORKER
- 役割: マスタ（同期/管理）
- 最大行/列（Excel上）: rows=800, cols=17
- テーブル想定: header_row=1 / last_col=G(7) / last_data_row=3

| No | 列 | ヘッダ名 | 備考（用途/想定） |
|---:|:--:|---|---|
| 1 | A | WORKER_ID |  |
| 2 | B | WORKER_NAME |  |
| 3 | C | DEPT_ID |  |
| 4 | D | DEPT_NAME_SNAP |  |
| 5 | E | IS_ACTIVE |  |
| 6 | F | UPDATED_AT |  |
| 7 | G | NOTE |  |

### T_LEAVE_REQUEST
- 役割: トランザクション（申請データ）
- 最大行/列（Excel上）: rows=501, cols=32
- テーブル想定: header_row=1 / last_col=V(22) / last_data_row=501

| No | 列 | ヘッダ名 | 備考（用途/想定） |
|---:|:--:|---|---|
| 1 | A | REQ_ID | 運用の中核項目 |
| 2 | B | FY_START_YEAR | 運用の中核項目 |
| 3 | C | STATUS | 運用の中核項目 |
| 4 | D | APPLIED_AT | 運用の中核項目 |
| 5 | E | DEPT_ID |  |
| 6 | F | DEPT_NAME |  |
| 7 | G | WORKER_ID |  |
| 8 | H | WORKER_NAME |  |
| 9 | I | LEAVE_DATE | 申請入力（モーダル） |
| 10 | J | DAY_TYPE | 申請入力（モーダル） |
| 11 | K | HALF_TYPE | 申請入力（モーダル） |
| 12 | L | LEAVE_KIND | 申請入力（モーダル） |
| 13 | M | SUBSTITUTE_FOR_WORK_DATE | 条件付き入力（振替/特別/有給詳細/任意詳細） |
| 14 | N | SPECIAL_REASON_REF |  |
| 15 | O | SPECIAL_REASON_TEXT | 条件付き入力（振替/特別/有給詳細/任意詳細） |
| 16 | P | PAIDLEAVE_DETAIL | 条件付き入力（振替/特別/有給詳細/任意詳細） |
| 17 | Q | DETAIL_FREE | 条件付き入力（振替/特別/有給詳細/任意詳細） |
| 18 | R | APPROVED_AT | 運用の中核項目 |
| 19 | S | SIGN_IMAGE_URL | 運用の中核項目 |
| 20 | T | PDF_URL | 運用の中核項目 |
| 21 | U | CREATED_BY_EMAIL |  |
| 22 | V | UPDATED_AT |  |

### V_LEAVE_SUMMARY
- 役割: ビュー/集計（参照）
- 最大行/列（Excel上）: rows=1, cols=21
- テーブル想定: header_row=1 / last_col=K(11) / last_data_row=1

| No | 列 | ヘッダ名 | 備考（用途/想定） |
|---:|:--:|---|---|
| 1 | A | FY_START_YEAR |  |
| 2 | B | DEPT_ID |  |
| 3 | C | DEPT_NAME |  |
| 4 | D | WORKER_ID |  |
| 5 | E | WORKER_NAME |  |
| 6 | F | PAIDLEAVE_COUNT |  |
| 7 | G | SPECIAL_COUNT |  |
| 8 | H | SUBSTITUTE_COUNT |  |
| 9 | I | TOTAL_COUNT |  |
| 10 | J | WARN_LEVEL |  |
| 11 | K | LAST_PAIDLEAVE_DATE |  |

### PDF_TEMPLATE
- 役割: PDF出力テンプレ
- 最大行/列（Excel上）: rows=43, cols=15
- テーブル想定: header_row=1 / last_col=A(1) / last_data_row=6

| No | 列 | ヘッダ名 | 備考（用途/想定） |
|---:|:--:|---|---|
| 1 | A | 休暇届（全日休・半日休） |  |

### PDF_MAP
- 役割: テンプレ差し込み対応表
- 最大行/列（Excel上）: rows=18, cols=15
- テーブル想定: header_row=1 / last_col=E(5) / last_data_row=18

| No | 列 | ヘッダ名 | 備考（用途/想定） |
|---:|:--:|---|---|
| 1 | A | FIELD_KEY |  |
| 2 | B | TARGET_CELL | PDF_TEMPLATEへの差し込みセル座標 |
| 3 | C | DESCRIPTION |  |
| 4 | D | SOURCE（T_LEAVE_REQUEST列など） |  |
| 5 | E | NOTE |  |

## 3. 申請DB：T_LEAVE_REQUEST 詳細（列番号・用途・必須/条件）
| ColNo | 列 | カラム名 | 生成/入力元 | 条件 |
|---:|:--:|---|---|---|
| 1 | A | REQ_ID | — | — |
| 2 | B | FY_START_YEAR | LEAVE_DATEから算出（4/1開始） | — |
| 3 | C | STATUS | 申請/承認/却下/取消 | — |
| 4 | D | APPLIED_AT | 申請確定時（サーバ時刻） | — |
| 5 | E | DEPT_ID | TOPのプルダウン選択（マスタ由来） | — |
| 6 | F | DEPT_NAME | 選択時にスナップショット保存 | — |
| 7 | G | WORKER_ID | TOPのプルダウン選択（マスタ由来） | — |
| 8 | H | WORKER_NAME | 選択時にスナップショット保存 | — |
| 9 | I | LEAVE_DATE | モーダル入力（日付） | — |
| 10 | J | DAY_TYPE | モーダル入力（全日休/半日休） | — |
| 11 | K | HALF_TYPE | モーダル入力（午前/午後） | DAY_TYPE=半日休のとき必須 |
| 12 | L | LEAVE_KIND | モーダル入力（種類） | DAY_TYPEで選択肢が変化 |
| 13 | M | SUBSTITUTE_FOR_WORK_DATE | モーダル入力（振替元出勤日） | LEAVE_KIND=振替休暇のとき必須（半日には出ない） |
| 14 | N | SPECIAL_REASON_REF | モーダル入力（特別休暇の参考/自由記述） | LEAVE_KIND=特別休暇のとき任意（ただし運用で必須化可） |
| 15 | O | SPECIAL_REASON_TEXT | モーダル入力（特別休暇の参考/自由記述） | LEAVE_KIND=特別休暇のとき任意（ただし運用で必須化可） |
| 16 | P | PAIDLEAVE_DETAIL | モーダル入力（有給詳細：病欠詳細など） | LEAVE_KIND=有給消化(私用)のとき任意 |
| 17 | Q | DETAIL_FREE | モーダル入力（任意詳細） | 『詳細追加Yes』のときのみ |
| 18 | R | APPROVED_AT | サイン確定時 | — |
| 19 | S | SIGN_IMAGE_URL | サイン画像をDrive保存したURL | — |
| 20 | T | PDF_URL | PDF生成後のDrive URL | — |
| 21 | U | CREATED_BY_EMAIL | GASが自動セット | — |
| 22 | V | UPDATED_AT | GASが自動セット | — |

## 4. PDF出力：PDF_TEMPLATE × PDF_MAP（差し込みセル完全対応）
### 4.1 PDF_MAP一覧（FIELD_KEY → TARGET_CELL）
| FIELD_KEY | TARGET_CELL | 説明 | 主な入力元 | NOTE |
|---|---|---|---|---|
| LEAVE_START_REIWA_YEAR | B6 | 開始：令和年（数値） | LEAVE_DATE | 西暦→令和変換はGAS側 |
| LEAVE_START_MONTH | E6 | 開始：月 | LEAVE_DATE | None |
| LEAVE_START_DAY | H6 | 開始：日 | LEAVE_DATE | None |
| LEAVE_START_WDAY | K6 | 開始：曜日（例：月） | LEAVE_DATE | None |
| LEAVE_END_REIWA_YEAR | B7 | 終了：令和年 | LEAVE_DATE（全日なら同日可） | None |
| LEAVE_END_MONTH | E7 | 終了：月 | LEAVE_DATE | None |
| LEAVE_END_DAY | H7 | 終了：日 | LEAVE_DATE | None |
| LEAVE_END_WDAY | K7 | 終了：曜日 | LEAVE_DATE | None |
| LEAVE_END_HOUR | M7 | 終了：時 | 半日区分/固定 | 午前/午後で固定時刻にしてもOK |
| LEAVE_END_MIN | B8 | 終了：分 | 半日区分/固定 | None |
| LEAVE_REASON_SHORT | A14 | 理由（短文） | LEAVE_KIND / SPECIAL_REASON_TEXT / PAIDLEAVE_DETAIL | 「私用」「結婚」「忌引」など |
| SUBMIT_REIWA_YEAR | J30 | 届出日：令和年 | APPLIED_AT | None |
| SUBMIT_MONTH | M30 | 届出日：月 | APPLIED_AT | None |
| SUBMIT_DAY | H31 | 届出日：日 | APPLIED_AT | None |
| WORKER_NAME | J34 | 氏名 | WORKER_NAME | None |
| APPROVAL_SIGN_BOSS | K46 | 部長サイン（画像貼り付け） | SIGN_IMAGE_URL | GASで画像貼付 |
| APPROVAL_SIGN_SECTION | M46 | 所属長サイン（画像貼り付け） | SIGN_IMAGE_URL | 運用に合わせて |

### 4.2 PDF_TEMPLATE：印字領域/目安
- 推奨印字範囲（雛形）: A1:O52（A4縦）
- 主要フィールド座標は **PDF_MAP** を正とします。

## 5. 運用設定：CONFIG（GASが参照するキー一覧）
| KEY | VALUE（例） | NOTE |
|---|---|---|
| ADMIN_EMAILS | imaizumi@lineworks.co.jp | 管理者（複数はカンマ区切り） |
| SOMU_EMAILS |  | 総務（複数はカンマ区切り） |
| ALLOW_ADMIN_PAGE_EMAILS | imaizumi@lineworks.co.jp | 管理ページ閲覧OKのGoogleアカウント（管理者＋総務） |
| MAIL_TO | imaizumi@lineworks.co.jp, | 申請通知/警告通知の宛先（カンマ区切り） |
| MAIL_FROM_NAME | 休暇届システム | メール差出人名 |
| PDF_FOLDER_ID |  | PDF保存先DriveフォルダID |
| FY_START_MONTH | 4 | 年度開始月（4=4月） |
| FY_START_DAY | 1 | 年度開始日（1=1日） |
| WARN_1_DEADLINE | 06-30 | 有給しきい値① 期限（月-日） |
| WARN_1_MIN_COUNT | 1 | 期限①までの必要回数 |
| WARN_2_DEADLINE | 08-31 | しきい値② |
| WARN_2_MIN_COUNT | 2 |  |
| WARN_3_DEADLINE | 12-31 | しきい値③ |
| WARN_3_MIN_COUNT | 3 |  |
| WARN_4_DEADLINE | 02-15 | しきい値④ |
| WARN_4_MIN_COUNT | 4 |  |
| WARN_5_DEADLINE | 03-31 | 年度末しきい値（参考） |
| WARN_5_MIN_COUNT | 5 |  |
| WARN_MAIL_TIMING | month_start | 警告メール送信：month_start / month_end / due_date |
