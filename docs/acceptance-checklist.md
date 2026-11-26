# 受け入れ表（休暇届 VPS 版）

作成: 2026-11-26（担当 C）。もと: `docs/vps-implementation-spec.md` 12-3（A-1〜A-29）、`docs/vps-migration/06-day1-findings-and-decisions.md`（K-0〜K-11）、残業の `docs/acceptance-checklist.md`。
切り替え（`docs/cutover-runbook.md` の A-1）の前に、すべての行の「結果」が「済」になっていることを確かめます。

## 使い方

- 1 つの項目を、3 つの確かめ方に分けて書きます。どれかが「—」の項目は、その確かめ方を使いません。
  - **自動**: 手元の自動テスト（`cd vps-app/api && .venv/Scripts/python -m pytest -q`。担当と試験のファイル名）と、ステージングで流す `vps-app/tools/qa/acceptance.py`・`compare_views.py` の項目の番号（`H-`・`E-`・`S-`・`P-`・`F-`・`O-`・`N-`・`B-`・`PM-`・`W-`・`C-`・`G-`）。
  - **JEV**: JEV ハンズで実機の画面を押して確かめる操作（下の 2 節の J- の番号）。押すのはオーナーの了承を得てから。
  - **オーナーの実機**: スマホ・タブレットの実物（JEV では動かせないもの。下の 3 節の O- の番号）。オーナーに写真か一言をもらう。
- 「結果」は 未／済／NG のどれかにし、済にした日と、確かめた人を「メモ」に書きます。
- 報告では「実機で押した／自動だけ／まだ」を分けて書きます。自動のテストが通っても、JEV とオーナーの実機の欄が「済」になるまで、その項目は「済」にしません。
- ステージングのメールはオーナーの 1 件（ステージングの `MAIL_TO`）にだけ届きます。`acceptance.py --write` は承認完了メールを 2 通出します。

### 道具の動かし方（ステージング。VPS の api コンテナの中）

```sh
ssh lineworks-vps-user 'bash -s' <<'EOF'
set -eu
test "$(hostname)" = srv1508169 || { echo "WRONG HOST"; exit 9; }
cd /opt/kyuka
mkdir -p cutover
# 読むだけ（書く関数は「存在しない申請」「入力の誤り」で止まることだけを見る）
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session" < /dev/null
# 書き込みの流れ（申請 4 件 → 承認 → 2 次承認 → 編集 → 取消。最後に全部取消）
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro api python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session --write" < /dev/null
# 読む関数の総当たり（VPS 版の返事と plan を保存。GAS 版があれば --gas-file で突き合わせる）
sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/kyuka/tools:/app/tools:ro -v /opt/kyuka/cutover:/out api python tools/qa/compare_views.py --base-url http://api:8000 --mint-session --out /out/compare_vps.json --plan-out /out/compare_plan.json" < /dev/null
EOF
```

- どの道具も `/healthz` が `staging` でなければ何もせずに終了コード 2 で止まります（本番では動きません）。
- `--mint-session` は BOOTSTRAP_ADMINS の 1 人目のセッションと、権限の無い架空のメール（`acceptance-outsider@lineworks-local.info`）のセッションを DB に作り、終わったら必ず消します（クッキーの値は表示しません）。
- 結果は表（項目・期待・実際）で出て、FAIL が 1 つでもあれば終了コード 1 です。最後の行の `PASS n / FAIL 0` と `session removed (…): True` を読みます。
- GAS 版の返事の取り方は `compare_views.py` の docstring にあります（写しの GAS のエディタで plan を流してドライブに保存）。

## 1. 実装仕様 12-3 の受け入れ表

| # | 項目 | 自動（pytest・acceptance・compare） | JEV ハンズで実機 | オーナーの実機 | 結果 | メモ |
|---|---|---|---|---|---|---|
| A-1 | 7 ページ＋権限エラー、`?page=`、不明は top、題名 | U `test_pages.py`。acceptance P-1〜P-7・E-1・E-2 | J-1（各ページを開き、ヘッダーのリンクを全部押す） | — | 未 | |
| A-2 | 入口の判定（admin・deptAdmin・somuAdmin） | U `test_pages.py`。acceptance PM-14 | J-14（権限の無い共有アカウントで 3 つの文と題名） | — | 未 | |
| A-3 | 27 関数の形・エラー文 | W `test_domain_*.py`。acceptance F-*・N-1〜N-8・F-all（34 関数）。compare G-1（GAS 版と全部署×全年度、作業員の抜き取り、全 REQ_ID の preview） | — | — | 未 | |
| A-4 | 申請（全日有給・半日特別・振替の期間の中と外・過去申請） | W `test_domain_requests.py`。acceptance W-1〜W-4・B-1〜B-7 | J-2〜J-5 | — | 未 | |
| A-5 | 二重押し | F `test_rpc.py`・U `shim.test.js`・W `test_concurrency.py`。acceptance W-1（同じ冪等キーで 1 件） | J-6（「申請する」を素早く 2 回で 1 件） | — | 未 | |
| A-6 | 編集・取消（申請中・承認済・2 次承認済） | W・G `internal.test.js`。acceptance W-15・W-16・W-20〜W-25（台帳の行を消す仕事、サインの列が空、PDF_URL は残る＝asis） | J-10・J-11 | — | 未 | |
| A-7 | 1 次承認（電子印・手書き）とサイン | W・G。acceptance W-10〜W-13（サイン画像の URL・承認完了メール・台帳の仕事） | J-7・J-8（PC のマウスで手書き） | O-2（**スマホ・iPad の指とペンで手書き**） | 未 | |
| A-8 | 1 次承認の状態の確かめ（K-4） | W `test_concurrency.py`。acceptance W-10（承認済みをもう一度 → `この申請はすでに承認済みです。`）・W-15（取消済み → `この申請は取消済みです。`）・W-27（メールは承認の回数だけ） | J-9（2 つのタブで同じ申請を承認、2 回目が「すでに承認済み」、メール 1 通） | — | 未 | |
| A-9 | 2 次承認（1 件・一括）と PDF | W・G・M。acceptance W-9・W-14・W-17〜W-19 | J-12（1 件と一括。PDF の見た目を本番の PDF と見比べる：印・令和・時間帯・理由） | — | 未 | |
| A-10 | 集計・取得状況（somuAdmin・deptAdmin・admin・leaveStatus） | W `test_summaries.py`。acceptance F-api_getSomuAdminView ほか。compare G-1・C-4 | J-13 | — | 未 | |
| A-11 | V_LEAVE_SUMMARY と統合業務メール | D `test_summary_sheet.py`。acceptance W-6・W-11・W-19・W-23（シートの行を統合業務メールの条件で読む）。手順書の「シートを数える」 | — | — | 未 | SB: VPS が書いたタブを GAS の `rebuildLeaveSummary_` の結果と 1 セルずつ比べる |
| A-12 | Safari の日付 | U `web/dev/pages.test.js` | — | O-1（iPhone の preview の届出日・承認日・曜日の写真） | 未 | |
| A-13 | 行の反映の型 | D `test_sheet_codec.py`・`test_row_sync.py`。acceptance W-6〜W-25 の「休暇日が日付・APPROVED_AT2」 | — | — | 未 | SB: 写しの GAS の `getValues()` で `Date`、`api_getLeaveRequestById` が元の値と同じ |
| A-14 | 行の位置・半端な行 | D `test_row_sync.py` | — | — | 未 | SB: 新しい行が最後の行の次、REQ_ID が空の 52 行が変わらない |
| A-15 | マスタ同期と DB の 4 タブ | D `test_master_sync.py`。acceptance F-ops_runMasterSync（--write） | J-15（運用の欄の［マスタ同期］） | — | 未 | SB: GAS の `syncAllMasters` の後と VPS の後でタブを 1 セルずつ比べる |
| A-16 | 設定の取り込み・権限 | F `test_permissions.py`・D `test_config_sync.py`。acceptance F-ops_syncSheetConfigNow（--write）・PM-0〜PM-13 | J-16（`SOMU_EMAILS` を変えて［今すぐ取り込む］で somuAdmin の入口が変わる） | — | 未 | |
| A-17 | GAS が止まったとき | M `test_gas_jobs.py` | J-17（写しの GAS の token を外して承認 → 運用の欄に出る → 戻すと流れる。メール 1 通） | — | 未 | |
| A-18 | 移転モード | G `cutover.test.js` | J-18（リハーサル: @106・@104 の `/exec?page=preview&reqId=LV-…` から新しいアプリの同じ申請が最上位の窓で開く。メールの `?page=somuAdmin` も） | — | 未 | |
| A-19 | 内部用の入口の呼び方 | M `test_spike_gas_auth.py`。acceptance O-1（gas_ping） | — | — | 済（2026-11-26） | スパイク S-1 で `bearer` に決定（06 の 7 節） |
| A-20 | **0:30 JST の申請** | W `test_fiscal_jst.py` | J-19（0 時台に 1 件出し、REQ_ID の日付・年度・シートの休暇日を見る） | — | 未 | |
| A-21 | CSP で全ボタンが動く | F `test_web_config.py`・U。acceptance E-10・E-11・P-1〜P-7（インラインの script と on 属性が無い） | J-1〜J-16 の全ボタン・選択・サインの描画 | — | 未 | |
| A-22 | スマホ・タブレット | F `test_web_config.py`（`mobile.css`） | J-20（狭い幅。PC の 1366px で今の画面と見比べて差が無い） | O-1〜O-4（iPhone・現場の Android・タブレットで申請・サイン・2 次承認） | 未 | |
| A-23 | XSS | U `web/dev/pages.test.js`。acceptance W-4・W-6D・W-8（氏名の `<img …>&"'</script>` が文字のまま） | J-21（その氏名の申請を作り、6 画面を開いて文字のまま） | — | 未 | |
| A-24 | 401 と書きかけ | U `shim.test.js` | J-22（mock-server の `status=401` で申請モーダル・編集・**手書きサイン**が戻る） | — | 未 | |
| A-25 | ステージングから本番に書かない | F `test_config_safety.py`。手順書 B-14 の安全装置の確かめ（ステージングで流しても `OK: refused`） | — | — | 未 | |
| A-26 | 運用の欄 | W `test_domain_ops.py`。acceptance PM-7〜PM-13（運用者でない人は 403） | J-23（運用者で somuAdmin の欄が出る、ほかでは出ない。有給警告メールは送信日以外「送信日ではありません」相当） | — | 未 | |
| A-27 | 移行の照合 | D `test_migrate_from_sheets.py`・`test_reconcile.py` | — | — | 未 | 手順書 B-12・B-13（`verifyProblems: []`、`--check` 差 0） |
| A-28 | 速さ | acceptance の最後の「速さ（中央値 ms）」（目安: 一覧 100ms 以下、承認 100ms 以下） | — | — | 未 | |
| A-29 | 手順書のリハーサルと戻しの練習 | — | — | — | 未 | ディレクターがステージングで 2 回（手順書 A-3） |

## 2. JEV ハンズで実機で押す操作（ステージング。押す前にオーナーの了承をもらう）

申請・承認・2 次承認・取消のボタンは `allow_sensitive=true` が要ります。「ステージングのアプリで、これらのボタンを押してよいか。承認完了メールは○通ほど、オーナーにだけ届く」と 1 回で聞きます。作る申請の作業員名か有給詳細に【受入テスト】を入れ、最後に取消します。

| # | 操作 | 確かめること | 結果 | メモ |
|---|---|---|---|---|
| J-1 | 7 ページを開き、ヘッダーのリンクを全部押す | 各ページが開き、題名が「休暇届 …」。ボタンが効く（CSP） | 未 | |
| J-2 | 全日・有給の申請 | 申請 TOP の承認待ちに出る。1 分以内にシートの末尾に行が出る | 未 | |
| J-3 | 半日（午前・午後）・特別休暇の申請 | 特別休暇の理由の選択肢が M_LOOKUP の順 | 未 | |
| J-4 | 振替休暇の申請（振替元の候補・同じ月次期間の外を選ぶと止まる） | 候補が休暇日の月次期間の休日だけ。期間の外は今と同じ文 | 未 | |
| J-5 | 過去申請 | 今年度の過去の勤務日だけ選べる。一覧に「過去申請」の印 | 未 | |
| J-6 | 「申請する」を素早く 2 回 | 申請は 1 件だけ | 未 | |
| J-7 | 1 次承認（電子印） | 承認済になる。承認完了メールが 1 通。台帳に行 | 未 | |
| J-8 | 1 次承認（手書きサイン。PC のマウス） | サインの PNG が休暇日のフォルダに 1 つ、SIGN_IMAGE_URL が入る | 未 | |
| J-9 | 2 つのタブで同じ申請を 1 次承認 | 2 回目が「この申請はすでに承認済みです。」、メール 1 通 | 未 | |
| J-10 | 承認済みの申請を編集して保存 | 申請中に戻る。台帳の行が消え、サインの列が空、PDF_URL は残る | 未 | |
| J-11 | 取消（申請中・承認済・2 次承認済）と、取消した申請の sign を開いて承認 | 取消になる。承認は「この申請は取消済みです。」 | 未 | |
| J-12 | 2 次承認（1 件と一括）、もう一度 2 次承認 | PDF ができ、台帳の 2 次承認日・2 次承認者・PDF のリンク。2 回目は「この申請はすでに2次承認済みです。」。PDF を本番の PDF と見比べる | 未 | |
| J-13 | 総務管理・部署管理・古い管理画面・取得状況（年度と部署を変える） | 件数・回数・最終有給日・警告が GAS 版（compare G-1）と同じ | 未 | |
| J-14 | 権限の無い部署の共有アカウントで admin・deptAdmin・somuAdmin を開く | 3 つの権限エラーの文と題名 | 未 | |
| J-15 | 運用の欄の［マスタ同期］・［集計タブの書き直し］ | 結果が出て、タブが変わらない（控えと同じ） | 未 | |
| J-16 | `SOMU_EMAILS` を変えて［今すぐ取り込む］ | somuAdmin の入口がすぐ変わる。元に戻す | 未 | |
| J-17 | 写しの GAS の token を外して 1 次承認 → 戻す | 画面は動き、運用の欄に失敗が出て、戻すと流れる。メール 1 通 | 未 | |
| J-18 | @106・@104 の古いリンク（リハーサル） | 移転ページのボタンで、新しいアプリの同じ申請が最上位の窓で開く | 未 | |
| J-19 | 0 時台（0:30 ごろ）の申請 | REQ_ID の日付・年度・シートの休暇日がその日 | 未 | |
| J-20 | 狭い幅（スマホの幅）と PC の 1366px | スマホ向けの配置が効く。PC は今の画面と差が無い | 未 | |
| J-21 | 氏名に `<img src=x onerror=alert(1)>` を入れた申請で 6 画面 | どの画面も文字のまま。警告が出ない | 未 | |
| J-22 | mock-server の `status=401` で申請モーダル・編集・手書きサイン | ログインの後、書きかけが戻る | 未 | |
| J-23 | 運用の欄（運用者とそうでない人）、有給警告メール（送信日以外） | 運用者にだけ出る。送信日以外は送らない | 未 | |

## 3. オーナーの実機（JEV では動かせないもの）

| # | 端末 | 確かめること | 結果 | メモ |
|---|---|---|---|---|
| O-1 | iPhone（Safari） | preview の届出日・承認日・曜日、申請 TOP の一覧、2 次承認の画面（写真） | 未 | |
| O-2 | 現場のタブレット（iPad・Android） | 手書きサインを指とペンで描いて 1 次承認できる。画面を回したときの動き（今のまま消える＝Q14） | 未 | K-7 |
| O-3 | 現場の Android スマホ | 申請と一覧、ログインが 90 日続く | 未 | |
| O-4 | PC（ふだんの端末） | 今の画面と見た目が同じ | 未 | |

## 4. 切り替えの当日に確かめること（手順書の B 章）

| # | 項目 | 確かめ方 | 結果 | メモ |
|---|---|---|---|---|
| S-1 | 進行中の件数が移行の前後で同じ | 手順書 B-1・B-13・B-17（統合業務メールの条件で本番シートを数える） | 未 | |
| S-2 | @106 と @104 の両方が移転モードの版 | 手順書 B-3（`list-deployments --json`）・B-6 | 未 | |
| S-3 | 安全装置 | 手順書 B-14（`NG` 0 行・`gas_ping` が本番の DB と移転モード） | 未 | |
| S-4 | 移行と照合 | 手順書 B-12・B-13・B-16（`verifyProblems: []`・差 0・控えとの差 0） | 未 | |
| S-5 | 古いリンク・`APP_URL`・ポータル | 手順書 B-18〜B-20 | 未 | |
| S-6 | 翌朝の統合業務メールと最初の PDF | 手順書 D 章 | 未 | |
