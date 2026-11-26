# 休暇届アプリ VPS 版 運用手順（骨組み）

デプロイ・戻し・秘密の置き場の手順です。値の正は `docs/vps-migration/06-day1-findings-and-decisions.md` と計画書・実装仕様（リポジトリの直下の docs）です。作業の前に必ず読み合わせます。切り替えの手順（移転モード・流しきり・戻し）は後の段階で足します。

## 最初に確かめること（毎回）

- `ssh lineworks-vps-user "hostname"` が `srv1508169` を返すこと。違えばすぐ中止します。
- `srv1300618` には絶対に触れません（別会社の本番が同居しています）。
- 置き場は `/opt/kyuka`。リポジトリの `vps-app/` の**中身**をここに直置きします（git のフォルダではありません。`.env` も直下、権限 600）。ほかの `/opt/*` とコンテナには触れません。
- docker は `sg docker -c '...'` で動かします（lineworks ユーザーは sudo 不可）。
- コンテナ・ボリューム・ネットワークは `kyuka-` で始まるものだけを扱います（`kyuka-db`・`kyuka-api`・`kyuka-worker`・`kyuka-web`・`kyuka-cloudflared`・`kyuka-backup`、`kyuka-pgdata`、`kyuka-net` 10.94.0.0/24、web の固定 IP 10.94.0.10）。

## 秘密の置き場

| もの | 置き場 | 備考 |
|---|---|---|
| `.env`（DB のパスワード、OAuth の secret、SESSION_SECRET、FERNET_KEY、GAS_INTERNAL_TOKEN、TUNNEL_TOKEN） | VPS の `/opt/kyuka/.env`（`chmod 600`） | git に入れない。見本は `.env.example`。確かめるときは `grep -E '^(APP_ENV\|APP_IMAGE_TAG)=' .env` のように秘密でない値だけを出す |
| 手元の控え | `C:\Users\imaizumi.LINEWORKS-NET\.secrets\kyuka\` | git の外。Tunnel のトークンは `tunnel-token.env` |
| 委任トークン（`imaizumi@lineworks-local.info` のスプレッドシート・ドライブ） | DB の `google_credentials` 表（FERNET_KEY で暗号化） | FERNET_KEY を変えたら、オーナーが `/auth/google-delegate/start` から許可し直す |

## 初回の準備（ディレクターが行う。実行前にオーナーへ値を見せて一声かける）

1. `/opt/kyuka` を作る（lineworks 所有、750）。Cloudflare Tunnel `leave-vps`（`leave.lineworks-local.net` → `http://web:80`）は作成済み（`cloudflare/setup_tunnel.py`）。
2. 手元の `.env`（`.env.example` から作る）を送る: `ssh lineworks-vps-user 'umask 077; cat > /opt/kyuka/.env && chmod 600 /opt/kyuka/.env' < vps.env`
3. OAuth クライアント（GCP の「Time Off Request (VPS)」）のリダイレクト URI は登録済み。
   - `https://leave.lineworks-local.net/auth/callback`
   - `https://leave.lineworks-local.net/auth/google-delegate/callback`
4. 下の「デプロイ」で起動する（staging。分からない ID は罠 50 のとおり仮の値で起動してよい）。
5. オーナーに、ログインしてから `https://leave.lineworks-local.net/auth/google-delegate/start` を開いてスプレッドシート・ドライブを許可してもらう（未ログインで開くと 1 回目はログインだけで終わるので、もう一度同じアドレスを開く）。
6. GAS の内部用の入口は `GAS_INTERNAL_AUTH=bearer`（第 1 案・既定。第 2 案は `anonymous`）で、委任のアクセストークンを付けて呼ぶ（GAS の Web アプリが `access: DOMAIN` のため）。ステージングで `internal.ping` が通るかを確かめ、ログイン画面（HTML）が返るなら `gas_ping` の記録に「HTML」と出る。

## デプロイ（git archive 方式）

手元のリポジトリで、**コミット済みの** `vps-app/` を送ります（未コミットの変更は送られません。手元で消したファイルは VPS に残るので、要らなくなったものは VPS でも消します）。

```sh
cd /d/projects/kyuka-todoke-app/repo
git archive --format=tar HEAD vps-app | ssh lineworks-vps-user 'test "$(hostname)" = srv1508169 || exit 9; rm -rf /tmp/kyuka-sync && mkdir -p /tmp/kyuka-sync && tar -x -C /tmp/kyuka-sync && cp -a /tmp/kyuka-sync/vps-app/. /opt/kyuka/ && rm -rf /tmp/kyuka-sync && echo synced'
```

イメージにタグを付けて出し、タグは `.env` の `APP_IMAGE_TAG` に書きます（コマンド行だけで渡すと、後の `docker compose up -d` で古いタグに戻ります＝罠 24e）。

```sh
TAG=$(git rev-parse --short HEAD)    # 手元で求める（VPS は git のフォルダではない）
ssh lineworks-vps-user 'grep -E "^APP_IMAGE_TAG=" /opt/kyuka/.env; sg docker -c "docker image ls kyuka-app"'   # 今のタグを控える
ssh lineworks-vps-user "test \"\$(hostname)\" = srv1508169 || exit 9; cd /opt/kyuka && sed -i 's/^APP_IMAGE_TAG=.*/APP_IMAGE_TAG=$TAG/' .env && sg docker -c 'docker compose build api' && sg docker -c 'docker compose up -d' && sg docker -c 'docker compose ps'"
```

- build と起動は時間がかかります。`timeout 590 ssh ...` を裏で流します（罠 2）。
- api は起動のたびに `alembic upgrade head` を流してから立ち上がります。
- 設定の値が欠けている・危ない組み合わせ（staging なのに本番の ID、本番なのにステージングの GAS など）のときは、api と worker は起動しません（ログに名前が出ます）。
- 確かめ: `sg docker -c "docker compose exec -T web wget -qO- http://127.0.0.1/healthz"` が `{"ok":true,"env":"staging"}`。外から `curl -s https://leave.lineworks-local.net/healthz` が 200。
- 入口の確かめ（curl）: `/` 302（ログインへ）、ログインなしの RPC 401、`/auth/login` 302（`hd`・PKCE・state・nonce 付き、`prompt` なし）、CSP などのヘッダ、JSON でない本文 415、別サイトの Origin 403、知らない関数 404。
- 初めて出すとき・`KYUKA_NET_SUBNET` を変えたときは `docker compose down` → `up -d` でネットワークを作り直します。
- 切り替えでは、先にシートの控え → **api だけ**起動して安全装置を確かめる → worker を起動（罠 24）。

## 試験

- 手元: `cd vps-app/api && .venv/Scripts/python -m pytest -q`（SQLite。最後の行の passed/failed まで読む＝罠 44）。
- VPS の結合試験（PostgreSQL）: `docker build --target test -t kyuka-test ./api` の後、DB `kyuka_test` を作り `TEST_DATABASE_URL` を渡して `python -m pytest`（罠 51。DB 名に test が無ければ conftest が拒む。`@pytest.mark.postgres` の試験は `TEST_DATABASE_URL` があるときだけ動く）。

## worker と定期処理

`python -m app.worker run <名前>` で手で動かせます。結果は `sync_runs`（kind=`schedule:<名前>`）に残ります。予定は担当 D・M が `register_worker()` で登録します（`app/worker.py` の REGISTRATION_MODULES）。土台が持つのは次の 1 つだけです。

| 名前 | いつ（日本時間） | 持ち主 |
|---|---|---|
| `purge` | 毎日 3:30（冪等キー 24 時間・期限切れのセッション・done の仕事 90 日。audit_log は消さない） | F |

GAS の仕事（`jobs` の `lane='gas'`）は全体で同時に 1 本。

- シートへ書く前の関所（`services/write_gate`）: 本番は、合格した `gas_ping` が `SHEET_WRITE_PING_MAX_MINUTES`（既定 10）分以内で、GAS が移転モード（`migrationMode=true`）のときだけ、行の反映・照合・集計タブ・マスタ同期がシートへ書きます。ステージングは、ステージングの GAS への合格した `gas_ping` があれば書きます。
- 書いてよいタブ: `T_LEAVE_REQUEST`（行単位）、`V_LEAVE_SUMMARY`（全行の書き直し）、`作業員マスタ`・`M_CALENDAR`・`M_STAMP`・`M_DEPT_APPROVERS`（マスタ同期）。`M_SYSTEM_SETTING`・`M_LOOKUP` は読むだけ。台帳・PDF フォルダ・マスタ元・読み先には書きません（`app/safety.py`）。
- `SKIP_DAILY_ON_STARTUP=1` で起動すると、起動した時に枠の中にある毎日の予定を、その回だけ飛ばします（`sync_runs` に `skipped`）。切り替えの夜に使い、済んだら `.env` から消します。

## ログインの取り消し

- 全員: `sg docker -c "docker compose exec -T db psql -U kyuka -d kyuka -c 'DELETE FROM sessions'"`（全員がもう一度ログインします）。
- 1 人（そのメールの全端末）: `auth.revoke_sessions(db, email)`（管理者の関数から呼ぶ。担当 W）。
- ログインの成功・失敗は api のログ（`login ok`・`login rejected`）に残ります。

## バックアップ

- `kyuka-backup` が毎日 3:00 に `backups/kyuka_*.dump` を取り、`BACKUP_KEEP_DAYS` 日で消します。1KB 未満は `FAILED`（初回起動の直後は正常＝罠 52）。
- 切り替えの控えなど手で取ったダンプは `/opt/kyuka/cutover/` に置きます（backups/ は消える＝罠 25）。

## 戻す（骨組み）

- アプリの版を戻す: `.env` の `APP_IMAGE_TAG` を控えたタグに書き戻して `docker compose up -d`。
- 本番の切り替えの戻し（GAS に戻す）は後の段階で書きます。
