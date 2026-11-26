#!/usr/bin/env python3
"""休暇届アプリ VPS 版の Cloudflare 登録（Tunnel・公開ホスト名・DNS だけ）。標準ライブラリのみ。

Cloudflare Access は使わない（入口はアプリの Google ログインだけ）。
API トークンは TOKEN_ENV のファイルの CLOUDFLARE_API_TOKEN から読む（画面に出さない）。
TUNNEL_TOKEN は画面に出さず、--out のファイルにだけ書く。

安全のため、同じ名前の Tunnel（leave-vps）が既にあれば何もせずに止まる（ほかの用途の Tunnel の ingress を
上書きしないため）。作り直すときは、Cloudflare の画面で既存の Tunnel を確かめてから人が判断する。
同じホスト名の DNS が既にあるときも止まる。

  python setup_tunnel.py --dry-run          # 現状確認だけ（何も作らない）
  python setup_tunnel.py --out <file>       # 作成（Tunnel・ingress・CNAME）
"""
import json
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

API = "https://api.cloudflare.com/client/v4"
ZONE = "lineworks-local.net"
HOSTNAME = "leave.lineworks-local.net"
TUNNEL_NAME = "leave-vps"
SERVICE_URL = "http://web:80"  # docker compose の web サービス
DNS_COMMENT = "kyuka (leave request) via cloudflared tunnel leave-vps"
TOKEN_ENV = Path(r"D:\projects\Production-Schedule-Viewer\remote-access\cloudflare\.env")


def main(argv):
    for s in (sys.stdout, sys.stderr):
        if s.encoding and s.encoding.lower() not in ("utf-8", "utf8"):
            s.reconfigure(encoding="utf-8")

    dry = "--dry-run" in argv
    out = Path(argv[argv.index("--out") + 1]) if "--out" in argv else None
    if not dry and out is None:
        sys.exit("ERROR: --out <TUNNEL_TOKEN を書くファイル> を指定してください")

    token = ""
    for line in TOKEN_ENV.read_text(encoding="utf-8-sig").splitlines():
        if line.strip().startswith("CLOUDFLARE_API_TOKEN="):
            token = line.split("=", 1)[1].strip()
    if not token:
        sys.exit("ERROR: CLOUDFLARE_API_TOKEN が見つかりません")

    def api(method, path, body=None):
        req = urllib.request.Request(
            API + path,
            data=json.dumps(body).encode() if body is not None else None,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            method=method,
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                doc = json.load(r)
        except urllib.error.HTTPError as e:
            sys.exit(f"ERROR: {method} {path} -> HTTP {e.code}: {e.read().decode('utf-8', 'replace')[:300]}")
        if not doc.get("success"):
            sys.exit(f"ERROR: {method} {path} -> {doc.get('errors')}")
        return doc["result"]

    api("GET", "/user/tokens/verify")
    accounts = api("GET", "/accounts?per_page=50")
    if len(accounts) != 1:
        sys.exit(f"ERROR: アカウントが {len(accounts)} 個あります。どれに作るかを確かめてから直してください")
    account_id = accounts[0]["id"]
    print(f"account: {accounts[0]['name']}")
    zone = api("GET", f"/zones?name={ZONE}")[0]
    print(f"zone: {zone['name']} ({zone['status']})")

    recs = api("GET", f"/zones/{zone['id']}/dns_records?name={urllib.parse.quote(HOSTNAME)}") or []
    tunnels = api("GET", f"/accounts/{account_id}/cfd_tunnel?name={urllib.parse.quote(TUNNEL_NAME)}&is_deleted=false") or []
    print(f"既存 DNS {HOSTNAME}: {[(r['type'], r['content']) for r in recs] or 'なし'}")
    print(f"既存 Tunnel {TUNNEL_NAME}: {'あり' if tunnels else 'なし'}")

    if tunnels:
        sys.exit(f"STOP: 同じ名前の Tunnel（{TUNNEL_NAME}）が既にあります。上書きしないので止めます")
    if recs:
        sys.exit(f"STOP: {HOSTNAME} の DNS が既にあります。ほかの用途の可能性があるので止めます")
    if dry:
        print("--dry-run: ここまで（作るものはありません: 問題なし）")
        return 0

    tunnel = api("POST", f"/accounts/{account_id}/cfd_tunnel", {"name": TUNNEL_NAME, "config_src": "cloudflare"})
    tunnel_id = tunnel["id"]
    print(f"tunnel: {TUNNEL_NAME} (id={tunnel_id[:8]}…)")
    api("PUT", f"/accounts/{account_id}/cfd_tunnel/{tunnel_id}/configurations",
        {"config": {"ingress": [{"hostname": HOSTNAME, "service": SERVICE_URL}, {"service": "http_status:404"}]}})
    print(f"ingress: {HOSTNAME} -> {SERVICE_URL}")
    api("POST", f"/zones/{zone['id']}/dns_records",
        {"type": "CNAME", "name": HOSTNAME, "content": f"{tunnel_id}.cfargotunnel.com", "proxied": True,
         "comment": DNS_COMMENT})
    print("dns: OK")
    tunnel_token = api("GET", f"/accounts/{account_id}/cfd_tunnel/{tunnel_id}/token")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(f"TUNNEL_TOKEN={tunnel_token}\n", encoding="utf-8")
    print(f"TUNNEL_TOKEN を書いた: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
