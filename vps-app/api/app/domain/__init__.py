"""業務（担当 W）。画面から呼ぶ関数と管理者の関数。

- 登録は `app/domain/register.py` の `register_all(register)`（rpc.load_domain_handlers が起動時に呼ぶ）。
  書き方は app/rpc.py の説明（register の write・role・request_id_arg）。このファイルが無いあいだは、
  閉じたリストは空（どの関数も 404）。
- 業務関数は commit しない（rpc が冪等キー・監査と一緒に確定する）。
- leave_requests の行を書いたら同じトランザクションで sheet_queue.mark_dirty(db, "leave_requests", req_id)、
  GAS の後処理（PDF・サイン画像・台帳・承認完了メール）は担当 M の services/gas_jobs で積む。
"""
