"""癖の切り替え（実装仕様 15 章・06 の決定）。オーナーの答えで変えるときは、ここの定数 1 つを直す。

担当 M・D は読むだけ（PDF_CLEANUP_MODE・LAST_PAID_DATE_STYLE）。
"""
from __future__ import annotations

# Q6: 前の PDF を捨てない（本番に PDF_FILE_ID 列が無く、今の GAS は捨てていない）。'trash' にすると、新しい PDF が
# できてから前の PDF を捨てる（GAS の内部用の入口が行う）。編集・取消で DB の pdf_url も残す
PDF_CLEANUP_MODE = "asis"

# Q7（K5）: 画面へ返す「最終有給日」の形。
# 'gas' = 今の GAS が返している文字。集計タブのセル（USER_ENTERED の 'yyyy-MM-dd' を、シートが DB_SHEET_TIMEZONE の
#         0 時の日付のセルに変える）を GAS が Date で読み、String(Date) にした長い英語の形
#         （GMT のシートで 'Fri Nov 27 2026 09:00:00 GMT+0900 (Japan Standard Time)'）。
# 'iso'（'ymd' も同じ）= 'yyyy-MM-dd' のまま。
LAST_PAID_DATE_STYLE = "iso"

# Q2（ディレクター判断で直す）: 2 次承認者のメールは引数ではなくログイン中のメールを使う。False なら今の GAS と同じく引数。
SECOND_APPROVER_FROM_LOGIN = True

# K-10（オーナーの決定で直す）: 承認待ち一覧を本当の申請日時の新しい順に並べる（表示の形 MM/dd HH:mm は今のまま）。
# False なら今の GAS と同じく 'MM/dd HH:mm' の文字の降順（年をまたぐと 12 月が 1 月より上）。
PENDING_SORT_WITH_YEAR = True

# 手書きサインの data URL の上限（文字数。nginx は 4m、rpc の本文は 3 MB。受け入れる差）
SIGN_MAX_CHARS = 3_000_000
# 手書きサインの PNG の上限（base64 を解いた後のバイト数）。今の画面の PNG は devicePixelRatio 3 のタブレットでも数十 KB
SIGN_MAX_BYTES = 1_500_000

# シートのセルの上限（GAS も書き込みで失敗する。受け入れる差）
MAX_CELL_CHARS = 50_000
