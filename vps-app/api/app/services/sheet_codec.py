"""DB の leave_requests ⇔ シート T_LEAVE_REQUEST のセル（実装仕様 5-1・6-1・6-2。担当 D）。

行の反映（row_sync）・毎晩の照合（reconcile）・移行の道具（tools/migrate_from_sheets.py）が同じ決まりを使う。

- 列はシートの見出しの文字で引く（COLUMN_RULES。GAS の buildHeaderIndex_ と同じく、英語の見出しと日本語の別名の
  どちらでも同じ列として受ける。見出しは normalize_ してから比べる）。列の位置は使わない。
- 区分（ColumnRule.cls）: 'V'＝VPS が書く、'G'＝GAS だけが書く（SIGN_IMAGE_URL・PDF_URL。行の反映は書かない）、
  'N'＝誰も書かない（SPECIAL_REASON_REF・DETAIL_FREE・UPDATED_AT。移行で値を持つだけ）。
- 書くときは updateCells のセル（CellData）を作る（encode_cell）。文字は stringValue（式にも日付にもならない）、
  数に見える文字は numberValue（GAS の setValues と同じ）、日付・日時は DB_SHEET_TIMEZONE の壁時計のシリアル値＋
  日付の表示形式。空は userEnteredValue を持たないセル（値を消す。表示形式は残る）。
- 読むときは UNFORMATTED_VALUE（SERIAL_NUMBER）と FORMATTED_VALUE を一緒に読む（decode_cell）。
- 行ごとの型は行の cell_kinds[属性名] に持つ（'serial'＝日付・日時のセル、'string'＝文字のセル、'number'＝数のセル）。
- 休暇日・振替元出勤日は「その日の JST 0 時」の瞬間で書く（GAS の new Date('yyyy-MM-ddT00:00:00+09:00')。
  GMT のシートでは前日 15:00＝小数 .625）。読むときはセルの瞬間の JST の日付にする（古い行の GMT 0:00＝JST 9:00 の
  整数のセルも同じ日になる）。変えていない列は書かないので、古い行の元のシリアル値はそのまま残る。

DB の列の名前は実装仕様 5-1（fiscal_year・leave_kubun ほか）。model_attr()・to_columns() は、持たない列（*_text など）
を捨てる小道具。移行のときのシートの生の値（休暇日の .625／整数など）は sheet_sync_state.synced_values の __raw__ に残す。
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from typing import Any
from zoneinfo import ZoneInfo

from app import jsutil
from app.jsutil import js_string, js_trim, normalize_gas

UTC = timezone.utc
JST = ZoneInfo("Asia/Tokyo")
EPOCH = datetime(1899, 12, 30)
MS_PER_DAY = 86_400_000

TAB = "T_LEAVE_REQUEST"
TABLE = "leave_requests"
HEADER_ROW = 1
DATA_START_ROW = 2

SERIAL = "serial"
STRING = "string"
NUMBER = "number"

# 新しく書く日付・日時のセルで、同じ列に写せる表示形式が無いとき（06 2-5 の表示）
DEFAULT_DATE_FORMAT = {"type": "DATE", "pattern": "yyyy-mm-dd"}
DEFAULT_DATETIME_FORMAT = {"type": "DATE_TIME", "pattern": "yyyy-mm-dd hh:mm"}
DATE_FORMAT_TYPES = ("DATE", "DATE_TIME", "TIME")

KINDS_KEY = "__kinds__"  # synced_values の中の「シートにある型」
RAW_KEY = "__raw__"  # synced_values の中の「移行のときのシートの生の値」（日付・日時の列だけ。元のシリアル値の控え）

# GAS の setValues が数にする文字（実装仕様 6-2）
NUMBER_TEXT_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)\Z")
REQ_ID_RE = re.compile(r"^LV-\d{8}-\d{3,}\Z")


class CodecError(ValueError):
    """セルの値が決まりの型として読めない。"""


@dataclass(frozen=True)
class ColumnRule:
    header: str  # 本番の英語の見出し
    attr: str  # 実装仕様 5-1 の DB の列の名前
    kind: str  # text・fy・date・datetime
    cls: str  # V・G・N
    alias: str = ""  # 日本語の見出し（HEADER_ALIAS_）

    @property
    def date_like(self) -> bool:
        return self.kind in ("date", "datetime")

    @property
    def text_attr(self) -> str:
        """日付・日時の列で、日付にならない元の文字を入れる列。"""
        return self.attr + "_text"


def _r(header: str, attr: str, kind: str, cls: str, alias: str = "") -> ColumnRule:
    return ColumnRule(header, attr, kind, cls, alias)


# 本番の見出し 26 列（06 2-5。並びは本番の並び。列は見出しの文字で引くので、並びは書き方に使わない）
COLUMN_RULES: tuple[ColumnRule, ...] = (
    _r("REQ_ID", "req_id", "text", "V"),
    _r("FY_START_YEAR", "fiscal_year", "fy", "V", "年度"),
    _r("STATUS", "status", "text", "V", "承認状態"),
    _r("APPLIED_AT", "applied_at", "datetime", "V", "申請日時"),
    _r("DEPT_ID", "dept_id", "text", "V", "部署ID"),
    _r("DEPT_NAME", "dept_name", "text", "V", "部署名"),
    _r("WORKER_ID", "worker_id", "text", "V", "作業員ID"),
    _r("WORKER_NAME", "worker_name", "text", "V", "作業員名"),
    _r("LEAVE_DATE", "leave_date", "date", "V", "休暇日"),
    _r("DAY_TYPE", "leave_kubun", "text", "V", "休暇区分"),
    _r("HALF_TYPE", "half_day_type", "text", "V", "半日区分"),
    _r("LEAVE_KIND", "leave_type", "text", "V", "休暇種類"),
    _r("SUBSTITUTE_FOR_WORK_DATE", "substitute_date", "date", "V", "振替元出勤日"),
    _r("SPECIAL_REASON_REF", "special_reason_ref", "text", "N", "特別理由参照"),
    _r("SPECIAL_REASON_TEXT", "special_reason", "text", "V", "特別理由"),
    _r("PAIDLEAVE_DETAIL", "paid_detail", "text", "V", "有給詳細"),
    _r("DETAIL_FREE", "additional_detail", "text", "N", "追加詳細"),
    _r("APPROVED_AT", "approved_at", "datetime", "V", "承認日時"),
    _r("SIGN_IMAGE_URL", "sign_image_url", "text", "G", "サイン画像URL"),
    _r("PDF_URL", "pdf_url", "text", "G"),
    _r("CREATED_BY_EMAIL", "created_by_email", "text", "V", "作成者メール"),
    _r("UPDATED_AT", "updated_at_text", "text", "N", "更新日時"),
    _r("APPROVED_BY1_EMAIL", "approved_by1_email", "text", "V", "1次承認者メール"),
    _r("APPROVED_BY2", "approved_by2", "text", "V", "2次承認者名"),
    _r("APPROVED_BY2_EMAIL", "approved_by2_email", "text", "V", "2次承認者メール"),
    _r("APPROVED_AT2", "approved_at2", "datetime", "V", "2次承認日時"),
)
HEADERS: tuple[str, ...] = tuple(c.header for c in COLUMN_RULES)
V_RULES = tuple(c for c in COLUMN_RULES if c.cls == "V")
G_RULES = tuple(c for c in COLUMN_RULES if c.cls == "G")
N_RULES = tuple(c for c in COLUMN_RULES if c.cls == "N")
ID_RULE = COLUMN_RULES[0]
# 将来 PDF_FILE_ID 列が足されたら G 区分（GAS が書く）。今の本番には無い（06 2-5）
OPTIONAL_G_HEADERS = ("PDF_FILE_ID",)
# 型を行ごとに持つ列（cell_kinds のキー＝属性名）
KIND_ATTRS = ("fiscal_year", "worker_id", "applied_at", "leave_date", "substitute_date", "approved_at", "approved_at2")


def rule(header: str) -> ColumnRule:
    for c in COLUMN_RULES:
        if c.header == header:
            return c
    raise KeyError(header)


# ===== 土台の models.py との読み替え（土台が仕様 5-1 の名前になれば、1 つ目の候補がそのまま使われる） =====

# 仕様の名前 → 実物の列の候補（土台は仕様 5-1 の名前になった＝master 5c4431e。読み替えは要らないので空）
LEAVE_ATTR_CANDIDATES: dict[str, tuple[str, ...]] = {}


def _leave_columns() -> dict[str, Any]:
    from app.models import LeaveRequest

    return {c.name: c for c in LeaveRequest.__table__.columns}


def model_attr(name: str) -> str | None:
    """仕様の属性名 → 実物の列の名前（無ければ None＝持たない）。"""
    cols = _leave_columns()
    for cand in LEAVE_ATTR_CANDIDATES.get(name, (name,)):
        if cand in cols:
            return cand
    return None


def _is_int_column(column_name: str) -> bool:
    from sqlalchemy import Integer

    col = _leave_columns().get(column_name)
    return col is not None and isinstance(col.type, Integer)


def get_attr(row: Any, name: str, default: Any = None) -> Any:
    """ORM の行か dict（列の名前がキー）から、仕様の属性名で値を取る。"""
    real = model_attr(name)
    if real is None:
        return default
    if isinstance(row, dict):
        return row.get(real, default)
    return getattr(row, real, default)


def to_columns(record: dict) -> dict:
    """仕様の属性名の dict → 実物の列の名前の dict（持たない列は捨てる。年度が整数の列なら整数にする）。"""
    out: dict = {}
    cols = _leave_columns()
    for name, value in record.items():
        real = model_attr(name) if name in LEAVE_ATTR_CANDIDATES or name not in cols else name
        if real is None:
            continue
        if _is_int_column(real) and not isinstance(value, int):
            text = js_trim(js_string(value))
            value = int(float(text)) if NUMBER_TEXT_RE.match(text) else None
        out[real] = value
    return out


def set_attr(row: Any, name: str, value: Any) -> bool:
    """ORM の行に仕様の属性名で値を入れる（持たない列なら False）。"""
    converted = to_columns({name: value})
    if not converted:
        return False
    for k, v in converted.items():
        setattr(row, k, v)
    return True


# ===== シリアル値 =====


def wall_to_serial(wall: datetime) -> float:
    delta = wall.replace(tzinfo=None) - EPOCH
    ms = delta.days * MS_PER_DAY + delta.seconds * 1000 + delta.microseconds // 1000
    return ms / MS_PER_DAY


def serial_to_wall(serial: float) -> datetime:
    ms = int(math.floor(float(serial) * MS_PER_DAY + 0.5))
    return EPOCH + timedelta(milliseconds=ms)


def as_tz(tz: Any) -> ZoneInfo:
    return tz if isinstance(tz, ZoneInfo) else ZoneInfo(str(tz))


def datetime_to_serial(moment: datetime, tz: Any) -> float:
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=UTC)
    return wall_to_serial(moment.astimezone(as_tz(tz)).replace(tzinfo=None))


def serial_to_datetime(serial: float, tz: Any) -> datetime:
    return serial_to_wall(serial).replace(tzinfo=as_tz(tz)).astimezone(UTC)


def jst_midnight(d: date) -> datetime:
    return datetime(d.year, d.month, d.day, tzinfo=JST)


def date_to_serial(d: date, tz: Any) -> float:
    """休暇日 → 「その日の JST 0 時」の瞬間のシート（tz）の壁時計のシリアル値（GMT では前日の .625）。"""
    return datetime_to_serial(jst_midnight(d), tz)


def serial_to_jst_date(serial: float, tz: Any) -> date:
    """日付のセルのシリアル値 → その瞬間の JST の日付（.625 の行も整数＝JST 9 時の行も同じ日）。"""
    return serial_to_datetime(serial, tz).astimezone(JST).date()


# ===== 文字の日時 =====

_DT_FORMATS = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M",
               "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%Y-%m-%dT%H:%M:%S")
_D_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y")


def parse_jst_datetime(text: str) -> datetime | None:
    s = js_trim(text or "")
    if not s:
        return None
    if "T" in s and re.search(r"(Z|[+-]\d{2}:?\d{2})\Z", s):
        try:
            return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(UTC)
        except ValueError:
            return None
    for fmt in _DT_FORMATS + _D_FORMATS:
        try:
            return datetime.strptime(s, fmt).replace(tzinfo=JST).astimezone(UTC)
        except ValueError:
            continue
    return None


def parse_date_text(text: str) -> date | None:
    s = js_trim(text or "")
    for fmt in _D_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def jst_text(moment: datetime | None) -> str:
    if moment is None:
        return ""
    return moment.astimezone(JST).strftime("%Y-%m-%d %H:%M:%S")


_DATE_DISPLAY_RE = re.compile(r"^\s*(\d{1,4}[/\-.年]\d{1,2}[/\-.月]\d{1,4}|\d{1,2}:\d{2})")


def looks_like_date_display(formatted: Any) -> bool:
    return isinstance(formatted, str) and bool(_DATE_DISPLAY_RE.match(formatted))


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def is_blank(value: Any) -> bool:
    return value is None or value == ""


def sheet_readback(text: Any) -> tuple[str, str]:
    """GAS の setValues で書いた文字を読み直した形（数に見える文字は数）。土台の jsutil.sheet_readback があればそれ。"""
    fn = getattr(jsutil, "sheet_readback", None)
    if fn is not None:
        return fn(text)
    s = js_string(text)
    if NUMBER_TEXT_RE.match(s):
        n = float(s)
        return js_string(int(n) if n == int(n) else n), NUMBER
    return s, STRING


def number_value(text: str) -> int | float:
    n = float(text)
    return int(n) if n == int(n) and abs(n) < 1e15 else n


# ===== DB の行 ⇔ 論理の値 =====


def _kinds_of(row: Any) -> dict:
    kinds = row.get("cell_kinds") if isinstance(row, dict) else getattr(row, "cell_kinds", None)
    return dict(kinds or {})


def db_value(r: ColumnRule, row: Any) -> Any:
    """ORM の行（か列の名前の dict）からその列の論理の値。

    text・fy: 文字（年度が整数の列なら String(値)）。date: date（文字のセルの古い行で、元の文字が同じ日を指すなら元の文字）
    か日付にならない文字。datetime: datetime（文字のセルの行は元の文字）か ''。
    """
    kinds = _kinds_of(row)
    if r.kind in ("text", "fy"):
        v = get_attr(row, r.attr)
        return "" if v is None else js_string(v)
    value = get_attr(row, r.attr)
    text = get_attr(row, r.text_attr) or ""
    if kinds.get(r.attr) == STRING and text:
        if value is None:
            return text
        if r.kind == "date" and parse_date_text(text) == value:
            return text
        if r.kind == "datetime" and (parse_jst_datetime(text) is not None
                                     and abs((parse_jst_datetime(text) - _aware(value)).total_seconds()) < 1.0):
            return text
    if value is not None:
        return value
    return text


def _aware(v: datetime) -> datetime:
    return v if v.tzinfo else v.replace(tzinfo=UTC)


def db_values(row: Any, rules: tuple[ColumnRule, ...] = COLUMN_RULES) -> dict:
    """見出し → 論理の値。"""
    return {r.header: db_value(r, row) for r in rules}


def assign(record: dict, r: ColumnRule, value: Any) -> None:
    """論理の値を、仕様の属性名の dict に入れる（移行）。"""
    if r.kind == "date":
        if isinstance(value, date):
            record[r.attr], record[r.text_attr] = value, ""
        else:
            record[r.attr], record[r.text_attr] = parse_date_text(js_string(value)), js_string(value or "")
        return
    if r.kind == "datetime":
        if isinstance(value, datetime):
            record[r.attr], record[r.text_attr] = value, ""
        else:
            text = js_string(value or "")
            record[r.attr], record[r.text_attr] = parse_jst_datetime(text), text
        return
    record[r.attr] = "" if value is None else value


def canonical(r: ColumnRule, value: Any) -> Any:
    """synced_values に入れる・比べる形（JSON にできる値）。"""
    if value is None or value == "":
        return ""
    if r.kind == "datetime" and isinstance(value, datetime):
        return _aware(value).astimezone(UTC).isoformat(timespec="milliseconds")
    if r.kind == "date" and isinstance(value, date):
        return value.isoformat()
    return js_string(value)


def canonical_stored(r: ColumnRule, stored: Any) -> Any:
    """synced_values の値を canonical の形にそろえる（ISO の書き方の揺れを吸収）。"""
    if r.kind == "datetime" and isinstance(stored, str) and "T" in stored:
        try:
            parsed = datetime.fromisoformat(stored.replace("Z", "+00:00"))
        except ValueError:
            return stored
        return canonical(r, _aware(parsed))
    return stored


def _as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return _aware(value).astimezone(JST).date()
    if isinstance(value, date):
        return value
    return parse_date_text(js_string(value)) if not is_blank(value) else None


def same_value(r: ColumnRule, a: Any, b: Any) -> bool:
    """シートから読んだ値 a と DB の値 b が同じか。日付は JST の日付、日時は秒、文字は GAS が読む形。"""
    if is_blank(a) and is_blank(b):
        return True
    if is_blank(a) != is_blank(b):
        return False
    if r.kind == "date":
        da, db_ = _as_date(a), _as_date(b)
        if da is not None and db_ is not None:
            return da == db_
        return js_string(a) == js_string(b)
    if r.kind == "datetime":
        ma = a if isinstance(a, datetime) else parse_jst_datetime(js_string(a))
        mb = b if isinstance(b, datetime) else parse_jst_datetime(js_string(b))
        if ma is not None and mb is not None:
            return abs((_aware(ma) - _aware(mb)).total_seconds()) < 1.0
        return js_string(a) == js_string(b)
    return normalize_gas(sheet_readback(a)[0]) == normalize_gas(sheet_readback(b)[0])


def snapshot(row: Any) -> dict:
    """行の全列の canonical（見出し → 値）。"""
    return {r.header: canonical(r, db_value(r, row)) for r in COLUMN_RULES}


def synced_values_for(record: Any, raw: dict | None = None) -> dict:
    """シートから読んだ行の sheet_sync_state.synced_values（今シートにある値と型。反映済みの印）。"""
    synced = snapshot(record)
    synced[KINDS_KEY] = _kinds_of(record)
    if raw:
        synced[RAW_KEY] = dict(raw)
    return synced


# ===== 書く（CellData） =====


def _string_cell(text: str) -> dict:
    return {"userEnteredValue": {"stringValue": text}} if text != "" else {}


def _number_cell(number: float) -> dict:
    return {"userEnteredValue": {"numberValue": number}}


VALUE_ONLY = "userEnteredValue"
WITH_FORMAT = "userEnteredValue,userEnteredFormat.numberFormat"


def encode_cell(r: ColumnRule, value: Any, kind: str | None, tz: Any, number_format: dict | None = None) -> dict:
    """論理の値 → {"cell": CellData, "fields", "kind": 書いた型（空なら None）}。

    kind は行の cell_kinds の型（無ければ None）。文字の列で kind が 'string' のときだけ、数に見える文字も文字のまま
    書く（シートの元の型を保つ）。日付の列で kind が 'string' なら文字で書く（呼ぶ側が db_value で元の文字を渡す）。
    """
    if is_blank(value):
        return {"cell": {}, "fields": VALUE_ONLY, "kind": None}
    if r.kind in ("text", "fy"):
        text = js_string(value)
        if kind != STRING and NUMBER_TEXT_RE.match(text):
            return {"cell": _number_cell(number_value(text)), "fields": VALUE_ONLY, "kind": NUMBER}
        return {"cell": _string_cell(text), "fields": VALUE_ONLY, "kind": STRING}
    fmt = dict(number_format or (DEFAULT_DATE_FORMAT if r.kind == "date" else DEFAULT_DATETIME_FORMAT))
    if r.kind == "date":
        if isinstance(value, date) and not isinstance(value, datetime):
            cell = _number_cell(date_to_serial(value, tz))
            cell["userEnteredFormat"] = {"numberFormat": fmt}
            return {"cell": cell, "fields": WITH_FORMAT, "kind": SERIAL}
        return {"cell": _string_cell(js_string(value)), "fields": VALUE_ONLY, "kind": STRING}
    if r.kind == "datetime":
        if isinstance(value, datetime):
            cell = _number_cell(datetime_to_serial(value, tz))
            cell["userEnteredFormat"] = {"numberFormat": fmt}
            return {"cell": cell, "fields": WITH_FORMAT, "kind": SERIAL}
        return {"cell": _string_cell(js_string(value)), "fields": VALUE_ONLY, "kind": STRING}
    raise ValueError(f"知らない列の種類です: {r.kind}")  # pragma: no cover


# ===== 読む =====


def decode_cell(r: ColumnRule, unformatted: Any, formatted: Any, tz: Any) -> tuple[Any, str | None]:
    """UNFORMATTED と FORMATTED の値 → (論理の値, 型)。空は ('' か None, None)。"""
    v = unformatted
    if is_blank(v):
        return (None if r.kind == "datetime" else ""), None
    if r.kind in ("text", "fy"):
        if isinstance(v, bool):
            return ("true" if v else "false"), STRING
        if is_number(v):
            if r.kind == "text" and looks_like_date_display(formatted):
                return js_string(formatted), SERIAL  # 自由入力が日付のセルになった行（表示の文字で持つ）
            return js_string(v), NUMBER
        return js_string(v), STRING
    if r.kind == "date":
        if is_number(v):
            return serial_to_jst_date(v, tz), SERIAL
        return js_string(v), STRING
    if r.kind == "datetime":
        if is_number(v):
            return serial_to_datetime(v, tz), SERIAL
        return js_string(v), STRING
    raise ValueError(f"知らない列の種類です: {r.kind}")  # pragma: no cover


def decode_row(col: dict[str, int], urow: list, frow: list, tz: Any, sheet_row: int) -> tuple[dict, dict]:
    """シートの 1 行 → (仕様の属性名の dict（sheet_row・row_version=1・cell_kinds を含む）, 日付・日時の生の値)。"""
    record: dict = {"sheet_row": sheet_row, "row_version": 1, "cell_kinds": {}}
    raw: dict = {}
    for r in COLUMN_RULES:
        c = col.get(r.header, -1)
        u, f = cell_at(urow, c), cell_at(frow, c)
        value, kind = decode_cell(r, u, f, tz)
        assign(record, r, value)
        if kind and (r.attr in KIND_ATTRS or (r.kind == "text" and (kind == NUMBER or (
                kind == STRING and NUMBER_TEXT_RE.match(js_string(value)))))):
            record["cell_kinds"][r.attr] = kind
        if r.date_like and not is_blank(u):
            raw[r.header] = u
    return record, raw


def pdf_file_id_from_url(url: Any) -> str:
    m = re.search(r"/d/([A-Za-z0-9_-]+)", js_string(url))
    return m.group(1) if m else ""


# ===== 見出し =====


def _alias_maps() -> tuple[dict[str, str], dict[str, str]]:
    forward = {r.header: r.alias for r in COLUMN_RULES if r.alias}
    return forward, {v: k for k, v in forward.items()}


def header_index(header_row: list) -> dict[str, int]:
    """GAS の getSheetHeaderIndex_＋buildHeaderIndex_ と同じ（normalize_ した見出し → 列。英語・日本語の別名も登録）。

    戻り値のキーは見出しの文字そのもの（英語・日本語・決まりに無いもの）。COLUMN_RULES の列は英語の見出しで引く。
    """
    forward, reverse = _alias_maps()
    idx: dict[str, int] = {}
    last = 0
    for i, h in enumerate(header_row or []):
        if js_trim(js_string(h)) != "":
            last = i + 1
    for i in range(last):
        h = normalize_gas(header_row[i])
        idx[h] = i
        if h in forward and forward[h] not in idx:
            idx[forward[h]] = i
        if h in reverse and reverse[h] not in idx:
            idx[reverse[h]] = i
    return idx


def rule_columns(header_row: list) -> dict[str, int]:
    """COLUMN_RULES の見出し → 0 始まりの列（無い見出しは入らない）。"""
    idx = header_index(header_row)
    return {r.header: idx[r.header] for r in COLUMN_RULES if r.header in idx}


def missing_headers(header_row: list) -> list[str]:
    cols = rule_columns(header_row)
    return [r.header for r in COLUMN_RULES if r.header not in cols]


def dropped_headers(header_row: list) -> list[str]:
    """シートにあって決まりに無い見出し（触らない）。"""
    known = set(HEADERS) | {r.alias for r in COLUMN_RULES if r.alias}
    return [normalize_gas(h) for h in header_row if normalize_gas(h) and normalize_gas(h) not in known]


def cell_at(row: list, index: int) -> Any:
    return row[index] if 0 <= index < len(row) else ""


__all__ = [
    "COLUMN_RULES", "ColumnRule", "decode_cell", "decode_row", "encode_cell", "synced_values_for", "header_index",
    "rule_columns", "missing_headers", "db_value", "db_values", "same_value", "canonical", "model_attr",
]
