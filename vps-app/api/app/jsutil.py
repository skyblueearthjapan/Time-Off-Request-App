"""GAS（JavaScript）と同じ結果になる小道具。

- js_round(x) = Math.round（.5 は上へ。負の数も JS と同じ）。Python の組み込みの丸め（偶数への丸め）は使わない。
- js_floor(x) = Math.floor。
- js_string(v) = String(v)（1.0 → "1"、true → "true"、null/undefined → ""）。
- normalize_gas(s) = src/config.gs の normalize_（前後の空白を取り、全角英数を半角に）。
- js_sort_key(s) = JS の文字の比べ（UTF-16 の符号単位の順）と同じ並びになるキー。
- collate_default(a, b)・collate_ja(a, b) = localeCompare（言語なし・'ja'）。
- sheet_readback(text) = シートへ書いた文字を読み直した値と型（数に見える文字は数）。
"""
from __future__ import annotations

import math
import re
import unicodedata
from decimal import Decimal
from typing import Any

# JS の String.prototype.trim が取る空白（WhiteSpace と LineTerminator）
_JS_WHITESPACE = (
    "\u0009\u000a\u000b\u000c\u000d          "
    "        　﻿"
)
# JS の正規表現の \s と同じ文字の組（Python の \s は U+001C・U+0085 を含み U+FEFF を含まない）
JS_WHITESPACE_RE = re.compile("[" + re.escape(_JS_WHITESPACE) + "]+")
_FULLWIDTH_ALNUM = re.compile(r"[Ａ-Ｚａ-ｚ０-９]")


def js_round(x: float) -> int:
    """Math.round と同じ（floor(x + 0.5)）。14.5 → 15、-14.5 → -14。"""
    return int(math.floor(x + 0.5))


def js_floor(x: float) -> int:
    return int(math.floor(x))


def _js_number_string(value: float) -> str:
    if math.isnan(value):
        return "NaN"
    if math.isinf(value):
        return "Infinity" if value > 0 else "-Infinity"
    if value == int(value) and abs(value) < 1e21:
        # JS は最短の桁で書いてから 0 を足す（2**64 → "18446744073709552000"）。repr も最短の桁
        return format(Decimal(repr(value)), "f") if "e" in repr(value) else str(int(value))
    text = repr(value)
    if "e" in text:
        mantissa, exp = text.split("e")
        exp_num = int(exp)
        return f"{mantissa}e{'+' if exp_num > 0 else '-'}{abs(exp_num)}"
    return text


def js_string(value: Any) -> str:
    """String(value) と同じ文字（null・undefined にあたる None は ''。GAS の `String(x ?? '')` の形）。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return _js_number_string(value)
    if isinstance(value, (list, tuple)):
        # Array.prototype.toString: 要素を ',' でつなぐ（null・undefined は空。正しさ 低-2）
        return ",".join(js_string(v) for v in value)
    if isinstance(value, dict):
        return "[object Object]"
    return str(value)


def js_property_key(value: Any) -> str:
    """オブジェクトのキーにしたときの文字（`obj[value]`）。null は 'null'（String(null)）。"""
    if value is None:
        return "null"
    if isinstance(value, (list, tuple)):
        return ",".join("" if v is None else js_string(v) for v in value)
    return js_string(value)


def _is_array_index(key: str) -> bool:
    return key.isdigit() and key.isascii() and (key == "0" or not key.startswith("0")) and int(key) < 2**32 - 1


def js_object_key_order(keys: list[str]) -> list[str]:
    """`for (k in obj)` の順: 配列の添字に見える数字のキーを小さい順に先に、残りは入れた順。"""
    index_keys = sorted((k for k in keys if _is_array_index(k)), key=int)
    return index_keys + [k for k in keys if not _is_array_index(k)]


def js_utf16_units(text: str) -> list[str]:
    """文字列を JS の添字（UTF-16 の符号単位）ごとに分ける（サロゲートのペアは 2 つになる）。"""
    data = text.encode("utf-16-be", "surrogatepass")
    return [data[i:i + 2].decode("utf-16-be", "surrogatepass") for i in range(0, len(data), 2)]


def js_array_like_items(value: Any) -> list:
    """`for (k = 0; k < x.length; k++) x[k]` で読む値（配列はそのまま、文字列は 1 字ずつ、ほかは length が無いので空）。"""
    if isinstance(value, (list, tuple)):
        return list(value)
    if isinstance(value, str):
        return js_utf16_units(value)
    return []


def js_trim(text: str) -> str:
    return text.strip(_JS_WHITESPACE)


def normalize_gas(value: Any) -> str:
    """src/config.gs の normalize_ と同じ: String(s ?? '').trim() の後、全角英数（Ａ-Ｚａ-ｚ０-９）を半角に。"""
    text = js_trim(js_string(value))
    return _FULLWIDTH_ALNUM.sub(lambda m: chr(ord(m.group(0)) - 0xFEE0), text)


def js_sort_key(value: Any) -> tuple[int, ...]:
    """JS の既定の並び（Array.prototype.sort・`<`）と同じ順になるキー（UTF-16 の符号単位の並び）。"""
    data = js_string(value).encode("utf-16-be", "surrogatepass")
    return tuple(int.from_bytes(data[i:i + 2], "big") for i in range(0, len(data), 2))


# ===== シートへ書いた文字の読み戻し（仕様書 6-2: GAS の setValues と同じく、数に見える文字は数） =====

_NUMBER_LIKE = re.compile(r"^[+-]?([0-9]+\.?[0-9]*|\.[0-9]+)\Z")


def sheet_readback(text: Any) -> tuple[str, str]:
    """シートへ書いた文字が、読み直したときにどうなるか。(DB に入れる文字, cell_kinds の型)。

    `^[+-]?(\\d+\\.?\\d*|\\.\\d+)$` に合う文字は数のセルになる: (String(Number(文字)), 'number')。
    例 '0042' → ('42', 'number')、'1.50' → ('1.5', 'number')。ほかは (文字, 'string')。
    """
    s = js_string(text)
    if _NUMBER_LIKE.match(s):
        return js_string(float(s)), "number"
    return s, "string"


# ===== localeCompare（仕様書 8-10） =====
# collate_default(a, b) = a.localeCompare(b)（言語の指定なし＝ICU の既定の照合。en_US とみなす）
# collate_ja(a, b)      = a.localeCompare(b, 'ja')
# 戻り値は -1・0・1。並べ替えは functools.cmp_to_key(collate_default) で使う。
# PyICU（V8 と同じ ICU の照合）を使う。本番のイメージには入れ、APP_ENV=production で読み込めなければ起動しない
# （require_icu）。手元の試験など PyICU が無い環境では近い並び（_fallback_key）で代える（記号 < 数字 < ラテン文字
# （大小を区別しない、同じなら小文字が先）< ほかの文字 < かな（ひらがなとカタカナは同じ、同じならひらがなが先）
# < 漢字（符号順。ICU の ja は JIS X 0208 の順＝おおむね読みの順なので、ここが ICU と違う））。


def _icu_module():
    try:
        import icu  # type: ignore[import-not-found]
    except ImportError:
        return None
    return icu


def icu_version() -> str | None:
    """読み込めた ICU の版（無ければ None）。/healthz?detail=1 に出す。"""
    icu = _icu_module()
    return None if icu is None else str(icu.ICU_VERSION)


def require_icu(app_env: str) -> None:
    """APP_ENV=production で PyICU が読み込めなければ起動を止める（並びが GAS と変わるのを防ぐ）。"""
    if app_env == "production" and icu_version() is None:
        raise RuntimeError("PyICU（ICU の照合）が読み込めません。本番は ICU なしでは起動しません")


def _icu_collator(locale: str):
    icu = _icu_module()
    if icu is None:
        return None
    return icu.Collator.createInstance(icu.Locale(locale))


_COLLATORS: dict[str, Any] = {}


def _collator(locale: str):
    if locale not in _COLLATORS:
        _COLLATORS[locale] = _icu_collator(locale)
    return _COLLATORS[locale]


def _char_class(ch: str) -> tuple[int, int, int]:
    """(種類, 1 次の値, 3 次の値) を返す（PyICU が無いときの近い並び）。"""
    cp = ord(ch)
    if ch.isspace():
        return 0, cp, 0
    if "0" <= ch <= "9":
        return 2, cp, 0
    if ch.isascii() and ch.isalpha():
        return 3, ord(ch.lower()), 0 if ch.islower() else 1
    if 0x3041 <= cp <= 0x3096 or 0x30A1 <= cp <= 0x30F6:  # ひらがな・カタカナ
        hira = cp - 0x60 if cp >= 0x30A1 else cp
        base = unicodedata.normalize("NFD", chr(hira))[0]  # 濁点を外した字（1 次は同じ、2 次で分ける）
        return 5, ord(base), (0 if cp < 0x30A1 else 1)
    if 0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or 0xF900 <= cp <= 0xFAFF:
        return 6, cp, 0
    category = unicodedata.category(ch)
    if category[0] in ("P", "S"):
        return 1, cp, 0
    return 4, cp, 0


def _fallback_key(text: str) -> tuple:
    normalized = unicodedata.normalize("NFKC", text)
    classes = [_char_class(ch) for ch in normalized]
    primary = tuple((c[0], c[1]) for c in classes)
    secondary = tuple(ord(ch) for ch in normalized)
    tertiary = tuple(c[2] for c in classes)
    return primary, tertiary, secondary, js_sort_key(text)


def _compare(locale: str, a: Any, b: Any) -> int:
    sa, sb = js_string(a), js_string(b)
    collator = _collator(locale)
    if collator is not None:
        result = collator.compare(sa, sb)
        return (result > 0) - (result < 0)
    ka, kb = _fallback_key(sa), _fallback_key(sb)
    return (ka > kb) - (ka < kb)


def collate_default(a: Any, b: Any) -> int:
    """a.localeCompare(b)（言語の指定なし）。-1・0・1。"""
    return _compare("en_US", a, b)


def collate_ja(a: Any, b: Any) -> int:
    """a.localeCompare(b, 'ja')。-1・0・1。"""
    return _compare("ja", a, b)
