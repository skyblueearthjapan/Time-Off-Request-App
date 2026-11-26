"""JS と同じ結果になる小道具。"""
from __future__ import annotations

import pytest

from app.jsutil import js_floor, js_round, js_sort_key, js_string, normalize_gas


@pytest.mark.parametrize("x,expected", [
    (14.5, 15), (15.5, 16), (2.5, 3), (0.5, 1), (-0.5, 0), (-14.5, -14), (14.4999, 14), (0.0, 0), (15.5, 16),
    # 予測の .5（1 分を 2 日で割り 31 日を掛ける = 15.5 → 16）
    (1 / 2 * 31, 16),
])
def test_js_round(x, expected):
    assert js_round(x) == expected


def test_js_floor():
    assert js_floor(2.9) == 2 and js_floor(-0.1) == -1 and js_floor(3) == 3


@pytest.mark.parametrize("value,expected", [
    (None, ""), (True, "true"), (False, "false"), (1, "1"), (1.0, "1"), (1.5, "1.5"), (0.1, "0.1"),
    (-0.0, "0"), ("abc", "abc"), (1e21, "1e+21"),
])
def test_js_string(value, expected):
    assert js_string(value) == expected


@pytest.mark.parametrize("value,expected", [
    ("  製造１課  ", "製造1課"),
    ("ＡＢＣ１２３ａｂｃ", "ABC123abc"),
    ("　総務　", "総務"),
    ("﻿x\n", "x"),
    (None, ""),
    (12, "12"),
    (12.0, "12"),
    ("全角カナ（そのまま）", "全角カナ（そのまま）"),
    ("Ｇｏｏｇｌｅアカウント", "Googleアカウント"),
])
def test_normalize_gas(value, expected):
    assert normalize_gas(value) == expected


def test_js_sort_key_is_utf16_code_unit_order():
    words = ["総務", "製造1課", "設計", "Zeta", "alpha", "営業", "𠮷野家", "ｱ", "ア"]
    # JS: [...].sort() は UTF-16 の符号単位で比べる（サロゲートは 0xD800 台。Python の文字の順とは違う）
    expected = sorted(words, key=lambda s: [ord(c) if ord(c) < 0x10000 else 0xD800 for c in s])
    assert sorted(words, key=js_sort_key) == expected
    assert js_sort_key("ア") < js_sort_key("𠮷") < js_sort_key("ｱ")
    assert sorted(["𠮷", "ｱ"]) == ["ｱ", "𠮷"]  # Python の既定の並びは JS と逆になる例
    assert js_sort_key("B") < js_sort_key("a")


# ===== 仕様書 10-2・12-2 F の約束 =====


def test_spec_examples():
    from app.jsutil import sheet_readback

    assert normalize_gas("  ＡＢｃ１２３　") == "ABc123"
    assert js_string(2026.0) == "2026"
    assert js_string(2 ** 64 * 1.0) == "18446744073709552000"  # JS の String と同じ最短の桁
    assert sheet_readback("0042") == ("42", "number")


@pytest.mark.parametrize("text,expected", [
    ("1.50", ("1.5", "number")), ("+5", ("5", "number")), ("-3", ("-3", "number")), (".5", ("0.5", "number")),
    ("5.", ("5", "number")), ("1e5", ("1e5", "string")), (" 4", (" 4", "string")), ("４", ("４", "string")),
    ("2026-11-26", ("2026-11-26", "string")), ("=1+1", ("=1+1", "string")), ("", ("", "string")),
])
def test_sheet_readback(text, expected):
    from app.jsutil import sheet_readback

    assert sheet_readback(text) == expected


def test_collate_returns_sign_and_orders_like_locale_compare():
    import functools

    from app.jsutil import collate_default, collate_ja

    for fn in (collate_default, collate_ja):
        assert fn("a", "a") == 0 and fn("a", "b") == -1 and fn("b", "a") == 1
        # 大文字小文字は 1 次で同じ（JS の既定の sort と違い a < B）、記号 < 数字 < 文字、数字は桁ごと
        items = ["b", "B", "a", "A", "9", "10", "-x"]
        assert sorted(items, key=functools.cmp_to_key(fn)) == ["-x", "10", "9", "a", "A", "b", "B"]
    # 日付の文字（yyyy/MM/dd・MM/dd HH:mm）は桁がそろっているので文字の順と同じ
    dates = ["2026/12/01", "2026/01/15", "2025/12/31", "12/01 09:00", "01/02 10:30"]
    assert sorted(dates, key=functools.cmp_to_key(collate_default)) == sorted(dates)
