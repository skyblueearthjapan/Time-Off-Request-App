"""services/sheet_codec: GMT のシートで休暇日が前日 15:00（.625）のシリアル値・古い JST 9 時の整数の行も同じ日・
日時のミリ秒・cell_kinds・数に見える文字・見出しの別名・数式にならない文字。"""
from __future__ import annotations

from datetime import date, datetime, timezone

from app.services import sheet_codec as codec
from d_sheet_fixtures import HEADERS, LEAVE_ROWS

GMT = "Etc/GMT"


def test_leave_date_is_jst_midnight_as_previous_day_1500_in_gmt():
    enc = codec.encode_cell(codec.rule("LEAVE_DATE"), date(2026, 11, 27), None, GMT)
    serial = enc["cell"]["userEnteredValue"]["numberValue"]
    assert serial == 46352.625  # 06 5 節の実物と同じ（表示は 2026-11-26、休暇日は JST の 11-27）
    assert enc["kind"] == codec.SERIAL
    assert enc["cell"]["userEnteredFormat"]["numberFormat"]["type"] == "DATE"
    assert enc["fields"] == codec.WITH_FORMAT


def test_decode_both_shapes_of_old_and_new_leave_dates():
    r = codec.rule("LEAVE_DATE")
    assert codec.decode_cell(r, 46352.625, "2026-11-26", GMT) == (date(2026, 11, 27), codec.SERIAL)
    # 古い行: new Date('yyyy-MM-dd')＝UTC 0 時＝JST 9 時（整数）
    assert codec.decode_cell(r, 46352.0, "2026-11-26", GMT) == (date(2026, 11, 26), codec.SERIAL)
    assert codec.decode_cell(r, "2026-04-05", "2026-04-05", GMT) == ("2026-04-05", codec.STRING)


def test_datetime_keeps_milliseconds_and_wall_clock():
    moment = datetime(2026, 11, 26, 1, 12, 3, 456000, tzinfo=timezone.utc)  # 10:12:03.456 JST
    enc = codec.encode_cell(codec.rule("APPLIED_AT"), moment, None, GMT, {"type": "DATE_TIME", "pattern": "x"})
    serial = enc["cell"]["userEnteredValue"]["numberValue"]
    assert codec.serial_to_wall(serial) == datetime(2026, 11, 26, 1, 12, 3, 456000)
    assert enc["cell"]["userEnteredFormat"]["numberFormat"] == {"type": "DATE_TIME", "pattern": "x"}
    back, kind = codec.decode_cell(codec.rule("APPLIED_AT"), serial, "2026-11-26 01:12", GMT)
    assert back == moment and kind == codec.SERIAL


def test_blank_datetime_clears_value_only():
    enc = codec.encode_cell(codec.rule("APPROVED_AT2"), None, None, GMT)
    assert enc == {"cell": {}, "fields": "userEnteredValue", "kind": None}


def test_number_like_text_becomes_number_and_formulas_stay_text():
    r = codec.rule("WORKER_ID")
    assert codec.encode_cell(r, "0042", None, GMT)["cell"] == {"userEnteredValue": {"numberValue": 42}}
    assert codec.encode_cell(r, "0042", codec.STRING, GMT)["cell"] == {"userEnteredValue": {"stringValue": "0042"}}
    reason = codec.rule("SPECIAL_REASON_TEXT")
    for text in ("=SUM(A1)", "+1-2", "@x", "1/2"):
        assert codec.encode_cell(reason, text, None, GMT)["cell"] == {"userEnteredValue": {"stringValue": text}}
    fy = codec.encode_cell(codec.rule("FY_START_YEAR"), "2026", None, GMT)
    assert fy["cell"] == {"userEnteredValue": {"numberValue": 2026}} and fy["kind"] == codec.NUMBER


def test_same_value_compares_jst_dates_seconds_and_gas_reads():
    d = codec.rule("LEAVE_DATE")
    assert codec.same_value(d, date(2026, 3, 20), date(2026, 3, 20))
    assert codec.same_value(d, "2026-04-05", date(2026, 4, 5))
    t = codec.rule("APPROVED_AT")
    a = datetime(2026, 1, 1, 0, 0, 0, 300000, tzinfo=timezone.utc)
    assert codec.same_value(t, a, a.replace(microsecond=900000))
    assert not codec.same_value(t, a, a.replace(second=2))
    assert codec.same_value(codec.rule("WORKER_ID"), "42", "0042")


def test_header_index_takes_english_and_japanese_aliases():
    cols = codec.rule_columns(HEADERS)
    assert cols["REQ_ID"] == 0 and cols["APPROVED_AT2"] == 25 and len(cols) == 26
    jp = ["REQ_ID", "年度", "承認状態"] + HEADERS[3:]
    assert codec.rule_columns(jp)["FY_START_YEAR"] == 1 and codec.rule_columns(jp)["STATUS"] == 2
    assert codec.missing_headers(HEADERS[:-1]) == ["APPROVED_AT2"]
    assert codec.dropped_headers(HEADERS + ["PDF_FILE_ID", "メモ"]) == ["PDF_FILE_ID", "メモ"]


def test_decode_row_keeps_kinds_and_raw_serials():
    col = codec.rule_columns(HEADERS)
    row = LEAVE_ROWS[0]
    record, raw = codec.decode_row(col, row, [""] * len(row), GMT, 2)
    assert record["fiscal_year"] == "2025" and record["cell_kinds"]["fiscal_year"] == codec.NUMBER
    assert record["worker_id"] == "1001" and record["cell_kinds"]["worker_id"] == codec.NUMBER
    assert record["leave_date"] == date(2026, 3, 20) and record["cell_kinds"]["leave_date"] == codec.SERIAL
    assert raw["LEAVE_DATE"] == row[HEADERS.index("LEAVE_DATE")]  # 元のシリアル値（整数）を控える
    assert record["sheet_row"] == 2 and record["row_version"] == 1


def test_string_date_cell_is_preserved_until_the_date_changes():
    r = codec.rule("LEAVE_DATE")
    row = {"leave_date": date(2026, 4, 5), "leave_date_text": "2026/04/05", "cell_kinds": {"leave_date": "string"}}
    assert codec.db_value(r, row) == "2026/04/05"
    row["leave_date"] = date(2026, 4, 6)  # 編集で日付が変わった → 日付のセルで書く
    assert codec.db_value(r, row) == date(2026, 4, 6)


def test_pdf_file_id_from_url():
    assert codec.pdf_file_id_from_url("https://drive.google.com/file/d/AbC_1-x/view") == "AbC_1-x"
    assert codec.pdf_file_id_from_url("") == ""
