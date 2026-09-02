# -*- coding: utf-8 -*-
"""健康診断申込: シート構成の検証（列名・スキーマ版・年度・パディング）。"""

import pytest

from services.health_apply import schema as S


def good_headers():
    return {name: list(S.HEADERS_BY_SHEET[name]) for name in S.ALL_SHEETS}


def good_settings():
    return {
        "スキーマ版": S.SCHEMA_VERSION, "年度": "2027", "前年度": "2026",
        "受付開始": "2027-02-01", "受付終了": "2027-02-28",
        "受診期間開始": "2027-04-01", "受診期間終了": "2028-03-31", "回答受付": "1",
    }


def test_verify_workbook_passes_with_exact_layout():
    assert S.verify_workbook(list(S.ALL_SHEETS), good_headers(), good_settings(), 2027) == []


def test_extra_trailing_columns_are_allowed():
    headers = good_headers()
    headers[S.SHEET_TARGETS] = headers[S.SHEET_TARGETS] + ["メモ", "担当"]
    assert S.verify_workbook(list(S.ALL_SHEETS), headers, good_settings(), 2027) == []


def test_missing_sheet_is_reported():
    titles = [n for n in S.ALL_SHEETS if n != S.SHEET_AUDIT]
    errors = S.verify_workbook(titles, good_headers(), good_settings(), 2027)
    assert errors == ["シートがありません: 監査ログ"]


def test_header_mismatch_names_column_and_expected():
    headers = good_headers()
    headers[S.SHEET_TARGETS][2] = "名前"
    errors = S.verify_workbook(list(S.ALL_SHEETS), headers, good_settings(), 2027)
    assert errors == ["対象者シートの3列目が想定外です: 「名前」（期待: 「氏名」）"]


def test_short_header_lists_missing_columns():
    headers = good_headers()
    headers[S.SHEET_AUDIT] = headers[S.SHEET_AUDIT][:5]
    errors = S.verify_workbook(list(S.ALL_SHEETS), headers, good_settings(), 2027)
    assert errors == ["監査ログシートに列がありません: 社員番号、詳細"]


def test_header_cells_are_stripped_and_none_tolerated():
    headers = good_headers()
    headers[S.SHEET_SETTINGS] = [" キー ", "値", None]
    errors = S.verify_workbook(list(S.ALL_SHEETS), headers, good_settings(), 2027)
    assert errors == ["設定シートに列がありません: 備考"]


def test_schema_version_mismatch():
    settings = good_settings()
    settings["スキーマ版"] = "2026.9"
    errors = S.verify_workbook(list(S.ALL_SHEETS), good_headers(), settings, 2027)
    assert errors == [f"設定シートのスキーマ版が違います: 「2026.9」（Hub は {S.SCHEMA_VERSION}）"]


def test_year_mismatch_between_sheet_and_json():
    errors = S.verify_workbook(list(S.ALL_SHEETS), good_headers(), good_settings(), 2028)
    assert errors == ["設定シートの年度が違います: 「2027」（年度設定JSONは 2028）"]


def test_missing_setting_keys_are_listed():
    settings = good_settings()
    del settings["受付終了"]
    settings["回答受付"] = ""
    errors = S.verify_workbook(list(S.ALL_SHEETS), good_headers(), settings, 2027)
    assert errors == ["設定シートに値がありません: 受付終了、回答受付"]


def test_settings_not_checked_when_settings_header_is_broken():
    """設定シートの列並びが壊れているときは、その中身の検査まで進めない（二重のエラーを避ける）。"""
    headers = good_headers()
    headers[S.SHEET_SETTINGS] = ["key", "value"]
    errors = S.verify_workbook(list(S.ALL_SHEETS), headers, {}, 2027)
    assert errors == [
        "設定シートの1列目が想定外です: 「key」（期待: 「キー」）",
        "設定シートの2列目が想定外です: 「value」（期待: 「値」）",
        "設定シートに列がありません: 備考",
    ]


def test_rows_to_dicts_pads_short_rows_and_skips_blank_rows():
    header = ("A", "B", "C")
    rows = [["1", "2"], [], ["", "", ""], ["x", "y", "z", "extra"], [None, "only-b"]]
    assert S.rows_to_dicts(header, rows) == [
        {"A": "1", "B": "2", "C": ""},
        {"A": "x", "B": "y", "C": "z"},
        {"A": "", "B": "only-b", "C": ""},
    ]


def test_split_header_and_settings_to_kv():
    header, rows = S.split_header([["キー", "値", "備考"], ["年度", "2027", ""], ["", "捨てる"], ["回答受付", "1"]])
    assert header == ["キー", "値", "備考"]
    assert S.settings_to_kv(rows) == {"年度": "2027", "回答受付": "1"}
    assert S.split_header([]) == ([], [])


def test_assert_writable_only_targets_and_audit():
    S.assert_writable(S.SHEET_TARGETS)
    S.assert_writable(S.SHEET_AUDIT)
    for sheet in (S.SHEET_RESPONSES, S.SHEET_OPTIONS, S.SHEET_SETTINGS, "適当"):
        with pytest.raises(S.SchemaError):
            S.assert_writable(sheet)


def test_target_headers_hub_columns_boundary():
    """Hub が書くのは先頭14列。15列目からは Apps Script の領域。"""
    assert S.TARGET_HEADERS[S.TARGET_HUB_COLUMNS - 1] == "登録者"
    assert S.TARGET_HEADERS[S.TARGET_HUB_COLUMNS] == "トークンハッシュ"
    assert len(set(S.TARGET_HEADERS)) == len(S.TARGET_HEADERS)
    assert len(set(S.RESPONSE_HEADERS)) == len(S.RESPONSE_HEADERS)
