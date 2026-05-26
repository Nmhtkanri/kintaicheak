"""jinjer_template_matcher.py の単体テスト"""

import os
import sys
import csv
import tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest

from services.jinjer_template_matcher import (
    load_jinjer_templates,
    find_matching_template,
    match_legend_to_templates,
    generate_new_templates_csv,
    TEMPLATE_CSV_HEADERS,
    _normalize_time_str,
    canonicalize_overnight_times,
)


# =============================================================================
# テスト用 fixture: ミニ jinjer 雛形 CSV
# =============================================================================

@pytest.fixture
def mini_template_csv(tmp_path):
    """本物の雛形（_2026-04-27 の新フォーマット）をベースにした最小 CSV"""
    path = tmp_path / "templates.csv"
    rows = [
        # No 16: 日勤 9:00-17:30
        {"No": "16", "＊スケジュール雛形名": "日勤", "略称(3文字以内)": "日勤",
         "＊スケジュール雛形ID": "1", "表示順": "9998", "半休ID": "1",
         "＊出勤時間(0:00~47:59)": "9:00:00", "＊退勤時間(0:00~47:59)": "17:30:00",
         "休憩開始時間1(0:00~47:59)": "12:00:00", "復帰時間1(0:00~47:59)": "13:00:00"},
        # No 20: 遅番1 13:00-21:30
        {"No": "20", "＊スケジュール雛形名": "遅番1", "略称(3文字以内)": "遅1",
         "＊スケジュール雛形ID": "5", "表示順": "9998", "半休ID": "1",
         "＊出勤時間(0:00~47:59)": "13:00:00", "＊退勤時間(0:00~47:59)": "21:30:00",
         "休憩開始時間1(0:00~47:59)": "17:00:00", "復帰時間1(0:00~47:59)": "18:00:00"},
        # No 21: D1 10:00-18:30
        {"No": "21", "＊スケジュール雛形名": "D1", "略称(3文字以内)": "D1",
         "＊スケジュール雛形ID": "6", "表示順": "9998", "半休ID": "1",
         "＊出勤時間(0:00~47:59)": "10:00:00", "＊退勤時間(0:00~47:59)": "18:30:00",
         "休憩開始時間1(0:00~47:59)": "13:00:00", "復帰時間1(0:00~47:59)": "14:00:00"},
    ]
    with open(path, "w", encoding="cp932", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=TEMPLATE_CSV_HEADERS)
        writer.writeheader()
        for r in rows:
            full = {h: "" for h in TEMPLATE_CSV_HEADERS}
            full.update(r)
            writer.writerow(full)
    return str(path)


# =============================================================================
# load_jinjer_templates
# =============================================================================

def test_load_templates_cp932(mini_template_csv):
    rows = load_jinjer_templates(mini_template_csv)
    assert len(rows) == 3
    assert rows[0]["＊スケジュール雛形名"] == "日勤"


def test_load_missing_file_returns_empty():
    rows = load_jinjer_templates("/nonexistent/path.csv")
    assert rows == []


# =============================================================================
# find_matching_template
# =============================================================================

def test_find_match_exact(mini_template_csv):
    templates = load_jinjer_templates(mini_template_csv)
    # G記号 = 10:00-18:30 → No21 D1 にマッチ
    tpl = find_matching_template("10:00", "18:30", 60, templates)
    assert tpl is not None
    assert tpl["＊スケジュール雛形名"] == "D1"


def test_find_match_none(mini_template_csv):
    templates = load_jinjer_templates(mini_template_csv)
    # B記号 = 12:30-21:00 → 一致なし（13:00-21:30 とはズレてる）
    tpl = find_matching_template("12:30", "21:00", 60, templates)
    assert tpl is None


def test_find_match_rolls_overnight_end_into_jinjer_time():
    templates = [{
        "＊スケジュール雛形名": "深夜",
        "＊スケジュール雛形ID": "NMHT03",
        "＊出勤時間(0:00~47:59)": "24:00:00",
        "＊退勤時間(0:00~47:59)": "33:00:00",
    }]

    tpl = find_matching_template("24:00", "09:00", 0, templates)

    assert tpl is templates[0]


# =============================================================================
# match_legend_to_templates
# =============================================================================

def test_match_legend_mixed(mini_template_csv):
    legend = [
        {"code": "G", "label": "G勤", "start_time": "10:00", "end_time": "18:30", "break_minutes": 60},
        {"code": "B", "label": "B勤", "start_time": "12:30", "end_time": "21:00", "break_minutes": 60},
        {"code": "A", "label": "日勤", "start_time": "9:00", "end_time": "17:30", "break_minutes": 60},
        {"code": "明", "label": "明け休", "is_off": True},  # 雛形マッチ対象外
    ]
    result = match_legend_to_templates(legend, mini_template_csv)
    assert len(result["matched"]) == 2  # G, A
    assert len(result["unmatched"]) == 1  # B
    matched_codes = {m["code"] for m in result["matched"]}
    assert matched_codes == {"G", "A"}
    assert result["unmatched"][0]["code"] == "B"


# =============================================================================
# generate_new_templates_csv
# =============================================================================

def test_generate_new_templates(mini_template_csv, tmp_path):
    unmatched = [
        {"code": "B", "label": "B勤", "start_time": "12:30", "end_time": "21:00", "break_minutes": 60},
        {"code": "●", "label": "夜勤", "start_time": "16:30", "end_time": "25:00", "break_minutes": 60},
    ]
    output = tmp_path / "new_templates.csv"
    result = generate_new_templates_csv(unmatched, mini_template_csv, str(output))

    assert result["count"] == 2
    assert os.path.exists(output)

    # 内容確認
    with open(output, encoding="cp932") as f:
        rows = list(csv.DictReader(f))
    assert len(rows) == 2

    # No は既存の最大(21) +1 から始まる
    assert int(rows[0]["No"]) == 22
    assert int(rows[1]["No"]) == 23

    # 時刻が正しく入っている
    assert rows[0]["＊出勤時間(0:00~47:59)"] == "12:30:00"
    assert rows[0]["＊退勤時間(0:00~47:59)"] == "21:00:00"
    assert rows[1]["＊出勤時間(0:00~47:59)"] == "16:30:00"
    assert rows[1]["＊退勤時間(0:00~47:59)"] == "25:00:00"

    # 名称・ID も入っている
    assert rows[0]["＊スケジュール雛形名"] == "B勤"
    assert rows[0]["略称(3文字以内)"] == "B"


def test_generate_empty_unmatched(mini_template_csv, tmp_path):
    """unmatched が空のときは何も生成しない"""
    output = tmp_path / "empty.csv"
    result = generate_new_templates_csv([], mini_template_csv, str(output))
    assert result["count"] == 0
    assert result["path"] is None


def test_generate_new_template_rolls_overnight_end_forward(mini_template_csv, tmp_path):
    output = tmp_path / "overnight.csv"
    generate_new_templates_csv(
        [{
            "code": "5",
            "label": "深夜",
            "start_time": "23:55",
            "end_time": "09:00",
            "break_minutes": 0,
        }],
        mini_template_csv,
        str(output),
    )

    with open(output, encoding="cp932") as f:
        row = next(csv.DictReader(f))

    assert row["＊退勤時間(0:00~47:59)"] == "33:00:00"


# =============================================================================
# _normalize_time_str
# =============================================================================

def test_normalize_time_variants():
    assert _normalize_time_str("9:00:00") == "9:00"
    assert _normalize_time_str("09:00") == "9:00"
    assert _normalize_time_str("17:30:00") == "17:30"
    assert _normalize_time_str("") == ""


def test_canonicalize_overnight_times_keeps_non_overnight_time():
    assert canonicalize_overnight_times("16:30", "24:00") == ("16:30", "24:00")
