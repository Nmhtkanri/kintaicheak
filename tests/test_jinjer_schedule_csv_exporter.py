"""jinjer スケジュール一括登録 CSV エクスポータのテスト"""
from __future__ import annotations

import csv
import os
import sys
import tempfile
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.jinjer_schedule_csv_exporter import (
    _is_ake_code,
    _off_value_for_weekday,
    _resolve_merged_cell_value,
    _resolve_cell_value,
    _sanitize_filename_part,
    build_legend_to_template_name,
    export_jinjer_schedule_csv,
    export_jinjer_schedule_csv_split,
)
from services.shift_resolver import normalize_legend
from services.jinjer_api_client import (
    build_name_to_id_map,
    pick_attendance_group_at,
)


# =============================================================================
# 基本ヘルパーのテスト
# =============================================================================

def test_ake_code_detection():
    assert _is_ake_code("明") is True
    assert _is_ake_code("明け") is True
    assert _is_ake_code("明休") is True
    assert _is_ake_code("0", label="明け休") is True
    assert _is_ake_code("B") is False
    assert _is_ake_code("") is False


def test_off_value_for_weekday_a_pattern():
    # Mon=0 ... Sun=6
    assert _off_value_for_weekday(0) == "所"  # 月
    assert _off_value_for_weekday(1) == "所"  # 火
    assert _off_value_for_weekday(2) == "所"  # 水
    assert _off_value_for_weekday(3) == "所"  # 木
    assert _off_value_for_weekday(4) == "所"  # 金
    assert _off_value_for_weekday(5) == "所"  # 土
    assert _off_value_for_weekday(6) == "法"  # 日


# =============================================================================
# セル値解決のテスト
# =============================================================================

def _legend_for_test():
    return [
        {"code": "B", "label": "B勤", "start_time": "12:30", "end_time": "21:00",
         "break_minutes": 60, "is_off": False},
        {"code": "明", "label": "明け休", "is_off": True},
        {"code": "公", "label": "公休", "is_off": True},
    ]


def test_resolve_cell_ake_returns_zero():
    legend = normalize_legend(_legend_for_test())
    code_to_name = {"B": "B勤"}
    val = _resolve_cell_value(
        "明", date(2026, 4, 1), legend, code_to_name, set()
    )
    assert val == "0"


def test_resolve_cell_off_on_sunday():
    legend = normalize_legend(_legend_for_test())
    code_to_name = {"B": "B勤"}
    # 2026-04-05 は日曜
    val = _resolve_cell_value(
        "公", date(2026, 4, 5), legend, code_to_name, set()
    )
    assert val == "法"


def test_resolve_cell_off_on_saturday():
    legend = normalize_legend(_legend_for_test())
    code_to_name = {"B": "B勤"}
    # 2026-04-04 は土曜
    val = _resolve_cell_value(
        "公", date(2026, 4, 4), legend, code_to_name, set()
    )
    assert val == "所"


def test_resolve_cell_blank_on_weekday():
    legend = normalize_legend(_legend_for_test())
    code_to_name = {"B": "B勤"}
    # 2026-04-01 は水曜
    val = _resolve_cell_value(
        "", date(2026, 4, 1), legend, code_to_name, set()
    )
    assert val == "所"


def test_resolve_cell_template_name():
    legend = normalize_legend(_legend_for_test())
    code_to_name = {"B": "BCS基本"}
    val = _resolve_cell_value(
        "B", date(2026, 4, 1), legend, code_to_name, set()
    )
    assert val == "BCS基本"


def test_resolve_cell_unknown_code_returned_as_is():
    legend = normalize_legend(_legend_for_test())
    code_to_name = {"B": "B勤"}
    val = _resolve_cell_value(
        "X", date(2026, 4, 1), legend, code_to_name, set()
    )
    assert val == "X"


# =============================================================================
# 氏名→IDマップ
# =============================================================================

def test_build_name_to_id_map_variants():
    employees = [
        {"id": 1234, "company": {"last_name": "小嶋", "first_name": "桃子"}},
        {"id": 5678, "company": {"last_name": "戸松", "first_name": "工"}},
    ]
    name_map = build_name_to_id_map(employees)
    # 連結
    assert name_map["小嶋桃子"] == "1234"
    assert name_map["小嶋 桃子"] == "1234"
    assert name_map["小嶋　桃子"] == "1234"
    # 姓のみ
    assert name_map["小嶋"] == "1234"


def test_build_name_to_id_map_skips_missing_id():
    employees = [
        {"company": {"last_name": "氏名のみ", "first_name": "ID無し"}},  # id 無し
    ]
    name_map = build_name_to_id_map(employees)
    assert name_map == {}


# =============================================================================
# CSV 出力統合テスト
# =============================================================================

def test_export_csv_basic_round_trip(tmp_path):
    legend = _legend_for_test()
    employees = [
        {
            "name": "小嶋桃子",
            "shifts": [
                {"date": "2026-04-01", "code": "B"},
                {"date": "2026-04-02", "code": "B"},
                {"date": "2026-04-03", "code": "B"},
                {"date": "2026-04-04", "code": "公"},   # 土
                {"date": "2026-04-05", "code": "公"},   # 日
                {"date": "2026-04-06", "code": "明"},  # 明け休
            ],
        },
    ]
    name_to_id = {"小嶋桃子": "1234"}
    out = tmp_path / "out.csv"

    result = export_jinjer_schedule_csv(
        legend=legend,
        employees=employees,
        year=2026,
        month=4,
        name_to_id=name_to_id,
        output_path=str(out),
    )

    assert result["rows"] == 1
    assert result["missing_ids"] == []
    assert os.path.exists(out)

    # CP932で読み戻す
    with open(out, "r", encoding="cp932", newline="") as f:
        reader = csv.reader(f)
        all_rows = list(reader)

    # 30日 + 先頭2セル
    assert len(all_rows) == 3  # ヘッダー2行 + 1人分
    assert all_rows[0][0] == "2026年"
    assert all_rows[0][1] == "4月"
    assert all_rows[1][0] == "氏名"
    assert all_rows[1][1] == "従業員ID"
    # 4月の日数 = 30
    assert len(all_rows[0]) == 32

    # 1人目の行
    emp_row = all_rows[2]
    assert emp_row[0] == "小嶋桃子"
    assert emp_row[1] == "1234"
    # 4/1 (水) → 新規雛形 CSV と共有する ID 候補
    assert emp_row[2] == "B"
    # 4/4 (土) → 公 → "所"
    assert emp_row[5] == "所"
    # 4/5 (日) → 公 → "法"
    assert emp_row[6] == "法"
    # 4/6 (月) → 明 → "0"
    assert emp_row[7] == "0"


def test_export_csv_records_missing_id(tmp_path):
    legend = _legend_for_test()
    employees = [
        {
            "name": "未登録さん",
            "shifts": [{"date": "2026-04-01", "code": "B"}],
        },
    ]
    out = tmp_path / "out.csv"
    result = export_jinjer_schedule_csv(
        legend=legend,
        employees=employees,
        year=2026,
        month=4,
        name_to_id={},
        output_path=str(out),
    )
    assert result["missing_ids"] == ["未登録さん"]


def test_export_csv_missing_year_raises(tmp_path):
    legend = _legend_for_test()
    out = tmp_path / "out.csv"
    with pytest.raises(ValueError):
        export_jinjer_schedule_csv(
            legend=legend,
            employees=[],
            year=None,
            month=4,
            name_to_id={},
            output_path=str(out),
        )


def test_export_csv_blank_day_for_active_employee(tmp_path):
    """シフト記載のない日も "所" 等で埋まること"""
    legend = _legend_for_test()
    employees = [
        {"name": "小嶋桃子", "shifts": [{"date": "2026-04-01", "code": "B"}]},
    ]
    name_to_id = {"小嶋桃子": "1234"}
    out = tmp_path / "out.csv"

    export_jinjer_schedule_csv(
        legend=legend, employees=employees,
        year=2026, month=4,
        name_to_id=name_to_id,
        output_path=str(out),
    )

    with open(out, "r", encoding="cp932", newline="") as f:
        all_rows = list(csv.reader(f))

    emp_row = all_rows[2]
    # 4/2 (木) は記録なし → "所"
    assert emp_row[3] == "所"
    # 4/5 (日) は記録なし → "法"
    assert emp_row[6] == "法"


def test_build_legend_to_template_name_no_csv():
    """雛形CSVが無い場合も jinjer 用の ID 候補へフォールバック"""
    legend = _legend_for_test()
    code_to_name = build_legend_to_template_name(legend, "")
    assert code_to_name == {"B": "B"}


def test_resolve_merged_unmatched_cell_uses_template_id_candidate():
    cell_value, unmatched = _resolve_merged_cell_value(
        {
            "merged_code": "4+a",
            "merged_label": "4番勤務+a勤務",
            "merged_start": "16:30",
            "merged_end": "33:00",
            "merged_break": 0,
        },
        templates=[],
    )

    assert cell_value == "4a"
    assert unmatched["code"] == "4+a"


# =============================================================================
# 打刻グループ判定 (pick_attendance_group_at)
# =============================================================================

def test_pick_attendance_group_picks_latest_at_or_before_target():
    affiliations = [
        {"date_of_issue": "2022-01-01",
         "attendance_group": {"id": "54", "name": "ワークフロー"}},
        {"date_of_issue": "2026-04-01",
         "attendance_group": {"id": "49", "name": "時給制"}},
        # 順番がバラバラでも OK（API レスポンスは時系列順とは限らない）
        {"date_of_issue": "2026-02-01",
         "attendance_group": {"id": "54", "name": "ワークフロー"}},
    ]
    # 2026-04-01 → "時給制"
    assert pick_attendance_group_at(affiliations, date(2026, 4, 1)) == ("49", "時給制")
    # 2026-03-15 → "ワークフロー"（2026-02-01 が最新の <= target）
    assert pick_attendance_group_at(affiliations, date(2026, 3, 15)) == ("54", "ワークフロー")
    # 2021-12-31 → 該当なし
    assert pick_attendance_group_at(affiliations, date(2021, 12, 31)) == ("", "")


def test_pick_attendance_group_skips_blank_id():
    """履歴に attendance_group.id が空のレコードが混ざっていても、遡って非空を採用"""
    affiliations = [
        {"date_of_issue": "2022-01-01",
         "attendance_group": {"id": "306", "name": "ホンダ芳賀"}},
        {"date_of_issue": "2024-04-01",
         "attendance_group": {"id": "", "name": ""}},  # 未設定の中間レコード
    ]
    # 2025-01-01 時点 → 空でない 306 を採用
    assert pick_attendance_group_at(affiliations, date(2025, 1, 1)) == ("306", "ホンダ芳賀")


def test_pick_attendance_group_accepts_iso_string():
    affiliations = [
        {"date_of_issue": "2026-04-01",
         "attendance_group": {"id": "43", "name": "140-160時間制"}},
    ]
    assert pick_attendance_group_at(affiliations, "2026-05-01") == ("43", "140-160時間制")


def test_pick_attendance_group_empty_input():
    assert pick_attendance_group_at([], date(2026, 5, 1)) == ("", "")


# =============================================================================
# ファイル名サニタイズ
# =============================================================================

def test_sanitize_filename_part_strips_invalid_chars():
    assert _sanitize_filename_part("140-160時間制") == "140-160時間制"
    assert _sanitize_filename_part("ホンダ/芳賀") == "ホンダ_芳賀"
    assert _sanitize_filename_part("a:b*c?d") == "a_b_c_d"
    assert _sanitize_filename_part("時給 制") == "時給_制"
    assert _sanitize_filename_part("") == "未分類"
    assert _sanitize_filename_part(None) == "未分類"


# =============================================================================
# 打刻グループ別 CSV 分割エクスポート
# =============================================================================

def test_export_split_creates_one_csv_per_group(tmp_path):
    """異なる打刻グループの従業員が混在 → グループ別に分割される"""
    legend = _legend_for_test()
    employees = [
        {"name": "小嶋桃子",
         "shifts": [{"date": "2026-05-01", "code": "B"}]},
        {"name": "戸松工",
         "shifts": [{"date": "2026-05-01", "code": "B"}]},
        {"name": "大堀広智",
         "shifts": [{"date": "2026-05-01", "code": "B"}]},
    ]
    name_to_id = {
        "小嶋桃子": "1001",
        "戸松工": "1002",
        "大堀広智": "1003",
    }
    attendance_group_map = {
        "1001": ("43", "140-160時間制"),
        "1002": ("43", "140-160時間制"),
        "1003": ("47", "140-180時間制"),
    }

    result = export_jinjer_schedule_csv_split(
        legend=legend,
        employees=employees,
        year=2026,
        month=5,
        name_to_id=name_to_id,
        attendance_group_map=attendance_group_map,
        output_dir=str(tmp_path),
        filename_hash="abc123",
    )

    assert len(result["csv_files"]) == 2
    by_group = {f["attendance_group_id"]: f for f in result["csv_files"]}
    assert by_group["43"]["rows"] == 2
    assert by_group["43"]["attendance_group_name"] == "140-160時間制"
    assert by_group["47"]["rows"] == 1
    assert by_group["47"]["attendance_group_name"] == "140-180時間制"

    # ファイル名にグループ名とハッシュが入る
    for f in result["csv_files"]:
        assert f["attendance_group_name"] in f["filename"]
        assert "abc123" in f["filename"]
        assert f["filename"].endswith(".csv")
        assert os.path.exists(f["path"])

    # 140-160 の CSV に大堀が混ざってない
    with open(by_group["43"]["path"], "r", encoding="cp932", newline="") as f:
        rows = list(csv.reader(f))
    names_in_140 = [r[0] for r in rows[2:]]
    assert "大堀広智" not in names_in_140


def test_export_split_routes_unknown_group_to_misc_bucket(tmp_path):
    """打刻グループが取れなかった従業員は "未分類" バケツに入る"""
    legend = _legend_for_test()
    employees = [
        {"name": "小嶋桃子",
         "shifts": [{"date": "2026-05-01", "code": "B"}]},
        {"name": "ID無しさん",
         "shifts": [{"date": "2026-05-01", "code": "B"}]},
    ]
    name_to_id = {"小嶋桃子": "1001"}
    attendance_group_map = {"1001": ("43", "140-160時間制")}

    result = export_jinjer_schedule_csv_split(
        legend=legend,
        employees=employees,
        year=2026,
        month=5,
        name_to_id=name_to_id,
        attendance_group_map=attendance_group_map,
        output_dir=str(tmp_path),
    )

    files_by_group = {f["attendance_group_name"]: f for f in result["csv_files"]}
    assert "140-160時間制" in files_by_group
    assert "未分類" in files_by_group
    assert "ID無しさん" in result["ungrouped"]


def test_export_split_with_empty_group_map_produces_one_misc_file(tmp_path):
    """attendance_group_map が空 → 全員 "未分類" にまとめて 1 ファイル出力"""
    legend = _legend_for_test()
    employees = [
        {"name": "小嶋桃子",
         "shifts": [{"date": "2026-05-01", "code": "B"}]},
        {"name": "戸松工",
         "shifts": [{"date": "2026-05-01", "code": "B"}]},
    ]
    name_to_id = {"小嶋桃子": "1001", "戸松工": "1002"}

    result = export_jinjer_schedule_csv_split(
        legend=legend,
        employees=employees,
        year=2026,
        month=5,
        name_to_id=name_to_id,
        attendance_group_map={},  # API失敗等で空
        output_dir=str(tmp_path),
    )

    assert len(result["csv_files"]) == 1
    assert result["csv_files"][0]["attendance_group_name"] == "未分類"
    assert result["csv_files"][0]["rows"] == 2
    # 2 名とも ungrouped 扱い
    assert len(result["ungrouped"]) == 2
