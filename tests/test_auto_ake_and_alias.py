"""夜勤明けの自動割当（退勤30:00以降）と、氏名エイリアス表のテスト

背景（2026-07-31）:
  - KDX勤務シフト表は姓のみ（"吉田"）で氏名を持つが、jinjer には吉田さんが2名いる。
    姓だけでは自動確定できず、従業員ID空欄の「未分類」CSVに落ちていた。
  - 夜勤明けの「休み」はシフト表の「ー」記号頼みだった。退勤時刻から機械的に決める。
"""
from __future__ import annotations

import csv
import io
import os
import sys
import tempfile
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.employee_alias import (
    alias_csv_path,
    apply_aliases,
    load_aliases_for_source,
    load_employee_aliases,
)
from services.jinjer_schedule_csv_exporter import (
    AKE_REST_VALUE,
    _apply_auto_ake_rest,
    _raw_time_to_minutes,
    export_jinjer_schedule_csv,
)


# =============================================================================
# 24時超表記のパース
# =============================================================================

@pytest.mark.parametrize("raw,expected", [
    ("34:00", 34 * 60),
    ("33:30", 33 * 60 + 30),
    ("30:00", 30 * 60),
    ("9:00", 9 * 60),
    ("17:30", 17 * 60 + 30),
    ("  34:00 ", 34 * 60),
])
def test_raw_time_to_minutes(raw, expected):
    assert _raw_time_to_minutes(raw) == expected


@pytest.mark.parametrize("raw", [None, "", "abc", "34:60", "34", "3400"])
def test_raw_time_to_minutes_invalid(raw):
    assert _raw_time_to_minutes(raw) is None


# =============================================================================
# 夜勤明けの自動割当
# =============================================================================

def _days(year, month, n):
    return [date(year, month, d) for d in range(1, n + 1)]


def test_auto_ake_sets_next_day_to_rest():
    """退勤34:00の翌日が休扱い（所）→「休み」に置き換わる"""
    cells = ["N20", "所", "所"]
    end_minutes = [34 * 60, None, None]
    applied, sched, conflicts = _apply_auto_ake_rest(cells, end_minutes, _days(2026, 8, 3), "池田")

    assert cells == ["N20", AKE_REST_VALUE, "所"]
    assert conflicts == []
    assert len(applied) == 1
    assert applied[0]["night_day"] == 1
    assert applied[0]["ake_day"] == 2
    assert applied[0]["before"] == "所"


def test_auto_ake_threshold_is_30h():
    """退勤29:59は対象外・30:00ちょうどは対象"""
    cells_under = ["X", "所"]
    _apply_auto_ake_rest(cells_under, [29 * 60 + 59, None], _days(2026, 8, 2), "A")
    assert cells_under == ["X", "所"]

    cells_exact = ["X", "所"]
    _apply_auto_ake_rest(cells_exact, [30 * 60, None], _days(2026, 8, 2), "B")
    assert cells_exact == ["X", AKE_REST_VALUE]


def test_auto_ake_day_shift_untouched():
    """日勤（17:30退勤）の翌日は触らない"""
    cells = ["1", "1", "所"]
    applied, sched, conflicts = _apply_auto_ake_rest(
        cells, [17 * 60 + 30, 17 * 60 + 30, None], _days(2026, 8, 3), "尾川")
    assert cells == ["1", "1", "所"]
    assert applied == [] and sched == [] and conflicts == []


def test_auto_ake_keeps_existing_rest():
    """すでに「休み」（シフト表の「ー」）なら二重に記録しない"""
    cells = ["N20", AKE_REST_VALUE]
    applied, sched, conflicts = _apply_auto_ake_rest(
        cells, [34 * 60, None], _days(2026, 8, 2), "菊池")
    assert cells == ["N20", AKE_REST_VALUE]
    assert applied == [] and sched == [] and conflicts == []


def test_auto_ake_does_not_overwrite_real_shift():
    """翌日に勤務予定があるときは上書きせず、シフト表を優先する（警告ではない）"""
    cells = ["N20", "1"]
    applied, sched, conflicts = _apply_auto_ake_rest(
        cells, [34 * 60, 17 * 60 + 30], _days(2026, 8, 2), "横山")
    assert cells == ["N20", "1"]          # 勝手に消さない
    assert applied == []
    assert conflicts == []                # 要確認ではない＝正常なパターン
    assert len(sched) == 1
    assert sched[0]["ake_day"] == 2
    assert sched[0]["next_value"] == "1"


def test_auto_ake_consecutive_night_shifts_keep_schedule():
    """連日夜勤（8/10・8/11 とも 16:45-33:30）はシフト表を優先し、明け休にしない

    2026-07-31 谷津さん指示。8/10 の退勤が 33:30 でも、8/11 にシフト記号が
    入っているならそれが正。8/11 を「休み」に潰してはいけない。
    8/11 の翌日（8/12）に予定が無ければ、そちらは明け休になる。
    """
    # [8/10, 8/11, 8/12]
    cells = ["夜20", "夜20", "所"]
    end = [33 * 60 + 30, 33 * 60 + 30, None]
    applied, sched, conflicts = _apply_auto_ake_rest(
        cells, end, [date(2026, 8, 10), date(2026, 8, 11), date(2026, 8, 12)], "瀧澤")

    assert cells == ["夜20", "夜20", AKE_REST_VALUE]
    assert conflicts == []
    assert len(sched) == 1 and sched[0]["night_day"] == 10 and sched[0]["ake_day"] == 11
    assert len(applied) == 1 and applied[0]["night_day"] == 11 and applied[0]["ake_day"] == 12


def test_export_consecutive_night_shifts_keep_schedule(tmp_path):
    """CSV出力まで通して、連日夜勤の2日目が潰れないことを確認する"""
    legend = [
        {"code": "C1", "label": "KDX夜勤", "start_time": "16:30", "end_time": "34:00",
         "break_minutes": 120, "is_off": False},
    ]
    employees = [{
        "name": "テスト太郎",
        "shifts": [{"date": "2026-08-10", "code": "C1"},
                   {"date": "2026-08-11", "code": "C1"},
                   {"date": "2026-08-12", "code": ""}],
    }]
    out = str(tmp_path / "consecutive.csv")
    result = export_jinjer_schedule_csv(
        legend=legend, employees=employees, year=2026, month=8,
        name_to_id={"テスト太郎": "2020001"}, output_path=out,
    )

    with open(out, encoding="cp932") as f:
        row = list(csv.reader(f))[2]
    # [氏名, 従業員ID, 1日, 2日, ...] → 10日は index 11
    assert row[11] == row[12], "連日夜勤の2日目がシフト表どおり残っていない"
    assert row[11] != AKE_REST_VALUE
    assert row[13] == AKE_REST_VALUE      # 8/12 は明け休になる
    assert len(result["ake_schedule_priority"]) == 1
    assert result["ake_schedule_priority"][0]["ake_day"] == 11


def test_auto_ake_month_end_reports_conflict():
    """月末の夜勤は明けが翌月 → 当月CSVでは設定できないので要確認"""
    cells = ["所", "N20"]
    applied, sched, conflicts = _apply_auto_ake_rest(
        cells, [None, 34 * 60], _days(2026, 8, 2), "大貫")
    assert applied == []
    assert len(conflicts) == 1
    assert conflicts[0]["ake_day"] is None


def test_auto_ake_applies_to_merged_overnight_shift(tmp_path):
    """深夜跨ぎ統合（16:45-24:00 + 00:00-09:30 → 16:45-33:30）の翌々日ではなく、
    **吸収された翌日**が「休み」になる。

    ⚠️ これは KDX 以外（夜勤突合フロー）にも効く挙動変更。
    従来は吸収日を休扱い（所/法）にしていたが、退勤33:30 の翌朝まで勤務が続く日を
    所定休日として立てるのは実態と合わないため「休み」に寄せる（2026-07-31 の
    「退勤30:00以降は翌日を休みにする」指示に従う）。
    """
    legend = [
        {"code": "夜", "label": "夜勤", "start_time": "16:45", "end_time": "24:00",
         "break_minutes": 60, "is_off": False},
        {"code": "明", "label": "明け", "start_time": "24:00", "end_time": "33:30",
         "break_minutes": 0, "is_off": False},
    ]
    employees = [{
        "name": "テスト太郎",
        "shifts": [{"date": "2026-08-01", "code": "夜"},
                   {"date": "2026-08-02", "code": "明"}],
    }]
    out = str(tmp_path / "merged.csv")
    result = export_jinjer_schedule_csv(
        legend=legend, employees=employees, year=2026, month=8,
        name_to_id={"テスト太郎": "2020001"}, output_path=out,
    )

    with open(out, encoding="cp932") as f:
        data_row = list(csv.reader(f))[2]
    assert len(result["merges"]) == 1                 # 統合は従来通り成立
    assert data_row[3] == AKE_REST_VALUE              # 8/2（吸収日）が「休み」
    assert len(result["ake_auto"]) == 1
    assert result["ake_auto"][0]["before"] in ("所", "法")


def test_export_applies_auto_ake_end_to_end(tmp_path):
    """CSV出力まで通して、夜勤の翌日が「休み」になる（「ー」記号が無い表でも）"""
    legend = [
        {"code": "C1", "label": "KDX夜勤", "start_time": "16:30", "end_time": "34:00",
         "break_minutes": 120, "is_off": False},
    ]
    employees = [{
        "name": "テスト太郎",
        # 8/1 夜勤 → 8/2 は空欄（明け記号なし）
        "shifts": [{"date": "2026-08-01", "code": "C1"},
                   {"date": "2026-08-02", "code": ""}],
    }]
    out = str(tmp_path / "out.csv")
    result = export_jinjer_schedule_csv(
        legend=legend, employees=employees, year=2026, month=8,
        name_to_id={"テスト太郎": "2020001"}, output_path=out,
    )

    with open(out, encoding="cp932") as f:
        rows = list(csv.reader(f))
    data_row = rows[2]
    # [氏名, 従業員ID, 1日, 2日, ...]
    assert data_row[3] == AKE_REST_VALUE
    assert len(result["ake_auto"]) == 1
    assert result["ake_auto"][0]["ake_day"] == 2


# =============================================================================
# 氏名エイリアス表
# =============================================================================

def _write_csv(path, rows, encoding="utf-8-sig"):
    with io.open(path, "w", encoding=encoding, newline="") as f:
        csv.writer(f).writerows(rows)


def test_load_employee_aliases(tmp_path):
    p = tmp_path / "alias.csv"
    _write_csv(p, [["シフト表氏名", "従業員ID", "備考"],
                   ["吉田", "2025007", "KDXの吉田は拓矢さん"]])
    assert load_employee_aliases(str(p)) == {"吉田": "2025007"}


def test_load_employee_aliases_cp932(tmp_path):
    """Excel で保存し直して CP932 になっても読める"""
    p = tmp_path / "alias_sjis.csv"
    _write_csv(p, [["シフト表氏名", "従業員ID"], ["吉田", "2025007"]], encoding="cp932")
    assert load_employee_aliases(str(p)) == {"吉田": "2025007"}


def test_load_employee_aliases_missing_file():
    assert load_employee_aliases(r"C:\存在しない\alias.csv") == {}
    assert load_employee_aliases("") == {}


def test_load_employee_aliases_rejects_non_employee_id(tmp_path):
    """派遣・テスト番号（5/6/9始まり）は給与計算対象外なので採用しない"""
    p = tmp_path / "alias.csv"
    _write_csv(p, [["シフト表氏名", "従業員ID"],
                   ["吉田", "2025007"],
                   ["派遣さん", "5000001"],
                   ["空ID", ""]])
    assert load_employee_aliases(str(p)) == {"吉田": "2025007"}


def test_apply_aliases_is_non_destructive():
    base = {"尾川": "2009006"}
    merged = apply_aliases(base, {"吉田": "2025007"})
    assert merged == {"尾川": "2009006", "吉田": "2025007"}
    assert base == {"尾川": "2009006"}   # 元の辞書は書き換えない


def test_apply_aliases_wins_over_base():
    """同姓で確定できず未登録だった姓を、エイリアスで確定させられる"""
    merged = apply_aliases({"吉田 拓矢": "2025007"}, {"吉田": "2025007"})
    assert merged["吉田"] == "2025007"


def test_alias_scope_is_limited_to_kdx(tmp_path):
    """KDX以外の系統ではエイリアス表を読まない（別現場の同姓を巻き込まないため）"""
    _write_csv(tmp_path / "スケジュール氏名エイリアス_KDX.csv",
               [["シフト表氏名", "従業員ID"], ["吉田", "2025007"]])

    kdx_aliases, warning = load_aliases_for_source("kdx", str(tmp_path))
    assert kdx_aliases == {"吉田": "2025007"}
    assert warning == ""

    for other in ("", "higashi", "monthly_xlsx", None):
        aliases, warn = load_aliases_for_source(other, str(tmp_path))
        assert aliases == {}, f"source={other!r} でエイリアスが適用されてしまった"
        assert warn == ""


def test_alias_csv_path_unknown_source(tmp_path):
    assert alias_csv_path("kdx", str(tmp_path)).endswith("スケジュール氏名エイリアス_KDX.csv")
    assert alias_csv_path("higashi", str(tmp_path)) == ""
    assert alias_csv_path("kdx", "") == ""


def test_export_with_alias_resolves_surname_only(tmp_path):
    """姓だけの「吉田」がエイリアス経由で従業員IDまで解決される"""
    legend = [{"code": "1", "label": "日勤", "start_time": "9:00", "end_time": "17:30",
               "break_minutes": 60, "is_off": False}]
    employees = [{"name": "吉田", "shifts": [{"date": "2026-08-01", "code": "1"}]}]

    # 同姓2名 → build_name_to_id_map は「吉田」を登録しない状態を再現
    base_name_to_id = {"吉田 拓矢": "2025007", "吉田 英伸": "2017016"}

    out_without = str(tmp_path / "without.csv")
    export_jinjer_schedule_csv(
        legend=legend, employees=employees, year=2026, month=8,
        name_to_id=base_name_to_id, output_path=out_without)
    with open(out_without, encoding="cp932") as f:
        assert list(csv.reader(f))[2][1] == ""   # 従業員ID空欄＝未分類行き

    out_with = str(tmp_path / "with.csv")
    export_jinjer_schedule_csv(
        legend=legend, employees=employees, year=2026, month=8,
        name_to_id=apply_aliases(base_name_to_id, {"吉田": "2025007"}),
        output_path=out_with)
    with open(out_with, encoding="cp932") as f:
        assert list(csv.reader(f))[2][1] == "2025007"
