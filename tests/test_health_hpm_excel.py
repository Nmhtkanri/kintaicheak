# -*- coding: utf-8 -*-
"""健診整形済みExcel（スキーマv2）の読み取りテスト

守りたいのは主に3つ:
  1. 血圧の1回目・2回目が別々のまま残ること（平均も複製も作らない）
  2. (-) は陰性として拾い、空欄や裸のハイフンを勝手に (-) にしないこと
  3. 旧スキーマ（v1）を確実に見分けて CSV 生成を止められること
"""

from __future__ import annotations

import os
import sys
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.health_hpm_excel import (  # noqa: E402
    NEGATIVE,
    Issue,
    format_cell_number,
    format_exam_no,
    normalize_qualitative,
    parse_exam_date,
    parse_health_workbook,
)
from tests.health_hpm_fixtures import bp_items, item, make_v2_workbook, person  # noqa: E402


def codes(result, level=None):
    return {i.code for i in result.all_issues() if level is None or i.level == level}


# ---------------------------------------------------------------------------
# 正常系
# ---------------------------------------------------------------------------

class TestParseHappyPath:
    def test_reads_persons_and_metrics(self, tmp_path):
        path = make_v2_workbook(tmp_path / "v2.xlsx", [
            person(items=[
                item("身体計測", "身長", "181.1", unit="cm", occurrence=1),
                item("身体計測", "体重", "89.3", unit="kg", occurrence=1),
                *bp_items(sys1="120", dia1="61", sys2="118", dia2="72"),
                item("尿検査", "尿蛋白", "(-)", occurrence=1, value_type="定性"),
            ]),
        ])
        result = parse_health_workbook(path, "v2.xlsx")

        assert result.schema_version == "2.0"
        assert result.genpyo_confirmed is True
        assert result.errors() == []
        assert len(result.persons) == 1

        p = result.persons[0]
        assert p.name == "友納 英彦"
        assert p.age == 58
        assert p.gender == "男性"
        assert p.exam_date == date(2026, 7, 1)
        assert p.exam_no == "000132", "受診No.は6桁ゼロ埋めで持つ"
        assert len(p.numeric()) == 6
        assert len(p.qualitative()) == 1

    def test_blood_pressure_keeps_both_rounds_separately(self, tmp_path):
        path = make_v2_workbook(tmp_path / "bp.xlsx", [
            person(items=bp_items(sys1="130", dia1="80", sys2="120", dia2="70")),
        ])
        result = parse_health_workbook(path)
        bp = result.persons[0].blood_pressure()

        assert bp == {1: {"sys": "130", "dia": "80"},
                      2: {"sys": "120", "dia": "70"}}
        # 平均（125/75）がどこにも現れないこと
        values = {m.value for m in result.persons[0].metrics}
        assert "125" not in values and "75" not in values

    def test_single_measurement_is_warning_not_duplicated(self, tmp_path):
        path = make_v2_workbook(tmp_path / "bp1.xlsx", [
            person(items=bp_items(sys1="120", dia1="61")),
        ])
        result = parse_health_workbook(path)
        p = result.persons[0]

        assert p.blood_pressure() == {1: {"sys": "120", "dia": "61"}}
        assert 2 not in p.blood_pressure(), "1回目を2回目へ複製しない"
        assert "BP_SINGLE" in codes(result, "warning")
        assert result.errors() == []

    def test_third_measurement_kept_but_warned(self, tmp_path):
        path = make_v2_workbook(tmp_path / "bp3.xlsx", [
            person(items=bp_items(sys1="130", dia1="80", sys2="120",
                                  dia2="70", sys3="118", dia3="68")),
        ])
        result = parse_health_workbook(path)
        bp = result.persons[0].blood_pressure()

        assert bp[1] == {"sys": "130", "dia": "80"}
        assert bp[2] == {"sys": "120", "dia": "70"}
        assert bp[3] == {"sys": "118", "dia": "68"}, "3回目もExcel側では保持する"
        assert "BP_THIRD_IGNORED" in codes(result, "warning")
        assert result.errors() == []


# ---------------------------------------------------------------------------
# スキーマ判定
# ---------------------------------------------------------------------------

class TestSchemaGate:
    def test_v1_workbook_is_rejected(self, tmp_path):
        path = make_v2_workbook(tmp_path / "v1.xlsx", [
            person(items=[item("血圧", "収縮期血圧", "120", unit="mmHg")]),
        ], schema_version=None, genpyo=False, bp_kept=False,
            qual_kept=False, v2_columns=False)
        result = parse_health_workbook(path)

        assert "SCHEMA_TOO_OLD" in codes(result, "error")
        messages = " ".join(i.message for i in result.errors())
        assert "スキーマ2.0で再整形" in messages
        # プレビューは出せる＝受診者は読めている
        assert len(result.persons) == 1

    def test_older_schema_version_rejected(self, tmp_path):
        path = make_v2_workbook(tmp_path / "v19.xlsx", [person()],
                                schema_version="1.9")
        result = parse_health_workbook(path)
        assert "SCHEMA_TOO_OLD" in codes(result, "error")

    def test_newer_schema_version_accepted(self, tmp_path):
        path = make_v2_workbook(tmp_path / "v21.xlsx", [person()],
                                schema_version="2.1")
        result = parse_health_workbook(path)
        assert "SCHEMA_TOO_OLD" not in codes(result)

    def test_genpyo_flag_required(self, tmp_path):
        path = make_v2_workbook(tmp_path / "nogenpyo.xlsx", [person()], genpyo=False)
        result = parse_health_workbook(path)
        assert "GENPYO_FLAG_MISSING" in codes(result, "error")

    def test_bp_occurrence_flag_required(self, tmp_path):
        path = make_v2_workbook(tmp_path / "nobp.xlsx", [person()], bp_kept=False)
        result = parse_health_workbook(path)
        assert "BP_OCCURRENCE_FLAG_MISSING" in codes(result, "error")


# ---------------------------------------------------------------------------
# 血圧の測定回
# ---------------------------------------------------------------------------

class TestOccurrence:
    def test_missing_occurrence_is_blocking(self, tmp_path):
        path = make_v2_workbook(tmp_path / "noocc.xlsx", [
            person(items=[item("血圧", "収縮期血圧", "120", unit="mmHg", occurrence=None)]),
        ])
        result = parse_health_workbook(path)

        assert "BP_OCCURRENCE_MISSING" in codes(result, "error")
        assert result.persons[0].blood_pressure() == {}, "推測して取り込まない"

    def test_non_bp_items_default_to_first_occurrence(self, tmp_path):
        path = make_v2_workbook(tmp_path / "occ1.xlsx", [
            person(items=[item("身体計測", "身長", "170.0", unit="cm", occurrence=None)]),
        ])
        result = parse_health_workbook(path)
        assert result.errors() == []
        assert result.persons[0].metrics[0].occurrence == 1

    def test_conflicting_values_same_occurrence(self, tmp_path):
        path = make_v2_workbook(tmp_path / "conflict.xlsx", [
            person(items=[
                item("血圧", "収縮期血圧", "120", unit="mmHg", occurrence=1),
                item("血圧", "収縮期血圧", "130", unit="mmHg", occurrence=1),
            ]),
        ])
        result = parse_health_workbook(path)
        assert "DUP_VALUE_CONFLICT" in codes(result, "error")

    def test_identical_duplicate_is_deduped(self, tmp_path):
        path = make_v2_workbook(tmp_path / "dup.xlsx", [
            person(items=[
                item("血圧", "収縮期血圧", "120", unit="mmHg", occurrence=1),
                item("血圧", "収縮期血圧", "120", unit="mmHg", occurrence=1),
            ]),
        ])
        result = parse_health_workbook(path)
        assert result.errors() == []
        assert len(result.persons[0].metrics) == 1


# ---------------------------------------------------------------------------
# 定性値
# ---------------------------------------------------------------------------

class TestQualitative:
    @pytest.mark.parametrize("raw", ["(-)", "（－）", "（−）", "( - )", "（ - ）",
                                     "(ー)", "陰性", "（陰性）"])
    def test_negative_variants_normalized(self, raw):
        value, issue = normalize_qualitative(raw)
        assert issue is None
        assert value == NEGATIVE

    @pytest.mark.parametrize("raw,expect", [
        ("(+)", "(+)"),
        ("（＋）", "(+)"),
        ("(±)", "(±)"),
        ("1+", "1+"),
        ("（２＋）", "(2+)"),
        ("3+", "3+"),
    ])
    def test_positive_values_preserved(self, raw, expect):
        value, issue = normalize_qualitative(raw)
        assert issue is None
        assert value == expect

    @pytest.mark.parametrize("raw", ["-", "－", "−", "ー", "--"])
    def test_bare_hyphen_is_not_negative(self, raw):
        value, issue = normalize_qualitative(raw)
        assert issue is not None and issue.code == "QUAL_BARE_HYPHEN"
        assert value != NEGATIVE

    @pytest.mark.parametrize("raw", ["", None, "   ", "　"])
    def test_blank_stays_blank(self, raw):
        value, issue = normalize_qualitative(raw)
        assert value == ""
        assert issue is None

    def test_bare_hyphen_blocks_generation(self, tmp_path):
        path = make_v2_workbook(tmp_path / "hyphen.xlsx", [
            person(items=[item("尿検査", "尿蛋白", "-", occurrence=1, value_type="定性")]),
        ])
        result = parse_health_workbook(path)
        assert "QUAL_BARE_HYPHEN" in codes(result, "error")

    def test_blank_qualitative_is_not_stored(self, tmp_path):
        path = make_v2_workbook(tmp_path / "blankqual.xlsx", [
            person(items=[item("尿検査", "尿糖", "", occurrence=1, value_type="定性")]),
        ])
        result = parse_health_workbook(path)
        assert result.errors() == []
        assert result.persons[0].qualitative() == []

    def test_needs_source_check_is_blocking(self, tmp_path):
        path = make_v2_workbook(tmp_path / "needs.xlsx", [
            person(items=[item("尿検査", "尿潜血", "(-)", occurrence=1,
                               value_type="要原票確認")]),
        ])
        result = parse_health_workbook(path)
        assert "NEEDS_SOURCE_CHECK" in codes(result, "error")

    def test_value_without_item_name_is_blocking(self, tmp_path):
        path = make_v2_workbook(tmp_path / "noitem.xlsx", [
            person(items=[item("尿検査", "", "(-)", occurrence=1, value_type="定性")]),
        ])
        result = parse_health_workbook(path)
        assert "QUAL_ATTRIBUTION_UNKNOWN" in codes(result, "error")


# ---------------------------------------------------------------------------
# セル値の読み取り
# ---------------------------------------------------------------------------

class TestCellValues:
    @pytest.mark.parametrize("raw,expect", [
        (89.3, "89.3"), (98.0, "98"), (98, "98"), ("89.3", "89.3"),
        ("120", "120"), (0.07, "0.07"), (1.10, "1.1"), ("1.10", "1.1"),
        (100.0, "100"), ("", ""), (None, ""), ("A", "A"),
        ("0301619", "0301619"),
    ])
    def test_format_cell_number(self, raw, expect):
        assert format_cell_number(raw) == expect

    @pytest.mark.parametrize("raw,expect", [
        ("132", "000132"), (132, "000132"), (132.0, "000132"),
        ("000132", "000132"), ("1", "000001"),
    ])
    def test_exam_no_zero_padded(self, raw, expect):
        value, issue = format_exam_no(raw)
        assert issue is None
        assert value == expect

    @pytest.mark.parametrize("raw", ["1234567", "abc", "", None])
    def test_exam_no_invalid(self, raw):
        _, issue = format_exam_no(raw)
        assert issue is not None and issue.level == "error"

    @pytest.mark.parametrize("raw,expect", [
        ("2026-07-01 00:00:00", date(2026, 7, 1)),
        ("2026-07-01", date(2026, 7, 1)),
        ("2026/7/1", date(2026, 7, 1)),
        ("20260701", date(2026, 7, 1)),
        (date(2026, 7, 1), date(2026, 7, 1)),
        ("", None), (None, None), ("なし", None),
    ])
    def test_parse_exam_date(self, raw, expect):
        assert parse_exam_date(raw) == expect

    def test_datetime_cell_for_exam_date(self, tmp_path):
        from datetime import datetime
        path = make_v2_workbook(tmp_path / "dt.xlsx", [
            person(exam_date=datetime(2026, 7, 2, 0, 0), items=[]),
        ])
        result = parse_health_workbook(path)
        assert result.persons[0].exam_date == date(2026, 7, 2)


# ---------------------------------------------------------------------------
# 構造の異常
# ---------------------------------------------------------------------------

class TestStructure:
    def test_unopenable_file_raises(self, tmp_path):
        bad = tmp_path / "broken.xlsx"
        bad.write_bytes(b"not an excel file")
        with pytest.raises(ValueError):
            parse_health_workbook(str(bad))

    def test_multiple_persons_are_separated(self, tmp_path):
        path = make_v2_workbook(tmp_path / "multi.xlsx", [
            person("友納 英彦", sheet="P02_友納英彦", exam_no="132",
                   items=bp_items(sys1="120", dia1="61")),
            person("高橋 和紀", age=47, sheet="P03_高橋和紀", exam_no="186",
                   items=bp_items(sys1="114", dia1="58")),
        ])
        result = parse_health_workbook(path)

        assert [p.name for p in result.persons] == ["友納 英彦", "高橋 和紀"]
        assert [p.exam_no for p in result.persons] == ["000132", "000186"]
        assert result.persons[0].blood_pressure()[1]["sys"] == "120"
        assert result.persons[1].blood_pressure()[1]["sys"] == "114"

    def test_judgement_is_kept_for_display_only(self, tmp_path):
        """原票判定は表示用に持つ。CSVへ書かないのは csv 側の責務。"""
        path = make_v2_workbook(tmp_path / "judge.xlsx", [
            person(items=bp_items(sys1="114", dia1="58", judgement="F")),
        ])
        result = parse_health_workbook(path)
        assert {m.source_judgement for m in result.persons[0].metrics} == {"F"}
