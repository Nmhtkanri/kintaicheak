# -*- coding: utf-8 -*-
"""HPM取込用302列CSVの組み立て・書き出しテスト

このファイルが守っているのは要件の「絶対ルール」そのもの:
  - 血圧の平均を作らない／1回分を2回目へ複製しない
  - (-) は陰性として出す。空欄を勝手に (-) にしない
  - 原票判定A〜GはHPMへ転記しない
  - CP932・302列・受診番号6桁ゼロ埋め・同名ファイル上書き禁止
  - 書いた後にバイトで読み直して検証する（行消失・ゼロ落ちの実事故対策）
"""

from __future__ import annotations

import csv
import io
import os
import sys
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.health_hpm_csv import (  # noqa: E402
    COL_AGE,
    COL_BIRTH,
    COL_COURSE,
    COL_EXAM_DATE,
    COL_EXAM_NO,
    COL_GENDER,
    COL_KANA,
    COL_LOCATION,
    COL_NAME,
    COL_VENUE,
    TOTAL_COLS,
    OutputExistsError,
    build_csv_rows,
    build_person_row,
    check_cp932,
    default_output_dir,
    default_output_filename,
    fiscal_year,
    mixed_fiscal_years,
    verify_written_csv,
    write_hpm_csv,
    zen_to_han_kana,
)
from services.health_hpm_excel import HealthMetric, PersonRecord  # noqa: E402
from services.health_hpm_master import find_course, load_master, resolve_institution  # noqa: E402
from services.health_hpm_match import JinjerCandidate  # noqa: E402
from tests.health_hpm_fixtures import make_master_xlsx  # noqa: E402

DOYUKAI = "医療法人社団 同友会 春日クリニック"
COURSE_VALUE = "人間ドックＣ　胃カメラ（４０歳以上）"


@pytest.fixture
def master(tmp_path):
    return load_master(make_master_xlsx(tmp_path / "master.xlsx"))


def metric(category, item, value, occurrence=1, value_type="数値",
           judgement="", original=""):
    return HealthMetric(category=category, item=item, occurrence=occurrence,
                        value=value, value_type=value_type,
                        source_judgement=judgement,
                        original_display=original or value)


def make_person(metrics=(), *, name="友納 英彦", exam_no="000132",
                exam_date=date(2026, 7, 1)):
    return PersonRecord(key="p01", name=name, age=58, gender="男性",
                        exam_date=exam_date, exam_no=exam_no,
                        sheet="P02", metrics=list(metrics))


def make_employee(**kwargs):
    base = dict(employee_id="2018013", last_name="友納", first_name="英彦",
                last_kana="トモノウ", first_kana="ヒデヒコ",
                birth_date=date(1968, 4, 13), gender="男性")
    base.update(kwargs)
    return JinjerCandidate(**base)


def build_row(master, metrics=(), person=None, employee=None):
    person = person or make_person(metrics)
    employee = employee or make_employee()
    institution = resolve_institution(master, DOYUKAI)
    course = find_course(master, DOYUKAI, COURSE_VALUE)
    return build_person_row(person, employee, course, institution, master)


# ---------------------------------------------------------------------------
# 血圧（最重要）
# ---------------------------------------------------------------------------

class TestBloodPressure:
    def test_two_rounds_pass_through_untouched(self, master):
        """1回目 132/86・2回目 118/72 が列50〜53へそのまま出る。"""
        row, issues = build_row(master, [
            metric("血圧", "収縮期血圧", "132", occurrence=1),
            metric("血圧", "拡張期血圧", "86", occurrence=1),
            metric("血圧", "収縮期血圧", "118", occurrence=2),
            metric("血圧", "拡張期血圧", "72", occurrence=2),
        ])

        assert row[50] == "132"
        assert row[51] == "86"
        assert row[52] == "118"
        assert row[53] == "72"

    def test_no_average_anywhere_in_the_row(self, master):
        """平均（125/79）が302セルのどこにも現れないこと。"""
        row, _ = build_row(master, [
            metric("血圧", "収縮期血圧", "132", occurrence=1),
            metric("血圧", "拡張期血圧", "86", occurrence=1),
            metric("血圧", "収縮期血圧", "118", occurrence=2),
            metric("血圧", "拡張期血圧", "72", occurrence=2),
        ])
        # 132と118の平均=125、86と72の平均=79。どちらも他の値と衝突しない値を選んである
        assert "125" not in row
        assert "79" not in row
        assert "125.0" not in row and "79.0" not in row

    def test_single_round_leaves_second_blank(self, master):
        """1回分しかないとき、2回目の列は空欄のまま（複製しない）。"""
        row, _ = build_row(master, [
            metric("血圧", "収縮期血圧", "120", occurrence=1),
            metric("血圧", "拡張期血圧", "61", occurrence=1),
        ])

        assert row[50] == "120"
        assert row[51] == "61"
        assert row[52] == "", "1回目を2回目へ複製しない"
        assert row[53] == ""

    def test_third_round_not_written(self, master):
        """3回目があっても列50〜53は1回目・2回目のまま。"""
        row, _ = build_row(master, [
            metric("血圧", "収縮期血圧", "132", occurrence=1),
            metric("血圧", "拡張期血圧", "86", occurrence=1),
            metric("血圧", "収縮期血圧", "118", occurrence=2),
            metric("血圧", "拡張期血圧", "72", occurrence=2),
            metric("血圧", "収縮期血圧", "111", occurrence=3),
            metric("血圧", "拡張期血圧", "66", occurrence=3),
        ])

        assert [row[50], row[51], row[52], row[53]] == ["132", "86", "118", "72"]
        assert "111" not in row, "3回目の値はどこにも出さない"
        assert "66" not in row

    def test_only_second_round_present(self, master):
        """2回目だけある場合、1回目は空欄のまま（繰り上げない）。"""
        row, _ = build_row(master, [
            metric("血圧", "収縮期血圧", "118", occurrence=2),
            metric("血圧", "拡張期血圧", "72", occurrence=2),
        ])

        assert row[50] == "" and row[51] == ""
        assert row[52] == "118" and row[53] == "72"


# ---------------------------------------------------------------------------
# 定性検査
# ---------------------------------------------------------------------------

class TestQualitative:
    def test_urine_protein_and_sugar_columns(self, master):
        row, _ = build_row(master, [
            metric("尿検査", "尿蛋白", "(-)", value_type="定性"),
            metric("尿検査", "尿糖", "(-)", value_type="定性"),
        ])
        assert row[54] == "(-)"
        assert row[55] == "(-)"

    def test_fecal_blood_split_by_occurrence(self, master):
        row, _ = build_row(master, [
            metric("便潜血", "便潜血", "(-)", occurrence=1, value_type="定性"),
            metric("便潜血", "便潜血", "(+)", occurrence=2, value_type="定性"),
        ])
        assert row[70] == "(-)"
        assert row[71] == "(+)"

    def test_blank_qualitative_column_stays_blank(self, master):
        row, _ = build_row(master, [
            metric("尿検査", "尿蛋白", "(-)", value_type="定性"),
        ])
        assert row[55] == "", "値が無い項目は空欄のまま（(-)にしない）"

    def test_item_name_with_method_needs_a_master_row(self, master):
        """項目名に方式を埋め込んだ表記は、マスタに行が無ければ出力せず警告する。

        近い列へ寄せるくらいなら出さない、という方針をここで固定する。
        """
        row, issues = build_row(master, [
            metric("感染症", "HBs抗原（CLIA）", "(-)", value_type="定性"),
        ])

        assert row[115] == "" and row[117] == ""
        assert any(i.code == "UNMAPPED_ITEM" for i in issues)

    def test_method_unknown_is_warned_and_left_blank(self, master):
        row, issues = build_row(master, [
            metric("感染症", "HBs抗原", "(-)", value_type="定性"),
        ])

        assert row[115] == "" and row[117] == "", "方式不明なら近い列へ推測しない"
        assert any(i.code == "METHOD_UNKNOWN" for i in issues)

    def test_method_matched_from_original_display(self, master):
        row, issues = build_row(master, [
            metric("感染症", "HBs抗原", "(-)", value_type="定性",
                   original="HBs抗原 CLIA (-)"),
        ])
        assert row[117] == "(-)"
        assert row[115] == ""
        assert not any(i.code == "METHOD_UNKNOWN" for i in issues)


# ---------------------------------------------------------------------------
# 判定列
# ---------------------------------------------------------------------------

class TestJudgement:
    def test_source_judgement_never_written(self, master):
        row, _ = build_row(master, [
            metric("血圧", "収縮期血圧", "114", occurrence=1, judgement="F"),
            metric("血圧", "拡張期血圧", "58", occurrence=1, judgement="F"),
            metric("身体計測", "BMI", "35.2", judgement="D"),
        ])

        for col in range(183, 198):
            assert row[col] == "", f"判定列 {col} は空欄でなければならない"
        assert "F" not in row and "D" not in row


# ---------------------------------------------------------------------------
# 識別列
# ---------------------------------------------------------------------------

class TestIdentityColumns:
    def test_identity_columns_from_jinjer(self, master):
        row, _ = build_row(master)

        assert row[COL_NAME] == "友納　英彦", "漢字氏名は全角スペース区切り"
        assert row[COL_KANA] == "ﾄﾓﾉｳ ﾋﾃﾞﾋｺ", "カナは半角カナ・半角スペース"
        assert row[COL_BIRTH] == "19680413"
        assert row[COL_GENDER] == "男"
        assert row[COL_AGE] == "58"
        assert row[COL_COURSE] == COURSE_VALUE
        assert row[COL_EXAM_DATE] == "20260701"
        assert row[COL_EXAM_NO] == "000132", "受診番号は6桁ゼロ埋めのまま"
        assert row[COL_VENUE] == "2"
        assert row[COL_LOCATION] == "1310528885"

    def test_age_is_computed_from_birth_date_not_excel(self, master):
        """Excelの年齢欄ではなく生年月日から出す（出典を1つに固定）。"""
        person = make_person([], exam_date=date(2026, 4, 1))
        person.age = 99  # Excel側が違っていても引きずられない
        row, _ = build_row(master, person=person)
        assert row[COL_AGE] == "57"

    def test_leading_zero_location_code_preserved(self, master):
        institution = resolve_institution(master, "医療法人徳洲会 生駒市立病院")
        course = find_course(master, "医療法人徳洲会 生駒市立病院", "13")
        row, _ = build_person_row(make_person(), make_employee(), course,
                                  institution, master)
        assert row[COL_LOCATION] == "0301619"

    def test_row_always_has_302_columns(self, master):
        row, _ = build_row(master, [metric("身体計測", "身長", "181.1")])
        assert len(row) == TOTAL_COLS


class TestZenToHanKana:
    @pytest.mark.parametrize("full,half", [
        ("トモノウ ヒデヒコ", "ﾄﾓﾉｳ ﾋﾃﾞﾋｺ"),
        ("オオツボ ジュンイチ", "ｵｵﾂﾎﾞ ｼﾞｭﾝｲﾁ"),
        ("タカハシ カズノリ", "ﾀｶﾊｼ ｶｽﾞﾉﾘ"),
        ("ヤマナガ ジン", "ﾔﾏﾅｶﾞ ｼﾞﾝ"),
        ("ガギグゲゴ", "ｶﾞｷﾞｸﾞｹﾞｺﾞ"),
        ("パピプペポ", "ﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ"),
        ("トモノウ　ヒデヒコ", "ﾄﾓﾉｳ ﾋﾃﾞﾋｺ"),   # 全角スペースも半角に
        ("ながやま", "ﾅｶﾞﾔﾏ"),                   # ひらがなもカタカナ経由で半角に
        ("ヴァイオリン", "ｳﾞｧｲｵﾘﾝ"),
        ("", ""),
    ])
    def test_conversion(self, full, half):
        assert zen_to_han_kana(full) == half


# ---------------------------------------------------------------------------
# マッピング外
# ---------------------------------------------------------------------------

class TestUnmapped:
    def test_unknown_item_is_warned_not_guessed(self, master):
        row, issues = build_row(master, [
            metric("謎の検査", "知らない項目", "1.23"),
        ])

        assert any(i.code == "UNMAPPED_ITEM" for i in issues)
        assert "1.23" not in row, "マッピングに無い値は出力しない"


# ---------------------------------------------------------------------------
# 書き出し
# ---------------------------------------------------------------------------

class TestWrite:
    def _rows(self, master, metrics=()):
        institution = resolve_institution(master, DOYUKAI)
        course = find_course(master, DOYUKAI, COURSE_VALUE)
        return build_csv_rows(
            [(make_person(metrics), make_employee(), course, institution)], master)

    def test_bytes_are_cp932_crlf_without_bom(self, master, tmp_path):
        rows, _ = self._rows(master, [metric("血圧", "収縮期血圧", "120", occurrence=1)])
        out = tmp_path / "out.csv"
        write_hpm_csv(str(out), rows)

        raw = out.read_bytes()
        assert raw[:3] != b"\xef\xbb\xbf", "BOMを付けない"
        assert raw.endswith(b"\r\n")
        assert raw.count(b"\r\n") == 2, "ヘッダー1行 + データ1行"
        text = raw.decode("cp932")
        assert '"' not in text, "QUOTE_MINIMAL なので引用符は出ないはず"

    def test_round_trip_keeps_negative_token(self, master, tmp_path):
        rows, _ = self._rows(master, [
            metric("尿検査", "尿蛋白", "(-)", value_type="定性"),
        ])
        out = tmp_path / "neg.csv"
        write_hpm_csv(str(out), rows)

        text = out.read_bytes().decode("cp932")
        back = list(csv.reader(io.StringIO(text, newline="")))
        assert back[1][54] == "(-)", "(-) がCP932往復で変わらない"

    def test_exam_no_zero_padding_survives_the_file(self, master, tmp_path):
        rows, _ = self._rows(master)
        out = tmp_path / "zero.csv"
        write_hpm_csv(str(out), rows)

        back = list(csv.reader(io.StringIO(out.read_bytes().decode("cp932"), newline="")))
        assert back[1][COL_EXAM_NO] == "000132"

    def test_refuses_to_overwrite(self, master, tmp_path):
        rows, _ = self._rows(master)
        out = tmp_path / "dup.csv"
        write_hpm_csv(str(out), rows)
        before = out.read_bytes()

        with pytest.raises(OutputExistsError):
            write_hpm_csv(str(out), rows)
        assert out.read_bytes() == before, "既存ファイルを壊さない"

    def test_wrong_column_count_refused(self, master, tmp_path):
        rows, _ = self._rows(master)
        rows[1] = rows[1][:-1]
        with pytest.raises(ValueError):
            write_hpm_csv(str(tmp_path / "short.csv"), rows)

    def test_unencodable_char_leaves_no_file(self, master, tmp_path):
        """CP932にできない文字があるとき、中途半端なファイルを残さない。"""
        rows, _ = self._rows(master)
        rows[1][24] = "1.5㎗"
        out = tmp_path / "bad.csv"

        with pytest.raises(UnicodeEncodeError):
            write_hpm_csv(str(out), rows)
        assert not out.exists(), "書き出し前に落ちるのでファイルはできない"

    def test_check_cp932_reports_before_writing(self, master):
        rows, _ = self._rows(master)
        rows[1][24] = "1.5㎗"
        issues = check_cp932(rows)

        assert len(issues) == 1
        assert issues[0].code == "CP932_UNENCODABLE"
        assert "U+3397" in issues[0].message

    @pytest.mark.parametrize("value", ["±0.5", "Ⅲ", "μ", "～", "㎎", "(-)"])
    def test_encodable_values_pass(self, master, value):
        rows, _ = self._rows(master)
        rows[1][24] = value
        assert check_cp932(rows) == []


class TestVerify:
    def _write(self, master, tmp_path, name="v.csv"):
        institution = resolve_institution(master, DOYUKAI)
        course = find_course(master, DOYUKAI, COURSE_VALUE)
        rows, _ = build_csv_rows([
            (make_person([metric("血圧", "収縮期血圧", "120", occurrence=1)]),
             make_employee(), course, institution),
            (make_person([], name="高橋 和紀", exam_no="000186"),
             make_employee(employee_id="2019022", last_name="高橋", first_name="和紀",
                           last_kana="タカハシ", first_kana="カズノリ",
                           birth_date=date(1978, 12, 25)),
             course, institution),
        ], master)
        out = tmp_path / name
        write_hpm_csv(str(out), rows)
        return out, rows

    def test_verify_passes_for_intact_file(self, master, tmp_path):
        out, rows = self._write(master, tmp_path)
        assert verify_written_csv(str(out), rows) == []

    def test_verify_detects_row_loss(self, master, tmp_path):
        """Excelで開き直して行が消えた状態を検出できること（実際に起きた事故）。"""
        out, rows = self._write(master, tmp_path)
        text = out.read_bytes().decode("cp932")
        lines = text.split("\r\n")
        out.write_bytes("\r\n".join(lines[:-2] + [""]).encode("cp932"))

        problems = verify_written_csv(str(out), rows)
        assert any("行数が違います" in p for p in problems)

    def test_verify_detects_zero_padding_loss(self, master, tmp_path):
        """受診番号 000132 が 132 に化けたのを検出できること。"""
        out, rows = self._write(master, tmp_path)
        text = out.read_bytes().decode("cp932")
        out.write_bytes(text.replace(",000132,", ",132,").encode("cp932"))

        problems = verify_written_csv(str(out), rows)
        assert any("000132" in p and "132" in p for p in problems)

    def test_verify_detects_bom(self, master, tmp_path):
        out, rows = self._write(master, tmp_path)
        out.write_bytes(b"\xef\xbb\xbf" + out.read_bytes())
        assert any("BOM" in p for p in verify_written_csv(str(out), rows))

    def test_verify_detects_changed_blood_pressure(self, master, tmp_path):
        out, rows = self._write(master, tmp_path)
        text = out.read_bytes().decode("cp932")
        out.write_bytes(text.replace(",120,", ",125,").encode("cp932"))

        problems = verify_written_csv(str(out), rows)
        assert any("120" in p for p in problems)


# ---------------------------------------------------------------------------
# 保存先
# ---------------------------------------------------------------------------

class TestOutputLocation:
    @pytest.mark.parametrize("d,expect", [
        (date(2026, 4, 1), 2026),
        (date(2026, 7, 1), 2026),
        (date(2027, 3, 31), 2026),
        (date(2027, 4, 1), 2027),
        (date(2026, 1, 15), 2025),
    ])
    def test_fiscal_year(self, d, expect):
        assert fiscal_year(d) == expect

    def test_default_output_dir(self):
        got = default_output_dir([date(2026, 7, 1), date(2026, 7, 3)], r"Z:\健康診断")
        assert got == os.path.join(r"Z:\健康診断", "2026", "2026年度健康診断受診者結果",
                                   "CSV格納")

    def test_mixed_fiscal_years_detected(self):
        assert mixed_fiscal_years([date(2026, 7, 1), date(2027, 5, 1)]) == [2026, 2027]
        assert mixed_fiscal_years([date(2026, 7, 1), date(2026, 8, 1)]) == [2026]

    def test_default_filename_matches_existing_convention(self, master):
        institution = resolve_institution(master, DOYUKAI)
        name = default_output_filename(
            master, [institution], [date(2026, 7, 1), date(2026, 7, 3)], 6)
        assert name == "HPM取込用_同友会_20260701-0703_6名.csv"

    def test_default_filename_single_day(self, master):
        institution = resolve_institution(master, DOYUKAI)
        name = default_output_filename(master, [institution], [date(2026, 7, 1)], 1)
        assert name == "HPM取込用_同友会_20260701_1名.csv"

    def test_default_filename_multiple_institutions(self, master):
        a = resolve_institution(master, DOYUKAI)
        b = resolve_institution(master, "医療法人徳洲会 生駒市立病院")
        name = default_output_filename(master, [a, b], [date(2026, 5, 2)], 4)
        assert "複数機関" in name
