# -*- coding: utf-8 -*-
"""健診受診者 → jinjer社員 の照合テスト

自動確定してよいのは「氏名が一意 かつ 性別一致 かつ 受診日時点の年齢一致」の
ときだけ。ここを緩めると別人の健診結果をHPMへ入れる事故になる。
"""

from __future__ import annotations

import os
import sys
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.health_hpm_match import (  # noqa: E402
    STATUS_OK,
    STATUS_SELECT,
    JinjerCandidate,
    age_at,
    build_candidates,
    gender_matches,
    gender_to_hpm,
    is_company_employee,
    match_person,
    validate_selection,
)
from tests.health_hpm_fixtures import employees_stub  # noqa: E402


@pytest.fixture
def candidates():
    return build_candidates(employees_stub())


class TestBuildCandidates:
    def test_extracts_nested_fields(self, candidates):
        tomono = next(c for c in candidates if c.employee_id == "2018013")

        assert tomono.name == "友納　英彦", "漢字氏名は全角スペース区切り"
        assert tomono.kana == "トモノウ ヒデヒコ", "カナは半角スペース区切り"
        assert tomono.birth_date == date(1968, 4, 13)
        assert tomono.gender == "男性"

    def test_non_company_ids_excluded(self, candidates):
        ids = {c.employee_id for c in candidates}
        assert "5551234" not in ids, "派遣番号は候補にしない"
        assert "3333008" not in ids, "テスト番号は候補にしない"
        assert len(candidates) == 7

    def test_missing_birth_date_is_none(self, candidates):
        kuma = next(c for c in candidates if c.employee_id == "2023077")
        assert kuma.birth_date is None

    def test_empty_input(self):
        assert build_candidates([]) == []
        assert build_candidates(None) == []

    def test_broken_records_do_not_crash(self):
        result = build_candidates([{"id": "2018099"}, {}, {"id": None},
                                   {"id": "2018098", "personal": None}])
        assert {c.employee_id for c in result} == {"2018099", "2018098"}


class TestCompanyEmployee:
    @pytest.mark.parametrize("eid,expect", [
        ("2018013", True), ("2026006", True),
        ("5551234", False), ("3333008", False), ("201801", False),
        ("20180133", False), ("", False), (None, False), ("abc", False),
    ])
    def test_is_company_employee(self, eid, expect):
        assert is_company_employee(eid) is expect


class TestAgeAt:
    @pytest.mark.parametrize("birth,on,expect", [
        (date(1968, 4, 13), date(2026, 7, 1), 58),   # 誕生日後
        (date(1968, 7, 2), date(2026, 7, 1), 57),    # 誕生日前日
        (date(1968, 7, 1), date(2026, 7, 1), 58),    # 誕生日当日
        (date(1968, 12, 25), date(2026, 7, 1), 57),
        (None, date(2026, 7, 1), None),
        (date(1968, 4, 13), None, None),
    ])
    def test_age_at(self, birth, on, expect):
        assert age_at(birth, on) == expect


class TestGender:
    @pytest.mark.parametrize("excel,jinjer,expect", [
        ("男性", "男性", True), ("男", "男性", True), ("男性", "男", True),
        ("女性", "女性", True), ("女", "女性", True),
        ("男性", "女性", False), ("女", "男性", False),
        ("", "男性", False), ("男性", "", False), ("", "", False),
    ])
    def test_gender_matches(self, excel, jinjer, expect):
        assert gender_matches(excel, jinjer) is expect

    @pytest.mark.parametrize("jinjer,expect", [
        ("男性", "男"), ("女性", "女"), ("男", "男"), ("", ""), ("不明", ""),
    ])
    def test_gender_to_hpm(self, jinjer, expect):
        assert gender_to_hpm(jinjer) == expect


class TestMatchPerson:
    def test_auto_ok_when_everything_agrees(self, candidates):
        result = match_person("友納 英彦", "男性", 58, date(2026, 7, 1), candidates)

        assert result.status == STATUS_OK
        assert result.employee_id == "2018013"
        assert result.is_ok is True
        assert result.reasons == []

    def test_full_width_space_in_name_matches(self, candidates):
        result = match_person("友納　英彦", "男性", 58, date(2026, 7, 1), candidates)
        assert result.status == STATUS_OK

    def test_same_name_twice_needs_human(self, candidates):
        result = match_person("吉田 拓矢", "男性", 36, date(2026, 7, 1), candidates)

        assert result.status == STATUS_SELECT
        assert result.employee_id == ""
        assert {c.employee_id for c in result.candidates} == {"2022001", "2022002"}
        assert "同姓同名" in result.reasons[0]

    def test_gender_mismatch_needs_human(self, candidates):
        result = match_person("友納 英彦", "女性", 58, date(2026, 7, 1), candidates)

        assert result.status == STATUS_SELECT
        assert any("性別" in r for r in result.reasons)
        assert result.candidates[0].employee_id == "2018013"

    def test_age_mismatch_needs_human(self, candidates):
        result = match_person("友納 英彦", "男性", 40, date(2026, 7, 1), candidates)

        assert result.status == STATUS_SELECT
        assert any("年齢" in r for r in result.reasons)

    def test_missing_birth_date_needs_human(self, candidates):
        result = match_person("熊崎 俊輔", "男性", 41, date(2026, 7, 3), candidates)

        assert result.status == STATUS_SELECT
        assert any("生年月日" in r for r in result.reasons)

    def test_not_found_needs_human(self, candidates):
        result = match_person("存在 しない", "男性", 30, date(2026, 7, 1), candidates)

        assert result.status == STATUS_SELECT
        assert result.candidates == []
        assert "見つかりません" in result.reasons[0]

    def test_surname_only_offers_candidates(self, candidates):
        """姓だけの表記でも候補は出す（自動確定はしない）。"""
        result = match_person("吉田", "男性", 36, date(2026, 7, 1), candidates)

        assert result.status == STATUS_SELECT
        assert {c.employee_id for c in result.candidates} == {"2022001", "2022002"}

    def test_no_age_in_excel_needs_human(self, candidates):
        result = match_person("友納 英彦", "男性", None, date(2026, 7, 1), candidates)

        assert result.status == STATUS_SELECT
        assert any("年齢" in r for r in result.reasons)


class TestItaiji:
    """原票の異体字（髙橋）と jinjer の常用字体（高橋）を同じ人として扱う。"""

    def test_takahashi_matches(self, candidates):
        result = match_person("髙橋 和紀", "男性", 47, date(2026, 7, 1), candidates)

        assert result.status == STATUS_OK
        assert result.employee_id == "2019022"

    @pytest.mark.parametrize("raw,folded", [
        ("髙橋", "高橋"), ("川﨑", "川崎"), ("濵田", "浜田"),
        ("渡邊", "渡辺"), ("渡邉", "渡辺"), ("齋藤", "斎藤"),
        ("高橋", "高橋"),
    ])
    def test_fold_itaiji(self, raw, folded):
        from services.health_hpm_match import fold_itaiji
        assert fold_itaiji(raw) == folded

    def test_display_name_stays_as_registered(self, candidates):
        """照合だけ寄せる。候補に出る氏名は jinjer の登録どおり。"""
        result = match_person("髙橋 和紀", "男性", 47, date(2026, 7, 1), candidates)
        assert result.candidates[0].name == "高橋　和紀"

    def test_unrelated_names_not_merged(self, candidates):
        result = match_person("友納 英彦", "男性", 58, date(2026, 7, 1), candidates)
        assert result.employee_id == "2018013", "無関係な氏名まで畳まない"

    def test_itaiji_still_checks_gender_and_age(self, candidates):
        result = match_person("髙橋 和紀", "女性", 47, date(2026, 7, 1), candidates)
        assert result.status == STATUS_SELECT
        assert any("性別" in r for r in result.reasons)


class TestValidateSelection:
    def test_accepts_known_employee(self, candidates):
        got = validate_selection("2018013", candidates)
        assert got.employee_id == "2018013"

    @pytest.mark.parametrize("eid,expect_in_message", [
        ("", "選ばれていません"),
        (None, "選ばれていません"),
        ("5551234", "自社の形式"),
        ("2099999", "在籍者一覧"),
    ])
    def test_rejects_bad_selection(self, candidates, eid, expect_in_message):
        with pytest.raises(ValueError) as e:
            validate_selection(eid, candidates)
        assert expect_in_message in str(e.value)


class TestCandidateSerialization:
    def test_as_dict_for_screen(self):
        c = JinjerCandidate("2018013", "友納", "英彦", "トモノウ", "ヒデヒコ",
                            date(1968, 4, 13), "男性")
        assert c.as_dict() == {
            "employee_id": "2018013",
            "name": "友納　英彦",
            "kana": "トモノウ ヒデヒコ",
            "birth_date": "1968-04-13",
            "gender": "男性",
        }

    def test_as_dict_without_birth(self):
        c = JinjerCandidate("2018013", "友納", "英彦")
        assert c.as_dict()["birth_date"] == ""
