# -*- coding: utf-8 -*-
"""健康診断HPM: 健診申込「選択肢」シートの機関・種別・追加検査を追記する合成ロジック。

変換マスタが正で、シートは追記。重複させない・その他は載せない・追加検査はCSVに
書かない、をここで固定する。
"""

from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.health_apply import schema as S  # noqa: E402
from services.health_apply.options import OptionCatalog  # noqa: E402
from services.health_hpm_master import load_master  # noqa: E402
from services.health_hpm_options import (  # noqa: E402
    INACTIVE_SUFFIX,
    SheetOptions,
    default_extra_codes,
    extras_issue,
    fallback_options,
    from_catalog,
    merged_institutions,
    resolve_course_choice,
    resolve_extras,
    resolve_institution_choice,
)
from tests.health_apply_fixtures import options_rows  # noqa: E402
from tests.health_hpm_fixtures import make_master_xlsx  # noqa: E402

DOYUKAI = "医療法人社団 同友会 春日クリニック"
OTEMACHI = "MYメディカルクリニック 大手町"
SHINAGAWA = "東京品川病院 総合健診センター"   # options_rows では 有効=0


def catalog() -> OptionCatalog:
    rows = options_rows()
    return OptionCatalog.from_rows(S.rows_to_dicts(rows[0], rows[1:]))


@pytest.fixture
def master(tmp_path):
    return load_master(make_master_xlsx(tmp_path / "master.xlsx"))


@pytest.fixture
def options() -> SheetOptions:
    return from_catalog(catalog(), source="テスト選択肢")


class TestFromCatalog:
    def test_other_is_dropped_and_inactive_sorted_last(self, options):
        codes = [i.code for i in options.institutions]
        assert "OTHER" not in codes
        assert codes[-1] == "130192", "無効の機関は末尾"
        assert options.institution_by_code("130192").active is False
        assert options.loaded is True
        assert options.source == "テスト選択肢"

    def test_exam_types_and_extras(self, options):
        assert [e.code for e in options.exam_types] == ["10", "11", "12", "13", "14", "15"]
        assert [(e.code, e.name) for e in options.extras] == [("GYN", "婦人科検診")]

    def test_extras_name_normalized_from_legacy(self):
        rows = [list(S.OPTION_HEADERS), ["追加検査", "GYN", "婦人病検査", "1", "1", "", ""]]
        cat = OptionCatalog.from_rows(S.rows_to_dicts(rows[0], rows[1:]))
        assert from_catalog(cat).extras[0].name == "婦人科検診"

    def test_no_extras_falls_back_to_gyn(self):
        rows = [list(S.OPTION_HEADERS), ["種別", "10", "定期健康診断", "1", "1", "", ""]]
        cat = OptionCatalog.from_rows(S.rows_to_dicts(rows[0], rows[1:]))
        assert [e.code for e in from_catalog(cat).extras] == ["GYN"]

    def test_fallback_keeps_gyn_and_reason(self):
        fb = fallback_options("鍵がありません")
        assert fb.loaded is False
        assert fb.error == "鍵がありません"
        assert fb.institutions == [] and fb.exam_types == []
        assert [e.code for e in fb.extras] == ["GYN"]
        assert fb.as_dict()["counts"] == {"institutions": 0, "exam_types": 0, "extras": 1}
        assert fb.as_dict()["exam_types"] == []

    def test_as_dict_exposes_exam_types(self, options):
        types = options.as_dict()["exam_types"]
        assert [t["code"] for t in types] == ["10", "11", "12", "13", "14", "15"]
        assert types[0] == {"code": "10", "name": "定期健康診断", "active": True}


class TestMergedInstitutions:
    def test_master_first_then_sheet_without_duplicates(self, master, options):
        merged = merged_institutions(master, options)
        sources = [m["source"] for m in merged]
        assert sources[:3] == ["master"] * 3 and set(sources[3:]) == {"sheet"}

        names = [m["name"] for m in merged]
        assert names.count(DOYUKAI) == 1, "同じ場所コードの機関はシート側を載せない"
        assert "医療法人徳洲会 生駒市立病院" in names and names.count("医療法人徳洲会 生駒市立病院") == 1
        assert OTEMACHI in names
        assert SHINAGAWA + INACTIVE_SUFFIX in names, "無効機関も選べるが表示名で分かる"
        assert "その他" not in names

    def test_sheet_institution_value_is_code(self, master, options):
        merged = merged_institutions(master, options)
        otemachi = next(m for m in merged if m["name"] == OTEMACHI)
        assert otemachi["value"] == "1310136358"
        assert otemachi["location_code"] == "1310136358"
        assert otemachi["hpm_confirmed"] is True
        doyukai = next(m for m in merged if m["name"] == DOYUKAI)
        assert doyukai["value"] == DOYUKAI, "変換マスタの機関は機関名のまま"

    def test_courses_master_then_generic(self, master, options):
        merged = merged_institutions(master, options)
        doyukai = next(m for m in merged if m["name"] == DOYUKAI)
        values = [c["hpm_value"] for c in doyukai["courses"]]
        assert values[:2] == ["人間ドックＣ　胃カメラ（４０歳以上）", "10"], "マスタのコースが先"
        assert values.count("10") == 1, "同じHPM出力値は追記しない"
        assert values[2:] == ["11", "12", "13", "14", "15"]
        assert doyukai["courses"][0]["source"] == "master"
        assert doyukai["courses"][-1]["source"] == "sheet"

        otemachi = next(m for m in merged if m["name"] == OTEMACHI)
        assert [c["hpm_value"] for c in otemachi["courses"]] == ["10", "11", "12", "13", "14", "15"]

    def test_fallback_gives_master_only(self, master):
        merged = merged_institutions(master, fallback_options("x"))
        assert [m["source"] for m in merged] == ["master"] * 3
        doyukai = next(m for m in merged if m["name"] == DOYUKAI)
        assert [c["hpm_value"] for c in doyukai["courses"]] == ["人間ドックＣ　胃カメラ（４０歳以上）", "10"]


class TestResolve:
    def test_master_name_and_alias_win(self, master, options):
        assert resolve_institution_choice(master, options, DOYUKAI).location_code == "1310528885"
        assert resolve_institution_choice(master, options, "同友会").location_code == "1310528885"

    def test_sheet_by_code_and_by_name(self, master, options):
        by_code = resolve_institution_choice(master, options, "1310136358")
        assert by_code is not None and by_code.name == OTEMACHI and by_code.hpm_confirmed
        assert by_code.location_code == "1310136358"
        by_name = resolve_institution_choice(master, options, OTEMACHI)
        assert by_name.location_code == "1310136358"
        with_suffix = resolve_institution_choice(master, options, SHINAGAWA + INACTIVE_SUFFIX)
        assert with_suffix.location_code == "130192"

    @pytest.mark.parametrize("value", ["", "  ", "存在しない", "OTHER", "その他"])
    def test_unknown_is_none(self, master, options, value):
        assert resolve_institution_choice(master, options, value) is None

    def test_course_master_then_sheet(self, master, options):
        doyukai = resolve_institution_choice(master, options, DOYUKAI)
        text = resolve_course_choice(master, options, doyukai, "人間ドックＣ　胃カメラ（４０歳以上）")
        assert text.display_name == "人間ドックC 胃カメラ（40歳以上）"
        generic = resolve_course_choice(master, options, doyukai, "12")
        assert generic.hpm_value == "12" and generic.display_name == "人間ドックB"
        assert generic.institution == DOYUKAI

        otemachi = resolve_institution_choice(master, options, "1310136358")
        assert resolve_course_choice(master, options, otemachi, "13").hpm_value == "13"
        assert resolve_course_choice(master, options, otemachi, "99") is None
        assert resolve_course_choice(master, options, otemachi, "") is None

    def test_course_without_sheet_only_master(self, master):
        fb = fallback_options("x")
        doyukai = resolve_institution_choice(master, fb, DOYUKAI)
        assert resolve_course_choice(master, fb, doyukai, "10").hpm_value == "10"
        assert resolve_course_choice(master, fb, doyukai, "12") is None

    def test_extras(self, options):
        found, unknown = resolve_extras(options, ["GYN", "GYN", "", "XXX"])
        assert [e.code for e in found] == ["GYN"]
        assert unknown == ["XXX"]
        assert resolve_extras(options, None) == ([], [])

    def test_default_extra_codes_female_only(self, options):
        assert default_extra_codes(options, "女性") == ["GYN"]
        assert default_extra_codes(options, "男性", "女性") == ["GYN"]
        assert default_extra_codes(options, "男性") == []
        assert default_extra_codes(options, "", None) == []

    def test_extras_issue_lists_people(self, options):
        gyn = options.extras
        issue = extras_issue([("友納 英彦", []), ("斉藤 ひとみ", gyn)])
        assert issue.level == "warning" and issue.code == "EXTRA_NOT_IN_CSV"
        assert "斉藤 ひとみ（婦人科検診）" in issue.message
        assert "友納" not in issue.message
        assert "CSV には出していません" in issue.message
        assert extras_issue([("友納 英彦", [])]) is None
