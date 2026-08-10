# -*- coding: utf-8 -*-
"""健康診断HPM変換マスタの読み込みテスト

マスタは人が Excel で直す。直し間違えたまま出力すると、検査値が別の欄に
入ったCSVができて人には気付けない。そのため「少しでも変なら止まる」ことを
ここで固定する。特に血圧の 50〜53 は動かせないこと。
"""

from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.health_hpm_master import (  # noqa: E402
    BP_EXPECTED_COLS,
    MasterError,
    courses_of,
    find_course,
    load_master,
    resolve_institution,
)
from tests.health_hpm_fixtures import (  # noqa: E402
    DEFAULT_ITEM_MAP,
    canonical_header,
    make_master_xlsx,
)


class TestLoadHappyPath:
    def test_loads_all_sheets(self, tmp_path):
        master = load_master(make_master_xlsx(tmp_path / "m.xlsx"))

        assert len(master.header) == 302
        assert master.venue_code == "2"
        assert len(master.institutions) == 3
        assert len(master.item_map) == len(DEFAULT_ITEM_MAP)
        assert master.courses["医療法人社団 同友会 春日クリニック"]

    def test_location_code_keeps_leading_zero(self, tmp_path):
        master = load_master(make_master_xlsx(tmp_path / "m.xlsx"))
        institution = master.institutions["医療法人徳洲会 生駒市立病院"]

        assert institution.location_code == "0301619", "前ゼロを数値化して落とさない"
        assert isinstance(institution.location_code, str)

    def test_alias_resolution(self, tmp_path):
        master = load_master(make_master_xlsx(tmp_path / "m.xlsx"))

        assert resolve_institution(master, "同友会").location_code == "1310528885"
        assert resolve_institution(master, "医療法人社団 同友会 春日クリニック") is not None
        assert resolve_institution(master, "知らない病院") is None
        assert resolve_institution(master, "") is None

    def test_hpm_confirmed_flag(self, tmp_path):
        master = load_master(make_master_xlsx(tmp_path / "m.xlsx"))

        assert master.institutions["医療法人社団 同友会 春日クリニック"].hpm_confirmed is True
        assert master.institutions["未確認クリニック"].hpm_confirmed is False

    def test_courses_and_find(self, tmp_path):
        master = load_master(make_master_xlsx(tmp_path / "m.xlsx"))

        courses = courses_of(master, "同友会")  # 別名でも引ける
        assert len(courses) == 2
        found = find_course(master, "同友会", "人間ドックＣ　胃カメラ（４０歳以上）")
        assert found is not None and found.display_name.startswith("人間ドックC")
        assert find_course(master, "同友会", "存在しない値") is None

    def test_rules_for_returns_method_variants(self, tmp_path):
        master = load_master(make_master_xlsx(tmp_path / "m.xlsx"))

        rules = master.rules_for("感染症", "HBs抗原", 1)
        assert {r.method for r in rules} == {"MAT", "CLIA"}
        assert {r.hpm_col for r in rules} == {115, 117}

    def test_blood_pressure_columns_are_fixed(self, tmp_path):
        master = load_master(make_master_xlsx(tmp_path / "m.xlsx"))
        actual = {(r.item, r.occurrence): r.hpm_col for r in master.item_map
                  if r.category == "血圧"}
        assert actual == BP_EXPECTED_COLS


class TestBrokenMaster:
    """壊れたマスタは必ず MasterError で止まる（部分的に読んで進めない）。"""

    @pytest.mark.parametrize("break_rule,expect_in_message", [
        ("bp_shift", "52"),
        ("missing_bp2", "2回目"),
        ("bp_header_name", "血圧（二回目）最高"),
        ("header_301", "302"),
        ("map_judgement", "判定列"),
        ("map_identity", "識別列"),
        ("dup_key", "両方"),
        ("dup_col", "両方"),
        ("no_venue", "会場コード"),
        ("alias_orphan", "存在しないクリニック"),
        ("course_orphan", "存在しないクリニック"),
        ("col_out_of_range", "範囲外"),
        ("missing_sheet", "機関別名"),
    ])
    def test_broken_master_raises(self, tmp_path, break_rule, expect_in_message):
        path = make_master_xlsx(tmp_path / f"{break_rule}.xlsx", break_rule=break_rule)
        with pytest.raises(MasterError) as e:
            load_master(path)
        assert expect_in_message in str(e.value)

    def test_bp_shift_message_explains_why(self, tmp_path):
        """列ズレは「平均のような値になる」と理由まで出す。"""
        path = make_master_xlsx(tmp_path / "shift.xlsx", break_rule="bp_shift")
        with pytest.raises(MasterError) as e:
            load_master(path)
        assert "平均" in str(e.value)

    def test_missing_file(self, tmp_path):
        with pytest.raises(MasterError) as e:
            load_master(str(tmp_path / "ない.xlsx"))
        assert "見つかりません" in str(e.value)

    def test_not_an_excel_file(self, tmp_path):
        bad = tmp_path / "bad.xlsx"
        bad.write_bytes("これはExcelではない".encode("utf-8"))
        with pytest.raises(MasterError):
            load_master(str(bad))

    def test_empty_institutions(self, tmp_path):
        path = make_master_xlsx(tmp_path / "noinst.xlsx", institutions=[], aliases=[],
                                courses=[])
        with pytest.raises(MasterError) as e:
            load_master(path)
        assert "健診機関" in str(e.value)

    def test_course_without_hpm_value(self, tmp_path):
        path = make_master_xlsx(
            tmp_path / "nocourseval.xlsx",
            courses=[("医療法人社団 同友会 春日クリニック", "定期健康診断", "")],
            aliases=[],
        )
        with pytest.raises(MasterError) as e:
            load_master(path)
        assert "HPM出力値" in str(e.value)


class TestCanonicalHeader:
    def test_bp_column_names(self):
        header = canonical_header()
        assert len(header) == 302
        assert header[50] == "血圧（一回目）最高"
        assert header[51] == "血圧（一回目）最低"
        assert header[52] == "血圧（二回目）最高"
        assert header[53] == "血圧（二回目）最低"

    def test_duplicate_names_exist_so_index_is_required(self):
        """列名は重複するので index で引くしかない、という前提を固定する。"""
        header = canonical_header()
        assert header[22] == header[23] == "場所コード"
        assert header[27] == header[183] == "ＢＭＩ"
        assert header[112] == header[113] == "コリンエステラーゼ"
