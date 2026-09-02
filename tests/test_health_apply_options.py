# -*- coding: utf-8 -*-
"""健康診断申込: 「選択肢」シートの解釈（コードは文字列のまま・別名・旧表記）。"""

import pytest

from services.health_apply import options as O
from services.health_apply.schema import OPTION_HEADERS, SchemaError, rows_to_dicts


def rows(*cells):
    """(区分, コード, 表示名, 有効, 並び順, 別名, 備考) のタプル列から dict 行を作る。"""
    raw = [list(c) for c in cells]
    return rows_to_dicts(OPTION_HEADERS, raw)


def sample_catalog():
    return O.OptionCatalog.from_rows(rows(
        ("機関", "1310528885", "医療法人社団 同友会 春日クリニック", "1", "10", "医療法人社団同友会 春日クリニック;同友会"),
        ("機関", "0301619", "医療法人徳洲会 生駒市立病院", "", "20"),
        ("機関", "13X5035440", "関東ITソフトウェア健康組合(大久保健診センター)", "1", "30"),
        ("機関", "1310436535", "医療法人社団　善仁会　新宿西口ヘルチェッククリニック", "1", "40"),
        ("機関", "130192", "東京品川病院 総合健診センター", "0", "50", "", "同名の別コードあり"),
        ("機関", "OTHER", "その他", "1", "999"),
        ("種別", "10", "定期健康診断", "1", "1", "基本健診"),
        ("種別", "12", "人間ドックB", "1", "3", "1日人間ドック・バリウム"),
        ("種別", "13", "人間ドックC", "1", "4", "1日人間ドック・胃カメラ"),
        ("種別", "15", "人間ドック1日コース", "1", "6", "1日人間ドック"),
        ("追加検査", "GYN", "婦人科検診", "1", "1", "婦人病検査;婦人科健診"),
        ("続柄", "妻", "妻", "1", "1"),
        ("続柄", "夫", "夫", "1", "2"),
    ))


def test_codes_stay_strings_with_leading_zero_and_letters():
    cat = sample_catalog()
    assert cat.lookup("機関", "0301619").name == "医療法人徳洲会 生駒市立病院"
    assert cat.lookup("機関", "13X5035440").name.startswith("関東IT")
    assert cat.lookup("機関", "301619") is None          # 前ゼロが落ちたコードは別物
    assert cat.lookup("種別", " 10 ").code == "10"       # 前後空白だけは許す


def test_blank_active_means_active_and_zero_means_inactive():
    cat = sample_catalog()
    assert cat.lookup("機関", "0301619").active is True
    assert cat.lookup("機関", "130192").active is False
    names = [o.code for o in cat.of_kind("機関")]
    assert "130192" not in names
    assert "130192" in [o.code for o in cat.of_kind("機関", active_only=False)]


def test_of_kind_sorted_by_order_then_name_with_other_last():
    codes = [o.code for o in sample_catalog().of_kind("機関")]
    assert codes == ["1310528885", "0301619", "13X5035440", "1310436535", "OTHER"]


def test_resolve_name_ignores_width_spacing_and_case():
    cat = sample_catalog()
    assert cat.resolve_name("機関", "医療法人社団善仁会 新宿西口ヘルチェッククリニック").code == "1310436535"
    assert cat.resolve_name("機関", "医療法人社団 同友会 春日クリニック\n").code == "1310528885"
    assert cat.resolve_name("機関", "同友会").code == "1310528885"            # 別名
    assert cat.resolve_name("機関", "13x5035440") is None                      # 名前でなくコードは引かない
    assert cat.resolve_name("種別", "基本健診 ").code == "10"                   # jinjer 側の末尾スペース
    assert cat.resolve_name("種別", "１日人間ドック・バリウム").code == "12"    # 全角数字
    assert cat.resolve_name("種別", "人間ドックA") is None


def test_extra_legacy_names_resolve_to_gyn():
    cat = sample_catalog()
    assert cat.resolve_name("追加検査", "婦人病検査").code == O.GYN_CODE
    assert cat.resolve_name("追加検査", "婦人科健診").code == O.GYN_CODE
    assert cat.resolve_name("追加検査", "婦人科検診").code == O.GYN_CODE
    assert O.normalize_extra_name("婦人病検査") == "婦人科検診"
    assert O.normalize_extra_name("胃カメラ") == "胃カメラ"


def test_display_unknown_code():
    cat = sample_catalog()
    assert cat.display("機関", "1310528885") == "医療法人社団 同友会 春日クリニック"
    assert cat.display("機関", "9999") == "（不明: 9999）"
    assert cat.display("機関", "") == "（不明: 空）"


def test_counts():
    assert sample_catalog().counts() == {"institutions": 5, "exam_types": 4, "extras": 1, "relationships": 2}


def test_duplicate_code_in_same_kind_is_rejected():
    with pytest.raises(SchemaError, match="重複"):
        O.OptionCatalog.from_rows(rows(
            ("機関", "130192", "東京品川病院 総合健診センター", "1", "1"),
            ("機関", "130192", "東京品川病院", "1", "2"),
        ))


def test_same_code_in_different_kinds_is_fine():
    cat = O.OptionCatalog.from_rows(rows(
        ("機関", "10", "十番クリニック", "1", "1"),
        ("種別", "10", "定期健康診断", "1", "1"),
    ))
    assert cat.lookup("機関", "10").name == "十番クリニック"
    assert cat.lookup("種別", "10").name == "定期健康診断"


def test_unknown_kind_blank_code_and_bad_order_are_rejected_with_row_number():
    with pytest.raises(SchemaError, match="2行目の区分"):
        O.OptionCatalog.from_rows(rows(("施設", "1", "x", "1", "1")))
    with pytest.raises(SchemaError, match="3行目のコードが空"):
        O.OptionCatalog.from_rows(rows(("機関", "1", "x", "1", "1"), ("機関", "", "y", "1", "1")))
    with pytest.raises(SchemaError, match="2行目の並び順"):
        O.OptionCatalog.from_rows(rows(("機関", "1", "x", "1", "abc")))


def test_blank_rows_are_skipped_and_blank_order_goes_last():
    cat = O.OptionCatalog.from_rows(rows(
        ("機関", "2", "後", "1", ""),
        ("", "", "", "", ""),
        ("機関", "1", "先", "1", "5"),
    ))
    assert [o.name for o in cat.of_kind("機関")] == ["先", "後"]


def test_hpm_txt_paste_shape():
    """hpm.txt をそのまま貼った形（区分列を足しただけ・有効/並び順/別名は空）でも読める。"""
    pasted = [
        ["機関", "0301619", "医療法人徳洲会 生駒市立病院"],
        ["機関", "040006", "医療法人財団 明理会 IMS Me‑Lifeクリニック仙台（旧 イムス仙台クリニック）"],
        ["機関", "14098", "MYメディカルクリニック 横浜みなとみらい"],
        ["種別", "10", "定期健康診断"],
        ["種別", "14", "雇用時の健康診断"],
    ]
    cat = O.OptionCatalog.from_rows(rows_to_dicts(OPTION_HEADERS, pasted))
    assert cat.lookup("機関", "040006").active is True
    assert cat.lookup("機関", "040006").order == 9999
    assert [o.code for o in cat.of_kind("種別")] == ["10", "14"]
