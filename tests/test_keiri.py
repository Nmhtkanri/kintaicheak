"""経理モードのユニットテスト（Z:\\API連携\\tests から移植 2026-07-27）。

実データを使う検証（生成4CSV vs 経理の最終CSV）は Z:\\API連携 側の開発ツールが担当する。
ここでは実データ無しで確かめられる純粋ロジックだけを見る。
"""

import os
import unittest

from services.keiri_engine import jp_date, split_halves, split_shaho_chosei, ym_add
from services.keiri_keihi_tenki import MAPPING_CSV, classify, decompose, load_mapping


def _pi(kenpo=0, kaigo=0, kounen=0, kodomo=0):
    """社保控除だけを持つ payroll_info を組み立てる。"""
    return {"salary_deduction_items": [
        {"id": "deduction29", "value": kenpo},
        {"id": "deduction30", "value": kaigo},
        {"id": "deduction31", "value": kounen},
        {"id": "child_support", "value": kodomo},
    ]}


class SplitShahoChoseiTests(unittest.TestCase):
    def test_splits_when_ratio_divides_exactly(self):
        """2026-07 小池2023019: 16,680 → 厚年10,980／健保5,562／子育て138（最終CSVと一致）。"""
        pi = _pi(kenpo=31518, kounen=62220, kodomo=782)
        parts = dict(split_shaho_chosei(-16680, pi))
        self.assertEqual(parts["salary_deduction_items:deduction31"], -10980)
        self.assertEqual(parts["salary_deduction_items:deduction29"], -5562)
        self.assertEqual(parts["salary_deduction_items:child_support"], -138)
        self.assertEqual(sum(parts.values()), -16680)

    def test_no_split_when_remainder(self):
        """2026-05 村山2013020 の 25円は月変差額ではない（最終CSVは雇用保険料）→ 分割しない。"""
        pi = _pi(kenpo=21785, kaigo=4230, kounen=43005, kodomo=540)
        self.assertEqual(split_shaho_chosei(25, pi), [])

    def test_no_split_without_basis(self):
        self.assertEqual(split_shaho_chosei(-1000, _pi()), [])

    def test_keeps_sign_of_positive_adjustment(self):
        pi = _pi(kenpo=1000, kounen=2000, kodomo=1000)
        parts = dict(split_shaho_chosei(400, pi))       # 400 × 各/4000 → 100/200/100
        self.assertEqual(sorted(parts.values()), [100, 100, 200])


class KeihiTenkiTests(unittest.TestCase):
    """C-4: 経費転記の品目判定（2026-05〜07 の最終CSVで全件一致を確認した組み合わせ）。"""

    @classmethod
    def setUpClass(cls):
        if not os.path.exists(MAPPING_CSV):
            raise unittest.SkipTest(f"マッピング表が参照できません: {MAPPING_CSV}")
        cls.mapping = load_mapping(MAPPING_CSV)

    def assertItem(self, uchiwake, memo, item, account):
        rule, _why = classify(uchiwake, memo, self.mapping)
        self.assertEqual((rule["freee_item"], rule["freee_account"]), (item, account),
                         f"内訳={uchiwake} 備考={memo}")

    def test_uchiwake_rules(self):
        self.assertItem("工具代", "", "消耗品費", "消耗品費")
        self.assertItem("資格取得手当", "", "社員育成費", "研修費")
        self.assertItem("定期健康診断費用", "", "福利厚生費", "福利厚生費")
        self.assertItem("郵便料金", "", "通信費", "通信費")

    def test_memo_beats_uchiwake(self):
        """備考のほうが具体的なので内訳より優先する（2026-07 矢嶋2018016）。"""
        self.assertItem("郵便料金", "顧客提出書類送付用", "雑費", "雑費")
        self.assertItem("その他経費", "甲府社宅備品購入代として", "消耗品費", "消耗品費")
        self.assertItem("その他経費", "メンバーマネジメントのための懇親会経費", "会議接待費", "会議費")

    def test_shuunyuu_inshi_is_sozei_kouka(self):
        self.assertItem("その他経費", "登記簿謄本×3部、印鑑証明書×3部取得", "雑費", "租税公課")

    def test_unmatched_falls_back_to_shomohin(self):
        rule, why = classify("見たことのない内訳", "", self.mapping)
        self.assertEqual(rule["freee_item"], "消耗品費")
        self.assertEqual(rule["status"], "要確認")
        self.assertIn("既定", why)

    def test_decompose_groups_by_item_and_keeps_total(self):
        """2026-07 岡野2024011: 62,580 → 社員育成費49,500 ＋ 会議接待費13,080。"""
        details = [("資格取得手当", 49500, "銀座コーチングスクール クラスA 受講費用"),
                   ("その他経費", 13080, "メンバーマネジメントのための懇親会経費")]
        by_item, biko, reasons, total = decompose(details, self.mapping)
        self.assertEqual(total, 62580)
        self.assertEqual(sum(by_item.values()), 62580)
        self.assertEqual({k[0]: v for k, v in by_item.items()},
                         {"社員育成費": 49500, "会議接待費": 13080})
        self.assertEqual(len(reasons), 2)
        self.assertTrue(all(b for b in biko.values()))

    def test_decompose_merges_same_item(self):
        """2026-07 矢嶋2018016: その他経費4件＋郵便料金 → 雑費1行 550。"""
        details = [("その他経費", 20, "顧客提出書類印刷"), ("その他経費", 60, "顧客提出書類スキャンメール添付用"),
                   ("その他経費", 30, "提出書類を送付するためにPDF化"), ("その他経費", 10, "提出書類印刷"),
                   ("郵便料金", 430, "顧客提出書類送付用")]
        by_item, biko, _reasons, total = decompose(details, self.mapping)
        self.assertEqual(total, 550)
        self.assertEqual({k[0]: v for k, v in by_item.items()}, {"雑費": 550})
        self.assertIn("ほか4件", list(biko.values())[0])


class DateHelperTests(unittest.TestCase):
    def test_ym_add_crosses_year(self):
        self.assertEqual(ym_add("2026-01", -1), "2025-12")
        self.assertEqual(ym_add("2026-12", 1), "2027-01")

    def test_jp_date_drops_zero_padding(self):
        self.assertEqual(jp_date("2026-07-05"), "2026/7/5")
        self.assertEqual(jp_date(""), "")

    def test_split_halves_puts_remainder_on_old_side(self):
        self.assertEqual(split_halves(101), (51, 50))
        self.assertEqual(split_halves(-101), (-51, -50))


if __name__ == "__main__":
    unittest.main()
