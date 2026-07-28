"""経理モードのユニットテスト（Z:\\API連携\\tests から移植 2026-07-27）。

実データを使う検証（生成4CSV vs 経理の最終CSV）は Z:\\API連携 側の開発ツールが担当する。
ここでは実データ無しで確かめられる純粋ロジックだけを見る。
"""

import os
import unittest

from services.keiri_engine import (YAKUIN_LOAN, build_kensan, jp_date, master_skipped_report,
                                   split_halves, split_shaho_chosei, ym_add)
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


def _master(skipped_rows=(), active_n=38, total_n=185):
    """検算シートが読むぶんだけのマスタ（load_master の戻り値の部分集合）。"""
    from collections import Counter
    return {"skipped_rows": list(skipped_rows), "active_n": active_n, "total_n": total_n,
            "skipped": Counter(r["status"] for r in skipped_rows),
            "offscope": Counter({"対象外": 146}), "min_status": "確定",
            "path": r"Z:\API連携\docs\経理モード_品目マッピングマスタ_draftC.csv"}


def _skipped_row(source_key, label, status="推定"):
    return {"source_key": source_key, "label": label, "status": status}


def _deduction_pi(source_key, value):
    array, item_id = source_key.split(":")
    return {array: [{"id": item_id, "value": value}]}


class KensanJuminzeiTests(unittest.TestCase):
    """住民税の『不一致』が毎月赤く出ていた件（谷津さん指摘 2026-07-28）。

    差は前月に会社が立て替えた分の相殺そのもの＝正常。相殺分を引いてから比べる。
    """

    def _line(self, api_total, csv_total, soosai):
        st_m = {"2026001": _deduction_pi("salary_deduction_items:deduction41", api_total)}
        files = {"住民税": [{"rows": [{"金額": csv_total, "勘定科目": "預り金", "品目": "住民税"}]}],
                 "給与": [], "健康保険": [], "厚生年金": []}
        alerts = {"juminzei_soosai": soosai}
        lines = build_kensan("2026-07", "2026-06", files, st_m, st_m, _master(), alerts)
        return next(l for l in lines if l.startswith("- 住民税:"))

    def test_offset_is_reported_as_match(self):
        """2026-07 実測: 3,419,600 − 16,900(奥山5,800+柳場11,100) = 3,402,700。"""
        line = self._line(3419600, 3402700,
                          {("2025029", "奥山 昌苗", 5800, 4700),
                           ("2026010", "柳場 涼馬", 11100, 10300)})
        self.assertIn("一致", line)
        self.assertNotIn("不一致", line)
        self.assertIn("前月立替の相殺 16,900円・2名", line)

    def test_plain_match_without_offset(self):
        line = self._line(3402700, 3402700, set())
        self.assertIn("一致", line)
        self.assertNotIn("相殺", line)

    def test_real_gap_still_flagged(self):
        """相殺で説明できない残差は今までどおり不一致で出す（見張りを殺さない）。"""
        line = self._line(3419600, 3400000, {("2025029", "奥山 昌苗", 5800, 4700)})
        self.assertIn("**不一致（残差 13,800円）**", line)


class MasterSkippedReportTests(unittest.TestCase):
    """『採用マスタ行 38行（status除外: 推定=1）』が何の欄か分からない件への対応。"""

    def test_no_amount_this_month_is_harmless(self):
        row = _skipped_row("salary_deduction_items:deduction3", "貸付金返済")
        rep = master_skipped_report(_master([row]), {"2026001": _deduction_pi("x:y", 0)})
        self.assertEqual(rep[0]["people"], 0)
        self.assertIn("影響なし", rep[0]["impact"])

    def test_amount_without_rule_is_flagged_as_missing(self):
        row = _skipped_row("salary_deduction_items:deduction3", "貸付金返済")
        st_m = {"2024001": _deduction_pi("salary_deduction_items:deduction3", 50000)}
        rep = master_skipped_report(_master([row]), st_m, {"2024001": {"name": "山田 太郎"}})
        self.assertIn("⚠️", rep[0]["impact"])
        self.assertIn("漏れています", rep[0]["impact"])
        self.assertIn("2024001 山田 太郎 50,000円", rep[0]["impact"])

    def test_special_rule_owner_is_not_a_false_alarm(self):
        """三谷さんの貸付金返済は役員貸付金の特別ルールで生成済み＝漏れではない。"""
        for key in YAKUIN_LOAN["source_keys"]:
            row = _skipped_row(key, "貸付金返済")
            st_m = {YAKUIN_LOAN["employee_id"]: _deduction_pi(key, YAKUIN_LOAN["payment"])}
            rep = master_skipped_report(_master([row]), st_m,
                                        {YAKUIN_LOAN["employee_id"]: {"name": "三谷 一志"}})
            self.assertNotIn("⚠️", rep[0]["impact"], key)
            self.assertIn("特別ルールで計上済み", rep[0]["impact"])


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
