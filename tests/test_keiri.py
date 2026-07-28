"""経理モードのユニットテスト（Z:\\API連携\\tests から移植 2026-07-27）。

実データを使う検証（生成4CSV vs 経理の最終CSV）は Z:\\API連携 側の開発ツールが担当する。
ここでは実データ無しで確かめられる純粋ロジックだけを見る。
"""

import os
import unittest
from collections import defaultdict

from services.keiri_engine import (GENBUTSU_KEY, KYUSHOKU_BIKO, Resolver, YAKUIN_LOAN,
                                   build_kensan, build_kyuyo, jp_date, master_skipped_report,
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


class KyuyoLayoutTests(unittest.TestCase):
    """給与CSVの行配置（2026-07-28 経理担当レビューで判明した2件）。

    金額の突合は開発ツールの diff が見るが、**行の位置と取引の切り方は突合されない**ので
    ここで押さえる。実データの根拠は 2026-07 の最終CSV。
    """

    MASTER = {
        "z_jinkenhi": [{"source_key": "salary_items:allowance1", "freee_item": "人件費（{人件費区分}）"}],
        "z_deduction": [{"source_key": "salary_deduction_items:deduction29",
                         "freee_account": "預り金", "freee_tax": "対象外",
                         "freee_item": "健康保険料（預り分）", "amount_sign": "-1"}],
        "j_items": [{"source_key": GENBUTSU_KEY, "freee_account": "給料手当", "freee_tax": "対象外",
                     "freee_item": "人件費（{人件費区分}）", "amount_sign": "1"}],
    }
    MASTER["z_by_key"] = {r["source_key"]: r for r in MASTER["z_deduction"]}

    def _run(self, st_m, alerts=None):
        alerts = alerts if alerts is not None else defaultdict(set)
        resolver = Resolver({}, alerts, {})
        ridx = {e: {"name": f"社員{e}"} for e in st_m}
        return build_kyuyo("2026-07", "2026-06", st_m, {}, ridx, resolver, self.MASTER,
                           "2026-07-25", alerts), alerts

    @staticmethod
    def _pi(base=0, kenpo=0, net=0, genbutsu=0):
        return {"salary_items": [{"id": "allowance1", "value": base, "label": "基本給"},
                                 {"id": GENBUTSU_KEY.split(":")[1], "value": genbutsu,
                                  "label": "現物支給"}],
                "salary_deduction_items": [{"id": "deduction29", "value": kenpo}],
                "salary_payment_items": [{"id": "payment1", "value": net}]}

    def test_kyushoku_karibarai_is_its_own_one_row_transaction(self):
        """休職者は預り金を暫定取引へ、仮払金だけ1行=1取引で末尾に分ける。"""
        tx, alerts = self._run({"2018025": self._pi(kenpo=12978, net=-50020)})
        zantei = [t for t in tx if t["管理番号"]][-1]
        self.assertEqual([r["品目"] for r in zantei["rows"]], ["健康保険料（預り分）"])
        tail = [t for t in tx if not t["管理番号"]]
        self.assertEqual(len(tail), 1)
        self.assertEqual(tail[0]["支払期日"], "")
        self.assertEqual(tail[0]["発生日"], "2026/7/15")
        self.assertEqual(tail[0]["取引先"], "従業員")
        self.assertEqual(len(tail[0]["rows"]), 1)
        row = tail[0]["rows"][0]
        self.assertEqual((row["勘定科目"], row["品目"], row["金額"]), ("仮払金", "仮払金", 50020))
        self.assertEqual(row["備考"], KYUSHOKU_BIKO)
        self.assertEqual({e for e, _n, _a in alerts["kyushoku"]}, {"2018025"})

    def test_karibarai_stays_inline_without_shaho(self):
        """社保が無い人（＝休職者ではない）のマイナス精算は暫定の中のまま。"""
        tx, _a = self._run({"2024001": self._pi(net=-1000)})
        self.assertEqual([t["管理番号"] for t in tx if not t["管理番号"]], [])
        zantei = [t for t in tx if t["管理番号"]][-1]
        self.assertEqual([r["品目"] for r in zantei["rows"]], ["仮払金"])

    def test_genbutsu_offset_follows_the_payout_row(self):
        """現物支給の相殺（仮払金）は本人の現物支給行の直下に入る。"""
        tx, _a = self._run({"2024001": self._pi(genbutsu=3000)})
        jisseki = [t for t in tx if t["発生日"] == "2026/6/30"][0]
        items = [(r["勘定科目"], r["金額"], r["従業員"]) for r in jisseki["rows"]]
        self.assertEqual(items, [("給料手当", 3000, "社員2024001"),
                                 ("仮払金", -3000, "その他本社経費")])
        self.assertEqual(jisseki["rows"][1]["部門"], "本社")

    def test_genbutsu_offset_is_per_employee_not_bundled_at_the_end(self):
        """複数人いても各自の直下に入る（末尾へまとめない）。"""
        tx, _a = self._run({e: self._pi(genbutsu=3000) for e in ("2024001", "2024002")})
        jisseki = [t for t in tx if t["発生日"] == "2026/6/30"][0]
        self.assertEqual([r["勘定科目"] for r in jisseki["rows"]],
                         ["給料手当", "仮払金", "給料手当", "仮払金"])


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
