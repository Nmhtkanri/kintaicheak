# -*- coding: utf-8 -*-
"""保険料一覧表PDFのパーサ。

**実PDFはリポジトリに置かない**（氏名・生年月日・報酬額が入るため）。
テキスト層から抜いた文字列を組み立てて純関数に流す。数字は実際の料率で
計算した値にしてあるので、料率検算のテストもそのまま通る。

最後の RealFileTests だけは共有フォルダの実物を読む（無ければ自動スキップ）。
"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.shaho_master import Rate  # noqa: E402
from services.shaho_pdf import (ShahoPdfError, parse_text, read_pdf,  # noqa: E402
                                verify_person_premiums, verify_totals)

REAL_PDF = (r"Z:\NMHT総務関係\社労士提出\提出書類\2026年度\2026.07\送付0730"
            r"\7月保険料一覧表（8月給与控除分）\7月保険料一覧表（8月給与控除分）.pdf")

HEADER = """事業所 263:株式会社 テスト 令和08年07月30日 印刷 1頁
保険料一覧表
令和08年07月分 (令和08年08月給与分)
健保№ 個人ｺｰﾄﾞ 氏 名 標準報酬 健康保険 (健康保険) (介護保険) 厚生年金 (厚生年金) (厚生年金基金)
厚年№ 生年月日 年齢 種別 健保 厚年 事業主 個人分 (事業主)(個人分)(子ども・子育て支援金)事業主 個人分 (事業主)(個人分)(事業主)(個人分) 改訂理由 介護区分
(内特定保険) (事業主)(個人分) ２社勤務"""

# 標報30万・第2号: 本体13,905 介護2,700 子育て345 → 健保計16,950 / 厚年27,450
P1 = """1001 2099001 試験 太郎 月額変更 第２号
1 昭55年12月06日 45 1 300 300 16,950 16,950 13,905 13,905 2,700 2,700 27,450 27,450 0 0 0 0 月額変更
( 0) 345 345"""
# 標報28万・非対象: 本体12,978 介護0 子育て322 → 健保計13,300 / 厚年25,620
P2 = """1002 2099002 試験 花子 取得時決定 非対象
2 平12年06月16日 26 2 280 280 13,300 13,300 12,978 12,978 0 0 25,620 25,620 0 0 0 0 取得時決定
( 0) 322 322"""

TOTALS = """【総 合 計】人数 2 30,250 26,883 2,700 53,070 0 0
子ども・子育て拠出金 1,000 500 30,250 26,883 2,700 53,070 0 0
( 0) 667
667
全体 60,500 53,766 5,400 106,140 0 0
1,334"""


def make_text(*people, header=HEADER, totals=TOTALS):
    parts = [header] + list(people or (P1, P2))
    if totals:
        parts.append(totals)
    return "\n".join(parts)


def fake_master():
    """料率だけを持つ等級表のダミー（本物と同じ令和8年度・関東IT健保の率）。"""
    class _M:
        rates = {
            "kenpo": Rate(total=0.0927, employee=0.04635),
            "kodomo": Rate(total=0.0023, employee=0.00115),
            "kaigo": Rate(total=0.018, employee=0.009),
            "konen": Rate(total=0.183, employee=0.0915),
        }
    return _M()


class ParseTests(unittest.TestCase):
    def test_reads_two_people(self):
        stmt = parse_text(make_text())
        self.assertEqual([p.emp for p in stmt.persons], ["2099001", "2099002"])
        self.assertEqual([p.name for p in stmt.persons], ["試験 太郎", "試験 花子"])

    def test_period_is_converted_to_seireki(self):
        stmt = parse_text(make_text())
        self.assertEqual(stmt.target_ym, "2026-07")   # 令和08年07月分
        self.assertEqual(stmt.pay_ym, "2026-08")      # 令和08年08月給与分
        self.assertEqual(stmt.office_code, "263")

    def test_standard_remuneration_is_converted_to_yen(self):
        """帳票は千円単位。円に直して渡さないと桁が1000分の1になる。"""
        p = parse_text(make_text()).persons[0]
        self.assertEqual(p.kenpo_smr, 300000)
        self.assertEqual(p.konen_smr, 300000)

    def test_split_smr_between_kenpo_and_konen(self):
        """厚年は上限65万で頭打ち。健保と厚年で値が違う人を取りこぼさない。"""
        person = """1003 2099003 上限 太郎 月額変更 第２号
3 昭40年08月28日 60 1 1150 650 64,975 64,975 53,303 53,302 10,350 10,350 59,475 59,475 0 0 0 0
( 0) 1,322 1,323"""
        p = parse_text(make_text(person, totals="")).persons[0]
        self.assertEqual((p.kenpo_smr, p.konen_smr), (1150000, 650000))

    def test_premium_columns(self):
        p = parse_text(make_text()).persons[0]
        self.assertEqual(p.premiums["kenpo_total_ee"], 16950)
        self.assertEqual(p.premiums["kenpo_ee"], 13905)
        self.assertEqual(p.premiums["kaigo_ee"], 2700)
        self.assertEqual(p.premiums["konen_ee"], 27450)
        self.assertEqual(p.premiums["kodomo_ee"], 345)

    def test_kaigo_and_reason(self):
        a, b = parse_text(make_text()).persons
        self.assertEqual((a.kaigo_kubun, a.reason), ("第２号", "月額変更"))
        self.assertEqual((b.kaigo_kubun, b.reason), ("非対象", "取得時決定"))

    def test_unknown_reason_is_kept_as_free_text(self):
        """9月は「定時決定」が来る。理由を列挙で受けていると新語で氏名が壊れる。"""
        person = P1.replace("月額変更", "定時決定")
        p = parse_text(make_text(person, totals="")).persons[0]
        self.assertEqual(p.reason, "定時決定")
        self.assertEqual(p.name, "試験 太郎")

    def test_reason_may_differ_between_kenpo_and_konen(self):
        """健保だけ改定・厚年は据え置き（上限）の人がいる。"""
        person = """1004 2099004 片側 改定 月額変更 第２号
4 昭40年08月28日 60 1 1150 650 64,975 64,975 53,303 53,302 10,350 10,350 59,475 59,475 0 0 0 0
( 0) 1,322 1,323"""
        p = parse_text(make_text(person, totals="")).persons[0]
        self.assertEqual(p.reason_kenpo, "月額変更")
        self.assertEqual(p.reason_konen, "")
        self.assertEqual(p.name, "片側 改定")

    def test_missing_reason_and_kaigo(self):
        person = P1.replace(" 月額変更 第２号", "").replace(" 0 0 月額変更", " 0 0")
        p = parse_text(make_text(person, totals="")).persons[0]
        self.assertEqual(p.name, "試験 太郎")
        self.assertEqual(p.reason, "")
        self.assertEqual(p.kenpo_smr, 300000)

    def test_romaji_name_with_space(self):
        person = P1.replace("試験 太郎", "MAHARJAN RAMITA")
        p = parse_text(make_text(person, totals="")).persons[0]
        self.assertEqual(p.name, "MAHARJAN RAMITA")

    def test_two_company_is_flagged(self):
        person = P1.replace("( 0) 345 345", "( 0) 345 345 ２社勤務")
        p = parse_text(make_text(person, totals="")).persons[0]
        self.assertTrue(p.two_company)
        self.assertTrue(p.issues)

    def test_duplicate_employee_code_is_flagged(self):
        dup = P2.replace("2099002", "2099001")
        stmt = parse_text(make_text(P1, dup, totals=""))
        self.assertTrue(all(p.issues for p in stmt.persons))

    def test_unexpected_column_count_is_flagged_not_fatal(self):
        broken = P1.replace(" 0 0 0 0 月額変更", " 0 0 月額変更")
        p = parse_text(make_text(broken, totals="")).persons[0]
        self.assertTrue(p.issues)
        self.assertEqual(p.kenpo_smr, 0)   # 値は採らない（推測しない）


class GuardTests(unittest.TestCase):
    def test_other_document_is_rejected(self):
        with self.assertRaises(ShahoPdfError):
            parse_text("これは別の書類です\n令和08年07月分 (令和08年08月給与分)")

    def test_other_office_is_rejected(self):
        with self.assertRaises(ShahoPdfError) as cm:
            parse_text(make_text(), expected_office="999")
        self.assertIn("別の事業所", str(cm.exception))

    def test_expected_office_passes(self):
        stmt = parse_text(make_text(), expected_office="263")
        self.assertEqual(stmt.office_code, "263")

    def test_no_person_rows_is_rejected(self):
        with self.assertRaises(ShahoPdfError):
            parse_text(HEADER)

    def test_mixed_periods_are_rejected(self):
        text = make_text() + "\n令和08年08月分 (令和08年09月給与分)"
        with self.assertRaises(ShahoPdfError) as cm:
            parse_text(text)
        self.assertIn("複数の対象月", str(cm.exception))


class ChecksumTests(unittest.TestCase):
    def test_totals_match(self):
        checked = verify_totals(parse_text(make_text()))
        self.assertIn("人数", checked)
        self.assertIn("konen_ee", checked)

    def test_person_count_mismatch_stops(self):
        stmt = parse_text(make_text(P1, totals=TOTALS))   # 合計は2名ぶんのまま
        with self.assertRaises(ShahoPdfError) as cm:
            verify_totals(stmt)
        self.assertIn("人数が合いません", str(cm.exception))

    def test_amount_mismatch_stops(self):
        bad = TOTALS.replace("53,070", "53,071")
        with self.assertRaises(ShahoPdfError) as cm:
            verify_totals(parse_text(make_text(P1, P2, totals=bad)))
        self.assertIn("合計が合いません", str(cm.exception))

    def test_no_total_block_is_tolerated(self):
        """合計行が取れない様式でも、人数チェック以外は動くようにしておく。"""
        stmt = parse_text(make_text(P1, P2, totals=""))
        self.assertEqual(verify_totals(stmt), [])


class PremiumVerifyTests(unittest.TestCase):
    def test_clean_pdf_has_no_problems(self):
        stmt = parse_text(make_text())
        self.assertEqual(verify_person_premiums(stmt, fake_master()), {})

    def test_internal_breakdown_mismatch(self):
        """健保計 ≠ 本体+介護+子育て はマスタが無くても捕まえる。"""
        broken = P1.replace("16,950 16,950", "16,950 17,000")
        issues = verify_person_premiums(parse_text(make_text(broken, totals="")))
        self.assertIn("2099001", issues)
        self.assertTrue(any("内訳と合いません" in m for m in issues["2099001"]))

    def test_rate_mismatch_is_detected(self):
        """標報だけ書き換えると、保険料と料率の関係が崩れる＝読み違いの網。"""
        broken = P1.replace(" 300 300 ", " 320 320 ")
        issues = verify_person_premiums(parse_text(make_text(broken, totals="")),
                                        fake_master())
        self.assertIn("2099001", issues)

    def test_kaigo_not_charged_is_not_flagged(self):
        """介護保険料0の人に介護の料率検算をかけない（第2号でないだけ）。"""
        issues = verify_person_premiums(parse_text(make_text(P2, totals="")), fake_master())
        self.assertEqual(issues, {})


class RealFileTests(unittest.TestCase):
    """共有フォルダの実物で回帰を見る（無い環境では自動スキップ）。"""

    @classmethod
    def setUpClass(cls):
        if not os.path.exists(REAL_PDF):
            raise unittest.SkipTest(f"実PDFがありません: {REAL_PDF}")
        cls.stmt = read_pdf(REAL_PDF, expected_office="263")

    def test_all_rows_and_checksum(self):
        self.assertEqual(len(self.stmt.persons), 36)
        self.assertEqual(self.stmt.total_count, 36)
        verify_totals(self.stmt)                     # 合わなければ例外

    def test_period(self):
        self.assertEqual((self.stmt.target_ym, self.stmt.pay_ym), ("2026-07", "2026-08"))

    def test_known_person(self):
        """手入力ミスが見つかった遠田さん（正しくは28万）。"""
        p = self.stmt.by_emp["2024027"]
        self.assertEqual((p.kenpo_smr, p.konen_smr), (280000, 280000))

    def test_upper_limit_person(self):
        p = self.stmt.by_emp["2024005"]
        self.assertEqual((p.kenpo_smr, p.konen_smr), (1150000, 650000))

    def test_premiums_verify_against_real_master(self):
        from config import Config
        from services.shaho_master import load_grade_table
        if not os.path.exists(Config.SHAHO_GRADE_TABLE_XLSX):
            self.skipTest("等級表がありません")
        master = load_grade_table(Config.SHAHO_GRADE_TABLE_XLSX, Config.SHAHO_INSURER, 2026)
        self.assertEqual(verify_person_premiums(self.stmt, master), {})


if __name__ == "__main__":
    unittest.main()
