# -*- coding: utf-8 -*-
"""メール台帳×jinjer同期の差分計算テスト（API・Excel・COMは使わない純ロジックのみ）。"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.mail_ledger_sync import compute_ledger_diff


def person(id_, name="山田 太郎", company="c@x.jp", personal="p@z.jp",
           enrollment="在籍", retirement_date=""):
    return {"id": id_, "name": name, "company_email": company, "personal_email": personal,
            "enrollment": enrollment, "retired": "退職" in enrollment,
            "retirement_date": retirement_date}


def ledger_entry(id_, name="山田 太郎", company=("c@x.jp",), personal=("p@z.jp",)):
    return {id_: [{"employee_id": id_, "name": name,
                   "company": tuple(company), "client": (), "personal": tuple(personal)}]}


class ComputeLedgerDiffTest(unittest.TestCase):
    def test_new_active_employee_is_addition(self):
        diff = compute_ledger_diff({}, [person("2026016", "髙垣 和希")])
        self.assertEqual([a["id"] for a in diff["additions"]], ["2026016"])
        self.assertFalse(diff["additions"][0]["no_email"])

    def test_non_regular_ids_are_ignored(self):
        directory = [person("3333003"), person("5000001"), person("9999999")]
        diff = compute_ledger_diff({}, directory)
        self.assertEqual(diff["additions"], [])

    def test_retired_new_employee_not_added(self):
        diff = compute_ledger_diff({}, [person("2026016", enrollment="退職")])
        self.assertEqual(diff["additions"], [])

    def test_no_email_addition_is_flagged(self):
        diff = compute_ledger_diff({}, [person("2026017", company="", personal="")])
        self.assertTrue(diff["additions"][0]["no_email"])

    def test_ledger_retiree_is_delete_candidate(self):
        book = ledger_entry("2020021", "二神 太郎")
        directory = [person("2020021", "二神 太郎", enrollment="退職", retirement_date="2026-06-30")]
        diff = compute_ledger_diff(book, directory)
        self.assertEqual([r["id"] for r in diff["retirees"]], ["2020021"])
        self.assertEqual(diff["retirees"][0]["retirement_date"], "2026-06-30")

    def test_missing_in_jinjer_is_report_only(self):
        book = ledger_entry("2019099", "謎の 番号")
        diff = compute_ledger_diff(book, [person("2026016")])
        self.assertEqual([m["id"] for m in diff["missing_in_jinjer"]], ["2019099"])
        self.assertEqual(diff["retirees"], [])

    def test_matching_active_employee_produces_nothing(self):
        book = ledger_entry("2024001")
        diff = compute_ledger_diff(book, [person("2024001")])
        self.assertEqual(diff["additions"], [])
        self.assertEqual(diff["retirees"], [])
        self.assertEqual(diff["mismatches"], [])

    def test_mismatch_is_reported(self):
        book = ledger_entry("2024001", company=("old@x.jp",))
        diff = compute_ledger_diff(book, [person("2024001", company="new@x.jp")])
        self.assertEqual(len(diff["mismatches"]), 1)
        self.assertIn("社用が台帳とjinjerで不一致", diff["mismatches"][0])


if __name__ == "__main__":
    unittest.main()
