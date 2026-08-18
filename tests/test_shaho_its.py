# -*- coding: utf-8 -*-
"""関東ITSの被保険者標準報酬決定通知書（PATPOSTでCSV化）の読み取りと突合。

実CSVはリポジトリに置かない（氏名・報酬額が入るため）。
PATPOSTが実際に出す崩れ方（`440 000` `1,330 000` `昭 和43年`）を再現した
文字列を組み立てて純関数に流す。
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.shaho_its import (build_statement, load_number_map,  # noqa: E402
                                name_key, read_rows, verify_sequence)
from services.shaho_pdf import ShahoPdfError  # noqa: E402

HEADER = ",被保険者 証番号,氏 名,生年月日,性別,従前の 標準報酬月額,平均額 (単純または修正),決定 標準報酬月額"
ROWS = [
    "1,1015,試験 太郎,昭和51年 6月25日,男,620 千円,\"595,833円\",590000",
    # PATPOSTが入れる空白・カンマの崩れ
    "2,1032,試験 花子,昭 和43年 4月13日,女,\"1,270千円\",\"1,348,570円\",\"1,330 000\"",
    "3,1147,ヴァー マーティン リエゴ,昭和60年 1月 1日,男,410千円,\"424,350 円\",440 000",
]


# 空リスト＝明細ゼロのテストなので、`or` で既定に落とさない
def make_csv(rows=None, header=HEADER):
    body = "\r\n".join(["〒 104 - 0061東京都中央区,,,,,,,", header] + list(ROWS if rows is None else rows))
    path = os.path.join(tempfile.mkdtemp(), "決定通知.csv")
    with open(path, "w", encoding="cp932", newline="") as f:
        f.write(body)
    return path


class FakeGrade:
    def __init__(self, kenpo, konen):
        self.kenpo_smr, self.konen_smr = kenpo, konen
        self.kenpo_grade = self.konen_grade = 1


class FakeMaster:
    """健保はそのまま・厚年は65万で頭打ちにする最小の等級表。"""

    def find_grade(self, amount):
        return FakeGrade(int(amount), min(int(amount), 650000))


def roster(**people):
    return {emp: {"name": name, "enrollment": "在籍", "retired_on": ""}
            for emp, name in people.items()}


class ReadTests(unittest.TestCase):
    def test_reads_only_detail_rows(self):
        rows = read_rows(make_csv())
        self.assertEqual([r["no"] for r in rows], ["1015", "1032", "1147"])

    def test_decided_amount_is_yen_even_with_spaces(self):
        """H列は円単位。`1,330 000` のような崩れを数字だけにして読む。"""
        rows = read_rows(make_csv())
        self.assertEqual([r["decided"] for r in rows], [590000, 1330000, 440000])

    def test_previous_amount_is_thousand_yen(self):
        """F列は千円単位。H列と単位が違う。"""
        self.assertEqual([r["prev"] for r in read_rows(make_csv())],
                         [620000, 1270000, 410000])

    def test_empty_file_is_rejected(self):
        with self.assertRaises(ShahoPdfError):
            read_rows(make_csv(rows=[]))

    def test_name_key_absorbs_spaces_only(self):
        self.assertEqual(name_key("試験 太郎"), name_key("試験　太郎"))
        self.assertNotEqual(name_key("井料 香凜"), name_key("井料 香凛"))  # 異体字は別物


class SequenceTests(unittest.TestCase):
    """PATPOSTの行落ち検知。A列（ページ内連番）は 1,2,3…でページが変わると1に戻る。"""

    def test_page_reset_is_normal(self):
        self.assertEqual(verify_sequence([1, 2, 3, 1, 2, 3, 4, 1]), [])

    def test_gap_means_a_dropped_row(self):
        problems = verify_sequence([1, 2, 4, 5])
        self.assertEqual(len(problems), 1)
        self.assertIn("2 の次に 4", problems[0])

    def test_not_starting_at_one_means_first_row_dropped(self):
        problems = verify_sequence([3, 4, 5])
        self.assertTrue(any("1から始まっていない" in m for m in problems))

    def test_duplicate_is_detected(self):
        self.assertTrue(verify_sequence([1, 2, 2, 3]))

    def test_read_rows_stops_on_gap(self):
        """連番が飛んだCSVは読み込みごと止める（欠けた1名が静かに消えるより良い）。"""
        rows = [
            "1,1015,試験 太郎,昭和51年 6月25日,男,620 千円,\"595,833円\",590000",
            "2,1022,試験 次郎,昭和52年 6月25日,男,410千円,\"424,350 円\",410000",
            "4,1032,試験 花子,昭和43年 4月13日,女,440千円,\"440,000円\",440000",
        ]
        with self.assertRaises(ShahoPdfError) as cm:
            read_rows(make_csv(rows=rows))
        self.assertIn("2 の次に 4", str(cm.exception))
        self.assertIn("PATPOST", str(cm.exception))

    def test_unreadable_detail_row_is_not_silently_dropped(self):
        """A列は数字なのにH列が読めない行＝以前は黙って捨てていた。今は止める。"""
        rows = [
            "1,1015,試験 太郎,昭和51年 6月25日,男,620 千円,\"595,833円\",590000",
            "2,1022,壊れ 次郎,昭和52年 6月25日,男,410千円,\"424,350 円\",",
        ]
        with self.assertRaises(ShahoPdfError) as cm:
            read_rows(make_csv(rows=rows))
        self.assertIn("壊れ 次郎", str(cm.exception))
        self.assertIn("読み取れません", str(cm.exception))


class MatchTests(unittest.TestCase):
    """本人の特定は証番号と氏名の両方で行う（片方だけでは事故る）。"""

    def _one(self, *, no="1015", name="試験 太郎", number_map=None, people=None):
        rows = [{"no": no, "name": name, "prev": 0, "avg": 0, "decided": 300000}]
        stmt = build_statement(rows, "2026-09",
                               roster=people if people is not None else roster(**{"2099001": "試験 太郎"}),
                               number_map=number_map if number_map is not None else {"1015": "2099001"},
                               master=FakeMaster())
        return stmt.persons[0]

    def test_both_keys_agree(self):
        p = self._one()
        self.assertEqual(p.emp, "2099001")
        self.assertEqual(p.issues, [])
        self.assertEqual(p.warnings, [])

    def test_keys_point_at_different_people_is_unresolved(self):
        """★実際に踏んだケース: 証番号と氏名が別人を指す（柴田さん/岡村さん）。"""
        p = self._one(no="1151", name="岡村 大士",
                      number_map={"1151": "2099001"},
                      people=roster(**{"2099001": "柴田 和浩", "2099002": "岡村 大士"}))
        self.assertEqual(p.emp, "")            # 誰に書くか決めない
        self.assertTrue(any("本人を特定できません" in m for m in p.issues))

    def test_number_wins_when_name_differs(self):
        """異体字・ローマ字は氏名では当たらない。証番号で拾い、警告を出す。"""
        p = self._one(no="1147", name="ヴァー マーティン リエゴ",
                      number_map={"1147": "2099001"},
                      people=roster(**{"2099001": "Ver Martin"}))
        self.assertEqual(p.emp, "2099001")
        self.assertEqual(p.issues, [])
        self.assertTrue(any("氏名が jinjer と違います" in m for m in p.warnings))

    def test_name_is_used_when_number_is_not_registered(self):
        p = self._one(number_map={})
        self.assertEqual(p.emp, "2099001")
        self.assertTrue(any("未登録" in m for m in p.warnings))

    def test_duplicate_name_without_number_is_unresolved(self):
        """同姓同名（谷津さんのように社員番号が2つある人）は特定しない。"""
        p = self._one(number_map={},
                      people=roster(**{"2099001": "試験 太郎", "3333001": "試験 太郎"}))
        self.assertEqual(p.emp, "")
        self.assertTrue(any("同じ氏名" in m for m in p.issues))

    def test_unknown_person(self):
        p = self._one(number_map={}, people={})
        self.assertEqual(p.emp, "")
        self.assertTrue(p.issues)


class PensionTests(unittest.TestCase):
    """この通知には健保しか載っていないので、厚年は等級表から導く。"""

    def _person(self, decided):
        rows = [{"no": "1015", "name": "試験 太郎", "prev": 0, "avg": 0,
                 "decided": decided}]
        return build_statement(rows, "2026-09", roster=roster(**{"2099001": "試験 太郎"}),
                               number_map={"1015": "2099001"},
                               master=FakeMaster()).persons[0]

    def test_pension_follows_the_grade_table(self):
        p = self._person(300000)
        self.assertEqual((p.kenpo_smr, p.konen_smr), (300000, 300000))
        self.assertEqual(p.warnings, [])

    def test_pension_is_capped(self):
        """健保139万でも厚年は65万で頭打ち。黙って上限を書くのではなく警告も出す。"""
        p = self._person(1390000)
        self.assertEqual((p.kenpo_smr, p.konen_smr), (1390000, 650000))
        self.assertTrue(any("上限" in m for m in p.warnings))

    def test_amount_off_the_grade_table_is_an_issue(self):
        class Off:
            def find_grade(self, amount):
                return FakeGrade(300000, 300000)      # 決定額と一致しない等級を返す
        rows = [{"no": "1015", "name": "試験 太郎", "prev": 0, "avg": 0,
                 "decided": 305000}]
        p = build_statement(rows, "2026-09", roster=roster(**{"2099001": "試験 太郎"}),
                            number_map={"1015": "2099001"}, master=Off()).persons[0]
        self.assertTrue(any("等級表" in m for m in p.issues))


class StatementTests(unittest.TestCase):
    def test_period_and_reason(self):
        stmt = build_statement(read_rows(make_csv()), "2026-09",
                               roster=roster(**{"2099001": "試験 太郎"}),
                               number_map={"1015": "2099001"}, master=FakeMaster())
        self.assertEqual((stmt.target_ym, stmt.pay_ym), ("2026-09", "2026-10"))
        self.assertEqual(stmt.total_count, 3)
        self.assertTrue(all(p.reason == "定時決定" for p in stmt.persons))

    def test_same_employee_twice_is_flagged(self):
        rows = [{"no": "1015", "name": "試験 太郎", "prev": 0, "avg": 0, "decided": 300000},
                {"no": "1016", "name": "試験 太郎", "prev": 0, "avg": 0, "decided": 320000}]
        stmt = build_statement(rows, "2026-09", roster=roster(**{"2099001": "試験 太郎"}),
                               number_map={"1015": "2099001", "1016": "2099001"},
                               master=FakeMaster())
        self.assertTrue(any("複数行" in m for p in stmt.persons for m in p.issues))


class NumberMapTests(unittest.TestCase):
    class FakeClient:
        base_url = "https://example.invalid"

        def _auth_headers(self):
            return {}

    def test_duplicate_numbers_are_dropped(self):
        """同じ証番号が2人に付いていたら、取り違えるより特定不能に倒す。"""
        import requests
        from unittest import mock

        payload = [
            {"employee_id": "2099001",
             "social_insurance": {"health_insurance": {"number": "1151"}}},
            {"employee_id": "2099002",
             "social_insurance": {"health_insurance": {"number": "1151"}}},
        ]

        class Res:
            status_code = 200

            @staticmethod
            def json():
                return {"data": payload}

        with mock.patch.object(requests, "get", return_value=Res()):
            result = load_number_map(self.FakeClient(), ["2099001", "2099002"])
        self.assertEqual(result, {})

    def test_normal_mapping(self):
        import requests
        from unittest import mock

        class Res:
            status_code = 200

            @staticmethod
            def json():
                return {"data": [
                    {"employee_id": "2099001",
                     "social_insurance": {"health_insurance": {"number": "1015"}}},
                    {"employee_id": "2099002",
                     "social_insurance": {"health_insurance": {"number": ""}}},
                ]}

        with mock.patch.object(requests, "get", return_value=Res()):
            result = load_number_map(self.FakeClient(), ["2099001", "2099002"])
        self.assertEqual(result, {"1015": "2099001"})


class PlanTests(unittest.TestCase):
    """build_plan に載せたときに、特定できない人が投入対象にならないこと。"""

    def test_unresolved_is_not_selectable(self):
        from services.shaho_writer import build_plan
        rows = [{"no": "1151", "name": "岡村 大士", "prev": 0, "avg": 0, "decided": 560000}]
        people = roster(**{"2099001": "柴田 和浩", "2099002": "岡村 大士"})
        stmt = build_statement(rows, "2026-09", roster=people,
                               number_map={"1151": "2099001"}, master=FakeMaster())
        plan = build_plan(stmt, {}, people, None)
        self.assertEqual(plan[0].status, "UNRESOLVED")
        self.assertFalse(plan[0].selectable)
        self.assertFalse(plan[0].needs_force)     # 承知のうえ投入でも通さない

    def test_capped_pension_warning_reaches_the_screen(self):
        from services.shaho_writer import build_plan
        rows = [{"no": "1015", "name": "試験 太郎", "prev": 0, "avg": 0,
                 "decided": 1390000}]
        people = roster(**{"2099001": "試験 太郎"})
        stmt = build_statement(rows, "2026-09", roster=people,
                               number_map={"1015": "2099001"}, master=FakeMaster())
        plan = build_plan(stmt, {}, people, None)
        self.assertTrue(any("上限" in n for n in plan[0].notes))


if __name__ == "__main__":
    unittest.main()
