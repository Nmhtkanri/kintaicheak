# -*- coding: utf-8 -*-
"""健保①主取引の子ども・子育て支援金を「当月明細」から読む（2026-09-10）。

jinjer の給与明細は 会社負担（other15/16/17）＝その月の分／本人控除＝前月分（翌月控除）で
1 か月ずれる。子ども・子育て支援金は API に会社負担側の項目が無く、同額の本人控除
（salary_deduction_items:child_support）で代用しているため、①主取引（前月分の費用）を
前月明細から作ると支援金だけ前々月分になっていた（有田 2018001 の 2026-07 分 506 円が欠落）。

実データでの裏取り（2026-09-10）: キャッシュ済みの 2026-07・2026-08 で修正前後の 4 CSV を比較。
  - 2026-07 実行分（6月分計上）: 差分なし
  - 2026-08 実行分（7月分計上）: 健保 CSV 35 行が計 3,402 円増。明細側で支援金が
    7月明細→8月明細で変わった 36 人（有田の復職 1 人＋4月昇給→7月月変で健保本人控除も
    同時に変わった人）から、7月入社 2 人（入社月は当月明細を読む経路なので前後で不変、
    計 644 円）を除いた 34 人＝35 行（1 人は月中異動で 2 行に分割）と金額まで一致。

実行: python -X utf8 -m pytest tests/test_keiri_shaho_child_support.py -q
"""
from __future__ import annotations

import sys
import unittest
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services.keiri_engine import KENPO_AZUKARI, Resolver, build_shaho, jp_date  # noqa: E402

CS = "salary_deduction_items:child_support"
KENPO_ROWS = [
    {"source_key": "salary_other_items:other15", "amount_sign": "1", "freee_item": "人件費（{人件費区分}）",
     "freee_account": "法定福利費", "freee_tax": "対象外"},
    {"source_key": "salary_other_items:other16", "amount_sign": "1", "freee_item": "人件費（{人件費区分}）",
     "freee_account": "法定福利費", "freee_tax": "対象外"},
    {"source_key": CS, "amount_sign": "1", "freee_item": "人件費（{人件費区分}）",
     "freee_account": "法定福利費", "freee_tax": "対象外"},
]


def _pi(kenpo_kaisha=0, kaigo_kaisha=0, kenpo_honnin=0, kodomo=0):
    """会社負担（当月分）と本人控除（前月分）を持つ payroll_info。"""
    return {"salary_other_items": [{"id": "other15", "value": kenpo_kaisha},
                                   {"id": "other16", "value": kaigo_kaisha},
                                   {"id": "other17", "value": 0}],
            "salary_deduction_items": [{"id": "deduction29", "value": kenpo_honnin},
                                       {"id": "deduction31", "value": 0},
                                       {"id": "child_support", "value": kodomo}]}


class ShahoChildSupportTests(unittest.TestCase):
    MONTH, PREV = "2026-08", "2026-07"

    def _run(self, st_prev, st_m, ridx_extra=None):
        alerts = defaultdict(set)
        ridx = {e: {"name": f"社員{e}"} for e in set(st_prev) | set(st_m)}
        for e, extra in (ridx_extra or {}).items():
            ridx[e].update(extra)
        resolver = Resolver({}, alerts, ridx)
        tx = build_shaho(self.MONTH, self.PREV, st_prev, st_m, ridx, resolver, KENPO_ROWS,
                         "K1", "健保組合", KENPO_AZUKARI, alerts,
                         kyoshutsukin=False, split_midmonth=False)
        return tx, alerts

    @staticmethod
    def _main_amounts(tx):
        """①主取引（発生日=前月末）の 従業員→金額。"""
        main = [t for t in tx if t["発生日"] == jp_date("2026-07-31") and t["支払期日"] == jp_date("2026-08-31")]
        assert len(main) == 1, main
        return {r["従業員"]: r["金額"] for r in main[0]["rows"]}

    @staticmethod
    def _azukari_items(tx):
        """②預り金（発生日=当月末・期日=当月末）の 品目→金額。"""
        az = [t for t in tx if t["発生日"] == jp_date("2026-08-31") and t["支払期日"] == jp_date("2026-08-31")]
        assert len(az) == 1, az
        return {r["品目"]: r["金額"] for r in az[0]["rows"]}

    def test_normal_employee_takes_child_support_from_current_statement(self):
        """月変で支援金が 500→506 になった人: ①は健保(前月明細の会社負担)＋支援金(当月明細)。"""
        st_prev = {"2011001": _pi(kenpo_kaisha=20000, kenpo_honnin=19000, kodomo=500)}   # 500 は前々月分
        st_m = {"2011001": _pi(kenpo_kaisha=21000, kenpo_honnin=20000, kodomo=506)}      # 506 が前月分
        tx, _ = self._run(st_prev, st_m)
        self.assertEqual(self._main_amounts(tx)["社員2011001"], 20000 + 506)

    def test_returning_from_leave_gets_the_first_month_child_support(self):
        """有田 2018001 型: 前月明細は本人控除ゼロ（前々月分が免除）だが、前月分は当月明細で控除されている。"""
        st_prev = {"2018001": _pi(kenpo_kaisha=20394, kenpo_honnin=0, kodomo=0)}
        st_m = {"2018001": _pi(kenpo_kaisha=20394, kenpo_honnin=20394, kodomo=506)}
        tx, alerts = self._run(st_prev, st_m)
        self.assertEqual(self._main_amounts(tx)["社員2018001"], 20394 + 506)
        self.assertEqual(alerts["shaho_menjo"], set())

    def test_still_on_leave_is_excluded_entirely(self):
        """当月明細でも本人控除ゼロ＝育休継続中。支援金の読み元を変えても①には載らない。"""
        st_prev = {"2017037": _pi(kenpo_kaisha=20000, kenpo_honnin=0, kodomo=0)}
        st_m = {"2017037": _pi(kenpo_kaisha=20000, kenpo_honnin=0, kodomo=0)}
        tx, alerts = self._run(st_prev, st_m)
        self.assertEqual({e for e, _n, _m in alerts["shaho_menjo"]}, {"2017037"})
        # ①主取引そのものが作られない（この人しかいないので rows が空＝取引なし）
        self.assertEqual([t for t in tx if t["発生日"] == jp_date("2026-07-31")], [])

    def test_month_end_retiree_path_is_unchanged(self):
        """当月末退職者は当月明細×0.5（2倍徴収の半分）を①と③に載せる。読み元は元から当月明細。"""
        st_prev = {"2015010": _pi(kenpo_kaisha=20000, kenpo_honnin=20000, kodomo=500)}
        st_m = {"2015010": _pi(kenpo_kaisha=40000, kenpo_honnin=40000, kodomo=1012)}
        tx, _ = self._run(st_prev, st_m, {"2015010": {"retired_on": "2026-08-31"}})
        self.assertEqual(self._main_amounts(tx)["社員2015010"], 20000 + 506)
        retire = [t for t in tx if t["発生日"] == jp_date("2026-08-31") and t["支払期日"] == jp_date("2026-09-30")]
        self.assertEqual(len(retire), 1)
        self.assertEqual([r["金額"] for r in retire[0]["rows"] if r["従業員"] == "社員2015010"], [20000 + 506])

    def test_new_hire_path_is_unchanged(self):
        """前月入社（前月明細に社保なし）は当月明細をそのまま読む。修正前後で同じ。"""
        st_prev = {"2026020": _pi()}
        st_m = {"2026020": _pi(kenpo_kaisha=13300, kenpo_honnin=13300, kodomo=644)}
        tx, alerts = self._run(st_prev, st_m, {"2026020": {"joined_on": "2026-07-01"}})
        self.assertEqual(self._main_amounts(tx)["社員2026020"], 13300 + 644)
        self.assertEqual({e for e, *_ in alerts["shaho_new_hire"]}, {"2026020"})

    def test_konen_rows_without_deduction_items_are_unchanged(self):
        """厚生年金ファイル（マスタ行は other17 だけ）は当月明細を変えても①が動かない。"""
        konen_rows = [{"source_key": "salary_other_items:other17", "amount_sign": "1",
                       "freee_item": "人件費（{人件費区分}）", "freee_account": "法定福利費", "freee_tax": "対象外"}]
        alerts = defaultdict(set)
        ridx = {"2011001": {"name": "社員2011001"}}
        prev = _pi(kenpo_kaisha=20000, kenpo_honnin=19000, kodomo=500)
        prev["salary_other_items"][2]["value"] = 40000          # other17 前月分
        cur = _pi(kenpo_kaisha=21000, kenpo_honnin=20000, kodomo=506)
        cur["salary_other_items"][2]["value"] = 41000           # other17 当月分（①では使わない）
        tx = build_shaho(self.MONTH, self.PREV, {"2011001": prev}, {"2011001": cur}, ridx,
                         Resolver({}, alerts, ridx), konen_rows, "K2", "年金", [], alerts,
                         kyoshutsukin=False, split_midmonth=False)
        self.assertEqual(self._main_amounts(tx)["社員2011001"], 40000)

    def test_azukari_still_reads_current_statement(self):
        """②預り金は当月給与から控除した分＝当月明細（変更なし）。"""
        st_prev = {"2011001": _pi(kenpo_kaisha=20000, kenpo_honnin=19000, kodomo=500)}
        st_m = {"2011001": _pi(kenpo_kaisha=21000, kaigo_kaisha=3000, kenpo_honnin=20000, kodomo=506)}
        tx, _ = self._run(st_prev, st_m)
        items = self._azukari_items(tx)
        self.assertEqual(items["子ども・子育て支援金（預り分）"], 506)
        self.assertEqual(items["健康保険料（預り分）"], 21000)


if __name__ == "__main__":
    unittest.main()
