# -*- coding: utf-8 -*-
"""健保①主取引を「当月明細の本人控除」から読む（2026-09-10。支援金→全項目へ拡張）。

jinjer の給与明細は 会社負担（other15/16/17）＝その月の分／本人控除＝前月分（翌月控除）で
1 か月ずれる。子ども・子育て支援金は API に会社負担側の項目が無く、同額の本人控除
（salary_deduction_items:child_support）で代用しているため、①主取引（前月分の費用）を
前月明細から作ると支援金だけ前々月分になっていた（有田 2018001 の 2026-07 分 506 円が欠落）。

実データでの裏取り（2026-09-10、同日 15:2x に API から取り直した 2026-07/08 明細＋退職日入り名簿）:
  - 2026-08 実行分（7月分計上）を経理担当の最終 CSV と突合: 健保 金額差 34→2、厚年 33→2。
    ①の金額が変わった 33 人は、基準年月 2026-07 の標準報酬が 7/31（7月給与計算の後）に
    管理者登録された人で、当月明細の本人控除（新額）が最終 CSV と 1 円まで一致。
    残る 2 件は②預り金の本社合算行が 1 人分少ないもので、修正前からある別件。
  - 2026-07 実行分（6月分計上）: 健保 1・厚年 1 で前後同じ（既存の差）。
  - 当月末退職者（森田 2025001）は前後とも①＝③（健保 19,210／厚年 31,110）。

実行: python -X utf8 -m pytest tests/test_keiri_shaho_child_support.py -q
"""
from __future__ import annotations

import sys
import unittest
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services.keiri_engine import KENPO_AZUKARI, Resolver, build_shaho, build_yokakunin, jp_date  # noqa: E402

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
        """月変で支援金が 500→506 になった人: ①は当月明細の本人控除（健保 20,000＋支援金 506）。"""
        st_prev = {"2011001": _pi(kenpo_kaisha=20000, kenpo_honnin=19000, kodomo=500)}   # 500 は前々月分
        st_m = {"2011001": _pi(kenpo_kaisha=21000, kenpo_honnin=20000, kodomo=506)}      # 506 が前月分
        tx, alerts = self._run(st_prev, st_m)
        self.assertEqual(self._main_amounts(tx)["社員2011001"], 20000 + 506)
        # 支援金だけが変わった人は「前月会社負担≠当月本人控除」の一覧には出ない（比較は支援金を除く）
        self.assertEqual(alerts["shaho_prev_diff"], set())

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
        # 会社負担 42,000 ≠ 本人控除 40,000 にして、①が本人控除側を読むことを固定する
        st_m = {"2015010": _pi(kenpo_kaisha=42000, kenpo_honnin=40000, kodomo=1012)}
        tx, _ = self._run(st_prev, st_m, {"2015010": {"retired_on": "2026-08-31"}})
        self.assertEqual(self._main_amounts(tx)["社員2015010"], 20000 + 506)        # ①=本人控除×0.5
        retire = [t for t in tx if t["発生日"] == jp_date("2026-08-31") and t["支払期日"] == jp_date("2026-09-30")]
        self.assertEqual(len(retire), 1)
        self.assertEqual([r["金額"] for r in retire[0]["rows"] if r["従業員"] == "社員2015010"], [21000 + 506])  # ③=会社負担×0.5（従来どおり）
        # ④退職者預り（発生日=翌月末・期日=翌月末）は本人控除×0.5＝1か月分
        retire_az = [t for t in tx if t["発生日"] == jp_date("2026-09-30") and t["支払期日"] == jp_date("2026-09-30")]
        self.assertEqual(len(retire_az), 1)
        items = {r["品目"]: r["金額"] for r in retire_az[0]["rows"]}
        self.assertEqual(items["健康保険料（預り分）"], 20000)
        self.assertEqual(items["子ども・子育て支援金（預り分）"], 506)

    def test_new_hire_path_is_unchanged(self):
        """前月入社（前月明細に社保なし）は当月明細をそのまま読む。修正前後で同じ。"""
        st_prev = {"2026020": _pi()}
        st_m = {"2026020": _pi(kenpo_kaisha=13500, kenpo_honnin=13300, kodomo=644)}   # 会社負担≠本人控除
        tx, alerts = self._run(st_prev, st_m, {"2026020": {"joined_on": "2026-07-01"}})
        self.assertEqual(self._main_amounts(tx)["社員2026020"], 13300 + 644)
        self.assertEqual({(e, kind, fl) for e, _n, _j, _a, kind, fl in alerts["shaho_new_hire"]}, {("2026020", "入社月", "K1")})
        self.assertEqual(alerts["shaho_prev_diff"], set())

    def test_konen_reads_deduction31_of_current_statement(self):
        """厚生年金ファイル（マスタ行は other17）も①は当月明細の本人控除 deduction31 から読む。"""
        konen_rows = [{"source_key": "salary_other_items:other17", "amount_sign": "1",
                       "freee_item": "人件費（{人件費区分}）", "freee_account": "法定福利費", "freee_tax": "対象外"}]
        alerts = defaultdict(set)
        ridx = {"2011001": {"name": "社員2011001"}}
        prev = _pi(kenpo_kaisha=20000, kenpo_honnin=19000, kodomo=500)
        prev["salary_other_items"][2]["value"] = 40000          # other17 前月分（旧額）
        cur = _pi(kenpo_kaisha=21000, kenpo_honnin=20000, kodomo=506)
        cur["salary_other_items"][2]["value"] = 41000           # other17 当月分（①では使わない）
        cur["salary_deduction_items"][1]["value"] = 40500       # deduction31 当月明細の本人控除＝前月分
        tx = build_shaho(self.MONTH, self.PREV, {"2011001": prev}, {"2011001": cur}, ridx,
                         Resolver({}, alerts, ridx), konen_rows, "K2", "年金", [], alerts,
                         kyoshutsukin=False, split_midmonth=False)
        self.assertEqual(self._main_amounts(tx)["社員2011001"], 40500)

    def test_late_registration_uses_current_statement_deduction(self):
        """標準報酬が前月の給与計算の後に登録された人: 前月明細の会社負担は旧額（20,000）、
        当月明細の本人控除は新額（22,000）。①は新額（経理担当の実績と同じ）。"""
        st_prev = {"2011001": _pi(kenpo_kaisha=20000, kenpo_honnin=20000, kodomo=500)}
        st_m = {"2011001": _pi(kenpo_kaisha=22000, kenpo_honnin=22000, kodomo=560)}
        tx, alerts = self._run(st_prev, st_m)
        self.assertEqual(self._main_amounts(tx)["社員2011001"], 22000 + 560)
        # 前月明細の会社負担（20,000）と①に計上した額（22,560）の差を要確認に出す
        # 比較は支援金を除いた同じ範囲（前月会社負担 20,000 vs 当月本人控除 22,000）
        self.assertEqual(alerts["shaho_prev_diff"], {("2011001", "社員2011001", 20000, 22000, "K1")})

    def test_missing_prev_company_burden_is_flagged(self):
        """前月明細に会社負担が無いのに当月明細で前月分が控除されている（後追いの資格取得）は
        ①に載せたうえで shaho_new_hire に出す（黙って通さない）。"""
        st_prev = {"2011001": _pi()}
        st_m = {"2011001": _pi(kenpo_kaisha=20000, kenpo_honnin=20000, kodomo=506)}
        tx, alerts = self._run(st_prev, st_m)
        self.assertEqual(self._main_amounts(tx)["社員2011001"], 20000 + 506)
        self.assertEqual({(e, amt, kind) for e, _n, _j, amt, kind, _f in alerts["shaho_new_hire"]},
                         {("2011001", 20506, "後追い（前月明細に会社負担なし）")})

    def test_azukari_still_reads_current_statement(self):
        """②預り金は当月給与から控除した分＝当月明細の本人控除（2026-09-10 に会社負担キーから変更）。"""
        st_prev = {"2011001": _pi(kenpo_kaisha=20000, kenpo_honnin=19000, kodomo=500)}
        st_m = {"2011001": _pi(kenpo_kaisha=21000, kaigo_kaisha=3000, kenpo_honnin=20000, kodomo=506)}
        tx, _ = self._run(st_prev, st_m)
        items = self._azukari_items(tx)
        self.assertEqual(items["子ども・子育て支援金（預り分）"], 506)
        self.assertEqual(items["健康保険料（預り分）"], 20000)   # 本人控除 deduction29（会社負担 21,000 ではない）
        self.assertNotIn("介護保険料（預り分）", items)           # deduction30 は 0


class YokakuninShahoSectionsTests(unittest.TestCase):
    """要確認 md の社保 2 節が、alerts のタプル構成（shaho_new_hire 6 要素／shaho_prev_diff 5 要素）で描画できる。"""

    def test_sections_render_with_one_row_each(self):
        alerts = defaultdict(set)
        alerts["shaho_new_hire"].add(("2026020", "社員2026020", "2026-07-01", 13944, "入社月", "健保"))
        alerts["shaho_prev_diff"].add(("2011001", "社員2011001", 20000, 22000, "健保"))
        md = chr(10).join(build_yokakunin("2026-08", alerts, {}))   # 行のリストを返す
        self.assertIn("## 前月分の社保を当月明細の本人控除から拾った人", md)
        self.assertIn("| 2026020 | 社員2026020 | 2026-07-01 | 入社月 | 健保 | 13,944 |", md)
        self.assertIn("## 前月明細の会社負担と当月明細の本人控除が違う人", md)
        self.assertIn("| 2011001 | 社員2011001 | 健保 | 20,000 | 22,000 | +2,000 |", md)

    def test_sections_say_none_when_empty(self):
        md = chr(10).join(build_yokakunin("2026-08", defaultdict(set), {}))
        i = md.index("## 前月明細の会社負担と当月明細の本人控除が違う人")
        self.assertIn("- なし", md[i:i + 400])


if __name__ == "__main__":
    unittest.main()
