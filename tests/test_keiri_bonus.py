# -*- coding: utf-8 -*-
"""賞与用 freee CSV（services/keiri_bonus.py）のテスト。合成データ。

実データでの裏取りは tests/test_keiri_bonus_realdata.py（共有フォルダが無ければ skip）。
実行: python -X utf8 -m pytest tests/test_keiri_bonus.py -q
"""
from __future__ import annotations

import io
import sys
import unittest
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services.keiri_bonus import (DEFAULT_RATES, build_bonus, file_names, load_bonus_csv,  # noqa: E402
                                  load_rates, round_gosha_rokunyu, round_half_up)


class FakeResolver:
    def __init__(self, bumon="FS：UAL（受託）", item="人件費（本社以外）"):
        self._bumon, self._item = bumon, item
        self.calls = []

    def bumon(self, emp, on_date, context):
        self.calls.append((emp, on_date, context))
        return self._bumon

    def jinkenhi_item(self, emp, on_date):
        return self._item


def _row(emp, name, total, *, sb=150000, kenpo=6952, kaigo=0, shien=172, konen=13725, koyo=1300,
         zei=8000, kyoshutsu=540, paid="2026/9/25", nencho=0, extra=None):
    ded = kenpo + kaigo + shien + konen + koyo + zei + nencho
    r = {"社員番号": emp, "氏名": name, "総支給額": str(total), "支給日": paid,
         "健保標準賞与額": str(sb), "厚年標準賞与額": str(sb),
         "健康保険料": str(kenpo), "介護保険料": str(kaigo), "子ども・子育て支援金": str(shien),
         "厚生年金": str(konen), "雇用保険料": str(koyo), "所得税": str(zei),
         "年調過不足額": str(nencho), "控除合計": str(ded), "差引支給額": str(total - ded),
         "事業主子ども・子育て拠出金": str(kyoshutsu),
         "事業主健康保険料": str(kenpo), "事業主介護保険料": str(kaigo),
         "事業主子ども・子育て支援金": str(shien), "事業主厚生年金": str(konen)}
    for i in range(1, 16):
        r[f"賞与控除項目{i}"] = "0"
    r.update(extra or {})
    return r


class RoundingTests(unittest.TestCase):
    def test_half_up_and_gosha_rokunyu(self):
        self.assertEqual(round_half_up(6952.5), 6953)
        self.assertEqual(round_gosha_rokunyu(6952.5), 6952)
        self.assertEqual(round_gosha_rokunyu(6952.6), 6953)
        self.assertEqual(round_half_up(10402.5), 10403)
        self.assertEqual(round_gosha_rokunyu(10150.65), 10151)

    def test_company_rule_matches_accountant_examples(self):
        """2026-09 FE部賞与の実測: 150,000×11.3%÷2=8,475、109,000×9.5%÷2=5,177.5→5,178、
        219,000×18.3%÷2=20,038.5→20,039。"""
        r = DEFAULT_RATES
        self.assertEqual(round_half_up(150000 * (r["kenpo"] + r["shien"] + r["kaigo"]) / 2), 8475)
        self.assertEqual(round_half_up(109000 * (r["kenpo"] + r["shien"]) / 2), 5178)
        self.assertEqual(round_half_up(219000 * r["konen"] / 2), 20039)


class BuildBonusTests(unittest.TestCase):
    def _build(self, rows, **kw):
        alerts = defaultdict(set)
        ridx = {"2018012": {"name": "山田 太郎"}, "2008003": {"name": "友納 英彦"}}
        res = FakeResolver()
        files = build_bonus(rows, month="2026-09", hassei="2026-09-15", resolver=res, ridx=ridx,
                            rates=dict(DEFAULT_RATES), alerts=alerts, **kw)
        return files, alerts, res

    def test_pay_file_rows_and_order(self):
        files, alerts, res = self._build([_row("2018012", "山田　太郎", 219000, sb=219000, kenpo=10151,
                                              shien=252, konen=20038, koyo=1300, zei=8000)])
        tx = files["支給"][0]
        self.assertEqual((tx["管理番号"], tx["発生日"], tx["支払期日"], tx["取引先"]),
                         ("2018012", "2026/9/15", "2026/9/25", "従業員"))
        self.assertEqual([(r["勘定科目"], r["品目"], r["金額"]) for r in tx["rows"]], [
            ("賞与手当", "人件費（本社以外）", 219000),
            ("預り金", "厚生年金保険料（預り分）", -20038),
            ("預り金", "健康保険料（預り分）", -10151),
            ("預り金", "雇用保険料（預り分）", -1300),
            ("預り金", "源泉所得税", -8000),
            ("預り金", "子ども・子育て支援金（預り分）", -252),   # 介護 0 は出ない
        ])
        self.assertEqual(tx["rows"][0]["従業員"], "山田 太郎")      # 氏名は名簿の「姓 名」
        self.assertEqual(res.calls[0], ("2018012", "2026-09-25", "2026-09賞与"))  # 部門は支給日時点
        self.assertFalse(alerts["bonus_rate"]); self.assertFalse(alerts["bonus_check"])

    def test_pay_file_is_one_transaction_for_everyone(self):
        """支給ファイルは全員で 1 取引。管理番号は先頭（最小）の社員番号、行は社員番号順に連結。"""
        rows = [_row("2019002", "鈴木　花子", 150000), _row("2018012", "山田　太郎", 219000, sb=219000,
                                                             kenpo=10151, shien=252, konen=20038)]
        files, alerts, _ = self._build(rows)
        self.assertEqual(len(files["支給"]), 1)
        tx = files["支給"][0]
        self.assertEqual(tx["管理番号"], "2018012")
        self.assertEqual([r["従業員"] for r in tx["rows"] if r["勘定科目"] in ("賞与手当", "役員賞与")],
                         ["山田 太郎", "鈴木 花子"])
        self.assertFalse(alerts["bonus_check"])

    def test_multiple_pay_dates_split_transactions_and_are_flagged(self):
        rows = [_row("2018012", "山田　太郎", 100000, sb=100000, kenpo=4635, shien=115, konen=9150),
                _row("2019002", "鈴木　花子", 150000, paid="2026/9/30")]
        files, alerts, _ = self._build(rows)
        self.assertEqual([(t["管理番号"], t["支払期日"]) for t in files["支給"]], [("2018012", "2026/9/25"), ("2019002", "2026/9/30")])
        self.assertTrue(any("支給日が" in msg for _e, _n, msg in alerts["bonus_check"]))

    def test_duplicate_employee_rows_are_flagged(self):
        rows = [_row("2018012", "山田　太郎", 100000, sb=100000, kenpo=4635, shien=115, konen=9150)] * 2
        _files, alerts, _ = self._build(rows)
        self.assertTrue(any(e == "2018012" and "2 行" in msg for e, _n, msg in alerts["bonus_check"]))

    def test_non_numeric_amount_and_missing_employer_columns_are_flagged(self):
        r = _row("2018012", "山田　太郎", 100000, sb=100000, kenpo=4635, shien=115, konen=9150,
                 extra={"雇用保険料": "￥1,300"})
        for c in ("事業主健康保険料", "事業主介護保険料", "事業主子ども・子育て支援金", "事業主厚生年金"):
            r.pop(c, None)
        _files, alerts, _ = self._build([r])
        msgs = [m for _e, _n, m in alerts["bonus_check"]]
        self.assertTrue(any("雇用保険料" in m and "数値として読めず" in m for m in msgs))
        self.assertTrue(any("事業主〜列が無く" in m for m in msgs))

    def test_shaho_files_company_burden_and_aggregates(self):
        rows = [_row("2018012", "山田　太郎", 219000, sb=219000, kenpo=10151, shien=252, konen=20038, kyoshutsu=788),
                _row("2019002", "鈴木　花子", 150000, sb=150000, kenpo=6952, kaigo=1350, shien=172, konen=13725, kyoshutsu=540)]
        files, alerts, _ = self._build(rows)
        k = files["健康保険"][0]
        self.assertEqual((k["管理番号"], k["発生日"], k["支払期日"], k["取引先"]),
                         ("KEMPO", "2026/9/30", "2026/10/31", "関東ＩＴソフトウェア健康保険組合"))
        self.assertEqual([(r["勘定科目"], r["品目"], r["金額"], r["部門"]) for r in k["rows"]], [
            ("法定福利費", "人件費（本社以外）", 10403, "FS：UAL（受託）"),   # 219,000×9.5%÷2=10,402.5→10,403
            ("法定福利費", "人件費（本社以外）", 8475, "FS：UAL（受託）"),    # 150,000×11.3%÷2
            ("預り金", "健康保険料（預り分）", 17103, "本社"),
            ("預り金", "介護保険料（預り分）", 1350, "本社"),
            ("預り金", "子ども・子育て支援金（預り分）", 424, "本社"),
        ])
        n = files["厚生年金"][0]
        self.assertEqual(n["取引先"], "厚生労働省")
        self.assertEqual([(r["勘定科目"], r["品目"], r["金額"], r["備考"]) for r in n["rows"]], [
            ("法定福利費", "人件費（本社以外）", 20039, ""),   # 219,000×18.3%÷2=20,038.5→20,039
            ("法定福利費", "人件費（本社以外）", 13725, ""),
            ("預り金", "厚生年金保険料（預り分）", 33763, ""),
            ("法定福利費", "人件費（本社）", 1328, "子ども・子育て拠出金"),
        ])
        self.assertFalse(alerts["bonus_rate"])

    def test_shaho_dates_can_be_overridden(self):
        files, _, _ = self._build([_row("2018012", "山田　太郎", 100000, sb=100000, kenpo=4635, shien=115, konen=9150)],
                                  shaho_hassei="2026-03-31", shaho_kigen="2026-05-31")
        self.assertEqual((files["健康保険"][0]["発生日"], files["健康保険"][0]["支払期日"]), ("2026/3/31", "2026/5/31"))

    def test_yakuin_uses_bonus_account(self):
        files, _, _ = self._build([_row("2008003", "友納　英彦", 500000, sb=500000, kenpo=23175, kaigo=4500,
                                        shien=575, konen=45750)])
        self.assertEqual(files["支給"][0]["rows"][0]["勘定科目"], "役員賞与")

    def test_rate_mismatch_and_unmapped_deductions_are_flagged(self):
        rows = [_row("2018012", "山田　太郎", 219000, sb=219000, kenpo=10151 + 100, shien=252, konen=20038,
                     nencho=-1200, extra={"賞与控除項目3": "5000"})]
        files, alerts, _ = self._build(rows)
        self.assertEqual({(e, lbl, got, exp) for e, _n, lbl, got, exp in alerts["bonus_rate"]},
                         {("2018012", "健康保険料", 10251, 10151)})
        self.assertEqual({(col, amt) for _e, _n, col, amt in alerts["bonus_unmapped"]},
                         {("年調過不足額", -1200), ("賞与控除項目3", 5000)})
        self.assertTrue(alerts["bonus_check"])   # 預り金に載せた控除 ≠ 控除合計

    def test_company_diff_against_jinjer_columns_is_listed(self):
        """jinjer の事業主列（項目ごと丸め）と計算値が 1 円違う人は要確認に出す（金額は計算値）。"""
        rows = [_row("2019002", "鈴木　花子", 150000, sb=150000, kenpo=6952, kaigo=1350, shien=172, konen=13725,
                     extra={"事業主健康保険料": "6952", "事業主介護保険料": "1350", "事業主子ども・子育て支援金": "172",
                            "事業主厚生年金": "13725"})]
        files, alerts, _ = self._build(rows)
        self.assertEqual(files["健康保険"][0]["rows"][0]["金額"], 8475)      # 計算値を採用
        self.assertEqual({(e, lbl, calc, jin) for e, _n, lbl, calc, jin in alerts["bonus_company_diff"]},
                         {("2019002", "健保（介護・支援金込み）", 8475, 8474)})   # 厚年は一致なので出ない

    def test_pay_date_in_other_month_and_zero_total_are_flagged(self):
        rows = [_row("2018012", "山田　太郎", 100000, sb=100000, kenpo=4635, shien=115, konen=9150, paid="2026/8/25"),
                _row("2019002", "鈴木　花子", 0, sb=0, kenpo=0, shien=0, konen=0, koyo=0, zei=0, kyoshutsu=0)]
        files, alerts, _ = self._build(rows)
        msgs = [m for _e, _n, m in alerts["bonus_check"]]
        self.assertTrue(any("支給月 2026-09 と違います" in m for m in msgs))
        self.assertTrue(any("総支給額が 0" in m for m in msgs))
        self.assertEqual([r["従業員"] for t in files["支給"] for r in t["rows"] if r["勘定科目"] == "賞与手当"], ["山田 太郎"])

    def test_file_names(self):
        self.assertEqual(file_names("FE部賞与", "2026-09"), {
            "支給": "freee_data_FE部賞与（202609）.csv",
            "健康保険": "freee_data_FE部賞与_健康保険（202609）.csv",
            "厚生年金": "freee_data_FE部賞与_厚生年金（202609）.csv"})


def _api_person(emp, last, first, *, total=219670, koyo=1098, kenpo=10151, kaigo=0, konen=20038, zei=23049,
                sb=219000, sashihiki=None, count=1, paid_on="2026-09-08", other15=None, other17=None, kyoshutsu=788):
    """実 API（2026-09-11 に 2018012 で確認した構造）を模した 1 人分。支援金は API に無い。"""
    shien = 252 if sb == 219000 else 0
    if sashihiki is None:
        sashihiki = total - (koyo + kenpo + kaigo + konen + zei + shien)
    ded = [("deduction6", "雇用保険料", koyo), ("deduction7", "健康保険料", kenpo), ("deduction8", "介護保険料", kaigo),
           ("deduction9", "厚生年金", konen), ("deduction10", "厚生年金基金", 0), ("deduction18", "年調過不足額", 0),
           ("deduction19", "所得税", zei)] + [(f"deduction{i}", f"賞与控除項目{i}", 0) for i in range(1, 6)]
    oth = [("other13", "雇用保険料", 1867), ("other14", "労災保険料", 659), ("other15", "健康保険料", other15 if other15 is not None else kenpo),
           ("other16", "介護保険料", kaigo), ("other17", "厚生年金", other17 if other17 is not None else konen),
           ("other19", "子ども・子育て拠出金", kyoshutsu), ("other31", "健保標準賞与額", sb), ("other32", "厚年標準賞与額", sb)]
    return {"employee_id": emp, "statements": [{
        "executed_on": "2026-09", "count": count, "paid_on": paid_on,
        "basic_info": {"last_name": last, "first_name": first},
        "payroll_info": {"is_calculated": True,
                         "bonus_items": [{"id": "allowance1", "label": "賞与", "value": total}],
                         "bonus_deduction_items": [{"id": i, "label": l, "value": v} for i, l, v in ded],
                         "bonus_payment_items": [{"id": "payment1", "label": "差引支給額", "value": sashihiki}],
                         "bonus_other_items": [{"id": i, "label": l, "value": v} for i, l, v in oth]}}]}


class ApiRowsTests(unittest.TestCase):
    def test_api_record_becomes_csv_shaped_row_with_computed_shien(self):
        from services.keiri_bonus import rows_from_api
        alerts = defaultdict(set)
        rows = rows_from_api([_api_person("2018012", "山田", "太郎")], rates=dict(DEFAULT_RATES), alerts=alerts,
                             paid_on="2026-09-25")
        r = rows[0]
        self.assertEqual((r["社員番号"], r["氏名"], r["総支給額"], r["支給日"]), ("2018012", "山田 太郎", "219670", "2026/09/25"))
        self.assertEqual((r["健康保険料"], r["厚生年金"], r["雇用保険料"], r["所得税"]), ("10151", "20038", "1098", "23049"))
        self.assertEqual(r["子ども・子育て支援金"], "252")            # 219,000×0.23%÷2=251.85→五捨六入 252
        self.assertEqual(r["事業主子ども・子育て支援金"], "252")      # 本人と同額（谷津さん確認）
        self.assertEqual((r["事業主健康保険料"], r["事業主厚生年金"], r["事業主子ども・子育て拠出金"]), ("10151", "20038", "788"))
        self.assertEqual((r["健保標準賞与額"], r["厚年標準賞与額"]), ("219000", "219000"))
        self.assertEqual(int(r["総支給額"]) - int(r["控除合計"]), int(r["差引支給額"]))
        self.assertFalse(alerts["bonus_check"])
        # そのまま build_bonus に流せる
        files = build_bonus(rows, month="2026-09", hassei="2026-09-15", resolver=FakeResolver(),
                            ridx={}, rates=dict(DEFAULT_RATES), alerts=alerts)
        self.assertEqual(files["支給"][0]["支払期日"], "2026/9/25")
        self.assertEqual(files["健康保険"][0]["rows"][0]["金額"], 10403)

    def test_special_target_and_unknown_deduction_and_paid_on_note(self):
        from services.keiri_bonus import rows_from_api
        alerts = defaultdict(set)
        p = _api_person("7777777", "監査", "太郎")                     # 非 20YY だが計上対象（SPECIAL_TARGET_EMPLOYEES）
        q = _api_person("2018012", "山田", "太郎")
        q["statements"][0]["payroll_info"]["bonus_deduction_items"].append({"id": "deduction3", "label": "貸付金返済", "value": 5000})
        q["statements"][0]["payroll_info"]["bonus_payment_items"][0]["value"] -= 5000
        r_skip = _api_person("5000001", "派遣", "太郎")
        rows = rows_from_api([p, q, r_skip], rates=dict(DEFAULT_RATES), alerts=alerts)      # paid_on 未指定
        self.assertEqual([r["社員番号"] for r in rows], ["7777777", "2018012"])
        msgs = [m for _e, _n, m in alerts["bonus_check"]]
        self.assertTrue(any("対象外の社員番号の人を 1 人読み飛ばしました: 5000001" in m for m in msgs))
        self.assertEqual({(e, col, amt) for e, _n, col, amt in alerts["bonus_unmapped"]}, {("2018012", "控除「貸付金返済」（API）", 5000)})
        self.assertTrue(any("支援金の計算値" in m for m in msgs))      # 未知の控除は listed に入らないので逆算と食い違う

    def test_nencho_refund_and_kaigo_keep_shien_check_consistent(self):
        from services.keiri_bonus import rows_from_api
        alerts = defaultdict(set)
        p = _api_person("2019002", "鈴木", "花子", total=150000, kenpo=6952, kaigo=1350, konen=13725, zei=8000, sb=150000)
        # 年調過不足額 −1,200（還付）と介護あり。差引 = 総支給 − (雇用+健保+介護+厚年+所得税+年調(−1200)+支援金172)
        pi = p["statements"][0]["payroll_info"]
        for it in pi["bonus_deduction_items"]:
            if it["label"] == "年調過不足額": it["value"] = -1200
        pi["bonus_payment_items"][0]["value"] = 150000 - (1098 + 6952 + 1350 + 13725 + 8000 - 1200 + 172)
        rows = rows_from_api([p], rates=dict(DEFAULT_RATES), alerts=alerts, paid_on="2026-09-25")
        self.assertEqual(rows[0]["子ども・子育て支援金"], "172")
        self.assertFalse(any("支援金の計算値" in m for _e, _n, m in alerts["bonus_check"]))
        # 年調過不足額の要確認は build_bonus 側で出す（変換関数の責務ではない）

    def test_shien_mismatch_with_sashihiki_is_flagged(self):
        from services.keiri_bonus import rows_from_api
        alerts = defaultdict(set)
        # 差引支給額が支援金 252 を引いていない（＝支援金 0 で計算された明細）と、計算値と食い違う
        rows_from_api([_api_person("2018012", "山田", "太郎", sashihiki=219670 - (1098 + 10151 + 20038 + 23049))],
                      rates=dict(DEFAULT_RATES), alerts=alerts)
        self.assertTrue(any("支援金の計算値" in m for _e, _n, m in alerts["bonus_check"]))

    def test_multiple_counts_require_selection(self):
        from services.keiri_bonus import rows_from_api
        a = _api_person("2018012", "山田", "太郎", count=1)
        b = _api_person("2019002", "鈴木", "花子", count=2)
        with self.assertRaises(ValueError):
            rows_from_api([a, b], rates=dict(DEFAULT_RATES), alerts=defaultdict(set))
        rows = rows_from_api([a, b], rates=dict(DEFAULT_RATES), alerts=defaultdict(set), count=2)
        self.assertEqual([r["社員番号"] for r in rows], ["2019002"])

    def test_non_target_and_uncalculated_are_skipped(self):
        from services.keiri_bonus import rows_from_api
        p = _api_person("5000001", "派遣", "太郎")
        q = _api_person("2018012", "山田", "太郎")
        q["statements"][0]["payroll_info"]["is_calculated"] = False
        with self.assertRaises(ValueError):
            rows_from_api([p, q], rates=dict(DEFAULT_RATES), alerts=defaultdict(set))


class FetchAndGenerateApiTests(unittest.TestCase):
    def _fake_client(self, data):
        class C:
            def __init__(self): self.calls = 0
            def get_bonus_statements(self, ym, employee_ids=None):
                self.calls += 1; return data
        return C()

    def test_empty_or_uncalculated_response_is_not_cached(self):
        import tempfile
        import os
        from services.keiri_bonus import fetch_bonus_statements
        d = tempfile.mkdtemp()
        with self.assertRaises(ValueError):
            fetch_bonus_statements(d, "2026-09", client=self._fake_client([]))
        unc = _api_person("2018012", "山田", "太郎"); unc["statements"][0]["payroll_info"]["is_calculated"] = False
        with self.assertRaises(ValueError):
            fetch_bonus_statements(d, "2026-09", client=self._fake_client([unc]))
        self.assertFalse(os.path.exists(os.path.join(d, "raw", "bonus_statements_2026-09.json")))
        # 計算済みがあれば保存され、2 回目はキャッシュから（client を呼ばない）
        c = self._fake_client([_api_person("2018012", "山田", "太郎")])
        data, t1 = fetch_bonus_statements(d, "2026-09", client=c)
        data2, t2 = fetch_bonus_statements(d, "2026-09", client=c)
        self.assertEqual((c.calls, len(data2), t1), (1, 1, t2))
        c2 = self._fake_client([_api_person("2018012", "山田", "太郎"), _api_person("2019002", "鈴木", "花子")])
        data3, _ = fetch_bonus_statements(d, "2026-09", client=c2, refresh=True)
        self.assertEqual(len(data3), 2)

    def test_generate_bonus_from_api_end_to_end(self):
        import tempfile
        import os
        import json
        from unittest import mock
        from services import keiri_bonus as kb
        d = tempfile.mkdtemp()
        os.makedirs(os.path.join(d, "raw"))
        roster = [{"id": "2018012", "company": {"last_name": "山田", "first_name": "太郎", "joined_on": "2018-04-01"}}]
        json.dump({"fetched_at": "x", "data": roster}, io.open(os.path.join(d, "raw", "roster.json"), "w", encoding="utf-8"))
        with mock.patch.object(kb, "load_custom_histories", lambda out_base, ids, refresh=False: {}), \
             mock.patch.object(kb, "load_or_fetch_roster", lambda client, path, refresh=False: roster):
            res = kb.generate_bonus("2026-09", None, "FE部賞与", "2026-09-15", out_base=d,
                                    client=self._fake_client([_api_person("2018012", "山田", "太郎")]),
                                    source="api", paid_on="2026-09-25", rates_csv=r"C:\__no_such__\rates.csv")
        self.assertEqual(res["people"], 1)
        self.assertIn("API bonus-statements 2026-09", res["input_src"])
        self.assertTrue(os.path.exists(os.path.join(d, "raw", "bonus_statements_2026-09.json")))
        md = res["yokakunin_md"]
        self.assertIn("- 入力: API bonus-statements 2026-09", md)
        pay = io.open(res["files"]["支給"]["path"], encoding="cp932").read()
        self.assertIn("2026/9/25", pay)          # 支払期日は画面指定
        self.assertIn("-252", pay)               # 支援金（計算値）の預り金行


class LoadTests(unittest.TestCase):
    def test_csv_requires_columns(self):
        with self.assertRaises(ValueError):
            load_bonus_csv("社員番号,氏名\n1,a\n".encode("cp932"))

    def test_csv_cp932_roundtrip(self):
        r = _row("2018012", "山田　太郎", 100)
        header = list(r.keys())
        text = ",".join(header) + "\n" + ",".join(r[h] for h in header) + "\n"
        rows = load_bonus_csv(text.encode("cp932"))
        self.assertEqual(rows[0]["社員番号"], "2018012")

    def test_rates_csv_is_percent(self):
        """料率 CSV は必ず % 表記（0.23 は 0.23% であって 23% ではない）。適用開始年月 ≤ 支給月 の最新行を使う。"""
        import tempfile
        import os
        d = tempfile.mkdtemp()
        p = os.path.join(d, "rates.csv")
        io.open(p, "w", encoding="utf-8-sig", newline="").write(chr(10).join([
            "適用開始年月,健保率,介護率,支援金率,厚年率",
            "2026-04,9.27,1.80,0.23,18.30",
            "2027-04,9.50,1.80,0.30,18.30", ""]))
        rates, src = load_rates("2026-09", p)
        self.assertEqual(rates, {"kenpo": 0.0927, "kaigo": 0.018, "shien": 0.0023, "konen": 0.183})
        self.assertIn("2026-04", src)
        rates2, _ = load_rates("2027-06", p)
        self.assertEqual(rates2["shien"], 0.003)

    def test_rates_csv_fraction_or_bad_month_falls_back(self):
        import tempfile
        import os
        d = tempfile.mkdtemp()
        p = os.path.join(d, "rates.csv")
        io.open(p, "w", encoding="utf-8-sig", newline="").write(chr(10).join([
            "適用開始年月,健保率,介護率,支援金率,厚年率", "2026/4,9.27,1.80,0.23,18.30", ""]))
        rates, src = load_rates("2026-09", p)
        self.assertEqual(rates, DEFAULT_RATES)
        self.assertIn("2026/4", src)
        io.open(p, "w", encoding="utf-8-sig", newline="").write(chr(10).join([
            "適用開始年月,健保率,介護率,支援金率,厚年率", "2026-04,0.0927,0.018,0.0023,0.183", ""]))
        rates, src = load_rates("2026-09", p)
        self.assertEqual(rates, DEFAULT_RATES)
        self.assertIn("範囲", src)

    def test_rates_default_when_csv_missing(self):
        rates, src = load_rates("2026-09", r"C:\__no_such__\rates.csv")
        self.assertEqual(rates, DEFAULT_RATES)
        self.assertIn("既定値", src)


class GenerateBonusTests(unittest.TestCase):
    """generate_bonus が部門キャッシュを『給与側と同じ全対象者』で作ること（賞与対象者だけで共有キャッシュを潰さない）。"""

    def test_custom_histories_target_includes_all_salary_targets(self):
        import tempfile
        import os
        import json
        from unittest import mock
        from services import keiri_bonus as kb
        d = tempfile.mkdtemp()
        os.makedirs(os.path.join(d, "raw"))
        roster = [{"id": e, "company": {"last_name": "社員", "first_name": e, "joined_on": "2020-04-01"}}
                  for e in ("2018012", "2019002", "2020005", "2021009")]
        json.dump({"fetched_at": "x", "data": roster}, io.open(os.path.join(d, "raw", "roster.json"), "w", encoding="utf-8"))
        r = _row("2018012", "山田　太郎", 100000, sb=100000, kenpo=4635, shien=115, konen=9150)
        header = list(r.keys())
        csv_bytes = (",".join(header) + chr(10) + ",".join(r[h] for h in header) + chr(10)).encode("cp932")
        seen = {}

        def fake_hist(out_base, ids, refresh=False):
            seen["ids"] = set(ids)
            return {}

        with mock.patch.object(kb, "load_custom_histories", fake_hist), \
             mock.patch.object(kb, "load_or_fetch_roster", lambda client, path, refresh=False: roster):
            res = kb.generate_bonus("2026-09", csv_bytes, "FE部賞与", "2026-09-15", out_base=d, client=None,
                                    rates_csv=r"C:\__no_such__\rates.csv")
        self.assertEqual(seen["ids"], {"2018012", "2019002", "2020005", "2021009"})   # 賞与 CSV に居ない人も含む
        self.assertEqual(res["people"], 1)
        self.assertEqual(sorted(res["files"]), ["健康保険", "厚生年金", "支給"])


if __name__ == "__main__":
    unittest.main()
