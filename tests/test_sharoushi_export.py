r"""社労士モード（前田事務所CSV）のユニットテスト。

純ロジックは最小の payroll_info を組み立てて確かめる。最後の RegressionTests だけは
API キャッシュ（outputs/keiri/raw/salary_statements_2026-07.json）と見本ファイルを使い、
2026-07 支給分 231 行が再現できることを見る。どちらも無ければスキップする。
"""

import csv
import io
import json
import os
import tempfile
import unittest

from services.sharoushi_export import (BIKO_ITEMS, CSV_COLUMNS, COL_KOJO_GOKEI, COL_KOZA1,
                                       COL_SASHIHIKI, COL_SHAHO_KEI, COL_SOUSHIKYU,
                                       DEFAULT_COLUMN_MAPPING, SharoushiExportError,
                                       build_biko_rows, build_row, build_rows,
                                       export_default_mapping, format_cell, load_biko_ledger,
                                       load_column_mapping, load_extra_ledger, save_biko_ledger,
                                       write_biko_csv, write_csv)

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(REPO, "outputs", "keiri", "raw", "salary_statements_2026-07.json")
SAMPLE = r"C:\Users\谷津晴香\Downloads\前田事務所20260813.csv"

MAPPING = load_column_mapping(path="")          # コード内の既定表（共有CSVに依存させない）


def _items(array, pairs, labels=None):
    """{id: value} → API のかたちの配列。labels で体系別名を付けられる。"""
    labels = labels or {}
    return [{"id": k, "value": v, "salary_system_label": labels.get(k, ""), "label": ""}
            for k, v in pairs.items()]


def _pi(shikyu=None, kojo=None, kintai=None, sonota=None, payment=None, labels=None):
    """支給・控除・勤怠・会社負担・支払を持つ payroll_info を組み立てる。"""
    return {
        "salary_items": _items("salary_items", shikyu or {}, labels),
        "salary_deduction_items": _items("salary_deduction_items", kojo or {}),
        "salary_attendance_items": _items("salary_attendance_items", kintai or {}),
        "salary_other_items": _items("salary_other_items", sonota or {}),
        "salary_payment_items": _items("salary_payment_items", payment or {}),
    }


def _basic(system="月給制1", last="山田", first="太郎"):
    return {"last_name": last, "first_name": first, "salary_system": {"name": system}}


def _row(system="月給制1", **kw):
    return build_row("2020001", _basic(system), _pi(**kw), MAPPING)


class TatekaeTests(unittest.TestCase):
    """立替金は総支給額と口座1から引き、差引支給額からは引かない（谷津さん指示）。"""

    def test_tatekae_excluded_from_total_and_transfer_but_not_from_net(self):
        """2026-07 友納2008003 の形: 立替金 229,292 が口座1だけ減らす。"""
        built = _row(
            system="管理監督者",
            shikyu={"allowance1": 1480000, "allowance35": 18570, "allowance51": 229292},
            sonota={"other5": 1498570},
            payment={"payment1": 1257582, "payment2": 1257582},
        )
        r = built["row"]
        self.assertEqual(r[COL_SOUSHIKYU], 1498570)      # 立替金を含まない
        self.assertEqual(r[COL_SASHIHIKI], 1257582)      # 立替金を引かない
        self.assertEqual(r[COL_KOZA1], 1028290)          # 1,257,582 − 229,292

    def test_customer_billed_tatekae_is_also_subtracted_from_transfer(self):
        """立替金（顧客請求分）も口座1から引く（2026-07 加藤2018012 で確認した挙動）。"""
        built = _row(
            shikyu={"allowance1": 399400, "allowance50": 398, "allowance51": 11982},
            sonota={"other5": 459400},
            payment={"payment1": 395923, "payment2": 395923},
        )
        self.assertEqual(built["row"][COL_KOZA1], 395923 - 11982 - 398)

    def test_sonota_is_added_to_total_but_not_in_koyo_taisho(self):
        """総支給額 ＝ 雇用保険対象額 ＋ その他（allowance52）。"""
        built = _row(shikyu={"allowance1": 230540, "allowance52": 16500},
                     sonota={"other5": 280920})
        self.assertEqual(built["row"][COL_SOUSHIKYU], 297420)

    def test_negative_net_pay_is_kept_and_transfer_stays_zero(self):
        """休職などで差引支給額がマイナスの人。口座1は jinjer の 0 のまま。"""
        built = _row(kojo={"deduction29": 12978, "deduction31": 25620},
                     payment={"payment1": -50020})
        self.assertEqual(built["row"][COL_SASHIHIKI], -50020)
        self.assertEqual(built["row"][COL_KOZA1], 0)


class DeductionTotalTests(unittest.TestCase):
    def test_shaho_kei_includes_shaho_chosei(self):
        """社会保険料計に社保調整(deduction4)が入る（2026-07 小池2023019 の 16,680）。"""
        built = _row(kojo={"deduction28": 1804, "deduction29": 31518, "deduction31": 62220,
                           "child_support": 782, "deduction4": 16680})
        self.assertEqual(built["row"][COL_SHAHO_KEI], 1804 + 31518 + 62220 + 782 + 16680)

    def test_kojo_gokei_includes_items_without_own_column(self):
        """社宅家賃・貸付金返済は専用列が無く、控除合計の中にだけ乗る。"""
        built = _row(kojo={"deduction29": 21785, "deduction31": 43005, "child_support": 540,
                           "deduction28": 3662, "deduction40": 53410, "deduction41": 23100,
                           "deduction2": 88000})
        shaho = 3662 + 21785 + 43005 + 540
        self.assertEqual(built["row"][COL_SHAHO_KEI], shaho)
        self.assertEqual(built["row"][COL_KOJO_GOKEI], shaho + 53410 + 23100 + 88000)


class MinashiKyuTests(unittest.TestCase):
    """allowance2 は体系で意味が変わるので、体系別名で「みなし給」列に入れるか決める。"""

    COL_MINASHI = 16

    def test_monthly_minashi_goes_to_column(self):
        built = _row(shikyu={"allowance2": 123850},
                     labels={"allowance2": "当月みなし時間外手当"})
        self.assertEqual(built["row"][self.COL_MINASHI], 123850)
        self.assertEqual(built["unknown"], [])

    def test_hourly_previous_month_overtime_is_not_output(self):
        """時給制の 2026-07 以前は「前月超過勤務」。差額調整に含まれるので出さない。"""
        built = _row(system="時給制1", shikyu={"allowance2": 692},
                     labels={"allowance2": "前月超過勤務"})
        self.assertEqual(built["row"][self.COL_MINASHI], "")
        self.assertEqual(built["unknown"], [])

    def test_hourly_minashi_from_202608_goes_to_column(self):
        """2026-08 支給分から時給制も「みなし給」になる。月で切らず体系別名で拾う。"""
        built = _row(system="時給制1", shikyu={"allowance2": 45120},
                     labels={"allowance2": "みなし給"})
        self.assertEqual(built["row"][self.COL_MINASHI], 45120)

    def test_unexpected_label_on_allowance2_is_flagged(self):
        """みなし給でも「前月〜」でもない名前になったら、設定変更を疑って未知項目にする。"""
        built = _row(shikyu={"allowance2": 1000}, labels={"allowance2": "謎の手当"})
        self.assertEqual([u["source_key"] for u in built["unknown"]],
                         ["salary_items:allowance2"])


class ChoseiTeateTests(unittest.TestCase):
    """調整手当は 2026-08 支給分から allowance15 → allowance12 へ移設された。"""

    COL_CHOSEI = 29
    COL_SHOKUNO = 26

    def test_old_id_before_202608(self):
        built = _row(shikyu={"allowance15": 7500})
        self.assertEqual(built["row"][self.COL_CHOSEI], 7500)
        self.assertEqual(built["unknown"], [])

    def test_new_id_from_202608(self):
        built = _row(shikyu={"allowance12": 7500})
        self.assertEqual(built["row"][self.COL_CHOSEI], 7500)
        self.assertEqual(built["unknown"], [])

    def test_shokuno_teate_column_is_always_empty(self):
        """職能手当には振り分けない（2026-08-14 決定。見本の 2017012 は是正する）。"""
        built = _row(shikyu={"allowance15": 10000})
        self.assertEqual(built["row"][self.COL_SHOKUNO], "")


class SonotaTeateTests(unittest.TestCase):
    """「その他手当」は jinjer に allowance20 と 21 の2つある（着地先はテンプレート設定次第）。"""

    COL_JOJOGAI = 32
    COL_SONOTA = 34

    def test_teijogai_gyomu_goes_to_own_column(self):
        """定常外業務対応手当(allowance19)。2026-08 支給分で初めて使われた。"""
        built = _row(shikyu={"allowance19": 30000})
        self.assertEqual(built["row"][self.COL_JOJOGAI], 30000)
        self.assertEqual(built["unknown"], [])

    def test_allowance20_goes_to_sonota_column(self):
        """2026-08 実測: 追加投入が allowance20 に着地した（2名 104,220円）。"""
        built = _row(shikyu={"allowance20": 60435})
        self.assertEqual(built["row"][self.COL_SONOTA], 60435)
        self.assertEqual(built["unknown"], [])

    def test_allowance21_goes_to_same_column(self):
        """2026-07 実測: 同じ追加投入が allowance21 に着地していた（3名）。"""
        built = _row(shikyu={"allowance21": 200000})
        self.assertEqual(built["row"][self.COL_SONOTA], 200000)

    def test_both_ids_are_summed(self):
        """同月に両方入ることは無いはずだが、入っても落とさず合計する。"""
        built = _row(shikyu={"allowance20": 1000, "allowance21": 2000})
        self.assertEqual(built["row"][self.COL_SONOTA], 3000)


class IgnoredItemTests(unittest.TestCase):
    """差額調整に含まれる前月精算系と基礎時給は、金額があっても未知項目にしない。"""

    def test_previous_month_settlement_items_are_ignored(self):
        built = _row(system="時給制1",
                     shikyu={"allowance3": 693, "allowance4": 100, "allowance5": 716,
                             "allowance6": 50, "allowance7": 2770, "allowance24": 49860})
        self.assertEqual(built["unknown"], [])
        self.assertEqual(built["row"][35], 49860)        # 過不足調整 ← 差額調整

    def test_unmapped_allowance_with_amount_is_unknown(self):
        built = _row(shikyu={"allowance30": 5000})
        self.assertEqual(len(built["unknown"]), 1)
        self.assertEqual(built["unknown"][0]["金額"], 5000)

    def test_unmapped_allowance_with_zero_is_not_unknown(self):
        self.assertEqual(_row(shikyu={"allowance30": 0})["unknown"], [])

    def test_attendance_items_are_not_watched(self):
        """勤怠項目は情報項目なので、出力に乗らなくても未知項目にしない。"""
        built = _row(kintai={"kintai16": 150, "kintai17": 150})
        self.assertEqual(built["unknown"], [])


class AttendanceMappingTests(unittest.TestCase):
    """勤怠列は給与体系ごとに中身が違う（見本ファイルの並びをそのまま踏襲する）。"""

    def test_monthly_layout(self):
        built = _row(kintai={"kintai1": 160, "kintai2": 40, "kintai4": 3, "kintai5": 11.033,
                             "kintai10": 20, "kintai13": 2, "kintai14": 37})
        r = built["row"]
        self.assertEqual([r[3], r[4], r[9], r[10]], [160, 40, 3, 11.033])
        self.assertEqual([r[11], r[13], r[14]], [20, 2, 37])
        self.assertEqual([r[6], r[7], r[8]], ["", "", ""])   # 月給制では使わない列

    def test_hourly_layout_follows_sample_even_though_labels_differ(self):
        """時給制は 06←前月総労働時間・07←前月深夜労働時間・08←前月総労働時間（見本どおり）。"""
        built = _row(system="時給制1",
                     kintai={"kintai1": 150, "kintai3": 167.5, "kintai4": 1, "kintai10": 22})
        r = built["row"]
        self.assertEqual([r[3], r[6], r[7], r[8], r[11]], [150, 167.5, 1, 167.5, 22])
        self.assertEqual([r[4], r[9], r[10]], ["", "", ""])

    def test_manager_has_no_attendance_columns(self):
        built = _row(system="管理監督者", kintai={"kintai1": 1})
        self.assertEqual(built["row"][3:15], [""] * 12)


class ExtraLedgerTests(unittest.TestCase):
    """給与計算後の追加支給は差引支給額と立替金列に足し、口座1と総支給額は据え置く。"""

    def test_ledger_adds_to_net_and_tatekae_only(self):
        """2026-07 出澤2017012: 3,280 円を後から支払った形を再現する。"""
        entries = [{"項目": "立替金（顧客請求分）", "金額": 3280, "メモ": ""}]
        built = build_row("2017012", _basic(), _pi(
            shikyu={"allowance1": 444050}, sonota={"other5": 533000},
            payment={"payment1": 412505, "payment2": 412505}), MAPPING, entries)
        r = built["row"]
        self.assertEqual(r[56], 3280)
        self.assertEqual(r[COL_SASHIHIKI], 415785)
        self.assertEqual(r[COL_KOZA1], 412505)           # 実際の振込額は変わらない
        self.assertEqual(r[COL_SOUSHIKYU], 533000)       # 総支給額も変わらない

    def test_rejects_unknown_item_name(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "ledger.csv")
            with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write("支給月,社員番号,氏名,項目,金額,メモ\n"
                        "2026-07,2017012,出澤 信晃,役職手当,3280,\n")
            with self.assertRaises(SharoushiExportError):
                load_extra_ledger(path)

    def test_rejects_non_numeric_amount(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "ledger.csv")
            with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write("支給月,社員番号,氏名,項目,金額,メモ\n"
                        "2026-07,2017012,出澤 信晃,立替金,あとで,\n")
            with self.assertRaises(SharoushiExportError):
                load_extra_ledger(path)

    def test_missing_file_is_empty(self):
        self.assertEqual(load_extra_ledger(os.path.join(tempfile.gettempdir(), "no_such.csv")), {})


class BikoTests(unittest.TestCase):
    """イレギュラー5項目の発生理由。金額は jinjer が正、理由だけを台帳に持たせる。"""

    def _person(self, emp, shikyu=None, kojo=None, system="月給制1", name=("山田", "太郎")):
        return {"employee_id": emp,
                "statements": [{"basic_info": _basic(system, *name),
                                "payroll_info": _pi(shikyu=shikyu, kojo=kojo)}]}

    def test_collects_all_five_items(self):
        data = [self._person("2020001",
                             shikyu={"allowance19": 30000, "allowance20": 60435,
                                     "allowance53": 5280, "allowance54": 174969},
                             kojo={"deduction4": -802})]
        got = build_biko_rows(data, "2026-08")
        self.assertEqual([r["項目"] for r in got["rows"]], list(BIKO_ITEMS))
        self.assertEqual([r["金額"] for r in got["rows"]],
                         [30000, 60435, 5280, 174969, -802])

    def test_zero_amount_is_not_listed(self):
        data = [self._person("2020001", shikyu={"allowance19": 0, "allowance53": 1000})]
        got = build_biko_rows(data, "2026-08")
        self.assertEqual([r["項目"] for r in got["rows"]], ["現物支給"])

    def test_sonota_teate_sums_both_ids(self):
        data = [self._person("2020001", shikyu={"allowance20": 1000, "allowance21": 2000})]
        self.assertEqual(build_biko_rows(data, "2026-08")["rows"][0]["金額"], 3000)

    def test_reason_comes_from_ledger_and_pending_lists_the_rest(self):
        data = [self._person("2020001", shikyu={"allowance19": 30000, "allowance53": 5280})]
        ledger = {("2026-08", "2020001", "定常外業務対応手当"): "夜間対応の臨時発生"}
        got = build_biko_rows(data, "2026-08", ledger)
        self.assertEqual(got["rows"][0]["理由"], "夜間対応の臨時発生")
        self.assertEqual([r["項目"] for r in got["pending"]], ["現物支給"])

    def test_reason_is_scoped_to_the_month(self):
        """別の月に入れた理由を引っ張ってこない。"""
        data = [self._person("2020001", shikyu={"allowance19": 30000})]
        ledger = {("2026-07", "2020001", "定常外業務対応手当"): "先月の理由"}
        got = build_biko_rows(data, "2026-08", ledger)
        self.assertEqual(got["rows"][0]["理由"], "")
        self.assertEqual(len(got["pending"]), 1)

    def test_excludes_non_target_employees(self):
        data = [self._person("9999999", shikyu={"allowance19": 1000}),
                self._person("2020001", shikyu={"allowance19": 1000}, system="テスト")]
        self.assertEqual(build_biko_rows(data, "2026-08")["rows"], [])

    def test_ledger_roundtrip(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "biko.csv")
            save_biko_ledger("2026-08", [
                {"社員番号": "2020001", "氏名": "山田 太郎",
                 "項目": "現物支給", "理由": "カタログギフト"},
                {"社員番号": "2020002", "氏名": "鈴木 花子",
                 "項目": "社保調整", "理由": "資格取得月の調整"},
            ], path)
            self.assertEqual(load_biko_ledger(path), {
                ("2026-08", "2020001", "現物支給"): "カタログギフト",
                ("2026-08", "2020002", "社保調整"): "資格取得月の調整",
            })

    def test_empty_reason_removes_the_entry(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "biko.csv")
            save_biko_ledger("2026-08", [
                {"社員番号": "2020001", "氏名": "山田 太郎", "項目": "現物支給", "理由": "初回"}], path)
            save_biko_ledger("2026-08", [
                {"社員番号": "2020001", "氏名": "山田 太郎", "項目": "現物支給", "理由": ""}], path)
            self.assertEqual(load_biko_ledger(path), {})

    def test_rejects_unknown_item(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "biko.csv")
            with self.assertRaises(SharoushiExportError):
                save_biko_ledger("2026-08", [
                    {"社員番号": "2020001", "項目": "役職手当", "理由": "x"}], path)

    def test_missing_ledger_file_is_empty(self):
        self.assertEqual(load_biko_ledger(os.path.join(tempfile.gettempdir(), "nope.csv")), {})

    def test_biko_csv_is_cp932_crlf(self):
        rows = [{"社員番号": "2020001", "氏名": "山田 太郎", "給与体系": "月給制1",
                 "項目": "現物支給", "金額": 5280.0, "理由": "カタログギフト"}]
        with tempfile.TemporaryDirectory() as d:
            path = write_biko_csv(rows, os.path.join(d, "備考.csv"))
            raw = open(path, "rb").read()
            self.assertIn(b"\r\n", raw)
            text = raw.decode("cp932")
            self.assertIn("社員番号,氏名,給与体系,項目,金額,理由", text)
            self.assertIn("2020001,山田 太郎,月給制1,現物支給,5280,カタログギフト", text)


class MappingCsvTests(unittest.TestCase):
    def test_roundtrip_matches_builtin_table(self):
        """書き出した CSV を読み直すと既定表と同じマッピングになる。"""
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "map.csv")
            export_default_mapping(path)
            loaded = load_column_mapping(path)
            self.assertEqual(loaded["rows_n"], len(DEFAULT_COLUMN_MAPPING))
            self.assertEqual(loaded["by_col"], MAPPING["by_col"])

    def test_refuses_to_overwrite(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "map.csv")
            export_default_mapping(path)
            with self.assertRaises(SharoushiExportError):
                export_default_mapping(path)

    def test_rejects_computed_column(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "map.csv")
            with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write("列番号,列名,給与体系,source_key,体系別名条件,備考\n"
                        "41,総支給額,,salary_items:allowance1,,\n")
            with self.assertRaises(SharoushiExportError):
                load_column_mapping(path)

    def test_rejects_bad_source_key(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "map.csv")
            with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write("列番号,列名,給与体系,source_key,体系別名条件,備考\n"
                        "15,基本給,,allowance1,,\n")
            with self.assertRaises(SharoushiExportError):
                load_column_mapping(path)


class FormatTests(unittest.TestCase):
    def test_drops_trailing_zeros_but_keeps_real_decimals(self):
        self.assertEqual(format_cell(396150.0), "396150")
        self.assertEqual(format_cell(11.033), "11.033")
        self.assertEqual(format_cell(167.5), "167.5")
        self.assertEqual(format_cell(0.0), "0")
        self.assertEqual(format_cell(""), "")
        self.assertEqual(format_cell(-50020.0), "-50020")

    def test_write_csv_rejects_names_outside_cp932(self):
        row = [""] * len(CSV_COLUMNS)
        row[0], row[1], row[2] = "2020001", "𠮷田 太郎", "月給制1"   # 𠮷 は cp932 に無い
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(SharoushiExportError):
                write_csv([row], os.path.join(d, "out.csv"))


class EmployeeFilterTests(unittest.TestCase):
    """社員番号 20YY 始まりだけを出す（テスト社員 7777777・9999999 は除外）。"""

    def _person(self, emp, system="月給制1"):
        return {"employee_id": emp,
                "statements": [{"basic_info": _basic(system),
                                "payroll_info": _pi(shikyu={"allowance1": 300000})}]}

    def test_excludes_non_employee_numbers(self):
        built = build_rows([self._person("2020001"), self._person("7777777"),
                            self._person("9999999")], "2026-07", MAPPING)
        self.assertEqual([r[0] for r in built["rows"]], ["2020001"])
        self.assertEqual(len(built["excluded"]), 2)

    def test_excludes_test_salary_system(self):
        built = build_rows([self._person("2020001", system="テスト")], "2026-07", MAPPING)
        self.assertEqual(built["rows"], [])
        self.assertIn("テスト", built["excluded"][0]["理由"])

    def test_flags_unknown_salary_system(self):
        built = build_rows([self._person("2020001", system="月給制9")], "2026-07", MAPPING)
        self.assertEqual(built["unmapped_systems"], ["月給制9"])

    def test_rows_are_sorted_by_employee_number(self):
        built = build_rows([self._person("2020005"), self._person("2007001")],
                           "2026-07", MAPPING)
        self.assertEqual([r[0] for r in built["rows"]], ["2007001", "2020005"])


class RegressionTests(unittest.TestCase):
    """2026-07 支給分の見本ファイル 231 行を再現できるか（キャッシュと見本が要る）。"""

    @classmethod
    def setUpClass(cls):
        for path in (CACHE, SAMPLE):
            if not os.path.exists(path):
                raise unittest.SkipTest(f"見つかりません: {path}")
        with io.open(CACHE, "r", encoding="utf-8") as f:
            cls.data = json.load(f)["data"]
        with io.open(SAMPLE, "r", encoding="cp932", newline="") as f:
            cls.sample = list(csv.reader(f))

    def test_header_matches_sample(self):
        self.assertEqual(self.sample[0], CSV_COLUMNS)

    def test_reproduces_sample_except_agreed_changes(self):
        """差分は 2026-08-14 の決定によるものだけ（職能手当→調整手当、テスト社員2名の除外）。"""
        ledger = {("2026-07", "2017012"): [
            {"項目": "立替金（顧客請求分）", "金額": 3280.0, "メモ": "給与計算後の申請"}]}
        built = build_rows(self.data, "2026-07", MAPPING, ledger)
        self.assertEqual(len(built["rows"]), 231)
        self.assertEqual(built["unknown"], [], "未知の支給・控除項目が出た")
        by_emp = {r[0]: r for r in built["rows"]}

        expected = {("2017012", 26), ("2017012", 29)}      # 職能手当 → 調整手当
        diffs = set()
        for row in self.sample[1:]:
            emp = row[0].strip()
            got = by_emp.get(emp)
            if got is None:
                self.assertIn(emp, ("7777777", "9999999"))
                continue
            for i in range(3, len(CSV_COLUMNS)):
                want = float(row[i]) if row[i].strip() else 0.0
                mine = float(got[i]) if str(got[i]).strip() else 0.0
                if abs(want - mine) > 0.0005:
                    diffs.add((emp, i))
        self.assertEqual(diffs, expected)


if __name__ == "__main__":
    unittest.main()
