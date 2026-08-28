r"""社労士モード（前田事務所CSV）のユニットテスト。

純ロジックは最小の payroll_info を組み立てて確かめる。最後の RegressionTests だけは
API キャッシュ（outputs/keiri/raw/salary_statements_2026-07.json）と見本ファイルを使い、
2026-07 支給分 231 行が再現できることを見る。どちらも無ければスキップする。

⚠ 値の検証は **列ID** で書く（`built["values"]["minashikyu"]`）。列番号を直に書いてよいのは
  RegressionTests だけ — あそこは見本ファイルの物理列番号を指しているので数字に意味がある。
"""

import csv
import io
import json
import os
import tempfile
import unittest

from services.sharoushi_export import (BIKO_ITEMS, COL_LABELS, DEFAULT_COLUMN_MAPPING,
                                       LAYOUT_V1, LAYOUT_V2, SharoushiExportError,
                                       build_biko_rows, build_row, build_rows,
                                       export_default_mapping, format_cell,
                                       hidden_with_amount, load_biko_ledger,
                                       load_column_mapping, load_extra_ledger, render_row,
                                       resolve_layout, save_biko_ledger, write_biko_csv,
                                       write_csv)

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(REPO, "outputs", "keiri", "raw", "salary_statements_2026-07.json")
SAMPLE = r"C:\Users\谷津晴香\Downloads\前田事務所20260813.csv"

MAPPING = load_column_mapping(path="")          # コード内の既定表（共有CSVに依存させない）

# 本番コードとは独立に手で凍結した列名（レイアウト定義がズレたらここで落ちる）
V1_COLUMN_NAMES = [
    "社員番号", "氏名", "給与体系", "基本勤務時間", "みなし時間", "超過時間",
    "深夜残業時間", "休日出勤時間", "休日深夜時間", "控除時間", "超過時間",
    "出勤日数", "欠勤日数", "有給休暇取得日数計", "有給休暇残（翌月繰越分）",
    "基本給", "みなし給", "超過時間分", "深夜残業分", "休日出勤分", "休日深夜分",
    "控除精算", "超過精算分", "超過調整分", "役職手当", "リーダー手当", "職能手当",
    "顧客対応当番手当", "営業手当", "調整手当", "現場管理費", "テレワーク手当",
    "定常外業務対応手当", "家賃手当", "その他手当", "過不足調整", "欠勤控除額",
    "課税通勤費", "非課税通勤費", "通信手当", "過不足調整", "総支給額", "課税対象額",
    "欠勤控除", "雇用保険料", "健康保険料", "介護保険料", "厚生年金保険料",
    "子ども・子育て支援金", "年調過不足額", "所得税", "住民税", "社会保険料計",
    "控除合計", "差引支給額", "口座1振込額", "立替金（顧客請求分）", "立替金",
    "その他", "現金支給額",
]
V2_COLUMN_NAMES = [
    "社員番号", "氏名", "給与体系", "基本勤務時間", "みなし時間",
    "深夜残業時間", "休日出勤時間", "休日深夜時間", "控除時間", "超過時間",
    "出勤日数", "欠勤日数", "有給休暇取得日数計", "有給休暇残（翌月繰越分）",
    "基本給", "みなし給", "役職手当", "リーダー手当", "顧客対応当番手当",
    "調整手当", "テレワーク手当", "定常外業務対応手当", "その他手当", "過不足調整",
    "課税通勤費", "非課税通勤費", "通信手当", "過不足調整", "総支給額", "課税対象額",
    "社保調整", "雇用保険料", "健康保険料", "介護保険料", "厚生年金保険料",
    "子ども・子育て支援金", "年調過不足額", "所得税", "住民税", "社宅家賃", "貸付金返済",
    "社会保険料計", "控除合計", "差引支給額", "口座1振込額", "立替金（顧客請求分）",
    "立替金", "その他", "現金支給額",
]
# V2 の列index → V1 の列index。**本番コードから導かず手で書く**（片方のレイアウト定義が
# 1つズレても同語反復にならず気づけるようにするため）。None は V1 に相当列が無い列。
V2_TO_V1_INDEX = {
    0: 0, 1: 1, 2: 2,
    3: 3, 4: 4,
    5: 6, 6: 7, 7: 8, 8: 9, 9: 10,
    10: 11, 11: 12, 12: 13, 13: 14,
    14: 15, 15: 16,
    16: 24, 17: 25, 18: 27, 19: 29, 20: 31, 21: 32, 22: 34, 23: 35,
    24: 37, 25: 38, 26: 39, 27: 40,
    28: 41,
    29: None,                       # 課税対象額（V1 は常に空）
    30: None,                       # 社保調整（V1 の同じ位置は欠勤控除）
    31: 44, 32: 45, 33: 46, 34: 47, 35: 48, 36: 49, 37: 50, 38: 51,
    39: None, 40: None,             # 社宅家賃 / 貸付金返済
    41: 52, 42: 53, 43: 54, 44: 55,
    45: 56, 46: 57, 47: 58, 48: 59,
}


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


def _row(system="月給制1", emp="2020001", **kw):
    return build_row(emp, _basic(system), _pi(**kw), MAPPING)


def _v(built, col_id, default=""):
    """列IDで値を取る（マッピングが当たらなかった列は空欄のまま）。"""
    return built["values"].get(col_id, default)


class TatekaeTests(unittest.TestCase):
    """立替金と「その他」は総支給額と口座1から外し、差引支給額からは引かない（谷津さん指示）。"""

    def test_tatekae_excluded_from_total_and_transfer_but_not_from_net(self):
        """2026-07 友納2008003 の形: 立替金 229,292 が口座1だけ減らす。"""
        built = _row(
            system="管理監督者",
            shikyu={"allowance1": 1480000, "allowance35": 18570, "allowance51": 229292},
            sonota={"other5": 1498570},
            payment={"payment1": 1257582, "payment2": 1257582},
        )
        self.assertEqual(_v(built, "soushikyu_gaku"), 1498570)    # 立替金を含まない
        self.assertEqual(_v(built, "sashihiki_shikyu"), 1257582)  # 立替金を引かない
        self.assertEqual(_v(built, "koza1_furikomi"), 1028290)    # 1,257,582 − 229,292

    def test_customer_billed_tatekae_is_also_subtracted_from_transfer(self):
        """立替金（顧客請求分）も口座1から引く（2026-07 加藤2018012 で確認した挙動）。"""
        built = _row(
            shikyu={"allowance1": 399400, "allowance50": 398, "allowance51": 11982},
            sonota={"other5": 459400},
            payment={"payment1": 395923, "payment2": 395923},
        )
        self.assertEqual(_v(built, "koza1_furikomi"), 395923 - 11982 - 398)

    def test_sonota_is_not_counted_as_pay(self):
        """「その他」は経費精算なので報酬ではない。総支給額に入れず、口座1からも引く。"""
        built = _row(shikyu={"allowance1": 230540, "allowance52": 16500},
                     sonota={"other5": 280920},
                     payment={"payment1": 244155, "payment2": 244155})
        self.assertEqual(_v(built, "soushikyu_gaku"), 280920)          # その他を足さない
        self.assertEqual(_v(built, "sonota"), 16500)                   # 明細列には出す
        self.assertEqual(_v(built, "sashihiki_shikyu"), 244155)        # 差引からは引かない
        self.assertEqual(_v(built, "koza1_furikomi"), 244155 - 16500)  # 口座1からは引く

    def test_tatekae_and_sonota_are_both_subtracted_from_transfer(self):
        """2026-07 矢嶋2018016 の形: 立替金（顧客請求分）840 と その他 550 の両方を引く。"""
        built = _row(
            shikyu={"allowance1": 349575, "allowance50": 840, "allowance52": 550},
            sonota={"other5": 420649},
            payment={"payment1": 327816, "payment2": 327816},
        )
        self.assertEqual(_v(built, "soushikyu_gaku"), 420649)
        self.assertEqual(_v(built, "koza1_furikomi"), 327816 - 840 - 550)

    def test_negative_net_pay_is_kept_and_transfer_stays_zero(self):
        """休職などで差引支給額がマイナスの人。口座1は jinjer の 0 のまま。"""
        built = _row(kojo={"deduction29": 12978, "deduction31": 25620},
                     payment={"payment1": -50020})
        self.assertEqual(_v(built, "sashihiki_shikyu"), -50020)
        self.assertEqual(_v(built, "koza1_furikomi"), 0)


class DeductionTotalTests(unittest.TestCase):
    def test_shaho_kei_includes_shaho_chosei(self):
        """社会保険料計に社保調整(deduction4)が入る（2026-07 小池2023019 の 16,680）。"""
        built = _row(kojo={"deduction28": 1804, "deduction29": 31518, "deduction31": 62220,
                           "child_support": 782, "deduction4": 16680})
        self.assertEqual(_v(built, "shaho_kei"), 1804 + 31518 + 62220 + 782 + 16680)

    def test_kojo_gokei_is_unchanged_when_columns_are_added(self):
        """⚠ 社宅家賃・貸付金返済・社保調整に専用列を作っても、合計の式は変えない。

        合計は source_key から直接足しているので、列を作ったことで二重計上にはならない。
        「合計に足し忘れでは？」と KOJO_GOKEI_KEYS を触るとここが落ちる。
        """
        kojo = {"deduction29": 21785, "deduction31": 43005, "child_support": 540,
                "deduction28": 3662, "deduction40": 53410, "deduction41": 23100,
                "deduction2": 88000, "deduction3": 341659, "deduction4": -802}
        built = _row(kojo=kojo)
        shaho = 3662 + 21785 + 43005 + 540 + (-802)
        self.assertEqual(_v(built, "shaho_kei"), shaho)
        self.assertEqual(_v(built, "kojo_gokei"), shaho + 53410 + 23100 + 88000 + 341659)
        # 専用列にも同じ額が出る（表示と合計の両方に乗るが、合計は列を見ていない）
        self.assertEqual(_v(built, "shataku_yachin"), 88000)
        self.assertEqual(_v(built, "kashitsukekin_hensai"), 341659)
        self.assertEqual(_v(built, "shaho_chosei"), -802)


class MinashiKyuTests(unittest.TestCase):
    """allowance2 は体系で意味が変わるので、体系別名で「みなし給」列に入れるか決める。"""

    def test_monthly_minashi_goes_to_column(self):
        built = _row(shikyu={"allowance2": 123850},
                     labels={"allowance2": "当月みなし時間外手当"})
        self.assertEqual(_v(built, "minashikyu"), 123850)
        self.assertEqual(built["unknown"], [])

    def test_hourly_previous_month_overtime_is_not_output(self):
        """時給制の 2026-07 以前は「前月超過勤務」。差額調整に含まれるので出さない。"""
        built = _row(system="時給制1", shikyu={"allowance2": 692},
                     labels={"allowance2": "前月超過勤務"})
        self.assertEqual(_v(built, "minashikyu"), "")
        self.assertEqual(built["unknown"], [])

    def test_hourly_minashi_from_202608_goes_to_column(self):
        """2026-08 支給分から時給制も「みなし給」になる。月で切らず体系別名で拾う。"""
        built = _row(system="時給制1", shikyu={"allowance2": 45120},
                     labels={"allowance2": "みなし給"})
        self.assertEqual(_v(built, "minashikyu"), 45120)

    def test_unexpected_label_on_allowance2_is_flagged(self):
        """みなし給でも「前月〜」でもない名前になったら、設定変更を疑って未知項目にする。"""
        built = _row(shikyu={"allowance2": 1000}, labels={"allowance2": "謎の手当"})
        self.assertEqual([u["source_key"] for u in built["unknown"]],
                         ["salary_items:allowance2"])


class ChoseiTeateTests(unittest.TestCase):
    """調整手当は 2026-08 支給分から allowance15 → allowance12 へ移設された。"""

    def test_old_id_before_202608(self):
        built = _row(shikyu={"allowance15": 7500})
        self.assertEqual(_v(built, "chosei_teate"), 7500)
        self.assertEqual(built["unknown"], [])

    def test_new_id_from_202608(self):
        built = _row(shikyu={"allowance12": 7500})
        self.assertEqual(_v(built, "chosei_teate"), 7500)
        self.assertEqual(built["unknown"], [])

    def test_shokuno_teate_column_is_always_empty(self):
        """職能手当には振り分けない（2026-08-14 決定。見本の 2017012 は是正する）。

        V2 では列そのものが無い。V1 では見出しだけあって中身は常に空。
        """
        built = _row(shikyu={"allowance15": 10000})
        self.assertNotIn("職能手当", LAYOUT_V2.column_names)
        self.assertEqual(render_row(built["values"], LAYOUT_V1)[26], "")


class SonotaTeateTests(unittest.TestCase):
    """「その他手当」は jinjer に allowance20 と 21 の2つある（着地先はテンプレート設定次第）。"""

    def test_teijogai_gyomu_goes_to_own_column(self):
        """定常外業務対応手当(allowance19)。2026-08 支給分で初めて使われた。"""
        built = _row(shikyu={"allowance19": 30000})
        self.assertEqual(_v(built, "teijogai_gyomu_teate"), 30000)
        self.assertEqual(built["unknown"], [])

    def test_allowance20_goes_to_sonota_column(self):
        """2026-08 実測: 追加投入が allowance20 に着地した（2名 104,220円）。"""
        built = _row(shikyu={"allowance20": 60435})
        self.assertEqual(_v(built, "sonota_teate"), 60435)
        self.assertEqual(built["unknown"], [])

    def test_allowance21_goes_to_same_column(self):
        """2026-07 実測: 同じ追加投入が allowance21 に着地していた（3名）。"""
        built = _row(shikyu={"allowance21": 200000})
        self.assertEqual(_v(built, "sonota_teate"), 200000)

    def test_both_ids_are_summed(self):
        """同月に両方入ることは無いはずだが、入っても落とさず合計する。"""
        built = _row(shikyu={"allowance20": 1000, "allowance21": 2000})
        self.assertEqual(_v(built, "sonota_teate"), 3000)


class IgnoredItemTests(unittest.TestCase):
    """差額調整に含まれる前月精算系と基礎時給は、金額があっても未知項目にしない。"""

    def test_previous_month_settlement_items_are_ignored(self):
        built = _row(system="時給制1",
                     shikyu={"allowance3": 693, "allowance4": 100, "allowance5": 716,
                             "allowance6": 50, "allowance7": 2770, "allowance24": 49860})
        self.assertEqual(built["unknown"], [])
        self.assertEqual(_v(built, "sagaku_chosei"), 49860)   # 過不足調整 ← 差額調整

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
        self.assertEqual([_v(built, c) for c in ("kihon_kinmu_jikan", "minashi_jikan",
                                                 "kojo_jikan", "choka_jikan")],
                         [160, 40, 3, 11.033])
        self.assertEqual([_v(built, c) for c in ("shukkin_nissu", "yukyu_shutoku_nissu",
                                                 "yukyu_zan")], [20, 2, 37])
        # 月給制では使わない列（時給制だけのマッピング）
        self.assertEqual([_v(built, c) for c in ("shinya_zangyo_jikan",
                                                 "kyujitsu_shukkin_jikan",
                                                 "kyujitsu_shinya_jikan")], ["", "", ""])

    def test_hourly_layout_follows_sample_even_though_labels_differ(self):
        """時給制は 深夜残業時間←前月総労働時間・休日出勤時間←前月深夜労働時間・
        休日深夜時間←前月総労働時間（列名と中身がずれているが見本どおり）。"""
        built = _row(system="時給制1",
                     kintai={"kintai1": 150, "kintai3": 167.5, "kintai4": 1, "kintai10": 22})
        self.assertEqual([_v(built, c) for c in ("kihon_kinmu_jikan", "shinya_zangyo_jikan",
                                                 "kyujitsu_shukkin_jikan",
                                                 "kyujitsu_shinya_jikan", "shukkin_nissu")],
                         [150, 167.5, 1, 167.5, 22])
        self.assertEqual([_v(built, c) for c in ("minashi_jikan", "kojo_jikan",
                                                 "choka_jikan")], ["", "", ""])

    def test_manager_has_no_attendance_columns(self):
        built = _row(system="管理監督者", kintai={"kintai1": 1})
        self.assertEqual(render_row(built["values"], LAYOUT_V1)[3:15], [""] * 12)
        self.assertEqual(render_row(built["values"], LAYOUT_V2)[3:14], [""] * 11)


class ExtraLedgerTests(unittest.TestCase):
    """給与計算後の追加支給は差引支給額と立替金列に足し、口座1と総支給額は据え置く。"""

    def test_ledger_adds_to_net_and_tatekae_only(self):
        """2026-07 出澤2017012: 3,280 円を後から支払った形を再現する。"""
        entries = [{"項目": "立替金（顧客請求分）", "金額": 3280, "メモ": ""}]
        built = build_row("2017012", _basic(), _pi(
            shikyu={"allowance1": 444050}, sonota={"other5": 533000},
            payment={"payment1": 412505, "payment2": 412505}), MAPPING, entries)
        self.assertEqual(_v(built, "tatekaekin_kyaku"), 3280)
        self.assertEqual(_v(built, "sashihiki_shikyu"), 415785)
        self.assertEqual(_v(built, "koza1_furikomi"), 412505)   # 実際の振込額は変わらない
        self.assertEqual(_v(built, "soushikyu_gaku"), 533000)   # 総支給額も変わらない

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


class LayoutTests(unittest.TestCase):
    """レイアウトの列並びが設計どおりか（列名はテスト側で独立に凍結してある）。"""

    def test_v1_has_60_columns_with_frozen_names(self):
        self.assertEqual(len(LAYOUT_V1), 60)
        self.assertEqual(LAYOUT_V1.column_names, V1_COLUMN_NAMES)

    def test_v2_has_49_columns_with_frozen_names(self):
        self.assertEqual(len(LAYOUT_V2), 49)
        self.assertEqual(LAYOUT_V2.column_names, V2_COLUMN_NAMES)

    def test_col_id_sets_differ_only_as_designed(self):
        """列名ではなく **列ID** の集合差で見る（「超過時間」は V1 に2回あるため）。"""
        self.assertEqual(set(LAYOUT_V1.col_ids) - set(LAYOUT_V2.col_ids), {"kekkin_kojo"})
        self.assertEqual(set(LAYOUT_V2.col_ids) - set(LAYOUT_V1.col_ids),
                         {"shaho_chosei", "shataku_yachin",
                          "kashitsukekin_hensai", "kazei_taisho_gaku"})

    def test_v1_blank_columns_are_the_13_deleted_ones_plus_kazei_taisho(self):
        blanks = [name for cid, name in LAYOUT_V1.entries if cid is None]
        self.assertEqual(blanks, ["超過時間", "超過時間分", "深夜残業分", "休日出勤分",
                                  "休日深夜分", "控除精算", "超過精算分", "超過調整分",
                                  "職能手当", "営業手当", "現場管理費", "家賃手当",
                                  "欠勤控除額", "課税対象額"])

    def test_duplicate_display_names_have_distinct_col_ids(self):
        """同じ見出しの列があっても列IDは別（過不足調整＝差額調整／支給過不足調整）。"""
        self.assertEqual(COL_LABELS["sagaku_chosei"], COL_LABELS["shikyu_kabusoku_chosei"])
        self.assertEqual(LAYOUT_V2.column_names.count("過不足調整"), 2)
        self.assertNotEqual(LAYOUT_V2.index_of("sagaku_chosei"),
                            LAYOUT_V2.index_of("shikyu_kabusoku_chosei"))

    def test_resolve_layout(self):
        self.assertIs(resolve_layout("V1"), LAYOUT_V1)
        self.assertIs(resolve_layout("v2"), LAYOUT_V2)
        self.assertIs(resolve_layout(None), LAYOUT_V2)      # 既定は新49列
        self.assertIs(resolve_layout(LAYOUT_V1), LAYOUT_V1)
        with self.assertRaises(SharoushiExportError):
            resolve_layout("V9")


class LayoutCorrespondenceTests(unittest.TestCase):
    """同じ値から V1 と V2 を並べて、共有列のセルが1つ残らず一致するか。"""

    def _built(self):
        return _row(
            kintai={"kintai1": 160, "kintai2": 40, "kintai4": 3, "kintai5": 11.033,
                    "kintai10": 20, "kintai11": 1, "kintai13": 2, "kintai14": 37},
            shikyu={"allowance1": 300000, "allowance2": 50000, "allowance10": 20000,
                    "allowance11": 10000, "allowance13": 5000, "allowance12": 7500,
                    "allowance16": 3000, "allowance18": 4000, "allowance19": 30000,
                    "allowance20": 1000, "allowance24": 900, "allowance34": 100,
                    "allowance35": 12000, "allowance50": 840, "allowance51": 550,
                    "allowance52": 660, "allowance53": 770, "allowance54": 80},
            kojo={"deduction1": 0, "deduction2": 88000, "deduction3": 341659,
                  "deduction4": -802, "deduction28": 1804, "deduction29": 31518,
                  "deduction30": 5000, "deduction31": 62220, "child_support": 782,
                  "deduction39": 0, "deduction40": 53410, "deduction41": 23100},
            sonota={"other5": 460000, "other7": 448000},
            payment={"payment1": 300000, "payment2": 300000},
            labels={"allowance2": "当月みなし時間外手当"},
        )

    def test_shared_columns_match_cell_by_cell(self):
        values = self._built()["values"]
        r1, r2 = render_row(values, LAYOUT_V1), render_row(values, LAYOUT_V2)
        self.assertEqual(len(r1), 60)
        self.assertEqual(len(r2), 49)
        for i2, i1 in V2_TO_V1_INDEX.items():
            if i1 is None:
                continue
            self.assertEqual(r1[i1], r2[i2],
                             "V2[%d] %s と V1[%d] %s が違う"
                             % (i2, V2_COLUMN_NAMES[i2], i1, V1_COLUMN_NAMES[i1]))

    def test_shared_column_headings_match(self):
        for i2, i1 in V2_TO_V1_INDEX.items():
            if i1 is None:
                continue
            self.assertEqual(V1_COLUMN_NAMES[i1], V2_COLUMN_NAMES[i2])

    def test_v2_only_columns_are_blank_in_v1(self):
        values = self._built()["values"]
        r1 = render_row(values, LAYOUT_V1)
        self.assertEqual(r1[42], "")            # 課税対象額は V1 では常に空
        self.assertEqual(r1[43], 0)             # V1 の同じ位置は欠勤控除（今回は0）


class LayoutV2ValueTests(unittest.TestCase):
    """V2 で新しく列を持たせた4項目。"""

    def test_kazei_taisho_comes_from_other7(self):
        built = _row(sonota={"other5": 460000, "other7": 448000})
        self.assertEqual(_v(built, "kazei_taisho_gaku"), 448000)
        self.assertEqual(render_row(built["values"], LAYOUT_V2)[29], 448000)
        self.assertEqual(render_row(built["values"], LAYOUT_V1)[42], "")

    def test_shaho_chosei_column_is_deduction4(self):
        built = _row(kojo={"deduction4": -3862})
        self.assertEqual(render_row(built["values"], LAYOUT_V2)[30], -3862)

    def test_shataku_and_kashitsuke_have_own_columns(self):
        built = _row(kojo={"deduction2": 88000, "deduction3": 341659})
        r = render_row(built["values"], LAYOUT_V2)
        self.assertEqual((r[39], r[40]), (88000, 341659))

    def test_unmapped_column_is_blank_not_zero(self):
        """⚠ マッピング行が1つも当たらなかった列は 0 ではなく空欄。

        社労士側の取り込みは空欄と 0 を区別するので、ここを 0 で埋めてはいけない。
        深夜残業時間は時給制1にしかマッピングが無いので、月給制では空欄になる。
        """
        built = _row(shikyu={"allowance1": 300000})
        self.assertNotIn("shinya_zangyo_jikan", built["values"])
        self.assertEqual(render_row(built["values"], LAYOUT_V2)[5], "")
        # 全体系共通の行がある列は、jinjer に項目が無くても 0 が入る（従来どおりの挙動）
        self.assertEqual(built["values"]["yakushoku_teate"], 0)


class HiddenDeductionWarningTests(unittest.TestCase):
    """この形式に列が無いのに金額がある控除項目を知らせる。"""

    def test_kekkin_kojo_with_amount_is_warned_in_v2(self):
        built = _row(kojo={"deduction1": 5000})
        got = hidden_with_amount(built["values"], LAYOUT_V2)
        self.assertEqual([(h["列id"], h["金額"]) for h in got], [("kekkin_kojo", 5000)])

    def test_kekkin_kojo_is_silent_in_v1(self):
        """V1 には欠勤控除の列があるので警告しない。"""
        built = _row(kojo={"deduction1": 5000})
        self.assertEqual(hidden_with_amount(built["values"], LAYOUT_V1), [])

    def test_v2_only_items_are_silent_in_v1(self):
        """V1 で社保調整・社宅家賃・貸付金返済・課税対象額に列が無いのは仕様。"""
        built = _row(kojo={"deduction2": 88000, "deduction3": 341659, "deduction4": -802},
                     sonota={"other7": 448000})
        self.assertEqual(hidden_with_amount(built["values"], LAYOUT_V1), [])
        self.assertEqual(hidden_with_amount(built["values"], LAYOUT_V2), [])

    def test_zero_amount_is_not_warned(self):
        built = _row(kojo={"deduction1": 0})
        self.assertEqual(hidden_with_amount(built["values"], LAYOUT_V2), [])

    def test_all_zero_column_is_flagged(self):
        """課税対象額が全員ゼロなら jinjer 側の項目移設を疑う（未知項目検知が効かない列）。"""
        people = [{"employee_id": "2020001",
                   "statements": [{"basic_info": _basic(),
                                   "payroll_info": _pi(shikyu={"allowance1": 300000},
                                                       sonota={"other5": 300000})}]}]
        built = build_rows(people, "2026-08", MAPPING, layout=LAYOUT_V2)
        self.assertIn("課税対象額", [z["項目"] for z in built["all_zero"]])
        self.assertNotIn("総支給額", [z["項目"] for z in built["all_zero"]])


class EmployeeConditionTests(unittest.TestCase):
    """社員番号条件つきの行は、同じ source_key の条件なし行より優先する。

    2023004 の役員貸付金返済は jinjer の入力先が月によって deduction2（社宅家賃の枠）へ
    揺れる（2026-05 実測）。その人の deduction2 は貸付金返済列に入れる。
    """

    def test_loan_repayment_in_shataku_slot_goes_to_loan_column(self):
        built = _row(emp="2023004", system="管理監督者", kojo={"deduction2": 341659})
        self.assertEqual(_v(built, "kashitsukekin_hensai"), 341659)
        # 社宅家賃には出さない。ただし列は在るので空欄ではなく 0（列の中身を不揃いにしない）
        self.assertEqual(_v(built, "shataku_yachin"), 0)

    def test_other_employees_keep_shataku_in_shataku_column(self):
        built = _row(emp="2020025", kojo={"deduction2": 88000})
        self.assertEqual(_v(built, "shataku_yachin"), 88000)
        self.assertEqual(_v(built, "kashitsukekin_hensai"), 0)

    def test_loan_in_its_own_slot_still_works(self):
        built = _row(emp="2023004", system="管理監督者", kojo={"deduction3": 341659})
        self.assertEqual(_v(built, "kashitsukekin_hensai"), 341659)

    def test_total_is_the_same_whichever_slot_is_used(self):
        """入力先が揺れても控除合計は変わらない（合計は source_key を直接足すため）。"""
        a = _row(emp="2023004", system="管理監督者", kojo={"deduction2": 341659})
        b = _row(emp="2023004", system="管理監督者", kojo={"deduction3": 341659})
        self.assertEqual(_v(a, "kojo_gokei"), _v(b, "kojo_gokei"))


class MappingCsvTests(unittest.TestCase):
    HEAD = "列id,列名,給与体系,社員番号条件,source_key,体系別名条件,備考\n"

    def _write(self, d, body, head=None):
        path = os.path.join(d, "map.csv")
        with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write((self.HEAD if head is None else head) + body)
        return path

    def test_roundtrip_matches_builtin_table(self):
        """書き出した CSV を読み直すと既定表と同じマッピングになる。"""
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "map.csv")
            export_default_mapping(path)
            loaded = load_column_mapping(path)
            self.assertEqual(loaded["rows_n"], len(DEFAULT_COLUMN_MAPPING))
            self.assertEqual(loaded["by_col"], MAPPING["by_col"])
            self.assertEqual(loaded["emp_specific"], MAPPING["emp_specific"])

    def test_refuses_to_overwrite(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "map.csv")
            export_default_mapping(path)
            with self.assertRaises(SharoushiExportError):
                export_default_mapping(path)

    def test_rejects_legacy_column_number_csv(self):
        """旧スキーマ（列番号）は読まずに止める。番号がずれて別の列へ入るのを防ぐ。"""
        with tempfile.TemporaryDirectory() as d:
            path = self._write(d, "15,基本給,,salary_items:allowance1,,\n",
                               head="列番号,列名,給与体系,source_key,体系別名条件,備考\n")
            with self.assertRaises(SharoushiExportError) as cm:
                load_column_mapping(path)
            self.assertIn("build_sharoushi_masters", str(cm.exception))

    def test_rejects_unknown_col_id(self):
        with tempfile.TemporaryDirectory() as d:
            path = self._write(d, "kihonkyuu,基本給,,,salary_items:allowance1,,\n")
            with self.assertRaises(SharoushiExportError):
                load_column_mapping(path)

    def test_rejects_mismatched_column_name(self):
        """列名は注記だが、列id と食い違っていたら人が誤解しているので止める。"""
        with tempfile.TemporaryDirectory() as d:
            path = self._write(d, "kihonkyu,役職手当,,,salary_items:allowance1,,\n")
            with self.assertRaises(SharoushiExportError):
                load_column_mapping(path)

    def test_rejects_reserved_column(self):
        with tempfile.TemporaryDirectory() as d:
            path = self._write(d, "emp_name,氏名,,,salary_items:allowance1,,\n")
            with self.assertRaises(SharoushiExportError):
                load_column_mapping(path)

    def test_rejects_computed_column(self):
        with tempfile.TemporaryDirectory() as d:
            path = self._write(d, "soushikyu_gaku,総支給額,,,salary_items:allowance1,,\n")
            with self.assertRaises(SharoushiExportError):
                load_column_mapping(path)

    def test_rejects_bad_source_key(self):
        with tempfile.TemporaryDirectory() as d:
            path = self._write(d, "kihonkyu,基本給,,,allowance1,,\n")
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
        row = [""] * len(LAYOUT_V2)
        row[0], row[1], row[2] = "2020001", "𠮷田 太郎", "月給制1"   # 𠮷 は cp932 に無い
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(SharoushiExportError):
                write_csv([row], os.path.join(d, "out.csv"), LAYOUT_V2)

    def test_write_csv_header_follows_layout(self):
        with tempfile.TemporaryDirectory() as d:
            for layout, names in ((LAYOUT_V2, V2_COLUMN_NAMES), (LAYOUT_V1, V1_COLUMN_NAMES)):
                path = os.path.join(d, "out_%s.csv" % layout.key)
                write_csv([[""] * len(layout)], path, layout)
                with io.open(path, encoding="cp932", newline="") as f:
                    self.assertEqual(next(csv.reader(f)), names)


class EmployeeFilterTests(unittest.TestCase):
    """社員番号 20YY 始まりだけを出す（テスト社員 7777777・9999999 は除外）。"""

    def _person(self, emp, system="月給制1"):
        return {"employee_id": emp,
                "statements": [{"basic_info": _basic(system),
                                "payroll_info": _pi(shikyu={"allowance1": 300000})}]}

    def test_excludes_non_employee_numbers(self):
        built = build_rows([self._person("2020001"), self._person("7777777"),
                            self._person("9999999")], "2026-07", MAPPING, layout=LAYOUT_V2)
        self.assertEqual([r[0] for r in built["rows"]], ["2020001"])
        self.assertEqual(len(built["excluded"]), 2)

    def test_excludes_test_salary_system(self):
        built = build_rows([self._person("2020001", system="テスト")], "2026-07", MAPPING,
                           layout=LAYOUT_V2)
        self.assertEqual(built["rows"], [])
        self.assertIn("テスト", built["excluded"][0]["理由"])

    def test_flags_unknown_salary_system(self):
        built = build_rows([self._person("2020001", system="月給制9")], "2026-07", MAPPING,
                           layout=LAYOUT_V2)
        self.assertEqual(built["unmapped_systems"], ["月給制9"])

    def test_rows_are_sorted_by_employee_number(self):
        built = build_rows([self._person("2020005"), self._person("2007001")],
                           "2026-07", MAPPING, layout=LAYOUT_V2)
        self.assertEqual([r[0] for r in built["rows"]], ["2007001", "2020005"])

    def test_row_width_follows_layout(self):
        for layout in (LAYOUT_V1, LAYOUT_V2):
            built = build_rows([self._person("2020001")], "2026-07", MAPPING, layout=layout)
            self.assertEqual(len(built["rows"][0]), len(layout))


class RegressionTests(unittest.TestCase):
    """2026-07 支給分の見本ファイル 231 行を **旧60列(V1)で** 再現できるか。

    見本は手作業時代の 60 列ファイルなので、突合は V1 レイアウトで行う。ここが通ることが
    「列ID化のリファクタで何も壊していない」証明になるので、期待値は動かさない。
    """

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
        self.assertEqual(self.sample[0], LAYOUT_V1.column_names)

    def test_reproduces_sample_except_agreed_changes(self):
        """差分は 2026-08-14 の決定によるものだけ。

        ①職能手当→調整手当 ②テスト社員2名の除外 ③「その他」を報酬から外した分
        （見本は総支給額に足して口座1から引いていなかったので、その人だけ2列ずれる）。
        """
        ledger = {("2026-07", "2017012"): [
            {"項目": "立替金（顧客請求分）", "金額": 3280.0, "メモ": "給与計算後の申請"}]}
        built = build_rows(self.data, "2026-07", MAPPING, ledger, layout=LAYOUT_V1)
        self.assertEqual(len(built["rows"]), 231)
        self.assertEqual(built["unknown"], [], "未知の支給・控除項目が出た")
        by_emp = {r[0]: r for r in built["rows"]}

        col_sonota = LAYOUT_V1.index_of("sonota")
        expected = {("2017012", 26), ("2017012", 29)}      # 職能手当 → 調整手当
        # 「その他」がある人は 総支給額 と 口座1 が見本と変わる
        sonota_emps = {r[0] for r in built["rows"] if r[col_sonota]}
        self.assertTrue(sonota_emps, "「その他」がある人が1人もいない")
        for emp in sonota_emps:
            expected.add((emp, LAYOUT_V1.index_of("soushikyu_gaku")))
            expected.add((emp, LAYOUT_V1.index_of("koza1_furikomi")))
        diffs = set()
        for row in self.sample[1:]:
            emp = row[0].strip()
            got = by_emp.get(emp)
            if got is None:
                self.assertIn(emp, ("7777777", "9999999"))
                continue
            for i in range(3, len(LAYOUT_V1)):
                want = float(row[i]) if row[i].strip() else 0.0
                mine = float(got[i]) if str(got[i]).strip() else 0.0
                if abs(want - mine) > 0.0005:
                    diffs.add((emp, i))
        self.assertEqual(diffs, expected)


if __name__ == "__main__":
    unittest.main()
