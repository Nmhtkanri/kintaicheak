r"""社労士モード: jinjer 給与明細 → 社労士（前田事務所）へ渡す給与CSV（49列・cp932）。

入力は jinjer の給与明細 API **1本だけ**。statements[].basic_info に氏名と給与体系が
入っているので、従業員マスタの別取得は要らない。キャッシュは経理モードと共用する
（outputs/keiri/raw/salary_statements_{ym}.json）。jinjer への書き込みは一切しない。

⚠ 列は「番号」ではなく **列ID（スラッグ）** で指す。
  2026-08-28 に「構造的に常に空だった13列」を削って 60列 → 49列 にしたため、番号で書くと
  古い表がそのまま**別の列**を指してしまう。列マッピングCSVも `列id` をキーにしてあり、
  旧スキーマ（列番号）のファイルは読まずにエラーで止める。
  出力の並びは Layout が持つ（LAYOUT_V2 = 新49列 / LAYOUT_V1 = 旧60列）。
  LAYOUT_V1 は社労士の受け入れ確認が済むまでの控え。**済んだら V1 一式を撤去する。**

列の対応は **給与体系ごとに変わる**（月給制1/2/3・時給制1・管理監督者）。対応表は
Z:\API連携\docs\社労士モード_列マッピング_v2.csv に外出ししてあり、直せば exe の再ビルド
なしで次回実行から効く。ファイルが無いときはこのモジュール内の既定表で動く。

⚠ 立替金・その他の扱い（2026-08-14 谷津さん指示・2026-07 実データ 231 行で検証済み）
    総支給額     = 雇用保険対象額(other5) そのまま
                   ← 立替金2種・その他は元々 other5 に入らないので、これだけで報酬計になる
    口座1振込額  = 口座1振込額(payment2) − 立替金(a51) − 立替金（顧客請求分）(a50) − その他(a52)
    差引支給額   = 差引支給額(payment1) そのまま（**何も引かない**。マイナスも据え置き）
  立替金2種と「その他」（経費精算）は報酬ではないので、総支給額と口座1振込額の両方から外す。
  差引支給額には乗ったままなので、この3つを引いた額が実際の振込額になる。

⚠ 社保調整・社宅家賃・貸付金返済は V2 で専用列を持つが、**合計欄の式は変えていない**。
  社会保険料計・控除合計はどちらも source_key から直接足しているので、専用列を作ったことで
  二重計上にはならない。「合計に足し忘れでは？」と KOJO_GOKEI_KEYS を触らないこと。

⚠ 給与計算後に発生した追加支給（jinjer に入っていない）は追加支給台帳 CSV で足す。
  台帳の額は 差引支給額 と立替金列に加算し、口座1 と総支給額 は据え置く。
  これで「jinjer に立替金として入っていたら」と同じ形になる。

決定事項（2026-08-14 谷津さん）:
  1. 時給制の超過・深夜・休日の**金額**は差額調整(allowance24)に含まれるので明細列を持たない
  2. 時給制の勤怠**時間**列は見本ファイルどおり（列名と中身がずれているが直さない）
  3. 社員番号 20YY 始まり以外は出さない（7777777・9999999 とも除外）
  4. allowance15／allowance12 は全員一律で「調整手当」列
  5. 追加支給は台帳 CSV から（上記）

決定事項（2026-08-28 谷津さん）:
  6. 「欠勤控除」列(deduction1・2026-04〜08 の5か月とも全員ゼロ)を「社保調整」(deduction4)へ転用
  7. 常に空だった13列を削除。課税対象額(other7)を埋め、社宅家賃・貸付金返済に専用列を作る
  8. 貸付金返済は月によって deduction2（社宅家賃の枠）へ入る（2026-05 実績）。
     列マッピングの「社員番号条件」で正しい列へ寄せる
"""

from __future__ import annotations

import csv
import datetime
import io
import os
from collections import defaultdict

from config import Config
from services.keiri_api import classify_employee, normalize_label, to_number
from services.keiri_engine import fetch_statements, ym_compact

# ---------------------------------------------------------------------------
# 列定義
# ---------------------------------------------------------------------------
# 出力1列を指す論理ID → 社労士へ渡すCSVの見出し。
# 列の並びが変わっても意味が変わらないので、マッピング表・台帳・テストはすべてこの ID で書く
# （物理的な列番号は Layout の中にしか出てこない）。
COL_LABELS = {
    # --- 見出し ---
    "emp_no": "社員番号",
    "emp_name": "氏名",
    "salary_system": "給与体系",
    # --- 勤怠時間（給与体系ごとに中身が変わる。列名と中身のズレは見本どおり据え置き） ---
    "kihon_kinmu_jikan": "基本勤務時間",
    "minashi_jikan": "みなし時間",
    "shinya_zangyo_jikan": "深夜残業時間",
    "kyujitsu_shukkin_jikan": "休日出勤時間",
    "kyujitsu_shinya_jikan": "休日深夜時間",
    "kojo_jikan": "控除時間",
    "choka_jikan": "超過時間",
    "shukkin_nissu": "出勤日数",
    "kekkin_nissu": "欠勤日数",
    "yukyu_shutoku_nissu": "有給休暇取得日数計",
    "yukyu_zan": "有給休暇残（翌月繰越分）",
    # --- 支給 ---
    "kihonkyu": "基本給",
    "minashikyu": "みなし給",
    "yakushoku_teate": "役職手当",
    "leader_teate": "リーダー手当",
    "kyakutaio_touban_teate": "顧客対応当番手当",
    "chosei_teate": "調整手当",
    "telework_teate": "テレワーク手当",
    "teijogai_gyomu_teate": "定常外業務対応手当",
    "sonota_teate": "その他手当",
    "sagaku_chosei": "差額調整",              # allowance24
    "kazei_tsukinhi": "課税通勤費",
    "hikazei_tsukinhi": "非課税通勤費",
    "tsushin_teate": "通信手当",
    "shikyu_kabusoku_chosei": "支給過不足調整",  # allowance54
    "soushikyu_gaku": "総支給額",
    "kazei_taisho_gaku": "課税対象額",
    # --- 控除 ---
    "shaho_chosei": "社保調整",
    "koyo_hokenryo": "雇用保険料",
    "kenko_hokenryo": "健康保険料",
    "kaigo_hokenryo": "介護保険料",
    "kosei_nenkin": "厚生年金保険料",
    "kodomo_kosodate": "子ども・子育て支援金",
    "nencho_kabusoku": "年調過不足額",
    "shotokuzei": "所得税",
    "juminzei": "住民税",
    "shataku_yachin": "社宅家賃",
    "kashitsukekin_hensai": "貸付金返済",
    "kekkin_kojo": "欠勤控除",                # V1 のみ（V2 では社保調整に転用した）
    # --- 合計・振込 ---
    "shaho_kei": "社会保険料計",
    "kojo_gokei": "控除合計",
    "sashihiki_shikyu": "差引支給額",
    "koza1_furikomi": "口座1振込額",
    "tatekaekin_kyaku": "立替金（顧客請求分）",
    "tatekaekin": "立替金",
    "sonota": "その他",
    "genkin_shikyu": "現金支給額",
}
ALL_COL_IDS = frozenset(COL_LABELS)

# 氏名などマッピングで上書きされては困る列
RESERVED_COL_IDS = frozenset({"emp_no", "emp_name", "salary_system"})
# 計算で埋める列（マッピング表では指定できない）
COMPUTED_COL_IDS = frozenset({"soushikyu_gaku", "shaho_kei", "kojo_gokei",
                              "sashihiki_shikyu", "koza1_furikomi"})


class Layout(object):
    """出力CSVの列の並び。entries は (列ID または None, 見出し) の並び。

    列IDが None の要素は「見出しだけあって中身は常に空」の列（V1 の名残）。
    silent_hidden は「この形式では列を持たないのが仕様」の列ID。ここに挙げたものは
    金額があっても hidden_with_amount の警告に出さない。
    labels はこのレイアウトだけの見出し上書き（V1 は旧ファイルの見出しを保つために使う）。
    """

    def __init__(self, key, label, items, silent_hidden=(), labels=None):
        names = dict(COL_LABELS)
        for cid in (labels or {}):
            assert cid in COL_LABELS, cid
        names.update(labels or {})
        entries = []
        for it in items:
            # ('見出し',) と書いたところは「見出しだけあって中身は常に空」の列
            entries.append((None, it[0]) if isinstance(it, tuple) else (it, names[it]))
        self.key = key
        self.label = label
        self.labels = dict(labels or {})
        self.entries = tuple(entries)
        self.silent_hidden = frozenset(silent_hidden)
        self.column_names = [name for _cid, name in self.entries]
        self.col_ids = tuple(cid for cid, _name in self.entries if cid)
        self._index = {cid: i for i, (cid, _name) in enumerate(self.entries) if cid}

    def __len__(self):
        return len(self.entries)

    def __repr__(self):
        return "<Layout %s %d列>" % (self.key, len(self.entries))

    def has(self, col_id):
        return col_id in self._index

    def index_of(self, col_id):
        return self._index[col_id]


def _entries(*items):
    """列IDの並び。('見出し',) と書いたところは「常に空の列」になる。"""
    return items


# 新49列（2026-08-28〜）。常に空だった13列を削り、課税対象額・社保調整・社宅家賃・
# 貸付金返済に列を与えたもの。社労士へ渡す正本。
LAYOUT_V2 = Layout("V2", "新49列", _entries(
    "emp_no", "emp_name", "salary_system",
    "kihon_kinmu_jikan", "minashi_jikan", "shinya_zangyo_jikan",
    "kyujitsu_shukkin_jikan", "kyujitsu_shinya_jikan", "kojo_jikan", "choka_jikan",
    "shukkin_nissu", "kekkin_nissu", "yukyu_shutoku_nissu", "yukyu_zan",
    "kihonkyu", "minashikyu", "yakushoku_teate", "leader_teate",
    "kyakutaio_touban_teate", "chosei_teate", "telework_teate",
    "teijogai_gyomu_teate", "sonota_teate", "sagaku_chosei",
    "kazei_tsukinhi", "hikazei_tsukinhi", "tsushin_teate", "shikyu_kabusoku_chosei",
    "soushikyu_gaku", "kazei_taisho_gaku",
    "shaho_chosei", "koyo_hokenryo", "kenko_hokenryo", "kaigo_hokenryo",
    "kosei_nenkin", "kodomo_kosodate", "nencho_kabusoku", "shotokuzei", "juminzei",
    "shataku_yachin", "kashitsukekin_hensai",
    "shaho_kei", "kojo_gokei", "sashihiki_shikyu", "koza1_furikomi",
    "tatekaekin_kyaku", "tatekaekin", "sonota", "genkin_shikyu",
))

# 旧60列（〜2026-08 支給分）。社労士の受け入れ確認が済むまでの控え。**済んだら撤去する。**
# 中身が空の13列は jinjer に対応項目が無いか一度も金額が出たことがなく、構造的に常に空だった。
LAYOUT_V1 = Layout("V1", "旧60列", _entries(
    "emp_no", "emp_name", "salary_system",
    "kihon_kinmu_jikan", "minashi_jikan", ("超過時間",),
    "shinya_zangyo_jikan", "kyujitsu_shukkin_jikan", "kyujitsu_shinya_jikan",
    "kojo_jikan", "choka_jikan",
    "shukkin_nissu", "kekkin_nissu", "yukyu_shutoku_nissu", "yukyu_zan",
    "kihonkyu", "minashikyu",
    ("超過時間分",), ("深夜残業分",), ("休日出勤分",), ("休日深夜分",),
    ("控除精算",), ("超過精算分",), ("超過調整分",),
    "yakushoku_teate", "leader_teate", ("職能手当",), "kyakutaio_touban_teate",
    ("営業手当",), "chosei_teate", ("現場管理費",), "telework_teate",
    "teijogai_gyomu_teate", ("家賃手当",), "sonota_teate", "sagaku_chosei",
    ("欠勤控除額",), "kazei_tsukinhi", "hikazei_tsukinhi", "tsushin_teate",
    "shikyu_kabusoku_chosei", "soushikyu_gaku", ("課税対象額",),
    "kekkin_kojo", "koyo_hokenryo", "kenko_hokenryo", "kaigo_hokenryo",
    "kosei_nenkin", "kodomo_kosodate", "nencho_kabusoku", "shotokuzei", "juminzei",
    "shaho_kei", "kojo_gokei", "sashihiki_shikyu", "koza1_furikomi",
    "tatekaekin_kyaku", "tatekaekin", "sonota", "genkin_shikyu",
), silent_hidden=("kazei_taisho_gaku", "shaho_chosei",
                  "shataku_yachin", "kashitsukekin_hensai"),
   # 見出しは旧ファイルのまま。V1 は「昔どおりの形」を再現するのが役目で、見本 231 行との
   # 突合とバイト一致の土台になっている。名前を直すのは V2 だけ（2026-08-28 谷津さん）。
   labels={"sagaku_chosei": "過不足調整", "shikyu_kabusoku_chosei": "過不足調整"})

LAYOUTS = {LAYOUT_V2.key: LAYOUT_V2, LAYOUT_V1.key: LAYOUT_V1}
DEFAULT_LAYOUT_KEY = LAYOUT_V2.key

# 見出しの取り違えを import 時に落とす（上書きを宣言していない列が COL_LABELS と違わないか）
for _lay in LAYOUTS.values():
    for _cid, _name in _lay.entries:
        assert _cid is None or _cid in _lay.labels or COL_LABELS[_cid] == _name,             (_lay.key, _cid, _name)

# 計算に使う source_key
K_KOYOHOKEN = "salary_deduction_items:deduction28"
K_KENPO = "salary_deduction_items:deduction29"
K_KAIGO = "salary_deduction_items:deduction30"
K_KONEN = "salary_deduction_items:deduction31"
K_KODOMO = "salary_deduction_items:child_support"
K_SHAHO_CHOSEI = "salary_deduction_items:deduction4"
K_SHOTOKUZEI = "salary_deduction_items:deduction40"
K_JUMINZEI = "salary_deduction_items:deduction41"
K_NENCHO = "salary_deduction_items:deduction39"
K_KEKKIN = "salary_deduction_items:deduction1"
K_SHATAKU = "salary_deduction_items:deduction2"
K_KASHITSUKE = "salary_deduction_items:deduction3"
K_KOYO_TAISHO = "salary_other_items:other5"        # 雇用保険対象額＝立替金・その他を含まない支給計
K_KAZEI_TAISHO = "salary_other_items:other7"       # 課税対象額（V2 の「課税対象額」列）
K_SONOTA = "salary_items:allowance52"
K_TATEKAE_KYAKU = "salary_items:allowance50"
K_TATEKAE = "salary_items:allowance51"
K_PAYMENT1 = "salary_payment_items:payment1"
K_PAYMENT2 = "salary_payment_items:payment2"

# 社会保険料計 ＝ この6つの合計（社保調整を含む。2023019 で実証）
SHAHO_KEI_KEYS = (K_KOYOHOKEN, K_KENPO, K_KAIGO, K_KONEN, K_KODOMO, K_SHAHO_CHOSEI)
# 控除合計 ＝ 社会保険料計 ＋ この6つ
# ⚠ V2 で社保調整・社宅家賃・貸付金返済に専用列を作ったが、**この式は変えない**。
#   列はあくまで表示で、合計はここで source_key から直接足している。足すと二重計上になる。
KOJO_GOKEI_KEYS = (K_SHOTOKUZEI, K_JUMINZEI, K_NENCHO, K_KEKKIN, K_SHATAKU, K_KASHITSUKE)

# 計算式で使うので「未知項目」に数えない source_key
FORMULA_KEYS = frozenset(
    SHAHO_KEI_KEYS + KOJO_GOKEI_KEYS + (K_KOYO_TAISHO, K_SONOTA, K_PAYMENT1, K_PAYMENT2)
)

# 金額があっても意図的に出さない項目（理由つき。未知項目の判定から除外する）。
# allowance3〜6 は「前月〜」の実績・精算項目か、差額調整(allowance24)の再掲のどちらか。
# どちらも実額は差額調整の側に入っているので、明細列に出すと二重計上になる。
IGNORED_KEYS = {
    "salary_items:allowance3": "時給制の前月深夜残業分／差額調整の再掲。差額調整に含まれる",
    "salary_items:allowance4": "時給制の前月休日出勤分／差額調整の再掲。差額調整に含まれる",
    "salary_items:allowance5": "時給制の前月実績分・超過精算分／差額調整の再掲。差額調整に含まれる",
    "salary_items:allowance6": "時給制の前月精算分／差額調整の再掲。差額調整に含まれる",
    "salary_items:allowance7": "前月基礎時給。単価であって支給額ではない",
}
# allowance2 だけは体系で意味が変わるので、体系別名で「出さなくてよい形」かを見分ける。
# 月給制＝みなし給（→みなし給の列へ）／時給制の 2026-07 以前＝前月超過勤務（→出さない）。
# どちらでもない名前になったら未知項目として止める（jinjer 側の設定変更に気づくため）。
K_ALLOWANCE2 = "salary_items:allowance2"
IGNORABLE_LABEL_PREFIXES = ("前月",)

# 「みなし給」列が拾ってよい体系別名。時給制の allowance2 は 2026-07 以前が
# 「前月超過勤務」、2026-08 支給分から「みなし給」に変わる。ID ではなく体系別名で
# 判定することで、月で切らずに両方の期間を同じコードで通せる。
MINASHI_LABELS = ("当月みなし時間外手当", "当月みなし深夜手当", "みなし手当", "みなし給")

# 発生した人を毎回内訳に出す控除項目。人数が少なく、jinjer 側の入力先が揺れるので
# 「先月と同じ顔ぶれか」を人が見て確かめられるようにしておく（→ 決定事項8）。
DETAIL_DEDUCTION_COL_IDS = ("shaho_chosei", "shataku_yachin",
                            "kashitsukekin_hensai", "kekkin_kojo")
# レイアウトに列があるのに全員ゼロなら知らせる列。
# salary_other_items は WATCHED_ARRAYS に入っていないため、jinjer が other7 を移設しても
# 未知項目検知には掛からない。7,900万円の列が丸ごとゼロで社労士へ渡るのを防ぐ唯一の網。
EXPECTED_NONZERO_COL_IDS = ("kazei_taisho_gaku", "soushikyu_gaku")

# 既定の列マッピング。(列ID, 給与体系, 社員番号条件, source_key, 体系別名条件)
#   給与体系が空 = 全体系共通 / 体系別名条件が空 = 条件なし / 社員番号条件が空 = 全員
#   同じ (列ID, 給与体系) の行が複数あるときは **合計** する
#     （調整手当は 2026-08 支給分から allowance15 → allowance12 へ移設されたので
#       両方を残す。片方は必ずゼロなので合計しても二重計上にならない）
#   社員番号条件つきの行は、同じ source_key の条件なし行より **優先** する
#     （その社員に限り条件なし行を無効にする。→ 決定事項8 の貸付金返済）
DEFAULT_COLUMN_MAPPING = [
    # --- 勤怠（給与体系ごとに中身が変わる。管理監督者は勤怠列を出さない） ---
    ("kihon_kinmu_jikan", "時給制1", "", "salary_attendance_items:kintai1", ""),
    ("shinya_zangyo_jikan", "時給制1", "", "salary_attendance_items:kintai3", ""),
    ("kyujitsu_shukkin_jikan", "時給制1", "", "salary_attendance_items:kintai4", ""),
    ("kyujitsu_shinya_jikan", "時給制1", "", "salary_attendance_items:kintai3", ""),
    ("shukkin_nissu", "時給制1", "", "salary_attendance_items:kintai10", ""),
    ("kekkin_nissu", "時給制1", "", "salary_attendance_items:kintai11", ""),
    ("yukyu_shutoku_nissu", "時給制1", "", "salary_attendance_items:kintai13", ""),
    ("yukyu_zan", "時給制1", "", "salary_attendance_items:kintai14", ""),
    ("kihon_kinmu_jikan", "月給制1", "", "salary_attendance_items:kintai1", ""),
    ("minashi_jikan", "月給制1", "", "salary_attendance_items:kintai2", ""),
    ("kojo_jikan", "月給制1", "", "salary_attendance_items:kintai4", ""),
    ("choka_jikan", "月給制1", "", "salary_attendance_items:kintai5", ""),
    ("shukkin_nissu", "月給制1", "", "salary_attendance_items:kintai10", ""),
    ("kekkin_nissu", "月給制1", "", "salary_attendance_items:kintai11", ""),
    ("yukyu_shutoku_nissu", "月給制1", "", "salary_attendance_items:kintai13", ""),
    ("yukyu_zan", "月給制1", "", "salary_attendance_items:kintai14", ""),
    ("kihon_kinmu_jikan", "月給制2", "", "salary_attendance_items:kintai1", ""),
    ("minashi_jikan", "月給制2", "", "salary_attendance_items:kintai2", ""),
    ("kojo_jikan", "月給制2", "", "salary_attendance_items:kintai4", ""),
    ("choka_jikan", "月給制2", "", "salary_attendance_items:kintai5", ""),
    ("shukkin_nissu", "月給制2", "", "salary_attendance_items:kintai10", ""),
    ("kekkin_nissu", "月給制2", "", "salary_attendance_items:kintai11", ""),
    ("yukyu_shutoku_nissu", "月給制2", "", "salary_attendance_items:kintai13", ""),
    ("yukyu_zan", "月給制2", "", "salary_attendance_items:kintai14", ""),
    ("kihon_kinmu_jikan", "月給制3", "", "salary_attendance_items:kintai1", ""),
    ("minashi_jikan", "月給制3", "", "salary_attendance_items:kintai2", ""),
    ("kojo_jikan", "月給制3", "", "salary_attendance_items:kintai4", ""),
    ("choka_jikan", "月給制3", "", "salary_attendance_items:kintai5", ""),
    ("shukkin_nissu", "月給制3", "", "salary_attendance_items:kintai10", ""),
    ("kekkin_nissu", "月給制3", "", "salary_attendance_items:kintai11", ""),
    ("yukyu_shutoku_nissu", "月給制3", "", "salary_attendance_items:kintai13", ""),
    ("yukyu_zan", "月給制3", "", "salary_attendance_items:kintai14", ""),
    # --- 支給（全体系共通） ---
    ("kihonkyu", "", "", "salary_items:allowance1", ""),
    ("minashikyu", "", "", "salary_items:allowance2", "|".join(MINASHI_LABELS)),
    ("yakushoku_teate", "", "", "salary_items:allowance10", ""),
    ("leader_teate", "", "", "salary_items:allowance11", ""),
    ("kyakutaio_touban_teate", "", "", "salary_items:allowance13", ""),
    ("chosei_teate", "", "", "salary_items:allowance15", ""),   # 〜2026-07 支給分の調整手当
    ("chosei_teate", "", "", "salary_items:allowance12", ""),   # 2026-08 支給分〜の調整手当
    ("telework_teate", "", "", "salary_items:allowance18", ""),
    ("teijogai_gyomu_teate", "", "", "salary_items:allowance19", ""),
    # 「その他手当」は jinjer に2つある（allowance20 と 21）。経費の追加投入がどちらの ID に
    # 着地するかは jinjer 側のテンプレート設定が決めるため、両方を同じ列に入れて合計する。
    # 実測: 2026-07 は allowance21 に3名、2026-08 は allowance20 に2名（同月に両方は出ない）。
    ("sonota_teate", "", "", "salary_items:allowance20", ""),
    ("sonota_teate", "", "", "salary_items:allowance21", ""),
    ("sagaku_chosei", "", "", "salary_items:allowance24", ""),
    ("kazei_tsukinhi", "", "", "salary_items:allowance34", ""),
    ("hikazei_tsukinhi", "", "", "salary_items:allowance35", ""),
    ("tsushin_teate", "", "", "salary_items:allowance16", ""),
    ("shikyu_kabusoku_chosei", "", "", "salary_items:allowance54", ""),
    ("tatekaekin_kyaku", "", "", K_TATEKAE_KYAKU, ""),
    ("tatekaekin", "", "", K_TATEKAE, ""),
    ("sonota", "", "", K_SONOTA, ""),
    ("genkin_shikyu", "", "", "salary_items:allowance53", ""),
    # --- 控除（全体系共通） ---
    ("kekkin_kojo", "", "", K_KEKKIN, ""),
    ("koyo_hokenryo", "", "", K_KOYOHOKEN, ""),
    ("kenko_hokenryo", "", "", K_KENPO, ""),
    ("kaigo_hokenryo", "", "", K_KAIGO, ""),
    ("kosei_nenkin", "", "", K_KONEN, ""),
    ("kodomo_kosodate", "", "", K_KODOMO, ""),
    ("nencho_kabusoku", "", "", K_NENCHO, ""),
    ("shotokuzei", "", "", K_SHOTOKUZEI, ""),
    ("juminzei", "", "", K_JUMINZEI, ""),
    ("shaho_chosei", "", "", K_SHAHO_CHOSEI, ""),
    ("shataku_yachin", "", "", K_SHATAKU, ""),
    ("kashitsukekin_hensai", "", "", K_KASHITSUKE, ""),
    # ⚠ 役員貸付金返済は jinjer の入力先が月によって deduction2（社宅家賃の枠）へ揺れる。
    #   2026-05 実測: 2023004 の 341,659 が deduction2 に入っていた（2026-06〜08 は deduction3）。
    #   社員番号条件つきのこの行が条件なし行より優先されるので、この人の deduction2 は
    #   社宅家賃列には出ず貸付金返済列に入る。経理モードの keiri_engine.YAKUIN_LOAN と対。
    #   jinjer 側で項目を整理するときは **両方** を直すこと。
    ("kashitsukekin_hensai", "", "2023004", K_SHATAKU, ""),
    # --- その他項目 ---
    ("kazei_taisho_gaku", "", "", K_KAZEI_TAISHO, ""),
]

MAPPING_CSV_COLS = ["列id", "列名", "給与体系", "社員番号条件",
                    "source_key", "体系別名条件", "備考"]
LEDGER_COLS = ["支給月", "社員番号", "氏名", "項目", "金額", "メモ"]
# 追加支給台帳で指定できる項目 → 列ID。差引支給額にも同額を足す
LEDGER_ITEM_COL_IDS = {"立替金（顧客請求分）": "tatekaekin_kyaku", "立替金": "tatekaekin"}

# ---------------------------------------------------------------------------
# 備考（イレギュラー発生分の理由）
# ---------------------------------------------------------------------------
# 経費チェックモードの「イレギュラー経費」から手入力する5項目
# （services/keihi_payroll_import.py の MANUAL_ITEM_KEYS と同じ並び）。
# 金額は jinjer が正、理由は台帳が正。社労士は金額だけ見ても理由が分からないので、
# 本体CSVとは別ファイルの備考CSVに「誰の・何が・いくら・なぜ」を並べて渡す。
BIKO_ITEMS = {
    "定常外業務対応手当": ("salary_items:allowance19",),
    # その他手当は jinjer に2つあり、着地先はテンプレート設定次第（→ 34列目と同じ扱い）
    "その他手当": ("salary_items:allowance20", "salary_items:allowance21"),
    "現物支給": ("salary_items:allowance53",),
    "支給過不足調整": ("salary_items:allowance54",),
    "社保調整": ("salary_deduction_items:deduction4",),   # これだけ控除項目
}
BIKO_LEDGER_COLS = ["支給月", "社員番号", "氏名", "項目", "理由"]
BIKO_CSV_COLS = ["社員番号", "氏名", "給与体系", "項目", "金額", "理由"]

# 未知項目を見張る配列（金額が動く項目だけ。勤怠・会社負担・定額減税は情報項目なので見ない）
WATCHED_ARRAYS = ("salary_items", "salary_deduction_items")

EXCLUDED_SALARY_SYSTEMS = ("テスト",)
# 勤怠列を持たない給与体系（見本ファイルでも 3〜14 列は空。マッピングが無くて正常）
NO_ATTENDANCE_SYSTEMS = ("管理監督者",)


class SharoushiExportError(ValueError):
    """マッピング・台帳の不備、未知項目の検出などで処理を止めるときの例外。"""


# ---------------------------------------------------------------------------
# マスタ読み込み
# ---------------------------------------------------------------------------
def _txt(v):
    return str(v if v is not None else "").strip()


def _read_csv_rows(path):
    """BOM付きUTF-8 → cp932 → UTF-8 の順に試して行 dict のリストを返す。"""
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return list(csv.DictReader(io.StringIO(raw.decode(enc))))
        except UnicodeDecodeError:
            continue
    raise SharoushiExportError(f"文字コードを判別できませんでした: {path}")


def _read_csv_table(path):
    """BOM付きUTF-8 → cp932 → UTF-8 の順に試して (見出しリスト, 行 dict のリスト) を返す。"""
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            rd = csv.DictReader(io.StringIO(raw.decode(enc)))
            rows = list(rd)
            return list(rd.fieldnames or []), rows
        except UnicodeDecodeError:
            continue
    raise SharoushiExportError(f"文字コードを判別できませんでした: {path}")


def resolve_layout(layout):
    """レイアウトのキー（'V1'/'V2'）または Layout そのもの → Layout。"""
    if isinstance(layout, Layout):
        return layout
    key = _txt(layout).upper() or DEFAULT_LAYOUT_KEY
    if key not in LAYOUTS:
        raise SharoushiExportError(
            "CSVの形式 '%s' は使えません（使えるのは %s）"
            % (layout, " / ".join("%s=%s" % (k, v.label) for k, v in LAYOUTS.items())))
    return LAYOUTS[key]


def load_column_mapping(path=None):
    """列マッピングを読む。CSV が無ければ既定表を使う。

    ⚠ キーは **列id**（列番号ではない）。2026-08-28 に列を削って番号が総ずれしたので、
      番号で書かれた旧スキーマのファイルはエラーで止める（黙って別の列へ入るのを防ぐ）。

    戻り値 dict:
      by_col       : {列ID: [(給与体系, 社員番号条件, source_key, 体系別名条件tuple)]}
      keys         : マッピングに出てくる source_key の集合（未知項目の判定に使う）
      emp_specific : {社員番号: {社員番号条件つき行が押さえた source_key}}
      path         : 実際に使ったパス（既定表なら None）
      rows_n       : 行数
    """
    path = path if path is not None else Config.SHAROUSHI_COLUMN_MAPPING_CSV
    rows = None
    used = None
    if path and os.path.exists(path):
        fieldnames, raw = _read_csv_table(path)
        if not fieldnames:
            raise SharoushiExportError(
                f"{path} の見出し行が読めませんでした（中身が空ではありませんか）")
        if "列id" not in fieldnames:
            raise SharoushiExportError(
                f"{path} は旧スキーマ（列番号）の列マッピングです。2026-08-28 に列を削って"
                "番号がずれたため、このまま読むと値が別の列へ入ります。"
                "`python tools/build_sharoushi_masters.py` で「列id」形式に作り直してください。")
        rows = []
        for i, r in enumerate(raw, start=2):
            col_id = _txt(r.get("列id"))
            if not col_id or col_id.startswith("#"):
                continue
            if col_id not in ALL_COL_IDS:
                raise SharoushiExportError(
                    f"{path} {i}行目: 列id '{col_id}' はありません"
                    f"（使えるのは {' / '.join(sorted(ALL_COL_IDS))}）")
            if col_id in RESERVED_COL_IDS:
                raise SharoushiExportError(
                    f"{path} {i}行目: 列 {col_id}（{COL_LABELS[col_id]}）は"
                    "社員番号・氏名など固定の列なのでマッピングでは指定できません")
            if col_id in COMPUTED_COL_IDS:
                raise SharoushiExportError(
                    f"{path} {i}行目: 列 {col_id}（{COL_LABELS[col_id]}）は計算で埋める列なので"
                    "マッピングでは指定できません")
            # 列名は注記だが、列id と食い違っていたら人が誤解しているので止める
            name = _txt(r.get("列名"))
            if name and name != COL_LABELS[col_id]:
                raise SharoushiExportError(
                    f"{path} {i}行目: 列id '{col_id}' の列名は "
                    f"'{COL_LABELS[col_id]}' です（'{name}' と書かれています）")
            key = _txt(r.get("source_key"))
            if ":" not in key:
                raise SharoushiExportError(
                    f"{path} {i}行目: source_key '{key}' は "
                    "'salary_items:allowance1' の形式で書いてください")
            rows.append((col_id, _txt(r.get("給与体系")), _txt(r.get("社員番号条件")),
                         key, _txt(r.get("体系別名条件"))))
        used = path
    if rows is None:
        rows = [tuple(r) for r in DEFAULT_COLUMN_MAPPING]
    by_col = defaultdict(list)
    keys = set()
    emp_specific = defaultdict(set)
    for col_id, system, emp_cond, key, labels in rows:
        cond = tuple(normalize_label(x) for x in labels.split("|") if x.strip())
        by_col[col_id].append((system, emp_cond, key, cond))
        keys.add(key)
        if emp_cond:
            emp_specific[emp_cond].add(key)
    return {"by_col": dict(by_col), "keys": keys,
            "emp_specific": {k: frozenset(v) for k, v in emp_specific.items()},
            "path": used, "rows_n": len(rows)}


# 既定表を書き出すときに付ける備考（列id + source_key で引く）
_MAPPING_NOTES = {
    ("minashikyu", "salary_items:allowance2"):
        "体系別名で判定。時給制の「前月超過勤務」を除くため",
    ("chosei_teate", "salary_items:allowance15"): "調整手当（〜2026-07 支給分）",
    ("chosei_teate", "salary_items:allowance12"):
        "調整手当（2026-08 支給分〜。allowance15 から移設）",
    ("sonota_teate", "salary_items:allowance20"):
        "その他手当は jinjer に2つあるので両方を同じ列に入れて合計する",
    ("sonota_teate", "salary_items:allowance21"):
        "その他手当は jinjer に2つあるので両方を同じ列に入れて合計する",
    ("kazei_taisho_gaku", K_KAZEI_TAISHO):
        "課税対象額。旧60列では常に空欄だったが 2026-08-28 から埋める",
    ("shaho_chosei", K_SHAHO_CHOSEI):
        "社保調整。旧60列の「欠勤控除」列を転用した（欠勤控除は5か月とも全員ゼロ）",
    ("shataku_yachin", K_SHATAKU): "社宅家賃",
    ("kashitsukekin_hensai", K_KASHITSUKE): "貸付金返済",
    ("kashitsukekin_hensai", K_SHATAKU):
        "⚠ 役員貸付金返済。jinjer の入力先が月によって deduction2（社宅家賃の枠）へ"
        "揺れる実績があるため、この社員の deduction2 は貸付金返済列へ寄せる（2026-05 実績）。"
        "経理モードの keiri_engine.YAKUIN_LOAN と対になっているので直すときは両方見ること",
}


def export_default_mapping(path, overwrite=False):
    """コード内の既定表を列マッピングCSVとして書き出す（谷津さんが編集する原本を作る）。

    既にファイルがあるときは上書きしない（レビュー済みの表を潰さないため）。
    """
    if os.path.exists(path) and not overwrite:
        raise SharoushiExportError(
            f"既に存在します: {path}（上書きするなら overwrite=True）")
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=MAPPING_CSV_COLS)
        w.writeheader()
        for col_id, system, emp_cond, key, labels in DEFAULT_COLUMN_MAPPING:
            w.writerow({"列id": col_id, "列名": COL_LABELS[col_id], "給与体系": system,
                        "社員番号条件": emp_cond, "source_key": key,
                        "体系別名条件": labels,
                        "備考": _MAPPING_NOTES.get((col_id, key), "")})
    return path


def load_extra_ledger(path=None):
    """追加支給台帳を {(支給月, 社員番号): [{項目, 金額, メモ}]} で読む。

    給与計算のあとに発生した支給（jinjer に入っていない分）を足すための台帳。
    ファイルが無ければ空を返す＝台帳を使わなくても従来どおり動く。
    """
    path = path if path is not None else Config.SHAROUSHI_EXTRA_LEDGER_CSV
    if not path or not os.path.exists(path):
        return {}
    out = defaultdict(list)
    for i, r in enumerate(_read_csv_rows(path), start=2):
        ym, emp = _txt(r.get("支給月")), _txt(r.get("社員番号"))
        if not ym or not emp:
            continue
        item = _txt(r.get("項目"))
        if item not in LEDGER_ITEM_COL_IDS:
            raise SharoushiExportError(
                f"{path} {i}行目: 項目 '{item}' は台帳では扱えません"
                f"（使えるのは {' / '.join(LEDGER_ITEM_COL_IDS)}）")
        amount = to_number(r.get("金額"))
        if amount is None:
            raise SharoushiExportError(
                f"{path} {i}行目: {emp} の金額 '{_txt(r.get('金額'))}' が数値ではありません")
        out[(ym, emp)].append({"項目": item, "金額": amount, "メモ": _txt(r.get("メモ"))})
    return dict(out)


def load_biko_ledger(path=None):
    """備考台帳を {(支給月, 社員番号, 項目): 理由} で読む。

    同じキーの行が複数あるときは**最後の行**を採る（画面からは追記していくため）。
    ファイルが無ければ空を返す＝台帳を使わなくても本体CSVは従来どおり作れる。
    """
    path = path if path is not None else Config.SHAROUSHI_BIKO_LEDGER_CSV
    if not path or not os.path.exists(path):
        return {}
    out = {}
    for i, r in enumerate(_read_csv_rows(path), start=2):
        ym, emp, item = _txt(r.get("支給月")), _txt(r.get("社員番号")), _txt(r.get("項目"))
        if not ym or not emp:
            continue
        if item not in BIKO_ITEMS:
            raise SharoushiExportError(
                f"{path} {i}行目: 項目 '{item}' は備考台帳では扱えません"
                f"（使えるのは {' / '.join(BIKO_ITEMS)}）")
        out[(ym, emp, item)] = _txt(r.get("理由"))
    return out


def save_biko_ledger(month, entries, path=None):
    """画面で入れた理由を備考台帳へ書く（同じ 支給月×社員番号×項目 は置き換え）。

    理由が空の行は台帳から**消す**（画面で空にしたら取り消せるようにするため）。
    項目名が不正な行が1つでもあれば1行も書かずに SharoushiExportError にする。
    Returns: 書き込んだ台帳のパス
    """
    path = path if path is not None else Config.SHAROUSHI_BIKO_LEDGER_CSV
    cleaned = []
    for e in entries or []:
        emp, item = _txt(e.get("社員番号")), _txt(e.get("項目"))
        if not emp:
            raise SharoushiExportError("社員番号が空の行があります")
        if item not in BIKO_ITEMS:
            raise SharoushiExportError(
                f"{emp}: 項目 '{item}' は備考台帳では扱えません")
        cleaned.append({"支給月": month, "社員番号": emp, "氏名": _txt(e.get("氏名")),
                        "項目": item, "理由": _txt(e.get("理由"))})
    ledger = load_biko_ledger(path)
    names = {}
    for r in cleaned:
        key = (r["支給月"], r["社員番号"], r["項目"])
        if r["理由"]:
            ledger[key] = r["理由"]
        else:
            ledger.pop(key, None)
        names[(r["支給月"], r["社員番号"])] = r["氏名"]
    # 既存行の氏名は読み直せないので、書き出し時に画面から来た氏名で補う
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=BIKO_LEDGER_COLS)
        w.writeheader()
        for (ym, emp, item) in sorted(ledger):
            w.writerow({"支給月": ym, "社員番号": emp,
                        "氏名": names.get((ym, emp), ""), "項目": item,
                        "理由": ledger[(ym, emp, item)]})
    return path


def build_biko_rows(data, month, biko_ledger=None):
    """イレギュラー5項目の発生分を {行, 理由待ち} で返す。

    金額は jinjer が正（台帳には金額を持たせない＝二重管理にしない）。
    理由が入っていない行は pending として画面へ返し、社労士へ渡す前に気づけるようにする。
    """
    biko_ledger = biko_ledger or {}
    rows, pending = [], []
    for person in data or []:
        emp = _txt(person.get("employee_id"))
        if classify_employee(emp) != "target":
            continue
        st = pick_statement(person.get("statements") or [])
        if st is None:
            continue
        basic_info = st.get("basic_info") or {}
        system = _txt((basic_info.get("salary_system") or {}).get("name"))
        if system in EXCLUDED_SALARY_SYSTEMS:
            continue
        flat = flatten_payroll(st.get("payroll_info") or {})
        name = ("%s %s" % (_txt(basic_info.get("last_name")),
                           _txt(basic_info.get("first_name")))).strip()
        for item, keys in BIKO_ITEMS.items():
            amount = sum(flat.get(k, 0.0) for k in keys)
            if not amount:
                continue
            reason = biko_ledger.get((month, emp, item), "")
            row = {"社員番号": emp, "氏名": name, "給与体系": system,
                   "項目": item, "金額": amount, "理由": reason}
            rows.append(row)
            if not reason:
                pending.append(row)
    rows.sort(key=lambda r: (r["社員番号"], list(BIKO_ITEMS).index(r["項目"])))
    pending.sort(key=lambda r: (r["社員番号"], list(BIKO_ITEMS).index(r["項目"])))
    return {"rows": rows, "pending": pending}


def write_biko_csv(rows, out_path):
    """備考CSVを cp932・CRLF で書く（本体CSVと同じ体裁）。"""
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="cp932", newline="", errors="replace") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(BIKO_CSV_COLS)
        for r in rows:
            w.writerow([format_cell(r[c]) if c == "金額" else r[c] for c in BIKO_CSV_COLS])
    return out_path


def default_biko_filename(today=None):
    stamp = (today or datetime.date.today()).strftime("%Y%m%d")
    return "備考%s.csv" % stamp


# ---------------------------------------------------------------------------
# 1人分の行を組み立てる
# ---------------------------------------------------------------------------
def flatten_payroll(payroll_info):
    """payroll_info → {source_key: float}（値が数値でない項目は入れない）。"""
    flat = {}
    for array_type, items in (payroll_info or {}).items():
        if not isinstance(items, list):
            continue
        for it in items:
            if not isinstance(it, dict):
                continue
            n = to_number(it.get("value"))
            flat["%s:%s" % (array_type, it.get("id"))] = 0.0 if n is None else n
    return flat


def item_labels(payroll_info):
    """{source_key: 体系別名}（体系別名が空なら共通ラベル）。"""
    out = {}
    for array_type, items in (payroll_info or {}).items():
        if not isinstance(items, list):
            continue
        for it in items:
            if not isinstance(it, dict):
                continue
            label = (_txt(it.get("salary_system_label")) or _txt(it.get("label")))
            out["%s:%s" % (array_type, it.get("id"))] = label
    return out


def build_row(emp, basic_info, payroll_info, mapping, ledger_entries=None):
    """1人分の値を **列IDの dict** で返す（並べるのは render_row の仕事）。

    ⚠ マッピングが1行もヒットしなかった列は dict に **入れない**。社労士側の取り込みは
      空欄と 0 を区別するので、ここで 0.0 を埋めてはいけない。

    Returns: {"values": {列ID: 値}, "unknown": [...], "ledger_applied": [...], "system": str}
    """
    flat = flatten_payroll(payroll_info)
    labels = item_labels(payroll_info)
    system = _txt((basic_info or {}).get("salary_system", {}).get("name"))
    g = lambda k: flat.get(k, 0.0)                                    # noqa: E731

    values = {
        "emp_no": emp,
        "emp_name": ("%s %s" % (_txt((basic_info or {}).get("last_name")),
                                _txt((basic_info or {}).get("first_name")))).strip(),
        "salary_system": system,
    }

    # --- マッピングで埋める列 ---
    # 社員番号条件つきの行が押さえた source_key は、その社員に限り条件なし行を無効にする
    # （例: 2023004 の deduction2 は「社宅家賃」ではなく「貸付金返済」列へ入れる）。
    claimed = mapping.get("emp_specific", {}).get(emp, frozenset())
    used_keys = set()
    for col_id, specs in mapping["by_col"].items():
        total = None
        for spec_system, spec_emp, key, cond in specs:
            if spec_system and spec_system != system:
                continue
            if spec_emp:
                if spec_emp != emp:
                    continue
            elif key in claimed:
                # この社員はこの source_key を別の列に取られている。列自体は在るので
                # 0 にしておく（その人だけ空欄になって列の中身が不揃いになるのを防ぐ）。
                total = total or 0.0
                continue
            if cond and normalize_label(labels.get(key, "")) not in cond:
                continue
            total = (total or 0.0) + g(key)
            used_keys.add(key)
        if total is not None:
            values[col_id] = total

    # --- 計算で埋める列 ---
    # 総支給額は雇用保険対象額そのもの。jinjer の雇用保険対象額は 立替金2種・その他 を
    # 含まないので、それだけで「報酬でないものを除いた支給計」になっている。
    values["soushikyu_gaku"] = g(K_KOYO_TAISHO)
    values["shaho_kei"] = sum(g(k) for k in SHAHO_KEI_KEYS)
    values["kojo_gokei"] = values["shaho_kei"] + sum(g(k) for k in KOJO_GOKEI_KEYS)
    values["sashihiki_shikyu"] = g(K_PAYMENT1)
    # 差引支給額には 立替金2種・その他 が乗っているので、実際の振込額はそれらを引いた額。
    values["koza1_furikomi"] = (g(K_PAYMENT2) - g(K_TATEKAE)
                                - g(K_TATEKAE_KYAKU) - g(K_SONOTA))

    # --- 追加支給台帳（jinjer に無い後追いの支給） ---
    applied = []
    for entry in (ledger_entries or []):
        col_id = LEDGER_ITEM_COL_IDS[entry["項目"]]
        values[col_id] = (values.get(col_id) or 0.0) + entry["金額"]
        values["sashihiki_shikyu"] += entry["金額"]   # 口座1と総支給額は据え置き
        applied.append(entry)

    # --- 未知項目の検知（金額があるのに出力にも計算にも乗っていない支給・控除） ---
    unknown = []
    for key, value in flat.items():
        if not value:
            continue
        if key.split(":", 1)[0] not in WATCHED_ARRAYS:
            continue
        if key in used_keys or key in FORMULA_KEYS or key in IGNORED_KEYS:
            continue
        label = labels.get(key, "")
        # allowance2 が「みなし給」列に乗らなかった＝みなし給ではない。時給制の「前月〜」なら
        # 意図どおりなので見逃し、それ以外の名前なら設定変更を疑って止める。
        if key == K_ALLOWANCE2 and label.startswith(IGNORABLE_LABEL_PREFIXES):
            continue
        unknown.append({"source_key": key, "label": label,
                        "金額": value, "給与体系": system})
    return {"values": values, "unknown": unknown,
            "ledger_applied": applied, "system": system}


def render_row(values, layout):
    """列IDの dict → そのレイアウトの1行。列IDが None の列と未設定の列は空欄。"""
    return [values.get(col_id, "") if col_id else "" for col_id, _name in layout.entries]


def hidden_with_amount(values, layout):
    """この形式に列が無いのに金額がある項目（silent_hidden に挙げたものは除く）。

    V2 では「欠勤控除」(deduction1) が該当する。金額は控除合計には入るが明細としては
    渡らないので、発生したことに誰も気づけなくなるのを防ぐ。
    """
    out = []
    for col_id, value in values.items():
        if layout.has(col_id) or col_id in layout.silent_hidden:
            continue
        if isinstance(value, str) or not value:
            continue
        out.append({"列id": col_id, "項目": COL_LABELS[col_id], "金額": value})
    out.sort(key=lambda d: d["列id"])
    return out


def detail_deductions(values, layout):
    """毎回内訳に出す控除項目のうち、金額があるもの。

    人数が少なく jinjer 側の入力先が揺れる項目なので、「先月と同じ顔ぶれか」を
    人が見て確かめられるようにしておく（→ 決定事項8 の貸付金返済↔社宅家賃）。
    """
    out = []
    for col_id in DETAIL_DEDUCTION_COL_IDS:
        value = values.get(col_id)
        if isinstance(value, str) or not value:
            continue
        out.append({"列id": col_id, "項目": COL_LABELS[col_id], "金額": value,
                    "列あり": layout.has(col_id)})
    return out


def all_zero_columns(rows, layout):
    """レイアウトに列があるのに全員ゼロ／空の列（jinjer 側の項目移設を疑う）。"""
    out = []
    for col_id in EXPECTED_NONZERO_COL_IDS:
        if not layout.has(col_id) or not rows:
            continue
        i = layout.index_of(col_id)
        if all(not r[i] for r in rows):
            out.append({"列id": col_id, "項目": COL_LABELS[col_id]})
    return out


# ---------------------------------------------------------------------------
# 全社員ぶん
# ---------------------------------------------------------------------------
def pick_statement(statements):
    """複数明細があるときは基本給が非ゼロのものを採る（経理モードと同じ規則）。"""
    if not statements:
        return None
    best = statements[0]
    for st in statements:
        for it in ((st.get("payroll_info") or {}).get("salary_items") or []):
            if str(it.get("id")) == "allowance1" and to_number(it.get("value")):
                return st
    return best


def build_rows(data, month, mapping, ledger=None, layout=None):
    """API の生 data[] → そのレイアウトの行リスト。社員番号順に並べる。

    layout は **必ず渡すこと**（既定値を持たせていないのは V1/V2 の取り違えを防ぐため）。

    Returns: {"rows", "unknown", "excluded", "ledger_applied", "systems",
              "multi_statement", "unmapped_systems", "hidden", "details", "all_zero"}
    """
    layout = resolve_layout(layout)
    ledger = ledger or {}
    built_rows, unknown, excluded, applied, multi = [], [], [], [], []
    hidden, details = [], []
    systems = defaultdict(int)
    known_systems = mapping_systems(mapping)
    unmapped_systems = set()
    for person in data or []:
        emp = _txt(person.get("employee_id"))
        statements = person.get("statements") or []
        if classify_employee(emp) != "target":
            excluded.append({"社員番号": emp, "理由": "社員番号が 20YY 始まりではない"})
            continue
        st = pick_statement(statements)
        if st is None:
            excluded.append({"社員番号": emp, "理由": "給与明細が無い"})
            continue
        if len(statements) > 1:
            multi.append(emp)
        basic_info = st.get("basic_info") or {}
        system = _txt((basic_info.get("salary_system") or {}).get("name"))
        if system in EXCLUDED_SALARY_SYSTEMS:
            excluded.append({"社員番号": emp, "理由": f"給与体系が「{system}」"})
            continue
        built = build_row(emp, basic_info, st.get("payroll_info") or {},
                          mapping, ledger.get((month, emp)))
        if (system and system not in known_systems
                and system not in NO_ATTENDANCE_SYSTEMS):
            # 体系別（勤怠列）のマッピングが1行も無い、想定外の給与体系。新しい体系が
            # 増えると勤怠列が丸ごと空で出てしまうので、止めずに記録して画面・ログで知らせる。
            unmapped_systems.add(system)
        values = built["values"]
        name = values["emp_name"]
        built_rows.append((emp, values))
        systems[system] += 1
        for u in built["unknown"]:
            unknown.append(dict(u, 社員番号=emp, 氏名=name))
        for a in built["ledger_applied"]:
            applied.append(dict(a, 社員番号=emp, 氏名=name))
        for h in hidden_with_amount(values, layout):
            hidden.append(dict(h, 社員番号=emp, 氏名=name, 給与体系=system))
        for d in detail_deductions(values, layout):
            details.append(dict(d, 社員番号=emp, 氏名=name, 給与体系=system))
    built_rows.sort(key=lambda pair: pair[0])
    rows = [render_row(values, layout) for _emp, values in built_rows]
    return {"rows": rows, "unknown": unknown, "excluded": excluded,
            "ledger_applied": applied, "systems": dict(systems),
            "multi_statement": multi, "unmapped_systems": sorted(unmapped_systems),
            "hidden": hidden, "details": details,
            "all_zero": all_zero_columns(rows, layout)}


def mapping_systems(mapping):
    """マッピングに体系別の行がある給与体系名の集合。"""
    out = set()
    for specs in mapping["by_col"].values():
        for system, _emp, _key, _cond in specs:
            if system:
                out.add(system)
    return out


# ---------------------------------------------------------------------------
# 出力
# ---------------------------------------------------------------------------
def format_cell(value):
    """数値は小数以下の余分なゼロを落として文字列に。空欄はそのまま空欄。"""
    if value == "" or value is None:
        return ""
    if isinstance(value, str):
        return value
    text = "%.3f" % float(value)
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def write_csv(rows, out_path, layout):
    """cp932・CRLF で書く（社労士側の取り込みが cp932 のため）。

    cp932 に無い文字があると社労士側で化けるので、書く前に検出して例外にする。
    """
    layout = resolve_layout(layout)
    head = [layout.index_of(c) for c in ("emp_no", "emp_name", "salary_system")]
    bad = []
    for r in rows:
        for i in head:
            try:
                str(r[i]).encode("cp932")
            except UnicodeEncodeError:
                bad.append("%s %s" % (r[head[0]], r[head[1]]))
                break
    if bad:
        raise SharoushiExportError(
            "cp932 で表せない文字が氏名等に含まれています: " + "、".join(bad[:5])
            + "（社労士側で文字化けするため、jinjer の登録名を確認してください）")
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(layout.column_names)
        for r in rows:
            w.writerow([format_cell(v) for v in r])
    return out_path


def default_filename(month, today=None, layout=None):
    """前田事務所YYYYMMDD.csv（日付は**作成日**。社労士へ渡す既存の命名に合わせる）。

    支給月ではなく作成日なのは見本ファイル（前田事務所20260813.csv＝2026-07 支給分）の
    命名がそうだったため。何月分かは中身と出力フォルダ（{YYYYMM}）で分かる。

    ⚠ 旧60列は末尾に `_旧60列` を付ける。同じ日に新旧を両方作ったとき、同名だと後から
      作った方が前のを黙って上書きし、開いたままのダウンロードリンクが別形式のファイルを
      返してしまうため。
    """
    stamp = (today or datetime.date.today()).strftime("%Y%m%d")
    suffix = "" if resolve_layout(layout).key == DEFAULT_LAYOUT_KEY else "_旧60列"
    return "前田事務所%s%s.csv" % (stamp, suffix)


def generate(month, out_base=None, mapping_csv=None, ledger_csv=None,
             client=None, refresh=False, allow_unknown=False, filename=None,
             biko_csv=None, layout=None):
    """指定支給月（'YYYY-MM'）の社労士CSVを作る。

    未知の支給・控除項目に金額が入っていたら **既定では例外で止める**
    （jinjer 側の項目移設に気づくための検知網。allow_unknown=True で続行できる）。

    layout は 'V2'（新49列・既定）か 'V1'（旧60列・社労士の確認が済むまでの控え）。

    Returns: dict（path / filename / rows / unknown / excluded / ledger_applied / ...）
    """
    layout = resolve_layout(layout)
    mapping = load_column_mapping(mapping_csv)
    ledger = load_extra_ledger(ledger_csv)
    out_base = out_base or Config.SHAROUSHI_OUTPUT_DIR
    cache_dir = Config.KEIRI_OUTPUT_DIR          # statements キャッシュは経理モードと共用
    data = fetch_statements(cache_dir, month, client=client, refresh=refresh)
    if not data:
        raise SharoushiExportError(
            f"{month} の給与明細が取得できませんでした"
            "（給与計算がまだ実行されていない可能性があります）")
    built = build_rows(data, month, mapping, ledger, layout=layout)
    if built["unknown"] and not allow_unknown:
        detail = "、".join(
            "%s %s（%s %s）%s円" % (u["社員番号"], u["氏名"], u["source_key"],
                                   u["label"], format_cell(u["金額"]))
            for u in built["unknown"][:5])
        raise SharoushiExportError(
            f"マッピングに無い支給・控除項目に金額が入っています（{len(built['unknown'])}件）: "
            f"{detail}"
            + ("…" if len(built["unknown"]) > 5 else "")
            + "。jinjer 側で項目が移設された可能性があります。"
            "列マッピングCSVに追記してから実行し直してください。")
    out_dir = os.path.join(out_base, ym_compact(month))
    path = write_csv(built["rows"],
                     os.path.join(out_dir, filename or default_filename(month, layout=layout)),
                     layout)
    # 備考CSV（イレギュラー5項目の発生理由）。理由が空でも行は出し、pending で画面に知らせる。
    biko = build_biko_rows(data, month, load_biko_ledger(biko_csv))
    biko_path = write_biko_csv(biko["rows"], os.path.join(out_dir, default_biko_filename()))
    return {
        "month": month,
        "path": path,
        "filename": os.path.basename(path),
        "out_dir": out_dir,
        "layout": layout.key,
        "layout_label": layout.label,
        "columns": len(layout),
        "column_names": list(layout.column_names),
        "biko_path": biko_path,
        "biko_filename": os.path.basename(biko_path),
        "biko_rows": biko["rows"],
        "biko_pending": biko["pending"],
        "biko_ledger_path": (biko_csv if biko_csv is not None
                             else Config.SHAROUSHI_BIKO_LEDGER_CSV),
        "rows": len(built["rows"]),
        "systems": built["systems"],
        "unknown": built["unknown"],
        "excluded": built["excluded"],
        "ledger_applied": built["ledger_applied"],
        "multi_statement": built["multi_statement"],
        "unmapped_systems": built["unmapped_systems"],
        "hidden": built["hidden"],
        "details": built["details"],
        "all_zero": built["all_zero"],
        "mapping_path": mapping["path"],
        "mapping_rows": mapping["rows_n"],
        "ledger_path": (ledger_csv if ledger_csv is not None
                        else Config.SHAROUSHI_EXTRA_LEDGER_CSV),
    }