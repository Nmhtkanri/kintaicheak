r"""社労士モード: jinjer 給与明細 → 社労士（前田事務所）へ渡す給与CSV（60列・cp932）。

入力は jinjer の給与明細 API **1本だけ**。statements[].basic_info に氏名と給与体系が
入っているので、従業員マスタの別取得は要らない。キャッシュは経理モードと共用する
（outputs/keiri/raw/salary_statements_{ym}.json）。jinjer への書き込みは一切しない。

列の対応は **給与体系ごとに変わる**（月給制1/2/3・時給制1・管理監督者）。対応表は
Z:\API連携\docs\社労士モード_列マッピング.csv に外出ししてあり、直せば exe の再ビルド
なしで次回実行から効く。ファイルが無いときはこのモジュール内の既定表で動く。

⚠ 立替金の扱い（2026-08-14 谷津さん指示・2026-07 実データ 231 行で検証済み）
    総支給額     = 雇用保険対象額(other5) ＋ その他(allowance52)   ← 立替金は元々 other5 に入らない
    口座1振込額  = 口座1振込額(payment2) − 立替金(a51) − 立替金（顧客請求分）(a50)
    差引支給額   = 差引支給額(payment1) そのまま（**立替金を引かない**。マイナスも据え置き）
  社宅家賃・貸付金返済・社保調整には専用列が無く、社会保険料計／控除合計の中にだけ乗る。

⚠ 給与計算後に発生した追加支給（jinjer に入っていない）は追加支給台帳 CSV で足す。
  台帳の額は 差引支給額(54) と立替金列(56/57) に加算し、口座1(55) と総支給額(41) は据え置く。
  これで「jinjer に立替金として入っていたら」と同じ形になる。

決定事項（2026-08-14 谷津さん）:
  1. 時給制の超過・深夜・休日の**金額**は差額調整(allowance24)に含まれるので明細列は空のまま
  2. 時給制の勤怠**時間**列は見本ファイルどおり（列名と中身がずれているが直さない）
  3. 社員番号 20YY 始まり以外は出さない（7777777・9999999 とも除外）
  4. allowance15／allowance12 は全員一律で「調整手当」列。「職能手当」列は常に空
  5. 追加支給は台帳 CSV から（上記）
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
CSV_COLUMNS = [
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
COL_N = len(CSV_COLUMNS)          # 60

# 計算で埋める列（マッピング表では扱わない）
COL_SOUSHIKYU = 41                # 総支給額
COL_KAZEI_TAISHO = 42             # 課税対象額（見本では常に空欄。埋めない）
COL_SHAHO_KEI = 52                # 社会保険料計
COL_KOJO_GOKEI = 53               # 控除合計
COL_SASHIHIKI = 54                # 差引支給額
COL_KOZA1 = 55                    # 口座1振込額
COMPUTED_COLUMNS = (COL_SOUSHIKYU, COL_SHAHO_KEI, COL_KOJO_GOKEI, COL_SASHIHIKI, COL_KOZA1)

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
K_SONOTA = "salary_items:allowance52"
K_TATEKAE_KYAKU = "salary_items:allowance50"
K_TATEKAE = "salary_items:allowance51"
K_PAYMENT1 = "salary_payment_items:payment1"
K_PAYMENT2 = "salary_payment_items:payment2"

# 社会保険料計 ＝ この6つの合計（社保調整を含む。2023019 で実証）
SHAHO_KEI_KEYS = (K_KOYOHOKEN, K_KENPO, K_KAIGO, K_KONEN, K_KODOMO, K_SHAHO_CHOSEI)
# 控除合計 ＝ 社会保険料計 ＋ この6つ（専用列が無い社宅家賃・貸付金返済もここに乗る）
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
# 月給制＝みなし給（→16列目へ）／時給制の 2026-07 以前＝前月超過勤務（→出さない）。
# どちらでもない名前になったら未知項目として止める（jinjer 側の設定変更に気づくため）。
K_ALLOWANCE2 = "salary_items:allowance2"
IGNORABLE_LABEL_PREFIXES = ("前月",)

# 「みなし給」列が拾ってよい体系別名。時給制の allowance2 は 2026-07 以前が
# 「前月超過勤務」、2026-08 支給分から「みなし給」に変わる。ID ではなく体系別名で
# 判定することで、月で切らずに両方の期間を同じコードで通せる。
MINASHI_LABELS = ("当月みなし時間外手当", "当月みなし深夜手当", "みなし手当", "みなし給")

# 既定の列マッピング。(列番号, 給与体系, source_key, 体系別名条件)
#   給与体系が空 = 全体系共通 / 体系別名条件が空 = 条件なし
#   同じ (列番号, 給与体系) の行が複数あるときは **合計** する
#     （調整手当は 2026-08 支給分から allowance15 → allowance12 へ移設されたので
#       両方を残す。片方は必ずゼロなので合計しても二重計上にならない）
DEFAULT_COLUMN_MAPPING = [
    # --- 勤怠（給与体系ごとに中身が変わる。管理監督者は勤怠列を出さない） ---
    (3, "時給制1", "salary_attendance_items:kintai1", ""),
    (6, "時給制1", "salary_attendance_items:kintai3", ""),
    (7, "時給制1", "salary_attendance_items:kintai4", ""),
    (8, "時給制1", "salary_attendance_items:kintai3", ""),
    (11, "時給制1", "salary_attendance_items:kintai10", ""),
    (12, "時給制1", "salary_attendance_items:kintai11", ""),
    (13, "時給制1", "salary_attendance_items:kintai13", ""),
    (14, "時給制1", "salary_attendance_items:kintai14", ""),
    (3, "月給制1", "salary_attendance_items:kintai1", ""),
    (4, "月給制1", "salary_attendance_items:kintai2", ""),
    (9, "月給制1", "salary_attendance_items:kintai4", ""),
    (10, "月給制1", "salary_attendance_items:kintai5", ""),
    (11, "月給制1", "salary_attendance_items:kintai10", ""),
    (12, "月給制1", "salary_attendance_items:kintai11", ""),
    (13, "月給制1", "salary_attendance_items:kintai13", ""),
    (14, "月給制1", "salary_attendance_items:kintai14", ""),
    (3, "月給制2", "salary_attendance_items:kintai1", ""),
    (4, "月給制2", "salary_attendance_items:kintai2", ""),
    (9, "月給制2", "salary_attendance_items:kintai4", ""),
    (10, "月給制2", "salary_attendance_items:kintai5", ""),
    (11, "月給制2", "salary_attendance_items:kintai10", ""),
    (12, "月給制2", "salary_attendance_items:kintai11", ""),
    (13, "月給制2", "salary_attendance_items:kintai13", ""),
    (14, "月給制2", "salary_attendance_items:kintai14", ""),
    (3, "月給制3", "salary_attendance_items:kintai1", ""),
    (4, "月給制3", "salary_attendance_items:kintai2", ""),
    (9, "月給制3", "salary_attendance_items:kintai4", ""),
    (10, "月給制3", "salary_attendance_items:kintai5", ""),
    (11, "月給制3", "salary_attendance_items:kintai10", ""),
    (12, "月給制3", "salary_attendance_items:kintai11", ""),
    (13, "月給制3", "salary_attendance_items:kintai13", ""),
    (14, "月給制3", "salary_attendance_items:kintai14", ""),
    # --- 支給（全体系共通） ---
    (15, "", "salary_items:allowance1", ""),
    (16, "", "salary_items:allowance2", "|".join(MINASHI_LABELS)),
    (24, "", "salary_items:allowance10", ""),
    (25, "", "salary_items:allowance11", ""),
    (27, "", "salary_items:allowance13", ""),
    (29, "", "salary_items:allowance15", ""),      # 〜2026-07 支給分の調整手当
    (29, "", "salary_items:allowance12", ""),      # 2026-08 支給分〜の調整手当
    (31, "", "salary_items:allowance18", ""),
    (32, "", "salary_items:allowance19", ""),
    # 「その他手当」は jinjer に2つある（allowance20 と 21）。経費の追加投入がどちらの ID に
    # 着地するかは jinjer 側のテンプレート設定が決めるため、両方を同じ列に入れて合計する。
    # 実測: 2026-07 は allowance21 に3名、2026-08 は allowance20 に2名（同月に両方は出ない）。
    (34, "", "salary_items:allowance20", ""),
    (34, "", "salary_items:allowance21", ""),
    (35, "", "salary_items:allowance24", ""),
    (37, "", "salary_items:allowance34", ""),
    (38, "", "salary_items:allowance35", ""),
    (39, "", "salary_items:allowance16", ""),
    (40, "", "salary_items:allowance54", ""),
    (56, "", K_TATEKAE_KYAKU, ""),
    (57, "", K_TATEKAE, ""),
    (58, "", K_SONOTA, ""),
    (59, "", "salary_items:allowance53", ""),
    # --- 控除（全体系共通） ---
    (43, "", K_KEKKIN, ""),
    (44, "", K_KOYOHOKEN, ""),
    (45, "", K_KENPO, ""),
    (46, "", K_KAIGO, ""),
    (47, "", K_KONEN, ""),
    (48, "", K_KODOMO, ""),
    (49, "", K_NENCHO, ""),
    (50, "", K_SHOTOKUZEI, ""),
    (51, "", K_JUMINZEI, ""),
]

MAPPING_CSV_COLS = ["列番号", "列名", "給与体系", "source_key", "体系別名条件", "備考"]
LEDGER_COLS = ["支給月", "社員番号", "氏名", "項目", "金額", "メモ"]
# 追加支給台帳で指定できる項目 → 列番号。差引支給額にも同額を足す
LEDGER_ITEM_COLUMNS = {"立替金（顧客請求分）": 56, "立替金": 57}

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


def load_column_mapping(path=None):
    """列マッピングを読む。CSV が無ければ既定表を使う。

    戻り値 dict:
      by_col   : {列番号: [(給与体系, source_key, 体系別名条件tuple)]}
      keys     : マッピングに出てくる source_key の集合（未知項目の判定に使う）
      path     : 実際に使ったパス（既定表なら None）
      rows_n   : 行数
    """
    path = path if path is not None else Config.SHAROUSHI_COLUMN_MAPPING_CSV
    rows = None
    used = None
    if path and os.path.exists(path):
        raw = _read_csv_rows(path)
        rows = []
        for i, r in enumerate(raw, start=2):
            col_txt = _txt(r.get("列番号"))
            if not col_txt or col_txt.startswith("#"):
                continue
            try:
                col = int(col_txt)
            except ValueError:
                raise SharoushiExportError(
                    f"{path} {i}行目: 列番号 '{col_txt}' が数値ではありません")
            if not 0 <= col < COL_N:
                raise SharoushiExportError(
                    f"{path} {i}行目: 列番号 {col} は 0〜{COL_N - 1} の範囲外です")
            if col in COMPUTED_COLUMNS:
                raise SharoushiExportError(
                    f"{path} {i}行目: 列 {col}（{CSV_COLUMNS[col]}）は計算で埋める列なので"
                    "マッピングでは指定できません")
            key = _txt(r.get("source_key"))
            if ":" not in key:
                raise SharoushiExportError(
                    f"{path} {i}行目: source_key '{key}' は "
                    "'salary_items:allowance1' の形式で書いてください")
            rows.append((col, _txt(r.get("給与体系")), key, _txt(r.get("体系別名条件"))))
        used = path
    if rows is None:
        rows = [(c, s, k, lab) for c, s, k, lab in DEFAULT_COLUMN_MAPPING]
    by_col = defaultdict(list)
    keys = set()
    for col, system, key, labels in rows:
        cond = tuple(normalize_label(x) for x in labels.split("|") if x.strip())
        by_col[col].append((system, key, cond))
        keys.add(key)
    return {"by_col": dict(by_col), "keys": keys, "path": used, "rows_n": len(rows)}


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
        for col, system, key, labels in DEFAULT_COLUMN_MAPPING:
            note = ""
            if col == 16:
                note = "体系別名で判定。時給制の「前月超過勤務」を除くため"
            elif key == "salary_items:allowance15":
                note = "調整手当（〜2026-07 支給分）"
            elif key == "salary_items:allowance12":
                note = "調整手当（2026-08 支給分〜。allowance15 から移設）"
            elif key in ("salary_items:allowance20", "salary_items:allowance21"):
                note = "その他手当は jinjer に2つあるので両方を同じ列に入れて合計する"
            w.writerow({"列番号": col, "列名": CSV_COLUMNS[col], "給与体系": system,
                        "source_key": key, "体系別名条件": labels, "備考": note})
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
        if item not in LEDGER_ITEM_COLUMNS:
            raise SharoushiExportError(
                f"{path} {i}行目: 項目 '{item}' は台帳では扱えません"
                f"（使えるのは {' / '.join(LEDGER_ITEM_COLUMNS)}）")
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
    """1人分の 60 列と、その人の未知項目・台帳適用結果を返す。

    Returns: {"row": [...], "unknown": [...], "ledger_applied": [...], "system": str}
    """
    flat = flatten_payroll(payroll_info)
    labels = item_labels(payroll_info)
    system = _txt((basic_info or {}).get("salary_system", {}).get("name"))
    g = lambda k: flat.get(k, 0.0)                                    # noqa: E731

    row = [""] * COL_N
    row[0] = emp
    row[1] = ("%s %s" % (_txt((basic_info or {}).get("last_name")),
                         _txt((basic_info or {}).get("first_name")))).strip()
    row[2] = system

    # --- マッピングで埋める列 ---
    used_keys = set()
    for col, specs in mapping["by_col"].items():
        total = None
        for spec_system, key, cond in specs:
            if spec_system and spec_system != system:
                continue
            if cond and normalize_label(labels.get(key, "")) not in cond:
                continue
            total = (total or 0.0) + g(key)
            used_keys.add(key)
        if total is not None:
            row[col] = total

    # --- 計算で埋める列 ---
    # 総支給額は雇用保険対象額（＝立替金2種・その他を含まない支給計）＋その他。
    # 立替金が引かれた形になるのはこのため。
    row[COL_SOUSHIKYU] = g(K_KOYO_TAISHO) + g(K_SONOTA)
    row[COL_SHAHO_KEI] = sum(g(k) for k in SHAHO_KEI_KEYS)
    row[COL_KOJO_GOKEI] = row[COL_SHAHO_KEI] + sum(g(k) for k in KOJO_GOKEI_KEYS)
    row[COL_SASHIHIKI] = g(K_PAYMENT1)
    row[COL_KOZA1] = g(K_PAYMENT2) - g(K_TATEKAE) - g(K_TATEKAE_KYAKU)

    # --- 追加支給台帳（jinjer に無い後追いの支給） ---
    applied = []
    for entry in (ledger_entries or []):
        col = LEDGER_ITEM_COLUMNS[entry["項目"]]
        row[col] = (row[col] or 0.0) + entry["金額"]
        row[COL_SASHIHIKI] += entry["金額"]       # 口座1と総支給額は据え置き
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
        # allowance2 が 16列目に乗らなかった＝みなし給ではない。時給制の「前月〜」なら
        # 意図どおりなので見逃し、それ以外の名前なら設定変更を疑って止める。
        if key == K_ALLOWANCE2 and label.startswith(IGNORABLE_LABEL_PREFIXES):
            continue
        unknown.append({"source_key": key, "label": label,
                        "金額": value, "給与体系": system})
    return {"row": row, "unknown": unknown, "ledger_applied": applied, "system": system}


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


def build_rows(data, month, mapping, ledger=None):
    """API の生 data[] → 行リスト。社員番号順に並べる。

    Returns: {"rows", "unknown", "excluded", "ledger_applied", "systems", "multi_statement"}
    """
    ledger = ledger or {}
    rows, unknown, excluded, applied, multi = [], [], [], [], []
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
        rows.append(built["row"])
        systems[system] += 1
        for u in built["unknown"]:
            unknown.append(dict(u, 社員番号=emp, 氏名=built["row"][1]))
        for a in built["ledger_applied"]:
            applied.append(dict(a, 社員番号=emp, 氏名=built["row"][1]))
    rows.sort(key=lambda r: r[0])
    return {"rows": rows, "unknown": unknown, "excluded": excluded,
            "ledger_applied": applied, "systems": dict(systems),
            "multi_statement": multi, "unmapped_systems": sorted(unmapped_systems)}


def mapping_systems(mapping):
    """マッピングに体系別の行がある給与体系名の集合。"""
    out = set()
    for specs in mapping["by_col"].values():
        for system, _key, _cond in specs:
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


def write_csv(rows, out_path):
    """cp932・CRLF で書く（社労士側の取り込みが cp932 のため）。

    cp932 に無い文字があると社労士側で化けるので、書く前に検出して例外にする。
    """
    bad = []
    for r in rows:
        for value in (r[0], r[1], r[2]):
            try:
                str(value).encode("cp932")
            except UnicodeEncodeError:
                bad.append("%s %s" % (r[0], r[1]))
                break
    if bad:
        raise SharoushiExportError(
            "cp932 で表せない文字が氏名等に含まれています: " + "、".join(bad[:5])
            + "（社労士側で文字化けするため、jinjer の登録名を確認してください）")
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(CSV_COLUMNS)
        for r in rows:
            w.writerow([format_cell(v) for v in r])
    return out_path


def default_filename(month, today=None):
    """前田事務所YYYYMMDD.csv（日付は**作成日**。社労士へ渡す既存の命名に合わせる）。

    支給月ではなく作成日なのは見本ファイル（前田事務所20260813.csv＝2026-07 支給分）の
    命名がそうだったため。何月分かは中身と出力フォルダ（{YYYYMM}）で分かる。
    """
    stamp = (today or datetime.date.today()).strftime("%Y%m%d")
    return "前田事務所%s.csv" % stamp


def generate(month, out_base=None, mapping_csv=None, ledger_csv=None,
             client=None, refresh=False, allow_unknown=False, filename=None,
             biko_csv=None):
    """指定支給月（'YYYY-MM'）の社労士CSVを作る。

    未知の支給・控除項目に金額が入っていたら **既定では例外で止める**
    （jinjer 側の項目移設に気づくための検知網。allow_unknown=True で続行できる）。

    Returns: dict（path / filename / rows / unknown / excluded / ledger_applied / ...）
    """
    mapping = load_column_mapping(mapping_csv)
    ledger = load_extra_ledger(ledger_csv)
    out_base = out_base or Config.SHAROUSHI_OUTPUT_DIR
    cache_dir = Config.KEIRI_OUTPUT_DIR          # statements キャッシュは経理モードと共用
    data = fetch_statements(cache_dir, month, client=client, refresh=refresh)
    if not data:
        raise SharoushiExportError(
            f"{month} の給与明細が取得できませんでした"
            "（給与計算がまだ実行されていない可能性があります）")
    built = build_rows(data, month, mapping, ledger)
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
    path = write_csv(built["rows"], os.path.join(out_dir, filename or default_filename(month)))
    # 備考CSV（イレギュラー5項目の発生理由）。理由が空でも行は出し、pending で画面に知らせる。
    biko = build_biko_rows(data, month, load_biko_ledger(biko_csv))
    biko_path = write_biko_csv(biko["rows"], os.path.join(out_dir, default_biko_filename()))
    return {
        "month": month,
        "path": path,
        "filename": os.path.basename(path),
        "out_dir": out_dir,
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
        "mapping_path": mapping["path"],
        "mapping_rows": mapping["rows_n"],
        "ledger_path": (ledger_csv if ledger_csv is not None
                        else Config.SHAROUSHI_EXTRA_LEDGER_CSV),
    }
