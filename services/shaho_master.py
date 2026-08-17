r"""標準報酬月額チェックのマスタ読み込み（等級表Excel・料率・報酬分類CSV）。

等級表: Z:\API連携\標準月額資料\令和8年度_標準報酬月額表_関東IT健保_協会けんぽ東京比較.xlsx
  - シート「標準報酬月額表」: 健保等級1〜50。ヘッダは4行目、データは5行目〜。
    下限「以上」・上限「未満」の半開区間で、最終等級だけ上限が空欄（＝上限なし）。
  - シート「設定・出典」: 料率8行（区分/保険者/項目/全体料率/本人負担率/適用・備考/出典URL）
    ＋ I列以降に厚生年金の等級マスタ（1〜32、88,000〜650,000）。

**マスタが壊れている限り1円も計算しない。** 列がズレたまま等級を引くと、全員の保険料が
静かに間違う。そのため読み込み時に総当たりで検証し、少しでも違えば ShahoMasterError を
投げる（health_hpm_master.py と同じ思想）:
  - 見出し名でパースする（列位置の決め打ちをしない。翌年度にExcelを差し替えても動く）
  - 等級の連番・下限昇順・半開区間の連続性（前の行の上限 = 次の行の下限）
  - 健保の各帯が厚年等級マスタと矛盾しないこと
  - 保険料額列 = 標準報酬月額 × 本人負担率 が全行で一致すること（±0.51円）
  - 年度表記（令和N年度）と実行側の対象年度の照合。ズレていたら停止

保険者は関東IT健保（its）が既定。協会けんぽ東京（kyokai_tokyo）は設定で切替できる
（実測: 控除 28,737÷620,000 = 4.635% = ITS本人負担率と完全一致。freee仕訳の取引先も関東IT健保）。
"""

from __future__ import annotations

import csv
import io
import os
import unicodedata
from dataclasses import dataclass, field

import openpyxl

SHEET_GRADES = "標準報酬月額表"
SHEET_RATES = "設定・出典"
GRADE_HEADER_ROW = 4
RATE_HEADER_ROW = 4

# 保険者キー → 等級表シートの保険料額列のプレフィックス
INSURERS = {"its": "ITS", "kyokai_tokyo": "協会東京"}
# 「設定・出典」の保険者名 → 保険者キー（「共通」は全保険者）
RATE_INSURER_NAMES = {"関東IT健保": "its", "協会けんぽ東京": "kyokai_tokyo", "共通": "共通"}
# 「設定・出典」の項目名 → 内部キー
RATE_ITEM_KEYS = {"健康保険": "kenpo", "子ども・子育て支援金": "kodomo",
                  "介護保険": "kaigo", "厚生年金": "konen", "厚生年金（参照）": "konen"}

REIWA_BASE_YEAR = 2018            # 令和N年 = 西暦 N + 2018
PREMIUM_TOLERANCE = 0.51          # 保険料額列の検証誤差（端数処理前の生floatのため）


class ShahoMasterError(ValueError):
    """マスタが読めない・壊れている。計算は必ず止める。"""


@dataclass(frozen=True)
class GradeRow:
    """等級表の1行。健保の帯（半開区間）と、対応する健保・厚年の標準報酬月額。"""
    kenpo_grade: int
    kenpo_smr: int                 # 健保 標準報酬月額（円）
    lower: int                     # 報酬月額下限（以上）
    upper: int | None              # 報酬月額上限（未満）。最終等級は None（上限なし）
    konen_grade: int
    konen_smr: int                 # 厚年 標準報酬月額（円）

    def contains(self, amount: float) -> bool:
        return amount >= self.lower and (self.upper is None or amount < self.upper)


@dataclass(frozen=True)
class Rate:
    """本人負担率（全体率の半分）。適用・備考と出典はレポートの「使用設定・出典」に出す。"""
    total: float                   # 全体料率
    employee: float                # 本人負担率
    applies: str = ""              # 適用・備考（例: 令和8年4月分～）
    source_url: str = ""


@dataclass
class ShahoMaster:
    path: str
    year: int                      # 対象年度（西暦。令和8年度 = 2026）
    insurer: str                   # its / kyokai_tokyo
    grades: list[GradeRow] = field(default_factory=list)
    rates: dict[str, Rate] = field(default_factory=dict)      # kenpo/kodomo/kaigo/konen
    konen_bands: list[tuple[int, int, int, int | None]] = field(default_factory=list)
    #                (厚年等級, 標準報酬月額, 下限, 上限None=∞)

    def find_grade(self, amount: float) -> GradeRow:
        """報酬月額 → 等級行（下限以上・上限未満の半開区間）。"""
        if amount < 0:
            raise ShahoMasterError(f"報酬月額がマイナスです: {amount}")
        for row in self.grades:
            if row.contains(amount):
                return row
        raise ShahoMasterError(f"報酬月額 {amount} がどの等級にも入りません")


def _norm(v) -> str:
    return unicodedata.normalize("NFKC", str(v if v is not None else "")).strip()


def _header_map(ws, row: int, sheet: str, required: tuple, first_col: int = 1,
                last_col: int | None = None) -> dict:
    """見出し行 → {見出し名: 列番号}。必須見出しが無ければ ShahoMasterError。

    照合は NFKC 正規化どうしで行う（全角括弧と半角括弧の揺れを吸収する）。
    呼び出し側が書いたままの名前でも引けるよう、required の原文も別名で入れておく。
    """
    cols = {}
    for c in range(first_col, (last_col or ws.max_column) + 1):
        name = _norm(ws.cell(row, c).value)
        if name and name not in cols:
            cols[name] = c
    missing = [h for h in required if _norm(h) not in cols]
    if missing:
        raise ShahoMasterError(
            f"シート「{sheet}」の{row}行目に必須の見出しがありません: {'、'.join(missing)}"
            "（列名を変えた場合はコード側の対応も必要です）")
    for h in required:
        cols[h] = cols[_norm(h)]
    return cols


def _to_int(v, what: str):
    if v is None or _norm(v) == "":
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        raise ShahoMasterError(f"{what} が数値ではありません: {v!r}")
    if abs(f - round(f)) > 1e-6:
        raise ShahoMasterError(f"{what} が整数ではありません: {v!r}")
    return int(round(f))


def _parse_reiwa_year(text: str) -> int | None:
    """'令和8年度...' → 2026。見つからなければ None。"""
    t = _norm(text)
    if "令和" not in t:
        return None
    digits = ""
    for ch in t.split("令和", 1)[1]:
        if ch.isdigit():
            digits += ch
        else:
            break
    return (int(digits) + REIWA_BASE_YEAR) if digits else None


# ---------------------------------------------------------------------------
# 等級表の読み込み
# ---------------------------------------------------------------------------
def load_grade_table(path: str, insurer: str, expected_year: int) -> ShahoMaster:
    """等級表Excelを読み、総当たり検証してから ShahoMaster を返す。"""
    if insurer not in INSURERS:
        raise ShahoMasterError(
            f"保険者 '{insurer}' は未対応です（使えるのは {' / '.join(INSURERS)}）")
    if not os.path.exists(path):
        raise ShahoMasterError(f"等級表Excelが見つかりません: {path}")
    wb = openpyxl.load_workbook(path, data_only=True)
    for sheet in (SHEET_GRADES, SHEET_RATES):
        if sheet not in wb.sheetnames:
            raise ShahoMasterError(f"シート「{sheet}」がありません: {path}")

    ws = wb[SHEET_GRADES]

    # --- 年度検証（タイトル行の「令和N年度」と対象年度を突き合わせる） ---
    year = None
    for r in range(1, GRADE_HEADER_ROW):
        year = _parse_reiwa_year(ws.cell(r, 1).value)
        if year:
            break
    if year is None:
        raise ShahoMasterError(
            f"シート「{SHEET_GRADES}」の冒頭に年度表記（令和N年度）が見つかりません")
    if year != expected_year:
        raise ShahoMasterError(
            f"等級表の年度（令和{year - REIWA_BASE_YEAR}年度={year}）と対象年度（{expected_year}）が"
            "一致しません。翌年度の等級表に差し替えるか、対象年度を確認してください")

    prefix = INSURERS[insurer]
    required = ("健保等級", "健保標準報酬月額", "報酬月額下限（以上）", "報酬月額上限（未満）",
                "厚年等級", "厚年標準報酬月額", "厚生年金",
                f"{prefix}健康", f"{prefix}支援金", f"{prefix}介護")
    cols = _header_map(ws, GRADE_HEADER_ROW, SHEET_GRADES, required)

    grades: list[GradeRow] = []
    premiums: list[dict] = []          # 保険料額列の検証用（行と同順）
    for r in range(GRADE_HEADER_ROW + 1, ws.max_row + 1):
        g = _to_int(ws.cell(r, cols["健保等級"]).value, f"{r}行目の健保等級")
        if g is None:
            break                      # 空行が来たら終わり
        row = GradeRow(
            kenpo_grade=g,
            kenpo_smr=_to_int(ws.cell(r, cols["健保標準報酬月額"]).value, f"{r}行目の健保標準報酬月額"),
            lower=_to_int(ws.cell(r, cols["報酬月額下限（以上）"]).value, f"{r}行目の下限") or 0,
            upper=_to_int(ws.cell(r, cols["報酬月額上限（未満）"]).value, f"{r}行目の上限"),
            konen_grade=_to_int(ws.cell(r, cols["厚年等級"]).value, f"{r}行目の厚年等級"),
            konen_smr=_to_int(ws.cell(r, cols["厚年標準報酬月額"]).value, f"{r}行目の厚年標準報酬月額"),
        )
        grades.append(row)
        premiums.append({
            "kenpo": ws.cell(r, cols[f"{prefix}健康"]).value,
            "kodomo": ws.cell(r, cols[f"{prefix}支援金"]).value,
            "kaigo": ws.cell(r, cols[f"{prefix}介護"]).value,
            "konen": ws.cell(r, cols["厚生年金"]).value,
        })
    if not grades:
        raise ShahoMasterError(f"シート「{SHEET_GRADES}」にデータ行がありません")

    # --- 等級の連番・区間の連続性 ---
    for i, row in enumerate(grades):
        if row.kenpo_grade != i + 1:
            raise ShahoMasterError(
                f"健保等級が連番ではありません: {i + 1} のはずが {row.kenpo_grade}")
        if row.upper is not None and row.upper <= row.lower:
            raise ShahoMasterError(
                f"健保等級{row.kenpo_grade}: 上限({row.upper}) ≤ 下限({row.lower}) です")
        if i > 0 and grades[i - 1].upper != row.lower:
            raise ShahoMasterError(
                f"健保等級{row.kenpo_grade}: 前の等級の上限({grades[i - 1].upper})と"
                f"下限({row.lower})が繋がっていません")
    if grades[0].lower != 0:
        raise ShahoMasterError(f"最初の等級の下限が0ではありません: {grades[0].lower}")
    if grades[-1].upper is not None:
        raise ShahoMasterError("最終等級の上限は空欄（上限なし）のはずです")
    for row in grades[:-1]:
        if row.upper is None:
            raise ShahoMasterError(
                f"健保等級{row.kenpo_grade}: 最終等級以外で上限が空欄です")

    # --- 料率と厚年等級マスタ ---
    rates = _load_rates(wb, insurer)
    konen_bands = _load_konen_bands(wb)

    # --- 健保の帯が厚年等級マスタと矛盾しないこと ---
    def konen_smr_at(amount: float) -> int:
        for _g, smr, lo, up in konen_bands:
            if amount >= lo and (up is None or amount < up):
                return smr
        raise ShahoMasterError(f"報酬月額 {amount} が厚年等級マスタのどの帯にも入りません")

    for row in grades:
        probes = [row.lower] if row.upper is None else [row.lower, row.upper - 1]
        for p in probes:
            expected = konen_smr_at(p)
            if expected != row.konen_smr:
                raise ShahoMasterError(
                    f"健保等級{row.kenpo_grade}（報酬月額{p}）の厚年標準報酬月額が"
                    f"厚年等級マスタと食い違います: 表={row.konen_smr} マスタ={expected}")

    # --- 保険料額列 = 標準報酬月額 × 本人負担率 の総当たり検証 ---
    for row, prem in zip(grades, premiums):
        checks = (("kenpo", row.kenpo_smr), ("kodomo", row.kenpo_smr),
                  ("kaigo", row.kenpo_smr), ("konen", row.konen_smr))
        for key, smr in checks:
            cell = prem[key]
            if cell is None:
                raise ShahoMasterError(
                    f"健保等級{row.kenpo_grade}: {key} の保険料額が空欄です")
            expected = smr * rates[key].employee
            if abs(float(cell) - expected) > PREMIUM_TOLERANCE:
                raise ShahoMasterError(
                    f"健保等級{row.kenpo_grade}: {key} の保険料額({cell})が "
                    f"標準報酬月額×本人負担率({expected:.2f})と合いません。"
                    "料率か表のどちらかが古い可能性があります")

    return ShahoMaster(path=path, year=year, insurer=insurer,
                       grades=grades, rates=rates, konen_bands=konen_bands)


def _load_rates(wb, insurer: str) -> dict[str, Rate]:
    """「設定・出典」の料率ブロックから、指定保険者＋共通の4本を取る。"""
    ws = wb[SHEET_RATES]
    cols = _header_map(ws, RATE_HEADER_ROW, SHEET_RATES,
                       ("区分", "保険者", "項目", "全体料率", "本人負担率"),
                       first_col=1, last_col=8)
    url_col = cols.get("出典URL")
    biko_col = cols.get("適用・備考")
    # 照合は NFKC 正規化どうし（全角括弧の揺れを吸収）
    item_keys = {_norm(k): v for k, v in RATE_ITEM_KEYS.items()}
    ref_item = _norm("厚生年金（参照）")
    rates: dict[str, Rate] = {}
    for r in range(RATE_HEADER_ROW + 1, ws.max_row + 1):
        ins = _norm(ws.cell(r, cols["保険者"]).value)
        item = _norm(ws.cell(r, cols["項目"]).value)
        if not ins and not item:
            break                      # 料率ブロックの終わり（空行）
        ins_key = RATE_INSURER_NAMES.get(ins)
        item_key = item_keys.get(item)
        if ins_key is None or item_key is None:
            raise ShahoMasterError(
                f"「{SHEET_RATES}」{r}行目: 保険者「{ins}」項目「{item}」を解釈できません")
        if ins_key not in ("共通", insurer):
            continue
        if item_key in rates and item != ref_item:
            raise ShahoMasterError(f"「{SHEET_RATES}」: {item} の料率が二重に定義されています")
        total = ws.cell(r, cols["全体料率"]).value
        emp = ws.cell(r, cols["本人負担率"]).value
        if not isinstance(total, (int, float)) or not isinstance(emp, (int, float)):
            raise ShahoMasterError(f"「{SHEET_RATES}」{r}行目: 料率が数値ではありません")
        if not 0 < emp <= total <= 0.5:
            raise ShahoMasterError(
                f"「{SHEET_RATES}」{r}行目: 料率が不自然です（全体{total}・本人{emp}）")
        if abs(emp - total / 2) > 1e-9:
            raise ShahoMasterError(
                f"「{SHEET_RATES}」{r}行目: 本人負担率({emp})が全体率({total})の半分ではありません")
        if item_key not in rates:
            rates[item_key] = Rate(
                total=float(total), employee=float(emp),
                applies=_norm(ws.cell(r, biko_col).value) if biko_col else "",
                source_url=_norm(ws.cell(r, url_col).value) if url_col else "")
    missing = [k for k in ("kenpo", "kodomo", "kaigo", "konen") if k not in rates]
    if missing:
        raise ShahoMasterError(
            f"「{SHEET_RATES}」に保険者の料率が揃っていません（不足: {'、'.join(missing)}）")
    return rates


def _load_konen_bands(wb) -> list[tuple[int, int, int, int | None]]:
    """「設定・出典」I列以降の厚生年金 等級マスタ（1〜32）。"""
    ws = wb[SHEET_RATES]
    cols = _header_map(ws, RATE_HEADER_ROW, SHEET_RATES,
                       ("厚年等級", "標準報酬月額", "報酬月額下限（以上）", "報酬月額上限（未満）"),
                       first_col=9)
    bands = []
    for r in range(RATE_HEADER_ROW + 1, ws.max_row + 1):
        g = _to_int(ws.cell(r, cols["厚年等級"]).value, f"厚年等級マスタ{r}行目の等級")
        if g is None:
            break
        bands.append((
            g,
            _to_int(ws.cell(r, cols["標準報酬月額"]).value, f"厚年等級{g}の標準報酬月額"),
            _to_int(ws.cell(r, cols["報酬月額下限（以上）"]).value, f"厚年等級{g}の下限") or 0,
            _to_int(ws.cell(r, cols["報酬月額上限（未満）"]).value, f"厚年等級{g}の上限"),
        ))
    if not bands:
        raise ShahoMasterError("厚生年金の等級マスタが読めません")
    for i, (g, _smr, lo, up) in enumerate(bands):
        if g != i + 1:
            raise ShahoMasterError(f"厚年等級が連番ではありません: {i + 1} のはずが {g}")
        if i > 0 and bands[i - 1][3] != lo:
            raise ShahoMasterError(f"厚年等級{g}: 前の等級と区間が繋がっていません")
    if bands[-1][3] is not None:
        raise ShahoMasterError("厚年の最終等級の上限は空欄（上限なし）のはずです")
    return bands


# ---------------------------------------------------------------------------
# 報酬分類マスタ（CSV）
# ---------------------------------------------------------------------------
# 適用開始月・適用終了月（YYYY-MM・両端含む・空=無期限）を持つのは、**同じ項目・同じ
# 体系別名でも月によって意味が変わる**実例があるため。時給制の「みなし給」(allowance2) は
# 2026-05〜07 支給分では差額調整(a24)への再掲（総支給に乗らない情報項目）だが、
# 2026-08 支給分からは実額の支給になった（経理モードの ZANTEI_LABELS_FROM と同じ事象）。
CLASS_MASTER_COLS = ["source_key", "salary_system_label", "label", "class", "fixed",
                     "適用開始月", "適用終了月", "note"]
# 対象   = 通貨の報酬（検算ゲートの基準＝雇用保険対象額に入っている）
# 現物   = 現物給与。報酬には数えるが、現金の総支給額の**外側**なので検算からは除く
#          （実測: allowance53 は other5 に不算入）
# 対象外 = 報酬に数えない（実費弁償・再掲・情報項目）
# 未設定 = 判断待ち。金額が出たらその人は要確認へ落ちる
CLASS_VALUES = ("対象", "現物", "対象外", "未設定")


@dataclass(frozen=True)
class ClassRule:
    source_key: str
    salary_system_label: str       # 正規化済み。空=全体系共通
    label: str
    cls: str                       # 対象 / 対象外 / 未設定
    fixed: str                     # "1"=固定的賃金 / "0"=非固定 / ""=未設定
    ym_from: str = ""              # 適用開始月 YYYY-MM（空=最初から）
    ym_to: str = ""                # 適用終了月 YYYY-MM（空=無期限）
    note: str = ""

    def covers(self, ym: str) -> bool:
        return ((not self.ym_from or self.ym_from <= ym)
                and (not self.ym_to or ym <= self.ym_to))


def load_class_master(path: str) -> dict:
    """報酬分類マスタを読む。

    戻り値 dict:
      lookup  : {(source_key, 正規化体系別名): ClassRule}（体系別名空の行は ("key","")）
      rules   : 全行
      path    : 読んだパス
    分類の解決は (source_key, 体系別名) 完全一致 → (source_key, 空) の順で
    resolve_class() を使う。
    """
    if not os.path.exists(path):
        raise ShahoMasterError(
            f"報酬分類マスタが見つかりません: {path}\n"
            "先に tools/build_shaho_class_master.py で初期版を作ってください")
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            rows = list(csv.DictReader(io.StringIO(raw.decode(enc))))
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ShahoMasterError(f"報酬分類マスタの文字コードを判別できません: {path}")
    if not rows:
        raise ShahoMasterError(f"報酬分類マスタが空です: {path}")
    missing = [c for c in CLASS_MASTER_COLS if c not in rows[0]]
    if missing:
        raise ShahoMasterError(
            f"報酬分類マスタに列が足りません: {'、'.join(missing)}（{path}）")

    def norm_ym(v, i, what):
        t = _norm(v)
        if t and (len(t) != 7 or t[4] != "-" or not (t[:4] + t[5:]).isdigit()):
            raise ShahoMasterError(
                f"報酬分類マスタ{i}行目: {what} '{t}' は YYYY-MM 形式で書いてください")
        return t

    lookup: dict[tuple[str, str], list[ClassRule]] = {}
    rules: list[ClassRule] = []
    for i, r in enumerate(rows, start=2):
        key = _norm(r.get("source_key"))
        if not key or key.startswith("#"):
            continue
        if ":" not in key:
            raise ShahoMasterError(
                f"報酬分類マスタ{i}行目: source_key '{key}' は "
                "'salary_items:allowance1' の形式で書いてください")
        cls = _norm(r.get("class"))
        if cls not in CLASS_VALUES:
            raise ShahoMasterError(
                f"報酬分類マスタ{i}行目: class '{cls}' は "
                f"{' / '.join(CLASS_VALUES)} のどれかにしてください")
        fixed = _norm(r.get("fixed"))
        if fixed not in ("", "0", "1"):
            raise ShahoMasterError(
                f"報酬分類マスタ{i}行目: fixed '{fixed}' は 1 / 0 / 空 のどれかです")
        rule = ClassRule(
            source_key=key,
            salary_system_label=_norm(r.get("salary_system_label")).replace(" ", ""),
            label=_norm(r.get("label")),
            cls=cls, fixed=fixed,
            ym_from=norm_ym(r.get("適用開始月"), i, "適用開始月"),
            ym_to=norm_ym(r.get("適用終了月"), i, "適用終了月"),
            note=_norm(r.get("note")))
        if rule.ym_from and rule.ym_to and rule.ym_from > rule.ym_to:
            raise ShahoMasterError(
                f"報酬分類マスタ{i}行目: 適用開始月({rule.ym_from})が"
                f"終了月({rule.ym_to})より後です")
        pair = (rule.source_key, rule.salary_system_label)
        for other in lookup.get(pair, []):
            lo1, hi1 = rule.ym_from or "0000-00", rule.ym_to or "9999-99"
            lo2, hi2 = other.ym_from or "0000-00", other.ym_to or "9999-99"
            if lo1 <= hi2 and lo2 <= hi1:
                raise ShahoMasterError(
                    f"報酬分類マスタ{i}行目: {pair} の適用期間が別の行と重なっています")
        lookup.setdefault(pair, []).append(rule)
        rules.append(rule)
    return {"lookup": lookup, "rules": rules, "path": path}


def resolve_class(master: dict, source_key: str, system_label: str,
                  ym: str = "9999-99") -> ClassRule | None:
    """(source_key, 体系別名) 完全一致 → (source_key, 空) の順で、支給月 ym に
    効いている行を引く。無ければ None。"""
    norm_label = _norm(system_label).replace(" ", "")
    for pair in ((source_key, norm_label), (source_key, "")):
        for rule in master["lookup"].get(pair, []):
            if rule.covers(ym):
                return rule
    return None
