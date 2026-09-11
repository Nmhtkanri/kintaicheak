# -*- coding: utf-8 -*-
"""賞与用の freee 取引インポート CSV（支給・健康保険・厚生年金）を作る（2026-09-10 新設）。

入力は jinjer API（既定。`GET /v1/employees/bonus-statements`。子ども・子育て支援金だけ API に無いので
標準賞与額×率÷2 で計算し、差引支給額との差で検算する）か、手で落とす「賞与支給控除項目一覧表」CSV
（121 列・cp932。`jinjer_賞与支給控除項目一覧表_9637_YYYYMMDD.csv`）。

規則は経理担当の最終 CSV（Y:\\給与明細\\R8年\\{3,4,6,9}月\\freee\\freee_data_*賞与*.csv）から
2026-09-10 に導いた。2026-09 の FE 部賞与 30 人で全項目一致:
  - 支給ファイル … **全員で 1 取引**（管理番号=先頭（最小）の社員番号・取引先=従業員・発生日=画面で指定・
      支払期日=支給日。給与ファイルと同じ慣例）。支給日が複数あるときは支給日ごとに取引を分けて要確認に出す。
      1 人ごとに 賞与手当（役員は役員賞与）＝総支給額、続けて預り金のマイナス行を
      厚年 → 健保 → 雇用 → 源泉所得税 → 介護 → 子ども・子育て支援金 の順（0 の行は出さない）。
  - 健康保険ファイル … KEMPO 1 取引（発生日=支給月の月末・支払期日=翌月末）。1 人 1 行の法定福利費＝
      **標準賞与額 × 総率 ÷ 2 を四捨五入**（総率＝健保＋支援金、介護保険料を払う人は＋介護）。
      jinjer の「事業主〜」列は項目ごとに端数処理しているため 7/30 人で 1 円ずれ、最終 CSV と合わない。
      末尾に 預り金（部門=本社・従業員は空欄）: 健保／介護／支援金 の各合計。
  - 厚生年金ファイル … KONEN 1 取引。法定福利費＝標準賞与額 × 厚年率 ÷ 2 を四捨五入。
      末尾に 預り金（本社）厚年合計と、法定福利費 人件費（本社）＝事業主子ども・子育て拠出金の合計
      （備考「子ども・子育て拠出金」）。
  - 本人控除の検算 … 本人＝標準賞与額 × 率 ÷ 2 を五捨六入（jinjer の丸め）。CSV の値と合わなければ
      料率が変わった可能性として要確認に出す。

料率は `Config.KEIRI_BONUS_RATES_CSV`（既定 Z:\\API連携\\docs\\経理モード_賞与料率.csv）にあれば
そこから（適用開始年月 ≤ 支給月 の最新行）、無ければ下の DEFAULT_RATES を使う。
"""
from __future__ import annotations

import csv
import io
import json
import math
import os
import re
from collections import defaultdict

from config import Config
from services.keiri_engine import (NAME_ALIASES, Resolver, detail_row,
                                   jinkenhi_account, jp_date, load_custom_histories,
                                   load_or_fetch_roster, month_last_day, roster_index,
                                   write_freee_csv, ym_add)
from services.keiri_api import (classify_employee, get_client, normalize_label, normalize_ymd, now_iso,
                                statement_flag, to_number)
from services.keiri_engine import SPECIAL_TARGET_EMPLOYEES

# 2026 年度（関東 IT ソフトウェア健康保険組合／厚生年金）。総率＝労使合計。
DEFAULT_RATES = {"kenpo": 0.0927, "kaigo": 0.0180, "shien": 0.0023, "konen": 0.1830}
RATES_CSV = Config.KEIRI_BONUS_RATES_CSV
RATE_COLS = {"kenpo": "健保率", "kaigo": "介護率", "shien": "支援金率", "konen": "厚年率"}

HONSHA_BUMON = "本社"
KEMPO_TORIHIKISAKI = "関東ＩＴソフトウェア健康保険組合"
KONEN_TORIHIKISAKI = "厚生労働省"
KYOSHUTSUKIN_BIKO = "子ども・子育て拠出金"

# jinjer 賞与 CSV の列名（実測 2026-09-04 のエクスポート）
COL_EMP, COL_NAME, COL_TOTAL, COL_PAID_ON = "社員番号", "氏名", "総支給額", "支給日"
COL_SB_KENPO, COL_SB_KONEN = "健保標準賞与額", "厚年標準賞与額"
COL_KYOSHUTSU = "事業主子ども・子育て拠出金"
# 預り金のマイナス行（出す順）: (CSV 列, freee 品目)
AZUKARI_ORDER = [
    ("厚生年金", "厚生年金保険料（預り分）"),
    ("健康保険料", "健康保険料（預り分）"),
    ("雇用保険料", "雇用保険料（預り分）"),
    ("所得税", "源泉所得税"),
    ("介護保険料", "介護保険料（預り分）"),
    ("子ども・子育て支援金", "子ども・子育て支援金（預り分）"),
]
REQUIRED_COLS = [COL_EMP, COL_NAME, COL_TOTAL, COL_PAID_ON, COL_SB_KENPO, COL_SB_KONEN,
                 COL_KYOSHUTSU, "年調過不足額", "控除合計", "差引支給額"] + [c for c, _ in AZUKARI_ORDER]
UNMAPPED_DEDUCTION_COLS = [f"賞与控除項目{i}" for i in range(1, 16)]
EMPLOYER_COLS = ["事業主健康保険料", "事業主介護保険料", "事業主子ども・子育て支援金", "事業主厚生年金"]
# 画面の「要確認の件数」と md に出す alerts のキー
MD_ALERT_KEYS = ("bonus_rate", "bonus_company_diff", "bonus_unmapped", "bonus_check", "bumon_missing",
                 "bumon_unknown", "ikusei_maybe")


def bonus_account(emp: str) -> str:
    """賞与の勘定科目。役員は 役員賞与、それ以外は 賞与手当（給与の 給料手当 とは別）。"""
    return "役員賞与" if jinkenhi_account(emp, is_bonus=True) == "役員賞与" else "賞与手当"


# ---------------------------------------------------------------------------
# 丸め
# ---------------------------------------------------------------------------
def round_half_up(x: float) -> int:
    """四捨五入（0.5 は切り上げ）。経理担当の会社負担の丸め。"""
    return int(math.floor(x + 0.5 + 1e-9))


def round_gosha_rokunyu(x: float) -> int:
    """五捨六入（0.5 は切り捨て、0.6 以上は切り上げ）。jinjer の本人控除の丸め。"""
    fl = math.floor(x)
    return fl + 1 if (x - fl) > 0.5 + 1e-9 else fl


# ---------------------------------------------------------------------------
# 料率
# ---------------------------------------------------------------------------
def load_rates(month: str, path: str | None = None) -> tuple[dict, str]:
    """支給月に効く料率と、その出どころ（説明文）を返す。"""
    path = path or RATES_CSV
    if not path or not os.path.exists(path):
        return dict(DEFAULT_RATES), "既定値（コード内 DEFAULT_RATES）"
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            rows = list(csv.DictReader(io.StringIO(raw.decode(enc))))
            break
        except UnicodeDecodeError:
            continue
    else:
        return dict(DEFAULT_RATES), f"料率 CSV を読めず既定値: {path}"
    chosen = None
    bad_start = []
    for r in rows:
        start = str(r.get("適用開始年月") or "").strip()
        if not re.fullmatch(r"\d{4}-\d{2}", start):
            bad_start.append(start)
            continue                      # Excel 保存で '2026/4' などに化けた行は無視
        if start <= month and (chosen is None or start >= chosen[0]):
            chosen = (start, r)
    if chosen is None:
        note = f"（適用開始年月が YYYY-MM でない行を無視: {'・'.join(bad_start)}）" if bad_start else ""
        return dict(DEFAULT_RATES), f"料率 CSV に {month} 以前の行が無く既定値: {path}{note}"
    rates = {}
    for key, col in RATE_COLS.items():
        v = to_number(chosen[1].get(col))
        if v is None:
            return dict(DEFAULT_RATES), f"料率 CSV の {col} が数値でなく既定値: {path}"
        v = float(v)
        if not (0.01 <= v <= 40):
            # % 表記の実務レンジ外（0.0927 のような小数表記や 100 超）は読み違いなので既定値に戻す
            return dict(DEFAULT_RATES), f"料率 CSV の {col}={v} が % 表記の範囲（0.01〜40）外で既定値: {path}"
        rates[key] = round(v / 100.0, 6)  # CSV は必ず % 表記（9.27 → 0.0927。0.23 は 0.23% であって 23% ではない）
    return rates, f"{path}（適用開始 {chosen[0]}）"


# ---------------------------------------------------------------------------
# 入力（API）
# ---------------------------------------------------------------------------
API_DEDUCTION_LABELS = ["雇用保険料", "健康保険料", "介護保険料", "厚生年金", "厚生年金基金", "年調過不足額", "所得税"]
API_OTHER_MAP = {          # bonus_other_items の項目名 → CSV の列名
    "健康保険料": "事業主健康保険料", "介護保険料": "事業主介護保険料", "厚生年金": "事業主厚生年金",
    "子ども・子育て拠出金": "事業主子ども・子育て拠出金",
    "健保標準賞与額": COL_SB_KENPO, "厚年標準賞与額": COL_SB_KONEN,
}


def _n_calculated(data) -> int:
    return sum(1 for p in data or [] for st in p.get("statements") or []
               if (st.get("payroll_info") or {}).get("is_calculated", True))


def fetch_bonus_statements(cache_dir: str, ym: str, client=None, refresh: bool = False) -> tuple[list, str]:
    """指定月の賞与計算結果を API から取って raw/bonus_statements_{ym}.json にキャッシュし、(生 data[], 取得時刻) を返す。

    空の応答や「計算済みのレコードが 1 件も無い」応答（jinjer で賞与計算を実行する前）はキャッシュせずに止める。
    キャッシュすると次回以降も既定でそれを読み、計算後に取り直さない限り古いまま進んでしまうため。
    """
    path = os.path.join(cache_dir, "raw", f"bonus_statements_{ym}.json")
    if not refresh and os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            obj = json.load(f)
        return obj["data"], str(obj.get("fetched_at") or "")
    data = (client or get_client()).get_bonus_statements(ym)
    if not data:
        raise ValueError(f"{ym} の賞与計算結果が jinjer から 0 件でした（賞与計算をしていない月か、API が空を返した）。"
                         "計算対象月を確認して、もう一度実行してください")
    if not _n_calculated(data):
        raise ValueError(f"{ym} の賞与計算結果は {len(data)} 人分ありますが、賞与計算が実行済みの人が 1 人もいません。"
                         "jinjer で賞与計算を実行してから、もう一度「賞与計算結果を API から取り直す」で実行してください（保存していません）")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    fetched_at = now_iso()
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"fetched_at": fetched_at, "executed_on": ym, "data": data}, f, ensure_ascii=False)
    return data, fetched_at


def _items_by_label(items, dup=None) -> dict:
    """項目名 → 値。項目名は normalize_label で寄せる。同じ名前が 2 件あれば dup に名前を積む（値は先勝ち）。"""
    out = {}
    for it in items or []:
        lab = normalize_label(str(it.get("label") or ""))
        if not lab:
            continue
        if lab in out:
            if dup is not None:
                dup.add(lab)
            continue
        out[lab] = it.get("value")
    return out


def rows_from_api(data: list, *, rates: dict, alerts, count: int | None = None,
                  paid_on: str | None = None) -> list[dict]:
    """API の賞与計算結果を、CSV（賞与支給控除項目一覧表）と同じ列名の行 dict にする。

    - 子ども・子育て支援金は API に無いので、本人＝標準賞与額×率÷2 の五捨六入、事業主＝本人と同額（谷津さん確認）。
      検算: 総支給額 − 差引支給額 − API に載っている控除の合計 ＝ 支援金 のはず（差引支給額は支援金控除後）。合わなければ要確認。
    - 同じ月に賞与が複数回あるとき（count が複数）は count を指定する。未指定で複数あれば ValueError。
    - paid_on を渡すとその日付を支給日にする（API の paid_on は処理基準日と同じ値のことがある）。
    """
    counts = set()
    for person in data:
        for st in person.get("statements") or []:
            if st.get("count") is not None:
                counts.add(int(to_number(st.get("count")) or 0))
    if count is None and len(counts) > 1:
        raise ValueError("この月は賞与が複数回あります（回数: " + "・".join(str(c) for c in sorted(counts))
                         + "）。回数を指定してください")
    rows = []
    skipped = defaultdict(list)
    known_ded = {normalize_label(x) for x in API_DEDUCTION_LABELS} | {normalize_label(f"賞与控除項目{i}") for i in range(1, 16)}
    for person in data:
        emp = str(person.get("employee_id") or "").strip()
        if not emp or (classify_employee(emp) != "target" and emp not in SPECIAL_TARGET_EMPLOYEES):
            skipped["対象外の社員番号"].append(emp or "(空)")
            continue
        for st in person.get("statements") or []:
            if count is not None and int(to_number(st.get("count")) or 0) != int(count):
                skipped["別の回数のレコード"].append(emp)
                continue
            pi = st.get("payroll_info") or {}
            if not pi.get("is_calculated", True):
                skipped["賞与計算が未実行"].append(emp)        # jinjer で賞与計算を実行する前のレコード
                continue
            bi = st.get("basic_info") or {}
            name = f"{bi.get('last_name', '')} {bi.get('first_name', '')}".strip()
            dup = set()
            ded = _items_by_label(pi.get("bonus_deduction_items"), dup)
            oth = _items_by_label(pi.get("bonus_other_items"), dup)
            pay = _items_by_label(pi.get("bonus_payment_items"), dup)
            if dup:
                alerts["bonus_check"].add((emp, name, "API の項目名が重複しています（先の値を採用）: " + "・".join(sorted(dup))))
            # 既知の名前に無い控除に金額が入っていたら、仕訳に載らないので要確認へ
            for lab, val in ded.items():
                if lab not in known_ded and to_number(val):
                    alerts["bonus_unmapped"].add((emp, name, f"控除「{lab}」（API）", int(round(to_number(val)))))
            total = sum(int(round(to_number(it.get("value")) or 0)) for it in pi.get("bonus_items") or [])
            # paid_on は給与側の実測では payroll_info 直下、賞与の実測（2026-09-11）では statement 直下。両方を見る
            api_paid_on = normalize_ymd(str(statement_flag(person, st, "paid_on") or ""))
            if not paid_on and api_paid_on:
                alerts["bonus_check"].add((emp, name, f"支払期日に API の支給日 {api_paid_on} を使いました"
                                                     "（処理基準日と同じ値のことがあるので、経理の支給日と違えば画面の「支払期日（支給日）」に入れて作り直す）"))
            row = {COL_EMP: emp, COL_NAME: name, COL_TOTAL: str(total),
                   COL_PAID_ON: (paid_on or api_paid_on).replace("-", "/"), "_source": "api"}
            g = lambda d, lab: str(int(round(to_number(d.get(normalize_label(lab))) or 0)))   # noqa: E731
            for lab in API_DEDUCTION_LABELS:
                row[lab] = g(ded, lab)
            for i in range(1, 16):
                row[f"賞与控除項目{i}"] = g(ded, f"賞与控除項目{i}")
            for lab, col in API_OTHER_MAP.items():
                row[col] = g(oth, lab)
            row["事業主雇用保険料"] = g(oth, "雇用保険料")
            # 支援金（API に無い）
            sb_k = int(row[COL_SB_KENPO])
            shien = round_gosha_rokunyu(sb_k * rates["shien"] / 2) if sb_k else 0
            row["子ども・子育て支援金"] = str(shien)
            row["事業主子ども・子育て支援金"] = str(shien)
            listed = sum(int(row[lab]) for lab in API_DEDUCTION_LABELS) + sum(int(row[f"賞与控除項目{i}"]) for i in range(1, 16))
            sashihiki = int(round(to_number(pay.get(normalize_label("差引支給額"))) or 0))
            row["差引支給額"] = str(sashihiki)
            row["控除合計"] = str(listed + shien)
            implied = total - sashihiki - listed
            if implied != shien:
                alerts["bonus_check"].add((emp, name, f"支援金の計算値 {shien:,} が明細から逆算した値 {implied:,}"
                                                     "（総支給額−差引支給額−API の控除合計）と合いません"))
            rows.append(row)
    if not rows:
        detail = "、".join(f"{k} {len(v)} 人" for k, v in skipped.items()) or "レコードなし"
        raise ValueError("API の賞与計算結果に使える人がいません（" + detail + "）。"
                         "jinjer で賞与計算を実行済みか、計算対象月と回数が合っているかを確認してください。"
                         "キャッシュを読んでいる場合は「賞与計算結果を API から取り直す」を付けて再実行してください")
    for reason, emps in skipped.items():
        # 読み飛ばした人は常に要確認へ（rows が空のときだけでなく）。7777777 のような例外は SPECIAL_TARGET_EMPLOYEES で対象に入れる
        alerts["bonus_check"].add(("", "", f"{reason}の人を {len(emps)} 人読み飛ばしました: " + "・".join(emps[:20])
                                   + ("…" if len(emps) > 20 else "")))
    return rows


# ---------------------------------------------------------------------------
# 入力（CSV）
# ---------------------------------------------------------------------------
def load_bonus_csv(source) -> list[dict]:
    """jinjer 賞与 CSV（パスまたは bytes）→ 行 dict のリスト。必須列が無ければ ValueError。"""
    raw = source if isinstance(source, (bytes, bytearray)) else open(source, "rb").read()
    last_missing = None
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        rows = list(csv.DictReader(io.StringIO(text)))
        if not rows:
            raise ValueError("賞与 CSV に明細行がありません")
        missing = [c for c in REQUIRED_COLS if c not in rows[0]]
        if not missing:
            return rows
        last_missing = missing
    if last_missing is None:
        raise ValueError("賞与 CSV の文字コードを判定できません（cp932 / UTF-8 のどちらでもない）")
    raise ValueError("賞与 CSV に必要な列がありません: " + "・".join(last_missing)
                     + "（jinjer の「賞与支給控除項目一覧表」を項目を絞らずに出してください。"
                     "列名が化けている場合は文字コード（cp932）も確認）")


def _num(row: dict, col: str, bad=None) -> int:
    """列の金額。空欄は 0。空欄でないのに数値化できない値は 0 にしつつ bad に列名を積む。"""
    raw = str(row.get(col) or "").strip()
    v = to_number(raw) if raw else 0
    if v is None:
        if bad is not None:
            bad.add(col)
        return 0
    return int(round(v))


def _iso_date(s: str) -> str:
    """'2026/9/25' or '2026-09-25' → '2026-09-25'。"""
    s = normalize_ymd(str(s or "").strip())
    if not s:
        return ""
    parts = s.replace("-", "/").split("/")
    if len(parts) != 3:
        raise ValueError(f"日付の形式が不正です: {s}")
    y, m, d = (int(p) for p in parts)
    return f"{y:04d}-{m:02d}-{d:02d}"


def _norm_name(s: str) -> str:
    return str(s or "").replace("\u3000", " ").strip()


# ---------------------------------------------------------------------------
# 本体
# ---------------------------------------------------------------------------
def build_bonus(rows: list[dict], *, month: str, hassei: str, resolver, ridx: dict,
                rates: dict, alerts, shaho_hassei: str | None = None,
                shaho_kigen: str | None = None) -> dict:
    """3 ファイル分の取引を作る。

    Args:
        rows: load_bonus_csv の結果
        month: 支給月 'yyyy-MM'（社保ファイルの発生日・期日の既定と、部門の文脈に使う）
        hassei: 支給ファイルの発生日 'yyyy-MM-dd'（経理担当が決める。9 月は 9/15、決算賞与は 3/31）
        resolver: keiri_engine.Resolver（部門・人件費区分）
        ridx: roster_index の結果（氏名の正は jinjer の名簿）
        rates: load_rates の結果
        alerts: defaultdict(set)
        shaho_hassei / shaho_kigen: 社保ファイルの発生日・支払期日（既定 月末／翌月末）
    Returns:
        {"支給": [取引...], "健康保険": [取引], "厚生年金": [取引]}
    """
    shaho_hassei = shaho_hassei or month_last_day(month)
    shaho_kigen = shaho_kigen or month_last_day(ym_add(month, 1))
    context = f"{month}賞与"

    pay_rows_by_date = defaultdict(list)     # 支給日 → [(社員番号, 行...)]
    kenpo_rows, konen_rows = [], []
    sums = defaultdict(int)
    seen = defaultdict(int)
    for row in rows:
        seen[str(row.get(COL_EMP) or "").strip()] += 1
    if seen.get("", 0):
        alerts["bonus_check"].add(("", "", f"社員番号が空の行を {seen['']} 行読み飛ばしました（末尾の合計行など）"))
    n_people = 0
    for emp, n in sorted(seen.items()):
        if emp and n > 1:
            alerts["bonus_check"].add((emp, "", f"社員番号が {n} 行あります（二重計上の疑い。行はすべて仕訳に載せています）"))
    if rows and any(c not in rows[0] for c in EMPLOYER_COLS):
        alerts["bonus_check"].add(("", "", "jinjer の事業主〜列が無く、会社負担の 1 円ずれ検知ができません（項目を絞らずに出力してください）"))
    for row in sorted(rows, key=lambda r: str(r.get(COL_EMP) or "")):
        emp = str(row.get(COL_EMP) or "").strip()
        if not emp:
            continue
        name = ridx.get(emp, {}).get("name") or _norm_name(row.get(COL_NAME))
        paid_on = _iso_date(row.get(COL_PAID_ON))
        if not paid_on:
            raise ValueError(f"{emp} {name}: 支給日が空です")
        if paid_on[:7] != month:
            alerts["bonus_check"].add((emp, name, f"支給日 {paid_on} が支給月 {month} と違います（隣の月の CSV や支給月の打ち間違いでないか確認）"))
        n_people += 1
        bad_cols = set()
        g = lambda col: _num(row, col, bad_cols)   # noqa: E731  この行の金額（数値化できない列は bad_cols へ）
        total = g(COL_TOTAL)
        item = resolver.jinkenhi_item(emp, paid_on)
        bumon = resolver.bumon(emp, paid_on, context)

        # --- 検算・要確認 ---
        ded_total = g("控除合計")
        if row.get("_source") != "api" and total - ded_total != g("差引支給額"):
            alerts["bonus_check"].add((emp, name, f"総支給額 {total:,} − 控除合計 {ded_total:,} ≠ 差引支給額 {g('差引支給額'):,}"))
        if g("年調過不足額"):
            alerts["bonus_unmapped"].add((emp, name, "年調過不足額", g("年調過不足額")))
        for col in UNMAPPED_DEDUCTION_COLS:
            if col in row and g(col):
                alerts["bonus_unmapped"].add((emp, name, col, g(col)))
        mapped = sum(g(c) for c, _ in AZUKARI_ORDER)
        if mapped != ded_total:
            alerts["bonus_check"].add((emp, name, f"預り金に載せた控除の合計 {mapped:,} ≠ 控除合計 {ded_total:,}"))

        # --- 支給ファイル ---
        prow = []
        if total:
            prow.append(detail_row(bonus_account(emp), "対象外", total, item, bumon, name))
        else:
            alerts["bonus_check"].add((emp, name, "総支給額が 0 なので賞与手当の行を出していません"))
        for col, freee_item in AZUKARI_ORDER:
            v = g(col)
            if v:
                prow.append(detail_row("預り金", "対象外", -v, freee_item, bumon, name))
        if prow:
            pay_rows_by_date[paid_on].append((emp, prow))

        # --- 社保ファイル: 会社負担 ---
        sb_k, sb_n = g(COL_SB_KENPO), g(COL_SB_KONEN)
        h_kenpo, h_kaigo, h_shien, h_konen = (g("健康保険料"), g("介護保険料"),
                                              g("子ども・子育て支援金"), g("厚生年金"))
        if sb_k:
            rate_k = rates["kenpo"] + rates["shien"] + (rates["kaigo"] if h_kaigo else 0.0)
            company_k = round_half_up(sb_k * rate_k / 2)
            kenpo_rows.append(detail_row("法定福利費", "対象外", company_k, item, bumon, name))
            jinjer_k = g("事業主健康保険料") + g("事業主介護保険料") + g("事業主子ども・子育て支援金")
            if jinjer_k and jinjer_k != company_k:
                # 項目ごとに丸めた jinjer の事業主列と 1 円ずれることがある（2026-09 は 7/30 人）。
                # 経理担当は 2026-09 は率÷2 の四捨五入、2026-06 は jinjer 列に近い値で計上しており揺れがある
                alerts["bonus_company_diff"].add((emp, name, "健保（介護・支援金込み）", company_k, jinjer_k))
            for label, honnin, rate in (("健康保険料", h_kenpo, rates["kenpo"]),
                                        ("介護保険料", h_kaigo, rates["kaigo"] if h_kaigo else 0.0),
                                        ("子ども・子育て支援金", h_shien, rates["shien"])):
                expected = round_gosha_rokunyu(sb_k * rate / 2) if rate else 0
                if expected != honnin:
                    alerts["bonus_rate"].add((emp, name, label, honnin, expected))
        elif h_kenpo or h_shien:
            alerts["bonus_check"].add((emp, name, "健保標準賞与額が 0 なのに本人控除がある"))
        if sb_n:
            company_n = round_half_up(sb_n * rates["konen"] / 2)
            konen_rows.append(detail_row("法定福利費", "対象外", company_n, item, bumon, name))
            jinjer_n = g("事業主厚生年金")
            if jinjer_n and jinjer_n != company_n:
                alerts["bonus_company_diff"].add((emp, name, "厚生年金", company_n, jinjer_n))
            expected = round_gosha_rokunyu(sb_n * rates["konen"] / 2)
            if expected != h_konen:
                alerts["bonus_rate"].add((emp, name, "厚生年金", h_konen, expected))
        elif h_konen:
            alerts["bonus_check"].add((emp, name, "厚年標準賞与額が 0 なのに本人控除がある"))
        sums["kenpo"] += h_kenpo
        sums["kaigo"] += h_kaigo
        sums["shien"] += h_shien
        sums["konen"] += h_konen
        sums["kyoshutsu"] += g(COL_KYOSHUTSU)
        if bad_cols:
            alerts["bonus_check"].add((emp, name, "数値として読めず 0 にした列: " + "・".join(sorted(bad_cols))))

    # --- 支給ファイル: 全員で 1 取引（管理番号=先頭の社員番号。支給日が複数なら支給日ごと）---
    pay_tx = []
    for paid_on in sorted(pay_rows_by_date):
        items = pay_rows_by_date[paid_on]
        pay_tx.append({"管理番号": items[0][0], "発生日": jp_date(hassei), "支払期日": jp_date(paid_on),
                       "取引先": "従業員", "rows": [r for _e, prow in items for r in prow]})
    if len(pay_tx) > 1:
        alerts["bonus_check"].add(("", "", "支給日が " + "・".join(jp_date(d) for d in sorted(pay_rows_by_date))
                                   + " の複数あり、支給ファイルの取引を支給日ごとに分けました"))

    # --- 社保ファイルの末尾（本社の合算行）---
    if sums["kenpo"]:
        kenpo_rows.append(detail_row("預り金", "対象外", sums["kenpo"], "健康保険料（預り分）", HONSHA_BUMON, ""))
    if sums["kaigo"]:
        kenpo_rows.append(detail_row("預り金", "対象外", sums["kaigo"], "介護保険料（預り分）", HONSHA_BUMON, ""))
    if sums["shien"]:
        kenpo_rows.append(detail_row("預り金", "対象外", sums["shien"], "子ども・子育て支援金（預り分）", HONSHA_BUMON, ""))
    if sums["konen"]:
        konen_rows.append(detail_row("預り金", "対象外", sums["konen"], "厚生年金保険料（預り分）", HONSHA_BUMON, ""))
    if sums["kyoshutsu"]:
        konen_rows.append(detail_row("法定福利費", "対象外", sums["kyoshutsu"], "人件費（本社）", HONSHA_BUMON,
                                     "", KYOSHUTSUKIN_BIKO))

    kenpo_tx = [{"管理番号": "KEMPO", "発生日": jp_date(shaho_hassei), "支払期日": jp_date(shaho_kigen),
                 "取引先": KEMPO_TORIHIKISAKI, "rows": kenpo_rows}] if kenpo_rows else []
    konen_tx = [{"管理番号": "KONEN", "発生日": jp_date(shaho_hassei), "支払期日": jp_date(shaho_kigen),
                 "取引先": KONEN_TORIHIKISAKI, "rows": konen_rows}] if konen_rows else []
    return {"支給": pay_tx, "健康保険": kenpo_tx, "厚生年金": konen_tx}


def file_names(label: str, month: str) -> dict:
    mc = month.replace("-", "")
    return {"支給": f"freee_data_{label}（{mc}）.csv",
            "健康保険": f"freee_data_{label}_健康保険（{mc}）.csv",
            "厚生年金": f"freee_data_{label}_厚生年金（{mc}）.csv"}


def build_yokakunin(month: str, label: str, rates: dict, rates_src: str, alerts, n_people: int,
                    input_src: str = "") -> list[str]:
    lines = [f"# 要確認リスト 賞与 {label} {month.replace('-', '')}", "",
             f"- 入力: {input_src}" if input_src else "- 入力: （不明）",
             f"- 対象 {n_people} 人。料率: 健保 {rates['kenpo']:.4%} / 介護 {rates['kaigo']:.4%} / "
             f"支援金 {rates['shien']:.4%} / 厚年 {rates['konen']:.4%}（{rates_src}）",
             "- 会社負担＝標準賞与額×総率÷2 の四捨五入（経理担当の最終 CSV に合わせた。jinjer の事業主列とは 1 円ずれることがある）", ""]
    lines += ["## 料率が違う可能性（本人控除が 標準賞与額×率÷2 の五捨六入 と合わない）", ""]
    if alerts["bonus_rate"]:
        lines += ["| 社員番号 | 氏名 | 項目 | CSV の本人控除 | 料率からの計算 |", "|---|---|---|---|---|"]
        for emp, name, label_, got, exp in sorted(alerts["bonus_rate"]):
            lines.append(f"| {emp} | {name} | {label_} | {got:,} | {exp:,} |")
        lines += ["", "料率 CSV（経理モード_賞与料率.csv）を直すか、jinjer 側の標準賞与額を確認すること。"]
    else:
        lines.append("- なし")
    lines += ["", "## 会社負担が jinjer の事業主列と 1 円ずれた人（計算値を採用）", ""]
    if input_src.startswith("API"):
        lines += ["※ API 入力では事業主子ども・子育て支援金は API に無く、本人と同額の計算値を jinjer 列の合計に含めている。", ""]
    lines += [
              "会社負担＝標準賞与額×総率÷2 の四捨五入（2026-09 の経理担当の実績）。jinjer の事業主列は項目ごとに"
              "丸めるので 1 円ずれることがある。経理担当の方法が月で揺れているので、どちらを正にするか確認して"
              "手で直す場合はこの表を使う。", ""]
    if alerts["bonus_company_diff"]:
        lines += ["| 社員番号 | 氏名 | 項目 | 計算値（採用） | jinjer 事業主列 |", "|---|---|---|---|---|"]
        for emp, name, label_, calc, jin in sorted(alerts["bonus_company_diff"]):
            lines.append(f"| {emp} | {name} | {label_} | {calc:,} | {jin:,} |")
    else:
        lines.append("- なし")
    lines += ["", "## 仕訳に載せていない控除（年調過不足額・賞与控除項目1〜15）", ""]
    if alerts["bonus_unmapped"]:
        lines += ["| 社員番号 | 氏名 | 項目 | 金額 |", "|---|---|---|---|"]
        for emp, name, col, amt in sorted(alerts["bonus_unmapped"]):
            lines.append(f"| {emp} | {name} | {col} | {amt:,} |")
        lines += ["", "**支給ファイルの預り金に入っていない。** 科目を決めて手で行を足すこと。"]
    else:
        lines.append("- なし")
    lines += ["", "## 検算で合わなかった行", ""]
    if alerts["bonus_check"]:
        for emp, name, msg in sorted(alerts["bonus_check"]):
            lines.append(f"- {emp} {name}: {msg}")
    else:
        lines.append("- なし")
    lines += ["", "## 育成期間だったかもしれない人（人件費区分が既定に落ちた新入社員）", "",
              "jinjer の人件費区分は 本社／育成 の例外だけを持ち、空欄は既定（本社以外）になる。"
              "入社まもない人が既定に落ちていたら、本当は 人件費（育成）・部門 本社 ではないか確認する。", ""]
    if alerts["ikusei_maybe"]:
        for emp, joined, ym in sorted(alerts["ikusei_maybe"]):
            lines.append(f"- {emp}（入社 {joined}、対象月 {ym}）")
    else:
        lines.append("- なし")
    lines += ["", "## 部門が対象日時点で解決できず、最古の履歴で代用した従業員", ""]
    if alerts["bumon_missing"]:
        for emp, ctx, v in sorted(alerts["bumon_missing"]):
            lines.append(f"- {emp}（{ctx}）→ {v}")
    else:
        lines.append("- なし")
    lines += ["", "## 部門マスタに無い部門値", ""]
    if alerts["bumon_unknown"]:
        for emp, v in sorted(alerts["bumon_unknown"]):
            lines.append(f"- {emp}: {v}")
    else:
        lines.append("- なし")
    return lines


def generate_bonus(month: str, csv_source, label: str, hassei: str, *, out_base: str | None = None,
                   client=None, refresh: bool = False, shaho_hassei: str | None = None,
                   shaho_kigen: str | None = None, rates_csv: str | None = None,
                   source: str = "csv", count: int | None = None, paid_on: str | None = None,
                   refresh_statements: bool = False) -> dict:
    """賞与 CSV または API → 3 つの freee CSV ＋ 要確認 md。戻り値は画面用の要約。

    source="api" のときは csv_source を使わず、month（計算対象月）の賞与計算結果を API から取る
    （キャッシュ raw/bonus_statements_{month}.json。refresh_statements で取り直し）。
    """
    from services.keiri_engine import OUT_BASE
    out_base = out_base or OUT_BASE
    rates, rates_src = load_rates(month, rates_csv)
    alerts = defaultdict(set)
    if source == "api":
        data, fetched_at = fetch_bonus_statements(out_base, month, client=client, refresh=refresh_statements)
        rows = rows_from_api(data, rates=rates, alerts=alerts, count=count, paid_on=paid_on)
        input_src = (f"API bonus-statements {month}" + (f"（回数 {count}）" if count is not None else "")
                     + f"（取得 {fetched_at}。支援金は標準賞与額×率÷2 の計算値）")
    else:
        rows = load_bonus_csv(csv_source)
        input_src = "CSV（賞与支給控除項目一覧表）"
    # 名簿は給与側と同様に「給与明細を取り直す」側の扱いなので、ここでは取り直さない（キャッシュが無ければ取得）
    roster = load_or_fetch_roster(client, os.path.join(out_base, "raw", "roster.json"), refresh=False)
    ridx = roster_index(roster)
    for emp, alias in NAME_ALIASES.items():
        if emp in ridx:
            ridx[emp]["name"] = alias
    if not rows or not any(_num(r, COL_TOTAL) for r in rows):
        raise ValueError("API の賞与計算結果の総支給額がすべて 0 です（計算対象月・回数を確認）" if source == "api"
                         else "賞与 CSV の総支給額がすべて 0 か、明細行がありません（列や文字コードを確認してください）")
    # 部門・人件費区分のキャッシュ（raw/custom_items.json）は給与モードと共用。取り直すときは給与側と同じ
    # 全対象者で作り直す（賞与対象者だけで上書きすると月次 4CSV の部門が解決できなくなる）
    from services.keiri_api import classify_employee
    emps = {str(r.get(COL_EMP) or "").strip() for r in rows}
    target_ids = {e for e in ridx if classify_employee(e) == "target"} | emps
    histories = load_custom_histories(out_base, target_ids, refresh=refresh)
    resolver = Resolver(histories, alerts, ridx)
    files_tx = build_bonus(rows, month=month, hassei=hassei, resolver=resolver, ridx=ridx,
                           rates=rates, alerts=alerts, shaho_hassei=shaho_hassei, shaho_kigen=shaho_kigen)
    pay_tx_people = {i: [r for r in t["rows"] if r["勘定科目"] in ("賞与手当", "役員賞与")]
                     for i, t in enumerate(files_tx["支給"])}
    mc = month.replace("-", "")
    out_dir = os.path.join(out_base, mc)
    os.makedirs(out_dir, exist_ok=True)
    names = file_names(label, month)
    result_files = {}
    overwritten = [names[k] for k in files_tx if os.path.exists(os.path.join(out_dir, names[k]))]
    for kind, tx in files_tx.items():
        path = os.path.join(out_dir, names[kind])
        write_freee_csv(path, tx)
        n_rows = sum(len(t["rows"]) for t in tx)
        result_files[kind] = {"path": path, "name": names[kind], "transactions": len(tx), "rows": n_rows,
                              "total": sum(r["金額"] for t in tx for r in t["rows"])}
    yk_path = os.path.join(out_dir, f"要確認_賞与_{label}_{mc}.md")
    yk_lines = build_yokakunin(month, label, rates, rates_src, alerts,
                               sum(len(v) for v in pay_tx_people.values()), input_src=input_src)
    with open(yk_path, "w", encoding="utf-8") as f:
        f.write("\n".join(yk_lines) + "\n")
    return {"month": month, "label": label, "hassei": hassei, "out_dir": out_dir,
            "files": result_files, "yokakunin_path": yk_path, "yokakunin_md": "\n".join(yk_lines),
            "people": sum(len(v) for v in pay_tx_people.values()), "rates": rates, "rates_src": rates_src,
            "overwritten": overwritten, "input_src": input_src,
            "alerts": {k: len(alerts[k]) for k in MD_ALERT_KEYS if alerts.get(k)}}
