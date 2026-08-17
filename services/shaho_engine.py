r"""標準報酬月額チェックのエンジン（報酬集計・支払基礎日数・定時決定）。

4〜6月に実際に支給された報酬（jinjer給与明細）から、9月適用予定の標準報酬月額を
計算する。**jinjer へは読みに行くだけで書き込まない。**

設計の要点:
- statements キャッシュは経理モードと共用（outputs/keiri/raw/salary_statements_{ym}.json）。
  basic_info の登録標準報酬月額は**その支給月の値**で、給与が確定した月
  （payroll_info.is_payroll_closed=True）は凍結される。2026-06 を3週間後に再取得して
  実社員232名が完全一致することを実測で確認済み（動いたのは未確定の
  テスト社員9999999だけ）。したがって過去月の再取得をしても期中改定の検知は壊れない。
  ただし**未確定の月は値が動く**ので、その月のスナップショットは信用しない。
- どの支給項目を報酬に入れるかは報酬分類マスタ（shaho_master.load_class_master）が決める。
  未分類・「未設定」の項目に金額があったら、その人は自動でOKにせず要確認へ落とす。
- **検算ゲート**: 報酬計（分類=対象の合計）は jinjer の雇用保険対象額(other5) と一致する
  はず（立替金・経費・情報項目は元々 other5 に入らない。2026-04〜08 の実データで確認済み）。
  1円でもズレたらその人の算定は信用しない（INSUFFICIENT_DATA）。
"""

from __future__ import annotations

import calendar
import unicodedata
from dataclasses import dataclass, field

from services.keiri_engine import fetch_statements, pi_item, to_number
from services.shaho_master import ShahoMaster, resolve_class

# 検算の基準（雇用保険対象額。0円のときは労災対象額で代用）
GROSS_KEYS = ("salary_other_items:other5", "salary_other_items:other6")
# 支払基礎日数の材料は **ラベルで引く**。項目IDは月によって意味が変わるため。
#   2026-04（旧体系）: kintai6=出勤日数 / kintai7=欠勤日数 / kintai8=前月有給消化日数
#   2026-05以降      : kintai10=出勤日数 / kintai11=欠勤日数 / kintai13=前月有休消化日数
# IDで引くと4月が全員「出勤日数0日」になり、時給制の4月が丸ごと算定から外れる
# （2026-08-17 に実データで発覚。齋藤2011001・MAHARJAN2022002 ほか7名の等級ズレの原因）。
# 「有給」と「有休」の表記ゆれもあるので、正規化して部分一致で拾う。
LABEL_SHUKKIN = ("出勤日数",)
LABEL_KEKKIN = ("欠勤日数",)
LABEL_YUKYU = ("前月有給消化日数", "前月有休消化日数", "有給消化日数", "有休消化日数")

# ⚠ 2026-05 支給分の時給制だけは**ラベルすら信用できない**。
# 体系移行の途中で、値は旧配置のままラベルだけ新体系に差し替わっている:
#   kintai6  ラベル「内前月実績超過勤怠時間60時間以上」だが中身は **出勤日数**
#            （46名・中央値20.0日。4月の出勤日数と同じ分布。kintai10 は空のまま）
#   kintai8  ラベル「前月の法定休日労働時間」だが中身は **前月有給消化日数**（kintai13と完全一致）
#   kintai9  ラベルなしだが中身は **有給休暇残**（kintai14と完全一致）
#   kintai11 ラベル「欠勤日数」だが中身は **総法定外残業時間**（最大60。日数ではない）
# 有休まわりは新スロットへコピーされたのに出勤日数だけ取り残された移行漏れ。
# 過去month なので値はもう動かない。該当月だけスロットを明示して読む。
SLOT_OVERRIDES = {
    ("2026-05", "時給制"): {"出勤": "kintai6", "欠勤": None, "有休": "kintai13"},
}


def _slot_override(system: str, ym: str):
    for (o_ym, o_sys), slots in SLOT_OVERRIDES.items():
        if ym == o_ym and system.startswith(o_sys):
            return slots
    return None
# 登録済みの標準報酬月額（basic_info。円）
BI_KENPO_SMR = "health_insurance"
BI_KONEN_SMR = "employee_pension"

# 「140-180時間制暫定」「時給制暫定」は 2026-04 支給分（暫定支給の初月）だけに現れる旧名
MONTHLY_SYSTEMS = ("月給制1", "月給制2", "月給制3", "管理監督者", "140-180時間制暫定")
HOURLY_SYSTEMS = ("時給制1", "時給制暫定")


def flatten_items(pi) -> dict:
    """payroll_info → {source_key: (float値, 体系別名)}。値が数値でない項目は0扱い。"""
    out = {}
    for array_type, items in (pi or {}).items():
        if not isinstance(items, list):
            continue
        for it in items:
            if not isinstance(it, dict):
                continue
            key = "%s:%s" % (array_type, it.get("id"))
            label = (str(it.get("salary_system_label") or "").strip()
                     or str(it.get("label") or "").strip())
            n = to_number(it.get("value"))
            out[key] = (0.0 if n is None else n, label)
    return out


def load_statements_full(cache_dir, ym, client=None, refresh=False) -> dict:
    """{社員番号: {"basic_info", "payroll_info", "n_nonzero"}} を返す。

    keiri_engine.load_statements は payroll_info しか返さないので、basic_info
    （氏名・給与体系・登録標報・生年月日・各種区分）も保持する版。statement の選択規則
    （基本給 allowance1 非ゼロを採用）は同じだが、**非ゼロ基本給の statement が2つ以上
    ある人は合算も推測もせず n_nonzero で知らせる**（判定側が要確認に落とす）。
    """
    from services.keiri_api import classify_employee
    data = fetch_statements(cache_dir, ym, client=client, refresh=refresh)
    out = {}
    for person in data or []:
        emp = str(person.get("employee_id", "")).strip()
        if classify_employee(emp) != "target":
            continue
        statements = person.get("statements") or []
        if not statements:
            continue
        nonzero = []
        for st in statements:
            pi = st.get("payroll_info") or {}
            for it in pi.get("salary_items") or []:
                if str(it.get("id")) == "allowance1" and to_number(it.get("value")):
                    nonzero.append(st)
                    break
        best = (nonzero or statements)[0]
        out[emp] = {
            "basic_info": best.get("basic_info") or {},
            "payroll_info": best.get("payroll_info") or {},
            "n_nonzero": len(nonzero),
        }
    return out


def salary_system_name(basic_info) -> str:
    return str(((basic_info or {}).get("salary_system") or {}).get("name") or "").strip()


def registered_smr(basic_info) -> tuple[float, float]:
    """登録済みの (健保標準報酬月額, 厚年標準報酬月額)。0=社保対象外。"""
    k = to_number((basic_info or {}).get(BI_KENPO_SMR))
    n = to_number((basic_info or {}).get(BI_KONEN_SMR))
    return (0.0 if k is None else k, 0.0 if n is None else n)


# ---------------------------------------------------------------------------
# 報酬集計（1人月）
# ---------------------------------------------------------------------------
@dataclass
class Remuneration:
    cash_total: float = 0.0                # 分類=対象（通貨報酬）の合計
    genbutsu_total: float = 0.0            # 分類=現物（現物給与）の合計
    fixed_total: float = 0.0               # うち固定的賃金（fixed=1）の合計
    gross_ref: float = 0.0                 # 検算基準（other5、0なら other6）
    gate_diff: float = 0.0                 # cash_total − gross_ref（±0.5超は信用しない）
    breakdown: list = field(default_factory=list)      # [(source_key, label, 金額, class, fixed)]
    unclassified: list = field(default_factory=list)   # 金額があるのに分類できない項目
    unresolved_fixed: list = field(default_factory=list)   # 対象なのに fixed が空の項目

    @property
    def total(self) -> float:
        """その月の報酬 ＝ 通貨 ＋ 現物。"""
        return self.cash_total + self.genbutsu_total

    @property
    def gate_ok(self) -> bool:
        return abs(self.gate_diff) <= 0.5


def collect_remuneration(pi, class_master, ym: str = "9999-99") -> Remuneration:
    """1明細の報酬算定。分類マスタで「対象」の支給項目を合算し、検算ゲートを通す。

    ym は支給月。適用期間つきの分類行（例: 時給制みなし給の再掲→実額の切替）を
    正しく引くために要る。
    """
    flat = flatten_items(pi)
    rem = Remuneration()
    for key, (value, label) in sorted(flat.items()):
        if not key.startswith("salary_items:"):
            continue
        rule = resolve_class(class_master, key, label, ym)
        if value == 0.0:
            continue
        if rule is None or rule.cls == "未設定":
            rem.unclassified.append({"source_key": key, "label": label, "金額": value,
                                     "理由": ("マスタに行が無い" if rule is None else "未設定")})
            continue
        if rule.cls == "対象外":
            rem.breakdown.append((key, label, value, "対象外", rule.fixed))
            continue
        if rule.cls == "現物":
            # 現物給与は報酬に数えるが、現金の総支給額の外側なので検算からは除く
            rem.genbutsu_total += value
            rem.breakdown.append((key, label, value, "現物", rule.fixed))
            continue
        rem.cash_total += value
        rem.breakdown.append((key, label, value, "対象", rule.fixed))
        if rule.fixed == "1":
            rem.fixed_total += value
        elif rule.fixed == "":
            rem.unresolved_fixed.append(key)
    gross = 0.0
    for gk in GROSS_KEYS:
        it = pi_item(pi, gk)
        n = to_number((it or {}).get("value"))
        if n:
            gross = n
            break
    rem.gross_ref = gross
    rem.gate_diff = rem.cash_total - gross
    return rem


# ---------------------------------------------------------------------------
# 支払基礎日数
# ---------------------------------------------------------------------------
@dataclass
class BaseDays:
    days: float | None                     # None = 判定材料が無い
    basis: str                             # 暦日 / 暦日-欠勤(概算) / 出勤+有給 / 不明
    approx: bool = False                   # 概算（所定日数がAPIに無いための代用）


def _kintai_value(pi, kintai_id) -> float | None:
    """勤怠項目をスロットID直接で読む（ラベルが信用できない月の補正用）。"""
    it = pi_item(pi, "salary_attendance_items:%s" % kintai_id)
    if it is None:
        return None
    n = to_number(it.get("value"))
    return 0.0 if n is None else n


def _kintai_by_label(pi, labels) -> float | None:
    """勤怠項目をラベルで引く。該当の項目が1つも無ければ None（0とは区別する）。"""
    wanted = {unicodedata.normalize("NFKC", s).replace(" ", "") for s in labels}
    for it in (pi or {}).get("salary_attendance_items") or []:
        if not isinstance(it, dict):
            continue
        label = (str(it.get("salary_system_label") or "").strip()
                 or str(it.get("label") or "").strip())
        if unicodedata.normalize("NFKC", label).replace(" ", "") in wanted:
            n = to_number(it.get("value"))
            return 0.0 if n is None else n
    return None


def payment_base_days(pi, system: str, ym: str) -> BaseDays:
    """支払基礎日数。月給者=暦日数（欠勤があれば−欠勤の概算）、時給者=出勤＋有休。

    材料はすべてラベルで引く（項目IDは月で意味が変わるため。上の定数のコメント参照）。
    """
    year, month = int(ym[:4]), int(ym[5:7])
    calendar_days = calendar.monthrange(year, month)[1]
    if system in MONTHLY_SYSTEMS:
        kekkin = _kintai_by_label(pi, LABEL_KEKKIN) or 0.0
        if kekkin > 0:
            # 正確には「就業規則の所定日数 − 欠勤日数」だが、所定日数は API に無い。
            # 暦日数−欠勤で代用し、概算フラグを立てる（採用しても最高 PROVISIONAL_OK）。
            return BaseDays(days=calendar_days - kekkin, basis="暦日-欠勤(概算)", approx=True)
        return BaseDays(days=calendar_days, basis="暦日")
    if system in HOURLY_SYSTEMS:
        override = _slot_override(system, ym)
        if override:
            # ラベルが当てにならない月。スロットを直接指定して読む
            shukkin = _kintai_value(pi, override["出勤"]) if override["出勤"] else None
            yukyu = _kintai_value(pi, override["有休"]) if override["有休"] else None
            days = (shukkin or 0.0) + (yukyu or 0.0)
            if shukkin is None and yukyu is None:
                return BaseDays(days=None, basis="不明（勤怠項目なし）")
            return BaseDays(days=days, basis="出勤+有給（%s の配置ずれを補正）" % ym)
        shukkin = _kintai_by_label(pi, LABEL_SHUKKIN)
        yukyu = _kintai_by_label(pi, LABEL_YUKYU)
        if shukkin is None and yukyu is None:
            return BaseDays(days=None, basis="不明（勤怠項目なし）")
        days = (shukkin or 0.0) + (yukyu or 0.0)
        # 出勤日数ゼロなのに報酬が出ている＝「0日働いた」ではなく jinjer 側の未入力を疑う。
        # 2026-05 支給分は時給制49名**全員**が出勤日数ゼロだった（体系移行月の入力漏れ）。
        # 0日として除外すると静かに算定から落ちるので、判定不能にして人の目に上げる。
        if days == 0:
            gross = 0.0
            for gk in GROSS_KEYS:
                it = pi_item(pi, gk)
                n = to_number((it or {}).get("value"))
                if n:
                    gross = n
                    break
            if gross > 0:
                return BaseDays(days=None, basis="不明（出勤日数が未入力の疑い）")
        return BaseDays(days=days, basis="出勤+有給")
    # 未知の給与体系。勝手に暦日扱いにせず「不明」で返す（判定側で要確認へ）
    return BaseDays(days=None, basis=f"不明（給与体系 {system or '空'}）")


# ---------------------------------------------------------------------------
# 定時決定（4〜6月の平均 → 等級）
# ---------------------------------------------------------------------------
@dataclass
class MonthAssessment:
    ym: str
    rem: Remuneration
    base_days: BaseDays
    adopted: bool = False
    reason: str = ""                       # 除外理由（採用時は空）


@dataclass
class TeijiKettei:
    months: list                            # MonthAssessment ×3（順序は算定月順）
    adopted_n: int = 0
    average: int | None = None              # 採用月平均（円未満切捨て）。採用0なら None
    kenpo_grade: int | None = None
    kenpo_smr: int | None = None
    konen_grade: int | None = None
    konen_smr: int | None = None
    approx_used: bool = False               # 概算基礎日数の月を採用した
    gate_ng: bool = False                   # 検算ゲートを割った月がある
    unclassified: list = field(default_factory=list)


def assess_month(ym, pi, system, class_master, threshold: int) -> MonthAssessment:
    """1人月の採用判定。基礎日数が閾値未満・判定不能・statements無しは除外。"""
    rem = collect_remuneration(pi, class_master, ym)
    bd = payment_base_days(pi, system, ym)
    ma = MonthAssessment(ym=ym, rem=rem, base_days=bd)
    if bd.days is None:
        ma.reason = f"支払基礎日数が判定できない（{bd.basis}）"
    elif bd.days < threshold:
        ma.reason = f"支払基礎日数 {bd.days:g}日 < {threshold}日"
    else:
        ma.adopted = True
    return ma


def calc_teiji_kettei(assessments: list, master: ShahoMaster) -> TeijiKettei:
    """採用月の報酬平均（円未満切捨て）→ 健保・厚年の等級。"""
    tk = TeijiKettei(months=list(assessments))
    adopted = [a for a in assessments if a.adopted]
    tk.adopted_n = len(adopted)
    tk.approx_used = any(a.base_days.approx for a in adopted)
    tk.gate_ng = any(not a.rem.gate_ok for a in assessments)
    for a in assessments:
        tk.unclassified.extend(a.rem.unclassified)
    if not adopted:
        return tk
    tk.average = int(sum(a.rem.total for a in adopted) / len(adopted))   # 円未満切捨て
    row = master.find_grade(tk.average)
    tk.kenpo_grade, tk.kenpo_smr = row.kenpo_grade, row.kenpo_smr
    tk.konen_grade, tk.konen_smr = row.konen_grade, row.konen_smr
    return tk
