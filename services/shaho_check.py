r"""標準報酬月額チェックの後半（保険料計算・控除突合・判定・レポート）。

- 保険料 = 基礎標報 × 本人負担率 を丸め設定（既定: 50銭以下切捨て・超切上げ）で計算する。
- 控除突合: C月明細の本人控除4種は **(C−lag)月分**（jinjer設定「翌月徴収」）。
  基礎標報は (C−lag)月スナップショットの登録値を使う（過去月の再取得はしない前提）。
- 判定は 標報判定・控除判定 の2系統＋総合。**要確認系はいかなる場合も自動OKに昇格しない。**
"""

from __future__ import annotations

import datetime
import json
import os
from dataclasses import dataclass, field
from decimal import ROUND_CEILING, ROUND_FLOOR, ROUND_HALF_UP, Decimal

from config import Config
from services.keiri_engine import is_shaho_menjo_prev, ym_add, ym_compact
from services.shaho_engine import (BI_KENPO_SMR, BI_KONEN_SMR, MONTHLY_SYSTEMS,  # noqa: F401
                                   assess_month, calc_teiji_kettei, load_statements_full,
                                   registered_smr, salary_system_name)
from services.shaho_master import (ShahoMasterError, load_class_master, load_grade_table,
                                   select_grade_table)

# 判定ステータス（強い順。総合は2系統の強い方）
STATUS_PRIORITY = ["INSUFFICIENT_DATA", "EXEMPTION_REVIEW", "TWO_MONTH_COLLECTION_REVIEW",
                   "ADJUSTMENT_PRESENT", "MONTHLY_REVISION_CANDIDATE", "DIFFERENCE",
                   "PROVISIONAL_OK", "OK", "NOT_APPLICABLE"]
REVIEW_STATUSES = set(STATUS_PRIORITY[:6])          # 要確認系（自動OKにしない）
STATUS_JA = {
    "OK": "OK", "PROVISIONAL_OK": "仮OK", "DIFFERENCE": "差異",
    "ADJUSTMENT_PRESENT": "社保調整あり", "EXEMPTION_REVIEW": "免除の確認",
    "TWO_MONTH_COLLECTION_REVIEW": "2か月徴収の確認",
    "MONTHLY_REVISION_CANDIDATE": "随時改定の候補",
    "INSUFFICIENT_DATA": "情報不足", "NOT_APPLICABLE": "対象外",
}

D_KENPO = "salary_deduction_items:deduction29"
D_KAIGO = "salary_deduction_items:deduction30"
D_KONEN = "salary_deduction_items:deduction31"
D_KODOMO = "salary_deduction_items:child_support"
D_CHOSEI = "salary_deduction_items:deduction4"


def _pi_value(pi, key):
    from services.keiri_engine import pi_value
    return pi_value(pi, key)


def round_premium(amount: float, mode: str) -> int:
    """本人負担額の円未満処理。既定 50sen ＝ 50銭以下切捨て・50銭超切上げ（健保法の原則）。"""
    d = Decimal(str(amount))
    if mode == "floor":
        return int(d.to_integral_value(rounding=ROUND_FLOOR))
    if mode == "ceil":
        return int(d.to_integral_value(rounding=ROUND_CEILING))
    if mode == "round":
        return int(d.to_integral_value(rounding=ROUND_HALF_UP))
    # 50sen: 端数がちょうど0.50なら切捨て、超えたら切上げ
    frac = d - int(d)
    return int(d) if frac <= Decimal("0.5") else int(d) + 1


def kaigo_applies(basic_info) -> bool:
    name = str(((basic_info or {}).get("care_insurance_calculation_classification") or {})
               .get("name") or "")
    return "第2号" in name


def konen_applies(basic_info) -> bool:
    name = str((((basic_info or {}).get("employees_pension") or {})
                .get("calculation_classification") or {}).get("name") or "")
    return bool(name) and "対象外" not in name


def kenpo_applies(basic_info) -> bool:
    name = str(((basic_info or {}).get("health_insurance_calculation_classification") or {})
               .get("name") or "")
    return bool(name) and "対象外" not in name


def calc_premiums(kenpo_smr: float, konen_smr: float, basic_info, master, rounding: str) -> dict:
    """本人負担4種。介護は第2号のみ・厚年は対象者のみ・健保対象外は全て0。"""
    if not kenpo_applies(basic_info) or kenpo_smr <= 0:
        return {"kenpo": 0, "kodomo": 0, "kaigo": 0, "konen": 0}
    r = master.rates
    return {
        "kenpo": round_premium(kenpo_smr * r["kenpo"].employee, rounding),
        "kodomo": round_premium(kenpo_smr * r["kodomo"].employee, rounding),
        "kaigo": (round_premium(kenpo_smr * r["kaigo"].employee, rounding)
                  if kaigo_applies(basic_info) else 0),
        "konen": (round_premium(konen_smr * r["konen"].employee, rounding)
                  if konen_applies(basic_info) and konen_smr > 0 else 0),
    }


def merge_status(*statuses: str) -> str:
    return min((s for s in statuses if s), key=STATUS_PRIORITY.index)


@dataclass
class PersonResult:
    emp: str
    name: str = ""
    system: str = ""
    teiji_status: str = "NOT_APPLICABLE"
    check_status: str = "NOT_APPLICABLE"
    notes: list = field(default_factory=list)
    teiji: object = None                    # TeijiKettei
    reg_kenpo: int = 0
    reg_konen: int = 0
    premiums: dict = field(default_factory=dict)      # 期待値（C−lag月分）
    actuals: dict = field(default_factory=dict)       # C月明細の控除実績
    chosei: float = 0.0

    @property
    def total_status(self) -> str:
        return merge_status(self.teiji_status, self.check_status)


def _parse_date(s):
    try:
        return datetime.date.fromisoformat(str(s))
    except (TypeError, ValueError):
        return None


def month_closed_stats(month_map: dict) -> dict:
    """その月の給与確定状況（payroll_info.is_payroll_closed）。

    **未確定の月は登録標準報酬月額も報酬額もまだ動く**ので、その月を含む算定・突合は
    信用しない。確定済みの月は値が凍結される（2026-06 を3週間後に再取得して実社員232名が
    完全一致することを実測。動いたのは未確定のテスト社員1名だけ）。
    """
    closed, open_emps = 0, []
    for emp, rec in month_map.items():
        if (rec.get("payroll_info") or {}).get("is_payroll_closed"):
            closed += 1
        else:
            open_emps.append(emp)
    return {"closed": closed, "open": len(open_emps), "open_emps": sorted(open_emps)}


def judge_person(emp, months_data, check_pair, year, master, class_master, cfg) -> PersonResult:
    """1人分の判定。months_data={ym: rec}（4〜6月）、check_pair=(C−1月rec, C月rec)。"""
    res = PersonResult(emp=emp)
    calc_months = [f"{year}-04", f"{year}-05", f"{year}-06"]
    recs = [months_data.get(m) for m in calc_months]
    bi = next((r["basic_info"] for r in recs if r), None) or {}
    for r in recs:
        if r:
            bi = r["basic_info"]                 # 最新月の basic_info を代表にする
    res.name = ("%s %s" % (bi.get("last_name", ""), bi.get("first_name", ""))).strip()
    res.system = salary_system_name(bi)
    res.reg_kenpo, res.reg_konen = (int(v) for v in registered_smr(bi))

    # ---------- 標報判定 ----------
    joined = _parse_date(bi.get("joined_on"))
    retired = _parse_date(bi.get("retirement_date"))
    if not any(recs):
        res.teiji_status = "NOT_APPLICABLE"
        res.notes.append("4〜6月に給与明細がない")
    elif res.system == "テスト":
        res.teiji_status = "NOT_APPLICABLE"
        res.notes.append("給与体系が「テスト」")
    elif not kenpo_applies(bi) and res.reg_kenpo == 0:
        res.teiji_status = "NOT_APPLICABLE"
        res.notes.append("社会保険の対象外（健保区分・登録標報とも対象外）")
    elif joined and joined >= datetime.date(year, 6, 1):
        res.teiji_status = "NOT_APPLICABLE"
        res.notes.append(f"{joined:%m/%d}入社（6/1以降の資格取得は定時決定の対象外）")
    elif retired and retired <= datetime.date(year, 8, 31):
        res.teiji_status = "NOT_APPLICABLE"
        res.notes.append(f"{retired:%m/%d}退職（9月適用前に資格喪失）")
    else:
        assessments = [assess_month(m, r["payroll_info"], salary_system_name(r["basic_info"]),
                                    class_master, cfg["threshold"])
                       for m, r in zip(calc_months, recs) if r]
        tk = calc_teiji_kettei(assessments, master)
        res.teiji = tk
        menjo = _menjo_months(calc_months, months_data)
        multi = any(r["n_nonzero"] > 1 for r in recs if r)
        # 給与が未確定の算定月は報酬額がまだ動くので、その人の算定は信用しない
        open_calc = [m for m in calc_months
                     if emp in (cfg.get("open_months") or {}).get(m, [])]
        if tk.gate_ng or tk.unclassified or multi or tk.adopted_n == 0 or open_calc:
            res.teiji_status = "INSUFFICIENT_DATA"
            if open_calc:
                res.notes.append("算定月の給与が未確定（" + "、".join(open_calc)
                                 + "）。確定してから再実行してください")
            if tk.gate_ng:
                res.notes.append("検算ゲート不一致の月がある（報酬計≠雇用保険対象額）")
            if tk.unclassified:
                res.notes.append("分類できない支給項目に金額がある")
            if multi:
                res.notes.append("同じ月に給与明細が複数ある")
            if tk.adopted_n == 0:
                res.notes.append("採用できる月がない（"
                                 + "／".join(a.reason for a in tk.months if a.reason) + "）")
        elif menjo:
            res.teiji_status = "EXEMPTION_REVIEW"
            res.notes.append("算定月に社保免除（産休・育休）の月がある: " + "、".join(menjo))
        elif _revision_candidate(res, months_data, master, class_master, cfg):
            res.teiji_status = "MONTHLY_REVISION_CANDIDATE"
        elif tk.kenpo_smr != res.reg_kenpo or tk.konen_smr != res.reg_konen:
            res.teiji_status = "DIFFERENCE"
            res.notes.append("計算標報と登録標報が不一致（9月適用前は「改定予定」の意味）")
        elif tk.approx_used or tk.adopted_n < 3:
            res.teiji_status = "PROVISIONAL_OK"
            if tk.approx_used:
                res.notes.append("欠勤月の支払基礎日数が概算（所定日数がAPIに無い）")
            if tk.adopted_n < 3:
                res.notes.append(f"採用月が{tk.adopted_n}か月")
        else:
            res.teiji_status = "OK"

    # ---------- 控除判定（C月明細 =（C−lag）月分） ----------
    prev_rec, c_rec = check_pair
    if c_rec is None:
        res.check_status = "NOT_APPLICABLE"
        res.notes.append("突合月に給与明細がない")
        return res
    c_pi = c_rec["payroll_info"]
    res.actuals = {"kenpo": _pi_value(c_pi, D_KENPO), "kaigo": _pi_value(c_pi, D_KAIGO),
                   "konen": _pi_value(c_pi, D_KONEN), "kodomo": _pi_value(c_pi, D_KODOMO)}
    res.chosei = _pi_value(c_pi, D_CHOSEI)
    base_bi = (prev_rec or c_rec)["basic_info"]
    base_kenpo, base_konen = registered_smr(base_bi)
    res.premiums = calc_premiums(base_kenpo, base_konen, base_bi, master, cfg["rounding"])

    actual_sum = sum(res.actuals.values())
    open_check = [m for m in (ym_add(cfg["check_month"], -cfg["lag"]), cfg["check_month"])
                  if emp in (cfg.get("open_months") or {}).get(m, [])]
    if open_check:
        res.check_status = "INSUFFICIENT_DATA"
        res.notes.append("突合に使う月の給与が未確定（" + "、".join(open_check)
                         + "）。確定してから再実行してください")
    elif base_kenpo == 0 and actual_sum == 0:
        res.check_status = "NOT_APPLICABLE"
    elif prev_rec and is_shaho_menjo_prev(prev_rec["payroll_info"], c_pi):
        res.check_status = "EXEMPTION_REVIEW"
        res.notes.append("控除対象月が社保免除（産休・育休）。控除0が正でも自動OKにしない")
    elif retired and retired.strftime("%Y-%m") >= ym_add(cfg["check_month"], -cfg["lag"]):
        res.check_status = "TWO_MONTH_COLLECTION_REVIEW"
        res.notes.append("退職者（2か月徴収・資格喪失の可能性）。目視で確認")
    elif res.chosei:
        res.check_status = "ADJUSTMENT_PRESENT"
        res.notes.append(f"社保調整 {int(res.chosei):+,}円 がある")
        if res.system == "管理監督者":
            res.notes.append("⚠管理監督者はマイナスの社保調整が0に潰される設定（一致でも信用しない）")
    elif base_kenpo == 0 and actual_sum != 0:
        res.check_status = "INSUFFICIENT_DATA"
        res.notes.append("登録標報が0なのに控除がある")
    else:
        diffs = {k: res.actuals[k] - res.premiums[k] for k in res.premiums}
        if any(abs(d) > cfg["tolerance"] for d in diffs.values()):
            res.check_status = "DIFFERENCE"
            res.notes.append("保険料の差: " + "、".join(
                f"{k}{d:+,.0f}円" for k, d in diffs.items() if abs(d) > cfg["tolerance"]))
        elif any(abs(d) > 0.5 for d in diffs.values()):
            res.check_status = "PROVISIONAL_OK"
            res.notes.append("±1円以内（端数処理の差）")
        else:
            res.check_status = "OK"
    return res


def _menjo_months(calc_months, months_data):
    """算定月のうち社保免除（前月分の会社負担あり・当月控除ゼロ）の月。"""
    out = []
    for m in calc_months:
        rec, prev = months_data.get(m), months_data.get(ym_add(m, -1))
        if rec and prev and is_shaho_menjo_prev(prev["payroll_info"], rec["payroll_info"]):
            out.append(m)
    return out


# 固定的賃金の変動検知から外すキー。通勤費は固定的賃金だが、実費精算の人は毎月
# 金額が動く（単価が変わったわけではない）ので、変動検知に入れると全員が候補になる。
REVISION_IGNORE_KEYS = {"salary_items:allowance34", "salary_items:allowance35"}


def _revision_candidate(res, months_data, master, class_master, cfg) -> bool:
    """4〜6月に固定的賃金が変動し、変動月からの3か月平均が登録と2等級以上差 → 候補。

    変動月を4〜6月に限るのは「7〜9月に随時改定される予定の人は定時決定を確定扱い
    しない」ため（4月変動→7月改定 … 6月変動→9月改定）。
    """
    from services.shaho_engine import collect_remuneration
    months = sorted(m for m in months_data if months_data.get(m))
    fixed, totals = {}, {}
    for m in months:
        rem = collect_remuneration(months_data[m]["payroll_info"], class_master, m)
        core = sum(v for key, _lab, v, cls, fx in rem.breakdown
                   if cls == "対象" and fx == "1" and key not in REVISION_IGNORE_KEYS)
        fixed[m], totals[m] = core, rem.total
    changes = [m for prev, m in zip(months, months[1:])
               if abs(fixed[m] - fixed[prev]) > 0.5 and m in cfg["revision_window"]]
    for change in changes:
        window = [ym_add(change, i) for i in range(3)]
        if not all(w in totals for w in window):
            res.notes.append(f"固定的賃金が{change}に変動（3か月窓が未完＝随時改定の様子見）")
            return True
        avg = int(sum(totals[w] for w in window) / 3)
        row = master.find_grade(avg)
        reg_row = next((g for g in master.grades if g.kenpo_smr == res.reg_kenpo), None)
        if reg_row and abs(row.kenpo_grade - reg_row.kenpo_grade) >= 2:
            res.notes.append(f"固定的賃金が{change}に変動し3か月平均{avg:,}円は"
                             f"登録と{abs(row.kenpo_grade - reg_row.kenpo_grade)}等級差"
                             "（随時改定候補）")
            return True
    return False


def detect_revisions(months_all: dict) -> list:
    """月次スナップショット間で登録標報が変わった人（期中改定の検知）。"""
    out = []
    months = sorted(months_all)
    for a, b in zip(months, months[1:]):
        common = set(months_all[a]) & set(months_all[b])
        for emp in sorted(common):
            ka, na = registered_smr(months_all[a][emp]["basic_info"])
            kb, nb = registered_smr(months_all[b][emp]["basic_info"])
            if (ka, na) != (kb, nb):
                bi = months_all[b][emp]["basic_info"]
                out.append({"社員番号": emp,
                            "氏名": ("%s %s" % (bi.get("last_name", ""),
                                              bi.get("first_name", ""))).strip(),
                            "検知区間": f"{a}→{b}",
                            "健保標報": f"{int(ka):,}→{int(kb):,}",
                            "厚年標報": f"{int(na):,}→{int(nb):,}", "区分": "情報"})
    return out


# ---------------------------------------------------------------------------
# 実行本体
# ---------------------------------------------------------------------------
def run_check(year: int, check_month: str, insurer: str = None, out_base: str = None,
              grade_xlsx: str = None, class_csv: str = None) -> dict:
    insurer = insurer or Config.SHAHO_INSURER
    out_base = out_base or Config.SHAHO_OUTPUT_DIR
    if grade_xlsx is None:
        # 控除を突き合わせる保険料の月（支給月 − 徴収ラグ）に合う等級表を選ぶ。
        # Codex の公式資料（標準報酬月額_YYYY_MM.xlsx）があればそれ、無ければ設定の手作りブック
        premium_ym = ym_add(check_month, -Config.SHAHO_DEDUCTION_LAG_MONTHS)
        grade_xlsx = select_grade_table(premium_ym, Config.SHAHO_GRADE_TABLE_XLSX).path
    master = load_grade_table(grade_xlsx, insurer, year)
    class_master = load_class_master(class_csv or Config.SHAHO_CLASS_MASTER_CSV)
    cfg = {"threshold": Config.SHAHO_BASE_DAYS_THRESHOLD,
           "rounding": Config.SHAHO_ROUNDING,
           "tolerance": Config.SHAHO_PREMIUM_TOLERANCE,
           "lag": Config.SHAHO_DEDUCTION_LAG_MONTHS,
           "check_month": check_month,
           # 変動月としてみる範囲＝4〜6月（7〜9月に随時改定が効く変動）
           "revision_window": {f"{year}-04", f"{year}-05", f"{year}-06"}}

    need = [f"{year}-04", f"{year}-05", f"{year}-06",
            ym_add(check_month, -1), check_month]
    months_all = {}
    for ym in dict.fromkeys(need):
        path = os.path.join(Config.KEIRI_OUTPUT_DIR, "raw", f"salary_statements_{ym}.json")
        if not os.path.exists(path):
            raise ShahoMasterError(
                f"{ym} の給与明細キャッシュがありません: {path}\n"
                "経理モードで取得するか shaho_check_run.py --fetch-missing を使ってください")
        months_all[ym] = load_statements_full(Config.KEIRI_OUTPUT_DIR, ym)

    # 給与が未確定の月が混じっていないか（混じっているとスナップショットがまだ動く）
    month_status = {ym: month_closed_stats(m) for ym, m in months_all.items()}
    cfg["open_months"] = {ym: s["open_emps"] for ym, s in month_status.items() if s["open"]}

    everyone = sorted(set().union(*[set(v) for v in months_all.values()]))
    prev_m, c_m = need[3], need[4]
    results = []
    for emp in everyone:
        # 全キャッシュ月を渡す（月変検知の3か月窓が突合月まで届くように）
        per_month = {m: months_all[m].get(emp) for m in months_all}
        pair = (months_all[prev_m].get(emp), months_all[c_m].get(emp))
        results.append(judge_person(emp, per_month, pair, year, master, class_master, cfg))
    revisions = detect_revisions(months_all)
    return {"year": year, "check_month": check_month, "insurer": insurer,
            "master": master, "class_master": class_master, "cfg": cfg,
            "results": results, "revisions": revisions, "out_base": out_base,
            "month_status": month_status}
