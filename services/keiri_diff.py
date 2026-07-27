r"""経理モード: 生成した4CSV と 経理担当の最終CSV を突合する。

Z:\API連携\scripts\keiri_diff_generated_vs_final.py から移植（2026-07-27）。

突合方法は「従業員×(品目,勘定科目,税区分) の合計額」。取引の分割の仕方が違っても
合計は保存されるため、分割規則の差に影響されずに中身を比較できる。
"""

from __future__ import annotations

import csv
import glob
import io
import os
import re
from collections import Counter, defaultdict

from config import Config
from services.keiri_api import normalize_label, to_number

TYPES = ["給与", "住民税", "健康保険", "厚生年金"]

# 名称ゆれ（同じものを指す別表記）。左辺を右辺に寄せて突合し、差は「名称ゆれ」として別途報告する。
#   どちらも 2026-07 の最終CSVだけに出る表記で、2026-04〜06 は右辺で統一されている。
ITEM_ALIASES = {"子ども・子育て支援金": "子ども・子育て支援金（預り分）"}
EMP_ALIASES = {"VerMartin": "ブアーマーティン"}


def read_rows(path):
    raw = open(path, "rb").read()
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            return list(csv.reader(io.StringIO(raw.decode(enc))))
        except UnicodeDecodeError:
            continue
    return list(csv.reader(io.StringIO(raw.decode("cp932", errors="replace"))))


def norm_date(s):
    m = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", str(s or "").strip())
    return f"{int(m.group(1))}/{int(m.group(2))}/{int(m.group(3))}" if m else str(s or "").strip()


def parse(path, by_date=False, variants=None):
    """{従業員: {キー: 合計額}} と 取引ヘッダ一覧を返す。

    by_date=True のときキーに発生日を含める（暫定取引と実績差分取引を分けて見る）。
    品目名・従業員名は名称ゆれを代表名へ寄せ、寄せた実績を variants に記録する。
    """
    rows = read_rows(path)
    idx = {c: i for i, c in enumerate(rows[0]) if c}
    per = defaultdict(lambda: defaultdict(float))
    heads = []
    cur_date = ""
    for r in rows[1:]:
        def v(col, r=r):
            i = idx.get(col)
            return r[i].strip() if i is not None and i < len(r) else ""
        if not any(str(x).strip() for x in r):
            continue
        if v("収支区分"):
            cur_date = norm_date(v("発生日"))
            heads.append((v("管理番号"), cur_date, norm_date(v("支払期日")), v("取引先")))
        amt = to_number(v("金額"))
        if amt is None:
            continue
        item, emp = v("品目"), normalize_label(v("従業員"))
        if item in ITEM_ALIASES:
            if variants is not None:
                variants.add(("品目", item, ITEM_ALIASES[item]))
            item = ITEM_ALIASES[item]
        if emp in EMP_ALIASES:
            if variants is not None:
                variants.add(("従業員", emp, EMP_ALIASES[emp]))
            emp = EMP_ALIASES[emp]
        key = (item, v("勘定科目"), v("税区分"))
        if by_date:
            key = (cur_date,) + key
        per[emp][key] += amt
    return per, heads


def pick(dirpath, ftype):
    cands = [p for p in glob.glob(os.path.join(dirpath, "*.csv"))
             if ftype in os.path.basename(p) and "賞与" not in os.path.basename(p)]
    return max(cands, key=os.path.getsize) if cands else None


def compare(gen, fin):
    res = {"match": 0, "diff": [], "gen_only": [], "fin_only": [],
           "emp_gen_only": set(), "emp_fin_only": set()}
    res["emp_gen_only"] = set(gen) - set(fin)
    res["emp_fin_only"] = set(fin) - set(gen)
    for emp in set(gen) & set(fin):
        g, f = gen[emp], fin[emp]
        for key in set(g) | set(f):
            gv, fv = g.get(key), f.get(key)
            if gv is not None and fv is not None:
                if abs(gv - fv) < 0.5:
                    res["match"] += 1
                else:
                    res["diff"].append((emp, key, gv, fv))
            elif gv is not None:
                res["gen_only"].append((emp, key, gv))
            else:
                res["fin_only"].append((emp, key, fv))
    return res


def mask(name):
    return (name[:1] + "＊＊") if name else "（空欄）"


def compare_month(month, gen_dir, final_dir=None, by_date=False, detail=15):
    """1か月ぶんの突合。 (summary, markdown行リスト) を返す。

    summary: [{"種別","一致","金額差","生成のみ","最終のみ","生成だけの人","最終だけの人"}]
    最終CSVのフォルダが無い月は「未突合」として summary に残す（エラーにしない）。
    """
    mc = month.replace("-", "")
    fin_dir = final_dir or os.path.join(Config.KEIRI_FINAL_CSV_DIR, f"{int(month[5:7])}月", "freee")
    lines = [f"# 生成 vs 最終CSV 差分 {mc}", "",
             "突合: 従業員×(品目,勘定科目,税区分) の合計額（取引分割の違いに影響されない方式）", "",
             f"- 生成: `{gen_dir}`", f"- 最終: `{fin_dir}`", "",
             "| 種別 | 一致 | 金額差 | 生成のみ | 最終のみ | 生成だけの人 | 最終だけの人 |",
             "|---|---|---|---|---|---|---|"]
    summary, details = [], []
    variants = defaultdict(set)
    for ftype in TYPES:
        gp, fp = pick(gen_dir, ftype), pick(fin_dir, ftype) if os.path.isdir(fin_dir) else None
        if not gp or not fp:
            lines.append(f"| {ftype} | 未突合（生成={bool(gp)} 最終={bool(fp)}） | | | | | |")
            summary.append({"種別": ftype, "状態": "未突合"})
            continue
        gen, gheads = parse(gp, by_date)
        fin, fheads = parse(fp, by_date, variants[ftype])
        r = compare(gen, fin)
        lines.append(f"| {ftype} | {r['match']} | {len(r['diff'])} | {len(r['gen_only'])} | "
                     f"{len(r['fin_only'])} | {len(r['emp_gen_only'])} | {len(r['emp_fin_only'])} |")
        summary.append({"種別": ftype, "状態": "突合済み", "一致": r["match"],
                        "金額差": len(r["diff"]), "生成のみ": len(r["gen_only"]),
                        "最終のみ": len(r["fin_only"]),
                        "生成だけの人": len(r["emp_gen_only"]),
                        "最終だけの人": len(r["emp_fin_only"])})
        details.append((ftype, os.path.basename(fp), r, gheads, fheads))

    for ftype, fname, r, gheads, fheads in details:
        lines += ["", f"## {ftype}（最終: `{fname}`）", "",
                  f"- 取引数: 生成 {len(gheads)} / 最終 {len(fheads)}"]
        if r["diff"]:
            agg = Counter(k[:-1] for _e, k, _g, _f in r["diff"])
            head = "| 発生日 | 品目 | 勘定科目 | 件数 |" if by_date else "| 品目 | 勘定科目 | 件数 |"
            sep = "|---|---|---|---|" if by_date else "|---|---|---|"
            lines += ["", f"### 金額差 {len(r['diff'])}件 — 内訳", "", head, sep]
            for k, n in agg.most_common(12):
                lines.append("| " + " | ".join(k) + f" | {n} |")
            lines += ["", "明細（差の大きい順）:", "",
                      "| 従業員 | " + ("発生日 | " if by_date else "")
                      + "品目 | 勘定科目 | 生成 | 最終 | 差 |",
                      "|---|---|---|---|---|---|" + ("---|" if by_date else "")]
            for emp, key, gv, fv in sorted(r["diff"], key=lambda x: -abs(x[2] - x[3]))[:detail]:
                lines.append(f"| {mask(emp)} | {' | '.join(key[:-1])} | {gv:,.0f} | "
                             f"{fv:,.0f} | {gv - fv:+,.0f} |")
        for label, key in (("生成のみ（余分に作っている）", "gen_only"),
                           ("最終のみ（作れていない）", "fin_only")):
            if r[key]:
                agg = Counter((k[0], k[1]) for _e, k, _v in r[key])
                sums = defaultdict(float)
                for _e, k, v in r[key]:
                    sums[(k[0], k[1])] += v
                lines += ["", f"### {label} {len(r[key])}件", "",
                          "| 品目 | 勘定科目 | 件数 | 合計額 |", "|---|---|---|---|"]
                for (item, acc), n in agg.most_common(12):
                    lines.append(f"| {item} | {acc} | {n} | {sums[(item, acc)]:,.0f} |")
        for label, key in (("生成だけに出る人", "emp_gen_only"), ("最終だけに出る人", "emp_fin_only")):
            if r[key]:
                lines.append(f"- {label}: " + "、".join(mask(e) for e in sorted(r[key])[:12])
                             + (f" 他{len(r[key]) - 12}名" if len(r[key]) > 12 else ""))

    if any(variants.values()):
        lines += ["", "## 名称ゆれ（最終CSV側の表記が他月と違う。突合では代表名へ寄せた）", "",
                  "| 種別 | 列 | 最終CSVの表記 | 代表名（生成側） |", "|---|---|---|---|"]
        for ftype in TYPES:
            for col, raw, canon in sorted(variants.get(ftype, ())):
                lines.append(f"| {ftype} | {col} | {raw} | {canon} |")
        lines += ["", "→ どちらの表記が freee の正しいマスタか要確認（谷津さん）。"]

    out = os.path.join(gen_dir, f"diff_{mc}.md")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return summary, lines, out
