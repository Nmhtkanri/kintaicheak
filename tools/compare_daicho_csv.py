# -*- coding: utf-8 -*-
"""派遣元管理台帳の一覧CSV 2本を突き合わせる（Phase 2 新旧パーサの比較検証用）。

使い方:
    python -X utf8 tools/compare_daicho_csv.py 旧_一覧.csv 新_一覧.csv [--max-detail 200]

キーは (契約No, 氏名)。出力は列ごとの差分件数と、行×列のセル差分の明細。
個人情報を含む出力になるため、リポジトリやdocsには件数サマリだけを転記すること。
"""
from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter
from pathlib import Path


def load(path: Path) -> dict[tuple[str, str], dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as fp:
        rows = list(csv.DictReader(fp))
    out: dict[tuple[str, str], dict] = {}
    for r in rows:
        key = ((r.get("契約No") or "").strip(), (r.get("氏名") or "").strip())
        if key in out:
            # 同一契約No・氏名が複数行（想定外）。連番を付けて衝突を避ける
            n = 2
            while (key[0] + f"#{n}", key[1]) in out:
                n += 1
            key = (key[0] + f"#{n}", key[1])
        out[key] = r
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="派遣元管理台帳 一覧CSVの比較")
    ap.add_argument("old_csv")
    ap.add_argument("new_csv")
    ap.add_argument("--max-detail", type=int, default=200, help="明細行の上限")
    args = ap.parse_args(argv)

    a = load(Path(args.old_csv))
    b = load(Path(args.new_csv))
    only_a = sorted(set(a) - set(b))
    only_b = sorted(set(b) - set(a))
    common = sorted(set(a) & set(b))

    print(f"旧: {len(a)}行 / 新: {len(b)}行 / 共通: {len(common)}行")
    if only_a:
        print(f"旧のみ {len(only_a)}行:")
        for k in only_a:
            print(f"  - {k[0]} {k[1]}")
    if only_b:
        print(f"新のみ {len(only_b)}行:")
        for k in only_b:
            print(f"  + {k[0]} {k[1]}")

    col_counter: Counter[str] = Counter()
    details: list[str] = []
    for key in common:
        ra, rb = a[key], b[key]
        for col in ra.keys() | rb.keys():
            va = (ra.get(col) or "").strip()
            vb = (rb.get(col) or "").strip()
            if va != vb:
                col_counter[col] += 1
                details.append(f"{key[0]} {key[1]} [{col}]\n    旧: {va[:160]}\n    新: {vb[:160]}")

    print()
    print(f"セル差分: {sum(col_counter.values())}箇所 / 差分のある行: "
          f"{len({d.split(' [', 1)[0] for d in details})}行")
    for col, n in col_counter.most_common():
        print(f"  {col}: {n}件")
    print()
    for d in details[: args.max_detail]:
        print(d)
    if len(details) > args.max_detail:
        print(f"…ほか {len(details) - args.max_detail} 箇所")
    return 0


if __name__ == "__main__":
    sys.exit(main())
