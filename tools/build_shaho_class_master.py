# -*- coding: utf-8 -*-
r"""標報チェックの報酬分類マスタ（初期版）を、給与明細の実データから作る。

    python tools/build_shaho_class_master.py                # 4〜6月のキャッシュから生成
    python tools/build_shaho_class_master.py --fetch        # 無い月をAPI取得してから生成
    python tools/build_shaho_class_master.py --show         # 現状を見るだけ
    python tools/build_shaho_class_master.py --force        # 既存をバックアップして作り直す

定時決定（4〜6月支給の報酬平均）の算定に、どの支給項目を入れるかを決めるマスタ。
実データに出てくる (source_key, 体系別名) を列挙し、シード分類を当てて CSV に書く。
**シードは出発点でしかない。「未設定」の行は谷津さんがレビューして確定する**
（未設定のまま金額が出た人は、チェック実行時に INSUFFICIENT_DATA へ落ちる）。

分類の根拠は note 列に残す。金額の出現統計（何名・いくら）も note に入れておくので、
レビュー時に「使われていない項目か」がその場で分かる。

背景・設計: docs/PLAN_標準報酬月額チェック.md
"""

from __future__ import annotations

import argparse
import csv
import datetime
import os
import shutil
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv                                        # noqa: E402

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
load_dotenv(os.path.join(REPO, ".env"))

from config import Config                                             # noqa: E402
from services.keiri_api import classify_employee, to_number           # noqa: E402
from services.keiri_engine import fetch_statements                    # noqa: E402
from services.shaho_master import CLASS_MASTER_COLS, load_class_master  # noqa: E402

CACHE_DIR = os.path.join(REPO, "outputs", "keiri")

# ---------------------------------------------------------------------------
# シード分類。(class, fixed, 根拠) を source_key または (source_key, 体系別名) で引く。
# fixed=1: 固定的賃金（単価・月額が決まっているもの。随時改定候補の検知に使う）
# ---------------------------------------------------------------------------
SEED_BY_KEY = {
    "salary_items:allowance1":  ("対象", "1", "基本給"),
    "salary_items:allowance10": ("対象", "1", "役職手当"),
    "salary_items:allowance11": ("対象", "1", "リーダー手当"),
    "salary_items:allowance12": ("対象", "1", "調整手当（2026-08支給分〜）"),
    "salary_items:allowance15": ("対象", "1", "調整手当（〜2026-07支給分。移設前）"),
    "salary_items:allowance16": ("対象", "1", "通信手当（毎月定額）"),
    "salary_items:allowance34": ("対象", "1", "課税通勤費（通勤手当は報酬）"),
    "salary_items:allowance35": ("対象", "1", "非課税通勤費（社会保険では非課税でも報酬）"),
    "salary_items:allowance13": ("対象", "0", "夜間当番手当（変動）"),
    "salary_items:allowance18": ("対象", "0", "テレワーク手当（変動）"),
    "salary_items:allowance19": ("対象", "0", "定常外業務対応手当（変動）"),
    "salary_items:allowance20": ("対象", "0", "その他手当（変動。着地先はテンプレート設定次第）"),
    "salary_items:allowance21": ("対象", "0", "その他手当（同上。2026-07はこちらに着地）"),
    "salary_items:allowance24": ("対象", "0", "差額調整（前月精算の実額はここに入る）"),
    "salary_items:allowance3":  ("対象外", "", "差額調整(a24)の内訳・再掲。入れると二重計上"),
    "salary_items:allowance4":  ("対象外", "", "差額調整(a24)の内訳・再掲。入れると二重計上"),
    "salary_items:allowance5":  ("対象外", "", "差額調整(a24)の内訳・再掲。入れると二重計上"),
    "salary_items:allowance6":  ("対象外", "", "差額調整(a24)の内訳・再掲。入れると二重計上"),
    "salary_items:allowance7":  ("対象外", "", "基礎時給（単価であって支給額ではない）"),
    "salary_items:allowance50": ("対象外", "", "立替金（顧客請求分）＝実費弁償"),
    "salary_items:allowance51": ("対象外", "", "立替金＝実費弁償"),
    "salary_items:allowance52": ("対象外", "", "その他＝経費精算（実費弁償）"),
    "salary_items:allowance53": ("未設定", "", "現物支給。恩恵的な物品なら対象外・報酬性があれば対象。要判断"
                                              "（実測: other5に不算入＝金銭給与の外側）"),
    "salary_items:allowance54": ("対象", "0", "支給過不足調整（実測: other5に算入＝実支給。2026-05/06 で30人月）"),
    "salary_items:allowance8":  ("対象外", "", "暫定支給額の再掲（2026-03/04のみ使用・総支給額に入らない情報項目。"
                                              "経理の最終CSVでも13名全員未計上。2026-04実測: other5に不算入）"),
}
# 体系別名（＋適用期間）で行を分けるもの。
# エントリ: (class, fixed, 適用開始月, 適用終了月, 根拠)。期間は両端含む・空=無期限。
# ⚠ 同じラベルでも月で意味が変わる実例があるため期間を持つ:
#   時給制の「みなし給」は 2026-05〜07 支給分では差額調整(a24)への再掲（総支給に乗らない）、
#   2026-08 支給分からは実額の支給（実測: 2026-05 守屋さん24,191円が other5 に不算入／
#   2026-08 は52名分が other5 に算入。経理モードの ZANTEI_LABELS_FROM と同じ事象）。
SEED_BY_LABEL = {
    ("salary_items:allowance2", "当月みなし時間外手当"):
        [("対象", "1", "", "", "みなし時間外手当（固定）")],
    ("salary_items:allowance2", "当月みなし深夜手当"):
        [("対象", "1", "", "", "みなし深夜手当（固定）")],
    ("salary_items:allowance2", "みなし手当"):
        [("対象", "1", "", "", "みなし手当（2026-04 暫定体系の固定みなし）")],
    ("salary_items:allowance2", "みなし給"):
        [("対象外", "", "", "2026-07", "時給制の再掲（実額はa24差額調整に折込。2026-05実測）"),
         ("対象", "1", "2026-08", "", "時給制のみなし給（2026-08支給分から実額）")],
    ("salary_items:allowance2", "前月超過勤務"):
        [("対象外", "", "", "", "時給制の前月実績の再掲（実額はa24に折込）")],
    # 2026-04（旧体系の最終月）だけ実支給だった精算項目。5月以降は a24 への再掲に変わる
    ("salary_items:allowance6", "過不足調整"):
        [("対象", "0", "", "2026-04", "旧体系の残業等精算（2026-04実測: other5に算入）")],
    ("salary_items:allowance6", "法定休日勤務分"):
        [("対象", "0", "", "2026-04", "旧体系の法定休日勤務（2026-04実測: other5に算入）")],
}

MONTHS = ("2026-04", "2026-05", "2026-06")


LABEL_SPLIT_KEYS = {"salary_items:allowance2", "salary_items:allowance6"}


def collect_items(months, fetch=False):
    """指定月の statements から salary_items の出現を集計する。

    Returns: ({(source_key, 体系別名グループ): {"label","n","total","months"}}, 取得できた月)
      体系別名グループ: ラベルで意味が分かれる項目（LABEL_SPLIT_KEYS）だけ体系別名、他は ""
    """
    stats = defaultdict(lambda: {"label": "", "n": 0, "total": 0.0, "months": set()})
    loaded = []
    for ym in months:
        path = os.path.join(CACHE_DIR, "raw", f"salary_statements_{ym}.json")
        if not os.path.exists(path) and not fetch:
            print(f"  {ym}: キャッシュなし（--fetch でAPI取得できます）")
            continue
        data = fetch_statements(CACHE_DIR, ym)
        loaded.append(ym)
        print(f"  {ym}: {len(data)}人分")
        for person in data:
            emp = str(person.get("employee_id", "")).strip()
            if classify_employee(emp) != "target":
                continue
            for st in person.get("statements") or []:
                pi = st.get("payroll_info") or {}
                for it in pi.get("salary_items") or []:
                    key = "salary_items:%s" % it.get("id")
                    label = (str(it.get("salary_system_label") or "").strip()
                             or str(it.get("label") or "").strip())
                    group = label if key in LABEL_SPLIT_KEYS else ""
                    rec = stats[(key, group)]
                    if label and not rec["label"]:
                        rec["label"] = label
                    v = to_number(it.get("value"))
                    if v:
                        rec["n"] += 1
                        rec["total"] += v
                        rec["months"].add(ym)
    return stats, loaded


def seed_for(key, group):
    """(key, グループ) のシード行リスト。期間つきエントリはそのまま複数行になる。"""
    if group:
        seeded = SEED_BY_LABEL.get((key, group))
        if seeded:
            return seeded
    if key in SEED_BY_KEY:
        cls, fixed, reason = SEED_BY_KEY[key]
        return [(cls, fixed, "", "", reason)]
    return [("未設定", "", "", "", "実データに出現。分類を決めてください")]


def build_rows(stats):
    rows = []
    for (key, group), rec in sorted(stats.items()):
        usage = ("未使用（全期間ゼロ）" if rec["n"] == 0 else
                 "実績 %d人月 %s円" % (rec["n"], format(int(rec["total"]), ",")))
        for cls, fixed, ym_from, ym_to, reason in seed_for(key, group):
            rows.append({
                "source_key": key,
                "salary_system_label": group,
                "label": rec["label"],
                "class": cls,
                "fixed": fixed,
                "適用開始月": ym_from,
                "適用終了月": ym_to,
                "note": f"{reason}／{usage}",
            })
    return rows


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--out", default=Config.SHAHO_CLASS_MASTER_CSV,
                   help="出力先CSV (default: %(default)s)")
    p.add_argument("--months", nargs="*", default=list(MONTHS),
                   help="集計する支給月 (default: %(default)s)")
    p.add_argument("--fetch", action="store_true",
                   help="キャッシュに無い月を jinjer API から取得する（数分かかる）")
    p.add_argument("--show", action="store_true", help="生成せず現状を表示する")
    p.add_argument("--force", action="store_true",
                   help="既存マスタをバックアップしてから作り直す")
    args = p.parse_args()

    print(f"報酬分類マスタ: {args.out}")
    print(f"  存在: {'あり' if os.path.exists(args.out) else 'なし'}")
    if args.show:
        if os.path.exists(args.out):
            m = load_class_master(args.out)
            by_cls = defaultdict(int)
            for r in m["rules"]:
                by_cls[r.cls] += 1
            print(f"  行数: {len(m['rules'])} / 内訳: {dict(by_cls)}")
            for r in m["rules"]:
                if r.cls == "未設定":
                    print(f"  未設定: {r.source_key} [{r.label}] {r.note}")
        return 0

    if os.path.exists(args.out) and not args.force:
        print("既にあるので触りません（作り直すなら --force。レビュー済みの分類を潰さないため）")
        return 1

    print("集計中:")
    stats, loaded = collect_items(args.months, fetch=args.fetch)
    if not loaded:
        print("1か月も読めませんでした。--fetch を付けるか、経理モードでキャッシュを作ってください")
        return 1
    if len(loaded) < len(args.months):
        print(f"⚠️ {len(loaded)}/{len(args.months)}か月分だけで生成します"
              f"（不足月を取得したら --force で作り直してください）")

    rows = build_rows(stats)
    if os.path.exists(args.out):
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        root, ext = os.path.splitext(args.out)
        backup = f"{root}_backup_{stamp}{ext}"
        shutil.copy2(args.out, backup)
        print(f"バックアップ → {backup}")
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CLASS_MASTER_COLS)
        w.writeheader()
        for r in rows:
            w.writerow(r)
    by_cls = defaultdict(int)
    for r in rows:
        by_cls[r["class"]] += 1
    print(f"生成しました: {len(rows)}行 / 内訳: {dict(by_cls)}")
    print("検証読み込み…")
    m = load_class_master(args.out)
    print(f"OK（{len(m['rules'])}行）。「未設定」{by_cls.get('未設定', 0)}行は"
          "谷津さんのレビューで確定してください")
    return 0


if __name__ == "__main__":
    sys.exit(main())
