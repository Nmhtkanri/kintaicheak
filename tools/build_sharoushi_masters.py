# -*- coding: utf-8 -*-
r"""社労士モードの共有マスタ（列マッピング・追加支給台帳）を共有フォルダに作る。

    python tools/build_sharoushi_masters.py            # 無いものだけ作る
    python tools/build_sharoushi_masters.py --show     # 今の状態を見るだけ
    python tools/build_sharoushi_masters.py --force    # 既存をバックアップして作り直す

列マッピングは services/sharoushi_export.py の既定表がそのまま書き出される。
作ったあとは谷津さんが Excel で直せばよく、exe の再ビルドは要らない。
台帳は空（ヘッダと記入例のコメント行だけ）で作る。

背景・設計: docs/PLAN_社労士モード.md
"""

from __future__ import annotations

import argparse
import csv
import datetime
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import Config                                              # noqa: E402
from services.sharoushi_export import (                                # noqa: E402
    LEDGER_COLS,
    LEDGER_ITEM_COL_IDS,
    SharoushiExportError,
    export_default_mapping,
    load_column_mapping,
    load_extra_ledger,
)


def backup(path: str) -> str:
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    root, ext = os.path.splitext(path)
    dest = f"{root}_backup_{stamp}{ext}"
    shutil.copy2(path, dest)
    return dest


def build_ledger(path: str) -> str:
    """追加支給台帳をヘッダだけで作る（記入例は備考欄に文章で残す）。"""
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=LEDGER_COLS)
        w.writeheader()
        w.writerow({
            "支給月": "", "社員番号": "", "氏名": "", "項目": "", "金額": "",
            "メモ": "給与計算のあとに発生した支給をここに書く。項目に使えるのは "
                    + " / ".join(LEDGER_ITEM_COL_IDS)
                    + "。書いた額は差引支給額と立替金列に足され、口座1振込額と総支給額は変わらない。"
                    "この行（支給月と社員番号が空の行）は読み飛ばされる。",
        })
    return path


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--mapping", default=Config.SHAROUSHI_COLUMN_MAPPING_CSV,
                   help="列マッピングCSVの出力先 (default: %(default)s)")
    p.add_argument("--ledger", default=Config.SHAROUSHI_EXTRA_LEDGER_CSV,
                   help="追加支給台帳CSVの出力先 (default: %(default)s)")
    p.add_argument("--show", action="store_true", help="作らずに現状だけ表示する")
    p.add_argument("--force", action="store_true",
                   help="既存をバックアップしてから作り直す")
    args = p.parse_args()

    for label, path in (("列マッピング", args.mapping), ("追加支給台帳", args.ledger)):
        print(f"{label}: {path}")
        print(f"  存在: {'あり' if os.path.exists(path) else 'なし'}")

    if args.show:
        try:
            m = load_column_mapping(args.mapping)
            print(f"\n列マッピング: {m['rows_n']}行 "
                  f"(読み元: {m['path'] or 'コード内の既定表'})")
            ledger = load_extra_ledger(args.ledger)
            print(f"追加支給台帳: {len(ledger)}件")
        except SharoushiExportError as e:
            print(f"  読み込みエラー: {e}")
            return 2
        return 0

    made = []
    for label, path, builder in (
        ("列マッピング", args.mapping, lambda pth: export_default_mapping(pth, overwrite=True)),
        ("追加支給台帳", args.ledger, build_ledger),
    ):
        if os.path.exists(path):
            if not args.force:
                print(f"\n{label}: 既にあるので触りません（作り直すなら --force）")
                continue
            print(f"\n{label}: バックアップ → {backup(path)}")
        builder(path)
        made.append(path)
        print(f"{label}: 作成しました → {path}")

    if made:
        m = load_column_mapping(args.mapping)
        print(f"\n確認: 列マッピング {m['rows_n']}行を {m['path']} から読めました")
    return 0


if __name__ == "__main__":
    sys.exit(main())
