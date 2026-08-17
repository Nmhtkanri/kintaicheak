# -*- coding: utf-8 -*-
r"""標準報酬月額チェック（定時決定・保険料突合）のCLI。jinjer へは書き込まない。

    python shaho_check_run.py                       # 算定年は自動（7月以降=当年）
    python shaho_check_run.py --year 2026 --check-month 2026-07
    python shaho_check_run.py --insurer kyokai_tokyo

4〜6月支給の報酬から9月適用予定の標準報酬月額を計算し、jinjer の登録値・控除実績と
突合して Excel＋JSON を outputs/shaho/{年}/ に出す。給与明細は経理モードと共用の
キャッシュ（outputs/keiri/raw/）を読む。**過去月のキャッシュは上書きしない**
（basic_info は取得時点の値しか返らず、上書きすると期中改定の検知が壊れるため）。

終了コード: 0=要確認ゼロ / 2=要確認あり（レポートは出力済み） / 1=実行エラー。
背景・設計: docs/PLAN_標準報酬月額チェック.md
"""

from __future__ import annotations

import argparse
import datetime
import os
import sys

from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

from config import Config                                             # noqa: E402
from services.shaho_check import REVIEW_STATUSES, STATUS_JA, run_check  # noqa: E402
from services.shaho_master import ShahoMasterError                    # noqa: E402
from services.shaho_report import write_reports                       # noqa: E402


def default_year(today=None) -> int:
    t = today or datetime.date.today()
    return t.year if t.month >= 7 else t.year - 1


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--year", type=int, default=default_year(),
                   help="算定年（YYYY年4〜6月支給 → 9月適用。default: %(default)s）")
    p.add_argument("--check-month", default=None,
                   help="控除突合に使う支給月 YYYY-MM（default: キャッシュにある最新月）")
    p.add_argument("--insurer", choices=["its", "kyokai_tokyo"], default=Config.SHAHO_INSURER,
                   help="保険者 (default: %(default)s)")
    p.add_argument("--output-dir", default=Config.SHAHO_OUTPUT_DIR,
                   help="出力先 (default: %(default)s)")
    args = p.parse_args()

    check_month = args.check_month
    if not check_month:
        raw = os.path.join(Config.KEIRI_OUTPUT_DIR, "raw")
        cached = sorted(f[len("salary_statements_"):-len(".json")]
                        for f in os.listdir(raw) if f.startswith("salary_statements_"))
        if not cached:
            print("給与明細のキャッシュがありません（経理モードで取得してください）")
            return 1
        check_month = cached[-1]
        print(f"突合月: {check_month}（キャッシュの最新月）")

    try:
        check = run_check(args.year, check_month, insurer=args.insurer,
                          out_base=args.output_dir)
        out = write_reports(check)
    except ShahoMasterError as e:
        print(f"[NG] {e}")
        return 1

    counts = {}
    for r in check["results"]:
        counts[r.total_status] = counts.get(r.total_status, 0) + 1
    print(f"対象 {out['n']}名 / 要確認 {out['review_n']}名")
    for st in STATUS_JA:
        if counts.get(st):
            mark = "⚠" if st in REVIEW_STATUSES else " "
            print(f" {mark} {STATUS_JA[st]:<12} {counts[st]:>4}名")
    print(f"Excel: {out['xlsx']}")
    print(f"JSON : {out['json']}")
    return 2 if out["review_n"] else 0


if __name__ == "__main__":
    sys.exit(main())
