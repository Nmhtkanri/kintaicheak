# -*- coding: utf-8 -*-
"""api_import_run — アップロード用CSVを jinjer API で投入するCLI（手順③の後工程）

使い方:
    python api_import_run.py <アップロード用CSV>                # dry-run（ガードと件数のみ）
    python api_import_run.py <アップロード用CSV> --execute      # 投入＋検証＋Excelレポート
    python api_import_run.py <CSV> --execute --executor 2026007 --month 2026-06

背景・設計: docs/PLAN_手順3_API直接投入.md
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()  # noqa: E402 — Config 参照前に .env を読む

from config import Config  # noqa: E402
from services.kintai_import_runner import run_api_import  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="アップロード用CSVを jinjer API で投入する")
    parser.add_argument("upload_csv", help="quick_export が生成した汎用データ形式CSV (CP932)")
    parser.add_argument("--execute", action="store_true",
                        help="実際に投入する（未指定なら dry-run）")
    parser.add_argument("--executor", default=Config.JINJER_IMPORT_EXECUTOR_ID,
                        help="実行者の社員番号（完了通知メール宛先。既定: %(default)s）")
    parser.add_argument("--month", default="", help="検証対象月 YYYY-MM（省略時はCSVから判定）")
    parser.add_argument("--output-dir", default=Config.OUTPUT_FOLDER,
                        help="レポート出力先（既定: %(default)s）")
    args = parser.parse_args()

    upload_csv = Path(args.upload_csv)
    if not upload_csv.is_file():
        print(f"CSVが見つかりません: {upload_csv}")
        return 1

    result = run_api_import(
        upload_csv=upload_csv,
        output_dir=Path(args.output_dir),
        executor_id=args.executor,
        dry_run=not args.execute,
        month=args.month,
    )
    if result.dry_run:
        print("\n※ dry-run でした。投入するには --execute を付けて再実行してください。")
        return 0
    if not result.ok:
        print("\n※ NG行があります。レポートの「手動対応リスト」を確認してください。")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
