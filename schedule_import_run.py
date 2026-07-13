# -*- coding: utf-8 -*-
"""schedule_import_run — グリッドCSVを jinjer API でスケジュール投入するCLI

スケジュールアップロードモードの後工程（Web UIと同じ処理のコマンド版）。
作成済みのグリッドCSV（当アプリの exporter 出力、または Z:\\jinjer移行\\カレンダー
配下の同形式ファイル）を、画面を通さずに差分投入できる。

使い方:
    python schedule_import_run.py <グリッドCSV|フォルダ>... --month 2026-08              # dry-run
    python schedule_import_run.py <グリッド...> --month 2026-08 --execute --expect 123   # 投入＋検証
    python schedule_import_run.py <グリッド...> --month 2026-08 --exclude 2009006,2024047

背景・設計: docs/PLAN_スケジュールAPI投入.md
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()  # noqa: E402 — Config 参照前に .env を読む

from config import Config  # noqa: E402
from services.schedule_import_runner import run_schedule_api_import  # noqa: E402


def _expand_grid_paths(args_paths: list[str]) -> list[Path]:
    """引数のファイル/フォルダを展開する（フォルダは直下の *.csv）"""
    out: list[Path] = []
    for a in args_paths:
        p = Path(a)
        if p.is_dir():
            out.extend(sorted(p.glob("*.csv")))
        else:
            out.append(p)
    return out


def main() -> int:
    parser = argparse.ArgumentParser(
        description="グリッドCSVを jinjer API でスケジュール投入する（差分のみ）")
    parser.add_argument("grid_csvs", nargs="+",
                        help="グリッドCSVファイル、またはフォルダ（直下の*.csvを展開）")
    parser.add_argument("--month", required=True,
                        help="対象月 YYYY-MM（グリッドのヘッダー年月と一致必須。未来月OK）")
    parser.add_argument("--execute", action="store_true",
                        help="実際に投入する（未指定なら dry-run）")
    parser.add_argument("--expect", type=int, default=None,
                        help="書込行数ガード（--execute 時は必須。dry-run の書込行数を指定）")
    parser.add_argument("--fingerprint", default="",
                        help="追加ガード: dry-run が出力した プランfingerprint")
    parser.add_argument("--exclude", default="",
                        help="除外する従業員ID（カンマ区切り）")
    parser.add_argument("--executor", default=Config.JINJER_IMPORT_EXECUTOR_ID,
                        help="実行者の社員番号（完了通知メール宛先。既定: %(default)s）")
    parser.add_argument("--output-dir", default=Config.OUTPUT_FOLDER,
                        help="レポート出力先（既定: %(default)s）")
    parser.add_argument("--template-csv", default="",
                        help="雛形一覧CSV（既定: 共有フォルダの最新を自動解決）")
    args = parser.parse_args()

    if args.execute and args.expect is None:
        print("--execute には --expect <書込行数> が必須です"
              "（dry-run の「書込 N行」を確認して指定してください）")
        return 3

    grid_paths = _expand_grid_paths(args.grid_csvs)
    if not grid_paths:
        print("グリッドCSVが指定されていません")
        return 3
    missing = [p for p in grid_paths if not p.is_file()]
    if missing:
        for p in missing:
            print(f"見つかりません: {p}")
        return 3

    exclude = {e.strip() for e in args.exclude.split(",") if e.strip()}
    result = run_schedule_api_import(
        grid_csvs=grid_paths,
        output_dir=Path(args.output_dir),
        month=args.month,
        executor_id=args.executor,
        dry_run=not args.execute,
        expected_fingerprint=args.fingerprint,
        expected_rows=args.expect,
        exclude_emps=exclude or None,
        template_csv=args.template_csv,
    )
    if result.dry_run:
        if not result.ok:
            print(f"\n※ 送信前チェックNGで中止しました。レポートの「要手動確認」シートを"
                  f"確認してください（{result.report_path}）")
            return 3
        print(f"\n※ dry-run でした。投入するには --execute --expect {result.plan_rows} "
              f"を付けて再実行してください（レポート: {result.report_path}）")
        return 0
    if not result.ok:
        print("\n※ NG があります。レポートの「要手動確認」シートを確認してください。")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
