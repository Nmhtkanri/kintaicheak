import csv
import os
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.batch_runner import run_batch_compare, gather_timesheet_files  # noqa: E402


def _write_estaffing_csv(path):
    """e-staffing形式の請求勤怠（AI不要で直接パースされる形式）。"""
    headers = [
        "スタッフ氏名", "就業年月日", "開始時刻", "終了時刻", "休憩時間", "備考コメント",
    ]
    with open(path, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerow(["山田 太郎", "2026/5/1", "9:00", "18:00", "1:00", ""])
        w.writerow(["山田 太郎", "2026/5/2", "9:00", "18:00", "1:00", ""])


def _write_jinjer_generic_csv(path):
    """jinjer 汎用データ（打刻ソース＆集計列）。最小限の列。"""
    headers = [
        "名前", "*従業員ID", "*年月日", "出勤1", "退勤1", "休憩1", "復帰1",
        "休憩時間", "総労働時間", "実労働時間", "実績確定状況",
    ]
    with open(path, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        # 5/1 は出勤が 9:30（請求勤怠 9:00 と 30分差）→ 出勤差異が出る
        w.writerow(["山田 太郎", "1001", "2026/5/1", "9:30", "18:00", "", "", "1:00", "8:30", "7:30", "FALSE"])
        w.writerow(["山田 太郎", "1001", "2026/5/2", "9:00", "18:00", "", "", "1:00", "9:00", "8:00", "FALSE"])


def test_run_batch_compare_end_to_end(tmp_path):
    ts_dir = tmp_path / "請求勤怠"
    ts_dir.mkdir()
    _write_estaffing_csv(ts_dir / "estaffing.csv")
    jinjer = tmp_path / "汎用データ.csv"
    _write_jinjer_generic_csv(jinjer)
    out = tmp_path / "差異一覧.xlsx"

    result, skipped, unsubmitted = run_batch_compare(
        timesheet_dir=ts_dir,
        jinjer_dir=jinjer,
        output_path=out,
        month_label="2026-05",
        application_csv=None,
        log_func=lambda _m: None,
    )

    assert result.ok, result.error
    assert out.exists()
    # 差異一覧シートに出勤差異が出る
    df = pd.read_excel(out, sheet_name="差異一覧", dtype=object)
    assert "確認区分" in df.columns
    assert (df["差異種別"] == "出勤").any()
    # 未処理・未マッチシートが追記されている
    xl = pd.ExcelFile(out)
    assert "未処理・未マッチ" in xl.sheet_names
    # e-staffing CSV は直接パースされ、解析スキップは無い
    assert skipped == []


def test_gather_timesheet_files_skips_temp_and_aggregate_dirs(tmp_path):
    (tmp_path / "通常").mkdir()
    (tmp_path / "通常" / "a.xlsx").write_text("x")
    (tmp_path / "通常" / "~$a.xlsx").write_text("x")        # 一時ファイル→除外
    (tmp_path / "月の総労働時間").mkdir()
    (tmp_path / "月の総労働時間" / "b.xlsx").write_text("x")  # 集計フォルダ→除外
    (tmp_path / "通常" / "note.docx").write_text("x")        # 対象外拡張子→除外

    files = [p.name for p in gather_timesheet_files(tmp_path)]
    assert files == ["a.xlsx"]


def test_run_batch_compare_no_timesheet_returns_error(tmp_path):
    ts_dir = tmp_path / "空"
    ts_dir.mkdir()
    jinjer = tmp_path / "汎用データ.csv"
    _write_jinjer_generic_csv(jinjer)
    out = tmp_path / "差異一覧.xlsx"

    result, skipped, unsubmitted = run_batch_compare(
        timesheet_dir=ts_dir, jinjer_dir=jinjer, output_path=out,
        month_label="2026-05", log_func=lambda _m: None,
    )
    assert not result.ok
    assert "請求勤怠" in result.error
