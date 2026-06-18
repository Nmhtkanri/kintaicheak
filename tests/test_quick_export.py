import csv
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from quick_export import run_quick_export  # noqa: E402


HEADERS = [
    "名前",
    "*従業員ID",
    "*年月日",
    "出勤1",
    "退勤1",
    "休憩1",
    "復帰1",
    "休憩時間",
    "実績確定状況",
]


def _write_jinjer_csv(path):
    with open(path, "w", encoding="cp932", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        writer.writerow(["上原 奏吾", "2018057", "2026/4/1", "9:00", "18:00", "", "", "0:00", "TRUE"])


def _read_output_row(path):
    with open(path, encoding="cp932", newline="") as f:
        rows = list(csv.DictReader(f))
    return rows[0]


def test_run_quick_export_writes_manual_break_columns(tmp_path):
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1,
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "差異種別": "休憩",
        "自動修正提案値": "",
        "手入力休憩1": "12:00",
        "手入力復帰1": "13:00",
        "手入力休憩時間": "1:00",
        "人間判断": "承認",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    row = _read_output_row(output_path)
    assert row["休憩1"] == "12:00"
    assert row["復帰1"] == "13:00"
    assert row["休憩時間"] == "1:00"
    assert result.stats.overwritten_break_start == 1
    assert result.stats.overwritten_break_end == 1
    assert result.stats.overwritten_break_total == 1
    assert result.stats.skipped_break == 0


def test_run_quick_export_skips_approved_break_without_manual_values(tmp_path):
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1,
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "差異種別": "休憩",
        "自動修正提案値": "",
        "人間判断": "承認",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    row = _read_output_row(output_path)
    assert row["休憩1"] == ""
    assert row["復帰1"] == ""
    assert row["休憩時間"] == "0:00"
    assert result.stats.skipped_break == 1


def test_run_quick_export_prefers_manual_fix_value_for_punch(tmp_path):
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1,
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "差異種別": "退勤",
        "自動修正提案値": "18:10",
        "打刻修正": "18:15",
        "人間判断": "承認",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert _read_output_row(output_path)["退勤1"] == "18:15"


def test_run_quick_export_reads_legacy_tenyuryoku_column(tmp_path):
    """旧フォーマット（列名「手入力修正値」）の差異一覧も後方互換で読める。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1,
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "差異種別": "退勤",
        "自動修正提案値": "18:10",
        "手入力修正値": "18:15",  # 旧列名（改名前の記入済みファイル）
        "人間判断": "承認",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert _read_output_row(output_path)["退勤1"] == "18:15"


def test_run_quick_export_recovers_judgment_in_wrong_column(tmp_path):
    """承認を「人間判断」列ではなく「手入力修正値」列に入力してしまったケース。

    判断は回収され、誤入力列の『承認』は時刻として書き込まれず、
    自動修正提案値が正しく反映される。
    """
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1,
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "差異種別": "出勤",
        "自動修正提案値": "08:30",
        "手入力修正値": "承認",   # ← 本来は「人間判断」列に入れるべき
        "人間判断": "",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert result.stats.approved == 1
    assert result.stats.recovered_misplaced == 1
    assert result.stats.overwritten_punch_in == 1
    row = _read_output_row(output_path)
    # 「承認」が時刻として書き込まれず、提案値 08:30 が出勤1へ反映される
    assert row["出勤1"] == "08:30"
    assert row["出勤1"] != "承認"


def _write_jinjer_csv_overnight(path):
    with open(path, "w", encoding="cp932", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        # 夜勤: 出勤 21:00 あり、退勤は空（請求勤怠で補う）
        writer.writerow(["上原 奏吾", "2018057", "2026/4/1", "21:00", "", "", "", "0:00", "TRUE"])


def test_run_quick_export_converts_overnight_punch_out(tmp_path):
    """夜勤の退勤提案値（32:15）がそのまま退勤1へ入る。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv_overnight(jinjer_path)

    pd.DataFrame([{
        "行ID": 1,
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "差異種別": "退勤",
        "自動修正提案値": "32:15",
        "人間判断": "承認",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert _read_output_row(output_path)["退勤1"] == "32:15"


def test_run_quick_export_fixes_overnight_manual_punch_out(tmp_path):
    """手入力で 08:15 と入れても、出勤1=21:00 と突き合わせて 32:15 へ補正される。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv_overnight(jinjer_path)

    pd.DataFrame([{
        "行ID": 1,
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "差異種別": "退勤",
        "自動修正提案値": "",
        "打刻修正": "08:15",
        "人間判断": "承認",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert _read_output_row(output_path)["退勤1"] == "32:15"


def test_run_quick_export_keeps_normal_punch_out(tmp_path):
    """通常勤務（出勤 9:00 / 退勤 18:00）は 24時超変換しない。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1,
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "差異種別": "退勤",
        "自動修正提案値": "18:30",
        "人間判断": "承認",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert _read_output_row(output_path)["退勤1"] == "18:30"


def test_run_quick_export_ignores_judgment_keyword_as_time(tmp_path):
    """判断キーワードが入力欄に紛れても、時刻として書き込まれない。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1,
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "差異種別": "退勤",
        "自動修正提案値": "18:10",
        "打刻修正": "承認",  # 誤入力。退勤1 には 18:10 が入るべき
        "人間判断": "承認",      # 判断は正しい列にもある
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert _read_output_row(output_path)["退勤1"] == "18:10"


def test_run_quick_export_label_seikyuukintai_overwrites(tmp_path):
    """新ラベル「請求勤怠」=請求勤怠を正 → jinjer へ書き戻す。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1, "従業員ID": "2018057", "氏名": "上原 奏吾",
        "対象日付": "2026-04-01", "差異種別": "出勤",
        "自動修正提案値": "8:30", "人間判断": "請求勤怠",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert _read_output_row(output_path)["出勤1"] == "8:30"
    assert result.stats.approved == 1


def test_run_quick_export_label_jinjer_keeps_jinjer(tmp_path):
    """新ラベル「jinjer勤怠」=jinjerを正 → 書き戻さない。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1, "従業員ID": "2018057", "氏名": "上原 奏吾",
        "対象日付": "2026-04-01", "差異種別": "出勤",
        "自動修正提案値": "8:30", "人間判断": "jinjer勤怠",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert _read_output_row(output_path)["出勤1"] == "9:00"  # 書き戻さない
    assert result.stats.approved == 0
    assert result.stats.rejected == 1
