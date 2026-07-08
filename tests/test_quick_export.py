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
    assert result.stats.skipped_break == 1
    # 何も書き込まれなかった日は出力しない（変更行のみ出力）
    with open(output_path, encoding="cp932", newline="") as f:
        assert list(csv.DictReader(f)) == []


def test_run_quick_export_applies_manual_break_on_non_break_rows(tmp_path):
    """休憩行が無くても、出勤/退勤/総労働時間の行に入力した手入力休憩を反映する。

    菅原孝さん 2026-06 の実例: 差異一覧に「休憩」種別の行が1件も無く、
    手入力休憩1/復帰1 を出勤・退勤・総労働時間の行に入力したところ、
    旧実装では全行無視され休憩1/復帰1 が空のままだった。
    """
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    common = {
        "従業員ID": "2018057",
        "氏名": "上原 奏吾",
        "対象日付": "2026-04-01",
        "手入力休憩1": "12:00",
        "手入力復帰1": "13:30",
    }
    pd.DataFrame([
        {"行ID": 1, "差異種別": "出勤", "自動修正提案値": "08:30", "人間判断": "請求勤怠", **common},
        {"行ID": 2, "差異種別": "退勤", "自動修正提案値": "19:45", "人間判断": "請求勤怠", **common},
        # 総労働時間を jinjer勤怠（却下）にしても、手入力休憩は反映される
        {"行ID": 3, "差異種別": "総労働時間", "自動修正提案値": "", "人間判断": "jinjer勤怠", **common},
    ]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    row = _read_output_row(output_path)
    assert row["休憩1"] == "12:00"
    assert row["復帰1"] == "13:30"
    assert row["出勤1"] == "08:30"
    assert row["退勤1"] == "19:45"
    # 同一 (従業員, 日付) なので反映は1回だけ
    assert result.stats.manual_break_days == 1
    assert result.stats.overwritten_break_start == 1
    assert result.stats.overwritten_break_end == 1
    assert result.stats.manual_break_conflicts == 0
    assert result.stats.skipped_break == 0


def test_run_quick_export_warns_on_conflicting_manual_breaks(tmp_path):
    """同じ日に食い違う手入力休憩が入力された場合、先の値を採用して警告する。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([
        {"行ID": 1, "従業員ID": "2018057", "氏名": "上原 奏吾", "対象日付": "2026-04-01",
         "差異種別": "出勤", "自動修正提案値": "08:30", "人間判断": "請求勤怠",
         "手入力休憩1": "12:00", "手入力復帰1": "13:00"},
        {"行ID": 2, "従業員ID": "2018057", "氏名": "上原 奏吾", "対象日付": "2026-04-01",
         "差異種別": "退勤", "自動修正提案値": "19:45", "人間判断": "請求勤怠",
         "手入力休憩1": "12:00", "手入力復帰1": "13:30"},  # 復帰が食い違う
    ]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    row = _read_output_row(output_path)
    assert row["休憩1"] == "12:00"
    assert row["復帰1"] == "13:00"  # 先に読んだ行の値を採用
    assert result.stats.manual_break_conflicts == 1
    assert any("食い違" in w for w in result.stats.warnings)


def test_run_quick_export_holds_manual_break_on_hold_row(tmp_path):
    """人間判断=保留 の行に入力された手入力休憩は反映しない。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)

    pd.DataFrame([{
        "行ID": 1, "従業員ID": "2018057", "氏名": "上原 奏吾", "対象日付": "2026-04-01",
        "差異種別": "総労働時間", "自動修正提案値": "", "人間判断": "保留",
        "手入力休憩1": "12:00", "手入力復帰1": "13:30",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert result.stats.manual_break_days == 0
    assert any("保留のため反映しません" in w for w in result.stats.warnings)
    # 保留の日は行ごとアップロードCSVから除外される（jinjer手修正の保護）
    assert result.stats.held_rows_removed == 1
    with open(output_path, encoding="cp932", newline="") as f:
        rows = list(csv.DictReader(f))
    assert rows == []


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
    assert result.stats.approved == 0
    assert result.stats.rejected == 1
    # 書き戻さない＝無変更の日なので、行自体を出力しない（jinjerに触らせない）
    with open(output_path, encoding="cp932", newline="") as f:
        assert list(csv.DictReader(f)) == []


# =============================================================================
# 出勤採用時にスケジュール開始（出勤予定時刻）も合わせる
# =============================================================================

def test_apply_approved_rows_syncs_schedule_start():
    from quick_export import (
        apply_approved_rows, build_jinjer_row_index, ApprovedRow, Stats,
        DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT,
    )

    headers = ["*従業員ID", "*年月日", "出勤予定時刻", "退勤予定時刻",
               "出勤1", "退勤1", "休憩1", "復帰1", "休憩時間", "実績確定状況"]
    rows = [["2020001", "2026/6/2", "9:00", "17:30", "8:50", "17:40", "", "", "", "FALSE"]]
    idx = build_jinjer_row_index(headers, rows)
    stats = Stats()
    approved = [
        ApprovedRow(emp_id="2020001", target_date_iso="2026-06-02", kind=DIFF_KIND_PUNCH_IN,
                    auto_fix_value="8:30", manual_fix_value="", manual_break_start="",
                    manual_break_end="", manual_break_total="", name="テスト",
                    warn_level="", source_diff_row_id=1),
        ApprovedRow(emp_id="2020001", target_date_iso="2026-06-02", kind=DIFF_KIND_PUNCH_OUT,
                    auto_fix_value="17:40", manual_fix_value="", manual_break_start="",
                    manual_break_end="", manual_break_total="", name="テスト",
                    warn_level="", source_diff_row_id=2),
    ]
    apply_approved_rows(headers, rows, idx, approved, stats)

    # 出勤1 と 出勤予定時刻 の両方が採用値 8:30 に揃う
    assert rows[0][headers.index("出勤1")] == "8:30"
    assert rows[0][headers.index("出勤予定時刻")] == "8:30"
    # 退勤側は予定を触らない
    assert rows[0][headers.index("退勤予定時刻")] == "17:30"
    assert stats.overwritten_punch_in == 1
    assert stats.overwritten_sched_in == 1


def test_apply_approved_rows_blank_adoption_keeps_schedule():
    """採用値が空（jinjer打刻を消すケース）のときは出勤予定時刻を消さない。"""
    from quick_export import (
        apply_approved_rows, build_jinjer_row_index, ApprovedRow, Stats,
        DIFF_KIND_PUNCH_IN,
    )
    headers = ["*従業員ID", "*年月日", "出勤予定時刻", "出勤1", "退勤1"]
    rows = [["2020002", "2026/6/3", "10:00", "9:00", "17:00"]]
    idx = build_jinjer_row_index(headers, rows)
    stats = Stats()
    approved = [
        ApprovedRow(emp_id="2020002", target_date_iso="2026-06-03", kind=DIFF_KIND_PUNCH_IN,
                    auto_fix_value="", manual_fix_value="", manual_break_start="",
                    manual_break_end="", manual_break_total="", name="空採用",
                    warn_level="", source_diff_row_id=3),
    ]
    apply_approved_rows(headers, rows, idx, approved, stats)
    assert rows[0][headers.index("出勤1")] == ""
    assert rows[0][headers.index("出勤予定時刻")] == "10:00"  # 予定は維持
    assert stats.overwritten_sched_in == 0


def test_run_quick_export_punch_uses_seikyu_value_when_proposal_is_label(tmp_path):
    """新フォーマット（自動修正提案値=採用ラベル）で、請求勤怠を採用した退勤を
    請求勤怠値の時刻で書き戻す。太田さん5/7の『退勤1が空で上書き＝打刻消失』の回帰防止。"""
    import csv as _csv
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    # 夜勤: 出勤20:45 / jinjer退勤09:01
    with open(jinjer_path, "w", encoding="cp932", newline="") as f:
        w = _csv.writer(f)
        w.writerow(HEADERS)
        w.writerow(["太田 太郎", "2020001", "2026/5/7", "20:45", "09:01", "", "", "0:00", "FALSE"])

    pd.DataFrame([{
        "行ID": 1, "従業員ID": "2020001", "氏名": "太田 太郎",
        "対象日付": "2026-05-07", "差異種別": "退勤",
        "請求勤怠値": "09:00",          # 表示している請求勤怠の退勤時刻（夜勤の翌朝）
        "自動修正提案値": "請求勤怠",     # 新フォーマット: 採用ラベル（時刻ではない）
        "打刻修正": "",                  # 手入力なし
        "人間判断": "請求勤怠",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)
    assert result.ok
    row = _read_output_row(output_path)
    # 空ではなく、請求勤怠値09:00が夜勤の24時超表記33:00に補正されて書き戻る
    assert row["退勤1"] == "33:00"
    assert result.stats.overwritten_punch_out == 1


def test_run_quick_export_reverse_missing_punch_clears_when_seikyu_blank(tmp_path):
    """逆向き片側欠落（請求勤怠値が空）を請求勤怠で承認したら、jinjer打刻を空で消す（意図どおり）。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv(jinjer_path)  # 上原 2018057 4/1 出勤9:00 退勤18:00
    pd.DataFrame([{
        "行ID": 1, "従業員ID": "2018057", "氏名": "上原 奏吾",
        "対象日付": "2026-04-01", "差異種別": "退勤",
        "請求勤怠値": "",               # 請求勤怠側に打刻なし → 消す
        "自動修正提案値": "保留",
        "人間判断": "請求勤怠",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)
    assert result.ok
    assert _read_output_row(output_path)["退勤1"] == ""


def _write_jinjer_csv_3days(path):
    with open(path, "w", encoding="cp932", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        writer.writerow(["上原 奏吾", "2018057", "2026/4/1", "9:00", "18:00", "", "", "0:00", "TRUE"])
        writer.writerow(["上原 奏吾", "2018057", "2026/4/2", "16:45", "33:33", "", "", "1:00", "TRUE"])
        writer.writerow(["上原 奏吾", "2018057", "2026/4/3", "9:00", "18:00", "", "", "0:00", "TRUE"])


def _read_output_rows(path):
    with open(path, encoding="cp932", newline="") as f:
        return list(csv.DictReader(f))


def test_run_quick_export_removes_held_day_rows(tmp_path):
    """人間判断=保留 の日は行ごとアップロードCSVから除外される。

    保留=「この日はツールで触らない」。DL後にjinjer画面で手修正した日を保留にする
    運用のため、DL時点の古い値を再インポートして手修正を巻き戻さない。
    """
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv_3days(jinjer_path)

    pd.DataFrame([
        {   # 4/2 は保留（jinjer手修正済みの想定）
            "行ID": 1, "従業員ID": "2018057", "氏名": "上原 奏吾",
            "対象日付": "2026-04-02", "差異種別": "退勤",
            "請求勤怠値": "33:30", "自動修正提案値": "", "人間判断": "保留",
        },
        {   # 4/3 は承認 → 通常どおり上書きして出力
            "行ID": 2, "従業員ID": "2018057", "氏名": "上原 奏吾",
            "対象日付": "2026-04-03", "差異種別": "退勤",
            "請求勤怠値": "18:15", "自動修正提案値": "18:15", "人間判断": "請求勤怠",
        },
    ]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert result.stats.held_rows_removed == 1
    rows = _read_output_rows(output_path)
    dates = [r["*年月日"] for r in rows]
    assert "2026/4/2" not in dates          # 保留日は行ごと消えている
    assert "2026/4/1" not in dates          # 無変更日も出力しない（変更行のみ出力）
    assert len(rows) == 1                    # 4/3(承認)のみ
    out_43 = rows[0]
    assert out_43["退勤1"] == "18:15"


def test_run_quick_export_keeps_row_when_hold_and_approve_mixed(tmp_path):
    """同じ日に保留と承認が混在する場合は行を残して警告を出す（承認の書き戻しを優先）。"""
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv_3days(jinjer_path)

    pd.DataFrame([
        {
            "行ID": 1, "従業員ID": "2018057", "氏名": "上原 奏吾",
            "対象日付": "2026-04-01", "差異種別": "出勤",
            "請求勤怠値": "8:55", "自動修正提案値": "8:55", "人間判断": "請求勤怠",
        },
        {
            "行ID": 2, "従業員ID": "2018057", "氏名": "上原 奏吾",
            "対象日付": "2026-04-01", "差異種別": "退勤",
            "請求勤怠値": "18:10", "自動修正提案値": "", "人間判断": "保留",
        },
    ]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    assert result.stats.held_rows_removed == 0
    rows = _read_output_rows(output_path)
    dates = [r["*年月日"] for r in rows]
    assert "2026/4/1" in dates  # 承認があるので行は残る
    out_41 = next(r for r in rows if r["*年月日"] == "2026/4/1")
    assert out_41["出勤1"] == "8:55"   # 承認分は反映
    assert any("混在" in w for w in result.stats.warnings)


def _make_approved(emp, date_iso, kind, value, row_id=1, name="テスト"):
    from quick_export import ApprovedRow
    return ApprovedRow(
        emp_id=emp, target_date_iso=date_iso, kind=kind,
        auto_fix_value=value, manual_fix_value="", manual_break_start="",
        manual_break_end="", manual_break_total="", name=name,
        warn_level="", source_diff_row_id=row_id,
    )


def test_sched_sync_skipped_when_no_schedule():
    """退勤予定時刻が空（スケジュール未設定）の行には出勤予定時刻を書かない。

    林広美さん 2026-06 の実例: スケジュール未設定の行に出勤予定だけ書いたところ、
    jinjerが「退勤予定時刻が正しくありません」で行ごと弾いた。
    """
    from quick_export import apply_approved_rows, build_jinjer_row_index, Stats, DIFF_KIND_PUNCH_IN

    headers = ["*従業員ID", "*年月日", "出勤予定時刻", "退勤予定時刻", "出勤1", "退勤1"]
    rows = [["2020003", "2026/6/2", "", "", "9:30", "18:01"]]
    idx = build_jinjer_row_index(headers, rows)
    stats = Stats()
    apply_approved_rows(headers, rows, idx,
                        [_make_approved("2020003", "2026-06-02", DIFF_KIND_PUNCH_IN, "09:00")],
                        stats)
    assert rows[0][headers.index("出勤1")] == "09:00"
    assert rows[0][headers.index("出勤予定時刻")] == ""  # 予定は書かない
    assert stats.overwritten_sched_in == 0


def test_work_status_flag_cleared_when_writing_punches():
    """打刻を書き込む日は勤務状況（0:未打刻1:欠勤）フラグをクリアする。

    矢野瑞穂さん 2026-06-04 の実例: 勤務状況=0(未打刻)のまま打刻を書いたところ、
    jinjerが「出勤1/退勤1/勤務状況が正しくありません」で弾いた。
    """
    from quick_export import apply_approved_rows, build_jinjer_row_index, Stats, DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT

    headers = ["*従業員ID", "*年月日", "出勤予定時刻", "退勤予定時刻",
               "勤務状況（0:未打刻1:欠勤）", "出勤1", "退勤1"]
    rows = [["2020004", "2026/6/4", "09:00", "17:30", "0", "", ""]]
    idx = build_jinjer_row_index(headers, rows)
    stats = Stats()
    apply_approved_rows(headers, rows, idx, [
        _make_approved("2020004", "2026-06-04", DIFF_KIND_PUNCH_IN, "08:50"),
        _make_approved("2020004", "2026-06-04", DIFF_KIND_PUNCH_OUT, "17:40", row_id=2),
    ], stats)
    assert rows[0][headers.index("出勤1")] == "08:50"
    assert rows[0][headers.index("退勤1")] == "17:40"
    assert rows[0][headers.index("勤務状況（0:未打刻1:欠勤）")] == ""


def test_clearing_both_punches_clears_break_leftovers():
    """出勤1・退勤1を両方削除した日は休憩の残り（休憩1/復帰1/休憩時間）もクリアする。

    塚本靖さん 2026-06-30 の実例: 打刻を消したのに休憩時間01:00が残り、
    jinjerが「出勤1が正しくありません」で弾いた。
    """
    from quick_export import apply_approved_rows, build_jinjer_row_index, Stats, DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT

    headers = ["*従業員ID", "*年月日", "出勤1", "退勤1", "休憩1", "復帰1", "休憩時間"]
    rows = [["2020005", "2026/6/30", "9:00", "18:23", "12:00", "13:00", "01:00"]]
    idx = build_jinjer_row_index(headers, rows)
    stats = Stats()
    apply_approved_rows(headers, rows, idx, [
        _make_approved("2020005", "2026-06-30", DIFF_KIND_PUNCH_IN, ""),
        _make_approved("2020005", "2026-06-30", DIFF_KIND_PUNCH_OUT, "", row_id=2),
    ], stats)
    assert rows[0][headers.index("出勤1")] == ""
    assert rows[0][headers.index("退勤1")] == ""
    assert rows[0][headers.index("休憩1")] == ""
    assert rows[0][headers.index("復帰1")] == ""
    assert rows[0][headers.index("休憩時間")] == "00:00"


def test_clearing_one_punch_warns_about_leftover():
    """片側の打刻だけを削除して他方が残る場合は警告する（自動では辻褄を合わせない）。

    塚本靖さん 2026-06-23 の実例: 退勤1だけ消して出勤1が残り、jinjerが
    「退勤1が正しくありません」で弾いた。
    """
    from quick_export import apply_approved_rows, build_jinjer_row_index, Stats, DIFF_KIND_PUNCH_OUT

    headers = ["*従業員ID", "*年月日", "出勤1", "退勤1", "休憩時間"]
    rows = [["2020006", "2026/6/23", "9:00", "18:40", "01:00"]]
    idx = build_jinjer_row_index(headers, rows)
    stats = Stats()
    apply_approved_rows(headers, rows, idx, [
        _make_approved("2020006", "2026-06-23", DIFF_KIND_PUNCH_OUT, ""),
    ], stats)
    assert rows[0][headers.index("退勤1")] == ""
    assert rows[0][headers.index("出勤1")] == "9:00"  # 自動では消さない
    assert rows[0][headers.index("休憩時間")] == "01:00"  # 触らない
    assert any("片側だけ" in w for w in stats.warnings)


def test_unchanged_rows_are_not_output(tmp_path):
    """ツールが変更していない日の行はアップロードCSVに出力しない。

    無変更行を送り返すと、jinjer側の既存不整合（有給残0・勤務状況の矛盾など）で
    エラー通知が汚れるだけ（池村さん・佐藤さん・福島さん 2026-06 の実例）。
    """
    diff_path = tmp_path / "diff.xlsx"
    jinjer_path = tmp_path / "jinjer.csv"
    output_path = tmp_path / "out.csv"
    _write_jinjer_csv_3days(jinjer_path)

    pd.DataFrame([{
        "行ID": 1, "従業員ID": "2018057", "氏名": "上原 奏吾",
        "対象日付": "2026-04-02", "差異種別": "退勤",
        "請求勤怠値": "33:30", "自動修正提案値": "33:30", "人間判断": "請求勤怠",
    }]).to_excel(diff_path, sheet_name="差異一覧", index=False)

    result = run_quick_export(diff_path, jinjer_path, output_path, dry_run=False, log_func=lambda _: None)

    assert result.ok
    rows = _read_output_rows(output_path)
    assert [r["*年月日"] for r in rows] == ["2026/4/2"]  # 変更した日だけ
    assert result.stats.unchanged_rows_removed == 2

