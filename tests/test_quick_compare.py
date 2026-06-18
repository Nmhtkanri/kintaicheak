import os
import sys
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from quick_compare import (  # noqa: E402
    DIFF_COLUMNS,
    DIFF_KIND_PUNCH_IN,
    DIFF_KIND_PUNCH_OUT,
    DIFF_KIND_TOTAL,
    JINJER_HEADERS,
    LogEntry,
    compute_diffs,
    kintai_total_minutes,
    normalize_kintai_result_columns,
    resolve_jinjer_extra_columns,
    to_jinjer_overnight_punch_out,
)


def _jinjer_row(**overrides):
    row = {
        JINJER_HEADERS["name"]: "上原 奏吾",
        JINJER_HEADERS["emp_id"]: "2018057",
        JINJER_HEADERS["date"]: "2026/4/1",
        JINJER_HEADERS["punch_in_1"]: "",
        JINJER_HEADERS["punch_out_1"]: "",
        JINJER_HEADERS["break_total"]: "",
        JINJER_HEADERS["total_work"]: "",
        JINJER_HEADERS["finalized"]: "",
    }
    row.update(overrides)
    return row


def test_compute_diffs_creates_punch_rows_when_jinjer_punches_are_blank():
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00",
        "jinjer_出勤": "",
        "出勤差分(分)": None,
        "勤務表_退勤": "18:00",
        "jinjer_退勤": "",
        "退勤差分(分)": None,
        "_source_file": "勤怠突合結果.xlsx",
    }])
    logs: list[LogEntry] = []

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): _jinjer_row()},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
    )

    punch_rows = [row for row in rows if row.kind in {DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT}]
    assert [row.kind for row in punch_rows] == [DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT]
    assert [row.auto_fix_value for row in punch_rows] == ["9:00", "18:00"]
    assert punch_rows[0].warn_reason == "jinjer出勤なし / 請求勤怠側に時刻あり"
    assert punch_rows[1].warn_reason == "jinjer退勤なし / 請求勤怠側に時刻あり"


def test_normalize_kintai_result_columns_accepts_seikyu_kintai_headers():
    df = pd.DataFrame([{
        "請求勤怠_出勤": "9:00",
        "請求勤怠_退勤": "18:00",
        "請求勤怠_総労働": "9:00",
    }])

    normalized = normalize_kintai_result_columns(df)

    assert normalized.iloc[0]["勤務表_出勤"] == "9:00"
    assert normalized.iloc[0]["勤務表_退勤"] == "18:00"
    assert normalized.iloc[0]["勤務表_総労働"] == "9:00"


def test_to_jinjer_overnight_punch_out():
    # 翌朝退勤（出勤 > 退勤）→ 24時超表記へ変換
    assert to_jinjer_overnight_punch_out("21:00", "08:15") == "32:15"
    assert to_jinjer_overnight_punch_out("17:00", "01:30") == "25:30"
    # 通常勤務（出勤 < 退勤）→ そのまま
    assert to_jinjer_overnight_punch_out("9:00", "18:00") == "18:00"
    # すでに24時超表記なら再変換しない（冪等）
    assert to_jinjer_overnight_punch_out("21:00", "32:15") == "32:15"
    # 出退勤のどちらかが空/不正ならそのまま
    assert to_jinjer_overnight_punch_out("", "08:15") == "08:15"
    assert to_jinjer_overnight_punch_out("21:00", "") == ""


def test_compute_diffs_converts_overnight_punch_out_to_jinjer_format():
    """夜勤（請求勤怠 出勤21:00→退勤翌08:15）の退勤提案値が 32:15 になる。"""
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "21:00",
        "jinjer_出勤": "21:00",
        "出勤差分(分)": 0,
        "勤務表_退勤": "08:15",
        "jinjer_退勤": "",
        "退勤差分(分)": None,
        "_source_file": "勤怠突合結果.xlsx",
    }])
    logs: list[LogEntry] = []

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): _jinjer_row(**{JINJER_HEADERS["punch_in_1"]: "21:00"})},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
    )

    out_rows = [r for r in rows if r.kind == DIFF_KIND_PUNCH_OUT]
    assert len(out_rows) == 1
    # 請求勤怠値（表示）は元の 08:15、自動修正提案値は jinjer 用の 32:15
    assert out_rows[0].kintai_value == "08:15"
    assert out_rows[0].auto_fix_value == "32:15"


def test_diff_columns_include_manual_review_fields():
    # 「手入力修正値」は「打刻修正」に改名済み
    assert "打刻修正" in DIFF_COLUMNS
    assert "手入力修正値" not in DIFF_COLUMNS
    assert "手入力休憩1" in DIFF_COLUMNS
    assert "手入力復帰1" in DIFF_COLUMNS
    assert "手入力休憩時間" in DIFF_COLUMNS


def test_diff_columns_include_schedule_and_leave_fields():
    for col in ("出勤予定", "退勤予定", "休憩予定", "有休", "AM有休", "PM有休"):
        assert col in DIFF_COLUMNS


def test_resolve_jinjer_extra_columns_exact_and_partial():
    cols = [
        "名前", "*従業員ID", "*年月日",
        "出勤予定時刻", "退勤予定時刻", "休憩予定時間",
        "有休", "AM有休", "PM有休",
    ]
    resolved = resolve_jinjer_extra_columns(cols)
    assert resolved["出勤予定"] == "出勤予定時刻"
    assert resolved["退勤予定"] == "退勤予定時刻"
    assert resolved["休憩予定"] == "休憩予定時間"
    assert resolved["有休"] == "有休"
    assert resolved["AM有休"] == "AM有休"
    assert resolved["PM有休"] == "PM有休"


def test_resolve_jinjer_extra_columns_leave_is_exact_only():
    # 「有休」列が無く AM有休/PM有休 だけある場合、「有休」を部分一致で誤ヒットさせない
    cols = ["名前", "AM有休", "PM有休"]
    resolved = resolve_jinjer_extra_columns(cols)
    assert "有休" not in resolved
    assert resolved.get("AM有休") == "AM有休"
    assert resolved.get("PM有休") == "PM有休"


def test_resolve_jinjer_extra_columns_missing_returns_no_key():
    resolved = resolve_jinjer_extra_columns(["名前", "*従業員ID", "*年月日"])
    assert resolved == {}


def test_compute_diffs_transcribes_schedule_and_leave():
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00",
        "jinjer_出勤": "",
        "出勤差分(分)": None,
        "勤務表_退勤": "18:00",
        "jinjer_退勤": "",
        "退勤差分(分)": None,
        "_source_file": "勤怠突合結果.xlsx",
    }])
    logs: list[LogEntry] = []
    jrow = _jinjer_row(**{
        "出勤予定時刻": "9:00",
        "退勤予定時刻": "18:00",
        "休憩予定時間": "1:00",
        "有休": "",
        "AM有休": "1",
        "PM有休": "",
    })
    extra_cols = resolve_jinjer_extra_columns(list(jrow.keys()))

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): jrow},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
        extra_cols,
    )

    assert rows, "差異行が生成されること"
    r = rows[0]
    assert r.sched_in == "9:00"
    assert r.sched_out == "18:00"
    assert r.sched_break == "1:00"
    assert r.am_yukyu == "1"
    assert r.yukyu == ""
    assert r.pm_yukyu == ""


def test_kintai_total_minutes_prefers_actual_work():
    """請求勤怠の正味労働(勤務表_実働時間)が拘束時間より優先される。"""
    row = pd.Series({
        "勤務表_実働時間": "7:00",     # 正味（休憩控除後）
        "勤務表_総労働時間": "9:00",   # 拘束（退勤−出勤）
        "勤務表_出勤": "9:00",
        "勤務表_退勤": "18:00",
    })
    minutes, hhmm = kintai_total_minutes(row)
    assert minutes == 420
    assert hhmm == "7:00"


def test_kintai_total_minutes_fallback_when_actual_blank():
    """実働列が空なら従来の拘束時間にフォールバックする（後方互換）。"""
    row = pd.Series({"勤務表_実働時間": "", "勤務表_総労働時間": "9:00"})
    minutes, _ = kintai_total_minutes(row)
    assert minutes == 540


def test_compute_diffs_total_compares_net_vs_net():
    """総労働差異は請求勤怠の正味(実働) vs jinjer 総労働 で突合される。

    拘束時間9:00ではなく実働8:00で比較されるため、jinjer総労働8:00とは差異なし。
    """
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00", "jinjer_出勤": "9:00", "出勤差分(分)": 0,
        "勤務表_退勤": "18:00", "jinjer_退勤": "18:00", "退勤差分(分)": 0,
        "勤務表_実働時間": "8:00",  # 正味（休憩1h控除後）
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    jrow = _jinjer_row(**{
        JINJER_HEADERS["total_work"]: "8:00",
        JINJER_HEADERS["break_total"]: "1:00",
        JINJER_HEADERS["punch_in_1"]: "9:00",
        JINJER_HEADERS["punch_out_1"]: "18:00",
    })
    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): jrow},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
    )
    total_rows = [r for r in rows if r.kind == DIFF_KIND_TOTAL]
    assert total_rows == []  # 実働8:00 = jinjer総労働8:00 → 差異なし
