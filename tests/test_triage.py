import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.triage import (  # noqa: E402
    classify,
    TRIAGE_NEEDS_CHECK,
    TRIAGE_AUTO_KINTAI,
    TRIAGE_AUTO_OK,
    TRIAGE_INFO_ONLY,
    JUDGE_KINTAI,
    JUDGE_JINJER,
)


def test_danger_is_needs_check():
    t, j = classify(kind="出勤", warn_level="DANGER")
    assert t == TRIAGE_NEEDS_CHECK
    assert j == ""


def test_full_day_leave_zero_punch_is_auto_ok():
    """出勤退勤が0＆休日休暇名1あり＆種別=全日 → 自動OK(jinjer勤怠)。"""
    t, j = classify(
        kind="出勤", warn_level="INFO",
        kintai_value="0:00", holiday_name1="年次有休", holiday_name1_type="全日",
    )
    assert t == TRIAGE_AUTO_OK
    assert j == JUDGE_JINJER


def test_full_day_leave_but_worked_is_not_auto_ok():
    """種別=全日でも請求勤怠に勤務がある（0でない）なら自動OKにしない。"""
    t, j = classify(
        kind="出勤", warn_level="INFO",
        kintai_value="9:00", holiday_name1="年次有休", holiday_name1_type="全日",
    )
    assert t == TRIAGE_AUTO_KINTAI  # 通常の出勤差異として請求勤怠を自動採用
    assert j == JUDGE_KINTAI


def test_comment_is_needs_check():
    t, j = classify(kind="出勤", warn_level="INFO", punch_comment="出勤: KDX出社")
    assert t == TRIAGE_NEEDS_CHECK
    t2, _ = classify(kind="退勤", warn_level="INFO", stamp_comment="打刻忘れ（承認済）")
    assert t2 == TRIAGE_NEEDS_CHECK


def test_warn_is_needs_check():
    t, _ = classify(kind="出勤", warn_level="WARN")
    assert t == TRIAGE_NEEDS_CHECK


def test_half_day_leave_is_needs_check():
    t, _ = classify(kind="出勤", warn_level="INFO", holiday_name1="有休", holiday_name1_type="PM有休")
    assert t == TRIAGE_NEEDS_CHECK


def test_plain_punch_diff_is_auto_kintai():
    """INFO・コメント無し・出退勤の小差分 → 請求勤怠を自動採用。"""
    t, j = classify(kind="出勤", warn_level="INFO")
    assert t == TRIAGE_AUTO_KINTAI
    assert j == JUDGE_KINTAI


def test_break_total_info_is_reference_only():
    """INFOの休憩・総労働は手順3で書き戻せず判断不要 → 参考のみ（要確認に積まない）。"""
    assert classify(kind="休憩", warn_level="INFO")[0] == TRIAGE_INFO_ONLY
    assert classify(kind="総労働時間", warn_level="INFO")[0] == TRIAGE_INFO_ONLY


def test_break_total_warn_or_danger_is_needs_check():
    """ただし休憩・総労働でも DANGER/WARN は要確認に残す。"""
    assert classify(kind="休憩", warn_level="DANGER")[0] == TRIAGE_NEEDS_CHECK
    assert classify(kind="総労働時間", warn_level="WARN")[0] == TRIAGE_NEEDS_CHECK
    # コメント付きの総労働も要確認
    assert classify(kind="総労働時間", warn_level="INFO", stamp_comment="理由")[0] == TRIAGE_NEEDS_CHECK


def test_danger_priority_over_comment_and_leave():
    """DANGER は全日休暇/コメントより優先して要確認。"""
    t, j = classify(
        kind="出勤", warn_level="DANGER",
        kintai_value="0:00", holiday_name1="年次有休", holiday_name1_type="全日",
        punch_comment="何か",
    )
    assert t == TRIAGE_NEEDS_CHECK
    assert j == ""
