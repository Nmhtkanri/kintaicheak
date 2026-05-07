"""shift_resolver.py の単体テスト

会話の中で実例として出てきた各パターンを regression test 化:
- 小嶋桃子（B 単純対応）
- 戸松工（●/△/明 夜勤跨ぎ）
- 田村（2/4/a/5 数字コード - 24:00 連結）
- 大堀チーム（A/B/明 + ×/空欄）
"""

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from datetime import date, time

import pandas as pd

from services.shift_resolver import (
    resolve_shifts,
    normalize_legend,
    _merge_consecutive_overnight,
    DEFAULT_OFF_MARKERS,
)


# =============================================================================
# normalize_legend
# =============================================================================

def test_normalize_legend_basic():
    raw = [
        {"code": "B", "label": "B勤", "start_time": "12:30", "end_time": "21:00", "break_minutes": 60},
    ]
    legend = normalize_legend(raw)
    assert "B" in legend
    assert legend["B"]["start_time"] == time(12, 30)
    assert legend["B"]["end_time"] == time(21, 0)
    assert legend["B"]["break_minutes"] == 60
    assert legend["B"]["is_off"] is False


def test_normalize_legend_off_when_no_times():
    raw = [
        {"code": "振", "label": "振休"},
    ]
    legend = normalize_legend(raw)
    # 時刻が無いものは自動的に休扱い
    assert legend["振"]["is_off"] is True


def test_normalize_legend_explicit_off():
    raw = [
        {"code": "明", "label": "明け休", "is_off": True},
    ]
    legend = normalize_legend(raw)
    assert legend["明"]["is_off"] is True


def test_normalize_legend_overnight_time():
    """25:00 などの 24h 超表記は %24 で正規化されること"""
    raw = [
        {"code": "●", "label": "夜勤", "start_time": "16:30", "end_time": "25:00"},
    ]
    legend = normalize_legend(raw)
    assert legend["●"]["start_time"] == time(16, 30)
    # 25:00 % 24 = 1:00（既存パイプラインとの整合性）
    assert legend["●"]["end_time"] == time(1, 0)


def test_normalize_legend_skips_invalid():
    raw = [
        {"code": "", "start_time": "9:00", "end_time": "17:00"},   # コード空
        {"code": None, "start_time": "9:00"},                       # コードNone
        "invalid",                                                  # dict じゃない
    ]
    legend = normalize_legend(raw)
    assert legend == {}


# =============================================================================
# 小嶋桃子パターン: B = 12:30-21:00 の単純対応
# =============================================================================

def test_kojima_pattern_simple_resolve():
    legend = [
        {"code": "B", "label": "B勤", "start_time": "12:30", "end_time": "21:00", "break_minutes": 60},
    ]
    employees = [
        {
            "name": "小嶋桃子",
            "shifts": [
                {"date": "2026-04-01", "code": "B"},
                {"date": "2026-04-02", "code": "B"},
                {"date": "2026-04-03", "code": "B"},
                {"date": "2026-04-04", "code": ""},   # 土曜=休
                {"date": "2026-04-05", "code": ""},   # 日曜=休
                {"date": "2026-04-06", "code": "B"},
            ],
        }
    ]
    df = resolve_shifts(legend, employees)
    assert len(df) == 4  # 休 2日を除いて 4 レコード
    assert all(df["氏名"] == "小嶋桃子")
    assert all(df["出勤時刻"] == time(12, 30))
    assert all(df["退勤時刻"] == time(21, 0))


# =============================================================================
# 戸松工パターン: ● = 夜勤(16:30-翌1:00), 明 = 明け休（休扱い）
# =============================================================================

def test_tomatsu_pattern_yakkin_and_meike():
    legend = [
        {"code": "●", "label": "夜勤", "start_time": "16:30", "end_time": "25:00", "break_minutes": 60},
        {"code": "△", "label": "日中受付", "start_time": "9:00", "end_time": "17:30", "break_minutes": 60},
        {"code": "/", "label": "夜勤明け", "is_off": True},
        {"code": "明", "label": "明け", "is_off": True},
        {"code": "振", "label": "振休", "is_off": True},
    ]
    employees = [
        {
            "name": "戸松工",
            "shifts": [
                {"date": "2026-04-01", "code": "振"},  # 振休
                {"date": "2026-04-02", "code": "△"},
                {"date": "2026-04-05", "code": "●"},  # 夜勤
                {"date": "2026-04-06", "code": "/"},  # 明け＝休
                {"date": "2026-04-07", "code": ""},
            ],
        }
    ]
    df = resolve_shifts(legend, employees)
    # 休扱い 3 つを除いて 2 レコード（4/2 と 4/5）
    assert len(df) == 2

    # 夜勤レコードの確認: 16:30 → 1:00（25:00 %24 = 1:00）
    yakkin = df[df["日付"] == date(2026, 4, 5)].iloc[0]
    assert yakkin["出勤時刻"] == time(16, 30)
    assert yakkin["退勤時刻"] == time(1, 0)

    nicchu = df[df["日付"] == date(2026, 4, 2)].iloc[0]
    assert nicchu["出勤時刻"] == time(9, 0)
    assert nicchu["退勤時刻"] == time(17, 30)


# =============================================================================
# 田村パターン: 4(16:30-24:00) + a(24:00-翌09:00) を統合 → 16:30-09:00
# =============================================================================

def test_tamura_pattern_consecutive_merge():
    """4日と5日の連続シフトが「4日 16:30-09:00」に統合されること"""
    legend = [
        {"code": "2", "label": "早番", "start_time": "8:30", "end_time": "17:00", "break_minutes": 60},
        {"code": "4", "label": "遅番", "start_time": "16:30", "end_time": "24:00", "break_minutes": 60},
        {"code": "a", "label": "深夜", "start_time": "24:00", "end_time": "33:00", "break_minutes": 60},
        {"code": "5", "label": "深夜+5分前", "start_time": "23:55", "end_time": "33:00", "break_minutes": 60},
    ]
    employees = [
        {
            "name": "田村",
            "shifts": [
                {"date": "2026-04-01", "code": "2"},
                {"date": "2026-04-02", "code": "2"},
                {"date": "2026-04-04", "code": "4"},  # ★ 16:30-24:00
                {"date": "2026-04-05", "code": "a"},  # ★ 24:00-翌09:00 → 統合される
                {"date": "2026-04-06", "code": "4"},  # ★ 16:30-24:00
                {"date": "2026-04-07", "code": "a"},  # ★ 24:00-翌09:00 → 統合される
                {"date": "2026-04-08", "code": "4"},  # 単独（翌日に a が無い）
            ],
        }
    ]
    df = resolve_shifts(legend, employees)
    # 4日: 4+a 統合 / 6日: 4+a 統合 / 8日: 4 単独 / 1日,2日: 通常 → 計5レコード
    assert len(df) == 5

    # 4日のレコード: start=16:30, end=09:00（33:00 %24）
    rec4 = df[df["日付"] == date(2026, 4, 4)].iloc[0]
    assert rec4["出勤時刻"] == time(16, 30)
    assert rec4["退勤時刻"] == time(9, 0)

    # 5日のレコードは存在しない（4日に統合された）
    assert len(df[df["日付"] == date(2026, 4, 5)]) == 0

    # 6日も同様に統合
    rec6 = df[df["日付"] == date(2026, 4, 6)].iloc[0]
    assert rec6["出勤時刻"] == time(16, 30)
    assert rec6["退勤時刻"] == time(9, 0)
    assert len(df[df["日付"] == date(2026, 4, 7)]) == 0

    # 8日は単独 → 退勤=24:00 → time(0,0)
    rec8 = df[df["日付"] == date(2026, 4, 8)].iloc[0]
    assert rec8["出勤時刻"] == time(16, 30)
    assert rec8["退勤時刻"] == time(0, 0)


def test_merge_does_not_join_non_consecutive_days():
    """日付が連続していない場合は統合しない"""
    legend = [
        {"code": "4", "label": "遅番", "start_time": "16:30", "end_time": "24:00"},
        {"code": "a", "label": "深夜", "start_time": "24:00", "end_time": "33:00"},
    ]
    employees = [{
        "name": "X",
        "shifts": [
            {"date": "2026-04-04", "code": "4"},
            {"date": "2026-04-06", "code": "a"},  # 1日空いている
        ],
    }]
    df = resolve_shifts(legend, employees)
    assert len(df) == 2  # 統合されない


def test_merge_disabled_flag():
    """merge_overnight=False で統合を無効化できる"""
    legend = [
        {"code": "4", "label": "遅番", "start_time": "16:30", "end_time": "24:00"},
        {"code": "a", "label": "深夜", "start_time": "24:00", "end_time": "33:00"},
    ]
    employees = [{
        "name": "X",
        "shifts": [
            {"date": "2026-04-04", "code": "4"},
            {"date": "2026-04-05", "code": "a"},
        ],
    }]
    df = resolve_shifts(legend, employees, merge_overnight=False)
    assert len(df) == 2


def test_merge_keeps_different_employees_separate():
    """別人同士は統合しない"""
    legend = [
        {"code": "4", "label": "遅番", "start_time": "16:30", "end_time": "24:00"},
        {"code": "a", "label": "深夜", "start_time": "24:00", "end_time": "33:00"},
    ]
    employees = [
        {"name": "A", "shifts": [{"date": "2026-04-04", "code": "4"}]},
        {"name": "B", "shifts": [{"date": "2026-04-05", "code": "a"}]},
    ]
    df = resolve_shifts(legend, employees)
    assert len(df) == 2


# =============================================================================
# 大堀チームパターン: A/B/明 + 空欄/×
# =============================================================================

def test_oohori_pattern_with_off_markers():
    legend = [
        {"code": "A", "label": "A勤", "start_time": "9:00", "end_time": "17:30", "break_minutes": 60},
        {"code": "B", "label": "B勤", "start_time": "16:45", "end_time": "33:30", "break_minutes": 105},
        {"code": "明", "label": "明け休", "is_off": True},
    ]
    employees = [{
        "name": "大堀",
        "shifts": [
            {"date": "2026-04-01", "code": "B"},
            {"date": "2026-04-02", "code": "明"},  # 休
            {"date": "2026-04-03", "code": ""},     # 休
            {"date": "2026-04-04", "code": "A"},
            {"date": "2026-04-05", "code": "×"},   # 休（DEFAULT_OFF_MARKERS）
        ],
    }]
    df = resolve_shifts(legend, employees)
    assert len(df) == 2  # 4/1 と 4/4 のみ

    # B勤は 16:45 → 9:30（33:30 %24）
    b = df[df["日付"] == date(2026, 4, 1)].iloc[0]
    assert b["出勤時刻"] == time(16, 45)
    assert b["退勤時刻"] == time(9, 30)

    a = df[df["日付"] == date(2026, 4, 4)].iloc[0]
    assert a["出勤時刻"] == time(9, 0)
    assert a["退勤時刻"] == time(17, 30)


# =============================================================================
# その他のエッジケース
# =============================================================================

def test_unknown_code_skipped():
    """凡例に無い記号はスキップ（エラーではなく）"""
    legend = [
        {"code": "A", "start_time": "9:00", "end_time": "17:30"},
    ]
    employees = [{
        "name": "X",
        "shifts": [
            {"date": "2026-04-01", "code": "A"},
            {"date": "2026-04-02", "code": "ZZ"},  # 未知
        ],
    }]
    df = resolve_shifts(legend, employees)
    assert len(df) == 1


def test_empty_input_returns_empty_df():
    df = resolve_shifts([], [])
    assert isinstance(df, pd.DataFrame)
    assert len(df) == 0
    # 列構成は維持されている
    assert list(df.columns) == ["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"]


def test_custom_off_markers():
    """ユーザー指定の追加 off_markers が効くこと"""
    legend = [
        {"code": "A", "start_time": "9:00", "end_time": "17:30"},
    ]
    employees = [{
        "name": "X",
        "shifts": [
            {"date": "2026-04-01", "code": "A"},
            {"date": "2026-04-02", "code": "★"},  # カスタム休扱い
        ],
    }]
    df = resolve_shifts(legend, employees, off_markers={"★"})
    assert len(df) == 1


def test_comment_includes_code_label():
    """コメント欄に [記号=ラベル] が入る（突合結果の追跡用）"""
    legend = [
        {"code": "B", "label": "B勤", "start_time": "12:30", "end_time": "21:00"},
    ]
    employees = [{"name": "X", "shifts": [{"date": "2026-04-01", "code": "B"}]}]
    df = resolve_shifts(legend, employees)
    assert "[B=B勤]" in df.iloc[0]["コメント"]


def test_date_string_formats():
    """複数の日付フォーマットを受け付ける"""
    legend = [{"code": "A", "start_time": "9:00", "end_time": "17:30"}]
    employees = [{
        "name": "X",
        "shifts": [
            {"date": "2026-04-01", "code": "A"},
            {"date": "2026/04/02", "code": "A"},
            {"date": "2026.04.03", "code": "A"},
        ],
    }]
    df = resolve_shifts(legend, employees)
    assert len(df) == 3
