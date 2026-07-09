# -*- coding: utf-8 -*-
"""kintai_import_runner の純粋関数テスト"""
import pytest

from services.kintai_import_runner import (
    breaks_of_row,
    compare_row,
    is_kyuka_row,
    norm_date_iso,
    t2m,
)


class TestT2m:
    def test_basic(self):
        assert t2m("8:45") == 525
        assert t2m("08:45") == 525

    def test_seconds_stripped(self):
        assert t2m("09:15:00") == 555

    def test_over_24h(self):
        """夜勤の24時超表記はそのまま分に換算する"""
        assert t2m("33:30") == 2010
        assert t2m("33:30:00") == 2010

    def test_empty_and_invalid(self):
        assert t2m("") is None
        assert t2m(None) is None
        assert t2m("1 day, 7:00:00") is None  # timedelta漏れは不正値として扱う


class TestNormDateIso:
    def test_formats(self):
        assert norm_date_iso("2026/6/1") == "2026-06-01"
        assert norm_date_iso("2026-06-01") == "2026-06-01"
        assert norm_date_iso("6/1/2026") == "2026-06-01"  # Excel米国式の巻き込み対策

    def test_invalid(self):
        assert norm_date_iso("") is None
        assert norm_date_iso("六月一日") is None


class TestIsKyukaRow:
    def test_no_kyuka(self):
        kyuka, why = is_kyuka_row({"休日休暇名1": "", "AM有休": "FALSE"})
        assert not kyuka

    def test_kyuka_name(self):
        kyuka, why = is_kyuka_row({"休日休暇名1": "年次有給"})
        assert kyuka
        assert "年次有給" in why

    def test_shubetsu_only(self):
        """守屋さん6/23パターン: 種別列にAM有休が入るケースを検知する"""
        kyuka, why = is_kyuka_row({"休日休暇名1": "年次有給", "休日休暇名1：種別": "AM有休"})
        assert kyuka
        assert "AM有休" in why

    def test_am_flag(self):
        kyuka, _ = is_kyuka_row({"AM有休": "TRUE"})
        assert kyuka

    def test_false_flags_ignored(self):
        kyuka, _ = is_kyuka_row({"AM有休": "FALSE", "PM有休": "0", "休日休暇名1": ""})
        assert not kyuka


class TestBreaksOfRow:
    def test_pair(self):
        row = {"休憩予定時刻1": "12:00", "復帰予定時刻1": "13:00"}
        assert breaks_of_row(row) == [(720, 780)]

    def test_half_pair_ignored(self):
        row = {"休憩予定時刻1": "12:00", "復帰予定時刻1": ""}
        assert breaks_of_row(row) == []

    def test_multiple_sorted(self):
        row = {"休憩予定時刻1": "15:00", "復帰予定時刻1": "15:30",
               "休憩予定時刻2": "12:00", "復帰予定時刻2": "13:00"}
        assert breaks_of_row(row) == [(720, 780), (900, 930)]

    def test_overnight_break(self):
        """夜勤の24時超休憩（30:00-31:30）"""
        row = {"休憩予定時刻1": "30:00", "復帰予定時刻1": "31:30"}
        assert breaks_of_row(row) == [(1800, 1890)]


class TestCompareRow:
    def test_all_match(self):
        row = {"出勤予定時刻": "8:50", "退勤予定時刻": "17:10",
               "休憩予定時刻1": "12:00", "復帰予定時刻1": "13:00",
               "出勤1": "8:50", "退勤1": "18:00"}
        sched = {"start": "08:50", "end": "17:10", "breaks": [("12:00", "13:00")]}
        att = {"in": "8:50", "out": "18:00"}
        assert compare_row(row, sched, att) == []

    def test_schedule_start_ng(self):
        row = {"出勤予定時刻": "8:50"}
        sched = {"start": "9:00", "breaks": []}
        ngs = compare_row(row, sched, {})
        assert len(ngs) == 1
        assert ngs[0][0] == "出勤予定"

    def test_empty_columns_not_compared(self):
        """CSVで空欄の項目は無処理なので比較しない"""
        row = {"出勤予定時刻": "", "退勤予定時刻": "", "出勤1": "", "退勤1": ""}
        assert compare_row(row, {"start": "9:00", "breaks": []}, {"in": "9:12"}) == []

    def test_punch_ng_has_marume_note(self):
        row = {"出勤1": "9:00"}
        ngs = compare_row(row, {}, {"in": "9:03"})
        assert len(ngs) == 1
        assert "まるめ前" in ngs[0][3]

    def test_breaks_ng(self):
        row = {"休憩予定時刻1": "12:00", "復帰予定時刻1": "13:00"}
        sched = {"start": "", "breaks": []}
        ngs = compare_row(row, sched, {})
        assert len(ngs) == 1
        assert ngs[0][0] == "休憩予定"
        assert "(なし)" in ngs[0][2]

    def test_missing_sched_and_att(self):
        row = {"出勤予定時刻": "8:50", "出勤1": "8:50"}
        ngs = compare_row(row, None, None)
        assert {n[0] for n in ngs} == {"出勤予定", "出勤打刻"}
