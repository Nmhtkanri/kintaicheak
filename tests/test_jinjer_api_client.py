# -*- coding: utf-8 -*-
"""jinjer_api_client の純粋パーサテスト"""
from services.jinjer_api_client import (
    parse_requested_day_offs_data,
    parse_work_schedules_data,
)


class TestParseWorkSchedulesData:
    def test_basic(self):
        data = {"work_schedules": [{
            "date": "2026-07-01",
            "work_schedule": {"start": "09:00:00", "end": "18:00:00"},
            "break_schedules": [{"start": "12:00:00", "end": "13:00:00"}],
            "store": {"id": 40, "name": "140-180時間制"},
        }]}
        out = parse_work_schedules_data(data)
        assert out == {"2026-07-01": {
            "start": "9:00", "end": "18:00",
            "breaks": [("12:00", "13:00")], "store": "140-180時間制",
        }}

    def test_duplicate_date_keeps_first(self):
        """同一日の新旧二重返却は先頭（新しい方）を採用する。

        能美(2024047)の2026-07実測: 先頭=誤上書きレイヤーではなく画面表示と
        一致する新バージョン。後勝ちで詰めると旧バージョンを拾ってしまう。
        """
        data = {"work_schedules": [
            {"date": "2026-07-03",
             "work_schedule": {"start": "09:00:00", "end": "17:30:00"},
             "break_schedules": [{"start": "12:00:00", "end": "13:00:00"}],
             "store": {"name": "140-180時間制"}},
            {"date": "2026-07-03",  # 旧バージョン（2件目）
             "work_schedule": {"start": "16:45:00", "end": "33:30:00"},
             "break_schedules": [{"start": "24:00:00", "end": "25:45:00"}],
             "store": {"name": "140-160時間制"}},
        ]}
        out = parse_work_schedules_data(data)
        assert out["2026-07-03"]["start"] == "9:00"
        assert out["2026-07-03"]["end"] == "17:30"
        assert out["2026-07-03"]["store"] == "140-180時間制"

    def test_break_pair_one_sided_excluded(self):
        data = {"work_schedules": [{
            "date": "2026-07-01",
            "work_schedule": {"start": "9:00", "end": "18:00"},
            "break_schedules": [{"start": "12:00:00", "end": ""},
                                {"start": "", "end": "13:00:00"}],
        }]}
        out = parse_work_schedules_data(data)
        assert out["2026-07-01"]["breaks"] == []

    def test_overnight_kept(self):
        data = {"work_schedules": [{
            "date": "2026-07-04",
            "work_schedule": {"start": "16:30:00", "end": "33:30:00"},
            "break_schedules": [{"start": "20:30:00", "end": "22:00:00"}],
        }]}
        out = parse_work_schedules_data(data)
        assert out["2026-07-04"]["end"] == "33:30"
        assert out["2026-07-04"]["breaks"] == [("20:30", "22:00")]

    def test_empty(self):
        assert parse_work_schedules_data({}) == {}
        assert parse_work_schedules_data({"work_schedules": None}) == {}


class TestParseRequestedDayOffsData:
    def test_slash_and_hyphen_dates(self):
        data = {"requested_day_offs": [
            {"date": "2026/07/06",
             "day_off_classification": {"name": "年次有休"}, "status": 1},
            {"date": "2026-7-25",
             "day_off_classification": {"name": "特別休暇"}, "status": 2},
        ]}
        out = parse_requested_day_offs_data(data)
        assert out == {"2026-07-06": "年次有休/status=1",
                       "2026-07-25": "特別休暇/status=2"}

    def test_missing_classification(self):
        data = {"requested_day_offs": [{"date": "2026/07/01", "status": None}]}
        out = parse_requested_day_offs_data(data)
        assert out["2026-07-01"].startswith("休暇/")

    def test_invalid_date_skipped(self):
        data = {"requested_day_offs": [{"date": "七月六日"}]}
        assert parse_requested_day_offs_data(data) == {}

    def test_empty(self):
        assert parse_requested_day_offs_data({}) == {}
        assert parse_requested_day_offs_data({"requested_day_offs": None}) == {}
