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
            "store_id": "40",
        }}

    def test_duplicate_date_without_timestamps_keeps_first(self):
        """タイムスタンプが無い同一日の複数返却は先頭を残す（従来動作の維持）。

        ⚠️ この「同一日に複数返る」現象の正体は**打刻グループごとに別スケジュールが
        あること**だと 2026-08-03 に判明した（store が別グループになっている）。
        本来は呼び出し側が現グループの store_id を渡して絞り込む
        （TestWorkSchedulesStoreFilter 参照）。ここは絞り込まなかった場合の
        フォールバック挙動を固定するためのテスト。
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


class TestWorkSchedulesNewestWins:
    """同一日に新旧2レコードが返るとき、**作成日時が最新**のものを採る

    2026-08-03 実測: 古い方が先頭で返るケースがあり（石下 8/5・木村 8/4 など12日）、
    そのとき汎用データのエクスポート（jinjerの正）は新しい方と一致していた。
    先頭採用のままだと差分判定と反映検証がともに誤る。
    """

    def _rec(self, date, start, end, stamp):
        return {"date": date, "work_schedule": {"start": start, "end": end},
                "break_schedules": [], "store": {"name": "g"},
                "created_at": stamp, "updated_at": stamp}

    def test_newest_wins_when_old_comes_first(self):
        from services.jinjer_api_client import parse_work_schedules_data
        data = {"work_schedules": [
            self._rec("2026-08-05", "09:00:00", "17:30:00", "2026-05-22 05:54:39"),
            self._rec("2026-08-05", "16:45:00", "33:30:00", "2026-08-03 10:40:44"),
        ]}
        got = parse_work_schedules_data(data)["2026-08-05"]
        assert (got["start"], got["end"]) == ("16:45", "33:30")

    def test_newest_wins_when_new_comes_first(self):
        from services.jinjer_api_client import parse_work_schedules_data
        data = {"work_schedules": [
            self._rec("2026-08-05", "16:45:00", "33:30:00", "2026-08-03 10:40:44"),
            self._rec("2026-08-05", "09:00:00", "17:30:00", "2026-05-22 05:54:39"),
        ]}
        got = parse_work_schedules_data(data)["2026-08-05"]
        assert (got["start"], got["end"]) == ("16:45", "33:30")

    def test_missing_timestamps_keeps_first(self):
        from services.jinjer_api_client import parse_work_schedules_data
        a = self._rec("2026-08-05", "09:00:00", "17:30:00", "")
        b = self._rec("2026-08-05", "16:45:00", "33:30:00", "")
        for r in (a, b):
            r.pop("created_at"); r.pop("updated_at")
        got = parse_work_schedules_data({"work_schedules": [a, b]})["2026-08-05"]
        assert (got["start"], got["end"]) == ("9:00", "17:30")


class TestWorkSchedulesStoreFilter:
    """jinjer は打刻グループごとに別のスケジュールを持つ（2026-08-03 実測）

    グループを移動した従業員は旧グループの予定が残り、同じ日について複数レコードが
    返る。旧グループの残骸を現在の予定と読むと「休日に予定残存」と誤判定し、本来
    書けるはずの予定を落とす（石下さん14日・木村さん15日で実害）。
    """

    def _rec(self, date, start, end, store_id, store_name, stamp):
        return {"date": date, "work_schedule": {"start": start, "end": end},
                "break_schedules": [], "store": {"id": store_id, "name": store_name},
                "created_at": stamp, "updated_at": stamp}

    def _data(self):
        return {"work_schedules": [
            # 旧グループ40の残骸（8/1に43へ移動する前に作られたもの）
            self._rec("2026-08-06", "09:00:00", "17:30:00", "40", "140-180時間制",
                      "2026-07-13 16:03:08"),
            # 現グループ43の予定
            self._rec("2026-08-05", "16:45:00", "33:30:00", "43", "140-160時間制",
                      "2026-08-03 10:40:44"),
        ]}

    def test_store_filter_hides_old_group(self):
        from services.jinjer_api_client import parse_work_schedules_data
        got = parse_work_schedules_data(self._data(), store_id="43")
        assert "2026-08-06" not in got, "旧グループの残骸を現在の予定として拾っている"
        assert got["2026-08-05"]["start"] == "16:45"
        assert got["2026-08-05"]["store_id"] == "43"

    def test_without_filter_keeps_everything(self):
        """store_id 未指定なら従来どおり全部返す（他の呼び出し元を壊さない）"""
        from services.jinjer_api_client import parse_work_schedules_data
        got = parse_work_schedules_data(self._data())
        assert set(got) == {"2026-08-05", "2026-08-06"}

    def test_same_group_duplicate_takes_newest(self):
        """同じグループ内で同じ日が複数返る場合は最新を採る"""
        from services.jinjer_api_client import parse_work_schedules_data
        data = {"work_schedules": [
            self._rec("2026-08-05", "09:00:00", "17:30:00", "43", "g", "2026-05-22 05:54:39"),
            self._rec("2026-08-05", "16:45:00", "33:30:00", "43", "g", "2026-08-03 10:40:44"),
        ]}
        got = parse_work_schedules_data(data, store_id="43")["2026-08-05"]
        assert (got["start"], got["end"]) == ("16:45", "33:30")
