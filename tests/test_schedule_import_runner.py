# -*- coding: utf-8 -*-
"""schedule_import_runner の純粋関数テスト（dict直渡し・HTTPモックなし）"""
from pathlib import Path

import pytest

from services.schedule_import_runner import (
    GENERIC_IMPORT_HEADER,
    KUBUN_DAYOFF_REST,
    KUBUN_DAYOFF_SKIP,
    KUBUN_DELETE_NEEDED,
    KUBUN_NO_GROUP,
    build_diff_plan,
    build_import_rows,
    build_template_index,
    find_unknown_cells,
    merge_grids,
    norm_cell,
    norm_time,
    parse_grid_rows,
    plan_fingerprint,
    rows_to_csv_bytes,
    verify_plan_rows,
)


class TestNormTime:
    def test_seconds_stripped(self):
        assert norm_time("09:00:00") == "9:00"

    def test_over_24h_kept(self):
        assert norm_time("33:30") == "33:30"
        assert norm_time("33:30:00") == "33:30"

    def test_invalid(self):
        assert norm_time("") == ""
        assert norm_time(None) == ""
        assert norm_time("休み") == ""


class TestNormCell:
    def test_fullwidth_to_halfwidth(self):
        """手作りグリッドの全角雛形ID（UAL時給制2026-07実データ）を半角へ"""
        assert norm_cell("１") == "1"
        assert norm_cell("ＪＴ１") == "JT1"

    def test_kanji_unchanged(self):
        assert norm_cell("所") == "所"
        assert norm_cell("休み") == "休み"

    def test_strip(self):
        assert norm_cell(" BBS3 ") == "BBS3"
        assert norm_cell(None) == ""


class TestParseGridRows:
    def _rows(self):
        return [
            ["2026年", "7月", "1", "2", "3"],
            ["氏名", "従業員ID", "水", "木", "金"],
            ["尾川", "2009006", "BBS3", "所", ""],
        ]

    def test_basic(self):
        g = parse_grid_rows(self._rows(), filename="grid.csv")
        assert (g["year"], g["month"]) == (2026, 7)
        emp = g["employees"]["2009006"]
        assert emp["name"] == "尾川"
        assert emp["file"] == "grid.csv"
        assert emp["days"] == {1: "BBS3", 2: "所", 3: ""}

    def test_bad_header_raises(self):
        rows = self._rows()
        rows[0][0] = "年度2026"
        with pytest.raises(ValueError):
            parse_grid_rows(rows, filename="bad.csv")

    def test_short_data_row_tolerated(self):
        rows = self._rows()
        rows[2] = ["椎津", "2009007", "BBS3"]  # 2日目以降のセルが無い
        g = parse_grid_rows(rows, filename="grid.csv")
        assert g["employees"]["2009007"]["days"] == {1: "BBS3", 2: "", 3: ""}

    def test_no_empid_row_separated(self):
        rows = self._rows() + [["山田", "", "1", "休", "1"]]
        g = parse_grid_rows(rows, filename="grid.csv")
        assert g["no_empid_names"] == ["山田"]
        assert "山田" not in {v["name"] for v in g["employees"].values()}

    def test_non_digit_day_column_ignored(self):
        rows = self._rows()
        rows[0].append("備考")  # 日番号でない列
        rows[2].append("メモ")
        g = parse_grid_rows(rows, filename="grid.csv")
        assert set(g["employees"]["2009006"]["days"]) == {1, 2, 3}


class TestMergeGrids:
    def _grid(self, days, file="a.csv"):
        return {"year": 2026, "month": 7,
                "employees": {"2009006": {"name": "尾川", "file": file, "days": days}},
                "no_empid_names": []}

    def test_same_content_first_wins(self):
        merged, errors = merge_grids([self._grid({1: "BBS3"}, "a.csv"),
                                      self._grid({1: "BBS3"}, "b.csv")])
        assert errors == []
        assert merged["2009006"]["file"] == "a.csv"

    def test_conflict_reported(self):
        merged, errors = merge_grids([self._grid({1: "BBS3"}, "a.csv"),
                                      self._grid({1: "所"}, "b.csv")])
        assert len(errors) == 1
        assert "a.csv" in errors[0] and "b.csv" in errors[0]


class TestBuildTemplateIndex:
    def test_break1_column_name_quirk(self):
        """雛形CSVは休憩1だけ「休憩開始時間1」、2〜5は「休憩時間N」という列名"""
        tpls = [{
            "＊スケジュール雛形ID": "BBS3",
            "＊スケジュール雛形名": "9:00~18:00",
            "＊出勤時間(0:00~47:59)": "09:00:00",
            "＊退勤時間(0:00~47:59)": "18:00:00",
            "休憩開始時間1(0:00~47:59)": "12:00:00",
            "復帰時間1(0:00~47:59)": "13:00:00",
            "休憩時間2(0:00~47:59)": "",
            "復帰時間2(0:00~47:59)": "",
        }]
        idx = build_template_index(tpls)
        assert idx["BBS3"]["start"] == "9:00"
        assert idx["BBS3"]["end"] == "18:00"
        assert idx["BBS3"]["breaks"] == [("12:00", "13:00")]

    def test_old_format_aliases(self):
        """旧フォーマット（勤怠スケジュール名ID / 出勤時刻）も _tpl_get 経由で読める"""
        tpls = [{
            "勤怠スケジュール名ID": "K9",
            "勤怠スケジュール名称": "テスト",
            "出勤時刻(0:00~47:59)": "08:00:00",
            "退勤時刻(0:00~47:59)": "16:00:00",
        }]
        idx = build_template_index(tpls)
        assert idx["K9"]["start"] == "8:00"
        assert idx["K9"]["end"] == "16:00"
        assert idx["K9"]["breaks"] == []

    def test_breaks_sorted_and_overnight(self):
        tpls = [{
            "＊スケジュール雛形ID": "K1",
            "＊スケジュール雛形名": "16:45~33:30",
            "＊出勤時間(0:00~47:59)": "16:45:00",
            "＊退勤時間(0:00~47:59)": "33:30:00",
            "休憩開始時間1(0:00~47:59)": "24:00:00",
            "復帰時間1(0:00~47:59)": "25:45:00",
        }]
        idx = build_template_index(tpls)
        assert idx["K1"]["breaks"] == [("24:00", "25:45")]

    def test_no_id_skipped(self):
        assert build_template_index([{"＊スケジュール雛形名": "IDなし"}]) == {}


class TestFindUnknownCells:
    def test_rest_markers_not_flagged(self):
        """'0' は手作りグリッドの明け休表記（KDX 140-160の2026-07実データ）"""
        employees = {"1": {"name": "a", "file": "f",
                           "days": {1: "所", 2: "法", 3: "休み", 4: "", 5: "休", 6: "0"}}}
        assert find_unknown_cells(employees, {}) == []

    def test_unknown_listed(self):
        employees = {"1": {"name": "a", "file": "f", "days": {1: "BBS3", 2: "XX"}}}
        assert find_unknown_cells(employees, {"BBS3": {}}) == [("XX", "1")]


def _tpl_index():
    return {
        "BBS3": {"name": "9:00~18:00", "start": "9:00", "end": "18:00",
                 "breaks": [("12:00", "13:00")]},
        "N": {"name": "9:00~17:30休憩なし", "start": "9:00", "end": "17:30", "breaks": []},
        "K1": {"name": "16:45~33:30", "start": "16:45", "end": "33:30",
               "breaks": [("24:00", "25:45")]},
    }


class TestBuildDiffPlan:
    def _run(self, days, current, dayoffs=None, groups=None, employees_extra=None,
             exclude=None, month=7):
        employees = {"E1": {"name": "尾川", "file": "grid.csv", "days": days}}
        employees.update(employees_extra or {})
        return build_diff_plan(
            employees, _tpl_index(),
            {"E1": current},
            {"E1": dayoffs or {}},
            groups if groups is not None else {"E1": ("40", "140-180時間制")},
            year=2026, month=month, exclude_emps=exclude,
        )

    def test_full_scenario(self):
        days = {1: "BBS3", 2: "所", 3: "BBS3", 4: "所", 5: "BBS3", 6: "N", 7: "BBS3"}
        current = {
            "2026-07-01": {"start": "9:00", "end": "17:30",
                           "breaks": [("12:00", "13:00")]},   # 修正（終業差）
            "2026-07-02": {"start": "9:00", "end": "17:30",
                           "breaks": [("12:00", "13:00")]},   # 休日に予定残存
            "2026-07-03": {"start": "09:00", "end": "18:00",
                           "breaks": [("12:00", "13:00")]},   # 一致（表記差はt2mで吸収）
            "2026-07-04": {"start": "9:00", "end": "17:30",
                           "breaks": [("12:00", "13:00")]},   # 休日×休暇登録→削除しないで確認
            # 7/5 行なし → 新規
            "2026-07-06": {"start": "9:00", "end": "17:30",
                           "breaks": [("12:00", "13:00")]},   # 休憩差のみ→修正
            "2026-07-07": {"start": "9:00", "end": "17:30",
                           "breaks": [("12:00", "13:00")]},   # 勤務×休暇登録→スキップ
        }
        dayoffs = {"2026-07-04": "年次有休/status=1", "2026-07-07": "年次有休/status=1"}
        r = self._run(days, current, dayoffs)

        kinds = {p["date_iso"]: p["kind"] for p in r.plan}
        assert kinds == {"2026-07-01": "修正", "2026-07-05": "新規", "2026-07-06": "修正"}
        assert r.matched == {"E1": 1}

        kubun = {m["日付"]: m["区分"] for m in r.manual}
        assert kubun["2026-07-02"] == KUBUN_DELETE_NEEDED
        assert kubun["2026-07-04"] == KUBUN_DAYOFF_REST
        assert kubun["2026-07-07"] == KUBUN_DAYOFF_SKIP

    def test_break_only_diff_is_plan_row(self):
        """開始終了が同じでも休憩が違えば書込対象（休憩あり→なし雛形）"""
        r = self._run({1: "N"},
                      {"2026-07-01": {"start": "9:00", "end": "17:30",
                                      "breaks": [("12:00", "13:00")]}})
        assert len(r.plan) == 1
        assert r.plan[0]["breaks"] == []

    def test_overnight_match(self):
        r = self._run({1: "K1"},
                      {"2026-07-01": {"start": "16:45", "end": "33:30",
                                      "breaks": [("24:00", "25:45")]}})
        assert r.plan == []
        assert r.matched == {"E1": 1}

    def test_no_group_all_days_skipped(self):
        r = self._run({1: "BBS3"}, {}, groups={"E1": ("", "")})
        assert r.plan == []
        assert [m["区分"] for m in r.manual] == [KUBUN_NO_GROUP]

    def test_exclude(self):
        r = self._run({1: "BBS3"}, {}, exclude={"E1"})
        assert r.plan == []
        assert r.manual == []
        assert any("除外" in w for w in r.warnings)

    def test_future_month_allowed(self):
        """スケジュール登録は翌月分が本命 → 未来月でも普通にプランが立つこと"""
        r = self._run({1: "BBS3"}, {}, month=12)
        assert len(r.plan) == 1
        assert r.plan[0]["date_iso"] == "2026-12-01"

    def test_plan_row_fields(self):
        r = self._run({5: "BBS3"}, {})
        p = r.plan[0]
        assert p["emp"] == "E1"
        assert p["date_iso"] == "2026-07-05"
        assert p["day"] == 5
        assert p["cell"] == "BBS3"
        assert p["store_id"] == "40"
        assert p["kind"] == "新規"
        assert p["cur"] == "(行なし)"


class TestPlanFingerprint:
    def _row(self, **over):
        base = {"emp": "E1", "date_iso": "2026-07-01", "start": "9:00", "end": "18:00",
                "breaks": [("12:00", "13:00")], "store_id": "40"}
        base.update(over)
        return base

    def test_order_independent(self):
        a = [self._row(), self._row(date_iso="2026-07-02")]
        b = [self._row(date_iso="2026-07-02"), self._row()]
        assert plan_fingerprint(a) == plan_fingerprint(b)

    def test_changes_on_any_field(self):
        assert plan_fingerprint([self._row()]) != plan_fingerprint([self._row(end="17:30")])
        assert plan_fingerprint([self._row()]) != plan_fingerprint([self._row(breaks=[])])
        assert plan_fingerprint([]) != plan_fingerprint([self._row()])


class TestBuildImportRows:
    def _plan_row(self, **over):
        base = {"emp": "2009006", "name": "尾川", "date_iso": "2026-07-06", "day": 6,
                "youbi": "月", "cell": "BBS3", "tpl_name": "9:00~18:00",
                "start": "9:00", "end": "18:00", "breaks": [("12:00", "13:00")],
                "cur": "(行なし)", "kind": "新規",
                "store_id": "40", "store_name": "140-180時間制"}
        base.update(over)
        return base

    def test_columns(self):
        rows = build_import_rows([self._plan_row()])
        assert len(rows) == 1
        row = rows[0]
        assert len(row) == len(GENERIC_IMPORT_HEADER) == 194
        h = GENERIC_IMPORT_HEADER
        assert row[h.index("名前")] == "尾川"
        assert row[h.index("*従業員ID")] == "2009006"
        assert row[h.index("*年月日")] == "2026/7/6"  # ゼロ埋めなし・秒なし
        assert row[h.index("*打刻グループID")] == "40"
        assert row[h.index("出勤予定時刻")] == "9:00"
        assert row[h.index("退勤予定時刻")] == "18:00"
        assert row[h.index("休憩予定時刻1")] == "12:00"
        assert row[h.index("復帰予定時刻1")] == "13:00"
        assert row[h.index("休憩予定時刻2")] == ""
        # スケジュール以外の列は空欄＝無処理（打刻・休暇に触れない）
        assert row[h.index("出勤1")] == ""
        assert row[h.index("休日休暇名1")] == ""
        assert row[h.index("スケジュール雛形ID")] == ""

    def test_sorted_by_emp_and_date(self):
        rows = build_import_rows([
            self._plan_row(emp="2", date_iso="2026-07-02", day=2),
            self._plan_row(emp="1", date_iso="2026-07-03", day=3),
            self._plan_row(emp="1", date_iso="2026-07-01", day=1),
        ])
        h = GENERIC_IMPORT_HEADER
        keys = [(r[h.index("*従業員ID")], r[h.index("*年月日")]) for r in rows]
        assert keys == [("1", "2026/7/1"), ("1", "2026/7/3"), ("2", "2026/7/2")]


class TestRowsToCsvBytes:
    def test_cp932_crlf(self):
        data = rows_to_csv_bytes(["名前", "*従業員ID"], [["尾川", "2009006"]])
        assert b"\r\n" in data
        text = data.decode("cp932")
        assert text.splitlines()[0] == "名前,*従業員ID"
        assert "尾川" in text


class TestVerifyPlanRows:
    def _plan_row(self, **over):
        base = {"emp": "E1", "name": "尾川", "date_iso": "2026-07-01",
                "start": "9:00", "end": "17:30", "breaks": []}
        base.update(over)
        return base

    def test_all_ok(self):
        verify, ng = verify_plan_rows(
            [self._plan_row()],
            {"E1": {"2026-07-01": {"start": "09:00", "end": "17:30", "breaks": []}}})
        assert ng == []
        assert verify[0]["判定"] == "OK"

    def test_missing_row_ng(self):
        verify, ng = verify_plan_rows([self._plan_row()], {"E1": {}})
        assert len(ng) == 1
        assert "(行なし)" in ng[0]["備考"]

    def test_start_mismatch_ng(self):
        verify, ng = verify_plan_rows(
            [self._plan_row()],
            {"E1": {"2026-07-01": {"start": "9:15", "end": "17:30", "breaks": []}}})
        assert len(ng) == 1

    def test_empty_breaks_strict(self):
        """期待休憩なし × 現物に休憩あり → NG（日丸ごと置換仕様の厳密検証）"""
        verify, ng = verify_plan_rows(
            [self._plan_row()],
            {"E1": {"2026-07-01": {"start": "9:00", "end": "17:30",
                                   "breaks": [("12:00", "13:00")]}}})
        assert len(ng) == 1

    def test_breaks_match(self):
        verify, ng = verify_plan_rows(
            [self._plan_row(breaks=[("12:00", "13:00")])],
            {"E1": {"2026-07-01": {"start": "9:00", "end": "17:30",
                                   "breaks": [("12:00:00", "13:00:00")]}}})
        assert ng == []


class TestGenericHeaderConstant:
    def test_matches_bundled_template(self):
        """埋め込み194列ヘッダーが同梱テンプレートCSVからドリフトしていないこと"""
        import csv
        path = (Path(__file__).resolve().parent.parent / "汎用データテンプレート"
                / "汎用データ(まるめ適用後)ダウンロード_9637_20260522150030.csv")
        if not path.exists():
            pytest.skip(f"テンプレートCSVなし: {path}")
        with open(path, encoding="cp932", newline="") as f:
            header = next(csv.reader(f))
        assert tuple(header) == GENERIC_IMPORT_HEADER
