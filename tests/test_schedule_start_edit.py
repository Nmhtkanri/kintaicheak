# -*- coding: utf-8 -*-
"""スケジュール開始時刻編集（services/schedule_start_edit.py）のテスト。

書き込み系なので、罠の再現を最重視する:
  - 開始だけ差し替えて 終了・休憩がフルセットで送られること（丸ごと置換対策）
  - 休暇日が送信前に弾かれること（サイレント無視対策）
  - fingerprint 不一致で送信しないこと（承認後変更の保険）
"""
import csv
import io
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import services.schedule_start_edit as sse
from services.schedule_import_runner import GENERIC_IMPORT_HEADER
from services.schedule_start_edit import (
    ST_CHANGE,
    ST_SAME,
    build_plan,
    parse_edit_lines,
    run_schedule_start_edit,
)


# ---------------------------------------------------------------------------
# parse_edit_lines
# ---------------------------------------------------------------------------

def test_parse_accepts_comma_tab_space_and_date_styles():
    text = "2020001,2026/8/1,9:30\n2020002\t2026-08-02\t10:00\n2020003 2026/8/3 8:15:00\n"
    edits, errors = parse_edit_lines(text)
    assert errors == []
    assert [(e["emp"], e["date_iso"], e["new_start"]) for e in edits] == [
        ("2020001", "2026-08-01", "9:30"),
        ("2020002", "2026-08-02", "10:00"),
        ("2020003", "2026-08-03", "8:15"),   # 秒は落とす
    ]


@pytest.mark.parametrize("line,frag", [
    ("2020001,2026/8/1", "3項目"),
    ("abc,2026/8/1,9:00", "社員番号"),
    ("2020001,8/1,9:00", "日付"),
    ("2020001,2026/2/30,9:00", "実在しない"),
    ("2020001,2026/8/1,930", "H:MM"),
    ("2020001,2026/8/1,24:00", "0:00〜23:59"),
])
def test_parse_rejects_bad_lines(line, frag):
    edits, errors = parse_edit_lines(line)
    assert edits == []
    assert any(frag in e for e in errors)


def test_parse_rejects_duplicate_emp_date():
    edits, errors = parse_edit_lines("2020001,2026/8/1,9:00\n2020001,2026/8/1,10:00")
    assert len(edits) == 1
    assert any("重複" in e for e in errors)


# ---------------------------------------------------------------------------
# build_plan
# ---------------------------------------------------------------------------

CUR = {"start": "9:00", "end": "18:00", "breaks": [("12:00", "13:00")],
       "store": "時給制", "store_id": "40"}


def _ctx(**over):
    ctx = {
        "schedules": {("2020001", "2026-08"): {"2026-08-01": dict(CUR)}},
        "day_offs": {("2020001", "2026-08"): {}},
        "names": {"2020001": "山田"},
        "groups": {("2020001", "2026-08-01"): ("40", "時給制")},
    }
    ctx.update(over)
    return ctx


def _edit(new="9:30"):
    edits, errors = parse_edit_lines(f"2020001,2026/8/1,{new}")
    assert errors == []
    return edits


def test_build_plan_change_preserves_end_and_breaks():
    plan, preview, errors = build_plan(_edit("9:30"), **_ctx())
    assert errors == []
    assert len(plan) == 1
    p = plan[0]
    assert (p["start"], p["end"], p["breaks"]) == ("9:30", "18:00", [("12:00", "13:00")])
    assert p["store_id"] == "40"
    assert preview[0]["状態"] == ST_CHANGE
    assert preview[0]["現在の開始"] == "9:00"


def test_build_plan_same_start_skips():
    plan, preview, errors = build_plan(_edit("9:00"), **_ctx())
    assert plan == [] and errors == []
    assert preview[0]["状態"] == ST_SAME


def test_build_plan_blocks_day_off():
    ctx = _ctx(day_offs={("2020001", "2026-08"): {"2026-08-01": "年次有休(全休)"}})
    plan, _preview, errors = build_plan(_edit(), **ctx)
    assert plan == []
    assert any("休暇" in e and "サイレント" in e for e in errors)


def test_build_plan_requires_existing_schedule():
    ctx = _ctx(schedules={("2020001", "2026-08"): {}})
    plan, _p, errors = build_plan(_edit(), **ctx)
    assert plan == []
    assert any("スケジュールがありません" in e for e in errors)


def test_build_plan_rejects_start_after_end():
    plan, _p, errors = build_plan(_edit("18:00"), **_ctx())
    assert plan == []
    assert any("退勤予定" in e for e in errors)


def test_build_plan_unknown_employee():
    plan, _p, errors = build_plan(_edit(), **_ctx(names={}))
    assert plan == []
    assert any("見つかりません" in e for e in errors)


# ---------------------------------------------------------------------------
# run_schedule_start_edit（fake client で end-to-end）
# ---------------------------------------------------------------------------

class FakeClient:
    def __init__(self):
        self.sched = {"2026-08-01": dict(CUR)}
        self.posted: list[tuple[str, bytes]] = []

    def get_work_schedules(self, emp, month, store_id=""):
        assert store_id == "40"   # 現グループで絞っていること（残骸対策）
        return {d: dict(v) for d, v in self.sched.items()}

    def get_requested_day_offs(self, emp, month):
        return {}

    def post_kintai_import(self, csv_bytes, file_name, executor_id=None):
        self.posted.append((file_name, csv_bytes))
        # 投入されたら jinjer 側の開始が変わる（反映検証がこれを読む）
        rows = list(csv.reader(io.StringIO(csv_bytes.decode("cp932"))))
        hdr = rows[0]
        for r in rows[1:]:
            self.sched[_iso(r[hdr.index("*年月日")])]["start"] = r[hdr.index("出勤予定時刻")]
        return {}


def _iso(ymd):
    y, m, d = ymd.split("/")
    return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"


@pytest.fixture
def patched(monkeypatch):
    monkeypatch.setattr(sse, "fetch_employee_id_map",
                        lambda: ({}, {"2020001": "山田"}, {}))
    monkeypatch.setattr(sse, "fetch_attendance_groups_at",
                        lambda emps, d: {e: ("40", "時給制") for e in emps})
    return FakeClient()


def test_dry_run_returns_preview_and_fingerprint(patched, tmp_path):
    r = run_schedule_start_edit("2020001,2026/8/1,9:30", dry_run=True,
                                output_dir=tmp_path, client=patched)
    assert r.ok and r.dry_run
    assert r.change_count == 1 and r.fingerprint
    assert patched.posted == []          # dry-run では送らない
    assert r.snapshot_path == ""         # ログも実行時のみ


def test_execute_requires_matching_fingerprint(patched, tmp_path):
    r = run_schedule_start_edit("2020001,2026/8/1,9:30", dry_run=False,
                                expected_fingerprint="deadbeef",
                                output_dir=tmp_path, client=patched)
    assert not r.ok
    assert patched.posted == []
    assert any("fingerprint" in e for e in r.errors)


def test_execute_sends_full_set_and_verifies(patched, tmp_path):
    pre = run_schedule_start_edit("2020001,2026/8/1,9:30", dry_run=True,
                                  output_dir=tmp_path, client=patched)
    r = run_schedule_start_edit("2020001,2026/8/1,9:30", dry_run=False,
                                expected_fingerprint=pre.fingerprint,
                                output_dir=tmp_path, client=patched,
                                poll_func=lambda cli, fn, log: "1")
    assert r.ok, r.errors
    assert r.import_status == "1" and r.verify_ng == []
    # 送信CSVの中身: 開始が差し替わり、終了・休憩がフルセットで入っている
    rows = list(csv.reader(io.StringIO(patched.posted[0][1].decode("cp932"))))
    hdr, row = rows[0], rows[1]
    assert hdr == list(GENERIC_IMPORT_HEADER)
    get = lambda c: row[hdr.index(c)]
    assert get("出勤予定時刻") == "9:30"
    assert get("退勤予定時刻") == "18:00"
    assert (get("休憩予定時刻1"), get("復帰予定時刻1")) == ("12:00", "13:00")
    assert get("*年月日") == "2026/8/1"       # 0埋めなし
    assert get("名前") == "山田"              # 姓のみ
    assert get("休日（0:法定休日1:所定休日2:法休(振替休出)3:所休(振替休出)4:法休(時間外休出)5:所休(時間外休出)）") == ""
    # 実行ログが残る
    assert r.snapshot_path and Path(r.snapshot_path).exists()
    body = Path(r.snapshot_path).read_text(encoding="utf-8-sig")
    assert "9:30" in body and "山田" in body


# ---------------------------------------------------------------------------
# Flask ルートの入力検証（ジョブは起動させない＝400 で返る経路のみ）
# ---------------------------------------------------------------------------

def test_route_validations():
    import app as app_module
    cli = app_module.app.test_client()
    r = cli.post("/schedule_start_edit", data={"edits": "  "})
    assert r.status_code == 400
    assert "入力されていません" in r.get_json()["errors"][0]
    r = cli.post("/schedule_start_edit", data={"edits": "2020001,2026/8/1,9:30", "execute": "1"})
    assert r.status_code == 400
    assert "fingerprint" in r.get_json()["errors"][0]


def test_execute_abort_when_any_row_has_error(patched, tmp_path):
    """1行でもエラーがあれば全体を送らない（部分書込の混乱防止）。"""
    text = "2020001,2026/8/1,9:30\n2020001,2026/8/2,9:30"   # 8/2 はスケジュール無し
    r = run_schedule_start_edit(text, dry_run=True, output_dir=tmp_path, client=patched)
    assert not r.ok
    assert any("スケジュールがありません" in e for e in r.errors)
    assert patched.posted == []
