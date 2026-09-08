# -*- coding: utf-8 -*-
"""標準報酬チェックの「不足月の取得」。

exe は各PCのローカルにキャッシュを持つので、開発PCで取った 4〜6月分は他のPCからは見えず、
「2026-04 の給与明細キャッシュがありません」で止まる（2026-09-08 実機）。ここで見張るのは:

  1. 無い月は **全部まとめて** 知らせる（1か月ずつ「取っては落ちる」を繰り返させない）
  2. 取得フラグ無しでは API に一切触れない（「読むだけ」の約束）
  3. 取得フラグ有りなら無い月だけ jinjer から取り、経理モードと同じ場所にキャッシュを残す
  4. 画面のチェック・CLI のフラグが run_check まで届く
"""

from __future__ import annotations

import json
import os
import sys
import types

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from config import Config  # noqa: E402
from services import shaho_check  # noqa: E402
from services.jinjer_api_client import JinjerAPIError  # noqa: E402
from services.shaho_master import ShahoMasterError  # noqa: E402


def _person(emp="2020001", closed=True):
    return {"employee_id": emp,
            "statements": [{"basic_info": {"last_name": "山田", "first_name": "太郎"},
                            "payroll_info": {"is_payroll_closed": closed,
                                             "salary_items": [{"id": "allowance1",
                                                               "value": "250000"}]}}]}


class FakeClient:
    """jinjer の代わり。呼ばれた月を記録して1人分を返すだけ。"""

    def __init__(self):
        self.calls = []

    def get_salary_statements(self, executed_on, employee_ids=None):
        self.calls.append(executed_on)
        return [_person()]


@pytest.fixture
def raw_dir(tmp_path, monkeypatch):
    monkeypatch.setattr(Config, "KEIRI_OUTPUT_DIR", str(tmp_path / "keiri"))
    path = tmp_path / "keiri" / "raw"
    path.mkdir(parents=True)
    return path


@pytest.fixture
def client():
    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as c:
        yield c


def _cache(raw_dir, ym):
    with open(raw_dir / f"salary_statements_{ym}.json", "w", encoding="utf-8") as f:
        json.dump({"executed_on": ym, "data": [_person()]}, f, ensure_ascii=False)


def _stub_masters(monkeypatch):
    """等級表・分類マスタ（共有フォルダの実ファイル）を読まずに run_check を通す。"""
    monkeypatch.setattr(shaho_check, "select_grade_table",
                        lambda ym, base: types.SimpleNamespace(path="dummy.xlsx"))
    monkeypatch.setattr(shaho_check, "load_grade_table", lambda *a, **k: None)
    monkeypatch.setattr(shaho_check, "load_class_master", lambda *a, **k: None)


_REPORT = {"n": 0, "review_n": 0, "open_months": {}, "xlsx": "x.xlsx", "json": "x.json"}


# ---------------------------------------------------------------- load_month_caches
def test_missing_months_are_reported_together(raw_dir):
    _cache(raw_dir, "2026-06")
    _cache(raw_dir, "2026-08")
    with pytest.raises(ShahoMasterError) as ei:
        shaho_check.load_month_caches(["2026-04", "2026-05", "2026-06", "2026-08", "2026-09"])
    msg = str(ei.value)
    assert "2026-04、2026-05、2026-09" in msg          # 無い月だけを、まとめて
    assert "2026-06" not in msg and "2026-08" not in msg
    assert str(raw_dir) in msg                        # どこを見たかも示す
    assert "--fetch-missing" in msg


def test_without_flag_never_touches_api(raw_dir):
    fake = FakeClient()
    with pytest.raises(ShahoMasterError):
        shaho_check.load_month_caches(["2026-04"], fetch_missing=False, client=fake)
    assert fake.calls == []


def test_fetch_missing_fetches_only_missing_months_and_caches_them(raw_dir):
    _cache(raw_dir, "2026-06")
    fake = FakeClient()
    out = shaho_check.load_month_caches(["2026-04", "2026-05", "2026-06", "2026-06"],
                                        fetch_missing=True, client=fake)
    assert fake.calls == ["2026-04", "2026-05"]       # 有る月・重複月は取らない
    assert sorted(out) == ["2026-04", "2026-05", "2026-06"]
    assert out["2026-04"]["2020001"]["payroll_info"]["is_payroll_closed"] is True
    # 取った月は経理モードと同じ場所に残り、次回は API を呼ばない
    assert (raw_dir / "salary_statements_2026-04.json").exists()
    again = shaho_check.load_month_caches(["2026-04", "2026-05"], fetch_missing=True,
                                          client=fake)
    assert fake.calls == ["2026-04", "2026-05"]
    assert sorted(again) == ["2026-04", "2026-05"]


def test_run_check_needs_four_to_six_and_check_pair_and_forwards_flag(monkeypatch):
    seen = {}

    class Stop(Exception):
        pass

    def fake_loader(months, fetch_missing=False, client=None):
        seen["months"] = list(months)
        seen["fetch_missing"] = fetch_missing
        raise Stop()

    _stub_masters(monkeypatch)
    monkeypatch.setattr(shaho_check, "load_month_caches", fake_loader)
    with pytest.raises(Stop):
        shaho_check.run_check(2026, "2026-08", grade_xlsx="dummy.xlsx",
                              class_csv="dummy.csv", fetch_missing=True)
    assert seen["fetch_missing"] is True
    assert seen["months"] == ["2026-04", "2026-05", "2026-06", "2026-07", "2026-08"]


# ---------------------------------------------------------------- 画面（/shaho_run）
def test_route_passes_checkbox_to_run_check(client, monkeypatch, raw_dir):
    seen = {}

    def fake_run_check(year, check_month, **kw):
        seen.clear()
        seen.update(year=year, check_month=check_month, **kw)
        return {"results": []}

    monkeypatch.setattr("services.shaho_check.run_check", fake_run_check)
    monkeypatch.setattr("services.shaho_report.write_reports", lambda check: dict(_REPORT))
    r = client.post("/shaho_run", data={"year": "2026", "check_month": "2026-08",
                                        "fetch_missing": "1"})
    assert r.status_code == 200, r.get_json()
    assert seen == {"year": 2026, "check_month": "2026-08", "fetch_missing": True}
    r = client.post("/shaho_run", data={"year": "2026", "check_month": "2026-08"})
    assert r.status_code == 200, r.get_json()
    assert seen["fetch_missing"] is False


def test_route_reports_missing_months_as_400(client, monkeypatch, raw_dir):
    """キャッシュ不足は入力の問題（400）で、不足月を全部並べる。"""
    _stub_masters(monkeypatch)
    _cache(raw_dir, "2026-08")
    r = client.post("/shaho_run", data={"year": "2026", "check_month": "2026-08"})
    assert r.status_code == 400
    err = r.get_json()["errors"][0]
    assert "2026-04、2026-05、2026-06、2026-07" in err
    assert "jinjer から取得する" in err


def test_route_without_cache_and_blank_month_asks_for_month(client, raw_dir):
    r = client.post("/shaho_run", data={"year": "", "check_month": "", "fetch_missing": "1"})
    assert r.status_code == 400
    assert "突合月" in r.get_json()["errors"][0]


def test_route_turns_api_error_into_message(client, monkeypatch, raw_dir):
    def boom(*a, **k):
        raise JinjerAPIError("429 Too Many Requests")

    monkeypatch.setattr("services.shaho_check.run_check", boom)
    r = client.post("/shaho_run", data={"year": "2026", "check_month": "2026-08",
                                        "fetch_missing": "1"})
    assert r.status_code == 500
    assert "jinjer API" in r.get_json()["errors"][0]


# ---------------------------------------------------------------- CLI（shaho_check_run.py）
def test_cli_flag_reaches_run_check(monkeypatch, raw_dir):
    import shaho_check_run
    seen = {}

    def fake_run_check(year, check_month, **kw):
        seen.update(kw)
        return {"results": []}

    monkeypatch.setattr(shaho_check_run, "run_check", fake_run_check)
    monkeypatch.setattr(shaho_check_run, "write_reports", lambda check: dict(_REPORT))
    monkeypatch.setattr(sys, "argv", ["shaho_check_run.py", "--check-month", "2026-08",
                                      "--fetch-missing"])
    assert shaho_check_run.main() == 0
    assert seen["fetch_missing"] is True


def test_cli_without_any_cache_explains_instead_of_crashing(tmp_path, monkeypatch, capsys):
    """raw フォルダ自体が無い PC（初回）でも os.listdir で落ちず、次の一手を出す。"""
    import shaho_check_run
    monkeypatch.setattr(Config, "KEIRI_OUTPUT_DIR", str(tmp_path / "nowhere"))
    monkeypatch.setattr(sys, "argv", ["shaho_check_run.py"])
    assert shaho_check_run.main() == 1
    out = capsys.readouterr().out
    assert "--check-month" in out and "--fetch-missing" in out
