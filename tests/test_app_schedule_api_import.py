# -*- coding: utf-8 -*-
"""/schedule_api_import ルートのテスト（ガードと手順③との相互ロック）"""
import json
import time
from pathlib import Path

import pytest

import app as app_module
from config import Config
from services.schedule_import_runner import ScheduleImportResult


@pytest.fixture
def client():
    app_module.app.config["TESTING"] = True
    # ジョブ辞書はモジュールグローバルなのでテスト間で必ず掃除する
    app_module._api_import_jobs.clear()
    with app_module.app.test_client() as c:
        yield c
    app_module._api_import_jobs.clear()


@pytest.fixture
def dummy_grid(tmp_path):
    """OUTPUT_FOLDER 内に存在チェックを通すだけのダミーグリッドCSVを置く"""
    out = Path(Config.OUTPUT_FOLDER)
    out.mkdir(parents=True, exist_ok=True)
    p = out / "_test_dummy_grid_for_route.csv"
    p.write_bytes("2026年,7月,1\n氏名,従業員ID,水\nテスト,9999999,所\n".encode("cp932"))
    yield p.name
    p.unlink(missing_ok=True)


def _post(client, **over):
    form = {"csv_filenames": json.dumps(["x.csv"]), "month": "2026-08", "execute": "0"}
    form.update(over)
    return client.post("/schedule_api_import", data=form)


class TestGuards:
    def test_no_files_400(self, client):
        res = _post(client, csv_filenames="[]")
        assert res.status_code == 400

    def test_bad_json_400(self, client):
        res = _post(client, csv_filenames="not-json")
        assert res.status_code == 400

    def test_bad_month_400(self, client):
        res = _post(client, month="2026/08")
        assert res.status_code == 400

    def test_execute_without_fingerprint_400(self, client, dummy_grid):
        res = _post(client, csv_filenames=json.dumps([dummy_grid]), execute="1")
        assert res.status_code == 400
        assert "fingerprint" in res.get_json()["errors"][0]

    def test_missing_file_400(self, client):
        res = _post(client, csv_filenames=json.dumps(["存在しない.csv"]))
        assert res.status_code == 400

    def test_path_traversal_stripped(self, client):
        """basename 強制により OUTPUT_FOLDER の外は参照できない"""
        res = _post(client, csv_filenames=json.dumps(["..\\..\\.env"]))
        assert res.status_code == 400  # basename後 ".env" は outputs に無い → 見つからない


class TestMutualLock:
    def test_409_when_other_import_running(self, client, dummy_grid):
        """手順③(/api_import)のジョブ実行中はスケジュール投入も409（同時予約1件）"""
        app_module._api_import_jobs["running-job"] = {
            "done": False, "ok": False, "log": [], "result": None}
        res = _post(client, csv_filenames=json.dumps([dummy_grid]))
        assert res.status_code == 409


class TestJobFlow:
    def test_dry_run_job_returns_result(self, client, dummy_grid, monkeypatch):
        """workerがrun_schedule_api_importの結果をjob.resultへ詰めること（本体はスタブ）"""
        stub_result = ScheduleImportResult(
            ok=True, dry_run=True, month="2026-08",
            plan_rows=2, matched_rows=5,
            plan=[{"emp": "1", "name": "テスト", "date_iso": "2026-08-01", "day": 1,
                   "youbi": "土", "cell": "BBS3", "tpl_name": "9:00~18:00",
                   "start": "9:00", "end": "18:00", "breaks": [("12:00", "13:00")],
                   "cur": "(行なし)", "kind": "新規",
                   "store_id": "40", "store_name": "140-180時間制"}],
            manual=[], fingerprint="abc123", report_path="outputs/dummy.xlsx",
        )
        monkeypatch.setattr(app_module, "run_schedule_api_import",
                            lambda **kw: stub_result)

        res = _post(client, csv_filenames=json.dumps([dummy_grid]))
        assert res.status_code == 200
        job_id = res.get_json()["job_id"]

        deadline = time.time() + 5
        status = None
        while time.time() < deadline:
            status = client.get(f"/api_import_status/{job_id}").get_json()
            if status["done"]:
                break
            time.sleep(0.05)
        assert status is not None and status["done"]
        assert status["ok"] is True
        r = status["result"]
        assert r["mode"] == "schedule"
        assert r["dry_run"] is True
        assert r["plan_rows"] == 2
        assert r["fingerprint"] == "abc123"
        assert r["plan_preview"][0]["new"] == "9:00-18:00"
        assert r["report_url"] == "/download/dummy.xlsx"
