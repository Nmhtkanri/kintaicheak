# -*- coding: utf-8 -*-
"""jinjer添付まわりの単体テスト（アダプタ・リトライ二層・fetch_records・run のフック）。

すべて合成データ・fakeクライアント。実APIにも Z:\\派遣元管理台帳 にも触れない。
移設まで jinjer_attach はテストゼロの領域だったため、判定ロジックをここで固定する。
"""
from __future__ import annotations

import datetime as dt
import sys
import time
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services.daicho import attach_job, jinjer_attach  # noqa: E402
from services.daicho import config as daicho_config  # noqa: E402
from services.daicho.hub_client import DaichoAttachClient  # noqa: E402
from services.jinjer_api_client import JinjerAPIError  # noqa: E402


# ---------------------------------------------------------------------------
# hub_client.DaichoAttachClient
# ---------------------------------------------------------------------------

class _FakeWrite:
    """JinjerWriteClient の代役。呼び出しを記録するだけ。"""

    def __init__(self):
        self.base_url = "https://api.example"
        self.timeout = 30
        self._access_token = "old-token"
        self.auth_calls = 0
        self.writes = []

    def authenticate(self):
        self.auth_calls += 1
        if self._access_token is None:
            self._access_token = "new-token"
        return self._access_token

    def _auth_headers(self):
        return {"Authorization": f"Bearer {self.authenticate()}"}

    def _write(self, method, path, body, *, what):
        self.writes.append((method, path, body, what))
        return {"id": "1"}


def _adapter_with_fake():
    client = DaichoAttachClient.__new__(DaichoAttachClient)
    client._c = _FakeWrite()
    return client


def test_adapter_authenticate_forces_token_refresh():
    client = _adapter_with_fake()
    client._c._access_token = "stale"
    client.authenticate()
    # 先にキャッシュを捨ててから取り直す＝失効トークンを使い回さない
    assert client._c._access_token == "new-token"
    assert client._c.auth_calls == 1


def test_adapter_dispatches_writes_to_hub_client():
    client = _adapter_with_fake()
    client._request("POST", "/v1/employees/addible-custom-items", json_body={"a": 1})
    client._request("PATCH", "/v1/async/files", json_body={"b": 2})
    methods = [(m, p) for m, p, _b, _w in client._c.writes]
    assert methods == [("POST", "/v1/employees/addible-custom-items"),
                       ("PATCH", "/v1/async/files")]


def test_adapter_get_returns_data_and_reauths_on_403(monkeypatch):
    client = _adapter_with_fake()
    calls = {"n": 0}

    class _Resp:
        def __init__(self, status, payload=None):
            self.status_code = status
            self.headers = {}
            self._payload = payload or {}

        def json(self):
            return self._payload

    def fake_get(url, headers=None, params=None, timeout=None):
        calls["n"] += 1
        if calls["n"] == 1:
            return _Resp(403)
        return _Resp(200, {"data": [{"employee_id": "1"}]})

    monkeypatch.setattr("services.daicho.hub_client.requests.get", fake_get)
    data = client._request("GET", "/v1/employees/addible-custom-items", params={"page": 1})
    assert data == [{"employee_id": "1"}]
    assert calls["n"] == 2                      # 403 → 再認証して1回だけやり直す


# ---------------------------------------------------------------------------
# _write_with_retry（429を粘る・403token→再認証）
# ---------------------------------------------------------------------------

def test_write_with_retry_survives_429_then_succeeds(monkeypatch):
    slept = []
    monkeypatch.setattr(time, "sleep", lambda s: slept.append(s))
    client = _FakeWrite()
    attempts = {"n": 0}

    def fn():
        attempts["n"] += 1
        if attempts["n"] <= 2:
            raise JinjerAPIError("POST xに失敗 (status=429): 混んでいます")
        return "ok"

    assert jinjer_attach._write_with_retry(client, fn, "テスト") == "ok"
    assert attempts["n"] == 3
    assert slept == [180.0, 180.0]              # cooldown 180秒×2回粘った


def test_write_with_retry_reauths_on_403_token(monkeypatch):
    monkeypatch.setattr(time, "sleep", lambda s: None)
    client = _FakeWrite()
    attempts = {"n": 0}

    def fn():
        attempts["n"] += 1
        if attempts["n"] == 1:
            raise JinjerAPIError("PATCH xに失敗 (status=403): Access token verification failed")
        return "ok"

    assert jinjer_attach._write_with_retry(client, fn, "テスト") == "ok"
    assert client.auth_calls == 1               # 再認証してから再試行した


def test_write_with_retry_raises_other_errors_immediately(monkeypatch):
    monkeypatch.setattr(time, "sleep", lambda s: None)
    client = _FakeWrite()

    def fn():
        raise JinjerAPIError("POST xに失敗 (status=400): 入力が不正")

    with pytest.raises(JinjerAPIError):
        jinjer_attach._write_with_retry(client, fn, "テスト")


# ---------------------------------------------------------------------------
# fetch_records（updated_at≠created_at の添付済み判定・単一dict・ページング）
# ---------------------------------------------------------------------------

class _FakeGetClient:
    def __init__(self, pages):
        self._pages = pages
        self.requests = []

    def _request(self, method, path, json_body=None, params=None):
        assert method == "GET"
        self.requests.append(dict(params or {}))
        page = int((params or {}).get("page", 1))
        return self._pages[page - 1] if page <= len(self._pages) else []


def _record(emp, rec_id, date_val, file_val="", upd=None, cre=None, single_menu=False):
    menu = {"id": "16", "customize_data": [{
        "id": rec_id,
        "customize_item": [
            {"id": "1", "value": date_val},
            {"id": "2", "value": file_val, "updated_at": upd, "created_at": cre},
        ]}]}
    return {"employee_id": emp, "customize_menu": (menu if single_menu else [menu])}


def test_fetch_records_attached_judgment_and_single_dict_menu():
    pages = [[
        _record("1001", "r1", "2026/4/1", upd="t2", cre="t1"),                 # 添付済み（進んだ）
        _record("1001", "r2", "2026/4/1", upd="t1", cre="t1"),                 # 未添付
        _record("1002", "r3", "2026-04-01", file_val="x.pdf", single_menu=True),  # value有り＝添付済み・単一dict
    ]]
    out = jinjer_attach.fetch_records(_FakeGetClient(pages), ["1001", "1002"])
    by_id = {r["id"]: r for rows in out.values() for r in rows}
    assert by_id["r1"]["attached"] is True
    assert by_id["r2"]["attached"] is False
    assert by_id["r3"]["attached"] is True
    assert by_id["r3"]["date"] == "2026-04-01"   # GETのYYYY/M/D表記も正規化される


def test_fetch_records_paginates_until_short_page():
    page1 = [_record(str(1000 + i), f"p{i}", "2026/4/1") for i in range(100)]
    page2 = [_record("2001", "last", "2026/4/1")]
    client = _FakeGetClient([page1, page2])
    out = jinjer_attach.fetch_records(client, [str(1000 + i) for i in range(100)] + ["2001"])
    assert any(r["id"] == "last" for rows in out.values() for r in rows)
    assert [p.get("page") for p in client.requests[:2]] == [1, 2]


# ---------------------------------------------------------------------------
# run() の on_progress / should_stop フック（書き込みはfake）
# ---------------------------------------------------------------------------

def test_scan_pdfs_reads_folders_and_skips_undecided(tmp_path):
    pdf_root = tmp_path / "PDF"
    (pdf_root / "1001_山田太郎").mkdir(parents=True)
    (pdf_root / "1001_山田太郎" / "1001_山田太郎_2026年4-6月分.pdf").write_bytes(b"%PDF a")
    (pdf_root / "未確定_岡田尚美").mkdir()
    (pdf_root / "未確定_岡田尚美" / "未確定_岡田尚美_2026年4-6月分.pdf").write_bytes(b"%PDF b")
    jobs, skipped = jinjer_attach.scan_pdfs(pdf_root=pdf_root)
    assert [(j[0], j[3]) for j in jobs] == [("1001", "2026/4/1")]
    assert skipped and "未確定" in skipped[0]


@pytest.fixture
def attach_env(tmp_path, monkeypatch):
    """PDF 2枚＋jinjer側レコード（1枚は添付済み）の合成環境。

    ⚠ scan_pdfs の pdf_root は**デフォルト引数でimport時に束縛**されるため、
    PDF_ROOT の monkeypatch は効かない（実フォルダを読んでしまう）。
    run() のテストでは scan_pdfs ごと差し替える。
    """
    pdf_root = tmp_path / "PDF"
    folder = pdf_root / "1001_山田太郎"
    folder.mkdir(parents=True)
    pdf1 = folder / "1001_山田太郎_2026年4-6月分.pdf"
    pdf1.write_bytes(b"%PDF-1.4 a")
    folder2 = pdf_root / "1002_佐藤花子"
    folder2.mkdir(parents=True)
    pdf2 = folder2 / "1002_佐藤花子_2026年4-6月分.pdf"
    pdf2.write_bytes(b"%PDF-1.4 b")
    jobs = [("1001", folder.name, pdf1, "2026/4/1"),
            ("1002", folder2.name, pdf2, "2026/4/1")]
    monkeypatch.setattr(jinjer_attach, "scan_pdfs",
                        lambda pdf_root=None, employees=None: (list(jobs), []))
    monkeypatch.setattr(daicho_config, "LOG_DIR", tmp_path / "ログ")
    # jinjer_attach は from .config import で名前を取り込んでいるため、こちらも差し替える
    monkeypatch.setattr(jinjer_attach, "LOG_DIR", tmp_path / "ログ")
    monkeypatch.setattr(jinjer_attach, "LOG_PATH", tmp_path / "ログ" / "jinjer添付ログ.txt")

    records = {"1001": [{"id": "r1", "date": "2026-04-01", "attached": True}],   # 済み
               "1002": [{"id": "r2", "date": "2026-04-01", "attached": False}]}  # 空きレコードあり
    monkeypatch.setattr(jinjer_attach, "_client", lambda interval: _FakeWrite())
    monkeypatch.setattr(jinjer_attach, "fetch_records", lambda c, ids: records)
    attached = []
    monkeypatch.setattr(jinjer_attach, "attach_file",
                        lambda c, emp, rec_id, pdf: attached.append((emp, rec_id, pdf.name)))
    return attached


def test_run_skips_attached_and_reports_progress(attach_env):
    progress = []
    done, skip, total = jinjer_attach.run(
        dry_run=False, on_progress=lambda d, s, t, cur: progress.append((d, s, t, cur)))
    assert (done, skip, total) == (1, 1, 2)      # 済み1をスキップ・空きレコードへ1枚添付
    assert attach_env == [("1002", "r2", "1002_佐藤花子_2026年4-6月分.pdf")]
    assert progress[-1] == (1, 1, 2, "")


def test_run_stops_on_should_stop(attach_env):
    done, skip, total = jinjer_attach.run(dry_run=False, should_stop=lambda: True)
    assert done == 0                             # 1件目の前に止まる
    assert attach_env == []


# ---------------------------------------------------------------------------
# attach_job（進捗ファイル・開始時刻・キャンセル）
# ---------------------------------------------------------------------------

@pytest.fixture
def job_paths(tmp_path, monkeypatch):
    monkeypatch.setattr(daicho_config, "LOG_DIR", tmp_path / "ログ")
    monkeypatch.setattr(daicho_config, "ATTACH_PROGRESS_JSON", tmp_path / "ログ" / "添付進捗.json")
    monkeypatch.setattr(daicho_config, "ATTACH_CANCEL_FLAG", tmp_path / "ログ" / "添付キャンセル.flag")
    from config import Config
    monkeypatch.setattr(Config, "HAKEN_ATTACH_LOCK_FILE", str(tmp_path / "own.lock"))
    monkeypatch.setattr(Config, "SHAHO_IMPORT_LOCK_FILE", str(tmp_path / "shaho.lock"))
    return tmp_path


def test_progress_roundtrip_and_unknown_on_broken(job_paths):
    attach_job.write_progress({"state": "running", "done": 3})
    p = attach_job.read_progress()
    assert p["state"] == "running" and p["done"] == 3 and p["updated_at"]
    daicho_config.ATTACH_PROGRESS_JSON.write_text("{broken", encoding="utf-8")
    assert attach_job.read_progress()["state"] == "unknown"


def test_parse_start_at():
    future = (dt.datetime.now() + dt.timedelta(hours=1)).strftime("%H:%M")
    past = (dt.datetime.now() - dt.timedelta(hours=1)).strftime("%H:%M")
    assert attach_job._parse_start_at(None) is None
    assert attach_job._parse_start_at(past) is None          # 過ぎていれば即時開始
    assert attach_job._parse_start_at(future) is not None
    with pytest.raises(attach_job.HakenAttachError):
        attach_job._parse_start_at("25時")


def test_scheduled_job_cancels_before_start(job_paths, monkeypatch):
    future = (dt.datetime.now() + dt.timedelta(hours=2)).strftime("%H:%M")
    attach_job.request_cancel()
    monkeypatch.setattr(time, "sleep", lambda s: None)
    code = attach_job.run_attach_job(execute=True, start_at=future)
    assert code == 0
    assert attach_job.read_progress()["state"] == "cancelled"
    assert not daicho_config.ATTACH_CANCEL_FLAG.exists()     # フラグは消費される


def test_job_waits_while_shaho_lock_alive(job_paths, monkeypatch):
    """標報投入ロックが生きている間は running に入らない（1周だけ回して確認）。"""
    import json as _json
    from config import Config

    Path(Config.SHAHO_IMPORT_LOCK_FILE).write_text(
        _json.dumps({"user": "他モード", "started_at": dt.datetime.now().isoformat()}),
        encoding="utf-8")
    calls = {"n": 0}

    def fake_sleep(s):
        calls["n"] += 1
        if calls["n"] >= 2:
            attach_job.request_cancel()          # 2周目でキャンセルして抜ける

    monkeypatch.setattr(time, "sleep", fake_sleep)
    code = attach_job.run_attach_job(execute=True)
    assert code == 0
    assert attach_job.read_progress()["state"] == "cancelled"


def test_shaho_import_waits_for_haken_attach_lock(job_paths, tmp_path):
    """相互ロックの逆方向: 派遣台帳の添付ロックが生きていれば標報投入は止まる（2026-08-28決定）。"""
    import json as _json
    from config import Config
    from services.shaho_writer import ShahoWriteError, acquire_lock

    shaho_lock = tmp_path / "shaho_own.lock"
    Path(Config.HAKEN_ATTACH_LOCK_FILE).write_text(
        _json.dumps({"user": "添付ジョブ", "started_at": dt.datetime.now().isoformat(), "count": 165}),
        encoding="utf-8")
    with pytest.raises(ShahoWriteError, match="派遣台帳"):
        acquire_lock("2026-09", 1, path=str(shaho_lock))
    # 添付ロックが消えれば取れる
    Path(Config.HAKEN_ATTACH_LOCK_FILE).unlink()
    assert acquire_lock("2026-09", 1, path=str(shaho_lock))
    assert shaho_lock.exists()


def test_job_lock_conflict_becomes_error(job_paths, monkeypatch):
    import json as _json
    from config import Config

    Path(Config.HAKEN_ATTACH_LOCK_FILE).write_text(
        _json.dumps({"user": "別PC", "started_at": dt.datetime.now().isoformat(), "count": 1}),
        encoding="utf-8")
    code = attach_job.run_attach_job(execute=True)
    assert code == 1
    p = attach_job.read_progress()
    assert p["state"] == "error"
    assert "別の添付ジョブ" in p["message"]
