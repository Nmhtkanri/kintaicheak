"""派遣元管理台帳モードのルート（鮮度・build・トリアージ・ダウンロード）。

daicho のパスは services.daicho.config のモジュール定数を monkeypatch で tmp に差し替え、
build 本体は fake に置き換える（実データ・実APIに触れない）。
"""

import csv
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app as app_module  # noqa: E402
from services.daicho import config as daicho_config  # noqa: E402


@pytest.fixture
def client():
    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as c:
        yield c


@pytest.fixture
def daicho_dirs(tmp_path, monkeypatch):
    """daicho の全パスとロック・許可CSVを tmp に向ける（実データ・NAS に触れない）。"""
    from config import Config

    data = tmp_path / "data"
    for name in ("input", "output", "ログ", "PDF"):
        (data / name).mkdir(parents=True)
    monkeypatch.setattr(daicho_config, "DATA_ROOT", data)
    monkeypatch.setattr(daicho_config, "INPUT_DIR", data / "input")
    monkeypatch.setattr(daicho_config, "OUTPUT_DIR", data / "output")
    monkeypatch.setattr(daicho_config, "LOG_DIR", data / "ログ")
    monkeypatch.setattr(daicho_config, "PDF_ROOT", data / "PDF")
    monkeypatch.setattr(daicho_config, "ATTACH_PROGRESS_JSON", data / "ログ" / "添付進捗.json")
    monkeypatch.setattr(daicho_config, "ATTACH_CANCEL_FLAG", data / "ログ" / "添付キャンセル.flag")
    monkeypatch.setattr(Config, "HAKEN_ATTACH_LOCK_FILE", str(tmp_path / "添付.lock"))
    monkeypatch.setattr(Config, "SHAHO_IMPORT_LOCK_FILE", str(tmp_path / "標報.lock"))
    monkeypatch.setattr(Config, "HAKEN_ATTACH_ALLOWED_USERS_CSV", str(tmp_path / "許可.csv"))
    return data


@pytest.fixture
def allow_attach(tmp_path, monkeypatch, daicho_dirs):
    """現在ユーザーを添付の許可リストに載せる。"""
    from config import Config
    from services.sap_import_ledger import current_user

    Path(Config.HAKEN_ATTACH_ALLOWED_USERS_CSV).write_text(
        f"ユーザー名,表示名,備考\n{current_user()},テスト実行者,\n", encoding="utf-8-sig")


def _write_warn_csv(out_dir: Path, quarter: str, rows: list[tuple]) -> None:
    path = out_dir / f"派遣元管理台帳_{quarter}_警告.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as fp:
        w = csv.writer(fp)
        w.writerow(["区分", "契約No", "氏名", "内容"])
        w.writerows(rows)


def test_build_rejects_bad_quarter(client):
    res = client.post("/haken_build", data={"quarter": "2026-08"})
    assert res.status_code == 400
    assert "四半期" in "".join(res.get_json()["errors"])


def test_build_triage_new_continued_resolved(client, daicho_dirs, monkeypatch):
    out = daicho_dirs / "output"
    # 前回の警告: A(解消される), B(継続する)
    _write_warn_csv(out, "2026Q2", [
        ("台帳", "K-1", "山田 太郎", "A: 直る警告"),
        ("台帳", "K-2", "佐藤 花子", "B: 残る警告"),
    ])

    def fake_build_quarter(quarter, **kwargs):
        _write_warn_csv(out, quarter, [
            ("台帳", "K-2", "佐藤 花子", "B: 残る警告"),
            ("全体", "", "", "C: 新しい警告"),
        ])
        return {"quarter": quarter, "label": "2026年4-6月期",
                "counts": {"total": 1, "estaffing": 1, "fieldglass": 0, "direct": 0, "people": 1},
                "match": {"ok": 1, "none": 0, "ambiguous": 0}, "n_warn": 1,
                "paths": {"xlsx": str(out / "x.xlsx"), "csv": "c", "warnings": "w"},
                "inputs": {}, "global_warnings": ["C: 新しい警告"], "notes": [],
                "summary": "テスト"}

    monkeypatch.setattr("services.daicho.build.build_quarter", fake_build_quarter)
    res = client.post("/haken_build", data={"quarter": "2026Q2"})
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert data["had_prev"] is True
    triage = {w["内容"]: w["triage"] for w in data["warnings"]}
    assert triage == {"B: 残る警告": "continued", "C: 新しい警告": "new"}
    assert [r["内容"] for r in data["resolved"]] == ["A: 直る警告"]


def test_build_first_run_marks_all_new(client, daicho_dirs, monkeypatch):
    out = daicho_dirs / "output"

    def fake_build_quarter(quarter, **kwargs):
        _write_warn_csv(out, quarter, [("台帳", "K-9", "田中 一", "初回の警告")])
        return {"quarter": quarter, "label": "2026年4-6月期",
                "counts": {"total": 1, "estaffing": 1, "fieldglass": 0, "direct": 0, "people": 1},
                "match": {"ok": 1, "none": 0, "ambiguous": 0}, "n_warn": 1,
                "paths": {"xlsx": "x", "csv": "c", "warnings": "w"},
                "inputs": {}, "global_warnings": [], "notes": [], "summary": "テスト"}

    monkeypatch.setattr("services.daicho.build.build_quarter", fake_build_quarter)
    data = client.post("/haken_build", data={"quarter": "2026Q2"}).get_json()
    assert data["had_prev"] is False
    assert [w["triage"] for w in data["warnings"]] == ["new"]
    assert data["resolved"] == []


def test_build_busy_returns_409(client):
    assert app_module._haken_build_lock.acquire(blocking=False)
    try:
        res = client.post("/haken_build", data={"quarter": "2026Q2"})
        assert res.status_code == 409
        assert "実行中" in "".join(res.get_json()["errors"])
    finally:
        app_module._haken_build_lock.release()


def test_build_reports_file_open_error(client, daicho_dirs, monkeypatch):
    def fake_build_quarter(quarter, **kwargs):
        raise PermissionError("book is open")

    monkeypatch.setattr("services.daicho.build.build_quarter", fake_build_quarter)
    res = client.post("/haken_build", data={"quarter": "2026Q2"})
    assert res.status_code == 400
    assert "Excel" in "".join(res.get_json()["errors"])


def test_download_serves_only_known_kinds(client, daicho_dirs):
    out = daicho_dirs / "output"
    (out / "派遣元管理台帳_2026Q2_一覧.csv").write_text("x", encoding="utf-8")

    assert client.get("/haken_download?quarter=2026Q2&kind=zip").status_code == 400
    assert client.get("/haken_download?quarter=2026Q2&kind=xlsx").status_code == 404
    res = client.get("/haken_download?quarter=2026Q2&kind=csv")
    assert res.status_code == 200
    assert res.data == b"x"


def test_freshness_reports_missing_required(client, daicho_dirs):
    data = client.get("/haken_freshness?quarter=2026Q2").get_json()
    assert data["success"] is True
    assert data["overall"] == "missing"
    verdicts = {r["key"]: r["verdict"] for r in data["inputs"]}
    assert verdicts["tc"] == "missing"
    assert verdicts["cpi"] == "missing"
    assert verdicts["roster"] == "missing"


def test_quarter_status_with_empty_dirs(client, daicho_dirs):
    data = client.get("/haken_quarter_status?quarter=2026Q2").get_json()
    assert data["success"] is True
    assert data["steps"]["build"]["exists"] is False
    assert data["steps"]["attach"]["state"] == "none"
    assert data["steps"]["pdf"]["count"] == 0


# --- ③ PDF出力ジョブ ---

def _wait_job(client, job_id, timeout=5.0):
    import time
    deadline = time.time() + timeout
    while time.time() < deadline:
        data = client.get(f"/haken_pdf_status/{job_id}").get_json()
        if data.get("done"):
            return data
        time.sleep(0.05)
    raise AssertionError("PDFジョブが終わらない")


def test_pdf_job_with_fake_export(client, daicho_dirs, monkeypatch):
    def fake_export(quarter, on_progress=None, **kwargs):
        if on_progress:
            on_progress(1, 2, "a.pdf")
            on_progress(2, 2, "b.pdf")
        return 2, ["W: テスト警告"]

    monkeypatch.setattr("services.daicho.export_pdf.export_quarter", fake_export)
    res = client.post("/haken_export_pdf", data={"quarter": "2026Q2"})
    assert res.status_code == 200
    job_id = res.get_json()["job_id"]
    data = _wait_job(client, job_id)
    assert data["ok"] is True
    assert data["n"] == 2
    assert data["warnings"] == ["W: テスト警告"]


def test_pdf_job_reports_missing_book(client, daicho_dirs, monkeypatch):
    def fake_export(quarter, on_progress=None, **kwargs):
        raise FileNotFoundError("台帳ブックが無い")

    monkeypatch.setattr("services.daicho.export_pdf.export_quarter", fake_export)
    job_id = client.post("/haken_export_pdf", data={"quarter": "2026Q2"}).get_json()["job_id"]
    data = _wait_job(client, job_id)
    assert data["ok"] is False
    assert "台帳ブック" in data["error"]


def test_pdf_rejects_while_attach_running(client, daicho_dirs):
    daicho_config.ATTACH_PROGRESS_JSON.write_text('{"state": "running"}', encoding="utf-8")
    res = client.post("/haken_export_pdf", data={"quarter": "2026Q2"})
    assert res.status_code == 409
    assert "添付" in "".join(res.get_json()["errors"])


def test_pdf_rejects_while_build_running(client, daicho_dirs):
    assert app_module._haken_build_lock.acquire(blocking=False)
    try:
        res = client.post("/haken_export_pdf", data={"quarter": "2026Q2"})
        assert res.status_code == 409
    finally:
        app_module._haken_build_lock.release()


def test_pdf_status_unknown_job(client):
    assert client.get("/haken_pdf_status/deadbeef").status_code == 404


# --- ④ jinjer添付 ---

def _exec_body(**over):
    body = {"mode": "now", "confirm": "添付"}
    body.update(over)
    return body


def test_attach_execute_denied_without_permission(client, daicho_dirs):
    # 許可CSVが無い＝フェイルクローズで403
    res = client.post("/haken_attach_execute", json=_exec_body())
    assert res.status_code == 403
    assert "書き込み" in "".join(res.get_json()["errors"])


def test_attach_execute_requires_confirm_word(client, allow_attach):
    res = client.post("/haken_attach_execute", json=_exec_body(confirm="てんぷ"))
    assert res.status_code == 400
    assert "添付" in "".join(res.get_json()["errors"])


def test_attach_execute_spawns_detached_now(client, allow_attach, daicho_dirs, monkeypatch):
    seen = {}

    def fake_spawn(cmd_args):
        seen["args"] = cmd_args
        return 4242

    monkeypatch.setattr(app_module, "_haken_attach_spawn", fake_spawn)
    res = client.post("/haken_attach_execute", json=_exec_body(limit=1))
    assert res.status_code == 200
    data = res.get_json()
    assert data["mode"] == "detached" and data["pid"] == 4242
    assert "--execute" in seen["args"]
    assert "--limit" in seen["args"] and "1" in seen["args"]
    assert "--start-at" not in seen["args"]
    # 二重クリック防止: 起動直後は spawning が入っている
    assert json.loads(daicho_config.ATTACH_PROGRESS_JSON.read_text(
        encoding="utf-8"))["state"] == "spawning"
    # spawning 中の再実行は409
    assert client.post("/haken_attach_execute", json=_exec_body()).status_code == 409


def test_attach_execute_tonight_passes_start_at(client, allow_attach, monkeypatch):
    seen = {}
    monkeypatch.setattr(app_module, "_haken_attach_spawn",
                        lambda args: seen.setdefault("args", args) and 1 or 1)
    res = client.post("/haken_attach_execute", json=_exec_body(mode="tonight"))
    assert res.status_code == 200
    assert "--start-at" in seen["args"]
    from config import Config
    assert Config.HAKEN_ATTACH_TONIGHT_AT in seen["args"]


def test_attach_execute_busy_and_lock_409(client, allow_attach, daicho_dirs):
    from services.daicho import attach_job

    attach_job.write_progress({"state": "running"})
    assert client.post("/haken_attach_execute", json=_exec_body()).status_code == 409
    attach_job.write_progress({"state": "done"})
    # 別PCの有効ロック
    from config import Config
    Path(Config.HAKEN_ATTACH_LOCK_FILE).write_text(
        json.dumps({"user": "他PC", "started_at": __import__("datetime").datetime.now().isoformat(),
                    "count": 1}), encoding="utf-8")
    assert client.post("/haken_attach_execute", json=_exec_body()).status_code == 409


def test_attach_cancel_writes_flag(client, allow_attach, daicho_dirs):
    from services.daicho import attach_job

    assert client.post("/haken_attach_cancel").status_code == 409  # ジョブなし
    attach_job.write_progress({"state": "running"})
    res = client.post("/haken_attach_cancel")
    assert res.status_code == 200
    assert daicho_config.ATTACH_CANCEL_FLAG.exists()


def test_attach_preview_and_verify_with_fakes(client, daicho_dirs, monkeypatch):
    monkeypatch.setattr("services.daicho.attach_job.preview",
                        lambda employees=None, limit=None: {
                            "total": 3, "people": 2, "already": 1, "to_attach": 2,
                            "plan": [], "skipped": ["X: 対象外"], "eta_minutes": 2})
    res = client.post("/haken_attach_preview", data={})
    assert res.status_code == 200
    data = res.get_json()
    assert data["to_attach"] == 2
    assert data["can_write"] is False          # 許可CSVなし＝閲覧はできるがボタンは無効
    assert data["state"] == "none"

    monkeypatch.setattr("services.daicho.attach_job.verify_data",
                        lambda employees=None: {"total": 3, "ok": 2,
                                                "missing": [{"emp": "1", "pdf": "a.pdf", "date": "2026/4/1"}]})
    data = client.post("/haken_verify", data={}).get_json()
    assert data["ok"] == 2 and len(data["missing"]) == 1
