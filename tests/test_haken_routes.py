"""派遣元管理台帳モードのルート（鮮度・build・トリアージ・ダウンロード）。

daicho のパスは services.daicho.config のモジュール定数を monkeypatch で tmp に差し替え、
build 本体は fake に置き換える（実データ・実APIに触れない）。
"""

import csv
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
    """daicho の全パスを tmp に向ける（実データ Z:\\派遣元管理台帳 に触れない）。"""
    data = tmp_path / "data"
    for name in ("input", "output", "ログ", "PDF"):
        (data / name).mkdir(parents=True)
    monkeypatch.setattr(daicho_config, "DATA_ROOT", data)
    monkeypatch.setattr(daicho_config, "INPUT_DIR", data / "input")
    monkeypatch.setattr(daicho_config, "OUTPUT_DIR", data / "output")
    monkeypatch.setattr(daicho_config, "LOG_DIR", data / "ログ")
    monkeypatch.setattr(daicho_config, "PDF_ROOT", data / "PDF")
    monkeypatch.setattr(daicho_config, "ATTACH_PROGRESS_JSON", data / "ログ" / "添付進捗.json")
    return data


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
