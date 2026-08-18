from pathlib import Path

import app as app_module
from services import invoice_mode


def _preview_payload():
    return {
        "month": "2026-07",
        "rows": [{"管理番号": "2017012", "_errors": []}],
        "documents": [],
        "selected_files": ["invoice.pdf"],
        "ignored": [],
        "missing_roots": [],
        "scan_errors": [],
        "parse_errors": [],
        "signature": "abc",
        "target_count": 28,
        "sales_book": "sales.xlsx",
    }


def test_invoice_preview_route(monkeypatch):
    monkeypatch.setattr(invoice_mode, "build_preview", lambda *args, **kwargs: _preview_payload())
    client = app_module.app.test_client()
    response = client.post("/invoice_preview", data={"month": "2026-07"})
    assert response.status_code == 200
    data = response.get_json()
    assert data["success"] is True
    assert data["target_count"] == 28
    assert data["error_rows"] == 0


def test_invoice_preview_rejects_bad_month():
    client = app_module.app.test_client()
    response = client.post("/invoice_preview", data={"month": "202607"})
    assert response.status_code == 400
    assert response.get_json()["success"] is False


def test_invoice_export_rejects_changed_sources(monkeypatch):
    monkeypatch.setattr(invoice_mode, "load_target_roots", lambda *_: ["root"])
    monkeypatch.setattr(invoice_mode, "current_scan_signature",
                        lambda *_: ("new", {"selected": [], "ignored": [],
                                             "missing_roots": [], "scan_errors": []}))
    client = app_module.app.test_client()
    response = client.post("/invoice_export", json={
        "month": "2026-07", "rows": [{}], "signature": "old",
    })
    assert response.status_code == 409
    assert "追加・更新" in response.get_json()["errors"][0]


def test_invoice_export_returns_download_links(monkeypatch, tmp_path):
    monkeypatch.setattr(invoice_mode, "load_target_roots", lambda *_: ["root"])
    scan = {"selected": ["invoice.pdf"], "ignored": [],
            "missing_roots": [], "scan_errors": []}
    monkeypatch.setattr(invoice_mode, "current_scan_signature", lambda *_: ("same", scan))

    captured = {}

    def fake_export(month, rows, output_root, log_context=None):
        captured["month"] = month
        captured["rows"] = rows
        captured["context"] = log_context
        return {"row_count": 1, "csv_name": "売上.csv", "log_name": "実行ログ.json"}

    monkeypatch.setattr(invoice_mode, "export_csv", fake_export)
    client = app_module.app.test_client()
    response = client.post("/invoice_export", json={
        "month": "2026-07", "rows": [{"管理番号": "1"}], "signature": "same",
    })
    assert response.status_code == 200
    data = response.get_json()
    assert data["csv_url"] == "/invoice_download/202607/売上.csv"
    assert data["log_url"] == "/invoice_download/202607/実行ログ.json"
    assert captured["context"]["source_files"] == ["invoice.pdf"]


def test_invoice_download_rejects_invalid_month():
    client = app_module.app.test_client()
    response = client.get("/invoice_download/not-a-month/file.csv")
    assert response.status_code == 400
