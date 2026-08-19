import pathlib

import app as app_module


def test_index_contains_invoice_mode_tab_and_preview_card():
    response = app_module.app.test_client().get("/")
    assert response.status_code == 200
    html = response.get_data(as_text=True)
    assert 'name="mode" value="invoice"' in html
    assert 'id="invoice-card"' in html
    assert 'id="invoice-preview-btn"' in html
    assert 'id="invoice-export-btn"' in html


def test_manual_row_button_is_offered():
    """スポット案件などPDFから拾えない請求分を任意で足せること（2026-08 谷津さん要望）。"""
    html = app_module.app.test_client().get("/").get_data(as_text=True)
    assert 'id="invoice-add-row-btn"' in html, "行を追加ボタンが無い"
    assert "手入力できます" in html, "任意で足せることが画面に書かれていない"
    js = pathlib.Path(app_module.__file__).parent.joinpath(
        "static", "script.js").read_text(encoding="utf-8")
    assert "function invoiceAddBlankRow()" in js
    assert "invoiceAddRowBtn.addEventListener" in js, "ボタンが処理につながっていない"
    assert "_manual_added: true" in js, "手入力行の目印が付いていない"
