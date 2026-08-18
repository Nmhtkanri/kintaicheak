import app as app_module


def test_index_contains_invoice_mode_tab_and_preview_card():
    response = app_module.app.test_client().get("/")
    assert response.status_code == 200
    html = response.get_data(as_text=True)
    assert 'name="mode" value="invoice"' in html
    assert 'id="invoice-card"' in html
    assert 'id="invoice-preview-btn"' in html
    assert 'id="invoice-export-btn"' in html
