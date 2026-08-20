"""提出用PDF作成モードの画面とルート。"""
import app as app_module


def _client():
    return app_module.app.test_client()


def test_pdf_and_csv_live_in_one_invoice_mode():
    """PDF作成とCSV作成は「請求書」モード1つにまとめる（モードは分けない）。"""
    html = _client().get("/").get_data(as_text=True)
    assert 'name="mode" value="invoice_pdf"' not in html, "モードを分けない"
    assert 'name="mode" value="invoice"' in html
    assert 'id="invoice-pdf-card"' in html, "提出用PDFのカードが無い"
    assert 'id="invoice-card"' in html, "freee CSVのカードが無い"
    assert 'id="invoice-pdf-month"' in html
    assert 'id="invoice-pdf-buttons"' in html, "会社ごとのボタンを並べる場所が無い"
    # 何をするモードかが画面に書かれていること
    assert "提出は手動" in html
    js = __import__("pathlib").Path(app_module.__file__).parent.joinpath(
        "static", "script.js").read_text(encoding="utf-8")
    assert "invoicePdfCard.style.display = isInvoice" in js,         "請求書モードで提出用PDFのカードが出ない"


def test_companies_are_grouped_by_partner():
    """会社ごとにボタンを出すので、設定CSVを取引先でまとめて返す。"""
    data = _client().get("/invoice_pdf_companies").get_json()
    assert data["success"], data
    assert data["companies"], "設定CSVが読めていない"
    for company in data["companies"]:
        assert company["partner"], "取引先名が空"
        assert company["people"], "その会社の対象者が空"


def test_bad_month_is_rejected():
    res = _client().post("/invoice_pdf_run",
                         data={"month": "2026-7", "partner": "x", "force": "0"})
    assert res.status_code == 400
    assert "YYYY-MM" in " ".join(res.get_json()["errors"])


def test_unknown_partner_is_rejected():
    """★設定に無い取引先を指定しても、勝手に他社を処理しない。"""
    res = _client().post("/invoice_pdf_run",
                         data={"month": "2026-07", "partner": "存在しない社", "force": "0"})
    assert res.status_code == 400
    assert "設定CSVに" in " ".join(res.get_json()["errors"])


def test_front_end_offers_confirm_button():
    """要確認が出たときだけ、そのまま作るボタンを出す。"""
    import pathlib
    js = pathlib.Path(app_module.__file__).parent.joinpath(
        "static", "script.js").read_text(encoding="utf-8")
    assert "invoicePdfLoadCompanies" in js, "会社ボタンの読み込みが無い"
    assert "invoice-pdf-force" in js, "確認して作るボタンが無い"
    assert "needs_confirm" in js, "要確認の表示が無い"


def test_button_label_is_short():
    """★ボタンは短く。「株式会社」や「（旧○○）」まで出すと押しにくい。

    APIへ渡すのは元の正式名のまま（設定CSVと突き合わせるため）。
    """
    import pathlib
    js = pathlib.Path(app_module.__file__).parent.joinpath(
        "static", "script.js").read_text(encoding="utf-8")
    assert "function invoicePdfShortName" in js
    assert "invoicePdfShortName(c.partner)" in js, "ボタン名に短縮を使っていない"
    assert "body.append('partner', partner)" in js, "APIへは正式名のまま送ること"
