"""派遣元管理台帳モードの画面配線。

タブ・カード・script.js の関数がそろっているか、インライン script に触れていないか
（2026-08-19 の全ボタン死の再発防止）、ルートの四半期バリデーションが効くかを見る。
"""

import pathlib
import re

import app as app_module

STATIC = pathlib.Path(app_module.__file__).parent / "static"
TEMPLATES = pathlib.Path(app_module.__file__).parent / "templates"


def _html():
    return app_module.app.test_client().get("/").get_data(as_text=True)


def _script_js():
    return (STATIC / "script.js").read_text(encoding="utf-8")


def test_index_has_haken_tab_and_card():
    html = _html()
    assert 'value="haken"' in html, "モードタブが無い"
    assert 'id="haken-card"' in html
    assert 'id="haken-quarter"' in html
    assert 'id="haken-freshness-btn"' in html
    assert 'id="haken-stepper"' in html
    assert 'id="haken-build-btn"' in html
    assert 'id="haken-warn-table"' in html
    assert 'id="haken-pdf-btn"' in html
    assert 'id="haken-attach-preview-btn"' in html
    assert 'id="haken-attach-confirm"' in html
    assert 'id="haken-attach-now-btn"' in html
    assert 'id="haken-attach-tonight-btn"' in html
    assert 'id="haken-verify-btn"' in html


def test_script_wires_haken_mode():
    js = _script_js()
    assert "haken:" in js, "MODE_HINTS に無い"
    assert "isHaken" in js, "applyModeUI に無い"
    assert "hakenCard" in js
    assert "/haken_freshness" in js
    assert "/haken_quarter_status" in js
    assert "/haken_build" in js
    assert "/haken_download" in js
    assert "/haken_export_pdf" in js
    assert "/haken_pdf_status" in js
    assert "/haken_attach_preview" in js
    assert "/haken_attach_execute" in js
    assert "/haken_attach_status" in js
    assert "/haken_attach_cancel" in js
    assert "/haken_verify" in js


def test_haken_is_not_in_inline_script():
    """index.html のインライン <script> は script.js より先に走るので、そこから
    script.js の世界に触ると全ボタンが死ぬ。派遣台帳の JS は script.js だけに置く。"""
    src = (TEMPLATES / "index.html").read_text(encoding="utf-8")
    inline = "".join(m.group(0) for m in
                     re.finditer(r"<script(?![^>]*\bsrc=)[^>]*>.*?</script>", src, re.S))
    assert inline, "インライン script が見つからない（この検査が空振りしている）"
    assert "haken" not in inline


def test_freshness_rejects_bad_quarter():
    res = app_module.app.test_client().get("/haken_freshness?quarter=2026-08")
    assert res.status_code == 400
    assert "四半期" in "".join(res.get_json()["errors"])


def test_freshness_requires_quarter():
    res = app_module.app.test_client().get("/haken_freshness")
    assert res.status_code == 400


def test_quarter_status_rejects_bad_quarter():
    res = app_module.app.test_client().get("/haken_quarter_status?quarter=Q9")
    assert res.status_code == 400
