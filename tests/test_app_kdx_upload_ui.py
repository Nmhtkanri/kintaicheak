"""KDX専用アップロード欄と、jinjer雛形の検索コンボの画面配線。

2026-09 の KDX シフト表は文字がアウトライン化されたベクターPDFで、
`is_kdx_shift_pdf()` が False になり構造化パースを通れず、AI 読み取りに落ちて
凡例（記号→時刻）を持たないまま生記号 "A1"/"C4" が jinjer CSV に書かれた。
専用欄に入れれば経路に関わらず固定時刻が当たる、という配線をここで固定する。
"""

import io
import pathlib
import re

import app as app_module

STATIC = pathlib.Path(app_module.__file__).parent / "static"
TEMPLATES = pathlib.Path(app_module.__file__).parent / "templates"


def _html():
    return app_module.app.test_client().get("/").get_data(as_text=True)


def _script_js():
    return (STATIC / "script.js").read_text(encoding="utf-8")


def _style_css():
    return (STATIC / "style.css").read_text(encoding="utf-8")


# --- KDX専用アップロード欄 -------------------------------------------------

def test_index_has_kdx_upload_section():
    html = _html()
    assert 'id="kdx-files-section"' in html
    assert 'id="kdx-files-input"' in html
    assert 'name="kdx_files"' in html
    assert 'id="kdx-files-drop-zone"' in html
    assert 'id="kdx-files-selected"' in html


def test_kdx_hint_states_the_forced_times():
    """何が起きる欄なのかを画面に書いておく（黙って時刻を書き換える欄にしない）"""
    html = _html()
    assert "9:00" in html and "17:30" in html
    assert "16:30" in html and "34:00" in html


def test_script_wires_kdx_dropzone_and_mode_toggle():
    js = _script_js()
    assert "setupDropZone('kdx-files-drop-zone', 'kdx-files-input', 'kdx-files-selected'" in js
    assert "kdxFilesSection" in js
    # disabled を切らないと FormData が送らない／他モードで誤送信になる
    assert "kdxFilesInput.disabled = !isSchedule" in js


def test_kdx_is_not_in_inline_script():
    """index.html のインライン <script> は script.js より先に走るので、そこから
    script.js の世界に触ると全ボタンが死ぬ。KDX の JS は script.js だけに置く。"""
    src = (TEMPLATES / "index.html").read_text(encoding="utf-8")
    inline = "".join(m.group(0) for m in
                     re.finditer(r"<script(?![^>]*\bsrc=)[^>]*>.*?</script>", src, re.S))
    assert inline, "インライン script が見つからない（この検査が空振りしている）"
    assert "kdx" not in inline.lower()


def test_upload_accepts_kdx_only_post():
    """KDX欄だけに入れて実行しても「ファイルが選択されていません」で弾かれない"""
    client = app_module.app.test_client()
    res = client.post(
        "/upload",
        data={
            "mode": "csv_export",
            "kdx_files": (io.BytesIO(b"%PDF-1.4 dummy"), "KDXオペチーム9月シフト.pdf"),
        },
        content_type="multipart/form-data",
    )
    # 解析まで進めばSSE(200)。少なくとも「選択されていません」の400にはならない。
    if res.status_code == 400:
        body = res.get_data(as_text=True)
        assert "選択されていません" not in body, body


def test_upload_rejects_empty_post():
    """どちらの欄も空なら従来どおり弾く"""
    client = app_module.app.test_client()
    res = client.post("/upload", data={"mode": "csv_export"},
                      content_type="multipart/form-data")
    assert res.status_code == 400
    assert any("選択されていません" in e for e in res.get_json()["errors"])


# --- jinjer雛形の検索コンボ -------------------------------------------------

def test_script_has_searchable_template_combo():
    js = _script_js()
    assert "function buildTemplateCombo(" in js
    assert "function filterTemplates(" in js
    assert "function tplScore(" in js
    assert "buildTemplateSelect" not in js, "旧プルダウンの残骸"


def test_template_id_reader_is_not_select_only():
    """コンボは hidden input で値を持つので select 限定セレクタでは読めない"""
    js = _script_js()
    assert 'select[data-field="template_id"]' not in js
    assert '[data-field="template_id"]' in js


def test_search_input_has_no_data_field():
    """collectLegendFromUI は行内の全 input を dataset.field で舐めるので、
    検索用テキスト入力に data-field を付けると打ちかけの文字列が凡例に混入する。"""
    js = _script_js()
    start = js.index("function buildTemplateCombo(")
    body = js[start:js.index(chr(10) + "}", start)]
    assert "hidden.dataset.field = 'template_id'" in body
    assert "input.dataset.role = 'tpl-combo-input'" in body
    assert "input.dataset.field" not in body


def test_combo_panel_escapes_clipping():
    """.legend-sheet(overflow:hidden) / .modal-body(overflow-y:auto) に
    切られないよう body 直下 + position:fixed + modal-overlay より上の z-index"""
    css = _style_css()
    m = re.search(r"\.tpl-combo-panel\s*\{(.*?)\}", css, re.S)
    assert m, ".tpl-combo-panel のCSSが無い"
    block = m.group(1)
    assert "position: fixed" in block
    assert "z-index: 1100" in block
    js = _script_js()
    assert "document.body.appendChild(tplComboPanel)" in js
    # .modal-body がスクロールするので capture:true でないと追従できない
    assert "window.addEventListener('scroll'" in js
    scroll = js[js.index("window.addEventListener('scroll'"):]
    assert scroll[:scroll.index(chr(10))].rstrip().endswith("true);")


def test_legend_template_select_css_is_kept():
    """経費モードの経路選択(ki-route-choice)が同じクラスを流用している"""
    assert ".legend-template-select" in _style_css()
    assert "legend-template-select ki-route-choice" in (TEMPLATES / "index.html").read_text(
        encoding="utf-8")


def test_available_templates_include_abbr():
    """略称は検索キー兼表示名。サーバが送っていないとコンボが痩せる。"""
    src = pathlib.Path(app_module.__file__).read_text(encoding="utf-8")
    assert '"abbr": _tpl_get(t, "略称(3文字以内)")' in src
