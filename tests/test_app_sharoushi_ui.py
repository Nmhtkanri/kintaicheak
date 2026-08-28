"""社労士モードの画面まわり（CSVの形式スイッチ）。

2026-08-28 に列を 60 → 49 に減らし、社労士の受け入れ確認が済むまでの控えとして
「旧60列のまま出す」チェックを付けた。画面とサーバの配線が切れていないかを見る。
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


def test_index_has_layout_switch():
    html = _html()
    assert 'id="sharoushi-legacy"' in html, "旧60列に戻すチェックが無い"
    assert "旧60列" in html


def test_lead_text_says_49_columns():
    """「60列CSV」の言い残しがあると、何列で渡すのかが画面から読めなくなる。"""
    html = _html()
    assert "49列CSV" in html
    assert "60列CSV" not in html


def test_script_sends_layout():
    js = _script_js()
    assert "fd.append('layout'" in js, "形式がサーバへ送られていない"
    assert "sharoushi-legacy" in js


def test_layout_switch_is_not_in_inline_script():
    """index.html のインライン <script> は script.js より先に走るので、そこから
    script.js の世界（要素ID の参照を含む処理）に触ると全ボタンが死ぬ。
    形式スイッチは静的HTMLと script.js だけで完結していること。"""
    src = (TEMPLATES / "index.html").read_text(encoding="utf-8")
    inline = "".join(m.group(0) for m in
                     re.finditer(r"<script(?![^>]*\bsrc=)[^>]*>.*?</script>", src, re.S))
    assert inline, "インライン script が見つからない（この検査が空振りしている）"
    assert "sharoushi-legacy" not in inline


def test_run_route_accepts_and_validates_layout():
    client = app_module.app.test_client()
    res = client.post("/sharoushi_run", data={"month": "2026-08", "layout": "V9"})
    assert res.status_code == 400
    assert "形式" in "".join(res.get_json()["errors"])


def test_run_route_rejects_bad_month_before_touching_jinjer():
    res = app_module.app.test_client().post("/sharoushi_run", data={"month": "202608"})
    assert res.status_code == 400
