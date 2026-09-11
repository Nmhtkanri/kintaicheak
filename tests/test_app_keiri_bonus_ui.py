# -*- coding: utf-8 -*-
"""経理モードの 給与／賞与 切替タブと /keiri_bonus_run の入力検証（2026-09-10）。

実行: python -X utf8 -m pytest tests/test_app_keiri_bonus_ui.py -q
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import app  # noqa: E402


@pytest.fixture(scope="module")
def html():
    return app.test_client().get("/").get_data(as_text=True)


@pytest.fixture(scope="module")
def js():
    return (ROOT / "static" / "script.js").read_text(encoding="utf-8")


def test_subtabs_and_panes_exist(html):
    for el_id in ("keiri-subtabs", "keiri-salary-pane", "keiri-bonus-pane", "keiri-bonus-month",
                  "keiri-bonus-label", "keiri-bonus-label-other", "keiri-bonus-hassei",
                  "keiri-bonus-shaho-hassei", "keiri-bonus-shaho-kigen", "keiri-bonus-refresh-custom",
                  "keiri-bonus-run-btn", "keiri-bonus-status", "keiri-bonus-result-area",
                  "keiri-bonus-files", "keiri-bonus-yokakunin"):
        assert f'id="{el_id}"' in html, el_id
    assert 'data-keiri-sub="salary"' in html and 'data-keiri-sub="bonus"' in html
    for el_id in ("keiri-bonus-paid-on", "keiri-bonus-count", "keiri-bonus-refresh-statements"):
        assert f'id="{el_id}"' in html, el_id
    # 賞与 CSV のアップロード欄は画面に出さない（入力は API のみ。2026-09-11）
    assert 'id="keiri-bonus-file"' not in html and 'name="keiri-bonus-source"' not in html
    # 既存の給与側 id はそのまま
    for el_id in ("keiri-month", "keiri-run-btn", "keiri-result-area", "keiri-error-area"):
        assert f'id="{el_id}"' in html, el_id
    # 賞与ペインはカードの中（エラー欄の前）にあり、既定で非表示
    assert html.index('id="keiri-salary-pane"') < html.index('id="keiri-bonus-pane"') < html.index('id="keiri-error-area"')
    assert 'id="keiri-bonus-pane" style="display:none"' in html


def test_subtab_wiring_mirrors_health_mode(js):
    assert "let keiriSubMode = 'salary';" in js
    assert js.index("let keiriSubMode = 'salary';") < js.index("function applyModeUI(")
    assert "const showKeiriBonus = isKeiri && keiriSubMode === 'bonus';" in js
    assert "#keiri-subtabs [data-keiri-sub]" in js
    assert "'/keiri_bonus_run'" in js
    assert "function keiriRenderBonus(" in js and "function keiriBonusLabel(" in js
    assert "function keiriBonusSource(" in js and "fd.append('source', source)" in js
    assert "getElementById('keiri-bonus-file')" not in js          # 画面からファイル入力を外した


def _post(client, **over):
    data = {"month": "2026-09", "label": "FE部賞与", "hassei": "2026-09-15", "source": "csv",
            "file": (io.BytesIO("社員番号,氏名\n1,a\n".encode("cp932")), "bonus.csv")}
    data.update(over)
    return client.post("/keiri_bonus_run", data=data, content_type="multipart/form-data")


def test_route_validates_inputs():
    c = app.test_client()
    assert _post(c, month="202609").status_code == 400
    assert _post(c, label="../x").status_code == 400
    assert _post(c, label="夏季一時金").status_code == 400          # 「賞与」を含まない
    assert _post(c, label="住民税賞与＜").status_code == 400          # 許可外の記号
    assert _post(c, hassei="2026/9/15").status_code == 400
    assert _post(c, shaho_kigen="2026-10").status_code == 400
    res = c.post("/keiri_bonus_run", data={"month": "2026-09", "label": "FE部賞与", "hassei": "2026-09-15"},
                 content_type="multipart/form-data")
    assert res.status_code == 400 and "賞与 CSV" in res.get_json()["errors"][0]


def test_route_api_source_does_not_require_file(monkeypatch):
    import services.keiri_bonus as kb
    seen = {}

    def fake_generate(month, raw, label, hassei, **kw):
        seen.update(kw); seen["raw"] = raw
        return {"people": 30, "out_dir": "x", "rates_src": "既定値", "yokakunin_md": "#", "alerts": {},
                "files": {}, "input_src": "API bonus-statements 2026-09"}

    monkeypatch.setattr(kb, "generate_bonus", fake_generate)
    c = app.test_client()
    res = c.post("/keiri_bonus_run", data={"month": "2026-09", "label": "FE部賞与", "hassei": "2026-09-15",
                                           "source": "api", "paid_on": "2026-09-25", "count": "1"},
                 content_type="multipart/form-data")
    assert res.status_code == 200, res.get_data(as_text=True)
    assert seen["raw"] is None and seen["source"] == "api" and seen["paid_on"] == "2026-09-25" and seen["count"] == 1
    assert res.get_json()["input_src"].startswith("API")
    res = c.post("/keiri_bonus_run", data={"month": "2026-09", "label": "FE部賞与", "hassei": "2026-09-15",
                                           "source": "api", "paid_on": "2026/9/25"}, content_type="multipart/form-data")
    assert res.status_code == 400


def test_route_reports_missing_columns_as_400():
    res = _post(app.test_client())
    assert res.status_code == 400
    assert "必要な列がありません" in res.get_json()["errors"][0]


def test_route_success_shape(monkeypatch):
    import services.keiri_bonus as kb

    def fake_generate(month, raw, label, hassei, **kw):
        assert isinstance(raw, (bytes, bytearray)) and label == "本社賞与" and hassei == "2026-06-15"
        assert kw["shaho_hassei"] is None and kw["shaho_kigen"] == "2026-07-31"
        return {"people": 10, "out_dir": r"outputs\keiri\202606", "rates_src": "既定値",
                "yokakunin_md": "# 要確認", "alerts": {"bonus_rate": 1},
                "files": {"支給": {"name": "freee_data_本社賞与（202606）.csv", "transactions": 10, "rows": 66, "total": 1},
                          "健康保険": {"name": "freee_data_本社賞与_健康保険（202606）.csv", "transactions": 1, "rows": 13, "total": 2},
                          "厚生年金": {"name": "freee_data_本社賞与_厚生年金（202606）.csv", "transactions": 1, "rows": 12, "total": 3}}}

    monkeypatch.setattr(kb, "generate_bonus", fake_generate)
    res = _post(app.test_client(), month="2026-06", label="本社賞与", hassei="2026-06-15", shaho_kigen="2026-07-31")
    assert res.status_code == 200, res.get_data(as_text=True)
    data = res.get_json()
    assert data["ym"] == "202606" and data["people"] == 10
    assert [f["種別"] for f in data["files"]] == ["支給", "健康保険", "厚生年金"]
    assert data["files"][0]["filename"] == "freee_data_本社賞与（202606）.csv"
    assert data["alerts"] == {"bonus_rate": 1}
