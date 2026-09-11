# -*- coding: utf-8 -*-
"""経理モード「その他」手入力の ＋／− 行と、保存ルートの合計チェック（2026-09-10）。

画面は実クリックできないので、配線（関数名・class・イベント委譲）の文字列と、
/keiri_sonota_save の入力検証を test_client で固定する。

実行: python -X utf8 -m pytest tests/test_app_keiri_sonota_ui.py -q
"""
from __future__ import annotations

import os
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


def test_sonota_area_ids_still_exist(html):
    for el_id in ("keiri-sonota-area", "keiri-sonota-rows", "keiri-sonota-choices",
                  "keiri-sonota-save-btn", "keiri-sonota-status", "keiri-sonota-path"):
        assert f'id="{el_id}"' in html, el_id
    assert "「＋」で行を増やして金額を分けます" in html


def test_row_html_has_plus_minus_and_amount_input(js):
    assert "function keiriSonotaRowHtml(" in js
    for cls in ("keiri-sonota-add", "keiri-sonota-del", "keiri-sonota-amt",
                "keiri-sonota-combo", "keiri-sonota-biko",
                "keiri-sonota-emp", "keiri-sonota-name", "keiri-sonota-bumon"):
        assert cls in js, cls
    assert 'data-src="' in js


def test_plus_minus_use_delegated_listener_not_inline_onclick(js, html):
    # 行は動的に増減するので、表の親でクリックを受ける。インライン onclick は使わない（全ボタン死の事故防止）
    i = js.index("keiriSonotaRowsEl.addEventListener('click'")
    assert "btn.classList.contains('keiri-sonota-add')" in js[i:i + 1200]
    assert "btn.classList.contains('keiri-sonota-del')" in js[i:i + 1200]
    assert "tr.remove()" in js[i:i + 1200]
    assert "onclick" not in js[js.index("function keiriSonotaRowHtml("):js.index("function keiriCollectSonota(")]
    assert 'keiri-sonota-add" onclick' not in html


def test_collect_sends_row_amount_and_statement_amount(js, html):
    i = js.index("function keiriCollectSonota(")
    body = js[i:js.index("const keiriSonotaRowsEl")]
    assert "'明細金額': p['金額']" in body
    assert "'金額': amt" in body
    assert "'明細社員番号': p['社員番号']" in body
    assert "(el.value || '').trim() === (el.dataset.init || '')" in body   # 初期値のままの氏名・部門は空で送る
    assert "data-init=" in js[js.index("function keiriSonotaRowHtml("):js.index("function keiriCollectSonota(")]
    assert 'id="keiri-sonota-bumon-choices"' in html
    assert "行の金額の合計" in body          # 合計不一致は保存前に画面で止める


def test_run_payload_carries_bumon_choices():
    """/keiri_run の応答に部門の候補が載る（載っていないと画面の datalist が空のまま）。"""
    src = open(os.path.join(os.path.dirname(__file__), "..", "app.py"), encoding="utf-8").read()
    i = src.index('"sonota_pending": result["sonota_pending"]')
    assert '"sonota_bumon_choices": result.get("sonota_bumon_choices", [])' in src[i:i + 400]


def test_save_route_rejects_split_rows_whose_total_differs():
    client = app.test_client()
    payload = {"month": "2026-08", "entries": [
        {"社員番号": "2024047", "氏名": "能美 龍郎", "金額": 20000, "明細金額": 33000,
         "勘定科目": "支払手数料", "品目": "雑費", "税区分": "課対仕入10%", "備考": ""},
        {"社員番号": "2024047", "氏名": "能美 龍郎", "金額": 10000, "明細金額": 33000,
         "勘定科目": "福利厚生費", "品目": "福利厚生費", "税区分": "対象外", "備考": "有給買取"},
    ]}
    res = client.post("/keiri_sonota_save", json=payload)
    assert res.status_code == 400
    data = res.get_json()
    assert data["success"] is False
    assert "30,000" in data["errors"][0] and "33,000" in data["errors"][0]


def test_save_route_rejects_fractional_amount_and_missing_statement_amount():
    client = app.test_client()
    base = {"社員番号": "2024047", "氏名": "能美 龍郎", "勘定科目": "支払手数料", "品目": "雑費", "税区分": "課対仕入10%", "備考": ""}
    res = client.post("/keiri_sonota_save", json={"month": "2026-08", "entries": [
        dict(base, 金額=0.5, 明細金額=33000), dict(base, 金額=32999.5, 明細金額=33000)]})
    assert res.status_code == 400 and "整数" in res.get_json()["errors"][0]
    res = client.post("/keiri_sonota_save", json={"month": "2026-08", "entries": [
        dict(base, 金額=20000, 明細金額=33000), dict(base, 金額=13000)]})
    assert res.status_code == 400 and "明細金額" in res.get_json()["errors"][0]


def test_save_route_groups_total_by_statement_employee_when_employee_edited():
    """社員番号を別人に書き換えた行も、合計は明細社員番号（jinjer の金額の持ち主）でまとめて検証する。"""
    client = app.test_client()
    base = {"氏名": "", "勘定科目": "支払手数料", "品目": "雑費", "税区分": "課対仕入10%", "備考": "", "明細金額": 33000,
            "明細社員番号": "2024047"}
    res = client.post("/keiri_sonota_save", json={"month": "2026-08", "entries": [
        dict(base, 社員番号="2025001", 金額=20000, 部門="OT：その他"),
        dict(base, 社員番号="2024047", 金額=10000)]})
    assert res.status_code == 400
    assert "2024047" in res.get_json()["errors"][0] and "30,000" in res.get_json()["errors"][0]


def test_save_route_writes_split_rows_when_total_matches(tmp_path, monkeypatch):
    from services import keiri_engine
    import config as cfg
    ledger = tmp_path / "sonota.csv"
    monkeypatch.setattr(cfg.Config, "KEIRI_SONOTA_MANUAL_CSV", str(ledger))
    monkeypatch.setattr(keiri_engine.Config, "KEIRI_SONOTA_MANUAL_CSV", str(ledger), raising=False)
    client = app.test_client()
    payload = {"month": "2026-08", "entries": [
        {"社員番号": "2024047", "氏名": "能美 龍郎", "金額": 20000, "明細金額": 33000,
         "勘定科目": "支払手数料", "品目": "雑費", "税区分": "課対仕入10%", "備考": ""},
        {"社員番号": "2024047", "氏名": "能美 龍郎", "金額": 13000, "明細金額": 33000,
         "勘定科目": "福利厚生費", "品目": "福利厚生費", "税区分": "対象外", "備考": "有給買取"},
    ]}
    res = client.post("/keiri_sonota_save", json=payload)
    assert res.status_code == 200, res.get_data(as_text=True)
    assert res.get_json()["saved"] == 2
    loaded = keiri_engine.load_sonota_manual(str(ledger))
    assert [r["金額"] for r in loaded[("2026-08", "2024047")]] == ["20000", "13000"]
