# -*- coding: utf-8 -*-
"""健康診断申込モードの画面配線（タブ・カード・script.js・CSS）。

インライン script に ha- が混ざると全ボタンが死ぬ事故につながるので、ここで機械的に止める。
"""

import pathlib
import re
import shutil
import subprocess

import pytest

import app as app_module

STATIC = pathlib.Path(app_module.__file__).parent / "static"


def _html():
    return app_module.app.test_client().get("/").get_data(as_text=True)


def _script_js():
    return (STATIC / "script.js").read_text(encoding="utf-8")


def _style_css():
    return (STATIC / "style.css").read_text(encoding="utf-8")


def test_index_has_health_apply_tab_and_card():
    """健診申込は専用タブではなく、健康診断HPMモードの項目（サブタブ）として出す（2026-09-04 統合）。"""
    html = _html()
    assert 'value="health_apply"' not in html               # 専用タブは廃止
    assert 'value="health_hpm"' in html
    assert 'id="health-subtabs"' in html
    assert 'data-health-sub="hpm"' in html
    assert 'data-health-sub="apply"' in html
    assert "健診申込" in html
    for el_id in ("ha-card", "ha-access-banner", "ha-year", "ha-reload-btn", "ha-status-line",
                  "ha-responses-btn", "ha-responses-area", "ha-cnt-targets", "ha-cnt-error",
                  "ha-workbook-issues", "ha-resp-filter", "ha-resp-search", "ha-resp-body", "ha-error-area"):
        assert f'id="{el_id}"' in html, el_id
    assert 'class="workflow-card mode-accent acc-happly" id="ha-card" style="display:none"' in html


def test_index_has_target_registration_section():
    html = _html()
    for el_id in ("ha-targets-section", "ha-target-ids", "ha-preview-btn", "ha-preview-status", "ha-preview-area",
                  "ha-pcnt-add", "ha-pcnt-conflict", "ha-pcnt-blocked", "ha-preview-input-issues", "ha-preview-body",
                  "ha-commit-guide", "ha-commit-confirm", "ha-commit-btn", "ha-commit-status", "ha-commit-result"):
        assert f'id="{el_id}"' in html, el_id
    assert 'id="ha-commit-btn" disabled' in html          # 確認語が一致するまで押せない
    assert "まだ何も書きません" in html


def test_script_js_wires_target_registration():
    js = _script_js()
    for fn in ("function haPreviewTargets", "function haRenderPreview", "function haCommitRefresh",
               "function haCommitTargets"):
        assert fn in js, fn
    assert "'/health_apply_targets_preview'" in js
    assert "'/health_apply_targets_commit'" in js
    assert "confirm.value.trim() === p.confirm_phrase" in js


def test_card_promises_no_jinjer_write_and_no_local_copy():
    html = _html()
    start = html.index('id="ha-card"')
    end = html.index('id="expense-card"')
    card = html[start:end]
    assert "jinjer には書きません" in card
    assert "保存しません" in card
    assert "本人が見た証拠にはしません" in card


def test_script_js_wires_mode_and_routes():
    js = _script_js()
    assert "health_apply: '" not in js                   # MODE_HINTS から専用モードは消えている
    assert "isHealthApply" not in js                     # 専用モードの分岐は残さない
    assert "let healthSubMode = 'hpm';" in js            # 健康診断HPMモードの項目（hpm / apply）
    assert js.index("let healthSubMode = 'hpm';") < js.index("function applyModeUI(mode)")
    assert "const showHealthApply = isHealthHpm && healthSubMode === 'apply';" in js
    assert "healthApplyCard.style.display = showHealthApply ? '' : 'none';" in js
    assert "if (showHealthApply) haLoadStatus();" in js
    assert "healthCard.style.display = showHealthHpm ? '' : 'none';" in js
    assert "#health-subtabs [data-health-sub]" in js    # 項目切替ボタンの配線
    for fn in ("function haLoadStatus", "function haRenderStatus", "function haLoadResponses",
               "function haPaintResponseTable", "function haSetForbidden", "function haShowError"):
        assert fn in js, fn
    assert "'/health_apply_status'" in js
    assert "'/health_apply_responses'" in js


def test_ha_state_is_declared_before_first_apply_mode_ui():
    """let 変数は applyModeUI の初回実行より前に宣言しないと TDZ で全ボタンが死ぬ。"""
    js = _script_js()
    assert js.index("let haState = {") < js.index("function applyModeUI(mode)")


def test_no_ha_code_in_inline_scripts():
    html = _html()
    inline = "".join(m.group(1) for m in re.finditer(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", html, re.S))
    assert "ha-" not in inline
    assert "haLoadStatus" not in inline


def test_style_has_accent_and_table_classes():
    css = _style_css()
    assert ".acc-happly" in css and ".icon-happly" in css
    assert ".ha-table" in css and ".ha-issue-error" in css


@pytest.mark.skipif(shutil.which("node") is None, reason="node が無い環境では構文検査を飛ばす")
def test_script_js_parses():
    subprocess.run(["node", "--check", str(STATIC / "script.js")], check=True, capture_output=True)
