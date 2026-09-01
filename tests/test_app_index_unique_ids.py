# -*- coding: utf-8 -*-
"""トップ画面の id が重複していないことを固定する。

2026-09-01 に経費チェックモードで id="ki-ledger-status" が2か所にあった。

  * ③-2 SAP重複の自動除外（取込済み費用シート台帳）の枠
  * 追加投入（前回の投入データを使う）の投入台帳の枠

JS はどちらも document.getElementById で引くので先頭要素（SAP側）しか取れず、
追加投入の枠は「投入台帳を読み込み中...」のまま永久に更新されないうえ、
SAP側の枠に投入台帳の文言が混ざっていた。id が被ると症状が
「動かない」ではなく「隣の枠が化ける」なので気付きにくい。ここで機械的に止める。

検査は画面に出る実物（app.test_client().get("/")）に対して行う。
"""

import collections
import pathlib
import re

import app as app_module

TEMPLATES = pathlib.Path(app_module.__file__).parent / "templates"

ID_RE = re.compile(r'id="([a-zA-Z0-9_:-]+)"')


def _html():
    return app_module.app.test_client().get("/").get_data(as_text=True)


def test_index_has_no_duplicate_ids():
    ids = ID_RE.findall(_html())
    assert ids, "id が1つも取れていない（この検査が空振りしている）"
    dups = sorted(i for i, n in collections.Counter(ids).items() if n > 1)
    assert not dups, "id が重複しています: " + " / ".join(dups)


def test_ledger_status_boxes_are_separate():
    """SAP台帳と投入台帳は別の枠。片方に寄せると相手の文言で上書きされる。"""
    html = _html()
    assert html.count('id="ki-ledger-status"') == 1
    assert html.count('id="ki-addon-ledger-status"') == 1


def test_each_loader_writes_to_its_own_box():
    """kiLoadLedger→SAP台帳、kiLoadImportLedger→投入台帳、の振り分けを固定する。"""
    src = (TEMPLATES / "index.html").read_text(encoding="utf-8")

    def body(name):
        start = src.index("function %s()" % name)
        end = src.index("\n    }", start)
        return src[start:end]

    sap = body("kiLoadLedger")
    assert "getElementById('ki-ledger-status')" in sap
    assert "ki-addon-ledger-status" not in sap
    assert "/sap_ledger_status" in sap

    addon = body("kiLoadImportLedger")
    assert "getElementById('ki-addon-ledger-status')" in addon
    assert "getElementById('ki-ledger-status')" not in addon
    assert "/keihi_import_ledger_status" in addon
