# -*- coding: utf-8 -*-
"""健康診断申込: Google Sheets ゲートウェイ（HTTP を偽セッションで受けて検査する）。"""

import json
import sys

import pytest

from services.health_apply import sheets_gateway as G
from services.health_apply.schema import SHEET_AUDIT, SHEET_OPTIONS, SHEET_RESPONSES, SHEET_SETTINGS, SHEET_TARGETS
from tests.health_apply_fixtures import FakeSheetsGateway, workbook


class FakeResponse:
    def __init__(self, status=200, payload=None, text=""):
        self.status_code = status
        self._payload = payload
        self.text = text

    def json(self):
        if self._payload is None:
            raise ValueError("no json")
        return self._payload


class FakeSession:
    def __init__(self, responses):
        self.responses = list(responses)
        self.requests = []

    def request(self, method, url, params=None, json=None, timeout=None):
        self.requests.append({"method": method, "url": url, "params": params, "json": json, "timeout": timeout})
        item = self.responses.pop(0)
        if isinstance(item, Exception):
            raise item
        return item


def gateway(monkeypatch, *responses, email="hc@example.iam.gserviceaccount.com"):
    gw = G.GoogleSheetsGateway("SHEET123", "unused.json", timeout=7)
    session = FakeSession(responses)
    gw.client_email = email
    monkeypatch.setattr(gw, "_open_session", lambda: session)
    return gw, session


def test_module_does_not_import_google_at_module_level():
    """google-auth 未導入でも import が通る（関数内 import の担保）。"""
    assert "services.health_apply.sheets_gateway" in sys.modules
    assert not any(name in ("google", "service_account", "AuthorizedSession") for name in vars(G))


def test_broken_key_file_is_config_error(tmp_path):
    """鍵JSONはあるがサービスアカウントの形ではない → 設定不備として止める（通信しない）。"""
    pytest.importorskip("google.oauth2.service_account")
    key = tmp_path / "sa.json"
    key.write_text(json.dumps({"type": "service_account"}), encoding="utf-8")
    gw = G.GoogleSheetsGateway("SHEET123", str(key))
    with pytest.raises(G.GatewayConfigError, match="読めません"):
        gw.get_metadata()


def test_missing_key_file_is_config_error(tmp_path):
    gw = G.GoogleSheetsGateway("SHEET123", str(tmp_path / "no.json"))
    with pytest.raises(G.GatewayConfigError, match="鍵JSONがありません"):
        gw.get_metadata()
    gw2 = G.GoogleSheetsGateway("", str(tmp_path / "no.json"))
    with pytest.raises(G.GatewayConfigError, match="スプレッドシートIDが空"):
        gw2.get_metadata()


def test_quote_range_encodes_sheet_title_and_range():
    assert G.quote_range("対象者", "A:Z") == "%27%E5%AF%BE%E8%B1%A1%E8%80%85%27%21A%3AZ"
    assert G.quote_range("it's") == "%27it%27%27s%27"


def test_get_metadata_shapes_titles(monkeypatch):
    gw, session = gateway(monkeypatch, FakeResponse(200, {
        "properties": {"title": "2027 健診"},
        "sheets": [{"properties": {"title": "設定", "sheetId": 1, "gridProperties": {"rowCount": 20, "columnCount": 3}}},
                   {"properties": {"title": "対象者", "sheetId": 2}}],
    }))
    meta = gw.get_metadata()
    assert meta["title"] == "2027 健診"
    assert G.sheet_titles(meta) == ["設定", "対象者"]
    assert meta["sheets"][0]["rowCount"] == 20
    req = session.requests[0]
    assert req["method"] == "GET" and req["url"] == f"{G.BASE}/SHEET123"
    assert req["params"]["fields"] == "properties.title,sheets.properties"
    assert req["timeout"] == 7


def test_read_values_uses_formatted_values_and_stringifies(monkeypatch):
    gw, session = gateway(monkeypatch, FakeResponse(200, {"values": [["キー", "値"], ["年度", 2027], [None, "x"]]}))
    values = gw.read_values("設定", "A:C")
    assert values == [["キー", "値"], ["年度", "2027"], ["", "x"]]
    req = session.requests[0]
    assert req["url"] == f"{G.BASE}/SHEET123/values/" + G.quote_range("設定", "A:C")
    assert req["params"] == {"valueRenderOption": "FORMATTED_VALUE", "dateTimeRenderOption": "FORMATTED_STRING"}


def test_append_rows_is_raw_insert_and_only_to_writable_sheets(monkeypatch):
    gw, session = gateway(monkeypatch, FakeResponse(200, {"updates": {"updatedRows": 2}}))
    n = gw.append_rows("対象者", [["2027", "0301619", None], ["2027", "13X5035440", "x"]])
    assert n == 2
    req = session.requests[0]
    assert req["method"] == "POST"
    assert req["url"].endswith(G.quote_range("対象者", "A1") + ":append")
    assert req["params"] == {"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"}
    assert req["json"] == {"values": [["2027", "0301619", ""], ["2027", "13X5035440", "x"]]}
    assert "0301619" in json.dumps(req["json"])  # 文字列のまま


@pytest.mark.parametrize("sheet", [SHEET_RESPONSES, SHEET_OPTIONS, SHEET_SETTINGS, "別のシート"])
def test_append_to_forbidden_sheet_never_touches_http(monkeypatch, sheet):
    gw, session = gateway(monkeypatch, FakeResponse(200, {}))
    with pytest.raises(G.GatewayError, match="許可されていません"):
        gw.append_rows(sheet, [["x"]])
    assert session.requests == []


def test_append_empty_rows_is_noop(monkeypatch):
    gw, session = gateway(monkeypatch)
    assert gw.append_rows("監査ログ", []) == 0
    assert session.requests == []


def test_http_errors_are_translated(monkeypatch):
    gw, _ = gateway(monkeypatch, FakeResponse(403, {"error": {"message": "The caller does not have permission"}}))
    with pytest.raises(G.GatewayError, match="403.*hc@example.iam.gserviceaccount.com.*共有"):
        gw.get_metadata()
    gw, _ = gateway(monkeypatch, FakeResponse(404, {"error": {"message": "Requested entity was not found."}}))
    with pytest.raises(G.GatewayError, match="404.*spreadsheet_id"):
        gw.get_metadata()
    gw, _ = gateway(monkeypatch, FakeResponse(429, {"error": {"message": "Quota exceeded"}}))
    with pytest.raises(G.GatewayError, match="HTTP 429: Quota exceeded"):
        gw.get_metadata()
    gw, _ = gateway(monkeypatch, FakeResponse(500, None, text="<html>oops</html>"))
    with pytest.raises(G.GatewayError, match="HTTP 500"):
        gw.get_metadata()
    gw, _ = gateway(monkeypatch, FakeResponse(200, None, text="not json"))
    with pytest.raises(G.GatewayError, match="JSON 以外"):
        gw.get_metadata()
    gw, _ = gateway(monkeypatch, ConnectionError("dns"))
    with pytest.raises(G.GatewayError, match="接続できません"):
        gw.get_metadata()


# --- FakeSheetsGateway 自体の振る舞い（他のテストが依存する） ----------------

def test_fake_gateway_trims_trailing_blanks_and_guards_writes():
    fake = FakeSheetsGateway(workbook())
    assert G.sheet_titles(fake.get_metadata()) == [SHEET_SETTINGS, SHEET_OPTIONS, SHEET_TARGETS, SHEET_RESPONSES, SHEET_AUDIT]
    header_only = fake.read_values(SHEET_TARGETS, "1:1")
    assert len(header_only) == 1 and header_only[0][0] == "年度"
    fake.sheets[SHEET_AUDIT].append(["2027-01-01T00:00:00", "SETUP", "", "", "", "", ""])
    assert fake.read_values(SHEET_AUDIT)[-1] == ["2027-01-01T00:00:00", "SETUP"]
    with pytest.raises(G.GatewayError):
        fake.append_rows(SHEET_RESPONSES, [["x"]])
    assert fake.append_rows(SHEET_TARGETS, [["2027", "2099002"]]) == 1
    assert fake.appended == [(SHEET_TARGETS, [["2027", "2099002"]])]
