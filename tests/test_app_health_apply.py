# -*- coding: utf-8 -*-
"""健康診断申込モードのルートテスト。

Google・jinjer には触らず、ゲートウェイは FakeSheetsGateway、jinjer はスタブに差し替える。
権限（403）→ 年度設定（400）→ シート構成（400）→ 取得（500）の順に止まること、
個人情報の行を読む前に構成検査が終わること、共有側フォルダに何も落ちないことを固定する。
"""

from __future__ import annotations

import json
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from config import Config  # noqa: E402
from services.health_apply import schema as S  # noqa: E402
from services.health_apply.sheets_gateway import GatewayConfigError, GatewayError  # noqa: E402
from services.sap_import_ledger import current_user  # noqa: E402
from tests.health_apply_fixtures import (  # noqa: E402
    FakeSheetsGateway, response_row, target_row, workbook,
)

SPREADSHEET_ID = "1TestSpreadsheetIdXYZ"


def write_users(path, users):
    lines = ["ユーザー名,表示名,備考"] + [f"{u},{u},テスト" for u in users]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return str(path)


def write_year_json(path, **years):
    payload = {"schema": 1, "default_year": "2027",
               "years": {"2027": {"spreadsheet_id": SPREADSHEET_ID, "webapp_url": "https://script.google.com/x/exec",
                                  "previous_year": 2026, "label": "2027年度"}}}
    payload["years"].update(years)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return str(path)


@pytest.fixture
def env(tmp_path, monkeypatch):
    """Config を tmp へ逃がし、ゲートウェイを偽物に差し替える。"""
    shared = tmp_path / "shared"          # 共有NAS側の代役（ここに何か出たらNG）
    local = tmp_path / "local_sessions"   # ローカル側
    shared.mkdir()
    sa = tmp_path / "sa.json"
    sa.write_text(json.dumps({"type": "service_account", "client_email": "hc@test.iam.gserviceaccount.com",
                              "private_key": "SECRET"}), encoding="utf-8")
    monkeypatch.setattr(Config, "SHIFT_SESSION_FOLDER", str(shared))
    monkeypatch.setattr(Config, "UPLOAD_FOLDER", str(shared))
    monkeypatch.setattr(Config, "HEALTH_APPLY_ALLOWED_USERS_CSV", write_users(tmp_path / "users.csv", [current_user()]))
    monkeypatch.setattr(Config, "HEALTH_APPLY_SETTINGS_JSON", write_year_json(tmp_path / "years.json"))
    monkeypatch.setattr(Config, "HEALTH_APPLY_SERVICE_ACCOUNT_JSON", str(sa))
    monkeypatch.setattr(Config, "HEALTH_APPLY_SESSION_DIR", str(local))
    monkeypatch.setattr(Config, "HEALTH_APPLY_PREVIEW_TTL_HOURS", 2.0)
    app_module.app.config["TESTING"] = True

    state = {"gateway": FakeSheetsGateway(workbook())}
    monkeypatch.setattr(app_module, "health_apply_gateway", lambda year: state["gateway"])
    state.update({"tmp": tmp_path, "shared": shared, "local": local, "sa": sa})
    return state


@pytest.fixture
def client(env):
    with app_module.app.test_client() as c:
        yield c


def use_workbook(env, **kw):
    env["gateway"] = FakeSheetsGateway(workbook(**kw))
    return env["gateway"]


def shared_files(env):
    return [p for p in env["shared"].rglob("*") if p.is_file()]


# --- 権限 -------------------------------------------------------------------

@pytest.mark.parametrize("path", ["/health_apply_status", "/health_apply_responses"])
def test_missing_allow_list_is_403_on_every_route(client, env, monkeypatch, path):
    monkeypatch.setattr(Config, "HEALTH_APPLY_ALLOWED_USERS_CSV", str(env["tmp"] / "none.csv"))
    res = client.get(path)
    body = res.get_json()
    assert res.status_code == 403
    assert body["success"] is False and body["forbidden"] is True
    assert "読めません" in body["errors"][0]
    assert env["gateway"].calls == []


def test_user_not_listed_is_403(client, env, monkeypatch):
    monkeypatch.setattr(Config, "HEALTH_APPLY_ALLOWED_USERS_CSV", write_users(env["tmp"] / "others.csv", ["someone_else"]))
    res = client.get("/health_apply_status")
    assert res.status_code == 403
    assert "someone_else" in res.get_json()["errors"][0]
    assert env["gateway"].calls == []


# --- 年度設定・鍵 ----------------------------------------------------------------

def test_missing_year_json_is_400_but_status_still_reports_environment(client, env, monkeypatch):
    monkeypatch.setattr(Config, "HEALTH_APPLY_SETTINGS_JSON", str(env["tmp"] / "no.json"))
    res = client.get("/health_apply_status")
    body = res.get_json()
    assert res.status_code == 400
    assert "年度設定JSONがありません" in body["errors"][0]
    assert body["service_account"]["client_email"] == "hc@test.iam.gserviceaccount.com"
    assert "SECRET" not in res.get_data(as_text=True)
    assert body["user"] == current_user()


def test_unknown_year_is_400(client, env):
    res = client.get("/health_apply_status?year=2031")
    assert res.status_code == 400
    assert "2031" in res.get_json()["errors"][0]
    assert res.get_json()["years"] == ["2027"]


def test_gateway_config_error_is_400(client, env, monkeypatch):
    def broken(year):
        raise GatewayConfigError("サービスアカウントの鍵JSONがありません: X")
    monkeypatch.setattr(app_module, "health_apply_gateway", broken)
    res = client.get("/health_apply_status")
    assert res.status_code == 400
    assert "鍵JSON" in res.get_json()["errors"][0]


def test_gateway_error_is_500(client, env):
    env["gateway"].fail_read = True
    res = client.get("/health_apply_responses")
    assert res.status_code == 500
    assert "失敗" in res.get_json()["errors"][0]


# --- シート構成 -----------------------------------------------------------------

def test_header_mismatch_is_400_and_personal_rows_are_not_read(client, env):
    gw = use_workbook(env)
    gw.sheets[S.SHEET_TARGETS][0][2] = "名前"
    gw.sheets[S.SHEET_TARGETS].append(target_row())
    res = client.get("/health_apply_responses")
    body = res.get_json()
    assert res.status_code == 400
    assert body["errors"] == ["対象者シートの3列目が想定外です: 「名前」（期待: 「氏名」）"]
    assert body["workbook"]["sheets"] == list(S.ALL_SHEETS)
    # 対象者・回答は1行目しか読んでいない
    assert ("read_values", S.SHEET_TARGETS, "1:1") in gw.calls
    assert ("read_values", S.SHEET_TARGETS, "") not in gw.calls
    assert ("read_values", S.SHEET_RESPONSES, "") not in gw.calls


def test_missing_sheet_and_schema_version_are_reported(client, env):
    gw = use_workbook(env)
    del gw.sheets[S.SHEET_AUDIT]
    gw.sheets[S.SHEET_SETTINGS][1][1] = "1999.0"   # スキーマ版
    res = client.get("/health_apply_status")
    assert res.status_code == 400
    assert res.get_json()["errors"] == [
        "シートがありません: 監査ログ",
        f"設定シートのスキーマ版が違います: 「1999.0」（Hub は {S.SCHEMA_VERSION}）",
    ]


def test_bad_option_row_is_400(client, env):
    gw = use_workbook(env)
    gw.sheets[S.SHEET_OPTIONS].append(["施設", "1", "x", "1", "1", "", ""])
    res = client.get("/health_apply_status")
    assert res.status_code == 400
    assert "区分が想定外" in res.get_json()["errors"][0]


# --- status -----------------------------------------------------------------------

def test_status_ok_payload(client, env):
    res = client.get("/health_apply_status")
    body = res.get_json()
    assert res.status_code == 200, body
    assert body["success"] is True
    assert body["user"] == current_user()
    assert body["years"] == ["2027"] and body["year"] == "2027"
    assert body["settings"]["fiscal_year"] == 2027 and body["settings"]["previous_year"] == 2026
    assert body["settings"]["accept_from"] == "2027-02-01" and body["settings"]["accepting"] == "1"
    assert body["settings"]["spreadsheet_id_tail"].endswith("IdXYZ")
    assert body["workbook"]["sheets"] == list(S.ALL_SHEETS)
    assert body["workbook"]["schema_version"] == S.SCHEMA_VERSION
    assert body["catalog"] == {"institutions": 5, "exam_types": 6, "extras": 1, "relationships": 2}
    assert body["service_account"]["exists"] is True
    assert body["service_account"]["client_email"] == "hc@test.iam.gserviceaccount.com"
    assert body["settings_file"]["mtime"]
    assert body["schema_version"] == S.SCHEMA_VERSION
    assert "SECRET" not in res.get_data(as_text=True)


# --- responses --------------------------------------------------------------------

def test_responses_report(client, env):
    use_workbook(env, targets=[
        target_row(社員番号="2099001"),
        target_row(社員番号="2099002", 申込状態=S.STATUS_ANSWERED, 送信日時="2027-02-01T09:00:00",
                   受付番号="HC-2027-2099002-01", 回答版="1"),
        target_row(社員番号="2099003", 申込状態=S.STATUS_ANSWERED, 送信日時="2027-02-01T09:00:00",
                   受付番号="HC-2027-2099003-01", 回答版="1"),
    ], responses=[
        response_row(社員番号="2099002", 受付番号="HC-2027-2099002-01"),
        response_row(社員番号="2099003", 受付番号="HC-2027-2099003-01", 健診種別コード="99"),
    ])
    res = client.get("/health_apply_responses?year=2027")
    body = res.get_json()
    assert res.status_code == 200, body
    assert body["success"] is True and body["year"] == "2027" and body["fetched_at"]
    assert body["counts"] == {"targets": 3, "unsent": 1, "sent_not_accessed": 0, "accessed_only": 0,
                              "answered": 1, "reanswer_pending": 0, "invalid": 0, "error": 1}
    assert [r["employee_id"] for r in body["rows"]] == ["2099003", "2099002", "2099001"]
    assert body["rows"][0]["bucket"] == "error"
    assert body["rows"][0]["issues"][0]["code"] == "unknown_exam_type"
    assert body["rows"][1]["latest"]["institution"]["name"] == "MYメディカルクリニック 大手町"
    assert body["rows"][1]["target"]["previous"]["source"] == S.SOURCE_HISTORY
    assert body["workbook_issues"] == []
    assert body["settings"]["accept_to"] == "2027-02-28"
    # 読んだ順: 構成検査 → 対象者・回答の本体
    calls = env["gateway"].calls
    assert calls.index(("read_values", S.SHEET_TARGETS, "")) > calls.index(("read_values", S.SHEET_TARGETS, "1:1"))
    assert env["gateway"].appended == []
    assert shared_files(env) == []
    assert not env["local"].exists() or list(env["local"].iterdir()) == []


def test_responses_default_year_when_param_missing(client, env):
    res = client.get("/health_apply_responses")
    assert res.status_code == 200
    assert res.get_json()["year"] == "2027"
    assert res.get_json()["counts"]["targets"] == 0
