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


# =============================================================================
# 対象者登録（preview → commit）
# =============================================================================

from services.health_apply.jinjer_source import EmployeeProfile, PreviousRaw  # noqa: E402
from services.jinjer_api_client import JinjerAPIError  # noqa: E402

JINJER = {
    "2099001": (EmployeeProfile("2099001", "試験 太郎", "t.shiken@nmht.co.jp", "0", "在籍"),
                PreviousRaw("履歴", "2026", "医療法人社団 同友会 春日クリニック", ["基本健診"], "2026-07-01")),
    "2099002": (EmployeeProfile("2099002", "二 号", "n2@nmht.co.jp", "0", "在籍"),
                PreviousRaw("履歴", "2026", "医療法人徳洲会 生駒市立病院", ["1日人間ドック・胃カメラ", "婦人病検査"], "2026-07-02")),
    "2099003": (EmployeeProfile("2099003", "三 号", "", "1", "退職", "2026-03-31"), PreviousRaw()),
}


def jinjer_stub(employee_ids, previous_year):
    assert previous_year == 2026
    profiles = {i: JINJER[i][0] for i in employee_ids if i in JINJER}
    previous = {i: JINJER[i][1] for i in employee_ids if i in JINJER}
    return profiles, previous, {}


@pytest.fixture
def jinjer(monkeypatch):
    monkeypatch.setattr(app_module, "fetch_health_apply_sources", jinjer_stub)


def preview(client, text, year="2027"):
    return client.post("/health_apply_targets_preview", data=json.dumps({"year": year, "employee_ids_text": text}),
                       content_type="application/json")


def commit(client, session_id, confirm):
    return client.post("/health_apply_targets_commit", data=json.dumps({"session_id": session_id, "confirm": confirm}),
                       content_type="application/json")


def local_previews(env):
    return sorted(p.name for p in env["local"].glob("happly_*.json")) if env["local"].exists() else []


@pytest.mark.parametrize("path", ["/health_apply_targets_preview", "/health_apply_targets_commit"])
def test_post_routes_are_403_without_allow_list(client, env, monkeypatch, path):
    monkeypatch.setattr(Config, "HEALTH_APPLY_ALLOWED_USERS_CSV", str(env["tmp"] / "none.csv"))
    res = client.post(path, data=json.dumps({}), content_type="application/json")
    assert res.status_code == 403 and res.get_json()["forbidden"] is True


def test_preview_requires_employee_ids(client, env, jinjer):
    res = preview(client, "abc\n")
    body = res.get_json()
    assert res.status_code == 400
    assert "1件以上" in body["errors"][0] and "abc" in body["errors"][1]
    assert local_previews(env) == []


def test_preview_returns_plan_and_saves_locally_only(client, env, jinjer):
    res = preview(client, "2099001\n2099002\n2099003\n2099009\n2099001")
    body = res.get_json()
    assert res.status_code == 200, body
    assert body["success"] is True and body["year"] == "2027" and body["fiscal_year"] == 2027
    assert body["session_id"].startswith("happly_")
    assert body["confirm_phrase"] == "REGISTER 2027 2"
    assert body["can_commit"] is False                      # 2099003（メール無し）と 2099009（jinjer無し）が blocked
    assert body["counts"] == {"input": 4, "add": 2, "unchanged": 0, "conflict": 0, "blocked": 2,
                              "warnings": 1, "input_errors": 0}
    assert [i["code"] for i in body["input_issues"]] == ["duplicate_employee_id"]
    by = {r["employee_id"]: r for r in body["rows"]}
    assert by["2099001"]["action"] == "add"
    assert by["2099001"]["previous"]["institution"]["code"] == "1310528885"
    assert by["2099002"]["previous"]["exam_type"]["code"] == "13"
    assert by["2099002"]["previous"]["extras"] == [{"code": "GYN", "name": "婦人科検診"}]
    assert by["2099003"]["action"] == "blocked" and by["2099003"]["enrollment_label"] == "退職"
    assert by["2099009"]["action"] == "blocked"
    assert len(local_previews(env)) == 1
    assert shared_files(env) == []
    assert env["gateway"].appended == []


def test_preview_jinjer_error_is_500(client, env, monkeypatch):
    def boom(ids, year):
        raise JinjerAPIError("HTTP 429")
    monkeypatch.setattr(app_module, "fetch_health_apply_sources", boom)
    res = preview(client, "2099001")
    assert res.status_code == 500 and "jinjer API エラー" in res.get_json()["errors"][0]
    assert local_previews(env) == []


def test_commit_full_flow_then_idempotent_rerun(client, env, jinjer):
    gw = env["gateway"]
    body = preview(client, "2099001\n2099002").get_json()
    assert body["can_commit"] is True and body["confirm_phrase"] == "REGISTER 2027 2"
    sid = body["session_id"]

    # 確認語違い → 何も書かない
    res = commit(client, sid, "REGISTER 2027 3")
    assert res.status_code == 400 and "REGISTER 2027 2" in res.get_json()["errors"][0]
    assert gw.appended == []

    res = commit(client, sid, " REGISTER 2027 2 ")
    body = res.get_json()
    assert res.status_code == 200, body
    assert body["success"] is True
    assert (body["added"], body["unchanged"], body["audit_rows"], body["verified"], body["missing"]) == (2, 0, 3, 2, [])
    assert [s for s, _ in gw.appended] == [S.SHEET_TARGETS, S.SHEET_AUDIT]
    target_rows = gw.appended[0][1]
    assert [r[1] for r in target_rows] == ["2099001", "2099002"]
    assert target_rows[0][:5] == ["2027", "2099001", "試験 太郎", "t.shiken@nmht.co.jp", "0"]
    assert target_rows[1][6:11] == ["0301619", "医療法人徳洲会 生駒市立病院", "13", "人間ドックC", "GYN"]
    assert all(len(r) == S.TARGET_HUB_COLUMNS for r in target_rows)
    audit = gw.appended[1][1]
    assert audit[0][1] == "REGISTER_BATCH" and audit[1][5] == "2099001" and audit[2][5] == "2099002"
    assert audit[0][3] == current_user()
    # 回答・選択肢・設定は不変、プレビューは消え、共有側には何も無い
    assert gw.sheets[S.SHEET_RESPONSES] == workbook()[S.SHEET_RESPONSES]
    assert gw.sheets[S.SHEET_OPTIONS] == workbook()[S.SHEET_OPTIONS]
    assert local_previews(env) == []
    assert shared_files(env) == []

    # 同じ貼り付けをもう一度 → 全員 unchanged・登録できない
    again = preview(client, "2099001\n2099002").get_json()
    assert again["counts"]["unchanged"] == 2 and again["counts"]["add"] == 0
    assert again["can_commit"] is False
    res = commit(client, again["session_id"], again["confirm_phrase"])
    assert res.status_code == 409
    assert len(gw.appended) == 2

    # 追記した対象者は回答読込にも見える
    rep = client.get("/health_apply_responses").get_json()
    assert rep["counts"]["targets"] == 2 and rep["counts"]["unsent"] == 2


def test_commit_rejects_unknown_expired_and_foreign_previews(client, env, jinjer):
    body = preview(client, "2099001").get_json()
    sid = body["session_id"]
    assert commit(client, "happly_" + "0" * 32, body["confirm_phrase"]).status_code == 400
    assert commit(client, "../x", body["confirm_phrase"]).status_code == 400

    path = env["local"] / f"{sid}.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    data["user"] = "someone_else"
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    res = commit(client, sid, body["confirm_phrase"])
    assert res.status_code == 403 and "別のユーザー" in res.get_json()["errors"][0]

    data["user"] = current_user()
    data["saved_at"] = "2020-01-01T00:00:00"
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    res = commit(client, sid, body["confirm_phrase"])
    assert res.status_code == 400 and "期限" in res.get_json()["errors"][0]
    assert not path.exists()
    assert env["gateway"].appended == []


def test_commit_stops_when_sheet_changed_after_preview(client, env, jinjer):
    gw = env["gateway"]
    body = preview(client, "2099001\n2099002").get_json()
    # Apps Script や別の人が同じ社員番号を先に入れた（内容が違う）
    gw.sheets[S.SHEET_TARGETS].append(target_row(社員番号="2099002", 氏名="二 号", 社用メール="other@nmht.co.jp"))
    res = commit(client, body["session_id"], body["confirm_phrase"])
    assert res.status_code == 409
    assert res.get_json()["counts"]["conflict"] == 1
    assert gw.appended == []
    assert local_previews(env) == [f"{body['session_id']}.json"]   # 期限内なので残す（作り直しは preview で）


def test_commit_reports_when_audit_write_fails(client, env, jinjer):
    gw = env["gateway"]
    body = preview(client, "2099001").get_json()
    gw.fail_append_on = S.SHEET_AUDIT
    res = commit(client, body["session_id"], body["confirm_phrase"])
    assert res.status_code == 500
    assert "対象者は追記済み" in res.get_json()["errors"][0]
    assert [s for s, _ in gw.appended] == [S.SHEET_TARGETS]
    # 再実行すると変更なしになる（冪等）
    gw.fail_append_on = None
    again = preview(client, "2099001").get_json()
    assert again["counts"] == {"input": 1, "add": 0, "unchanged": 1, "conflict": 0, "blocked": 0, "warnings": 0, "input_errors": 0}


def test_commit_with_conflict_or_blocked_is_refused(client, env, jinjer):
    body = preview(client, "2099001\n2099003").get_json()      # 2099003 は blocked
    assert body["can_commit"] is False
    res = commit(client, body["session_id"], body["confirm_phrase"])
    assert res.status_code == 409
    assert env["gateway"].appended == []
