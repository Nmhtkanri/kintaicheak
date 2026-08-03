# -*- coding: utf-8 -*-
"""/travel_expense_members ルート（移動交通費対象者リストのUI表示・編集）のテスト"""
import pytest

import app as app_module
from config import Config
from services.expense_check import load_travel_expense_members


@pytest.fixture
def client(tmp_path, monkeypatch):
    """リストCSVのパスを一時フォルダへ差し替えた Flask テストクライアント"""
    csv_path = tmp_path / "移動交通費対象者.csv"
    monkeypatch.setattr(Config, "KEIHI_TRAVEL_EXPENSE_MEMBERS_CSV", str(csv_path))
    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as c:
        yield c, csv_path


def test_get_missing_file_returns_empty(client):
    c, csv_path = client
    res = c.get("/travel_expense_members")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert data["exists"] is False
    assert data["members"] == []
    assert str(csv_path) == data["path"]


def test_post_then_get_roundtrip(client):
    c, csv_path = client
    res = c.post("/travel_expense_members", json={"members": [
        {"id": " 2018017 ", "name": "中村 淳一"},
        {"id": "2026001", "name": "佐久間歩"},
        {"id": "", "name": ""},                       # 空行は無視される
    ]})
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True and data["count"] == 2
    assert data["backup"] == ""                       # 新規作成
    assert load_travel_expense_members(csv_path) == {
        "2018017": "中村 淳一", "2026001": "佐久間歩"}

    res2 = c.get("/travel_expense_members")
    got = res2.get_json()
    assert got["exists"] is True
    assert got["members"] == [
        {"id": "2018017", "name": "中村 淳一"},
        {"id": "2026001", "name": "佐久間歩"},
    ]

    # 2回目の保存はバックアップ付き
    res3 = c.post("/travel_expense_members", json={"members": [
        {"id": "2018017", "name": "中村 淳一"}]})
    assert res3.get_json()["backup"]


def test_post_validation_error_400(client):
    c, csv_path = client
    res = c.post("/travel_expense_members", json={"members": [
        {"id": "2018017", "name": "A"}, {"id": "2018017", "name": "重複"}]})
    assert res.status_code == 400
    assert any("重複" in e for e in res.get_json()["errors"])
    assert not csv_path.exists()                      # エラー時は書かない


def test_post_empty_members_400(client):
    c, csv_path = client
    res = c.post("/travel_expense_members", json={"members": []})
    assert res.status_code == 400
    assert "0名" in res.get_json()["errors"][0]
    assert not csv_path.exists()


def test_post_bad_payload_400(client):
    c, _ = client
    res = c.post("/travel_expense_members", json={"members": "not-a-list"})
    assert res.status_code == 400
    res2 = c.post("/travel_expense_members", data="junk",
                  content_type="application/json")
    assert res2.status_code == 400


def test_post_warning_passthrough(client):
    c, _ = client
    res = c.post("/travel_expense_members", json={"members": [
        {"id": "5000001", "name": "派遣さん"}]})
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert any("5000001" in w for w in data["warnings"])
