# -*- coding: utf-8 -*-
"""健康診断申込: 年度設定JSON・鍵JSONの所在・利用許可CSV。"""

import json

import pytest

from services.health_apply import access as A
from services.health_apply import settings as S


def write_json(path, payload):
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return str(path)


def sample_config():
    return {
        "schema": 1,
        "default_year": "2027",
        "years": {
            "2027": {"spreadsheet_id": "1AbCdEfGhIjKlMnOpQrStUvWxYz", "webapp_url": "https://script.google.com/x/exec",
                     "previous_year": 2026, "label": "2027年度"},
            "2028": {"spreadsheet_id": "2ZyXwVuTsRqPoNmLkJiHgFeDcBa"},
        },
    }


def test_load_year_config_and_pick(tmp_path):
    cfg = S.load_year_config(write_json(tmp_path / "y.json", sample_config()))
    assert cfg.year_keys() == ["2028", "2027"]
    assert cfg.default_year == "2027"
    y = cfg.pick(None)
    assert (y.fiscal_year, y.previous_year, y.label) == (2027, 2026, "2027年度")
    assert y.webapp_url.endswith("/exec")
    assert y.spreadsheet_id_tail == "…UvWxYz"
    y2 = cfg.pick("2028")
    assert (y2.previous_year, y2.label, y2.webapp_url) == (2027, "2028年度", "")
    assert cfg.mtime


def test_pick_unknown_year_lists_known_years(tmp_path):
    cfg = S.load_year_config(write_json(tmp_path / "y.json", sample_config()))
    with pytest.raises(S.SettingsError, match="2026.*登録済み: 2028、2027"):
        cfg.pick("2026")


def test_missing_or_broken_json(tmp_path):
    with pytest.raises(S.SettingsError, match="ありません"):
        S.load_year_config(str(tmp_path / "nope.json"))
    bad = tmp_path / "bad.json"
    bad.write_text("{not json", encoding="utf-8")
    with pytest.raises(S.SettingsError, match="読めません"):
        S.load_year_config(str(bad))
    with pytest.raises(S.SettingsError, match="years"):
        S.load_year_config(write_json(tmp_path / "empty.json", {"years": {}}))
    with pytest.raises(S.SettingsError, match="spreadsheet_id"):
        S.load_year_config(write_json(tmp_path / "noid.json", {"years": {"2027": {}}}))
    with pytest.raises(S.SettingsError, match="default_year"):
        S.load_year_config(write_json(tmp_path / "dy.json", {"default_year": "2020", "years": {"2027": {"spreadsheet_id": "x"}}}))


def test_default_year_falls_back_to_latest(tmp_path):
    cfg = S.load_year_config(write_json(tmp_path / "y.json", {"years": {"2027": {"spreadsheet_id": "a"}, "2028": {"spreadsheet_id": "b"}}}))
    assert cfg.default_year == "2028"


def test_service_account_info_reads_only_client_email(tmp_path):
    missing = S.service_account_info(str(tmp_path / "sa.json"))
    assert missing["exists"] is False and "ありません" in missing["error"]
    key = tmp_path / "sa.json"
    key.write_text(json.dumps({"type": "service_account", "client_email": "hc@example.iam.gserviceaccount.com",
                               "private_key": "-----BEGIN PRIVATE KEY-----SECRET"}), encoding="utf-8")
    info = S.service_account_info(str(key))
    assert info == {"path": str(key), "exists": True, "client_email": "hc@example.iam.gserviceaccount.com", "error": ""}
    assert "SECRET" not in json.dumps(info)
    key.write_text(json.dumps({"foo": 1}), encoding="utf-8")
    assert "client_email" in S.service_account_info(str(key))["error"]


# --- 利用許可 -------------------------------------------------------------

def write_users(path, rows):
    lines = ["ユーザー名,表示名,備考"] + [",".join(r) for r in rows]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return str(path)


def test_access_fails_closed_when_csv_missing(tmp_path):
    ok, reason = A.check_access(str(tmp_path / "none.csv"), user="yatsu")
    assert ok is False
    assert "読めません" in reason and "yatsu" in reason


def test_access_allows_listed_user_case_and_width_insensitive(tmp_path):
    csv = write_users(tmp_path / "u.csv", [("Yatsu", "谷津", ""), ("taira", "平良", "")])
    assert A.check_access(csv, user="yatsu")[0] is True
    assert A.check_access(csv, user=" ＹＡＴＳＵ ")[0] is True
    ok, reason = A.check_access(csv, user="someone")
    assert ok is False
    assert "谷津・平良" in reason and "someone" in reason and str(csv) in reason
