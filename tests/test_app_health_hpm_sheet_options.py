# -*- coding: utf-8 -*-
"""健康診断HPMモード: 健診申込「選択肢」シート（Google）を機関・種別のプルダウンへ追記する経路。

- 読めたら追記され、読めなくても変換マスタだけで先へ進める（HPMモードは止まらない）
- シート由来の機関・種別で CSV が作れる（列18=種別コード、列23=場所コード）
- 追加検査（婦人科検診）は CSV に書かれず、警告とログに残る
"""

from __future__ import annotations

import csv
import io
import json
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from config import Config  # noqa: E402
from services.health_apply import schema as S  # noqa: E402
from services.health_apply.sheets_gateway import GatewayConfigError  # noqa: E402
from tests.health_apply_fixtures import FakeSheetsGateway, workbook  # noqa: E402
from tests.health_hpm_fixtures import employees_stub, make_master_xlsx, make_v2_workbook  # noqa: E402
from tests.test_app_health_hpm import sample_persons  # noqa: E402

DOYUKAI = "医療法人社団 同友会 春日クリニック"
OTEMACHI = "MYメディカルクリニック 大手町"
OTEMACHI_CODE = "1310136358"


def write_year_json(path):
    payload = {"schema": 1, "default_year": "2027",
               "years": {"2027": {"spreadsheet_id": "1TestSpreadsheetIdXYZ", "webapp_url": "",
                                  "previous_year": 2026, "label": "2027年度（テスト）"}}}
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return str(path)


@pytest.fixture
def env(tmp_path, monkeypatch):
    shared = tmp_path / "shared_sessions"
    local = tmp_path / "local_health"
    output = tmp_path / "健康診断"
    shared.mkdir()
    monkeypatch.setattr(Config, "SHIFT_SESSION_FOLDER", str(shared))
    monkeypatch.setattr(Config, "HEALTH_HPM_SESSION_DIR", str(local))
    monkeypatch.setattr(Config, "HEALTH_HPM_OUTPUT_BASE", str(output))
    monkeypatch.setattr(Config, "HEALTH_HPM_MASTER_XLSX",
                        make_master_xlsx(tmp_path / "master.xlsx"))
    monkeypatch.setattr(Config, "HEALTH_APPLY_SETTINGS_JSON", write_year_json(tmp_path / "years.json"))
    monkeypatch.setattr(Config, "HEALTH_HPM_OPTIONS_SNAPSHOT_JSON", str(tmp_path / "写し" / "選択肢_写し.json"))
    (tmp_path / "写し").mkdir()
    monkeypatch.setattr(app_module, "fetch_employees_for_health", employees_stub)
    state = {"gateway": FakeSheetsGateway(workbook())}
    monkeypatch.setattr(app_module, "health_apply_gateway", lambda year: state["gateway"])
    app_module.app.config["TESTING"] = True
    return {"tmp": tmp_path, "output": output, "state": state,
            "snapshot": tmp_path / "写し" / "選択肢_写し.json"}


@pytest.fixture
def client(env):
    with app_module.app.test_client() as c:
        yield c


def preview(client, env):
    path = make_v2_workbook(env["tmp"] / "v2.xlsx", sample_persons())
    with open(path, "rb") as f:
        data = {"health_excel": (io.BytesIO(f.read()), "v2.xlsx")}
    res = client.post("/health_hpm_preview", data=data, content_type="multipart/form-data")
    assert res.status_code == 200, res.get_json()
    return res.get_json()


def generate(client, session_id, persons, **extra):
    payload = {"session_id": session_id, "genpyo_confirmed": True, "persons": persons}
    payload.update(extra)
    return client.post("/health_hpm_generate", data=json.dumps(payload),
                       content_type="application/json")


def picks(data, institution, course, extras=None):
    out = []
    for p in data["persons"]:
        pick = {"key": p["key"], "employee_id": p["jinjer"]["employee_id"],
                "institution": institution, "course": course}
        if extras is not None:
            pick["extras"] = list(extras)
        out.append(pick)
    return out


def read_csv(path):
    return list(csv.reader(io.StringIO(open(path, "rb").read().decode("cp932"), newline="")))


class TestPreview:
    def test_sheet_institutions_and_types_are_appended(self, client, env):
        data = preview(client, env)
        so = data["master"]["sheet_options"]
        assert so["loaded"] is True
        assert "2027年度（テスト）" in so["source"]
        assert so["extras"] == [{"code": "GYN", "name": "婦人科検診"}]

        insts = data["master"]["institutions"]
        names = [i["name"] for i in insts]
        assert names.count(DOYUKAI) == 1, "同じ場所コードはシート側を載せない"
        assert OTEMACHI in names
        assert "その他" not in names
        otemachi = next(i for i in insts if i["name"] == OTEMACHI)
        assert otemachi["source"] == "sheet" and otemachi["value"] == OTEMACHI_CODE
        assert [c["hpm_value"] for c in otemachi["courses"]] == ["10", "11", "12", "13", "14", "15"]

        doyukai = next(i for i in insts if i["name"] == DOYUKAI)
        assert doyukai["source"] == "master" and doyukai["value"] == DOYUKAI
        assert [c["hpm_value"] for c in doyukai["courses"]][:2] == ["人間ドックＣ　胃カメラ（４０歳以上）", "10"]
        assert "13" in [c["hpm_value"] for c in doyukai["courses"]], "種別はマスタ機関にも追記"

    def test_only_options_sheet_is_read(self, client, env):
        preview(client, env)
        sheets_read = [c[1] for c in env["state"]["gateway"].calls if c[0] == "read_values"]
        assert sheets_read == [S.SHEET_OPTIONS], "対象者・回答（個人情報）は読まない"
        assert not env["state"]["gateway"].appended

    def test_missing_year_json_falls_back_to_master(self, client, env, monkeypatch):
        monkeypatch.setattr(Config, "HEALTH_APPLY_SETTINGS_JSON", str(env["tmp"] / "no.json"))
        data = preview(client, env)
        so = data["master"]["sheet_options"]
        assert so["loaded"] is False and so["error"]
        assert so["extras"] == [{"code": "GYN", "name": "婦人科検診"}], "追加検査の既定は残す"
        assert all(i["source"] == "master" for i in data["master"]["institutions"])
        assert data["can_generate"] is True

    def test_gateway_config_error_falls_back(self, client, env, monkeypatch):
        def broken(year):
            raise GatewayConfigError("鍵ファイルがありません（テスト）")
        monkeypatch.setattr(app_module, "health_apply_gateway", broken)
        so = preview(client, env)["master"]["sheet_options"]
        assert so["loaded"] is False
        assert "鍵ファイル" in so["error"]

    def test_read_failure_falls_back(self, client, env):
        env["state"]["gateway"].fail_read = True
        so = preview(client, env)["master"]["sheet_options"]
        assert so["loaded"] is False
        assert "読み取りに失敗" in so["error"]

    def test_success_writes_snapshot_for_other_pcs(self, client, env):
        assert not env["snapshot"].exists()
        preview(client, env)
        payload = json.loads(env["snapshot"].read_text(encoding="utf-8"))
        assert payload["schema"] == 1
        assert payload["saved_at"]
        assert [i["code"] for i in payload["institutions"]][:1] == ["1310528885"]
        assert "OTHER" not in [i["code"] for i in payload["institutions"]]
        assert [t["code"] for t in payload["exam_types"]] == ["10", "11", "12", "13", "14", "15"]
        assert payload["extras"] == [{"code": "GYN", "name": "婦人科検診"}]
        text = env["snapshot"].read_text(encoding="utf-8")
        assert "試験 太郎" not in text and "@" not in text, "写しに個人情報を入れない"

    def test_no_key_pc_reads_snapshot(self, client, env, monkeypatch):
        preview(client, env)                     # 鍵のあるPCが写しを書く
        assert env["snapshot"].exists()

        def no_key(year):
            raise GatewayConfigError("サービスアカウントの鍵JSONがありません（テスト）")
        monkeypatch.setattr(app_module, "health_apply_gateway", no_key)
        data = preview(client, env)
        so = data["master"]["sheet_options"]
        assert so["loaded"] is True
        assert "写し" in so["source"] and so["saved_at"]
        assert "鍵" in so["note"] and "鍵JSONがありません" in so["note"]
        names = [i["name"] for i in data["master"]["institutions"]]
        assert OTEMACHI in names
        otemachi = next(i for i in data["master"]["institutions"] if i["name"] == OTEMACHI)
        assert [c["hpm_value"] for c in otemachi["courses"]] == ["10", "11", "12", "13", "14", "15"]

    def test_no_key_pc_can_generate_from_snapshot(self, client, env, monkeypatch):
        preview(client, env)
        def no_key(year):
            raise GatewayConfigError("鍵なし")
        monkeypatch.setattr(app_module, "health_apply_gateway", no_key)
        data = preview(client, env)
        res = generate(client, data["session_id"], picks(data, OTEMACHI_CODE, "12"))
        assert res.status_code == 200, res.get_json()
        rows = read_csv(res.get_json()["output_path"])
        assert rows[1][18] == "12" and rows[1][23] == OTEMACHI_CODE

    def test_failed_read_does_not_overwrite_snapshot(self, client, env):
        preview(client, env)
        before = env["snapshot"].read_bytes()
        env["state"]["gateway"].fail_read = True
        so = preview(client, env)["master"]["sheet_options"]
        assert so["loaded"] is True and "写し" in so["source"]
        assert env["snapshot"].read_bytes() == before

    def test_broken_snapshot_falls_back(self, client, env, monkeypatch):
        env["snapshot"].write_text("{not json", encoding="utf-8")
        monkeypatch.setattr(Config, "HEALTH_APPLY_SETTINGS_JSON", str(env["tmp"] / "no.json"))
        so = preview(client, env)["master"]["sheet_options"]
        assert so["loaded"] is False

    def test_unwritable_snapshot_does_not_break_preview(self, client, env, monkeypatch):
        monkeypatch.setattr(Config, "HEALTH_HPM_OPTIONS_SNAPSHOT_JSON",
                            str(env["tmp"] / "no_such_dir" / "x.json"))
        data = preview(client, env)
        assert data["master"]["sheet_options"]["loaded"] is True

    def test_bad_header_falls_back(self, client, env):
        wb = workbook()
        wb[S.SHEET_OPTIONS][0][1] = "code"
        env["state"]["gateway"] = FakeSheetsGateway(wb)
        so = preview(client, env)["master"]["sheet_options"]
        assert so["loaded"] is False and "選択肢シート" in so["error"]


class TestGenerate:
    def test_sheet_institution_and_type_go_into_csv(self, client, env):
        data = preview(client, env)
        res = generate(client, data["session_id"], picks(data, OTEMACHI_CODE, "13"))
        assert res.status_code == 200, res.get_json()
        body = res.get_json()
        rows = read_csv(body["output_path"])
        assert rows[1][18] == "13" and rows[2][18] == "13"
        assert rows[1][23] == OTEMACHI_CODE
        assert rows[1][50:54] == ["132", "86", "118", "72"], "血圧はそのまま"
        assert "MYメディカルクリニック" in os.path.basename(body["output_path"])
        assert any("選択肢シート" in line for line in body["console"])

    def test_generic_type_on_master_institution(self, client, env):
        data = preview(client, env)
        res = generate(client, data["session_id"], picks(data, DOYUKAI, "12"))
        assert res.status_code == 200, res.get_json()
        rows = read_csv(res.get_json()["output_path"])
        assert rows[1][18] == "12" and rows[1][23] == "1310528885"

    def test_extras_not_written_but_warned(self, client, env):
        data = preview(client, env)
        res = generate(client, data["session_id"], picks(data, OTEMACHI_CODE, "10", extras=["GYN"]))
        assert res.status_code == 200, res.get_json()
        body = res.get_json()
        rows = read_csv(body["output_path"])
        assert "GYN" not in rows[1] and "婦人科検診" not in rows[1]
        assert rows[1][18] == "10"
        warn = [w for w in body["warnings"] if w.get("code") == "EXTRA_NOT_IN_CSV"]
        assert len(warn) == 1
        assert "友納 英彦（婦人科検診）" in warn[0]["message"]
        assert "高橋 和紀（婦人科検診）" in warn[0]["message"]
        assert any("追加検査" in line for line in body["console"])

    def test_no_extras_no_warning(self, client, env):
        data = preview(client, env)
        body = generate(client, data["session_id"], picks(data, OTEMACHI_CODE, "10", extras=[])).get_json()
        assert not [w for w in body["warnings"] if w.get("code") == "EXTRA_NOT_IN_CSV"]
        assert not any("追加検査" in line for line in body["console"])

    def test_unknown_extra_is_400(self, client, env):
        data = preview(client, env)
        res = generate(client, data["session_id"], picks(data, OTEMACHI_CODE, "10", extras=["XYZ"]))
        assert res.status_code == 400
        assert any("追加検査のコードが不明" in e for e in res.get_json()["errors"])

    def test_unknown_type_is_400(self, client, env):
        data = preview(client, env)
        res = generate(client, data["session_id"], picks(data, OTEMACHI_CODE, "99"))
        assert res.status_code == 400
        assert any("健診種別を選んでください" in e for e in res.get_json()["errors"])

    def test_other_code_is_400(self, client, env):
        data = preview(client, env)
        res = generate(client, data["session_id"], picks(data, "OTHER", "10"))
        assert res.status_code == 400
        assert any("健診機関を選んでください" in e for e in res.get_json()["errors"])

    def test_generate_uses_options_from_preview_time(self, client, env):
        """生成のたびに Google を叩かない。読み込み後にシートが読めなくなっても作れる。"""
        data = preview(client, env)
        env["state"]["gateway"].fail_read = True
        calls_before = len(env["state"]["gateway"].calls)
        res = generate(client, data["session_id"], picks(data, OTEMACHI_CODE, "13"))
        assert res.status_code == 200, res.get_json()
        assert len(env["state"]["gateway"].calls) == calls_before

    def test_fallback_session_still_generates_with_master(self, client, env, monkeypatch):
        monkeypatch.setattr(Config, "HEALTH_APPLY_SETTINGS_JSON", str(env["tmp"] / "no.json"))
        data = preview(client, env)
        ok = generate(client, data["session_id"], picks(data, DOYUKAI, "10", extras=["GYN"]))
        assert ok.status_code == 200, ok.get_json()
        assert any("変換マスタのみ" in line for line in ok.get_json()["console"])
        # シートが無いので、シート由来の機関・種別は使えない
        data2 = preview(client, env)
        ng = generate(client, data2["session_id"], picks(data2, OTEMACHI_CODE, "13"))
        assert ng.status_code == 400


def test_ui_wiring():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    js = open(os.path.join(root, "static", "script.js"), encoding="utf-8").read()
    html = open(os.path.join(root, "templates", "index.html"), encoding="utf-8").read()

    assert "function hhExtrasHtml(person, master)" in js
    assert "function hhIsFemale(person)" in js
    assert "=== '女性'" in js
    assert "class=\"hh-extra\"" in js
    assert "extras: extras," in js                       # 送信値に追加検査
    assert "optgroup label=\"健診申込の選択肢シート\"" in js
    assert "function hhInstitutionFilterHtml(cls, key, id)" in js   # 機関の絞り込み入力
    assert "class=\"hh-inst-filter " in js
    assert "e.target.classList.contains('hh-inst-filter')" in js
    assert "master.sheet_options.exam_types" in js                  # 機関未選択でも種別を出す
    assert "so.note" in js                                            # 写しから読んだ旨
    assert "optgroup label=\"申込では無効の種別\"" in js
    assert "function hhSheetOptionsNote(master)" in js
    assert js.index("function hhSheetOptionsNote(master)") < js.index("function hhRenderPreview(data)")
    assert "hhExtrasHtml(person, master)" in js
    assert "健診申込の「選択肢」シート（Google）の内容を<b>追記</b>" in html
    assert "女性なら既定でオン" in html
