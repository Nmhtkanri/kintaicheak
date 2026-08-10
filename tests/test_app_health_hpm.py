# -*- coding: utf-8 -*-
"""健康診断HPMモードのルートテスト

CSVが実際にできる経路と、できてはいけない経路の両方を固定する。
健診結果は要配慮個人情報なので、一時ファイルが共有フォルダ側に落ちないことも
ここで見張る（launcher.py が作業フォルダを共有NASにするため事故りやすい）。
"""

from __future__ import annotations

import csv
import io
import json
import os
import sys
import time
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from config import Config  # noqa: E402
from tests.health_hpm_fixtures import (  # noqa: E402
    bp_items,
    employees_stub,
    item,
    make_master_xlsx,
    make_v2_workbook,
    person,
)

DOYUKAI = "医療法人社団 同友会 春日クリニック"
COURSE_VALUE = "人間ドックＣ　胃カメラ（４０歳以上）"


def sample_persons():
    """jinjerと自動一致する2名（employees_stub と生年月日・性別が合う）。"""
    return [
        person("友納 英彦", age=58, gender="男性", exam_date="2026-07-01 00:00:00",
               exam_no="132", sheet="P02_友納英彦", items=[
                   item("身体計測", "身長", "181.1", unit="cm", occurrence=1),
                   item("身体計測", "体重", "89.3", unit="kg", occurrence=1),
                   *bp_items(sys1="132", dia1="86", sys2="118", dia2="72"),
                   item("尿検査", "尿蛋白", "(-)", occurrence=1, value_type="定性"),
                   item("尿検査", "尿糖", "(-)", occurrence=1, value_type="定性"),
               ]),
        person("高橋 和紀", age=47, gender="男性", exam_date="2026-07-01 00:00:00",
               exam_no="186", sheet="P03_高橋和紀", items=[
                   item("身体計測", "身長", "166.5", unit="cm", occurrence=1),
                   *bp_items(sys1="114", dia1="58"),
               ]),
    ]


@pytest.fixture
def env(tmp_path, monkeypatch):
    """Configのパスを全部tmpへ逃がし、jinjerはスタブに差し替える。"""
    shared = tmp_path / "shared_sessions"      # 共有NAS側（ここに健診データが出たらNG）
    local = tmp_path / "local_health"          # ローカル側（ここに出るのが正しい）
    output = tmp_path / "健康診断"
    shared.mkdir()
    monkeypatch.setattr(Config, "SHIFT_SESSION_FOLDER", str(shared))
    monkeypatch.setattr(Config, "HEALTH_HPM_SESSION_DIR", str(local))
    monkeypatch.setattr(Config, "HEALTH_HPM_OUTPUT_BASE", str(output))
    monkeypatch.setattr(Config, "HEALTH_HPM_MASTER_XLSX",
                        make_master_xlsx(tmp_path / "master.xlsx"))
    monkeypatch.setattr(app_module, "fetch_employees_for_health", employees_stub)
    app_module.app.config["TESTING"] = True
    return {"shared": shared, "local": local, "output": output, "tmp": tmp_path}


@pytest.fixture
def client(env):
    with app_module.app.test_client() as c:
        yield c


def upload(client, path, **form):
    with open(path, "rb") as f:
        data = {"health_excel": (io.BytesIO(f.read()), os.path.basename(path))}
    data.update(form)
    return client.post("/health_hpm_preview", data=data,
                       content_type="multipart/form-data")


def generate(client, session_id, persons, *, confirmed=True, **extra):
    payload = {"session_id": session_id, "genpyo_confirmed": confirmed,
               "persons": persons}
    payload.update(extra)
    return client.post("/health_hpm_generate", data=json.dumps(payload),
                       content_type="application/json")


def selections(preview, institution=DOYUKAI, course=COURSE_VALUE):
    return [{"key": p["key"], "employee_id": p["jinjer"]["employee_id"],
             "institution": institution, "course": course}
            for p in preview["persons"]]


# ---------------------------------------------------------------------------
# プレビュー
# ---------------------------------------------------------------------------

class TestPreview:
    def test_happy_path(self, client, env):
        path = make_v2_workbook(env["tmp"] / "v2.xlsx", sample_persons())
        res = upload(client, path)
        assert res.status_code == 200

        data = res.get_json()
        assert data["success"] is True
        assert data["session_id"].startswith("health_")
        assert data["can_generate"] is True
        assert data["counts"] == {"persons": 2, "errors": 0, "warnings": 1}

        first = data["persons"][0]
        assert first["name"] == "友納 英彦"
        assert first["exam_no"] == "000132"
        assert first["jinjer"]["status"] == "ok"
        assert first["jinjer"]["employee_id"] == "2018013"
        assert first["blood_pressure"] == {
            "r1": {"sys": "132", "dia": "86"},
            "r2": {"sys": "118", "dia": "72"},
            "r3_present": False,
        }
        assert [q["item"] for q in first["qualitative"]] == ["尿蛋白", "尿糖"]
        assert first["qualitative"][0]["hpm_col"] == 54

    def test_blood_pressure_payload_has_no_average_field(self, client, env):
        path = make_v2_workbook(env["tmp"] / "v2.xlsx", sample_persons())
        bp = upload(client, path).get_json()["persons"][0]["blood_pressure"]
        assert set(bp) == {"r1", "r2", "r3_present"}, "平均を入れる枠を作らない"

    def test_master_and_roster_for_dropdowns(self, client, env):
        path = make_v2_workbook(env["tmp"] / "v2.xlsx", sample_persons())
        data = upload(client, path).get_json()

        names = [i["name"] for i in data["master"]["institutions"]]
        assert DOYUKAI in names
        unconfirmed = next(i for i in data["master"]["institutions"]
                           if i["name"] == "未確認クリニック")
        assert unconfirmed["hpm_confirmed"] is False
        assert all(len(r["employee_id"]) == 7 for r in data["roster"])

    def test_rejects_non_xlsx(self, client, env):
        bad = env["tmp"] / "x.csv"
        bad.write_text("a,b", encoding="utf-8")
        res = upload(client, bad)
        assert res.status_code == 400
        assert "xlsx" in res.get_json()["errors"][0]

    def test_rejects_missing_file(self, client):
        res = client.post("/health_hpm_preview", data={},
                          content_type="multipart/form-data")
        assert res.status_code == 400

    def test_master_missing_is_400(self, client, env, monkeypatch):
        monkeypatch.setattr(Config, "HEALTH_HPM_MASTER_XLSX",
                            str(env["tmp"] / "ない.xlsx"))
        path = make_v2_workbook(env["tmp"] / "v2.xlsx", sample_persons())
        res = upload(client, path)
        assert res.status_code == 400
        assert "見つかりません" in res.get_json()["errors"][0]

    def test_jinjer_error_is_500(self, client, env, monkeypatch):
        from services.jinjer_api_client import JinjerAPIError

        def boom():
            raise JinjerAPIError("トークン取得に失敗")

        monkeypatch.setattr(app_module, "fetch_employees_for_health", boom)
        path = make_v2_workbook(env["tmp"] / "v2.xlsx", sample_persons())
        res = upload(client, path)
        assert res.status_code == 500
        assert "jinjer" in res.get_json()["errors"][0]

    def test_v1_workbook_previews_but_blocks_generation(self, client, env):
        path = make_v2_workbook(env["tmp"] / "v1.xlsx", sample_persons(),
                                schema_version=None, genpyo=False, bp_kept=False,
                                v2_columns=False)
        data = upload(client, path).get_json()

        assert data["success"] is True
        assert len(data["persons"]) == 2, "プレビューは出せる"
        assert data["can_generate"] is False
        codes = {i["code"] for i in data["workbook_issues"]}
        assert "SCHEMA_TOO_OLD" in codes


# ---------------------------------------------------------------------------
# 一時ファイルの置き場所（要配慮個人情報のガード）
# ---------------------------------------------------------------------------

class TestSessionIsolation:
    def test_nothing_written_to_shared_folder(self, client, env):
        path = make_v2_workbook(env["tmp"] / "v2.xlsx", sample_persons())
        upload(client, path)

        assert list(env["shared"].iterdir()) == [], "共有フォルダに健診データを置かない"
        assert list(env["local"].glob("health_*.pkl")), "ローカル側にセッションができる"

    def test_uploaded_xlsx_is_removed(self, client, env):
        path = make_v2_workbook(env["tmp"] / "v2.xlsx", sample_persons())
        upload(client, path)
        assert list(env["local"].glob("*.xlsx")) == [], "アップロードした実体は残さない"

    def test_old_sessions_are_swept(self, client, env):
        env["local"].mkdir(parents=True, exist_ok=True)
        old = env["local"] / ("health_" + "a" * 32 + ".pkl")
        other = env["local"] / "keep_me.txt"
        old.write_bytes(b"x")
        other.write_bytes(b"x")
        stale = time.time() - (Config.HEALTH_HPM_SESSION_MAX_AGE_HOURS + 1) * 3600
        os.utime(old, (stale, stale))
        os.utime(other, (stale, stale))

        path = make_v2_workbook(env["tmp"] / "v2.xlsx", sample_persons())
        upload(client, path)

        assert not old.exists(), "古い健診セッションは消す"
        assert other.exists(), "健診モード以外のファイルには触らない"

    @pytest.mark.parametrize("session_id", [
        "../../.env", "health_xx", "", None, "health_" + "z" * 32,
        "abc123", "health_" + "a" * 31,
    ])
    def test_bad_session_id_rejected(self, client, session_id):
        res = generate(client, session_id, [])
        assert res.status_code == 400

    def test_shift_session_id_rejected(self, client, env):
        """スケジュールモードのセッションIDでは動かない（kindガード）。"""
        import pickle
        import uuid

        env["local"].mkdir(parents=True, exist_ok=True)
        sid = "health_" + uuid.uuid4().hex
        (env["local"] / f"{sid}.pkl").write_bytes(
            pickle.dumps({"kind": "shift", "code_sheets": []}))

        res = generate(client, sid, [])
        assert res.status_code == 400
        assert "やり直して" in res.get_json()["errors"][0]


# ---------------------------------------------------------------------------
# CSV生成
# ---------------------------------------------------------------------------

class TestGenerate:
    def _preview(self, client, env, persons=None):
        path = make_v2_workbook(env["tmp"] / "v2.xlsx", persons or sample_persons())
        return upload(client, path).get_json()

    def test_happy_path_writes_verified_csv(self, client, env):
        preview = self._preview(client, env)
        res = generate(client, preview["session_id"], selections(preview))
        assert res.status_code == 200, res.get_json()

        data = res.get_json()
        assert data["success"] is True
        assert data["row_count"] == 2
        assert data["column_count"] == 302
        assert data["verified"] is True

        out = data["output_path"]
        assert out.endswith(".csv")
        assert os.path.join("2026", "2026年度健康診断受診者結果", "CSV格納") in out

        raw = open(out, "rb").read()
        assert raw[:3] != b"\xef\xbb\xbf"
        assert raw.endswith(b"\r\n")
        rows = list(csv.reader(io.StringIO(raw.decode("cp932"), newline="")))
        assert len(rows) == 3 and all(len(r) == 302 for r in rows)

        first = rows[1]
        assert first[6] == "友納　英彦"
        assert first[7] == "ﾄﾓﾉｳ ﾋﾃﾞﾋｺ"
        assert first[8] == "19680413"
        assert first[9] == "男"
        assert first[19] == "20260701"
        assert first[20] == "000132", "受診番号のゼロ埋めがファイル上で保たれる"
        assert first[21] == "2"
        assert first[23] == "1310528885"

    def test_blood_pressure_written_without_average(self, client, env):
        preview = self._preview(client, env)
        out = generate(client, preview["session_id"],
                       selections(preview)).get_json()["output_path"]
        rows = list(csv.reader(io.StringIO(
            open(out, "rb").read().decode("cp932"), newline="")))

        assert rows[1][50:54] == ["132", "86", "118", "72"]
        assert "125" not in rows[1] and "79" not in rows[1], "平均を書かない"
        assert rows[2][50:54] == ["114", "58", "", ""], "1回分なら2回目は空欄"

    def test_qualitative_and_judgement_columns(self, client, env):
        preview = self._preview(client, env)
        out = generate(client, preview["session_id"],
                       selections(preview)).get_json()["output_path"]
        rows = list(csv.reader(io.StringIO(
            open(out, "rb").read().decode("cp932"), newline="")))

        assert rows[1][54] == "(-)" and rows[1][55] == "(-)"
        assert all(rows[1][c] == "" for c in range(183, 198)), "判定列は空欄"

    def test_session_is_dropped_after_success(self, client, env):
        preview = self._preview(client, env)
        sid = preview["session_id"]
        generate(client, sid, selections(preview))

        assert not (env["local"] / f"{sid}.pkl").exists()
        again = generate(client, sid, selections(preview))
        assert again.status_code == 400, "同じセッションで二重生成できない"

    def test_refuses_existing_file(self, client, env):
        preview = self._preview(client, env)
        out = generate(client, preview["session_id"],
                       selections(preview)).get_json()["output_path"]
        before = open(out, "rb").read()

        preview2 = self._preview(client, env)
        res = generate(client, preview2["session_id"], selections(preview2),
                       output_filename=os.path.basename(out))
        assert res.status_code == 400
        assert "上書きはしません" in res.get_json()["errors"][0]
        assert open(out, "rb").read() == before

    def test_requires_genpyo_confirmation(self, client, env):
        preview = self._preview(client, env)
        res = generate(client, preview["session_id"], selections(preview),
                       confirmed=False)
        assert res.status_code == 400
        assert "原票どおり" in res.get_json()["errors"][0]

    def test_unconfirmed_institution_blocked(self, client, env):
        preview = self._preview(client, env)
        res = generate(client, preview["session_id"],
                       selections(preview, institution="未確認クリニック", course="10"))
        assert res.status_code == 400
        assert "未確認" in res.get_json()["errors"][0]

    def test_missing_institution_blocked(self, client, env):
        preview = self._preview(client, env)
        picks = selections(preview)
        picks[0]["institution"] = ""
        res = generate(client, preview["session_id"], picks)
        assert res.status_code == 400
        assert "健診機関" in res.get_json()["errors"][0]

    def test_missing_course_blocked(self, client, env):
        preview = self._preview(client, env)
        picks = selections(preview)
        picks[0]["course"] = ""
        res = generate(client, preview["session_id"], picks)
        assert res.status_code == 400
        assert "健診種別" in res.get_json()["errors"][0]

    def test_unresolved_employee_blocked(self, client, env):
        preview = self._preview(client, env)
        picks = selections(preview)
        picks[0]["employee_id"] = ""
        res = generate(client, preview["session_id"], picks)
        assert res.status_code == 400
        assert "社員番号" in res.get_json()["errors"][0]

    def test_foreign_employee_id_blocked(self, client, env):
        preview = self._preview(client, env)
        picks = selections(preview)
        picks[0]["employee_id"] = "5551234"
        res = generate(client, preview["session_id"], picks)
        assert res.status_code == 400
        assert "自社の形式" in res.get_json()["errors"][0]

    def test_v1_workbook_cannot_generate(self, client, env):
        path = make_v2_workbook(env["tmp"] / "v1.xlsx", sample_persons(),
                                schema_version=None, genpyo=False, bp_kept=False,
                                v2_columns=False)
        preview = upload(client, path).get_json()
        res = generate(client, preview["session_id"], selections(preview))

        assert res.status_code == 400
        assert any("スキーマ2.0で再整形" in e for e in res.get_json()["errors"])
        assert not list(env["output"].rglob("*.csv")), "CSVを1つも作らない"

    def test_needs_source_check_blocks(self, client, env):
        persons = [person("友納 英彦", items=[
            item("尿検査", "尿潜血", "(-)", occurrence=1, value_type="要原票確認"),
        ])]
        preview = self._preview(client, env, persons)
        res = generate(client, preview["session_id"], selections(preview))

        assert res.status_code == 400
        assert any("要原票確認" in e for e in res.get_json()["errors"])

    def test_mixed_fiscal_years_blocked(self, client, env):
        persons = sample_persons()
        # 受診日を翌年度にずらす。年齢も受診日時点のものに直さないと
        # 照合の方で先に止まってしまい、年度チェックまで届かない
        persons[1]["exam_date"] = "2027-05-01 00:00:00"
        persons[1]["age"] = 48
        preview = self._preview(client, env, persons)
        assert all(p["jinjer"]["status"] == "ok" for p in preview["persons"])

        res = generate(client, preview["session_id"], selections(preview))
        assert res.status_code == 400
        assert "年度" in res.get_json()["errors"][0]
        assert not list(env["output"].rglob("*.csv"))

    def test_console_log_is_returned(self, client, env):
        preview = self._preview(client, env)
        data = generate(client, preview["session_id"],
                        selections(preview)).get_json()
        joined = " ".join(data["console"])
        assert "読み込み" in joined and "検証" in joined

    def test_default_filename_convention(self, client, env):
        preview = self._preview(client, env)
        out = generate(client, preview["session_id"],
                       selections(preview)).get_json()["output_path"]
        assert os.path.basename(out) == "HPM取込用_同友会_20260701_2名.csv"
