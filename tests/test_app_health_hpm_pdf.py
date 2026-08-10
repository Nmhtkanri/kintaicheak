# -*- coding: utf-8 -*-
"""健診PDF直接読み取りのルートテスト

Claude API と PDF レンダリングは差し替えて動かす。見張るのは主に4つ:

  1. SSEで進捗が流れ、最後の done が既存プレビューと同じ形で返ること
  2. 原票画像が「そのセッションの人にだけ」出せること（要配慮個人情報）
  3. 健診データが共有フォルダに落ちないこと・後片付けされること
  4. CSVと一緒に監査用Excelができ、それをExcel経路で読み直せること
"""

from __future__ import annotations

import io
import json
import os
import pickle
import sys
import time

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
import services.health_hpm_pdf as pdf_module  # noqa: E402
from config import Config  # noqa: E402
from tests.health_hpm_fixtures import employees_stub, make_master_xlsx  # noqa: E402

DOYUKAI = "医療法人社団 同友会 春日クリニック"
COURSE_VALUE = "人間ドックＣ　胃カメラ（４０歳以上）"


def reading(name="友納 英彦", exam_no="132", *, bp=None, **overrides):
    base = {
        "identity": {"氏名": name, "年齢": 58, "性別": "男性",
                     "受診日": "2026-07-01", "受診No": exam_no},
        "blood_pressure": bp if bp is not None else [
            {"occurrence": 1, "systolic": 132, "diastolic": 86},
            {"occurrence": 2, "systolic": 118, "diastolic": 72}],
        "metrics": [{"category": "身体計測", "item": "身長", "value": "181.1"},
                    {"category": "身体計測", "item": "体重", "value": "89.3"}],
        "qualitative": [{"category": "尿検査", "item": "尿蛋白", "value": "(-)"}],
        "judgements": {"身体計測": "B"},
        "needs_check": ["ZTTは空欄"],
    }
    base.update(overrides)
    return base


TAKAHASHI = reading("高橋 和紀", "186")
TAKAHASHI["identity"]["年齢"] = 47


def tiny_png() -> bytes:
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (40, 60), "white").save(buf, format="PNG")
    return buf.getvalue()


@pytest.fixture
def env(tmp_path, monkeypatch):
    shared = tmp_path / "shared_sessions"    # 共有NAS相当。ここに健診データが出たらNG
    local = tmp_path / "local_health"        # ローカル。ここに出るのが正しい
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
def fake_reader(monkeypatch):
    """PDFレンダリングとClaude呼び出しを差し替える。readings を後から入れ替えられる。"""
    state = {"readings": [reading(), TAKAHASHI], "pages": 2}

    monkeypatch.setattr(pdf_module, "render_pdf_pages",
                        lambda path, dpi=200: [tiny_png()] * state["pages"])
    monkeypatch.setattr(pdf_module, "page_images_for_reading", lambda png: [png])
    monkeypatch.setattr(pdf_module, "build_client", lambda: object())

    calls = {"n": 0}

    def fake_call(client, images):
        calls["n"] += 1
        value = state["readings"][calls["n"] - 1]
        if isinstance(value, Exception):
            raise value
        return value

    monkeypatch.setattr(pdf_module, "_call_claude_for_page", fake_call)
    state["calls"] = calls
    return state


@pytest.fixture
def client(env):
    with app_module.app.test_client() as c:
        yield c


def post_pdf(client, name="健診.pdf", content=b"%PDF-1.4 dummy", **form):
    data = {"health_pdf": (io.BytesIO(content), name)}
    data.update(form)
    return client.post("/health_hpm_pdf_preview", data=data,
                       content_type="multipart/form-data")


def sse_events(res):
    """SSEを最後まで読み切ってイベント一覧にする。"""
    events = []
    for part in res.get_data(as_text=True).split("\n\n"):
        kind = payload = None
        for line in part.splitlines():
            if line.startswith("event: "):
                kind = line[len("event: "):]
            elif line.startswith("data: "):
                payload = json.loads(line[len("data: "):])
        if kind:
            events.append((kind, payload))
    return events


def done_of(res):
    for kind, payload in sse_events(res):
        if kind == "done":
            return payload
    return None


def selections(done, institution=DOYUKAI, course=COURSE_VALUE):
    return [{"key": p["key"], "employee_id": p["jinjer"]["employee_id"],
             "institution": institution, "course": course}
            for p in done["persons"]]


def generate(client, session_id, persons, *, confirmed=True, **extra):
    payload = {"session_id": session_id, "genpyo_confirmed": confirmed,
               "persons": persons}
    payload.update(extra)
    return client.post("/health_hpm_generate", data=json.dumps(payload),
                       content_type="application/json")


# ---------------------------------------------------------------------------
# プレビュー（SSE）
# ---------------------------------------------------------------------------

class TestPdfPreview:
    def test_happy_path(self, client, fake_reader):
        res = post_pdf(client)
        assert res.status_code == 200
        assert res.mimetype == "text/event-stream"

        events = sse_events(res)
        kinds = [k for k, _ in events]
        assert kinds[-1] == "done"
        assert kinds.count("progress") >= 3, "ページごとに進捗が出ること"

        done = events[-1][1]
        assert done["success"] is True
        assert done["source"] == "pdf"
        assert done["session_id"].startswith("health_")
        assert done["schema_version"] == "2.0"
        assert done["can_generate"] is True
        assert done["counts"]["persons"] == 2

    def test_person_payload_has_page_image(self, client, fake_reader):
        done = done_of(post_pdf(client))
        first = done["persons"][0]

        assert first["page"] == 1
        assert first["page_image_url"] == \
            f"/health_hpm_page_image/{done['session_id']}/1"
        assert first["jinjer"]["status"] == "ok"
        assert first["exam_no"] == "000132"
        assert first["blood_pressure"] == {"r1": {"sys": "132", "dia": "86"},
                                           "r2": {"sys": "118", "dia": "72"},
                                           "r3_present": False}

    def test_needs_check_is_info_not_counted(self, client, fake_reader):
        done = done_of(post_pdf(client))
        levels = [i["level"] for p in done["persons"] for i in p["issues"]]

        assert "info" in levels, "AIの読み取りメモは残る"
        assert done["counts"]["errors"] == 0, "メモは件数に数えない"

    def test_rejects_non_pdf(self, client, fake_reader):
        res = post_pdf(client, name="健診.xlsx")
        assert res.status_code == 400
        assert "PDF" in res.get_json()["errors"][0]

    def test_rejects_missing_file(self, client, fake_reader):
        res = client.post("/health_hpm_pdf_preview", data={},
                          content_type="multipart/form-data")
        assert res.status_code == 400

    def test_master_error_is_json_400(self, client, env, fake_reader, monkeypatch):
        monkeypatch.setattr(Config, "HEALTH_HPM_MASTER_XLSX",
                            str(env["tmp"] / "ない.xlsx"))
        res = post_pdf(client)
        assert res.status_code == 400
        assert res.mimetype == "application/json"

    def test_jinjer_error_stops_before_calling_claude(self, client, fake_reader,
                                                      monkeypatch):
        from services.jinjer_api_client import JinjerAPIError

        def boom():
            raise JinjerAPIError("トークン取得に失敗")

        monkeypatch.setattr(app_module, "fetch_employees_for_health", boom)
        events = sse_events(post_pdf(client))

        assert events[-1][0] == "error"
        assert "jinjer" in events[-1][1]["message"]
        assert fake_reader["calls"]["n"] == 0, "課金する前に止まること"

    def test_pdf_read_error_becomes_error_event(self, client, fake_reader, monkeypatch):
        def boom(path, dpi=200):
            raise pdf_module.PdfReadError("PDFを開けません")

        monkeypatch.setattr(pdf_module, "render_pdf_pages", boom)
        events = sse_events(post_pdf(client))

        assert events[-1][0] == "error"
        assert events[-1][1]["code"] == "PDF_READ_FAILED"

    def test_page_failure_blocks_generation(self, client, fake_reader):
        fake_reader["readings"] = [pdf_module.PageReadError("読めない"), TAKAHASHI]
        done = done_of(post_pdf(client))

        assert done["can_generate"] is False
        assert done["counts"]["errors"] >= 1
        assert "PDF_PAGE_READ_FAILED" in {i["code"] for i in done["workbook_issues"]}
        assert done["counts"]["persons"] == 1, "残りのページは読める"

    def test_itaiji_name_still_matches(self, client, fake_reader):
        """原票が髙橋（はしごだか）でも jinjer の高橋に当たること。"""
        takahashi = reading("髙橋 和紀", "186")
        takahashi["identity"]["年齢"] = 47
        fake_reader["readings"] = [takahashi]
        fake_reader["pages"] = 1

        done = done_of(post_pdf(client))
        assert done["persons"][0]["jinjer"]["status"] == "ok"
        assert done["persons"][0]["jinjer"]["employee_id"] == "2019022"


# ---------------------------------------------------------------------------
# 原票画像の配信
# ---------------------------------------------------------------------------

class TestPageImage:
    def test_served_inline(self, client, fake_reader):
        done = done_of(post_pdf(client))
        res = client.get(done["persons"][0]["page_image_url"])

        assert res.status_code == 200
        assert res.mimetype == "image/png"
        assert "attachment" not in (res.headers.get("Content-Disposition") or "")
        assert res.get_data()[:8] == b"\x89PNG\r\n\x1a\n"

    @pytest.mark.parametrize("session_id", [
        "health_xx", "abc", "health_" + "z" * 32, "health_" + "a" * 31, "..%2F..%2Fx",
    ])
    def test_bad_session_id_404(self, client, fake_reader, session_id):
        assert client.get(f"/health_hpm_page_image/{session_id}/1").status_code == 404

    def test_unknown_page_404(self, client, fake_reader):
        done = done_of(post_pdf(client))
        assert client.get(
            f"/health_hpm_page_image/{done['session_id']}/9").status_code == 404

    def test_other_session_not_leaked(self, client, fake_reader):
        done = done_of(post_pdf(client))
        other = "health_" + "b" * 32
        assert client.get(f"/health_hpm_page_image/{other}/1").status_code == 404
        assert client.get(done["persons"][0]["page_image_url"]).status_code == 200


# ---------------------------------------------------------------------------
# 一時ファイルの扱い
# ---------------------------------------------------------------------------

class TestSessionFiles:
    def test_nothing_lands_in_shared_folder(self, client, env, fake_reader):
        post_pdf(client)
        assert list(env["shared"].iterdir()) == []

    def test_pngs_local_and_pdf_removed(self, client, env, fake_reader):
        done = done_of(post_pdf(client))
        sid = done["session_id"]

        assert sorted(p.name for p in env["local"].glob(f"{sid}_p*.png")) == \
            [f"{sid}_p1.png", f"{sid}_p2.png"]
        assert not (env["local"] / f"{sid}.pdf").exists(), "PDFの実体は残さない"

    def test_pkl_has_no_image_bytes(self, client, env, fake_reader):
        done = done_of(post_pdf(client))
        with open(env["local"] / f"{done['session_id']}.pkl", "rb") as f:
            session = pickle.load(f)

        assert session["source"] == "pdf"
        assert session["pages"] == {"p01": 1, "p02": 2}
        assert all(isinstance(v, int) for v in session["pages"].values())

        def has_bytes(value, depth=0):
            if depth > 6:
                return False
            if isinstance(value, (bytes, bytearray)):
                return True
            if isinstance(value, dict):
                return any(has_bytes(v, depth + 1) for v in value.values())
            if isinstance(value, (list, tuple, set)):
                return any(has_bytes(v, depth + 1) for v in value)
            return False

        assert not has_bytes(session), "画像はpklに入れない"

    def test_generate_removes_every_session_file(self, client, env, fake_reader):
        done = done_of(post_pdf(client))
        sid = done["session_id"]
        assert generate(client, sid, selections(done)).status_code == 200

        assert list(env["local"].glob(f"{sid}*")) == [], "pklもPNGも残さない"

    def test_old_files_are_swept(self, client, env, fake_reader):
        env["local"].mkdir(parents=True, exist_ok=True)
        stale_png = env["local"] / ("health_" + "c" * 32 + "_p1.png")
        keep = env["local"] / "keep_me.txt"
        stale_png.write_bytes(b"x")
        keep.write_bytes(b"x")
        old = time.time() - (Config.HEALTH_HPM_SESSION_MAX_AGE_HOURS + 1) * 3600
        os.utime(stale_png, (old, old))
        os.utime(keep, (old, old))

        post_pdf(client)
        assert not stale_png.exists(), "古い原票画像も掃除される"
        assert keep.exists(), "健診モード以外のファイルは触らない"


# ---------------------------------------------------------------------------
# CSV生成と監査用Excel
# ---------------------------------------------------------------------------

class TestGenerateFromPdf:
    def test_writes_csv_and_audit_workbook(self, client, env, fake_reader):
        from services.health_hpm_excel import parse_health_workbook

        done = done_of(post_pdf(client, name="0721～0724受診者3名分.pdf"))
        res = generate(client, done["session_id"], selections(done))
        assert res.status_code == 200, res.get_json()
        data = res.get_json()

        csv_path = data["output_path"]
        audit_path = data["audit_xlsx_path"]
        assert os.path.exists(csv_path)
        assert audit_path and os.path.exists(audit_path)
        assert os.path.dirname(audit_path) == os.path.dirname(csv_path), \
            "CSVと同じフォルダに置く"
        assert os.path.basename(audit_path).endswith("_Excel変換_整形済.xlsx")

        # 監査Excelをチェッカーで読み直せること（往復の実地確認）
        back = parse_health_workbook(audit_path)
        assert back.errors() == [], [i.message for i in back.errors()]
        assert back.genpyo_confirmed is True
        assert len(back.persons) == 2
        assert back.persons[0].blood_pressure() == {1: {"sys": "132", "dia": "86"},
                                                    2: {"sys": "118", "dia": "72"}}

    def test_csv_keeps_blood_pressure_untouched(self, client, env, fake_reader):
        import csv as csv_mod

        done = done_of(post_pdf(client))
        data = generate(client, done["session_id"], selections(done)).get_json()
        rows = list(csv_mod.reader(io.StringIO(
            open(data["output_path"], "rb").read().decode("cp932"), newline="")))

        assert rows[1][50:54] == ["132", "86", "118", "72"]
        assert "125" not in rows[1] and "79" not in rows[1], "平均を作らない"
        assert all(rows[1][c] == "" for c in range(183, 198)), "判定列は空欄"

    def test_audit_failure_keeps_csv(self, client, env, fake_reader, monkeypatch):
        def boom(*args, **kwargs):
            raise RuntimeError("書けません")

        monkeypatch.setattr(pdf_module, "write_audit_workbook", boom)
        done = done_of(post_pdf(client))
        res = generate(client, done["session_id"], selections(done))

        assert res.status_code == 200, "CSVは有効なまま"
        data = res.get_json()
        assert data["audit_xlsx_path"] is None
        assert "AUDIT_XLSX_FAILED" in {w["code"] for w in data["warnings"]}
        assert os.path.exists(data["output_path"])

    def test_still_requires_confirmation(self, client, fake_reader):
        done = done_of(post_pdf(client))
        res = generate(client, done["session_id"], selections(done), confirmed=False)
        assert res.status_code == 400
        assert "原票どおり" in res.get_json()["errors"][0]

    def test_bp_occurrence_unknown_blocks_generation(self, client, env, fake_reader):
        fake_reader["readings"] = [
            reading(bp=[{"occurrence": None, "systolic": 120, "diastolic": 61}])]
        fake_reader["pages"] = 1

        done = done_of(post_pdf(client))
        assert done["can_generate"] is False
        res = generate(client, done["session_id"], selections(done))

        assert res.status_code == 400
        assert any("BP" in e or "測定回" in e or "何回目" in e
                   for e in res.get_json()["errors"])
        assert not list(env["output"].rglob("*.csv"))


# ---------------------------------------------------------------------------
# 既存Excel経路への影響
# ---------------------------------------------------------------------------

class TestExcelRouteUnaffected:
    def test_excel_preview_still_works(self, client, env):
        from tests.health_hpm_fixtures import bp_items, item, make_v2_workbook, person

        path = make_v2_workbook(env["tmp"] / "v2.xlsx", [
            person("友納 英彦", items=[
                item("身体計測", "身長", "181.1", occurrence=1),
                *bp_items(sys1="120", dia1="61")])])
        with open(path, "rb") as f:
            res = client.post("/health_hpm_preview",
                              data={"health_excel": (io.BytesIO(f.read()), "v2.xlsx")},
                              content_type="multipart/form-data")

        assert res.status_code == 200
        data = res.get_json()
        assert data["success"] is True
        assert data["persons"][0]["page_image_url"] == "", "Excel経路では画像は付かない"
        assert data["persons"][0]["page"] is None
        assert "source" not in data, "Excel経路のJSONは従来どおり"
