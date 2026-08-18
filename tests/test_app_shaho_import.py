# -*- coding: utf-8 -*-
"""標報投入のルート。

jinjer は差し替える。ここで見張るのは**書かせないための門**が全部効いているか:

  1. 許可ユーザー以外は実行できない（許可リストが読めないときも書けない）
  2. dry-run で見た内容と違う計画は実行させない（計画ハッシュ）
  3. 対象年月の手入力照合を通らないと書かない
  4. 同時に2つ走らせない
  5. PDFの写しがセッションフォルダに残らない
"""

from __future__ import annotations

import csv
import io
import json
import os
import sys
import tempfile

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from config import Config  # noqa: E402
from services import shaho_writer  # noqa: E402
from services.shaho_pdf import PdfPerson, PdfStatement  # noqa: E402
from services.shaho_writer import CalcResult  # noqa: E402

TARGET = "2026-07"


@pytest.fixture
def client():
    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as c:
        yield c


@pytest.fixture
def session_dir(tmp_path, monkeypatch):
    path = tmp_path / "sessions"
    path.mkdir()
    monkeypatch.setattr(Config, "SHAHO_IMPORT_SESSION_DIR", str(path))
    return path


def _write_allowlist(tmp_path, monkeypatch, force_mark):
    from services.sap_import_ledger import current_user
    path = tmp_path / "許可.csv"
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["ユーザー名", "表示名", "承知投入", "備考"])
        writer.writerow([current_user(), "テスト実行者", force_mark, ""])
        writer.writerow(["谷津晴香", "谷津さん", "○", ""])
    monkeypatch.setattr(Config, "SHAHO_IMPORT_ALLOWED_USERS_CSV", str(path))
    monkeypatch.setattr(Config, "SHAHO_IMPORT_LOCK_FILE", str(tmp_path / "lock"))
    return path


@pytest.fixture
def allow_write(tmp_path, monkeypatch):
    """このテストのユーザーを「承知投入まで可」で許可リストに載せる。"""
    return _write_allowlist(tmp_path, monkeypatch, "○")


@pytest.fixture
def allow_write_only(tmp_path, monkeypatch):
    """投入はできるが、要確認の承知投入はできない実行者（平良さんら5名の想定）。"""
    return _write_allowlist(tmp_path, monkeypatch, "")


def deny_write(tmp_path, monkeypatch):
    monkeypatch.setattr(Config, "SHAHO_IMPORT_ALLOWED_USERS_CSV",
                        str(tmp_path / "ない.csv"))


def fake_pdf(name="7月保険料一覧表.pdf"):
    return {"hoken_pdf": (io.BytesIO(b"%PDF-1.4 dummy"), name)}


def statement():
    return PdfStatement(
        target_ym=TARGET, pay_ym="2026-08", office_code="263", office_name="株式会社 テスト",
        total_count=2,
        persons=[PdfPerson(emp="2099001", name="試験 太郎", kenpo_smr=300000,
                           konen_smr=300000, reason_kenpo="月額変更"),
                 PdfPerson(emp="2099002", name="試験 花子", kenpo_smr=280000,
                           konen_smr=280000, reason_kenpo="取得時決定")])


class FakeJinjer:
    """プレビューが使う GET だけを持つ代役。

    **本物のAPIは year/month を渡すとその月だけに絞る**ので、ここでも同じ挙動を
    再現する。プレビューが月で絞って取ると「過去のレコードが効いている人」を
    見失う（2026-08-17 に本番で踏んだ）。
    """

    def __init__(self, current=None):
        self.current = current or {}
        self.calls = []

    def get_monthly_remunerations(self, emps, year=None, month=None):
        self.calls.append({"year": year, "month": month})
        out = {e: list(v) for e, v in self.current.items() if e in emps}
        if year:
            out = {e: [r for r in recs
                       if r.get("year") == year
                       and (not month or r.get("month") == month)]
                   for e, recs in out.items()}
        return out

    def get_employees(self, only_active=True):
        return []


def install_preview_stubs(monkeypatch, *, current=None, calc=None, jinjer=None):
    """PDF読み取り・jinjer・従業員一覧を差し替える（APIも実PDFも使わない）。"""
    import services.shaho_pdf as pdf_module

    monkeypatch.setattr(pdf_module, "read_pdf",
                        lambda path, expected_office="": statement())
    monkeypatch.setattr(pdf_module, "verify_totals", lambda stmt: ["人数"])
    monkeypatch.setattr(pdf_module, "verify_person_premiums",
                        lambda stmt, master=None, **kw: {})
    monkeypatch.setattr(shaho_writer, "load_calc_context",
                        lambda ym, **kw: shaho_writer.CalcContext(master=object()))
    monkeypatch.setattr(shaho_writer, "expected_smr",
                        lambda p, ym, ctx: calc or CalcResult(
                            kenpo=p.kenpo_smr, konen=p.konen_smr, source="随時改定"))
    fake = jinjer if jinjer is not None else FakeJinjer(current)
    monkeypatch.setattr(app_module, "_shaho_import_client", lambda: fake)
    monkeypatch.setattr(app_module, "load_or_fetch_roster", lambda *a, **kw: [],
                        raising=False)
    import services.keiri_api as keiri_api
    monkeypatch.setattr(keiri_api, "load_or_fetch_roster", lambda *a, **kw: [])
    monkeypatch.setattr(keiri_api, "roster_index", lambda emps: {
        "2099001": {"name": "試験 太郎", "enrollment": "在籍", "retired_on": ""},
        "2099002": {"name": "試験 花子", "enrollment": "在籍", "retired_on": ""}})


def preview(client, monkeypatch, **kw):
    install_preview_stubs(monkeypatch, **kw)
    res = client.post("/shaho_import_preview", data=fake_pdf(),
                      content_type="multipart/form-data")
    return res, res.get_json()


# ---------------------------------------------------------------------------
# プレビュー
# ---------------------------------------------------------------------------

class TestPreview:
    def test_requires_a_file(self, client, session_dir):
        res = client.post("/shaho_import_preview", data={},
                          content_type="multipart/form-data")
        assert res.status_code == 400
        assert "PDF" in res.get_json()["errors"][0]

    def test_rejects_non_pdf(self, client, session_dir):
        res = client.post("/shaho_import_preview",
                          data={"hoken_pdf": (io.BytesIO(b"x"), "一覧表.xlsx")},
                          content_type="multipart/form-data")
        assert res.status_code == 400

    def test_rejects_bad_expected_ym(self, client, session_dir):
        res = client.post("/shaho_import_preview",
                          data={**fake_pdf(), "expected_ym": "2026/07"},
                          content_type="multipart/form-data")
        assert res.status_code == 400

    def test_returns_plan(self, client, session_dir, monkeypatch, allow_write):
        res, data = preview(client, monkeypatch)
        assert res.status_code == 200 and data["success"]
        assert data["target_ym"] == TARGET and data["pay_ym"] == "2026-08"
        assert {r["emp"] for r in data["rows"]} == {"2099001", "2099002"}
        assert data["plan_hash"]
        assert data["can_write"] is True

    def test_month_mismatch_is_refused(self, client, session_dir, monkeypatch):
        install_preview_stubs(monkeypatch)
        res = client.post("/shaho_import_preview",
                          data={**fake_pdf(), "expected_ym": "2026-08"},
                          content_type="multipart/form-data")
        assert res.status_code == 400
        assert "2026-07" in res.get_json()["errors"][0]

    def test_pdf_copy_is_deleted(self, client, session_dir, monkeypatch):
        preview(client, monkeypatch)
        assert [p.name for p in session_dir.iterdir() if p.suffix == ".pdf"] == []

    def test_writer_is_blocked_when_allowlist_missing(self, client, session_dir,
                                                      monkeypatch, tmp_path):
        deny_write(tmp_path, monkeypatch)
        _res, data = preview(client, monkeypatch)
        assert data["can_write"] is False
        assert "許可リスト" in data["write_reason"]

    def test_already_registered_person_is_no_change(self, client, session_dir,
                                                    monkeypatch, allow_write):
        current = {"2099001": [{"year": "2026", "month": "07",
                                "health_insurance": {"fee": "300000"},
                                "employee_pension": {"fee": "300000"}}]}
        _res, data = preview(client, monkeypatch, current=current)
        by_emp = {r["emp"]: r for r in data["rows"]}
        assert by_emp["2099001"]["status"] == "NO_CHANGE"
        assert by_emp["2099001"]["selectable"] is False

    def test_history_is_fetched_unfiltered(self, client, session_dir, monkeypatch,
                                           allow_write):
        """★回帰: 対象月で絞って取ると、過去のレコードが効いている人を見失う。

        2026-08-17 に本番exeで踏んだ（有田さん: 2026-07のレコードは無いが
        2025-09の440,000が効いていて、PDFと同額なので書く必要がなかった）。
        """
        fake = FakeJinjer({"2099001": [{"year": "2025", "month": "09",
                                        "health_insurance": {"fee": "300000"},
                                        "employee_pension": {"fee": "300000"}}]})
        _res, data = preview(client, monkeypatch, jinjer=fake)
        assert fake.calls and fake.calls[0] == {"year": None, "month": None}, \
            "報酬月額は月で絞らずに取ること"
        by_emp = {r["emp"]: r for r in data["rows"]}
        assert by_emp["2099001"]["status"] == "NO_CHANGE"
        assert by_emp["2099001"]["cur_kenpo"] == 300000
        assert by_emp["2099001"]["cur_ym"] == "2025-09"


# ---------------------------------------------------------------------------
# 実行（門のテスト）
# ---------------------------------------------------------------------------

class TestExecuteGuards:
    def _prepared(self, client, session_dir, monkeypatch, **kw):
        _res, data = preview(client, monkeypatch, **kw)
        return data

    def test_unknown_session(self, client, session_dir, allow_write):
        res = client.post("/shaho_import_execute",
                          json={"session_id": "shimp_" + "0" * 32,
                                "plan_hash": "x", "selected": []})
        assert res.status_code == 400
        assert "プレビュー" in res.get_json()["errors"][0]

    def test_plan_hash_mismatch(self, client, session_dir, monkeypatch, allow_write):
        data = self._prepared(client, session_dir, monkeypatch)
        res = client.post("/shaho_import_execute",
                          json={"session_id": data["session_id"],
                                "plan_hash": "ちがう", "selected": [{"emp": "2099001"}]})
        assert res.status_code == 400
        assert "もう一度プレビュー" in res.get_json()["errors"][0]

    def test_not_allowed_user_cannot_run(self, client, session_dir, monkeypatch,
                                         tmp_path, allow_write):
        data = self._prepared(client, session_dir, monkeypatch)
        deny_write(tmp_path, monkeypatch)
        res = client.post("/shaho_import_execute",
                          json={"session_id": data["session_id"],
                                "plan_hash": data["plan_hash"],
                                "selected": [{"emp": "2099001"}], "dry_run": True})
        assert res.status_code == 403

    def test_nothing_selected(self, client, session_dir, monkeypatch, allow_write):
        data = self._prepared(client, session_dir, monkeypatch)
        res = client.post("/shaho_import_execute",
                          json={"session_id": data["session_id"],
                                "plan_hash": data["plan_hash"], "selected": []})
        assert res.status_code == 400
        assert "選ばれていません" in res.get_json()["errors"][0]

    def test_unknown_employee_is_refused(self, client, session_dir, monkeypatch,
                                         allow_write):
        data = self._prepared(client, session_dir, monkeypatch)
        res = client.post("/shaho_import_execute",
                          json={"session_id": data["session_id"],
                                "plan_hash": data["plan_hash"],
                                "selected": [{"emp": "9999999"}]})
        assert res.status_code == 400

    def test_review_row_needs_explicit_force(self, client, session_dir, monkeypatch,
                                             allow_write):
        """計算と食い違う人は「承知のうえ投入」なしでは通らない。"""
        data = self._prepared(client, session_dir, monkeypatch,
                              calc=CalcResult(kenpo=999000, konen=999000,
                                              source="随時改定"))
        by_emp = {r["emp"]: r for r in data["rows"]}
        assert by_emp["2099001"]["status"] == "CALC_MISMATCH"
        res = client.post("/shaho_import_execute",
                          json={"session_id": data["session_id"],
                                "plan_hash": data["plan_hash"],
                                "selected": [{"emp": "2099001"}], "dry_run": True})
        assert res.status_code == 400
        assert "承知のうえ投入" in res.get_json()["errors"][0]

        ok = client.post("/shaho_import_execute",
                         json={"session_id": data["session_id"],
                               "plan_hash": data["plan_hash"],
                               "selected": [{"emp": "2099001", "forced": True}],
                               "dry_run": True})
        assert ok.status_code == 200

    def test_import_only_user_can_run_matching_rows(self, client, session_dir,
                                                    monkeypatch, allow_write_only):
        """3点が一致した人の投入は、承知投入の権限が無くてもできる。"""
        data = self._prepared(client, session_dir, monkeypatch)
        assert data["can_write"] is True and data["can_force"] is False
        res = client.post("/shaho_import_execute",
                          json={"session_id": data["session_id"],
                                "plan_hash": data["plan_hash"],
                                "selected": [{"emp": "2099001"}], "dry_run": True})
        assert res.status_code == 200

    def test_import_only_user_cannot_force_a_review_row(self, client, session_dir,
                                                        monkeypatch, allow_write_only):
        """要確認の人は、画面で承知投入を選んでも権限が無ければ通らない。"""
        data = self._prepared(client, session_dir, monkeypatch,
                              calc=CalcResult(kenpo=999000, konen=999000,
                                              source="随時改定"))
        assert data["rows"][0]["status"] == "CALC_MISMATCH"
        res = client.post("/shaho_import_execute",
                          json={"session_id": data["session_id"],
                                "plan_hash": data["plan_hash"],
                                "selected": [{"emp": "2099001", "forced": True}],
                                "dry_run": True})
        assert res.status_code == 400
        assert "承知のうえ投入" in res.get_json()["errors"][0]

    def test_dry_run_writes_nothing(self, client, session_dir, monkeypatch, allow_write):
        data = self._prepared(client, session_dir, monkeypatch)
        called = []
        monkeypatch.setattr(app_module, "_shaho_import_write_client",
                            lambda: called.append(1))
        res = client.post("/shaho_import_execute",
                          json={"session_id": data["session_id"],
                                "plan_hash": data["plan_hash"],
                                "selected": [{"emp": "2099001"}], "dry_run": True})
        body = res.get_json()
        assert res.status_code == 200 and body["dry_run"] is True
        assert body["count"] == 1 and body["results"][0]["result"] == "dry-run"
        assert called == []

    def test_real_run_needs_matching_confirm_ym(self, client, session_dir, monkeypatch,
                                                allow_write):
        data = self._prepared(client, session_dir, monkeypatch)
        for bad in ("", "2026-08", "ちがう"):
            res = client.post("/shaho_import_execute",
                              json={"session_id": data["session_id"],
                                    "plan_hash": data["plan_hash"],
                                    "selected": [{"emp": "2099001"}],
                                    "dry_run": False, "confirm_ym": bad})
            assert res.status_code == 400
            assert TARGET in res.get_json()["errors"][0]

    def test_lock_blocks_a_second_run(self, client, session_dir, monkeypatch,
                                      allow_write, tmp_path):
        data = self._prepared(client, session_dir, monkeypatch)
        shaho_writer.acquire_lock(TARGET, 1, path=str(tmp_path / "lock"))
        res = client.post("/shaho_import_execute",
                          json={"session_id": data["session_id"],
                                "plan_hash": data["plan_hash"],
                                "selected": [{"emp": "2099001"}],
                                "dry_run": False, "confirm_ym": TARGET})
        assert res.status_code == 409
        assert "別の投入" in res.get_json()["errors"][0]
        shaho_writer.release_lock(str(tmp_path / "lock"))


# ---------------------------------------------------------------------------
# 進捗
# ---------------------------------------------------------------------------

class TestStatus:
    def test_rejects_bad_session_id(self, client, session_dir):
        res = client.get("/shaho_import_status?session_id=../../etc/passwd")
        assert res.status_code == 400

    def test_no_progress_yet(self, client, session_dir):
        res = client.get("/shaho_import_status?session_id=shimp_" + "a" * 32)
        assert res.get_json()["state"] == "none"

    def test_returns_written_progress(self, client, session_dir):
        session_id = "shimp_" + "b" * 32
        app_module._write_shaho_progress(session_id, {"state": "running", "done": 3,
                                                      "total": 10})
        res = client.get(f"/shaho_import_status?session_id={session_id}")
        body = res.get_json()
        assert body["state"] == "running" and body["done"] == 3


class TestDownload:
    def test_rejects_bad_month(self, client):
        res = client.get("/shaho_import_download/20xx07/x.json")
        assert res.status_code == 400
