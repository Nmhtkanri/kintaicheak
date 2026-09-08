# -*- coding: utf-8 -*-
"""標準報酬チェックの1人分判定（judge_person）— 随時改定の候補と休職の確認。

2026-09-08 のレポートで分かったこと:
  - 休職者2名が「随時改定の候補」に出た（休職給は固定的賃金の変動ではなく、3か月とも
    支払基礎日数17日以上という要件も満たさない）
  - 4月入社で5月から報酬が大きく変わった人（取得時決定の見込み額と実績のずれ）が候補に出なかった
  - 候補者の計算欄が定時決定の式の値のままで、随時改定の通知額と違って見えた
ここで見張るのは、その3つが二度と起きないこと。等級表と分類マスタは tests/test_shaho.py の
合成物を流用する（健保 1:58,000 / 2:68,000（63,000〜） / 3:78,000（73,000〜））。
"""

from __future__ import annotations

import json
import os
import sys
import tempfile

import openpyxl
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from services import shaho_check  # noqa: E402
from services.keiri_engine import ym_add  # noqa: E402
from services.shaho_check import judge_person  # noqa: E402
from services.shaho_master import load_class_master  # noqa: E402
from services.shaho_report import EMPLOYEE_COLUMNS, write_reports  # noqa: E402
from tests.test_shaho import _ClassMasterMixin, _pi_shaho, build_workbook, load_from_workbook  # noqa: E402

EMP = "2020001"


@pytest.fixture(scope="module")
def cm():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "class.csv")
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(_ClassMasterMixin.CSV)
        yield load_class_master(path)


@pytest.fixture(scope="module")
def master():
    return load_from_workbook(build_workbook())


def _rec(base, system="月給制1", kenpo=58000, konen=88000, joined="2025-04-01",
         kintai=None, labels=None, extra=None):
    """1人月のキャッシュレコード。報酬計＝雇用保険対象額（other5）にして検算ゲートを通す。"""
    extra = dict(extra or {})
    bi = {"last_name": "山田", "first_name": "太郎", "joined_on": joined,
          "salary_system": {"id": "1", "name": system},
          "health_insurance": kenpo, "employee_pension": konen,
          "is_social_insurance_calculated": True,
          "social_insurance_classification": {"id": "1", "name": "加入"},
          "health_insurance_calculation_classification": {"id": "0", "name": "被保険者（対象）"},
          "care_insurance_calculation_classification": {"id": "0", "name": "対象外"},
          "employees_pension": {"calculation_classification": {"id": "0", "name": "被保険者（対象）"}}}
    shikyu = {"allowance1": base, **extra}
    pi = _pi_shaho(shikyu=shikyu, kintai=kintai, other={"other5": base + sum(extra.values())},
                   labels=labels)
    pi["is_payroll_closed"] = True
    return {"basic_info": bi, "payroll_info": pi, "n_nonzero": 1}


def _months(spec: dict) -> dict:
    """{"2026-04": _rec(...), ...}"""
    return dict(spec)


def _cfg(check_month="2026-08"):
    return {"threshold": 17, "rounding": "50sen", "tolerance": 1, "lag": 1,
            "check_month": check_month, "open_months": {},
            "revision_window": {"2026-04", "2026-05", "2026-06"}}


def _judge(months, master, cm, check_month="2026-08"):
    pair = (months.get(ym_add(check_month, -1)), months.get(check_month))
    return judge_person(EMP, months, pair, 2026, master, cm, _cfg(check_month))


def _flat(base_by_month: dict, **kw) -> dict:
    return {ym: _rec(base, **kw) for ym, base in base_by_month.items()}


# ---------------------------------------------------------------- 随時改定の候補
def test_fixed_wage_change_with_full_base_days_is_candidate(master, cm):
    months = _flat({"2026-04": 60000, "2026-05": 80000, "2026-06": 80000,
                    "2026-07": 80000, "2026-08": 80000})
    r = _judge(months, master, cm)
    assert r.teiji_status == "MONTHLY_REVISION_CANDIDATE"
    rv = r.revision
    assert (rv.change_month, rv.apply_month, rv.pending) == ("2026-05", "2026-08", False)
    assert "固定的賃金の変動" in rv.trigger
    assert rv.kettei.average == 80000 and rv.kettei.kenpo_smr == 78000
    assert rv.grade_diff == 2
    assert r.calc_basis == "随時改定"
    assert any("2026-08改定" in n and "78,000" in n for n in r.notes)


def test_short_base_days_in_window_blocks_candidate(master, cm):
    """随時改定は3か月とも支払基礎日数17日以上が要件。7月が16日なら候補にしない。"""
    months = _flat({"2026-04": 60000, "2026-05": 80000, "2026-06": 80000, "2026-08": 80000})
    months["2026-07"] = _rec(80000, kintai={"kintai11": 15}, labels={"kintai11": "欠勤日数"})
    r = _judge(months, master, cm)
    assert r.revision is None
    assert r.teiji_status != "MONTHLY_REVISION_CANDIDATE"
    assert any("2026-07は支払基礎日数 16日 < 17日" in n and "要件" in n for n in r.notes)


def test_leave_months_go_to_leave_review_not_candidate(master, cm):
    """休職（無給）: 候補にせず「休職の確認」。定時決定の計算値は出さない。"""
    months = _flat({"2026-04": 60000, "2026-05": 80000, "2026-06": 0,
                    "2026-07": 0, "2026-08": 0})
    r = _judge(months, master, cm)
    assert r.teiji_status == "LEAVE_REVIEW"
    assert r.total_status == "LEAVE_REVIEW" and r.total_status in shaho_check.REVIEW_STATUSES
    assert r.leave_months == ["2026-06", "2026-07", "2026-08"]
    assert r.revision is None
    assert "保険者算定" in r.calc_basis
    assert any("休職の疑い" in n and "2026-06" in n for n in r.notes)


def test_salary_system_switch_triggers_candidate(master, cm):
    """基本給は同額のまま体系が変わり、変動扱いの差額調整で報酬が増えた（2026004型）。"""
    months = {"2026-04": _rec(60000, system="時給制1",
                              kintai={"kintai10": 20}, labels={"kintai10": "出勤日数"})}
    for ym in ("2026-05", "2026-06", "2026-07", "2026-08"):
        months[ym] = _rec(60000, extra={"allowance24": 20000})
    r = _judge(months, master, cm)
    assert r.teiji_status == "MONTHLY_REVISION_CANDIDATE"
    assert "給与体系の切替（時給制1→月給制1）" in r.revision.trigger
    assert (r.revision.change_month, r.revision.apply_month) == ("2026-05", "2026-08")
    assert r.revision.kettei.kenpo_smr == 78000


def test_switch_from_provisional_system_is_not_a_trigger(master, cm):
    """「〜暫定」からの切替は 2026-04 の全社的な体系移行なので、それだけでは候補にしない。"""
    months = {"2026-04": _rec(60000, system="時給制暫定",
                              kintai={"kintai10": 20}, labels={"kintai10": "出勤日数"})}
    for ym in ("2026-05", "2026-06", "2026-07", "2026-08"):
        months[ym] = _rec(60000, extra={"allowance24": 20000})
    r = _judge(months, master, cm)
    assert r.revision is None
    assert r.teiji_status == "DIFFERENCE"


def test_joined_previous_month_is_a_trigger(master, cm):
    """4月入社（取得時決定は見込み額）→ 5月からの3か月で2等級以上ずれれば8月改定の候補。"""
    months = {"2026-04": _rec(60000, joined="2026-04-10", kenpo=0, konen=0)}
    for ym in ("2026-05", "2026-06", "2026-07", "2026-08"):
        months[ym] = _rec(60000, joined="2026-04-10", extra={"allowance24": 20000})
    r = _judge(months, master, cm)
    assert r.teiji_status == "MONTHLY_REVISION_CANDIDATE"
    assert "資格取得（04/10入社）の翌月" in r.revision.trigger
    assert (r.revision.change_month, r.revision.apply_month) == ("2026-05", "2026-08")


def test_registered_flip_is_fallback_when_joined_on_is_old(master, cm):
    """joined_on が古い人が途中で被保険者になった: 登録標報 0→正 の月の翌月を変動月にする。"""
    months = {"2026-04": _rec(60000, joined="2020-04-01", kenpo=0, konen=0)}
    for ym in ("2026-05", "2026-06", "2026-07", "2026-08"):
        months[ym] = _rec(60000, joined="2020-04-01", extra={"allowance24": 20000})
    r = _judge(months, master, cm)
    assert r.teiji_status == "MONTHLY_REVISION_CANDIDATE"
    assert "0→58,000円" in r.revision.trigger
    assert (r.revision.change_month, r.revision.apply_month) == ("2026-06", "2026-09")


def test_change_outside_april_to_june_is_ignored(master, cm):
    months = _flat({"2026-04": 60000, "2026-05": 60000, "2026-06": 60000,
                    "2026-07": 80000, "2026-08": 80000})
    r = _judge(months, master, cm)
    assert r.revision is None
    assert r.teiji_status == "OK" and r.calc_basis == "定時決定"


def test_one_grade_difference_is_not_candidate_and_silent(master, cm):
    months = _flat({"2026-04": 60000, "2026-05": 68000, "2026-06": 68000,
                    "2026-07": 68000, "2026-08": 68000})
    r = _judge(months, master, cm)
    assert r.revision is None
    assert r.teiji_status == "DIFFERENCE"
    assert not any("随時改定" in n for n in r.notes)


def test_incomplete_window_is_pending_candidate(master, cm):
    months = _flat({"2026-04": 60000, "2026-05": 60000, "2026-06": 80000, "2026-07": 80000})
    r = _judge(months, master, cm, check_month="2026-07")
    assert r.teiji_status == "MONTHLY_REVISION_CANDIDATE"
    rv = r.revision
    assert rv.pending and rv.kettei.adopted_n == 2
    assert (rv.change_month, rv.apply_month) == ("2026-06", "2026-09")
    assert r.calc_basis == "随時改定（様子見）"
    assert "未完" in rv.reason


def test_revision_to_dict_shape(master, cm):
    months = _flat({"2026-04": 60000, "2026-05": 80000, "2026-06": 80000,
                    "2026-07": 80000, "2026-08": 80000})
    d = _judge(months, master, cm).revision.to_dict()
    assert {"change_month", "apply_month", "trigger", "window", "pending", "grade_diff",
            "reason", "average", "kenpo_grade", "kenpo_smr", "konen_smr", "months"} <= set(d)
    assert [m["ym"] for m in d["months"]] == ["2026-05", "2026-06", "2026-07"]
    assert d["months"][0]["days"] == 31 and d["kenpo_smr"] == 78000


# ---------------------------------------------------------------- Excel / JSON
def _check(results, master, cm, out_base):
    return {"year": 2026, "check_month": "2026-08", "insurer": "its", "master": master,
            "class_master": {"path": "class.csv", "rules": []}, "cfg": _cfg(),
            "results": results, "revisions": [], "out_base": out_base, "month_status": {}}


def _rows(ws):
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    return hdr, [dict(zip(hdr, r)) for r in it]


def test_reports_show_calc_basis_and_revision_columns(master, cm, tmp_path):
    cand = _judge(_flat({"2026-04": 60000, "2026-05": 80000, "2026-06": 80000,
                         "2026-07": 80000, "2026-08": 80000}), master, cm)
    leave = _judge(_flat({"2026-04": 60000, "2026-05": 80000, "2026-06": 0,
                          "2026-07": 0, "2026-08": 0}), master, cm)
    plain = _judge(_flat({"2026-04": 60000, "2026-05": 60000, "2026-06": 60000,
                          "2026-07": 80000, "2026-08": 80000}), master, cm)
    cand.emp, leave.emp, plain.emp = "2020001", "2020002", "2020003"
    out = write_reports(_check([cand, leave, plain], master, cm, str(tmp_path)))
    wb = openpyxl.load_workbook(out["xlsx"], read_only=True)

    hdr, rows = _rows(wb["社員別判定"])
    assert hdr == EMPLOYEE_COLUMNS
    by = {r["社員番号"]: r for r in rows}
    assert by["2020001"]["計算根拠"] == "随時改定"
    assert by["2020001"]["計算健保標報"] == 78000            # 定時決定（73,333→78,000）ではなく窓の値
    assert by["2020001"]["平均報酬月額"] == 80000
    assert (by["2020001"]["随時改定 変動月"], by["2020001"]["随時改定 改定月"]) == ("2026-05", "2026-08")
    assert by["2020001"]["随時改定 基礎日数"] == "31/30/31"
    assert by["2020002"]["計算健保標報"] is None and "保険者算定" in by["2020002"]["計算根拠"]
    assert by["2020002"]["標報判定"] == "休職の確認"
    assert by["2020003"]["計算根拠"] == "定時決定" and by["2020003"]["計算健保標報"] == 58000

    _, detail = _rows(wb["算定明細"])
    rev_rows = [r for r in detail if r["社員番号"] == "2020001" and str(r["用途"]).startswith("随時改定")]
    assert [r["支給月"] for r in rev_rows] == ["2026-05", "2026-06", "2026-07"]
    assert sum(1 for r in detail if r["用途"] == "定時決定") == 9

    hdr, review = _rows(wb["要確認"])
    assert "計算根拠" in hdr
    assert {r["社員番号"] for r in review} >= {"2020001", "2020002"}

    _, summary = _rows(wb["サマリー"])
    values = {r["項目"]: r["値"] for r in summary}
    assert values["随時改定の候補（7〜9月改定の見込み・定時決定の対象外）"] == "1名"
    assert values["休職の確認（無給の月あり・保険者算定の可能性）"] == "1名"

    payload = json.load(open(out["json"], encoding="utf-8"))
    emp = {e["emp"]: e for e in payload["employees"]}
    assert emp["2020001"]["revision"]["apply_month"] == "2026-08"
    assert emp["2020001"]["calc_kenpo_smr"] == 78000 and emp["2020001"]["calc_basis"] == "随時改定"
    assert emp["2020002"]["leave_months"] == ["2026-06", "2026-07", "2026-08"]
    assert emp["2020002"]["calc_kenpo_smr"] is None
    assert [c["emp"] for c in payload["revision_candidates"]] == ["2020001"]
    assert payload["revision_candidates"][0]["calc_kenpo"] == 78000
    assert [l["emp"] for l in payload["leave_review"]] == ["2020002"]
    assert out["revision_candidates"] == payload["revision_candidates"]
    assert out["leave_review"] == payload["leave_review"]


# ---------------------------------------------------------------- 画面（/shaho_run）
@pytest.fixture
def client():
    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as c:
        yield c


def test_route_returns_candidate_and_leave_lists(client, monkeypatch, tmp_path):
    monkeypatch.setattr(shaho_check.Config, "KEIRI_OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr("services.shaho_check.run_check", lambda *a, **k: {"results": []})
    report = {"n": 3, "review_n": 2, "open_months": {}, "xlsx": "x.xlsx", "json": "x.json",
              "revision_candidates": [{"emp": "2020001", "name": "山田 太郎", "pending": False}],
              "leave_review": [{"emp": "2020002", "name": "山田 花子", "months": ["2026-06"]}]}
    monkeypatch.setattr("services.shaho_report.write_reports", lambda check: dict(report))
    r = client.post("/shaho_run", data={"year": "2026", "check_month": "2026-08"})
    assert r.status_code == 200, r.get_json()
    data = r.get_json()
    assert data["revision_candidates"] == report["revision_candidates"]
    assert data["leave_review"] == report["leave_review"]
    # 旧い write_reports（キー無し）でも落ちない
    monkeypatch.setattr("services.shaho_report.write_reports",
                        lambda check: {"n": 0, "review_n": 0, "open_months": {},
                                       "xlsx": "x.xlsx", "json": "x.json"})
    r = client.post("/shaho_run", data={"year": "2026", "check_month": "2026-08"})
    assert r.status_code == 200 and r.get_json()["revision_candidates"] == []


def test_index_wires_the_two_tables(client):
    html = client.get("/").get_data(as_text=True)
    assert 'id="shaho-revisions"' in html and 'id="shaho-leave"' in html
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    js = open(os.path.join(root, "static", "script.js"), encoding="utf-8").read()
    assert "revision_candidates" in js and "leave_review" in js
