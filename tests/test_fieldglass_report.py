# -*- coding: utf-8 -*-
"""FG新レポートパーサ（ユニアデックスXLSX・エリクソンCSV）の合成データテスト。

実サンプル（Downloads配下・個人情報）はリポジトリに入れないため、列名だけ実物どおりの
ミニデータをテスト内で組み立てる。実データでの突合は tools/verify（手動）と
2026Q2 の新旧build比較（docs/PLAN_派遣台帳モード.md に記録）で行った。
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services.daicho.fieldglass import contract_from_workorder, workorders_in_quarter  # noqa: E402
from services.daicho.fieldglass_report import (apply_ericsson_report, load_ericsson_report,  # noqa: E402
                                               load_ual_report, _name_tokens)

# ---------------------------------------------------------------------------
# ユニアデックス XLSX
# ---------------------------------------------------------------------------

_UAL_COLS = [
    "求人情報 ID", "求人情報タイトル", "事業単位", "コストセンター", "応募者の ID", "応募者/スタッフ",
    "改訂番号", "勤務地", "作業オーダー ID", "作業オーダー開始日", "最新の作業オーダー終了日",
    "派遣先責任者の氏名", "派遣先責任者の役職", "派遣先責任者の部署", "派遣先責任者の電話番号",
    "従事する業務に伴う責任の程度", "就業日", "組織の長の職名", "組織単位の名称",
    "Break Start Time", "Break End Time",
    "Complaints Handling Representative Info (Name, Title, Dept, Tel, Email)",
    "Matters about education and training to provide necessary capabilities for the work",
    "Shift Pattern", "Site tenure", "Site address",
    "Supervisor Info (Name, Title, Dept, Tel, Email)",
    "Typical Working days remarks", "Workers subject to Labor Agreement or not",
    "Work Location UAL", "Work Hours End Time", "Work Hours Start Time",
    "Typical Working days(new)", "Shift confirmation timing",
    "Monthly ST : Lower Limit", "Monthly ST : Upper Limit",
    "Description of Work (Please be as descriptive as possible)",
    "Division Name UAL",
    "Client Side Responsible Person Info (Name, Title, Dept, Tel, Email)",
    "Level of Responsibility",
]


def _ual_row(**vals) -> list:
    return [vals.get(c, "") for c in _UAL_COLS]


@pytest.fixture
def ual_xlsx(tmp_path):
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["合成レポート（テスト）"])
    ws.append(_UAL_COLS)
    common = dict({
        "求人情報 ID": "JP1", "求人情報タイトル": "T", "事業単位": "NW本部。二部．一課",
        "応募者の ID": "S1", "応募者/スタッフ": "山田, 太郎", "勤務地": "本社",
        "作業オーダー ID": "WO1",
        "Client Side Responsible Person Info (Name, Title, Dept, Tel, Email)":
            "松尾　和善;部長;テクサポ部;050-1;m@example.com",
        "Complaints Handling Representative Info (Name, Title, Dept, Tel, Email)":
            "阿野　弘喜;本部長;NW本部;050-2;a@example.com",
        "Supervisor Info (Name, Title, Dept, Tel, Email)": "大須賀　怜;室長;テク２;050-3;o@example.com",
        "Division Name UAL": "NW本部　テクサポ二部;部長",
        "Site tenure": dt.datetime(2027, 10, 1),
        "Site address": "東京都江東区1-1",
        "Work Location UAL": "UAL東京SC;東京都江東区東雲1-7-12;「食堂」無;「更衣室」有",
        "Workers subject to Labor Agreement or not": "労使協定方式に限る",
        "Matters about education and training to provide necessary capabilities for the work": "OJT研修",
        "Monthly ST : Lower Limit": "140", "Monthly ST : Upper Limit": "180",
        "Description of Work (Please be as descriptive as possible)": "NW機器の検証",
        "Level of Responsibility": "一般業務 役職なし",
        "Typical Working days(new)": "月曜 勤務-火曜 勤務-水曜 勤務-木曜 勤務-金曜 勤務-土曜 非勤務-日曜 非勤務",
        "Typical Working days remarks": "休日は土日祝",
        "Work Hours Start Time": "09:00", "Work Hours End Time": "17:30",
        "Break Start Time": "12:00", "Break End Time": "13:00",
    })
    # 改訂0（コストセンター違いで2行に重複）と改訂1。終了日はどちらも「最新の」9/30
    ws.append(_ual_row(**common, **{"改訂番号": "0.0", "コストセンター": "CC-A",
                                    "作業オーダー開始日": dt.datetime(2026, 4, 1),
                                    "最新の作業オーダー終了日": dt.datetime(2026, 9, 30)}))
    ws.append(_ual_row(**common, **{"改訂番号": "0.0", "コストセンター": "CC-B",
                                    "作業オーダー開始日": dt.datetime(2026, 4, 1),
                                    "最新の作業オーダー終了日": dt.datetime(2026, 9, 30)}))
    ws.append(_ual_row(**common, **{"改訂番号": "1.0", "コストセンター": "CC-A",
                                    "作業オーダー開始日": dt.datetime(2026, 7, 1),
                                    "最新の作業オーダー終了日": dt.datetime(2026, 9, 30)}))
    # シフト系の人（Work Hours 空・Shift Pattern に入る）
    ws.append(_ual_row(**{**common,
                          "応募者の ID": "S2", "応募者/スタッフ": "佐藤, 花子", "作業オーダー ID": "WO2",
                          "改訂番号": "0.0", "コストセンター": "CC-C",
                          "作業オーダー開始日": dt.datetime(2026, 4, 1),
                          "最新の作業オーダー終了日": dt.datetime(2026, 6, 30),
                          "Work Hours Start Time": "", "Work Hours End Time": "",
                          "Break Start Time": "", "Break End Time": "",
                          "Shift Pattern": "4勤4休", "Shift confirmation timing": "前月20日",
                          "Typical Working days remarks": "シフト表による"}))
    # DTS行（英語列が空・日本語列に入る）。終了日が空＝継続中扱い
    ws.append(_ual_row(**{"求人情報 ID": "JP3", "事業単位": "DTS事業部",
                          "応募者の ID": "S3", "応募者/スタッフ": "岡崎, 修司", "作業オーダー ID": "WO3",
                          "改訂番号": "0.0", "作業オーダー開始日": dt.datetime(2026, 5, 1),
                          "派遣先責任者の氏名": "田中 一", "派遣先責任者の役職": "課長",
                          "派遣先責任者の部署": "DTS部", "派遣先責任者の電話番号": "03-1",
                          "従事する業務に伴う責任の程度": "役職なし（DTS）",
                          "就業日": "月曜 勤務-火曜 勤務-水曜 勤務-木曜 勤務-金曜 勤務-土曜 非勤務-日曜 非勤務",
                          "組織単位の名称": "DTS一課", "組織の長の職名": "課長"}))
    path = tmp_path / "ユニアデックス_WorkOrder_基本_業務内容_20260101.xlsx"
    wb.save(path)
    return path


def test_ual_merges_cost_centers_and_truncates_revision_ends(ual_xlsx):
    wos, by_staff, by_wo = load_ual_report(ual_xlsx)
    wo1 = {w.revision: w for w in wos if w.wo_id == "WO1"}
    assert wo1[0].cost_centers == ["CC-A", "CC-B"]          # 重複行はコストセンターに畳む
    assert wo1[0].end == dt.date(2026, 6, 30)                # 「最新の終了日」を次改訂開始-1日へ切り詰め
    assert wo1[1].end == dt.date(2026, 9, 30)                # 最終改訂はそのまま
    q2 = workorders_in_quarter(wos, dt.date(2026, 4, 1), dt.date(2026, 6, 30))
    assert {(w.wo_id, w.revision) for w in q2} == {("WO1", 0), ("WO2", 0), ("WO3", 0)}


def test_ual_entry_normalization(ual_xlsx):
    _wos, by_staff, _by_wo = load_ual_report(ual_xlsx)
    e = by_staff["S1"]
    assert e["clientResponsible"] == {"name": "松尾　和善", "title": "部長", "department": "テクサポ部",
                                      "phone": "050-1", "email": "m@example.com"}
    assert e["supervisor"]["name"] == "大須賀　怜"
    assert e["orgUnit"] == "NW本部　テクサポ二部" and e["orgChief"] == "部長"
    assert e["siteTenure"] == "2027/10/1"
    assert e["workOffice"] == "UAL東京SC"
    assert e["workAddress"] == "東京都江東区東雲1-7-12"
    assert e["conveniences"] == "「食堂」無、「更衣室」有"
    assert e["agreementTarget"] == "該当する"
    assert e["training"] == "OJT研修"
    assert (e["monthlyStdLower"], e["monthlyStdUpper"]) == ("140", "180")
    assert e["workStart"] == "9:00" and e["breakEnd"] == "13:00"
    assert e["workdaysHolidays"].startswith("月曜 勤務")


def test_ual_shift_row_has_notes_but_no_hours(ual_xlsx):
    _wos, by_staff, _by_wo = load_ual_report(ual_xlsx)
    e = by_staff["S2"]
    assert "workStart" not in e                              # シフト系は時間が入らない
    assert "シフト: 4勤4休" in e["workdaysHolidaysNotes"]
    assert "シフト確定時期: 前月20日" in e["workdaysHolidaysNotes"]


def test_ual_dts_row_falls_back_to_japanese_columns(ual_xlsx):
    wos, by_staff, _by_wo = load_ual_report(ual_xlsx)
    e = by_staff["S3"]
    assert e["clientResponsible"]["name"] == "田中 一"
    assert e["responsibilityDegree"] == "役職なし（DTS）"
    assert e["orgUnit"] == "DTS一課" and e["orgChief"] == "課長"
    wo3 = next(w for w in wos if w.wo_id == "WO3")
    assert wo3.end is None                                   # 終了日空＝継続中扱いのまま


def test_contract_from_workorder_applies_report_only_keys(ual_xlsx):
    wos, by_staff, _by_wo = load_ual_report(ual_xlsx)
    wo = next(w for w in wos if w.wo_id == "WO1" and w.revision == 0)
    c = contract_from_workorder(wo, None, dt.date(2026, 6, 30), None, by_staff["S1"])
    cmd = "事業所の名称及び所在地その他派遣就業場所"
    assert c.cpi[f"{cmd} 指揮命令者氏名"] == "大須賀　怜"
    assert c.cpi[f"{cmd} 指揮命令者役職"] == "室長"
    assert c.tc["組織単位"] == "NW本部　テクサポ二部"
    assert c.tc["組織の長の職名"] == "部長"
    assert c.tc["事業所抵触日"] == "2027/10/1"
    assert c.tc["便宜供与：その他1"] == "「食堂」無、「更衣室」有"
    assert c.cpi["協定対象派遣労働者に該当するか否かの別"] == "該当する"
    assert c.cpi["教育訓練"] == "OJT研修"
    assert c.cpi[f"{cmd} 事業所の名称"] == "UAL東京SC"


def test_contract_from_workorder_without_new_keys_is_unchanged(ual_xlsx):
    """旧details JSON相当（新キーなし）では従来動作＝legacyモードの出力が変わらない。"""
    wos, _s, _w = load_ual_report(ual_xlsx)
    wo = next(w for w in wos if w.wo_id == "WO1" and w.revision == 0)
    legacy_detail = {"clientResponsible": {"name": "旧 責任者", "title": "", "department": "", "phone": ""}}
    c = contract_from_workorder(wo, None, dt.date(2026, 6, 30), None, legacy_detail)
    cmd = "事業所の名称及び所在地その他派遣就業場所"
    assert c.cpi[f"{cmd} 指揮命令者氏名"] == "大須賀　怜"   # WOのスーパーバイザ結合（従来経路）
    assert "組織の長の職名" not in c.tc
    assert "教育訓練" not in c.cpi


# ---------------------------------------------------------------------------
# エリクソン CSV
# ---------------------------------------------------------------------------

_ERIC_CSV = "﻿" + "\r\n".join([
    ",".join(["Job Posting ID", "Job Seeker",
              "Client Side Responsible Person", "Client Side Responsible Person",
              "Client Side Responsible Person Position", "Client Side Responsible Person Department",
              "Client Side Responsible Person Telephone Number", "Client Side Responsible Person Telephone Number",
              "Complaints Handling Representative Position", "Complaints Handling Representative Telephone Number",
              "Supervisor level 1",
              '"Notes on breaks, holidays or work hours"', '"Notes on breaks, holidays or work hours"',
              "Work Hours Start Time", "Work Hours End Time", "Break Start Time", "Break End Time",
              "Typical Working days", "Typical Non-working days", "Level of Responsibility",
              "Is the workers age greater than or equal to 60 years old?", "Worker Location (at client site)"]),
    ",".join(["JP1", '"Ohta, Takuya"', "Li Stephen", "Li Stephen", "Manager", "CSS JP", "070-1", "070-1",
              "Manager", "070-1", "Nobuhito Miki(nobuhito@example.com)",
              "就業時間は変形労働時間制（4勤4休）", "特に無し",
              "8:45", "21:00", "12:00", "13:15", "月 火 水 木 金 土 日", "シフト表による", "役職無し",
              "No", "Yokohama"]),
    ",".join(["JP2", '"Nara, Takahiro"', "Li Stephen", "Li Stephen", "Manager", "CSS JP", "070-1", "070-1",
              "Manager", "070-1", "Stephen Li(stephen@example.com)",
              "", "特に無し",
              "9:00", "18:00", "12:00", "13:00", "月 火 水 木 金", "土日祝", "役職無し 専念",
              "No", "Yokohama"]),
]) + "\r\n"


def test_name_tokens_collapse_long_vowels():
    assert _name_tokens("Ohta, Takuya") == frozenset({"ota", "takuya"})
    assert _name_tokens("MAHARJAN RAMITA") == frozenset({"maharjan", "ramita"})


def test_ericsson_dup_headers_and_notes(tmp_path):
    path = tmp_path / "派遣元管理台帳作成用_20260101.csv"
    path.write_text(_ERIC_CSV, encoding="utf-8")
    rows = load_ericsson_report(path)
    assert [r["name"] for r in rows] == ["Ohta, Takuya", "Nara, Takahiro"]
    ohta = rows[0]
    assert ohta["client_responsible"]["name"] == "Li Stephen"       # 重複ヘッダーは最初の1個
    assert ohta["supervisor_name"] == "Nobuhito Miki"               # (email) を落とす
    assert ohta["notes"] == ["就業時間は変形労働時間制（4勤4休）"]   # Notesは両列から・「特に無し」は除外


def test_ericsson_overlay_updates_fixed_and_guards_shift(tmp_path):
    path = tmp_path / "派遣元管理台帳作成用_20260101.csv"
    path.write_text(_ERIC_CSV, encoding="utf-8")
    rows = load_ericsson_report(path)
    master = [
        {"派遣先名称": "エリクソン・ジャパン株式会社", "氏名": "奈良 隆宏", "氏名カナ": "Nara, Takahiro",
         "就業時間": "9:00～17:00", "休憩時間": "12:00～12:45", "就業曜日": "月 火 水", "休日": "土日",
         "派遣先責任者_氏名": "旧責任者", "苦情申出先_氏名": "三木 暢人", "苦情申出先_役職": "旧役職",
         "指揮命令者_氏名": "鈴木 紀生", "責任の程度": "役職なし", "備考": ""},
        {"派遣先名称": "エリクソン・ジャパン株式会社", "氏名": "太田 琢也", "氏名カナ": "Ohta, Takuya",
         "就業時間": "シフト制（原則 日勤 08:45～21:00）", "休憩時間": "1時間15分",
         "就業曜日": "月 火 水 木 金 土 日 祝 シフトあり", "休日": "別途シフト表による",
         "派遣先責任者_氏名": "服部 幸治", "苦情申出先_氏名": "三木 暢人",
         "指揮命令者_氏名": "三木 信人", "責任の程度": "旧記載", "備考": ""},
        {"派遣先名称": "エリクソン・ジャパン株式会社", "氏名": "新人 未突合", "氏名カナ": "",
         "就業時間": "9:00～18:00", "備考": ""},
        {"派遣先名称": "テスト商事", "氏名": "無関係 太郎", "備考": ""},
    ]
    n_upd, unmatched = apply_ericsson_report(rows, master, source_name="test.csv")
    assert n_upd == 2
    assert unmatched == ["新人 未突合"]
    nara = master[0]
    assert nara["派遣先責任者_氏名"] == "Li Stephen"
    assert nara["指揮命令者_氏名"] == "Stephen Li"
    assert nara["就業時間"] == "9:00～18:00"                # 固定時間は上書き
    assert nara["苦情申出先_氏名"] == "三木 暢人"            # 氏名はレポートに無い＝維持
    assert nara["苦情申出先_役職"] == "Manager"
    ohta = master[1]
    assert ohta["就業時間"].startswith("シフト制")           # シフト制は就業時間を維持
    assert ohta["就業曜日"] == "月 火 水 木 金 土 日 祝 シフトあり"
    assert ohta["指揮命令者_氏名"] == "Nobuhito Miki"
    assert "変形労働時間制" in ohta["備考"]                  # Notesは備考へ
    assert "_エリクソン注記" in ohta and "シフト制のため" in ohta["_エリクソン注記"]
    assert master[3].get("_エリクソン注記") is None          # エリクソン以外は触らない
