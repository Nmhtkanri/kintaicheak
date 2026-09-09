# -*- coding: utf-8 -*-
"""所見（聴力・診察・胸部X線・心電図・眼底・胃部・腹部超音波）と医師名（2026-09-09）。

守ること:
  - 所見は原票の文章のまま所見列へ。聴力の「所見なし」も変換しない（値変換列が空なら）
  - 原票判定A〜Gは、変換マスタの「判定」行がある臓器別判定列へだけ出す。183〜197 のまとめ判定列は空のまま
  - 胃部は健診種別でX線／内視鏡へ振り分け、決められなければ出さない（警告）
  - 医師名は CSV に出さず、画面と監査用Excelに残す
"""

from __future__ import annotations

import io
import os
import sys
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from config import Config  # noqa: E402
from services.health_hpm_csv import (  # noqa: E402
    STOMACH_ENDOSCOPY,
    STOMACH_XRAY,
    build_person_row,
    stomach_method,
)
from services.health_hpm_excel import VT_FINDING, HealthMetric, PersonRecord, parse_health_workbook  # noqa: E402
from services.health_hpm_master import (  # noqa: E402
    JUDGEMENT_COLS,
    Course,
    MasterError,
    find_course,
    load_master,
    parse_value_map,
    resolve_institution,
)
from services.health_hpm_match import JinjerCandidate  # noqa: E402
from services.health_hpm_pdf import build_item_lookup, readings_to_person, write_audit_workbook  # noqa: E402
from tests.health_hpm_fixtures import DEFAULT_ITEM_MAP, employees_stub, make_master_xlsx  # noqa: E402
from tests.test_app_health_hpm_pdf import reading, tiny_png  # noqa: E402

DOYUKAI = "医療法人社団 同友会 春日クリニック"
COURSE_TEXT = "人間ドックＣ　胃カメラ（４０歳以上）"


@pytest.fixture
def master(tmp_path):
    return load_master(make_master_xlsx(tmp_path / "master.xlsx"))


def finding(category, item, value, judgement=""):
    return HealthMetric(category=category, item=item, occurrence=1, value=value,
                        value_type=VT_FINDING, source_judgement=judgement,
                        original_display=value)


def person_with(metrics, doctor=""):
    return PersonRecord(key="p02", name="友納 英彦", age=58, gender="男性",
                        exam_date=date(2026, 7, 1), exam_no="000132",
                        sheet="P02_友納英彦", metrics=list(metrics), doctor_name=doctor)


def employee():
    return JinjerCandidate(employee_id="2008003", last_name="友納", first_name="英彦",
                           last_kana="トモノウ", first_kana="ヒデヒコ",
                           birth_date=date(1968, 4, 13), gender="男性")


def row_for(master, metrics, course_value=COURSE_TEXT):
    inst = resolve_institution(master, DOYUKAI)
    course = find_course(master, DOYUKAI, course_value)
    return build_person_row(person_with(metrics), employee(), course, inst, master)


# ---------------------------------------------------------------------------
# 変換マスタ
# ---------------------------------------------------------------------------

class TestMaster:
    def test_finding_and_judgement_rows_coexist(self, master):
        rules = master.rules_for("腹部超音波", "肝臓", 1)
        assert sorted((r.value_type, r.hpm_col) for r in rules) == [("判定", 219), ("所見", 220)]

    def test_value_map_parsed(self, master):
        left = next(r for r in master.item_map if r.item == "聴力1000（左）")
        right = next(r for r in master.item_map if r.item == "聴力1000（右）")
        assert left.value_map == (("所見なし", "正常"), ("所見あり", "異常"))
        assert right.value_map == ()

    @pytest.mark.parametrize("text, expected", [
        ("", ()),
        ("a=b", (("a", "b"),)),
        ("a=b;c=d", (("a", "b"), ("c", "d"))),
        ("a=b；c=d", (("a", "b"), ("c", "d"))),
        (" a = b ;", (("a", "b"),)),
    ])
    def test_parse_value_map(self, text, expected):
        assert parse_value_map(text) == expected

    @pytest.mark.parametrize("text", ["ab", "=b", "a="])
    def test_parse_value_map_rejects(self, text):
        with pytest.raises(MasterError):
            parse_value_map(text, "聴力")

    def test_judgement_without_finding_is_rejected(self, tmp_path):
        item_map = list(DEFAULT_ITEM_MAP) + [("眼底", "眼底", 1, 237, "判定", "", "")]
        with pytest.raises(MasterError, match="対になる"):
            load_master(make_master_xlsx(tmp_path / "m.xlsx", item_map=item_map))

    def test_judgement_row_cannot_point_to_summary_judgement_cols(self, tmp_path):
        item_map = list(DEFAULT_ITEM_MAP) + [
            ("聴力", "聴力1000（右）", 1, 184, "判定", "", ""),  # 聴力判定=184 はまとめ判定列
        ]
        with pytest.raises(MasterError, match="判定列"):
            load_master(make_master_xlsx(tmp_path / "m.xlsx", item_map=item_map))

    def test_old_master_without_value_map_column_still_loads(self, tmp_path):
        """値変換列を持たない既存マスタも読める（列は任意）。"""
        import openpyxl
        path = make_master_xlsx(tmp_path / "m.xlsx")
        wb = openpyxl.load_workbook(path)
        ws = wb["項目マッピング"]
        ws.delete_cols(8)
        wb.save(path)
        m = load_master(path)
        assert all(r.value_map == () for r in m.item_map)


# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------

class TestCsv:
    def test_hearing_written_verbatim(self, master):
        row, issues = row_for(master, [finding("聴力", "聴力1000（右）", "所見なし")])
        assert row[41] == "所見なし", "聴力は原票のまま（変換列が空なら変換しない）"
        assert row[184] == "", "聴力判定（まとめ判定列）は書かない"
        assert not [i for i in issues if i.code.startswith("FINDING")]

    def test_value_map_applies_when_configured(self, master):
        row, issues = row_for(master, [finding("聴力", "聴力1000（左）", "所見なし")])
        assert row[40] == "正常"
        row, issues = row_for(master, [finding("聴力", "聴力1000（左）", "難聴の疑い")])
        assert row[40] == "難聴の疑い", "表に無い値はそのまま出す"
        assert any(i.code == "FINDING_VALUE_UNMAPPED" for i in issues)

    def test_finding_and_organ_judgement(self, master):
        row, issues = row_for(master, [finding("腹部超音波", "肝臓", "脂肪肝", "B"),
                                       finding("診察", "診察", "異常なし", "A")])
        assert row[220] == "脂肪肝" and row[219] == "B"
        assert row[204] == "異常なし" and row[203] == "A"
        assert all(row[c] == "" for c in JUDGEMENT_COLS), "まとめ判定列は空のまま"
        assert not [i for i in issues if i.code.startswith("FINDING")]

    def test_judgement_missing_leaves_column_blank(self, master):
        row, _ = row_for(master, [finding("腹部超音波", "肝臓", "異常なし")])
        assert row[220] == "異常なし" and row[219] == ""

    def test_judgement_with_digit_allowed_and_garbage_warned(self, master):
        row, issues = row_for(master, [finding("腹部超音波", "肝臓", "肝嚢胞", "C3")])
        assert row[219] == "C3"
        row, issues = row_for(master, [finding("腹部超音波", "肝臓", "肝嚢胞", "要精密")])
        assert row[219] == "" and row[220] == "肝嚢胞"
        assert any(i.code == "FINDING_JUDGEMENT_INVALID" for i in issues)

    def test_stomach_routed_by_course(self, master):
        # 同友会の胃カメラコース（テキスト） → 内視鏡列
        row, issues = row_for(master, [finding("胃部", "胃部", "萎縮性胃炎", "B")])
        assert row[218] == "萎縮性胃炎" and row[217] == "B"
        assert row[216] == "" and row[215] == ""
        # コード 10（定期健診）では判別できない → 出さない・警告
        row, issues = row_for(master, [finding("胃部", "胃部", "萎縮性胃炎", "B")],
                              course_value="10")
        assert row[215:219] == ["", "", "", ""]
        assert any(i.code == "FINDING_METHOD_UNKNOWN" for i in issues)

    @pytest.mark.parametrize("display, value, expected", [
        ("人間ドックB（胃部X線：バリウム）", "12", STOMACH_XRAY),
        ("人間ドックC（胃部内視鏡）", "13", STOMACH_ENDOSCOPY),
        ("人間ドックC 胃カメラ（40歳以上）", "人間ドックＣ　胃カメラ（４０歳以上）", STOMACH_ENDOSCOPY),
        ("定期健康診断（基本健診）", "10", ""),
        ("人間ドックA（胃部検査なし）", "11", ""),
    ])
    def test_stomach_method(self, display, value, expected):
        assert stomach_method(Course(institution="x", display_name=display, hpm_value=value)) == expected

    def test_unmapped_finding_is_warned_not_written(self, master):
        row, issues = row_for(master, [finding("心電図", "負荷心電図", "異常なし", "A")])
        assert "異常なし" not in row
        assert any(i.code == "UNMAPPED_ITEM" and "負荷心電図" in i.message for i in issues)

    def test_doctor_name_never_in_csv(self, master):
        inst = resolve_institution(master, DOYUKAI)
        course = find_course(master, DOYUKAI, COURSE_TEXT)
        person = person_with([finding("診察", "診察", "異常なし", "A")], doctor="堀口 実")
        row, _ = build_person_row(person, employee(), course, inst, master)
        assert "堀口 実" not in row and "堀口　実" not in row


# ---------------------------------------------------------------------------
# PDF 読み取り
# ---------------------------------------------------------------------------

class TestPdfReading:
    @pytest.fixture
    def item_lookup(self, master):
        return build_item_lookup(master)

    def raw(self, **over):
        base = reading(findings=[
            {"category": "聴力", "item": "聴力1000（右）", "judgement": "A", "value": "所見なし"},
            {"category": "聴力", "item": "聴力1KHz（左）", "judgement": "A", "value": "所見なし"},
            {"category": "腹部超音波", "item": "肝臓", "judgement": "A", "value": "異常なし"},
            {"category": "腹部超音波", "item": "脾他", "judgement": "b", "value": "副脾"},
            {"category": "胃部", "item": "胃部", "judgement": None, "value": None},
            {"category": "診察", "item": "診察", "judgement": "A", "value": ""},
            {"category": "心電図", "item": "心電図（安）", "judgement": "A", "value": "異常なし"},
        ], doctor="堀口 実")
        base.update(over)
        return base

    def test_findings_become_metrics(self, item_lookup):
        person = readings_to_person(self.raw(), 2, item_lookup)
        found = {m.item: m for m in person.findings()}
        assert set(found) == {"聴力1000（右）", "聴力1000（左）", "肝臓", "脾臓", "安静時心電図"}
        assert found["聴力1000（右）"].value == "所見なし"
        assert found["肝臓"].category == "腹部超音波" and found["肝臓"].source_judgement == "A"
        assert found["脾臓"].source_judgement == "B", "判定は大文字にそろえる"
        assert found["安静時心電図"].category == "心電図", "マスタに無い項目は読み取り側の分類"
        assert person.doctor_name == "堀口 実"

    def test_blank_findings_are_not_stored(self, item_lookup):
        person = readings_to_person(self.raw(), 2, item_lookup)
        assert "胃部" not in {m.item for m in person.findings()}
        assert "診察" not in {m.item for m in person.findings()}, "判定だけで所見が空なら出さない"

    def test_no_findings_key_is_fine(self, item_lookup):
        person = readings_to_person(reading(), 2, item_lookup)
        assert person.findings() == [] and person.doctor_name == ""

    def test_findings_do_not_pollute_numeric_or_qualitative(self, item_lookup):
        person = readings_to_person(self.raw(), 2, item_lookup)
        assert all(m.value_type != VT_FINDING for m in person.numeric() + person.qualitative())

    def test_audit_workbook_roundtrip_keeps_findings_and_doctor(self, item_lookup, master, tmp_path):
        from services.health_hpm_excel import WorkbookParseResult
        person = readings_to_person(self.raw(), 1, item_lookup)
        parse = WorkbookParseResult(schema_version="2.0", genpyo_confirmed=False,
                                    bp_occurrence_kept=True, qualitative_kept=True,
                                    source_filename="x.pdf", persons=[person])
        out = write_audit_workbook(str(tmp_path / "audit.xlsx"), parse, {person.key: 1},
                                   {1: tiny_png()}, pdf_name="x.pdf", confirmed_at="2026-09-09")
        back = parse_health_workbook(out, "audit.xlsx")
        assert not back.errors(), [i.message for i in back.errors()]
        p = back.persons[0]
        assert getattr(p, "doctor_name", "") == "堀口 実"
        assert {m.item: (m.value, m.source_judgement) for m in p.findings()}["肝臓"] == ("異常なし", "A")


# ---------------------------------------------------------------------------
# 画面へのペイロード
# ---------------------------------------------------------------------------

class TestPayload:
    def test_person_payload_has_findings_and_doctor(self, master):
        person = person_with([finding("腹部超音波", "肝臓", "脂肪肝", "B"),
                              finding("胃部", "胃部", "慢性胃炎", "B"),
                              finding("心電図", "負荷心電図", "異常なし", "A")], doctor="堀口 実")
        from services.health_hpm_match import MatchResult
        match = MatchResult(status="ok", employee_id="2008003", candidates=[employee()], reasons=[])
        payload = app_module._health_person_payload(person, match, master)
        assert payload["doctor_name"] == "堀口 実"
        assert payload["finding_count"] == 3
        by_item = {f["item"]: f for f in payload["findings"]}
        assert by_item["肝臓"]["hpm_col"] == 220 and by_item["肝臓"]["judgement_col"] == 219
        assert by_item["肝臓"]["judgement"] == "B"
        assert by_item["胃部"]["hpm_col"] is None and "健診種別" in by_item["胃部"]["note"]
        assert by_item["負荷心電図"]["hpm_col"] is None and "変換マスタに無い" in by_item["負荷心電図"]["note"]
        assert "負荷心電図" in [u["item"] for u in payload["unmapped_items"]]


def test_ui_and_prompt_wiring():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    js = open(os.path.join(root, "static", "script.js"), encoding="utf-8").read()
    pdf = open(os.path.join(root, "services", "health_hpm_pdf.py"), encoding="utf-8").read()
    assert "person.findings" in js and "所見（原票の文章どおり）" in js
    assert "person.doctor_name" in js
    assert '"findings": [' in pdf and '"doctor":' in pdf
    assert "「簡易」の欄は読まない" in pdf
