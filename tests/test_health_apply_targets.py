# -*- coding: utf-8 -*-
"""健康診断申込: 対象者登録の差分計画（社員番号解析・前年度の解決・add/unchanged/conflict/blocked・冪等）。"""

from services.health_apply import audit as AU
from services.health_apply import schema as S
from services.health_apply import targets as T
from services.health_apply.jinjer_source import EmployeeProfile, PreviousRaw
from services.health_apply.options import OptionCatalog
from tests.health_apply_fixtures import options_rows, response_row, target_row


def catalog():
    return OptionCatalog.from_rows(S.rows_to_dicts(S.OPTION_HEADERS, options_rows()[1:]))


def profile(emp_id, name="試験 太郎", email="t.shiken@nmht.co.jp", enr="0", enr_name="在籍", retire=""):
    return EmployeeProfile(emp_id, name, email, enr, enr_name, retire)


def raw(source="履歴", year="2026", inst="医療法人社団 同友会 春日クリニック", labels=("基本健診",),
        date="2026-07-01", notes=()):
    return PreviousRaw(source, year, inst, list(labels), date, list(notes))


def targets(*rows):
    return S.rows_to_dicts(S.TARGET_HEADERS, [list(r) for r in rows])


def responses(*rows):
    return S.rows_to_dicts(S.RESPONSE_HEADERS, [list(r) for r in rows])


# --- 社員番号の解析 ---------------------------------------------------------------

def test_parse_employee_ids_accepts_pasted_shapes():
    ids, issues = T.parse_employee_ids("2099001\n２０９９００２, 2099003;2099001\n\n12345 abc\n2099004\t2099005")
    assert ids == ["2099001", "2099002", "2099003", "2099004", "2099005"]
    assert [(i.level, i.code) for i in issues] == [
        ("warning", "duplicate_employee_id"), ("error", "bad_employee_id"), ("error", "bad_employee_id")]
    assert "4行目" in issues[1].message and "12345" in issues[1].message
    assert T.parse_employee_ids("") == ([], [])
    assert T.parse_employee_ids(None) == ([], [])


# --- 前年度をコードへ ------------------------------------------------------------

def test_resolve_previous_maps_institution_alias_type_and_extra():
    prev = T.resolve_previous(raw(inst="同友会", labels=("基本健診 ", "婦人病検査")), catalog())
    assert prev.source == "履歴"
    assert (prev.institution_code, prev.institution_name) == ("1310528885", "医療法人社団 同友会 春日クリニック")
    assert prev.institution_raw == "同友会"
    assert (prev.exam_type_code, prev.exam_type_name) == ("10", "定期健康診断")
    assert prev.extra_codes == ("GYN",) and prev.extra_names == ("婦人科検診",)
    assert prev.notes == []
    assert prev.as_dict()["extras"] == [{"code": "GYN", "name": "婦人科検診"}]


def test_resolve_previous_unknown_institution_becomes_none_but_keeps_raw():
    prev = T.resolve_previous(raw(inst="どこかの病院", labels=("1日人間ドック・胃カメラ",)), catalog())
    assert prev.source == "なし"
    assert prev.institution_code == "" and prev.institution_raw == "どこかの病院"
    assert prev.content_labels == ("1日人間ドック・胃カメラ",)
    assert any("どこかの病院" in n and "別名" in n for n in prev.notes)


def test_resolve_previous_needs_exactly_one_exam_type():
    prev = T.resolve_previous(raw(labels=("婦人病検査",)), catalog())
    assert prev.source == "なし" and any("健診種別を決められません" in n for n in prev.notes)
    prev = T.resolve_previous(raw(labels=("基本健診", "1日人間ドック")), catalog())
    assert prev.source == "なし" and any("複数" in n for n in prev.notes)
    prev = T.resolve_previous(raw(labels=("基本健診", "謎の検査")), catalog())
    assert prev.source == "履歴" and any("謎の検査" in n for n in prev.notes)


def test_resolve_previous_none_source_and_missing():
    assert T.resolve_previous(PreviousRaw(), catalog()).source == "なし"
    assert T.resolve_previous(None, catalog()).source == "なし"
    prev = T.resolve_previous(raw(source="現在値", year="", inst="医療法人徳洲会 生駒市立病院", labels=("1日人間ドック・バリウム",),
                                  notes=["現在値を使いました"]), catalog())
    assert prev.source == "現在値" and prev.exam_type_code == "12" and prev.notes == ["現在値を使いました"]


# --- 候補 ------------------------------------------------------------------------

def test_build_candidates_flags_missing_profile_bad_email_and_retired():
    profiles = {
        "2099001": profile("2099001"),
        "2099002": profile("2099002", "空 メール", ""),
        "2099003": profile("2099003", "外 ドメイン", "x@example.com"),
        "2099004": profile("2099004", "退 職", "r@nmht.co.jp", "1", "退職", "2026-03-31"),
    }
    prev = {"2099001": raw(), "2099004": raw()}
    cands = T.build_candidates(["2099001", "2099002", "2099003", "2099004", "2099009"], profiles, prev, catalog())
    by = {c.employee_id: c for c in cands}
    assert by["2099001"].issues == [] and by["2099001"].blocked is False
    assert [i.code for i in by["2099002"].issues] == ["email_missing", "no_previous"]
    assert by["2099002"].blocked is True
    assert [i.code for i in by["2099003"].issues][0] == "email_domain"
    assert [i.code for i in by["2099004"].issues] == ["not_enrolled"]
    assert "2026-03-31" in by["2099004"].issues[0].message and by["2099004"].blocked is False
    assert [i.code for i in by["2099009"].issues] == ["not_in_jinjer"] and by["2099009"].blocked is True


def test_previous_notes_become_warnings():
    cands = T.build_candidates(["2099001"], {"2099001": profile("2099001")},
                               {"2099001": raw(inst="不明院")}, catalog())
    codes = [i.code for i in cands[0].issues]
    assert codes == ["previous_note", "no_previous"]
    assert cands[0].previous.source == "なし"


# --- 差分計画 ------------------------------------------------------------------------

def plan_for(existing=(), responded=(), **profiles_kw):
    profiles = {"2099001": profile("2099001"), "2099002": profile("2099002", "二 号", "n2@nmht.co.jp"),
                "2099003": profile("2099003", "三 号", "")}
    prev = {"2099001": raw(), "2099002": raw(inst="医療法人徳洲会 生駒市立病院", labels=("1日人間ドック・胃カメラ", "婦人病検査"))}
    cands = T.build_candidates(["2099001", "2099002", "2099003"], profiles, prev, catalog())
    return T.plan_targets(cands, targets(*existing), responses(*responded), 2027)


def test_plan_add_unchanged_conflict_blocked():
    plan = plan_for(existing=[
        target_row(社員番号="2099002", 氏名="二 号", 社用メール="n2@nmht.co.jp", 前年度健診機関コード="0301619",
                   前年度健診機関名="医療法人徳洲会 生駒市立病院", 前年度健診種別コード="13", 前年度健診種別名="人間ドックC",
                   前年度追加検査="GYN"),
    ])
    actions = {r.candidate.employee_id: r.action for r in plan.rows}
    assert actions == {"2099001": "add", "2099002": "unchanged", "2099003": "blocked"}
    assert plan.counts() == {"input": 3, "add": 1, "unchanged": 1, "conflict": 0, "blocked": 1, "warnings": 1, "input_errors": 0}
    assert plan.can_commit() is False           # blocked があるうちは登録しない
    assert plan.as_dict()["confirm_phrase"] == "REGISTER 2027 1"


def test_plan_conflict_when_sheet_differs_lists_columns():
    plan = plan_for(existing=[target_row(社員番号="2099001", 氏名="別 名", 前年度健診種別コード="13", 前年度健診種別名="人間ドックC")])
    row = next(r for r in plan.rows if r.candidate.employee_id == "2099001")
    assert row.action == "conflict"
    assert row.reasons == ["氏名: シート「別 名」→ jinjer「試験 太郎」", "前年度健診種別コード: シート「13」→ jinjer「10」"]


def test_plan_conflict_when_already_answered_or_response_exists():
    plan = plan_for(existing=[target_row(社員番号="2099001", 申込状態=S.STATUS_ANSWERED, 受付番号="HC-2027-2099001-01")])
    row = next(r for r in plan.rows if r.candidate.employee_id == "2099001")
    assert row.action == "conflict" and "回答" in row.reasons[0]

    plan = plan_for(existing=[target_row(社員番号="2099001")], responded=[response_row(社員番号="2099001")])
    row = next(r for r in plan.rows if r.candidate.employee_id == "2099001")
    assert row.action == "conflict"


def test_plan_ignores_other_year_rows_and_enrollment_change_is_warning():
    plan = plan_for(existing=[
        target_row(社員番号="2099001", 年度="2026", 氏名="去年の名前"),
        target_row(社員番号="2099001", 在籍区分="2"),
    ])
    row = next(r for r in plan.rows if r.candidate.employee_id == "2099001")
    assert row.action == "unchanged"
    assert [i.code for i in row.candidate.issues] == ["enrollment_changed"]


def test_can_commit_requires_add_and_no_conflict():
    plan = plan_for(existing=[target_row(社員番号="2099001"),
                              target_row(社員番号="2099002", 氏名="二 号", 社用メール="n2@nmht.co.jp", 前年度健診機関コード="0301619",
                                         前年度健診種別コード="13", 前年度追加検査="GYN")])
    assert plan.counts()["add"] == 0 and plan.can_commit() is False
    plan2 = T.plan_targets([c for c in (r.candidate for r in plan.rows) if not c.blocked][:1], [], [], 2027)
    assert plan2.can_commit() is True


# --- 行・指紋・監査 ---------------------------------------------------------------------

def test_build_target_rows_column_order_and_hub_columns_only():
    plan = plan_for()
    rows = T.build_target_rows(plan, "yatsu", "2027-01-10T09:00:00")
    assert len(rows) == 2 and all(len(r) == S.TARGET_HUB_COLUMNS for r in rows)
    assert rows[0] == ["2027", "2099001", "試験 太郎", "t.shiken@nmht.co.jp", "0", "履歴",
                       "1310528885", "医療法人社団 同友会 春日クリニック", "10", "定期健康診断", "",
                       "医療法人社団 同友会 春日クリニック", "2027-01-10T09:00:00", "yatsu"]
    assert rows[1][1] == "2099002" and rows[1][10] == "GYN" and rows[1][6] == "0301619"


def test_fingerprint_is_order_independent_and_content_sensitive():
    a = plan_for()
    b = plan_for()
    b.rows.reverse()
    assert T.plan_fingerprint(a) == T.plan_fingerprint(b)
    next(r for r in b.rows if r.action == "add").candidate.email = "changed@nmht.co.jp"
    assert T.plan_fingerprint(a) != T.plan_fingerprint(b)
    assert len(T.plan_fingerprint(a)) == 64


def test_same_input_twice_is_idempotent():
    """1回目の追加行をシートに入れた体で2回目を計画すると、全員 unchanged になる。"""
    first = plan_for()
    written = T.build_target_rows(first, "yatsu", "2027-01-10T09:00:00")
    sheet_rows = [r + [""] * (len(S.TARGET_HEADERS) - len(r)) for r in written]
    second_existing = S.rows_to_dicts(S.TARGET_HEADERS, sheet_rows)
    cands = [r.candidate for r in first.rows]
    for c in cands:
        c.issues = [i for i in c.issues if i.code != "enrollment_changed"]
    second = T.plan_targets(cands, second_existing, [], 2027)
    assert [r.action for r in second.rows] == ["unchanged", "unchanged", "blocked"]
    assert second.counts()["add"] == 0


def test_audit_rows():
    plan = plan_for()
    added = [r.as_dict() for r in plan.by_action("add")]
    rows = AU.target_register_rows("yatsu", 2027, added, "2027-01-10T09:00:00", "abcdef0123456789")
    assert len(rows) == 3 and all(len(r) == len(S.AUDIT_HEADERS) for r in rows)
    assert rows[0] == ["2027-01-10T09:00:00", "REGISTER_BATCH", "Hub", "yatsu", "2027", "", "2名を対象者へ追記 fingerprint=abcdef012345"]
    assert rows[1][1] == "REGISTER" and rows[1][5] == "2099001"
    assert rows[2][6] == "前年度情報元=履歴 機関=0301619 種別=13 追加検査=GYN"
