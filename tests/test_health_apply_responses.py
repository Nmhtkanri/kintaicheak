# -*- coding: utf-8 -*-
"""健康診断申込: 対象者×回答の突合（バケット分類・取込不可ルール・並び）。"""

from services.health_apply import responses as R
from services.health_apply import schema as S
from services.health_apply.options import OptionCatalog
from tests.health_apply_fixtures import options_rows, response_row, target_row


def catalog():
    return OptionCatalog.from_rows(S.rows_to_dicts(S.OPTION_HEADERS, options_rows()[1:]))


def targets(*rows):
    return S.rows_to_dicts(S.TARGET_HEADERS, [list(r) for r in rows])


def responses(*rows):
    return S.rows_to_dicts(S.RESPONSE_HEADERS, [list(r) for r in rows])


def report(target_rows, response_rows, year=2027):
    return R.build_report(targets(*target_rows), responses(*response_rows), catalog(), year)


def view_of(rep, emp):
    return next(r for r in rep["rows"] if r["employee_id"] == emp)


def codes(rep, emp):
    return [i["code"] for i in view_of(rep, emp)["issues"]]


# --- バケット ---------------------------------------------------------------

def test_buckets_follow_target_progress():
    rep = report([
        target_row(社員番号="2099001"),
        target_row(社員番号="2099002", 申込状態=S.STATUS_SENT, 送信日時="2027-02-01T09:00:00"),
        target_row(社員番号="2099003", 申込状態=S.STATUS_SENT, 送信日時="2027-02-01T09:00:00",
                   初回アクセス日時="2027-02-01T09:05:00"),
        target_row(社員番号="2099004", 申込状態=S.STATUS_ANSWERED, 送信日時="2027-02-01T09:00:00",
                   初回アクセス日時="2027-02-01T09:05:00", 受付番号="HC-2027-2099004-01", 回答版="1"),
        target_row(社員番号="2099005", 申込状態=S.STATUS_REANSWER, 送信日時="2027-02-01T09:00:00",
                   受付番号="HC-2027-2099005-01", 回答版="1"),
        target_row(社員番号="2099006", 申込状態=S.STATUS_INVALID),
    ], [
        response_row(社員番号="2099004", 受付番号="HC-2027-2099004-01"),
        response_row(社員番号="2099005", 受付番号="HC-2027-2099005-01"),
    ])
    assert rep["counts"] == {"targets": 6, "unsent": 1, "sent_not_accessed": 1, "accessed_only": 1,
                             "answered": 1, "reanswer_pending": 1, "invalid": 1, "error": 0}
    assert view_of(rep, "2099001")["bucket"] == "unsent"
    assert view_of(rep, "2099002")["bucket"] == "sent_not_accessed"
    assert view_of(rep, "2099003")["bucket"] == "accessed_only"
    assert view_of(rep, "2099004")["bucket"] == "answered"
    assert view_of(rep, "2099004")["importable"] is True
    assert view_of(rep, "2099005")["bucket"] == "reanswer_pending"
    assert view_of(rep, "2099006")["bucket"] == "invalid"
    assert rep["workbook_issues"] == []


def test_latest_response_is_max_version_not_latest_time():
    rep = report(
        [target_row(社員番号="2099001", 申込状態=S.STATUS_ANSWERED, 受付番号="HC-2027-2099001-02", 回答版="2",
                    送信日時="2027-02-01T09:00:00")],
        [response_row(社員番号="2099001", 受付番号="HC-2027-2099001-02", 回答版="2", 回答日時="2027-02-03T00:00:00",
                      健診機関コード="0301619", 健診機関名="医療法人徳洲会 生駒市立病院"),
         response_row(社員番号="2099001", 受付番号="HC-2027-2099001-01", 回答版="1", 回答日時="2027-02-09T00:00:00")],
    )
    v = view_of(rep, "2099001")
    assert v["bucket"] == "answered"
    assert v["history_count"] == 2
    assert v["latest"]["version"] == "2"
    assert v["latest"]["institution"]["code"] == "0301619"
    assert v["latest"]["exam_type"]["name"] == "人間ドックC"


def test_answered_rows_carry_display_names_from_catalog():
    rep = report(
        [target_row(社員番号="2099001", 申込状態=S.STATUS_ANSWERED, 受付番号="HC-2027-2099001-01", 回答版="1",
                    送信日時="x")],
        [response_row(社員番号="2099001", 追加検査="GYN", 被扶養者申込="1", 続柄="妻", 被扶養者氏名="試験 花子",
                      健診機関コード="OTHER", 健診機関名="その他", その他医療機関名="南町健診センター")],
    )
    latest = view_of(rep, "2099001")["latest"]
    assert latest["extras"] == [{"code": "GYN", "name": "婦人科検診", "known": True, "active": True}]
    assert latest["dependent"] == {"requested": True, "relationship": "妻", "name": "試験 花子"}
    assert latest["other_institution"] == "南町健診センター"
    assert latest["kind_label"] == "変更"
    assert view_of(rep, "2099001")["target"]["previous"]["institution"]["name"] == "医療法人社団 同友会 春日クリニック"
    assert view_of(rep, "2099001")["target"]["enrollment_label"] == "在籍"


# --- 取込不可 -----------------------------------------------------------------

def base_answered(**kw):
    t = dict(社員番号="2099001", 申込状態=S.STATUS_ANSWERED, 受付番号="HC-2027-2099001-01", 回答版="1", 送信日時="x")
    t.update(kw)
    return target_row(**t)


def test_duplicate_version_and_receipt_are_errors():
    rep = report([base_answered()], [
        response_row(社員番号="2099001"),
        response_row(社員番号="2099001", 回答日時="2027-02-06T00:00:00"),
    ])
    assert view_of(rep, "2099001")["bucket"] == "error"
    assert "duplicate_version" in codes(rep, "2099001")
    assert "duplicate_receipt" in codes(rep, "2099001")


def test_receipt_shared_between_two_employees_flags_both():
    rep = report([base_answered(), base_answered(社員番号="2099002", 受付番号="HC-2027-2099001-01")], [
        response_row(社員番号="2099001"),
        response_row(社員番号="2099002", 受付番号="HC-2027-2099001-01"),
    ])
    assert "duplicate_receipt" in codes(rep, "2099001")
    assert "duplicate_receipt" in codes(rep, "2099002")


def test_response_from_non_target_is_error_and_not_counted_as_target():
    rep = report([target_row(社員番号="2099001")], [response_row(社員番号="2088888")])
    assert rep["counts"]["targets"] == 1 and rep["counts"]["error"] == 1
    v = view_of(rep, "2088888")
    assert v["bucket"] == "error" and v["issues"][0]["code"] == "not_a_target"
    assert v["target"] == {}


def test_year_mismatch_in_response():
    rep = report([base_answered()], [response_row(社員番号="2099001", 年度="2026")])
    assert "year_mismatch" in codes(rep, "2099001")


def test_unknown_choice_codes_are_errors_and_inactive_is_warning():
    rep = report([base_answered()], [
        response_row(社員番号="2099001", 健診機関コード="9999999", 健診種別コード="99", 追加検査="XRAY"),
    ])
    c = codes(rep, "2099001")
    assert {"unknown_institution", "unknown_exam_type", "unknown_extra"} <= set(c)
    rep = report([base_answered()], [
        response_row(社員番号="2099001", 健診機関コード="130192", 健診機関名="東京品川病院 総合健診センター"),
    ])
    v = view_of(rep, "2099001")
    assert v["bucket"] == "answered"
    assert [i["level"] for i in v["issues"] if i["code"] == "inactive_institution"] == ["warning"]


def test_status_mismatches_between_target_and_response():
    rep = report([base_answered()], [])
    assert "status_answered_without_response" in codes(rep, "2099001")

    rep = report([target_row(社員番号="2099001")], [response_row(社員番号="2099001")])
    assert "status_unsent_with_response" in codes(rep, "2099001")

    rep = report([base_answered(受付番号="HC-2027-2099001-09")], [response_row(社員番号="2099001")])
    assert "receipt_mismatch" in codes(rep, "2099001")

    rep = report([base_answered(回答版="2")], [response_row(社員番号="2099001")])
    assert "version_mismatch" in codes(rep, "2099001")


def test_same_without_previous_is_error():
    rep = report(
        [base_answered(前年度情報元=S.SOURCE_NONE, 前年度健診機関コード="", 前年度健診機関名="",
                       前年度健診種別コード="", 前年度健診種別名="")],
        [response_row(社員番号="2099001", 申込区分=S.KIND_SAME)],
    )
    assert "same_without_previous" in codes(rep, "2099001")


def test_bad_kind_and_bad_version():
    rep = report([base_answered()], [response_row(社員番号="2099001", 申込区分="keep", 回答版="一")])
    assert {"bad_kind", "bad_version"} <= set(codes(rep, "2099001"))


def test_warnings_do_not_block_import():
    rep = report(
        [base_answered(在籍区分="1")],
        [response_row(社員番号="2099001", 健診機関コード="OTHER", 健診機関名="その他",
                      被扶養者申込="1", 続柄="", 被扶養者氏名="")],
    )
    v = view_of(rep, "2099001")
    assert v["bucket"] == "answered" and v["importable"] is True
    assert {i["code"] for i in v["issues"]} == {"not_enrolled", "other_without_name",
                                                "dependent_without_relationship", "dependent_without_name"}
    assert all(i["level"] == "warning" for i in v["issues"])


def test_unknown_relationship_is_error():
    rep = report([base_answered()], [response_row(社員番号="2099001", 被扶養者申込="1", 続柄="子", 被扶養者氏名="x")])
    assert "unknown_relationship" in codes(rep, "2099001")


# --- シート全体 -----------------------------------------------------------------

def test_target_sheet_issues_other_year_duplicate_blank():
    rep = report([
        target_row(社員番号="2099001"),
        target_row(社員番号="2099001", 氏名="別人"),
        target_row(社員番号="", 氏名="番号なし"),
        target_row(社員番号="2088001", 年度="2026"),
    ], [])
    assert rep["counts"]["targets"] == 1
    codes_ = [i["code"] for i in rep["workbook_issues"]]
    assert codes_ == ["target_duplicate", "target_blank_id", "target_other_year"]
    assert view_of(rep, "2099001")["name"] == "試験 太郎"


def test_response_with_blank_employee_id_is_workbook_issue():
    rep = report([target_row(社員番号="2099001")], [response_row(社員番号="")])
    assert [i["code"] for i in rep["workbook_issues"]] == ["response_blank_id"]


def test_rows_are_sorted_errors_first_then_bucket_then_id():
    rep = report([
        target_row(社員番号="2099009"),
        target_row(社員番号="2099002", 申込状態=S.STATUS_SENT, 送信日時="x"),
        base_answered(社員番号="2099005"),
        base_answered(社員番号="2099001", 受付番号="HC-2027-2099001-01"),
    ], [
        response_row(社員番号="2099005", 受付番号="HC-2027-2099005-01"),
        response_row(社員番号="2099001", 受付番号="HC-2027-2099001-01", 健診種別コード="99"),
        response_row(社員番号="2077777"),
    ])
    assert [r["employee_id"] for r in rep["rows"]] == ["2077777", "2099001", "2099005", "2099002", "2099009"]
    assert [r["bucket"] for r in rep["rows"]] == ["error", "error", "answered", "sent_not_accessed", "unsent"]


def test_split_codes_accepts_common_separators():
    assert R.split_codes("GYN;X、Y, Z；") == ["GYN", "X", "Y", "Z"]
    assert R.split_codes("") == []
