# -*- coding: utf-8 -*-
"""健康診断申込 Apps Script v2（Code.gs）を node の偽 GAS 環境で検査する。

Hub の schema.py と列並びが同じであること、トークンのハッシュ、受付期間、回答の検証、
回答行の組み立て、案内メール送信の流れを固定する。実データは使わない。
"""

import hashlib
import json
import pathlib
import shutil
import subprocess

import pytest

from services.health_apply import schema as S
from tests.health_apply_fixtures import options_rows, settings_rows, target_row, workbook

ROOT = pathlib.Path(__file__).resolve().parent.parent
CODE = ROOT / "prototypes" / "health_check_application" / "apps_script_v2" / "Code.gs"
HARNESS = ROOT / "tests" / "health_apply_gas_harness.js"

pytestmark = pytest.mark.skipif(shutil.which("node") is None, reason="node が無い環境では飛ばす")

TOKEN = "tok-2099001-abcdef"
HASH = hashlib.sha256(TOKEN.encode("utf-8")).hexdigest()
NOW = "2027-02-05T10:00:00+09:00"   # 受付期間内の日時


def run_gas(wb, scenario, **extra):
    payload = {"code_path": str(CODE), "workbook": wb, "scenario": scenario}
    payload.update(extra)
    proc = subprocess.run(["node", str(HARNESS)], input=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                          capture_output=True, check=False)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    return json.loads(proc.stdout.decode("utf-8"))


def sent_target(**kw):
    base = dict(社員番号="2099001", 申込状態=S.STATUS_SENT, トークンハッシュ=HASH, 送信日時="2027-02-01T09:00:00", 送信回数="1")
    base.update(kw)
    return target_row(**base)


def change_payload(**kw):
    p = {"applicationType": "change", "clinicCode": "1310136358", "courseCode": "13", "extraCodes": ["GYN"],
         "customClinic": "", "otherPlannedDate": "", "dependentRequested": False, "dependentRelationship": "",
         "dependentName": "", "remarks": "テスト", "agreement": True}
    p.update(kw)
    return p


def rows_by_header(rows, headers):
    return [dict(zip(headers, r + [""] * (len(headers) - len(r)))) for r in rows[1:]]


# --- 列定義が Hub と同じ ---------------------------------------------------------------

def test_constants_match_hub_schema():
    out = run_gas(workbook(), "({TARGET_HEADERS, RESPONSE_HEADERS, AUDIT_HEADERS, OPTION_HEADERS, SETTINGS_HEADERS, SCHEMA_VERSION, SHEETS, REQUIRED_SETTING_KEYS})")
    r = out["result"]
    assert r["TARGET_HEADERS"] == list(S.TARGET_HEADERS)
    assert r["RESPONSE_HEADERS"] == list(S.RESPONSE_HEADERS)
    assert r["AUDIT_HEADERS"] == list(S.AUDIT_HEADERS)
    assert r["OPTION_HEADERS"] == list(S.OPTION_HEADERS)
    assert r["SETTINGS_HEADERS"] == list(S.SETTINGS_HEADERS)
    assert r["SCHEMA_VERSION"] == S.SCHEMA_VERSION
    assert sorted(r["SHEETS"].values()) == sorted(S.ALL_SHEETS)
    assert r["REQUIRED_SETTING_KEYS"] == list(S.REQUIRED_SETTING_KEYS)


def test_hash_token_is_sha256_hex():
    out = run_gas(workbook(), "[hashToken_('abc'), hashToken_('日本語トークン')]")
    assert out["result"] == [hashlib.sha256(b"abc").hexdigest(), hashlib.sha256("日本語トークン".encode("utf-8")).hexdigest()]


def test_is_accepting_boundaries():
    kv = "{'回答受付':'1','受付開始':'2027-02-01','受付終了':'2027-02-28'}"
    scenario = f"""[
        isAccepting_({kv}, new Date('2027-01-31T23:30:00+09:00')),
        isAccepting_({kv}, new Date('2027-02-01T00:30:00+09:00')),
        isAccepting_({kv}, new Date('2027-02-28T23:00:00+09:00')),
        isAccepting_({kv}, new Date('2027-03-01T00:30:00+09:00')),
        isAccepting_({{'回答受付':'0','受付開始':'2027-02-01','受付終了':'2027-02-28'}}, new Date('2027-02-10T12:00:00+09:00')),
        isAccepting_({{'回答受付':'1','受付開始':'','受付終了':''}}, new Date('2030-01-01T00:00:00+09:00')),
    ]"""
    assert run_gas(workbook(), scenario)["result"] == [False, True, True, False, False, True]


def test_receipt_and_version_helpers():
    out = run_gas(workbook(), "[receiptId_('2027','2099001',3), nextVersion_({'回答版':''}), nextVersion_({'回答版':'2'}), nextVersion_({'回答版':'x'})]")
    assert out["result"] == ["HC-2027-2099001-03", 1, 3, 1]


def test_require_settings_rejects_schema_mismatch_and_missing_keys():
    wb = workbook(settings=settings_rows(スキーマ版="1999.0"))
    assert "スキーマ版" in run_gas(wb, "requireSettings_(readSettings_())")["error"]
    wb = workbook(settings=settings_rows(受付終了=""))
    assert "受付終了" in run_gas(wb, "requireSettings_(readSettings_())")["error"]
    assert run_gas(workbook(), "requireSettings_(readSettings_())['年度']")["result"] == "2027"


def test_read_options_sorts_and_filters_active():
    out = run_gas(workbook(), "(() => { const o = readOptions_(); return {inst: activeOptions_(o, KIND.institution).map(x => x.code), all: o[KIND.institution].length, types: activeOptions_(o, KIND.examType).map(x => x.code)}; })()")
    assert out["result"]["inst"] == ["1310528885", "0301619", "13X5035440", "1310136358", "OTHER"]
    assert out["result"]["all"] == 6
    assert out["result"]["types"] == ["10", "11", "12", "13", "14", "15"]


# --- 回答の検証 -----------------------------------------------------------------------------

def validate(target_kwargs, payload):
    wb = workbook(targets=[sent_target(**target_kwargs)])
    scenario = (f"(() => {{ const o = readOptions_(); const t = findTargetByHash_('{HASH}'); const kv = readSettings_();"
                f" const n = validatePayload_(t, {json.dumps(payload, ensure_ascii=False)}, o, kv);"
                " return {type: n.applicationType, inst: n.institution.code, instName: n.institution.name, other: n.otherInstitution,"
                " date: n.otherPlannedDate, course: n.examType.code, extras: n.extras.map(e => e.code), dep: n.dependentRequested,"
                " rel: n.relationship, depName: n.dependentName, remarks: n.remarks}; })()")
    return run_gas(wb, scenario)


def test_validate_change_uses_codes_and_sheet_names():
    out = validate({}, change_payload())
    assert out["error"] is None
    assert out["result"] == {"type": "change", "inst": "1310136358", "instName": "MYメディカルクリニック 大手町", "other": "",
                             "date": "", "course": "13", "extras": ["GYN"], "dep": False, "rel": "", "depName": "", "remarks": "テスト"}


def test_validate_same_copies_previous_and_needs_previous():
    out = validate({}, {"applicationType": "same", "agreement": True})
    assert out["error"] is None
    assert (out["result"]["inst"], out["result"]["course"], out["result"]["extras"]) == ("1310528885", "10", [])
    out = validate({"前年度情報元": S.SOURCE_NONE, "前年度健診機関コード": "", "前年度健診種別コード": ""},
                   {"applicationType": "same", "agreement": True})
    assert "前年度の情報が無い" in out["error"]


@pytest.mark.parametrize("payload, message", [
    (change_payload(clinicCode="9999999"), "健診機関"),
    (change_payload(clinicCode="130192"), "健診機関"),                       # 無効化された機関は選べない
    (change_payload(courseCode="99"), "健診種別"),
    (change_payload(extraCodes=["XRAY"]), "追加検査"),
    (change_payload(clinicCode="OTHER", otherPlannedDate="2028-04-01"), "受診期間"),
    (change_payload(clinicCode="OTHER", otherPlannedDate="2027/05/01"), "形式"),
    (change_payload(dependentRequested=True, dependentRelationship="妻", dependentName=""), "続柄と氏名"),
    (change_payload(dependentRequested=True, dependentRelationship="子", dependentName="x"), "続柄と氏名"),
    (change_payload(agreement=False), "確認欄"),
    (change_payload(applicationType="keep"), "申込内容"),
])
def test_validate_rejections(payload, message):
    assert message in validate({}, payload)["error"]


def test_validate_other_clinic_and_dependent():
    out = validate({}, change_payload(clinicCode="OTHER", customClinic="南町健診センター", otherPlannedDate="2027-09-10",
                                       dependentRequested=True, dependentRelationship="夫", dependentName="試験 一郎"))
    assert out["error"] is None
    r = out["result"]
    assert (r["inst"], r["other"], r["date"], r["dep"], r["rel"], r["depName"]) == ("OTHER", "南町健診センター", "2027-09-10", True, "夫", "試験 一郎")


# --- 回答の受付 ------------------------------------------------------------------------------

def test_submit_application_appends_response_and_updates_target():
    wb = workbook(targets=[sent_target()])
    out = run_gas(wb, f"submitApplication('{TOKEN}', {json.dumps(change_payload(), ensure_ascii=False)})",
                  now="2027-02-05T10:00:00+09:00")
    assert out["error"] is None, out["error"]
    assert out["result"] == {"ok": True, "receiptId": "HC-2027-2099001-01", "employeeId": "2099001", "version": 1}
    responses = rows_by_header(out["sheets"][S.SHEET_RESPONSES], S.RESPONSE_HEADERS)
    assert len(responses) == 1 and len(out["sheets"][S.SHEET_RESPONSES][1]) == len(S.RESPONSE_HEADERS)
    r = responses[0]
    assert r["回答日時"] == "2027-02-05T10:00:00" and r["受付番号"] == "HC-2027-2099001-01"
    assert (r["年度"], r["社員番号"], r["回答版"], r["氏名"], r["社用メール"]) == ("2027", "2099001", "1", "試験 太郎", "t.shiken@nmht.co.jp")
    assert (r["申込区分"], r["健診機関コード"], r["健診機関名"], r["健診種別コード"], r["健診種別名"], r["追加検査"]) == \
        ("change", "1310136358", "MYメディカルクリニック 大手町", "13", "人間ドックC", "GYN")
    assert (r["被扶養者申込"], r["トークンハッシュ"], r["回答元"]) == ("0", HASH, "Web")
    target = rows_by_header(out["sheets"][S.SHEET_TARGETS], S.TARGET_HEADERS)[0]
    assert (target["申込状態"], target["受付番号"], target["回答版"], target["回答日時"]) == (S.STATUS_ANSWERED, "HC-2027-2099001-01", "1", "2027-02-05T10:00:00")
    audit = out["sheets"][S.SHEET_AUDIT]
    assert audit[-1][1] == "SUBMIT" and audit[-1][2] == "AppsScript" and audit[-1][5] == "2099001"
    # Hub 側の突合でも取込可になる形
    from services.health_apply.options import OptionCatalog
    from services.health_apply.responses import build_report
    catalog = OptionCatalog.from_rows(S.rows_to_dicts(S.OPTION_HEADERS, options_rows()[1:]))
    report = build_report([target], responses, catalog, 2027)
    assert report["counts"]["answered"] == 1 and report["rows"][0]["importable"] is True


def test_submit_twice_is_rejected_until_reopened():
    wb = workbook(targets=[sent_target(申込状態=S.STATUS_ANSWERED, 受付番号="HC-2027-2099001-01", 回答版="1")])
    out = run_gas(wb, f"submitApplication('{TOKEN}', {json.dumps(change_payload(), ensure_ascii=False)})", now=NOW)
    assert "既に申込済み" in out["error"] and "HC-2027-2099001-01" in out["error"]
    assert len(out["sheets"][S.SHEET_RESPONSES]) == 1

    wb = workbook(targets=[sent_target(申込状態=S.STATUS_REANSWER, 受付番号="HC-2027-2099001-01", 回答版="1")])
    out = run_gas(wb, f"submitApplication('{TOKEN}', {json.dumps({'applicationType': 'same', 'agreement': True}, ensure_ascii=False)})", now=NOW)
    assert out["error"] is None
    assert out["result"]["receiptId"] == "HC-2027-2099001-02" and out["result"]["version"] == 2
    target = rows_by_header(out["sheets"][S.SHEET_TARGETS], S.TARGET_HEADERS)[0]
    assert (target["申込状態"], target["回答版"]) == (S.STATUS_ANSWERED, "2")


def test_submit_rejects_bad_token_unsent_and_closed():
    wb = workbook(targets=[sent_target()])
    assert "無効" in run_gas(wb, f"submitApplication('wrong', {json.dumps(change_payload())})", now=NOW)["error"]
    wb = workbook(targets=[sent_target(申込状態=S.STATUS_UNSENT, 送信日時="")])
    assert "現在使えません" in run_gas(wb, f"submitApplication('{TOKEN}', {json.dumps(change_payload())})", now=NOW)["error"]
    wb = workbook(targets=[sent_target()], settings=settings_rows(回答受付="0"))
    assert "受付期間外" in run_gas(wb, f"submitApplication('{TOKEN}', {json.dumps(change_payload())})", now=NOW)["error"]
    wb = workbook(targets=[sent_target()])
    assert "受付期間外" in run_gas(wb, f"submitApplication('{TOKEN}', {json.dumps(change_payload())})", now="2027-03-15T10:00:00+09:00")["error"]


def test_first_access_is_recorded_once():
    wb = workbook(targets=[sent_target()])
    scenario = f"(() => {{ const t = findTargetByHash_('{HASH}'); recordFirstAccess_(t); recordFirstAccess_(findTargetByHash_('{HASH}')); return true; }})()"
    out = run_gas(wb, scenario, now="2027-02-02T08:00:00+09:00")
    assert out["error"] is None
    target = rows_by_header(out["sheets"][S.SHEET_TARGETS], S.TARGET_HEADERS)[0]
    assert target["初回アクセス日時"] == "2027-02-02T08:00:00"
    assert [r[1] for r in out["sheets"][S.SHEET_AUDIT][1:]] == ["FIRST_ACCESS"]


# --- 案内メール ------------------------------------------------------------------------------

def test_send_invitations_generates_hashes_and_updates_rows():
    wb = workbook(targets=[
        target_row(社員番号="2099001", 社用メール="a@nmht.co.jp"),
        target_row(社員番号="2099002", 社用メール="b@nmht.co.jp", 氏名="二 号"),
        sent_target(社員番号="2099003", 社用メール="c@nmht.co.jp"),
    ])
    out = run_gas(wb, "sendInvitations()", prompt_answer="SEND 2027 2", now="2027-02-01T09:00:00+09:00")
    assert out["error"] is None
    assert [m["to"] for m in out["mails"]] == ["a@nmht.co.jp", "b@nmht.co.jp"]
    assert "二 号 さん" in out["mails"][1]["body"]
    url = out["mails"][0]["body"].split("https://script.google.com/macros/s/test/exec?token=")[1].split()[0]
    targets = rows_by_header(out["sheets"][S.SHEET_TARGETS], S.TARGET_HEADERS)
    assert targets[0]["トークンハッシュ"] == hashlib.sha256(url.encode("utf-8")).hexdigest()
    assert (targets[0]["送信日時"], targets[0]["送信回数"], targets[0]["申込状態"]) == ("2027-02-01T09:00:00", "1", S.STATUS_SENT)
    assert targets[2]["トークンハッシュ"] == HASH                     # 送信済みは触らない
    assert [r[1] for r in out["sheets"][S.SHEET_AUDIT][1:]] == ["SEND", "SEND"]
    assert any("送信 2件 / 失敗 0件" in a for a in out["alerts"])


def test_send_invitations_wrong_phrase_or_failure():
    wb = workbook(targets=[target_row(社員番号="2099001", 社用メール="a@nmht.co.jp")])
    out = run_gas(wb, "sendInvitations()", prompt_answer="SEND 2027 9")
    assert out["mails"] == [] and any("中止" in a for a in out["alerts"])
    assert rows_by_header(out["sheets"][S.SHEET_TARGETS], S.TARGET_HEADERS)[0]["送信日時"] == ""

    wb = workbook(targets=[target_row(社員番号="2099001", 社用メール="a@nmht.co.jp"),
                           target_row(社員番号="2099002", 社用メール="bad@nmht.co.jp")])
    out = run_gas(wb, "sendInvitations()", prompt_answer="SEND 2027 2", mail_fail="bad@nmht.co.jp")
    targets = rows_by_header(out["sheets"][S.SHEET_TARGETS], S.TARGET_HEADERS)
    assert targets[0]["送信日時"] != "" and targets[1]["送信日時"] == ""      # 失敗行は未送信のまま（再実行で再送）
    assert [r[1] for r in out["sheets"][S.SHEET_AUDIT][1:]] == ["SEND", "SEND_FAILED"]
    assert any("送信 1件 / 失敗 1件" in a for a in out["alerts"])


def test_build_invitation_email_and_targets_to_send():
    wb = workbook(settings=settings_rows(案内メール件名="{年度}年度 {氏名}", 案内メール本文="{URL} まで {受付終了}"))
    out = run_gas(wb, "(() => { const kv = readSettings_(); return buildInvitationEmail_(kv, {'氏名': '試験 太郎', '社用メール': 'a@nmht.co.jp'}, 'https://x/exec?token=T'); })()")
    assert out["result"] == {"to": "a@nmht.co.jp", "subject": "2027年度 試験 太郎", "body": "https://x/exec?token=T まで 2027-02-28"}
    out = run_gas(wb, "targetsToSend_([{'送信日時':'','申込状態':'','社用メール':'a'}, {'送信日時':'x','申込状態':'送信済','社用メール':'b'}, {'送信日時':'','申込状態':'無効','社用メール':'c'}, {'送信日時':'','申込状態':'未送信','社用メール':''}]).map(t => t['社用メール'])")
    assert out["result"] == ["a"]


def test_setup_workbook_creates_sheets_with_headers():
    out = run_gas({}, "(() => { setupWorkbook(); return Object.keys(SHEETS).map(k => SHEETS[k]); })()")
    assert out["error"] is None
    for name, headers in S.HEADERS_BY_SHEET.items():
        assert out["sheets"][name][0] == list(headers), name
    settings = rows_by_header(out["sheets"][S.SHEET_SETTINGS], S.SETTINGS_HEADERS)
    assert settings[0]["キー"] == "スキーマ版" and settings[0]["値"] == S.SCHEMA_VERSION
    assert {s["キー"] for s in settings} >= set(S.REQUIRED_SETTING_KEYS)
    assert out["sheets"][S.SHEET_AUDIT][-1][1] == "SETUP"
