import json
from pathlib import Path
import subprocess

from prototypes.health_check_application.preview_server import (
    ApplicationStore,
    CLINIC_OPTIONS,
    HEALTH_OPTIONS,
    create_app,
)


def make_client():
    store = ApplicationStore()
    app = create_app(store)
    app.config.update(TESTING=True)
    return app.test_client(), store


def valid_payload(**overrides):
    payload = {
        "application_type": "same",
        "remarks": "テスト回答",
        "dependent_requested": False,
        "agreement": True,
    }
    payload.update(overrides)
    return payload


def standard_change_payload(**overrides):
    payload = valid_payload(
        application_type="change",
        clinic_choice="Myメディカルクリニック（大手町）",
        health_options=["基本健診"],
    )
    payload.update(overrides)
    return payload


def test_apps_script_separates_first_access_and_answered_at_columns():
    source = Path(
        "prototypes/health_check_application/apps_script/Code.gs"
    ).read_text(encoding="utf-8")
    assert "'回答日時', '初回アクセス日時'" in source
    assert "firstAccessAt: 10" in source
    assert "recordFirstAccess_(employee);" in source
    assert "appendAudit_('FIRST_ACCESS'" in source
    assert "targetSheet.getRange(employee.rowNumber, TARGET_COLUMNS.status, 1, 3)" in source


def test_apps_script_records_first_access_only_once():
    source_path = Path(
        "prototypes/health_check_application/apps_script/Code.gs"
    ).resolve()
    harness = f"""
const fs = require('fs');
const source = fs.readFileSync({json.dumps(str(source_path))}, 'utf8');
const firstAccessCell = {{
  value: '',
  writes: 0,
  getValue() {{ return this.value; }},
  setValue(value) {{ this.value = value; this.writes += 1; }},
}};
const auditRows = [];
const targetSheet = {{
  getRange(row, column) {{
    if (row !== 2 || column !== 10) throw new Error('unexpected target cell');
    return firstAccessCell;
  }},
}};
const auditSheet = {{ appendRow(row) {{ auditRows.push(row); }} }};
const spreadsheet = {{
  getSheetByName(name) {{
    if (name === '対象者') return targetSheet;
    if (name === '監査ログ') return auditSheet;
    return null;
  }},
}};
global.LockService = {{
  getScriptLock() {{ return {{ waitLock() {{}}, releaseLock() {{}} }}; }},
}};
global.SpreadsheetApp = {{ openById() {{ return spreadsheet; }} }};
eval(source);
recordFirstAccess_({{ rowNumber: 2, employeeId: 'TEST-001' }});
recordFirstAccess_({{ rowNumber: 2, employeeId: 'TEST-001' }});
if (firstAccessCell.writes !== 1) throw new Error('first access was overwritten');
if (auditRows.length !== 1 || auditRows[0][1] !== 'FIRST_ACCESS') {{
  throw new Error('unexpected audit rows');
}}
"""
    subprocess.run(
        ["node", "-e", harness],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )


def test_individual_link_has_locked_identity_and_new_fields():
    client, _ = make_client()
    response = client.get("/apply/demo-aoki-2027")
    html = response.get_data(as_text=True)
    assert response.status_code == 200
    assert "TEST-001" in html
    assert "青木 花子" in html
    assert 'name="preferred_date"' not in html
    assert 'name="other_planned_date"' in html
    assert 'name="dependent_requested"' in html
    assert 'name="dependent_relationship"' in html
    assert 'name="dependent_name"' in html
    for clinic in CLINIC_OPTIONS:
        assert clinic in html
    for option in HEALTH_OPTIONS:
        assert option in html
    assert html.index('id="change-fields"') < html.index("</form>")
    assert html.index('id="dependent-fields"') < html.index("</form>")


def test_invalid_link_is_rejected():
    client, _ = make_client()
    response = client.get("/apply/not-a-valid-token")
    assert response.status_code == 404


def test_same_submission_ignores_forged_employee_identity():
    client, store = make_client()
    payload = valid_payload(employee_id="REAL-999", name_ja="別人")
    response = client.post("/api/applications/demo-aoki-2027", json=payload)
    assert response.status_code == 200
    assert response.get_json()["employee_id"] == "TEST-001"
    saved = store.get("demo-aoki-2027")
    assert saved["employee_id"] == "TEST-001"
    assert saved["name_ja"] == "青木 花子"
    assert saved["clinic"] == "青葉健診センター"
    assert saved["other_planned_date"] == ""


def test_duplicate_submission_is_blocked():
    client, _ = make_client()
    first = client.post("/api/applications/demo-aoki-2027", json=valid_payload())
    second = client.post("/api/applications/demo-aoki-2027", json=valid_payload())
    assert first.status_code == 200
    assert second.status_code == 409
    assert second.get_json()["duplicate"] is True


def test_admin_view_lists_submission_without_general_preferred_date():
    client, _ = make_client()
    response = client.post("/api/applications/demo-aoki-2027", json=valid_payload())
    assert response.status_code == 200
    admin = client.get("/admin")
    html = admin.get_data(as_text=True)
    assert admin.status_code == 200
    assert "TEST-2027-001" in html
    assert "TEST-001" in html
    assert "青木 花子" in html
    assert "その他の健診予定日" in html
    assert "受診希望日" not in html


def test_change_requires_a_listed_clinic():
    client, _ = make_client()
    response = client.post(
        "/api/applications/demo-sato-2027",
        json=standard_change_payload(clinic_choice="一覧にない健診機関"),
    )
    assert response.status_code == 400


def test_listed_clinic_requires_at_least_one_option():
    client, _ = make_client()
    response = client.post(
        "/api/applications/demo-sato-2027",
        json=standard_change_payload(health_options=[]),
    )
    assert response.status_code == 400


def test_listed_clinic_and_multiple_options_are_saved():
    client, store = make_client()
    response = client.post(
        "/api/applications/demo-sato-2027",
        json=standard_change_payload(
            clinic_choice="医療法人社団同友会 春日クリニック",
            health_options=["基本健診", "婦人科検診"],
        ),
    )
    assert response.status_code == 200
    saved = store.get("demo-sato-2027")
    assert saved["clinic"] == "医療法人社団同友会 春日クリニック"
    assert saved["course"] == "基本健診、婦人科検診"


def test_unlisted_health_option_is_rejected():
    client, _ = make_client()
    response = client.post(
        "/api/applications/demo-sato-2027",
        json=standard_change_payload(health_options=["未登録オプション"]),
    )
    assert response.status_code == 400


def test_other_clinic_accepts_blank_optional_details():
    client, store = make_client()
    response = client.post(
        "/api/applications/demo-lee-2027",
        json=valid_payload(application_type="change", clinic_choice="その他"),
    )
    assert response.status_code == 200
    saved = store.get("demo-lee-2027")
    assert saved["clinic"] == "その他（医療機関未定）"
    assert saved["course"] == "その他"
    assert saved["other_planned_date"] == ""


def test_other_clinic_saves_optional_clinic_and_date():
    client, store = make_client()
    response = client.post(
        "/api/applications/demo-lee-2027",
        json=valid_payload(
            application_type="change",
            clinic_choice="その他",
            custom_clinic="南町健診センター",
            other_planned_date="2027-09-10",
        ),
    )
    assert response.status_code == 200
    saved = store.get("demo-lee-2027")
    assert saved["clinic"] == "南町健診センター"
    assert saved["other_planned_date"] == "2027-09-10"


def test_other_clinic_date_must_be_within_fiscal_period():
    client, _ = make_client()
    response = client.post(
        "/api/applications/demo-lee-2027",
        json=valid_payload(
            application_type="change",
            clinic_choice="その他",
            other_planned_date="2028-04-01",
        ),
    )
    assert response.status_code == 400


def test_dependent_requires_relationship_and_name():
    client, _ = make_client()
    response = client.post(
        "/api/applications/demo-aoki-2027",
        json=valid_payload(dependent_requested=True),
    )
    assert response.status_code == 400


def test_dependent_relationship_and_name_are_saved():
    client, store = make_client()
    response = client.post(
        "/api/applications/demo-aoki-2027",
        json=valid_payload(
            dependent_requested=True,
            dependent_relationship="夫",
            dependent_name="青木 一郎",
        ),
    )
    assert response.status_code == 200
    saved = store.get("demo-aoki-2027")
    assert saved["dependent_requested"] is True
    assert saved["dependent_relationship"] == "夫"
    assert saved["dependent_name"] == "青木 一郎"

