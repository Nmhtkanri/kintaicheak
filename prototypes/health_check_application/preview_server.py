"""Local-only health-check application prototype.

Uses fictitious employees, stores answers only in memory, and has no email or
jinjer integration. The server selects employee identity from an individual
token; identity values sent by the browser are never trusted.
"""

from __future__ import annotations

import argparse
import threading
import webbrowser
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, redirect, render_template, request, url_for


FISCAL_YEAR = 2027
EXAM_DATE_MIN = date(2027, 4, 1)
EXAM_DATE_MAX = date(2028, 3, 31)
OTHER_CLINIC = "その他"
CLINIC_OPTIONS = (
    "Myメディカルクリニック（大手町）",
    "Myメディカルクリニック（渋谷）",
    "Myメディカルクリニック（新宿）",
    "Myメディカルクリニック（横浜）",
    "Myメディカルクリニック（田町）",
    "Myメディカルクリニック（八重洲）",
    "医療法人社団同友会 春日クリニック",
    OTHER_CLINIC,
)
HEALTH_OPTIONS = (
    "基本健診",
    "1日人間ドック",
    "1日人間ドック・胃カメラ",
    "1日人間ドック・バリウム",
    "婦人科検診",
)
DEPENDENT_RELATIONSHIPS = {"妻", "夫"}

TEST_EMPLOYEES: dict[str, dict[str, str]] = {
    "demo-aoki-2027": {
        "employee_id": "TEST-001",
        "name_ja": "青木 花子",
        "name_en": "Hanako Aoki",
        "email": "hanako.aoki@example.invalid",
        "previous_clinic": "青葉健診センター",
        "previous_course": "定期健康診断（一般）",
    },
    "demo-sato-2027": {
        "employee_id": "TEST-002",
        "name_ja": "佐藤 健太",
        "name_en": "Kenta Sato",
        "email": "kenta.sato@example.invalid",
        "previous_clinic": "中央メディカルクリニック",
        "previous_course": "定期健康診断（一般）",
    },
    "demo-lee-2027": {
        "employee_id": "TEST-003",
        "name_ja": "李 美玲",
        "name_en": "Meiling Li",
        "email": "meiling.li@example.invalid",
        "previous_clinic": "新宿ヘルスケアセンター",
        "previous_course": "生活習慣病健診",
    },
}


class ApplicationStore:
    """Thread-safe, in-memory test store."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._submissions: dict[str, dict[str, Any]] = {}
        self._sequence = 0

    def submit(self, token: str, record: dict[str, Any]) -> dict[str, Any] | None:
        with self._lock:
            if token in self._submissions:
                return None
            self._sequence += 1
            stored = deepcopy(record)
            stored["receipt_id"] = f"TEST-{FISCAL_YEAR}-{self._sequence:03d}"
            self._submissions[token] = stored
            return deepcopy(stored)

    def get(self, token: str) -> dict[str, Any] | None:
        with self._lock:
            item = self._submissions.get(token)
            return deepcopy(item) if item else None

    def all(self) -> list[dict[str, Any]]:
        with self._lock:
            return [deepcopy(item) for item in self._submissions.values()]

    def reset(self) -> None:
        with self._lock:
            self._submissions.clear()
            self._sequence = 0


def _parse_exam_date(raw_value: Any) -> date | None:
    if not isinstance(raw_value, str) or not raw_value:
        return None
    try:
        parsed = date.fromisoformat(raw_value)
    except ValueError:
        return None
    if not EXAM_DATE_MIN <= parsed <= EXAM_DATE_MAX:
        return None
    return parsed


def _clean_text(value: Any, max_length: int) -> str:
    if value is None:
        return ""
    return str(value).strip()[:max_length]


def create_app(store: ApplicationStore | None = None) -> Flask:
    root = Path(__file__).resolve().parent
    app = Flask(
        __name__,
        template_folder=str(root / "templates"),
        static_folder=str(root / "static"),
    )
    app.config.update(TESTING=False, JSON_AS_ASCII=False)
    app.extensions["health_check_store"] = store or ApplicationStore()

    def current_store() -> ApplicationStore:
        return app.extensions["health_check_store"]

    @app.get("/")
    def test_menu():
        people = []
        for token, employee in TEST_EMPLOYEES.items():
            people.append(
                {
                    **employee,
                    "token": token,
                    "submitted": current_store().get(token) is not None,
                }
            )
        return render_template(
            "test_menu.html",
            fiscal_year=FISCAL_YEAR,
            people=people,
        )

    @app.get("/apply/<token>")
    def application_form(token: str):
        employee = TEST_EMPLOYEES.get(token)
        if employee is None:
            return render_template("invalid_link.html"), 404
        existing = current_store().get(token)
        return render_template(
            "application.html",
            fiscal_year=FISCAL_YEAR,
            employee=employee,
            token=token,
            existing=existing,
            exam_date_min=EXAM_DATE_MIN.isoformat(),
            exam_date_max=EXAM_DATE_MAX.isoformat(),
            clinic_options=CLINIC_OPTIONS,
            health_options=HEALTH_OPTIONS,
            other_clinic=OTHER_CLINIC,
        )

    @app.post("/api/applications/<token>")
    def submit_application(token: str):
        employee = TEST_EMPLOYEES.get(token)
        if employee is None:
            return jsonify(ok=False, message="申込URLが無効です。 / Invalid application link."), 404

        payload = request.get_json(silent=True) or {}
        application_type = _clean_text(payload.get("application_type"), 20)
        if application_type not in {"same", "change"}:
            return jsonify(ok=False, message="申込内容を選択してください。 / Select an application option."), 400
        if payload.get("agreement") is not True:
            return jsonify(ok=False, message="確認欄にチェックしてください。 / Please confirm the details."), 400

        clinic_choice = ""
        selected_options: list[str] = []
        other_planned_date = ""
        if application_type == "same":
            clinic = employee["previous_clinic"]
            course = employee["previous_course"]
        else:
            clinic_choice = _clean_text(payload.get("clinic_choice"), 100)
            if clinic_choice not in CLINIC_OPTIONS:
                return jsonify(
                    ok=False,
                    message="健診機関を選択してください。 / Select a clinic.",
                ), 400
            if clinic_choice == OTHER_CLINIC:
                custom_clinic = _clean_text(payload.get("custom_clinic"), 100)
                clinic = custom_clinic or "その他（医療機関未定）"
                course = "その他"
                raw_date = _clean_text(payload.get("other_planned_date"), 10)
                if raw_date:
                    parsed_date = _parse_exam_date(raw_date)
                    if parsed_date is None:
                        return jsonify(
                            ok=False,
                            message="健診予定日は2027年4月1日から2028年3月31日の間で入力してください。 / Enter a date within the examination period.",
                        ), 400
                    other_planned_date = parsed_date.isoformat()
            else:
                raw_options = payload.get("health_options")
                if not isinstance(raw_options, list):
                    raw_options = []
                requested_options = {_clean_text(item, 100) for item in raw_options}
                if requested_options - set(HEALTH_OPTIONS):
                    return jsonify(
                        ok=False,
                        message="選択できない健診オプションが含まれています。 / An invalid health-check option was selected.",
                    ), 400
                selected_options = [item for item in HEALTH_OPTIONS if item in requested_options]
                if not selected_options:
                    return jsonify(
                        ok=False,
                        message="健診オプションを1つ以上選択してください。 / Select at least one health-check option.",
                    ), 400
                clinic = clinic_choice
                course = "、".join(selected_options)

        dependent_requested = payload.get("dependent_requested") is True
        dependent_relationship = _clean_text(payload.get("dependent_relationship"), 10)
        dependent_name = _clean_text(payload.get("dependent_name"), 100)
        if dependent_requested:
            if dependent_relationship not in DEPENDENT_RELATIONSHIPS or not dependent_name:
                return jsonify(
                    ok=False,
                    message="被扶養者の続柄（妻・夫）と氏名を入力してください。 / Enter the dependent's relationship and name.",
                ), 400
        else:
            dependent_relationship = ""
            dependent_name = ""

        # Identity fields come only from the server-side token mapping. Values
        # such as employee_id sent by a modified browser are intentionally ignored.
        record = {
            "received_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "fiscal_year": FISCAL_YEAR,
            "employee_id": employee["employee_id"],
            "name_ja": employee["name_ja"],
            "name_en": employee["name_en"],
            "email": employee["email"],
            "application_type": application_type,
            "clinic_choice": clinic_choice,
            "clinic": clinic,
            "course": course,
            "health_options": "、".join(selected_options),
            "other_planned_date": other_planned_date,
            "dependent_requested": dependent_requested,
            "dependent_relationship": dependent_relationship,
            "dependent_name": dependent_name,
            "remarks": _clean_text(payload.get("remarks"), 500),
            "status": "受付済（テスト）",
        }
        stored = current_store().submit(token, record)
        if stored is None:
            existing = current_store().get(token)
            return jsonify(
                ok=False,
                duplicate=True,
                receipt_id=existing["receipt_id"] if existing else "",
                message="このURLでは既に申込済みです。 / This link has already been used.",
            ), 409
        return jsonify(
            ok=True,
            receipt_id=stored["receipt_id"],
            employee_id=stored["employee_id"],
            message="テスト申込を受け付けました。 / Your test application has been received.",
        )

    @app.get("/admin")
    def admin_view():
        return render_template(
            "admin.html",
            fiscal_year=FISCAL_YEAR,
            submissions=current_store().all(),
            total_people=len(TEST_EMPLOYEES),
        )

    @app.post("/admin/reset")
    def reset_test_data():
        current_store().reset()
        return redirect(url_for("admin_view"))

    return app


def main() -> None:
    parser = argparse.ArgumentParser(description="健康診断申込のローカル試用版")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    url = f"http://127.0.0.1:{args.port}/"
    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print("健康診断申込テスト版を起動しました。")
    print(f"URL: {url}")
    print("終了するときは、この画面で Ctrl+C を押してください。")
    create_app().run(host="127.0.0.1", port=args.port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
