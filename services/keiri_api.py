"""経理モード共通: jinjer 給与明細（salary-statements）とカスタム項目の取得・時点解決。

Z:\\API連携\\scripts\\keiri_api.py から移植（2026-07-27）。本体はこちらが正で、
Z:\\API連携 側にはマッピング表の作成・差分突合などの開発ツールだけを残す。
移植にあたり jinjer クライアントの土台を JinjerLaborContractClient から
本アプリの JinjerClient へ差し替えた（GET しか使わないので `_get` 1本で足りる）。
"""

from __future__ import annotations

import datetime
import json
import os
import re
import time as _time
import unicodedata

import requests

from services.jinjer_api_client import JinjerAPIError, JinjerClient

STATEMENTS_PATH = "/v1/employees/salary-statements"
CUSTOM_ITEMS_PATH = "/v1/employees/custom-items"

# payroll_info 配下の項目配列（各要素 {id,label,value}）
ARRAY_TYPES = (
    "salary_items",
    "salary_deduction_items",
    "salary_attendance_items",
    "salary_payment_items",
    "salary_other_items",
    "salary_fixed_tax_abatement_items",
)

# 給与計算関連カスタム項目（menu13）
PAYROLL_MENU_ID = 13
JINKENHI_KUBUN_NAMES = {"1": "本社", "2": "本社以外", "3": "育成"}
JINKENHI_KUBUN_DEFAULT = "本社以外"  # Phase B 確定: 未設定は本社以外


class KeiriJinjerClient(JinjerClient):
    """給与明細・カスタム項目取得用クライアント（GET のみ・jinjer への書き込みなし）。"""

    MAX_RETRIES = 3

    def _get(self, path: str, params: dict) -> list | dict:
        """GET して data を返す。401 は再認証、429 と接続断はリトライする。"""
        url = f"{self.base_url}{path}"
        refreshed = False
        for attempt in range(1, self.MAX_RETRIES + 1):
            try:
                response = requests.get(
                    url, headers=self._auth_headers(), params=params, timeout=self.timeout
                )
            except requests.RequestException as exc:
                if attempt >= self.MAX_RETRIES:
                    raise JinjerAPIError(f"jinjer API へ接続できませんでした: {exc}") from exc
                _time.sleep(min(2 ** (attempt - 1), 5))
                continue
            if response.status_code == 401 and not refreshed:
                self._access_token = None      # 次の _auth_headers で取り直す
                refreshed = True
                continue
            if response.status_code == 429 and attempt < self.MAX_RETRIES:
                try:
                    wait = max(float(response.headers.get("Retry-After", "")), 1.0)
                except (TypeError, ValueError):
                    wait = 65.0
                _time.sleep(wait)
                continue
            if response.status_code != 200:
                raise JinjerAPIError(
                    f"jinjer API エラー HTTP {response.status_code}: {path}"
                )
            try:
                return response.json().get("data", {})
            except ValueError as exc:
                raise JinjerAPIError("jinjer API から JSON 以外の応答が返されました。") from exc
        raise JinjerAPIError("jinjer API の再試行回数を超えました。")

    def get_salary_statements(self, executed_on: str, employee_ids=None) -> list:
        """指定月（yyyy-MM）の給与計算結果を全ページ取得して data[] を返す。

        実測（2026-07-24）: このエンドポイントは 10件/ページ（X-Item-Counts=総人数）。
        ページサイズ既定の 100 件と異なるため「空ページが返るまで」ループする。
        """
        params = {"executed-on": executed_on}
        if employee_ids:
            params["employee-ids"] = ",".join(str(v) for v in employee_ids)
        results: list = []
        page = 1
        while page <= 500:  # 安全上限
            page_params = dict(params, page=page)
            data = self._get(STATEMENTS_PATH, page_params)
            if not isinstance(data, list) or not data:
                break
            results.extend(data)
            page += 1
            _time.sleep(0.25)
        return results

    def get_custom_items(self, employee_ids, chunk: int = 50) -> dict:
        """custom-items を 50 件ずつ取得。{社員番号: person} を返す。"""
        result: dict = {}
        ids = [str(i) for i in employee_ids]
        for start in range(0, len(ids), chunk):
            data = self._get(
                CUSTOM_ITEMS_PATH,
                {"employee-ids": ",".join(ids[start: start + chunk])},
            )
            for person in data if isinstance(data, list) else []:
                result[str(person.get("employee_id"))] = person
            _time.sleep(0.3)
        return result

    def get_custom_menus(self) -> list:
        """カスタム項目のマスタ（メニュー→項目→選択肢）。選択肢IDを名前に戻すのに使う。

        健診申込モードが「健康診断履歴」（menu 15）の健診内容IDを名前へ戻すために追加
        （2026-09-02）。GET のみ。
        """
        data = self._get("/v1/master/custom-menus", {})
        return data if isinstance(data, list) else []


def get_client() -> KeiriJinjerClient:
    """認証済みクライアントを返す。"""
    client = KeiriJinjerClient()
    client.authenticate()
    return client


def now_iso() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def load_or_fetch_roster(client, cache_path, refresh: bool = False) -> list:
    """従業員一覧（在籍区分フィルタなし＝退職者含む全件）を取得して JSON キャッシュする。"""
    if not refresh and cache_path and os.path.exists(cache_path):
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)["employees"]
    employees = (client or get_client()).get_employees(only_active=False)
    if cache_path:
        os.makedirs(os.path.dirname(cache_path) or ".", exist_ok=True)
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump({"fetched_at": now_iso(), "employees": employees}, f, ensure_ascii=False)
    return employees


def _dig(record, *paths):
    """'company.joined_on' 形式のパス候補を順に試し、最初に見つかった非空値を返す。"""
    for path in paths:
        cur = record
        ok = True
        for key in path.split("."):
            if isinstance(cur, dict) and key in cur:
                cur = cur[key]
            else:
                ok = False
                break
        if ok and cur not in (None, ""):
            return cur
    return ""


def _as_text(value) -> str:
    if isinstance(value, dict):
        return str(value.get("name") or value.get("id") or "")
    return str(value or "")


def roster_index(employees) -> dict:
    """従業員一覧 → {社員番号: {name, enrollment, joined_on, retired_on}}。"""
    index = {}
    for emp in employees:
        emp_id = str(emp.get("id", "")).strip()
        if not emp_id:
            continue
        last = _as_text(_dig(emp, "company.last_name", "last_name"))
        first = _as_text(_dig(emp, "company.first_name", "first_name"))
        index[emp_id] = {
            "name": f"{last} {first}".strip(),
            "enrollment": _as_text(
                _dig(emp, "enrollment_classification", "company.enrollment_classification")
            ),
            "joined_on": _as_text(_dig(emp, "company.joined_on", "joined_on")),
            "retired_on": _as_text(
                _dig(
                    emp,
                    "company.retirement_date",  # 実測 2026-07-24: 退職日はこのキー
                    "company.resigned_on",
                    "company.retired_on",
                    "resigned_on",
                    "retired_on",
                )
            ),
        }
    return index


def classify_employee(emp_code) -> str:
    """社員番号の判別: 20YY始まり=自社（対象）、5/6/9始まり=派遣・テスト（対象外）。"""
    code = str(emp_code or "").strip()
    if re.match(r"^20\d{2}", code):
        return "target"
    if re.match(r"^[569]", code):
        return "excluded"
    return "other"


def normalize_label(s) -> str:
    """項目名照合用の正規化: NFKC（全角括弧→半角等）＋空白除去。"""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(s or "")))


def to_number(value):
    """API 値（文字列/数値）を float へ。数値でなければ None（'160:00' 等）。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    t = str(value).replace(",", "").strip()
    if not t:
        return None
    try:
        return float(t)
    except ValueError:
        return None


def normalize_ymd(s) -> str:
    """'2026/7/1'・'2026-03-10' 等を 'YYYY-MM-DD' へ正規化（不明は空文字）。"""
    m = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", str(s or "").strip())
    if not m:
        return ""
    return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"


def parse_payroll_custom_history(person, menu_id: int = PAYROLL_MENU_ID) -> list:
    """custom-items レスポンス 1 名分 → 給与計算関連の履歴行リスト（日付昇順）。

    customize_item[] の `id` は履歴行番号（同じ id の要素群 = 1 行）。
    戻り値: [{"date": "YYYY-MM-DD", "row": {title: value}}]
    """
    rows: dict = {}
    for menu in (person or {}).get("customize_menu", []) or []:
        if str(menu.get("id")) != str(menu_id):
            continue
        for it in menu.get("customize_item", []) or []:
            rows.setdefault(str(it.get("id")), {})[str(it.get("title"))] = it.get("value")
    out = [{"date": normalize_ymd(row.get("日付")), "row": row} for row in rows.values()]
    out.sort(key=lambda x: x["date"])
    return out


def resolve_custom_value(history, title, on_date, fallback_earliest: bool = True) -> str:
    """on_date('YYYY-MM-DD') 時点の title の値。日付<=on_date の最新非空値
    （空値レコードはさらに過去へ遡る＝affiliations の空 attendance_group と同じ扱い）。

    fallback_earliest: 対象日より前にレコードが無い場合、最も古い非空値を使う。
    人件費区分のように後から一括投入した項目をそれ以前の月に適用するために必要。
    """
    best = ""
    for h in history:
        if h["date"] and h["date"] <= on_date:
            v = str(h["row"].get(title) or "").strip()
            if v:
                best = v
    if best or not fallback_earliest:
        return best
    for h in history:  # history は日付昇順
        v = str(h["row"].get(title) or "").strip()
        if v:
            return v
    return ""


def midmonth_records(history, ym) -> list:
    """ym('YYYY-MM') 内の 2 日〜末日の履歴行＝月中異動の分割スイッチ（ルール9検知用）。"""
    return [h for h in history
            if h["date"] and h["date"][:7] == ym and h["date"][8:10] != "01"]


def statement_flag(person, st, key):
    """paid_on / is_payroll_closed / is_calculated を取得する。

    実測（2026-07-24）: これらは payroll_info 直下にある。念のため
    statement 直下・従業員直下もフォールバックで見る。
    """
    if isinstance(st, dict):
        pi = st.get("payroll_info")
        if isinstance(pi, dict) and key in pi:
            return pi.get(key)
        if key in st:
            return st.get(key)
    if isinstance(person, dict):
        return person.get(key)
    return None
