"""jinjer API クライアント（従業員ID取得用）

CSV変換モードで「氏名 → 従業員ID」のマップを構築するために使用する。
既存プロジェクト Z:\\API連携\\jinjer_client.py をベースに、kintai-checker 用に最小化。

主な機能:
- アクセストークン取得（4時間有効、簡易キャッシュ）
- /v1/employees から在籍者一覧を取得
- 氏名 → ID マップ生成（姓名連結のバリエーションを複数登録）
"""

from __future__ import annotations

import logging
import re
import time as _time
from typing import Any

import requests

from config import Config

logger = logging.getLogger(__name__)


class JinjerAPIError(Exception):
    """jinjer API 関連の例外"""


def _strip_seconds(v) -> str:
    """'09:15:00' → '9:15'。24時超（'33:30:00'）もそのまま扱う。不正値は空文字。"""
    s = str(v or "").strip()
    m = re.match(r"^(\d{1,3}):(\d{2})(?::\d{2})?$", s)
    if not m:
        return ""
    return f"{int(m.group(1))}:{m.group(2)}"


def parse_work_schedules_data(data: dict) -> dict[str, dict]:
    """work-schedules レスポンスの data 部を日別 dict に変換する（純粋関数）。

    jinjer は同一日のスケジュールを新旧2バージョンで二重返却することがある
    （実測: 先頭が新しい方＝画面表示と一致）。**先頭レコードを採用**し、
    2件目以降の同一日は捨てる。
    """
    result: dict[str, dict] = {}
    for w in (data or {}).get("work_schedules", []) or []:
        d = str(w.get("date") or "").strip()
        if not d or d in result:
            continue  # 同一日の2件目以降は旧バージョン → 先頭採用
        sched = w.get("work_schedule") or {}
        breaks = []
        for b in w.get("break_schedules", []) or []:
            bs = _strip_seconds(b.get("start"))
            be = _strip_seconds(b.get("end"))
            if bs and be:
                breaks.append((bs, be))
        result[d] = {
            "start": _strip_seconds(sched.get("start")),
            "end": _strip_seconds(sched.get("end")),
            "breaks": breaks,
            "store": str((w.get("store") or {}).get("name") or ""),
        }
    return result


def parse_requested_day_offs_data(data: dict) -> dict[str, str]:
    """requested-day-offs レスポンスの data 部を {date_iso: 説明文} に変換する（純粋関数）。

    date は "2026/07/06" / "2026-7-6" のどちらでも ISO (YYYY-MM-DD) に正規化する。
    説明文は「休暇名/status=N」形式（status の意味は jinjer 側仕様。表示用に残す）。
    """
    result: dict[str, str] = {}
    for rec in (data or {}).get("requested_day_offs", []) or []:
        d_raw = str(rec.get("date") or "").replace("/", "-")
        m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", d_raw)
        if not m:
            continue
        d_iso = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        name = ((rec.get("day_off_classification") or {}).get("name") or "休暇")
        result[d_iso] = f"{name}/status={rec.get('status')}"
    return result


class JinjerClient:
    """jinjer API への薄いクライアント。

    トークンはインスタンス内にキャッシュする（4 時間有効）。
    """

    TOKEN_TTL_SEC = 60 * 60 * 4  # 4時間

    def __init__(
        self,
        api_key: str | None = None,
        secret_key: str | None = None,
        base_url: str | None = None,
        timeout: int = 30,
    ):
        self.api_key = api_key or Config.JINJER_API_KEY
        self.secret_key = secret_key or Config.JINJER_SECRET_KEY
        self.base_url = (base_url or Config.JINJER_BASE_URL).rstrip("/")
        self.timeout = timeout
        self._access_token: str | None = None
        self._token_acquired_at: float = 0.0

    # ------------------------------------------------------------------
    # 認証
    # ------------------------------------------------------------------
    def authenticate(self) -> str:
        """アクセストークンを取得（既にキャッシュがあれば再利用）"""
        if self._access_token and (_time.time() - self._token_acquired_at) < self.TOKEN_TTL_SEC:
            return self._access_token

        if not self.api_key or not self.secret_key:
            raise JinjerAPIError(
                "JINJER_API_KEY / JINJER_SECRET_KEY が .env に設定されていません"
            )

        url = f"{self.base_url}/v2/token"
        headers = {
            "Content-Type": "application/json",
            "X-API-KEY": self.api_key,
            "X-SECRET-KEY": self.secret_key,
        }

        try:
            response = requests.get(url, headers=headers, timeout=self.timeout)
        except requests.RequestException as e:
            raise JinjerAPIError(f"jinjer API へ接続できませんでした: {e}") from e

        if response.status_code != 200:
            raise JinjerAPIError(
                f"jinjer 認証に失敗しました (status={response.status_code})"
            )

        try:
            data = response.json()
            token = data["data"]["access_token"]
        except (KeyError, ValueError) as e:
            raise JinjerAPIError(f"jinjer 認証レスポンスが不正です: {e}") from e

        self._access_token = token
        self._token_acquired_at = _time.time()
        logger.info("jinjer API 認証成功")
        return token

    def _auth_headers(self) -> dict:
        token = self.authenticate()
        return {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}",
        }

    # ------------------------------------------------------------------
    # 所属履歴（打刻グループ判定用）
    # ------------------------------------------------------------------
    AFFILIATIONS_BATCH_SIZE = 50  # employee-ids クエリの実用上限

    def get_affiliations(self, employee_ids: list[str]) -> dict[str, list[dict]]:
        """`/v1/employees/affiliations` を 50件ずつバッチで叩く

        Args:
            employee_ids: 従業員ID（文字列）のリスト

        Returns:
            {employee_id(str): [affiliation, ...]} の辞書。
            各 affiliation には ``date_of_issue``, ``attendance_group: {id, name}`` 等が含まれる。
            **chronological order とは限らない** ので呼び出し側で `date_of_issue` でソートすること。
        """
        url = f"{self.base_url}/v1/employees/affiliations"
        headers = self._auth_headers()
        result: dict[str, list[dict]] = {}

        # 重複排除＆空文字除外
        unique_ids = []
        seen = set()
        for emp_id in employee_ids:
            s = str(emp_id or "").strip()
            if not s or s in seen:
                continue
            seen.add(s)
            unique_ids.append(s)

        if not unique_ids:
            return result

        for i in range(0, len(unique_ids), self.AFFILIATIONS_BATCH_SIZE):
            chunk = unique_ids[i : i + self.AFFILIATIONS_BATCH_SIZE]
            params = {"employee-ids": ",".join(chunk)}
            try:
                response = requests.get(
                    url, headers=headers, params=params, timeout=self.timeout
                )
            except requests.RequestException as e:
                raise JinjerAPIError(f"所属履歴取得に失敗: {e}") from e

            if response.status_code != 200:
                raise JinjerAPIError(
                    f"所属履歴取得に失敗 (status={response.status_code})"
                )

            data = response.json().get("data", []) or []
            for item in data:
                if not isinstance(item, dict):
                    continue
                emp_id = str(item.get("employee_id") or "").strip()
                if not emp_id:
                    continue
                result[emp_id] = item.get("affiliations", []) or []

            if i + self.AFFILIATIONS_BATCH_SIZE < len(unique_ids):
                _time.sleep(0.3)  # 軽いペーシング

        logger.info("jinjer 所属履歴取得: %d 件", len(result))
        return result

    # ------------------------------------------------------------------
    # 従業員一覧
    # ------------------------------------------------------------------
    def get_employees(self, only_active: bool = True) -> list[dict]:
        """従業員一覧を全ページ取得する

        Args:
            only_active: True の場合 enrollment_classification_id=0 (在籍者のみ)

        Returns:
            list of employee dict（jinjer のレスポンスそのまま）
        """
        url = f"{self.base_url}/v1/employees"
        headers = self._auth_headers()
        params: dict[str, Any] = {}
        if only_active:
            params["enrollment-classification-id"] = "0"

        all_employees: list[dict] = []
        page = 1
        while True:
            params["page"] = page
            try:
                response = requests.get(
                    url, headers=headers, params=params, timeout=self.timeout
                )
            except requests.RequestException as e:
                raise JinjerAPIError(f"従業員一覧取得に失敗: {e}") from e

            if response.status_code != 200:
                raise JinjerAPIError(
                    f"従業員一覧取得に失敗 (status={response.status_code})"
                )

            payload = response.json()
            employees = payload.get("data", []) or []
            if not employees:
                break

            all_employees.extend(employees)

            try:
                total_count = int(response.headers.get("X-Item-Counts", 0))
            except (TypeError, ValueError):
                total_count = 0

            if total_count and len(all_employees) >= total_count:
                break
            if len(employees) < 100:
                # 1ページ最大100件未満なら最終ページ
                break
            page += 1
            _time.sleep(0.1)  # 軽いペーシング

        logger.info("jinjer 従業員一覧取得: %d 件", len(all_employees))
        return all_employees

    # ------------------------------------------------------------------
    # 通勤情報（出発・到着・経由・経路・支給間隔・支給金額）
    # ------------------------------------------------------------------
    def get_commuting_information(self, employee_ids: list[str] | None = None) -> list[dict]:
        """`/v1/employees/commuting-information` を全ページ取得する。

        各要素: ``{"employee_id": str, "commuting": [route, ...]}``。
        route には ``departure / arrival / transit_1 / transit_2 / path / type / one_way_distance``
        と ``payment{ start_date, interval{name}, method{name}, total, tax_exemption_amount,
        taxable_amount }`` が含まれる（出発・到着・経由・通勤経路・支給間隔・支給金額が取れる）。

        Args:
            employee_ids: 指定時はその従業員のみ（カンマ区切り、未指定なら全員）。
                          ※クエリは ``employee-ids``（複数形・ハイフン）のみ許可。
                          ``employee-id``（単数）は 400 になる。
        """
        url = f"{self.base_url}/v1/employees/commuting-information"
        headers = self._auth_headers()
        base_params: dict[str, Any] = {}
        if employee_ids:
            base_params["employee-ids"] = ",".join(str(e).strip() for e in employee_ids if str(e).strip())

        all_items: list[dict] = []
        page = 1
        while True:
            params = dict(base_params, page=page)
            try:
                response = requests.get(url, headers=headers, params=params, timeout=self.timeout)
            except requests.RequestException as e:
                raise JinjerAPIError(f"通勤情報取得に失敗: {e}") from e
            if response.status_code != 200:
                raise JinjerAPIError(f"通勤情報取得に失敗 (status={response.status_code})")

            data = response.json().get("data", []) or []
            if not data:
                break
            all_items.extend(data)

            try:
                total_count = int(response.headers.get("X-Item-Counts", 0))
            except (TypeError, ValueError):
                total_count = 0
            if total_count and len(all_items) >= total_count:
                break
            if len(data) < 100:
                break
            page += 1
            _time.sleep(0.1)

        logger.info("jinjer 通勤情報取得: %d 件", len(all_items))
        return all_items

    # ------------------------------------------------------------------
    # 勤怠インポート（汎用データCSVのAPI投入）と日別スケジュール
    #   背景: 画面の汎用データインポートがスケジュール列をサイレントに
    #   反映しない不具合(2026-07-09)を、POST /v1/kintai-imports で回避できる
    #   ことを実証済み（docs/PLAN_手順3_API直接投入.md）。
    # ------------------------------------------------------------------
    def post_kintai_import(
        self, csv_bytes: bytes, file_name: str, executor_id: str | None = None
    ) -> dict:
        """汎用データCSVを `POST /v1/kintai-imports`（種別5=新規登録・編集）で投入予約する。

        Args:
            csv_bytes: Shift_JIS(CP932) でエンコード済みの汎用データCSV（5000行以内）
            file_name: インポートファイル名（70字以内。jinjer側で一意IDが付与される）
            executor_id: 実行者の社員番号（勤怠管理者権限が必要。完了通知メールの宛先。
                         未指定時はマスタアカウント扱い）

        Returns:
            レスポンス JSON の ``data``（``executor`` / ``type`` を含む）

        Raises:
            JinjerAPIError: 認証失敗・投入予約の失敗（同時予約は1件までの制限あり）
        """
        import base64 as _base64

        headers = dict(self._auth_headers())
        body: dict[str, Any] = {
            "type": {"id": "5"},
            "input_file": {
                "name": file_name,
                "encoded_string": _base64.b64encode(csv_bytes).decode("ascii"),
            },
        }
        if executor_id:
            body["executor"] = {"id": str(executor_id)}
        try:
            r = requests.post(
                f"{self.base_url}/v1/kintai-imports",
                headers=headers, json=body, timeout=120,
            )
        except requests.RequestException as e:
            raise JinjerAPIError(f"勤怠インポートの投入に失敗: {e}") from e
        if r.status_code != 200:
            raise JinjerAPIError(
                f"勤怠インポートの投入に失敗 (status={r.status_code}): {r.text[:300]}"
            )
        logger.info("kintai-imports 投入予約 OK: %s", file_name)
        return r.json().get("data") or {}

    def find_kintai_import(self, file_name: str) -> dict | None:
        """`GET /v1/kintai-imports` を全ページ走査し、ファイル名が一致する最新レコードを返す。

        jinjer は登録時にファイル名へ一意IDを付与するため、拡張子を除いた
        ベース名の部分一致で照合する。見つからなければ None。
        ``status``: "0"=予約 / "1"=成功 / "2"=失敗。
        """
        base_name = file_name.rsplit(".", 1)[0]
        headers = self._auth_headers()
        found: dict | None = None
        page = 1
        while page <= 100:
            try:
                r = requests.get(
                    f"{self.base_url}/v1/kintai-imports",
                    headers=headers, params={"page": page}, timeout=30,
                )
            except requests.RequestException as e:
                raise JinjerAPIError(f"勤怠インポート状況の取得に失敗: {e}") from e
            if r.status_code != 200:
                raise JinjerAPIError(
                    f"勤怠インポート状況の取得に失敗 (status={r.status_code})"
                )
            items = r.json().get("data", []) or []
            if not items:
                break
            for item in items:
                name = str((item.get("input_file") or {}).get("name") or "")
                if base_name in name:
                    found = item  # 後のページほど新しい想定で上書き
            page += 1
            _time.sleep(0.1)
        return found

    # ------------------------------------------------------------------
    # 人事インポート（給与計算インポートのAPI投入）
    #   POST /v1/jinji-imports はメニュー種別「給与計算」の入力テンプレートに対応しており、
    #   経費インポートCSV（テンプレート「経費APIインポート用」id=44450）をAPIで投入できる。
    #   給与計算テンプレートは type.id="3"（新規登録/更新）必須・salary_setting.executed_on=処理月。
    # ------------------------------------------------------------------
    def post_jinji_import(
        self,
        csv_bytes: bytes,
        file_name: str,
        template_id: str,
        executed_on: str,
        apply_formulas_off: bool = True,
        type_id: str = "3",
    ) -> dict:
        """人事インポート（給与計算等）を `POST /v1/jinji-imports` で投入予約する。

        Args:
            csv_bytes: Shift_JIS(CP932) でエンコード済みのインポートCSV
            file_name: インポートファイル名（255字以内）
            template_id: 入力テンプレートID（GET /v1/master/jinji-import-templates の id）
            executed_on: 処理月 "YYYY-MM"（給与計算/賞与計算/退職金計算のテンプレートで必須）
            apply_formulas_off: 「編集した項目は計算式を適用しない」を有効にするか
            type_id: インポート種別（給与計算テンプレートは "3"=新規登録/更新 必須）

        Returns:
            レスポンス JSON の ``data``

        Raises:
            JinjerAPIError: 認証失敗・投入予約の失敗
        """
        import base64 as _base64

        headers = dict(self._auth_headers())
        body: dict[str, Any] = {
            "type": {"id": str(type_id)},
            "template": {"id": str(template_id)},
            "salary_setting": {
                "executed_on": executed_on,
                "is_checked_not_to_apply_formulas": bool(apply_formulas_off),
            },
            "file": {
                "name": file_name,
                "encoded_string": _base64.b64encode(csv_bytes).decode("ascii"),
            },
        }
        try:
            r = requests.post(
                f"{self.base_url}/v1/jinji-imports",
                headers=headers, json=body, timeout=120,
            )
        except requests.RequestException as e:
            raise JinjerAPIError(f"人事インポートの投入に失敗: {e}") from e
        if r.status_code != 200:
            raise JinjerAPIError(
                f"人事インポートの投入に失敗 (status={r.status_code}): {r.text[:300]}"
            )
        logger.info("jinji-imports 投入予約 OK: %s (template=%s, month=%s)",
                    file_name, template_id, executed_on)
        return r.json().get("data") or {}

    def find_jinji_import(self, import_id: str) -> dict | None:
        """`GET /v1/jinji-imports` から指定 id のレコードを返す。見つからなければ None。

        kintai-imports と違い、jinji-imports の GET 応答には ``status`` やファイル名は無い。
        レコードは ``id`` / ``menu`` / ``template`` / ``number_of_total_rows`` /
        ``number_of_failed_rows`` / ``created_at`` / ``updated_at`` を持つ（2026-07-16 実測）。
        成否は ``number_of_failed_rows`` で判定する（0=全行成功）。
        """
        headers = self._auth_headers()
        target = str(import_id)
        page = 1
        while page <= 100:
            try:
                r = requests.get(
                    f"{self.base_url}/v1/jinji-imports",
                    headers=headers, params={"page": page}, timeout=30,
                )
            except requests.RequestException as e:
                raise JinjerAPIError(f"人事インポート状況の取得に失敗: {e}") from e
            if r.status_code != 200:
                raise JinjerAPIError(
                    f"人事インポート状況の取得に失敗 (status={r.status_code})"
                )
            items = r.json().get("data", []) or []
            if not items:
                break
            for item in items:
                if str(item.get("id") or "") == target:
                    return item
            page += 1
            _time.sleep(0.1)
        return None

    def get_work_schedules(self, employee_id: str, month: str) -> dict[str, dict]:
        """`/v1/employees/work-schedules` から日別スケジュールを取得する。

        Returns:
            ``{date_iso: {"start": "9:00", "end": "17:30", "breaks": [("12:00","13:00"), ...],
                          "store": 打刻グループ名}}``
            夜勤は 24時超表記（例 "33:30"）のまま返る。スケジュールがない日はキーごと無い。
            同一日の新旧二重返却は先頭（新しい方）を採用する。
        """
        headers = self._auth_headers()
        try:
            r = requests.get(
                f"{self.base_url}/v1/employees/work-schedules",
                headers=headers,
                params={"employee-id": str(employee_id), "month": month},
                timeout=30,
            )
        except requests.RequestException as e:
            raise JinjerAPIError(f"スケジュール取得に失敗 emp={employee_id}: {e}") from e
        if r.status_code != 200:
            logger.warning("work-schedules 取得失敗 emp=%s status=%s", employee_id, r.status_code)
            return {}
        return parse_work_schedules_data(r.json().get("data") or {})

    def get_requested_day_offs(self, employee_id: str, month: str) -> dict[str, str]:
        """`/v1/employees/requested-day-offs` から休暇登録日を取得する。

        休暇登録がある日はjinjerがスケジュール書込をサイレントに無視するため、
        投入前ガードに使う。取得失敗は警告のみで空dictを返す
        （失敗しても休暇日への書込はjinjerが無視 → 事後検証NGとして顕在化する）。

        Returns:
            ``{date_iso: "休暇名/status=N"}``
        """
        headers = self._auth_headers()
        r = None
        for attempt in range(3):
            try:
                r = requests.get(
                    f"{self.base_url}/v1/employees/requested-day-offs",
                    headers=headers,
                    params={"employee-id": str(employee_id), "month": month},
                    timeout=30,
                )
            except requests.RequestException as e:
                logger.warning("requested-day-offs 取得失敗 emp=%s: %s", employee_id, e)
                return {}
            if r.status_code == 429 and attempt < 2:
                _time.sleep(3 * (attempt + 1))
                continue
            break
        if r is None or r.status_code != 200:
            logger.warning("requested-day-offs 取得失敗 emp=%s status=%s",
                           employee_id, r.status_code if r is not None else "N/A")
            return {}
        return parse_requested_day_offs_data(r.json().get("data") or {})

    def get_attendance_times(self, employee_id: str, month: str) -> dict[str, dict]:
        """`/v1/employees/attendances` から日別の出退勤時刻を取得する。

        Returns:
            ``{date_iso: {"in": "16:45", "out": "33:30", "absent": bool}}``
            退勤が翌日以降の場合は 24時超表記へ補正する。重複レコードは先勝ちで除去。
        """
        import re as _re
        from datetime import date as _date

        headers = self._auth_headers()
        try:
            r = requests.get(
                f"{self.base_url}/v1/employees/attendances",
                headers=headers,
                params={"employee-id": str(employee_id), "month": month},
                timeout=60,
            )
        except requests.RequestException as e:
            raise JinjerAPIError(f"勤怠実績取得に失敗 emp={employee_id}: {e}") from e
        result: dict[str, dict] = {}
        if r.status_code != 200:
            logger.warning("attendances 取得失敗 emp=%s status=%s", employee_id, r.status_code)
            return result
        data = r.json().get("data") or {}

        def _ts_to_time(ts: str, base_iso: str) -> str:
            m = _re.match(r"^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})", (ts or "").strip())
            if not m:
                return ""
            dt_date = _date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            base = _date.fromisoformat(base_iso)
            hh = int(m.group(4)) + (dt_date - base).days * 24
            return f"{hh}:{m.group(5)}"

        for day in data.get("attendances", []) or []:
            d = str(day.get("date") or "").strip()
            if not d or d in result:  # 同一日の完全同一レコード二重返却対策
                continue
            result[d] = {
                "in": _ts_to_time(day.get("attended_at") or "", d),
                "out": _ts_to_time(day.get("left_at") or "", d),
                "absent": bool(day.get("is_absent")),
            }
        return result

    # ------------------------------------------------------------------
    # 日次勤怠の打刻コメント（打刻修正申請の理由など）
    # ------------------------------------------------------------------
    def get_stamp_comments(
        self, employee_ids: list[str], month: str
    ) -> dict[tuple[str, str], list[dict]]:
        """指定従業員・対象月の「打刻に付いた非空コメント」を取得する。

        `/v1/employees/attendances`（1名×1月ずつ）の各 `stamps[]` から、
        `comment` が空でないものだけを拾う。打刻修正申請の従業員コメント
        （`stamp_method=="打刻修正申請"`）や管理者修正の理由、打刻時コメントが入る。

        Args:
            employee_ids: 従業員ID（文字列）のリスト
            month: 対象月 ``"YYYY-MM"``

        Returns:
            ``{(employee_id, "YYYY-MM-DD"): [{"type","method","comment"}, ...]}``。
            非空コメントが無い従業員・日付はキーごと含めない。
            同一打刻が重複して返る既知問題に備え ``(date, type, stamped_at, comment)``
            でユニーク化する。1名でも取得失敗したら、その従業員はスキップして続行。
        """
        headers = self._auth_headers()
        url = f"{self.base_url}/v1/employees/attendances"
        result: dict[tuple[str, str], list[dict]] = {}
        seen: set[tuple] = set()

        unique_ids = []
        _seen_id = set()
        for emp_id in employee_ids:
            s = str(emp_id or "").strip()
            if s and s not in _seen_id:
                _seen_id.add(s)
                unique_ids.append(s)

        for idx, emp_id in enumerate(unique_ids):
            params = {"employee-id": emp_id, "month": month}
            attendances = None
            for attempt in range(1, 4):
                try:
                    r = requests.get(url, headers=headers, params=params, timeout=self.timeout)
                except requests.RequestException as e:
                    logger.warning("打刻コメント取得 通信失敗 emp=%s: %s", emp_id, e)
                    break
                if r.status_code == 200:
                    attendances = (r.json().get("data") or {}).get("attendances", []) or []
                    break
                if r.status_code == 429:
                    _time.sleep(2 * attempt)
                    continue
                if r.status_code == 404:
                    attendances = []
                    break
                logger.warning("打刻コメント取得 失敗 emp=%s status=%s", emp_id, r.status_code)
                break

            for day in attendances or []:
                date_str = str(day.get("date") or "").strip()
                if not date_str:
                    continue
                for st in day.get("stamps", []) or []:
                    comment = str(st.get("comment") or "").strip()
                    if not comment:
                        continue
                    stype = str((st.get("stamp_type") or {}).get("name") or "").strip()
                    method = str(st.get("stamp_method") or "").strip()
                    stamped_at = str(st.get("stamped_at") or "").strip()
                    dedup_key = (emp_id, date_str, stype, stamped_at, comment)
                    if dedup_key in seen:
                        continue
                    seen.add(dedup_key)
                    result.setdefault((emp_id, date_str), []).append(
                        {"type": stype, "method": method, "comment": comment}
                    )

            if idx + 1 < len(unique_ids):
                _time.sleep(0.2)  # 軽いペーシング

        logger.info("jinjer 打刻コメント取得: %d (emp,date) 件", len(result))
        return result


def fetch_stamp_comments(employee_ids: list[str], month: str) -> dict[tuple[str, str], list[dict]]:
    """API を叩いて打刻コメントの辞書を返す薄いラッパ。

    Raises:
        JinjerAPIError: 認証失敗時のみ（個別従業員の取得失敗はスキップして続行する）。
    """
    client = JinjerClient()
    return client.get_stamp_comments(employee_ids, month)


# ----------------------------------------------------------------------
# 氏名 → ID マッピング
# ----------------------------------------------------------------------

def _safe_get(d: dict, *keys, default=""):
    cur: Any = d
    for k in keys:
        if not isinstance(cur, dict):
            return default
        cur = cur.get(k)
        if cur is None:
            return default
    return cur or default


def _name_key_candidates(employees: list[dict]) -> dict[str, list[tuple[str, str]]]:
    """氏名キー → [(従業員ID, フルネーム), ...] を集める（build_name_to_id_map の内部）。"""
    key_hits: dict[str, list[tuple[str, str]]] = {}
    for emp in employees:
        if not isinstance(emp, dict):
            continue
        emp_id = emp.get("id") or emp.get("employee_id")
        if emp_id is None:
            continue
        emp_id_str = str(emp_id).strip()
        if not emp_id_str:
            continue

        last = str(_safe_get(emp, "company", "last_name") or "").strip()
        first = str(_safe_get(emp, "company", "first_name") or "").strip()
        full = f"{last} {first}".strip()

        candidates = []
        if last and first:
            candidates.extend([
                f"{last}{first}",
                f"{last} {first}",
                f"{last}　{first}",
            ])
        if last:
            candidates.append(last)
        if first:
            candidates.append(first)

        for name in candidates:
            if not name:
                continue
            hits = key_hits.setdefault(name, [])
            if all(eid != emp_id_str for eid, _ in hits):
                hits.append((emp_id_str, full))
    return key_hits


def build_name_to_id_map(employees: list[dict]) -> dict[str, str]:
    """jinjer 従業員一覧 → 氏名 → ID の辞書

    氏名のバリエーション（"姓名" / "姓 名" / "姓　名" / 姓のみ / 名のみ）を
    同じ ID にマップしておくことで、勤務表側の表記揺れに対応する。

    **同じ氏名キーが複数の従業員に該当する場合（同姓の「吉田」等）は登録しない**。
    以前は先勝ちで最初の1人の ID を返しており、勤務表の「吉田」が別人の吉田の
    ID に静かに化ける事故リスクがあった（2026-07-22 KDXシフト表で実例）。
    曖昧キーは build_ambiguous_names() で候補ごと取得し、警告表示に使う。

    Args:
        employees: JinjerClient.get_employees() の戻り値

    Returns:
        {氏名(str): 従業員ID(str)}
    """
    return {name: hits[0][0]
            for name, hits in _name_key_candidates(employees).items()
            if len(hits) == 1}


def build_ambiguous_names(employees: list[dict]) -> dict[str, list[tuple[str, str]]]:
    """複数の従業員に該当してしまう氏名キー → [(従業員ID, フルネーム), ...]。

    同姓（例: 吉田英伸 / 吉田拓矢 の「吉田」）など、自動では確定できない
    キーの候補一覧。勤務表側でこの名前が出た場合、警告に候補を並べて
    人間に選ばせるために使う。
    """
    return {name: hits
            for name, hits in _name_key_candidates(employees).items()
            if len(hits) > 1}


def build_id_to_official_name(employees: list[dict]) -> dict[str, str]:
    """jinjer 従業員一覧 → 従業員ID → jinjer 登録の「姓のみ」の辞書

    jinjer のスケジュール CSV インポートは 氏名 + 従業員ID の組で照合するため、
    勤務表側の表記揺れ（"大堀 広智" など）をそのまま書き出すと
    "登録されていない従業員ID" として弾かれる。
    公式サンプル CSV が姓のみ（"大貫" 等）で記載されているのに合わせる。
    """
    id_to_name: dict[str, str] = {}
    for emp in employees:
        if not isinstance(emp, dict):
            continue
        emp_id = emp.get("id") or emp.get("employee_id")
        if emp_id is None:
            continue
        emp_id_str = str(emp_id).strip()
        if not emp_id_str:
            continue
        last = str(_safe_get(emp, "company", "last_name") or "").strip()
        if last:
            id_to_name[emp_id_str] = last
    return id_to_name


def fetch_employee_id_map() -> tuple[dict[str, str], dict[str, str], dict[str, list[tuple[str, str]]]]:
    """API を叩いて (氏名→ID マップ, 従業員ID→姓 マップ, 曖昧氏名→候補) を返す

    Raises:
        JinjerAPIError: 認証失敗 / 取得失敗
    """
    client = JinjerClient()
    employees = client.get_employees(only_active=True)
    return (build_name_to_id_map(employees),
            build_id_to_official_name(employees),
            build_ambiguous_names(employees))


# ----------------------------------------------------------------------
# 打刻グループ判定
# ----------------------------------------------------------------------

def _parse_iso_date_safe(s):
    """'YYYY-MM-DD' を date に変換。失敗時 None"""
    from datetime import date as _date
    if not s:
        return None
    if isinstance(s, _date):
        return s
    try:
        return _date.fromisoformat(str(s).strip())
    except (ValueError, TypeError):
        return None


def pick_attendance_group_at(affiliations: list[dict], target_date) -> tuple[str, str]:
    """対象日時点で適用される打刻グループ (id, name) を返す

    アルゴリズム:
        1. ``date_of_issue`` 昇順にソート
        2. ``date_of_issue <= target_date`` のレコードを抽出
        3. 抽出後の最新レコードの ``attendance_group.id/name`` を返す
        4. ``attendance_group.id`` が空なら、さらに古い方へ遡って非空のものを採用
           （履歴上「未設定」レコードが挟まることがあるため）
        5. それでも取れなければ ``("", "")`` を返す

    Args:
        affiliations: ``JinjerClient.get_affiliations()`` の値（特定従業員1名分）
        target_date: ``datetime.date`` 想定。文字列なら ISO 形式

    Returns:
        (attendance_group_id, attendance_group_name)
    """
    target_d = _parse_iso_date_safe(target_date) if not hasattr(target_date, "year") else target_date
    if target_d is None:
        return ("", "")

    dated = []
    for a in affiliations or []:
        if not isinstance(a, dict):
            continue
        d = _parse_iso_date_safe(a.get("date_of_issue"))
        if d is None:
            continue
        ag = a.get("attendance_group") or {}
        gid = str(ag.get("id") or "").strip()
        gname = str(ag.get("name") or "").strip()
        dated.append((d, gid, gname))

    if not dated:
        return ("", "")

    dated.sort(key=lambda x: x[0])
    candidates = [t for t in dated if t[0] <= target_d]
    if not candidates:
        return ("", "")

    # 末尾から非空の attendance_group を探す
    for d, gid, gname in reversed(candidates):
        if gid:
            return (gid, gname)

    # 全部空 → 末尾 (空文字) を返す
    _, gid, gname = candidates[-1]
    return (gid, gname)


def fetch_attendance_groups_at(
    employee_ids: list[str],
    target_date,
) -> dict[str, tuple[str, str]]:
    """指定された従業員IDリストについて、対象日時点の打刻グループ (id, name) を返す

    Args:
        employee_ids: 従業員ID 文字列のリスト
        target_date: 対象日（``datetime.date`` または ``"YYYY-MM-DD"``）

    Returns:
        {employee_id(str): (group_id, group_name)}。
        所属履歴が取れなかった従業員は ``("", "")`` を返す。

    Raises:
        JinjerAPIError: 認証/通信エラー
    """
    client = JinjerClient()
    affiliations_map = client.get_affiliations(employee_ids)
    result: dict[str, tuple[str, str]] = {}
    for emp_id in employee_ids:
        s = str(emp_id or "").strip()
        if not s:
            continue
        result[s] = pick_attendance_group_at(affiliations_map.get(s, []), target_date)
    return result
