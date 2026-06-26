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
import time as _time
from typing import Any

import requests

from config import Config

logger = logging.getLogger(__name__)


class JinjerAPIError(Exception):
    """jinjer API 関連の例外"""


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


def build_name_to_id_map(employees: list[dict]) -> dict[str, str]:
    """jinjer 従業員一覧 → 氏名 → ID の辞書

    氏名のバリエーション（"姓名" / "姓 名" / "姓　名"）を全て同じ ID に
    マップしておくことで、勤務表側の表記揺れに対応する。

    Args:
        employees: JinjerClient.get_employees() の戻り値

    Returns:
        {氏名(str): 従業員ID(str)}
    """
    name_map: dict[str, str] = {}
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
            if name and name not in name_map:
                name_map[name] = emp_id_str

    return name_map


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


def fetch_employee_id_map() -> tuple[dict[str, str], dict[str, str]]:
    """API を叩いて (氏名→ID マップ, 従業員ID→姓 マップ) を返す

    Raises:
        JinjerAPIError: 認証失敗 / 取得失敗
    """
    client = JinjerClient()
    employees = client.get_employees(only_active=True)
    return build_name_to_id_map(employees), build_id_to_official_name(employees)


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
