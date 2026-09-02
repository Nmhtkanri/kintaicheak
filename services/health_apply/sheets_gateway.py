# -*- coding: utf-8 -*-
"""健康診断申込: Google Sheets API v4 への薄い出入口。

- 認証は健診専用サービスアカウントの鍵JSON（管理者PCのローカル）。google-auth の
  AuthorizedSession で REST を直接叩く（gspread は使わない）。
- google-auth の import は `_open_session` の中だけ。未導入の環境でも本モジュールの
  import は通り、実際に使おうとした時点で GatewayConfigError になる。
- 書き込みは append のみ、かつ schema.WRITABLE_SHEETS（対象者・監査ログ）に限る。
  回答・選択肢・設定には構造上書けない。
- 値は valueInputOption=RAW で送る。前ゼロ付きの HPM コード（0301619）や英字入り
  （13X5035440）を Google に数値へ変えさせないため。
"""

from __future__ import annotations

import os
import urllib.parse
from typing import Protocol

from services.health_apply.schema import SchemaError, assert_writable

SCOPES = ("https://www.googleapis.com/auth/spreadsheets",)
BASE = "https://sheets.googleapis.com/v4/spreadsheets"


class GatewayError(RuntimeError):
    """通信・認可の失敗（ルートは 500 にする）。"""


class GatewayConfigError(GatewayError):
    """鍵JSONが無い・google-auth 未導入など運用設定の不備（ルートは 400 にする）。"""


class SheetsGateway(Protocol):
    """テストで FakeSheetsGateway に差し替える境界。"""

    def get_metadata(self) -> dict: ...

    def read_values(self, sheet: str, a1_range: str = "") -> list[list[str]]: ...

    def append_rows(self, sheet: str, rows: list[list[str]]) -> int: ...


def quote_range(sheet: str, a1_range: str = "") -> str:
    """'対象者'!A:Z をURLパスに載せられる形にする（シート名のクォートも含めて全部エンコード）。"""
    title = str(sheet).replace("'", "''")
    ref = f"'{title}'" + (f"!{a1_range}" if a1_range else "")
    return urllib.parse.quote(ref, safe="")


def sheet_titles(metadata: dict) -> list[str]:
    return [str(s.get("title", "")) for s in (metadata or {}).get("sheets", [])]


def _cells(rows) -> list[list[str]]:
    return [["" if v is None else str(v) for v in row] for row in rows]


class GoogleSheetsGateway:
    def __init__(self, spreadsheet_id: str, service_account_path: str, timeout: float = 30.0):
        self.spreadsheet_id = str(spreadsheet_id or "").strip()
        self.service_account_path = str(service_account_path or "")
        self.timeout = float(timeout)
        self.client_email = ""
        self._session = None

    # --- 認証 ---------------------------------------------------------
    def _open_session(self):
        if self._session is not None:
            return self._session
        if not self.spreadsheet_id:
            raise GatewayConfigError("スプレッドシートIDが空です（年度設定JSONを確認してください）")
        if not self.service_account_path or not os.path.exists(self.service_account_path):
            raise GatewayConfigError(
                f"サービスアカウントの鍵JSONがありません: {self.service_account_path}"
                "（管理者PCのローカルに置いてください。共有フォルダやリポジトリには置かない）")
        try:
            from google.oauth2 import service_account
            from google.auth.transport.requests import AuthorizedSession
        except ImportError as e:  # pragma: no cover - 環境依存
            raise GatewayConfigError(
                "google-auth が入っていません（pip install google-auth）") from e
        try:
            creds = service_account.Credentials.from_service_account_file(
                self.service_account_path, scopes=list(SCOPES))
        except (OSError, ValueError) as e:
            raise GatewayConfigError(f"サービスアカウントの鍵JSONを読めません: {e}") from e
        self.client_email = str(getattr(creds, "service_account_email", "") or "")
        self._session = AuthorizedSession(creds)
        return self._session

    # --- HTTP ---------------------------------------------------------
    def _request(self, method: str, url: str, *, params: dict | None = None,
                 json_body: dict | None = None) -> dict:
        session = self._open_session()
        try:
            resp = session.request(method, url, params=params, json=json_body, timeout=self.timeout)
        except Exception as e:  # requests / google.auth の例外をまとめて
            raise GatewayError(f"Google Sheets API に接続できません: {e}") from e
        status = getattr(resp, "status_code", 0)
        if status == 403:
            who = self.client_email or "サービスアカウント"
            raise GatewayError(
                f"スプレッドシートにアクセスできません（HTTP 403）。{who} に"
                "編集者として共有されているか、Sheets API が有効かを確認してください")
        if status == 404:
            raise GatewayError(
                "スプレッドシートが見つかりません（HTTP 404）。年度設定JSONの spreadsheet_id を確認してください")
        if status != 200:
            detail = ""
            try:
                detail = str(resp.json().get("error", {}).get("message", ""))
            except Exception:
                detail = str(getattr(resp, "text", ""))[:200]
            raise GatewayError(f"Google Sheets API エラー HTTP {status}: {detail}")
        try:
            return resp.json()
        except ValueError as e:
            raise GatewayError("Google Sheets API から JSON 以外の応答が返されました") from e

    # --- 公開メソッド -------------------------------------------------
    def get_metadata(self) -> dict:
        data = self._request("GET", f"{BASE}/{self.spreadsheet_id}",
                             params={"fields": "properties.title,sheets.properties"})
        sheets = []
        for s in data.get("sheets", []) or []:
            props = s.get("properties", {}) or {}
            grid = props.get("gridProperties", {}) or {}
            sheets.append({
                "title": str(props.get("title", "")),
                "sheetId": props.get("sheetId"),
                "rowCount": grid.get("rowCount"),
                "columnCount": grid.get("columnCount"),
            })
        return {"title": str((data.get("properties", {}) or {}).get("title", "")), "sheets": sheets}

    def read_values(self, sheet: str, a1_range: str = "") -> list[list[str]]:
        data = self._request(
            "GET", f"{BASE}/{self.spreadsheet_id}/values/{quote_range(sheet, a1_range)}",
            params={"valueRenderOption": "FORMATTED_VALUE",
                    "dateTimeRenderOption": "FORMATTED_STRING"})
        return _cells(data.get("values", []) or [])

    def append_rows(self, sheet: str, rows: list[list[str]]) -> int:
        try:
            assert_writable(sheet)
        except SchemaError as e:
            raise GatewayError(str(e)) from e
        if not rows:
            return 0
        data = self._request(
            "POST", f"{BASE}/{self.spreadsheet_id}/values/{quote_range(sheet, 'A1')}:append",
            params={"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"},
            json_body={"values": _cells(rows)})
        updates = data.get("updates", {}) or {}
        try:
            return int(updates.get("updatedRows", len(rows)) or 0)
        except (TypeError, ValueError):
            return len(rows)
