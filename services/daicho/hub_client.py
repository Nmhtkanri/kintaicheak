# -*- coding: utf-8 -*-
"""jinjer_attach が使うクライアント（本アプリ JinjerWriteClient への橋渡し）。

旧 JinjerLaborContractClient（Z:\\API連携）が提供していた authenticate() と
_request(method, path, json_body=, params=) の2点だけを同じ形で提供する
（2026-08-28 ハブ移設。jinjer_attach 側のロジックは無修正で動かすため）。

リトライの分担（二層構造を維持）:
- 書き込み間隔・429 の Retry-After・401 再認証 → JinjerWriteClient._write が持つ
- 429 を 180 秒×10回粘る・403 トークン失効→再認証 → jinjer_attach._write_with_retry が持つ
  （_write が投げる「…に失敗 (status=429): …」形式のメッセージ文字列で判定が噛み合う）
"""
from __future__ import annotations

import time as _time

import requests

from services.jinjer_api_client import JinjerAPIError, JinjerWriteClient


class DaichoAttachClient:
    """authenticate() / _request() だけの薄いラッパ。

    継承でなく内包にしているのは authenticate() の意味を変えるため:
    JinjerClient.authenticate() は TTL 内キャッシュを返すが、jinjer_attach の
    _write_with_retry は「403 トークン失効→ authenticate() で必ず取り直す」を期待する。
    """

    MAX_RETRIES = 3

    def __init__(self, write_interval: float):
        self._c = JinjerWriteClient(write_interval=write_interval)

    def authenticate(self) -> str:
        self._c._access_token = None  # キャッシュを捨てて必ず取り直す
        return self._c.authenticate()

    def _request(self, method: str, path: str, json_body: dict | None = None,
                 params: dict | None = None):
        m = str(method).upper()
        if m == "GET":
            return self._get(path, params or {})
        return self._c._write(m, path, json_body or {}, what=f"{m} {path}")

    def _get(self, path: str, params: dict) -> list | dict:
        """GET して data 部を返す（keiri_api._get と同型＋403 トークン失効も1回だけ再認証）。

        fetch_records は list / dict 両対応で受けるので data 部を返せばよい。
        """
        url = f"{self._c.base_url}{path}"
        refreshed = False
        for attempt in range(1, self.MAX_RETRIES + 1):
            try:
                response = requests.get(url, headers=self._c._auth_headers(),
                                        params=params, timeout=self._c.timeout)
            except requests.RequestException as exc:
                if attempt >= self.MAX_RETRIES:
                    raise JinjerAPIError(f"jinjer API へ接続できませんでした: {exc}") from exc
                _time.sleep(min(2 ** (attempt - 1), 5))
                continue
            if response.status_code in (401, 403) and not refreshed:
                self._c._access_token = None  # 次の _auth_headers で取り直す
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
                raise JinjerAPIError(f"GET {path} に失敗 (status={response.status_code})")
            try:
                return response.json().get("data", {})
            except ValueError as exc:
                raise JinjerAPIError("jinjer API から JSON 以外の応答が返されました。") from exc
        raise JinjerAPIError(f"GET {path}: 再試行回数を超えました。")
