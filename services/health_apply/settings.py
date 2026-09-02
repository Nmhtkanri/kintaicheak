# -*- coding: utf-8 -*-
"""健康診断申込: 年度設定JSON（NAS）とサービスアカウント鍵の所在。

年度設定JSON（Config.HEALTH_APPLY_SETTINGS_JSON）:
    {"schema": 1, "default_year": "2027",
     "years": {"2027": {"spreadsheet_id": "...", "webapp_url": "...",
                        "previous_year": 2026, "label": "2027年度"}}}
翌年度は years に1項目足して default_year を変えるだけ。

サービスアカウント鍵JSONは管理者PCのローカルにだけ置く。ここでは存在確認と
client_email（スプレッドシートの共有先として画面に出す）しか読まない。
"""

from __future__ import annotations

import datetime as _dt
import json
import os
from dataclasses import dataclass


class SettingsError(ValueError):
    """年度設定JSONが読めない・年度が無いなど、運用設定の不備（画面に文言を出す）。"""


@dataclass(frozen=True)
class YearSettings:
    fiscal_year: int
    previous_year: int
    spreadsheet_id: str
    webapp_url: str = ""
    label: str = ""

    @property
    def key(self) -> str:
        return str(self.fiscal_year)

    @property
    def spreadsheet_id_tail(self) -> str:
        """画面表示用。IDそのものは秘密ではないが、取り違え確認に末尾だけ出す。"""
        return "…" + self.spreadsheet_id[-6:] if len(self.spreadsheet_id) > 6 else self.spreadsheet_id


@dataclass(frozen=True)
class YearConfig:
    years: dict[str, YearSettings]
    default_year: str
    path: str
    mtime: str

    def year_keys(self) -> list[str]:
        return sorted(self.years.keys(), reverse=True)

    def pick(self, requested: str | None) -> YearSettings:
        key = str(requested or "").strip() or self.default_year
        if key not in self.years:
            raise SettingsError(
                f"年度「{key}」は年度設定JSONにありません（登録済み: "
                + "、".join(self.year_keys()) + f"）。{self.path} に足してください")
        return self.years[key]


def _mtime_iso(path: str) -> str:
    try:
        return _dt.datetime.fromtimestamp(os.path.getmtime(path)).isoformat(timespec="seconds")
    except OSError:
        return ""


def load_year_config(path: str) -> YearConfig:
    """年度設定JSONを読む。無い・壊れている・年度が空なら SettingsError。"""
    if not path or not os.path.exists(path):
        raise SettingsError(f"年度設定JSONがありません: {path}")
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            raw = json.load(f)
    except (OSError, ValueError) as e:
        raise SettingsError(f"年度設定JSONを読めません: {path}（{e}）") from e
    years_raw = raw.get("years") if isinstance(raw, dict) else None
    if not isinstance(years_raw, dict) or not years_raw:
        raise SettingsError(f"年度設定JSONに years がありません: {path}")

    years: dict[str, YearSettings] = {}
    for key, item in years_raw.items():
        try:
            fiscal_year = int(str(key).strip())
        except ValueError as e:
            raise SettingsError(f"年度設定JSONの年度キーが数値ではありません: 「{key}」") from e
        if not isinstance(item, dict):
            raise SettingsError(f"年度設定JSONの {key} が辞書ではありません")
        spreadsheet_id = str(item.get("spreadsheet_id", "")).strip()
        if not spreadsheet_id:
            raise SettingsError(f"年度設定JSONの {key} に spreadsheet_id がありません")
        previous_raw = item.get("previous_year", fiscal_year - 1)
        try:
            previous_year = int(previous_raw)
        except (TypeError, ValueError) as e:
            raise SettingsError(f"年度設定JSONの {key} の previous_year が数値ではありません") from e
        years[str(fiscal_year)] = YearSettings(
            fiscal_year=fiscal_year,
            previous_year=previous_year,
            spreadsheet_id=spreadsheet_id,
            webapp_url=str(item.get("webapp_url", "")).strip(),
            label=str(item.get("label", "")).strip() or f"{fiscal_year}年度",
        )
    default_year = str(raw.get("default_year", "")).strip() or max(years.keys())
    if default_year not in years:
        raise SettingsError(f"年度設定JSONの default_year「{default_year}」が years にありません")
    return YearConfig(years=years, default_year=default_year, path=path, mtime=_mtime_iso(path))


def service_account_info(path: str) -> dict:
    """鍵JSONの存在と client_email だけを返す（秘密鍵は読まない・返さない）。"""
    info = {"path": path, "exists": False, "client_email": "", "error": ""}
    if not path or not os.path.exists(path):
        info["error"] = f"サービスアカウントの鍵JSONがありません: {path}"
        return info
    info["exists"] = True
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        info["client_email"] = str(raw.get("client_email", "")).strip() if isinstance(raw, dict) else ""
        if not info["client_email"]:
            info["error"] = "鍵JSONに client_email がありません（サービスアカウントの鍵ではない可能性）"
    except (OSError, ValueError) as e:
        info["error"] = f"鍵JSONを読めません: {e}"
    return info
