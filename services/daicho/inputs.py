# -*- coding: utf-8 -*-
"""入力ファイルの解決（build と鮮度チェックで共有し、glob 定義の二重管理を防ぐ）。

input/ 直下から「mtime 最新の1本」を自動選択する。パターンを変えるときはここだけ直す。
"""
from __future__ import annotations

from pathlib import Path

# build が input/ から自動選択する glob
PATTERN_TC = "TCnmht*.csv"
PATTERN_CPI = "CPInmht*.csv"
PATTERN_ROSTER = "従業員一覧*.xlsx"
PATTERN_FG_WO = "*WorkOrder*.csv"
PATTERN_FG_DETAILS = "*fieldglass_details*.json"


def newest(folder: Path, pattern: str) -> Path:
    """folder 直下で pattern に合う mtime 最新の1本。無ければ FileNotFoundError。"""
    files = sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"入力が見つかりません: {folder}\\{pattern}")
    return files[0]


def newest_or_none(folder: Path, pattern: str) -> Path | None:
    """newest と同じ選び方で、無ければ None（Fieldglass 系は任意入力のため）。"""
    files = sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None
