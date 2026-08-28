# -*- coding: utf-8 -*-
"""パスと四半期ヘルパー。データは Z:\派遣元管理台帳（リポジトリ外・個人情報）、コードはここ。"""
from __future__ import annotations

import datetime as dt
import os
import re
from pathlib import Path

DATA_ROOT = Path(os.environ.get("DAICHO_DATA_ROOT", r"Z:\派遣元管理台帳"))
INPUT_DIR = DATA_ROOT / "input"
OUTPUT_DIR = DATA_ROOT / "output"
TEMPLATE_DIR = DATA_ROOT / "テンプレート"
LOG_DIR = DATA_ROOT / "ログ"

# 旧 .xlsm の台帳フォーム（レイアウトの元）。元ファイルは読み取りのみ。
SOURCE_FORM_XLSM = TEMPLATE_DIR / "元_派遣元管理台帳202504.xlsm"
SOURCE_FORM_SHEET = "派遣元管理台帳"
TEMPLATE_XLSX = TEMPLATE_DIR / "派遣元管理台帳_テンプレート.xlsx"
TEMPLATE_SHEET = "台帳"

# 派遣元の固定情報（契約データにも入っているが、空のときの既定値）
DISPATCH_SOURCE_NAME = "株式会社エヌエム・ヒューマテック"
DISPATCH_LICENSE_NO = "派13-301312"

_Q = re.compile(r"^(\d{4})Q([1-4])$")


def quarter_range(q: str) -> tuple[dt.date, dt.date]:
    """'2025Q3' → (2025-07-01, 2025-09-30)。Q1=1-3月（暦四半期）。"""
    m = _Q.match(q.strip().upper())
    if not m:
        raise ValueError(f"四半期の指定は 2025Q3 のような形式で: {q!r}")
    y, n = int(m.group(1)), int(m.group(2))
    start = dt.date(y, 3 * (n - 1) + 1, 1)
    end_month = 3 * n
    last_day = (dt.date(y + (end_month // 12), end_month % 12 + 1, 1) - dt.timedelta(days=1))
    return start, last_day


def quarter_label(q: str) -> str:
    start, end = quarter_range(q)
    return f"{start.year}年{start.month}-{end.month}月期"


def overlaps(a_start: dt.date | None, a_end: dt.date | None, b_start: dt.date, b_end: dt.date) -> bool:
    if a_start is None or a_end is None:
        return False
    return a_start <= b_end and a_end >= b_start
