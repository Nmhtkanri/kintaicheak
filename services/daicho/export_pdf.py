# -*- coding: utf-8 -*-
"""生成済みの四半期台帳ブックから、契約ごとのPDFを人別フォルダへ書き出す。

出力先: Z:\\派遣元管理台帳\\PDF\\{社員番号}_{氏名}\\{社員番号}_{氏名}_{契約期間}分.pdf
- 「契約期間」はそのシートの契約期間（E17）から作る（例: 2025年7-9月分／2025年7月分／
  2025年10月-2026年2月分）。四半期をまたぐ契約は複数の期のブックに同内容で載るが、
  ファイル名が同じになるので自然に1本にまとまる（上書き）
- 社員番号が空のシートは「未確定_氏名」フォルダに出し、警告として返す

Excel COM（pywin32）を使う。読み取り専用で開き、元のブックは変更しない。
"""
from __future__ import annotations

import re
from pathlib import Path

from .config import DATA_ROOT, OUTPUT_DIR, quarter_range
from .estaffing import parse_date

PDF_ROOT = DATA_ROOT / "PDF"
_INVALID = re.compile(r'[\\/:*?"<>|]')


def sanitize(name: str) -> str:
    return _INVALID.sub("", re.sub(r"[\s　]+", "", str(name or ""))).strip() or "名前不明"


def period_label(period_text: str) -> str | None:
    """'2025/07/01～2025/09/30' → '2025年7-9月'。単月・年またぎにも対応。"""
    parts = re.split(r"[～〜]", str(period_text or ""))
    if len(parts) != 2:
        return None
    s, e = parse_date(parts[0].strip()), parse_date(parts[1].strip())
    if not (s and e):
        return None
    if s.year == e.year:
        if s.month == e.month:
            return f"{s.year}年{s.month}月"
        return f"{s.year}年{s.month}-{e.month}月"
    return f"{s.year}年{s.month}月-{e.year}年{e.month}月"


def export_quarter(quarter: str, xlsx_path: Path | str | None = None,
                   pdf_root: Path | str = PDF_ROOT) -> tuple[int, list[str]]:
    """1四半期分のブックをPDF化する。戻り値: (出力枚数, 警告)。"""
    import pythoncom
    import win32com.client.dynamic

    q = quarter.upper()
    quarter_range(q)                      # 形式チェック
    xlsx = Path(xlsx_path) if xlsx_path else OUTPUT_DIR / f"派遣元管理台帳_{q}.xlsx"
    if not xlsx.exists():
        raise FileNotFoundError(f"台帳ブックが無い: {xlsx}（先に build を実行）")
    root = Path(pdf_root)
    root.mkdir(parents=True, exist_ok=True)

    warnings: list[str] = []
    n = 0
    pythoncom.CoInitialize()
    # 新規インスタンスで開く（ユーザーが開いているExcelに乗らない）。
    # gencacheが壊れてDispatchExが落ちる場合は %LOCALAPPDATA%\Temp\gen_py を削除して再実行
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Open(str(xlsx), ReadOnly=True, UpdateLinks=0)
        try:
            for ws in wb.Worksheets:
                if ws.Name in ("目次", "警告"):
                    continue
                emp = str(ws.Range("D2").Value or "").strip()
                if emp.endswith(".0"):
                    emp = emp[:-2]
                name = sanitize(ws.Range("E2").Value)
                label = period_label(str(ws.Range("E17").Value or ""))
                if label is None:
                    label = q
                    warnings.append(f"{ws.Name}: 契約期間が読めないためファイル名を {q} にした")
                if not emp:
                    emp = "未確定"
                    warnings.append(f"{ws.Name}: 社員番号が空 → 未確定_{name}")
                folder = root / f"{emp}_{name}"
                folder.mkdir(parents=True, exist_ok=True)
                out = folder / f"{emp}_{name}_{label}分.pdf"
                ws.ExportAsFixedFormat(0, str(out))
                n += 1
        finally:
            wb.Close(False)
    finally:
        xl.Quit()
        pythoncom.CoUninitialize()
    return n, warnings
