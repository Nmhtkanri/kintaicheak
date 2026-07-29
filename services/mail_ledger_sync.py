# -*- coding: utf-8 -*-
"""メール台帳（メール一斉送信マクロ.xlsm）を jinjer API と同期する。

方針（2026-07-29 谷津さん依頼）:
- 追加: jinjer 在籍・社員番号が 20YY 形式・台帳に無い人。D=社用メール(company.email)、
  F=個人メール(personal.email)、E=就業先は jinjer に無いので空欄（手動維持）。
- 削除: 台帳にいて jinjer が「退職」の人。行ごと削除するが、実行前に必ずバックアップを作る。
- 既存行の D/E/F は更新しない（台帳側の手修正を守る）。jinjer と違う値は「不一致」として報告のみ。
- 台帳シートの実構造: シート「メール送信」、3行目ヘッダー（B=社員番号/C=氏名/D=社用/E=就業先/F=個人）、
  データは4行目から。A列は一斉送信マクロのチェック列なので触らない。1〜2行目は件名・本文の設定領域。
- 差分計算（プレビュー）は読み取りのみ。書き込みは Excel COM（xlsm の VBA・書式を完全保持するため）。
"""
from __future__ import annotations

import csv
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from config import Config

REGULAR_ID_PATTERN = re.compile(r"^20\d{5}$")
LEDGER_SHEET = "メール送信"
LOG_FILENAME = "台帳更新ログ.csv"
LOG_HEADERS = ["処理日時", "種別", "社員番号", "氏名", "詳細"]
XL_UP = -4162


def _clean_email(value: Any) -> str:
    return str(value or "").strip()


def fetch_jinjer_directory(client: Any = None) -> list[dict[str, Any]]:
    """jinjer の全従業員（退職者込み）を台帳同期用に正規化して返す。"""
    if client is None:
        from services.jinjer_api_client import JinjerClient
        client = JinjerClient()
    employees = client.get_employees(only_active=False)
    directory: list[dict[str, Any]] = []
    for emp in employees:
        company = emp.get("company") or {}
        personal = emp.get("personal") or {}
        enrollment = (company.get("enrollment_classification") or {}).get("name") or ""
        directory.append({
            "id": str(emp.get("id") or "").strip(),
            "name": f"{company.get('last_name') or ''} {company.get('first_name') or ''}".strip(),
            "company_email": _clean_email(company.get("email")),
            "personal_email": _clean_email(personal.get("email")),
            "enrollment": enrollment,
            "retired": "退職" in enrollment,
            "retirement_date": str(company.get("retirement_date") or ""),
        })
    return directory


def compute_ledger_diff(
    address_book: dict[str, list[dict[str, Any]]],
    directory: list[dict[str, Any]],
) -> dict[str, Any]:
    """台帳と jinjer を突合し、追加候補・削除候補・報告事項を返す（純ロジック）。

    - 対象は 20YY 形式の自社社員番号のみ。派遣・テスト番号（5/6/9始まり等）は触らない。
    - jinjer に存在しない台帳行は自動削除しない（番号誤りの可能性があるため報告のみ）。
    """
    by_id = {person["id"]: person for person in directory if person["id"]}

    additions: list[dict[str, Any]] = []
    for person in directory:
        if not REGULAR_ID_PATTERN.fullmatch(person["id"]):
            continue
        if person["retired"] or person["id"] in address_book:
            continue
        additions.append({
            "id": person["id"],
            "name": person["name"],
            "company_email": person["company_email"],
            "personal_email": person["personal_email"],
            "no_email": not (person["company_email"] or person["personal_email"]),
        })

    retirees: list[dict[str, Any]] = []
    missing_in_jinjer: list[dict[str, Any]] = []
    mismatches: list[str] = []
    for employee_id, entries in address_book.items():
        if not REGULAR_ID_PATTERN.fullmatch(employee_id):
            continue
        ledger_name = entries[0].get("name", "") if entries else ""
        person = by_id.get(employee_id)
        if person is None:
            missing_in_jinjer.append({"id": employee_id, "name": ledger_name})
            continue
        if person["retired"]:
            retirees.append({
                "id": employee_id,
                "name": ledger_name or person["name"],
                "retirement_date": person["retirement_date"],
            })
            continue
        entry = entries[0] if entries else {}
        ledger_company = tuple(entry.get("company") or ())
        ledger_personal = tuple(entry.get("personal") or ())
        jinjer_company = person["company_email"]
        jinjer_personal = person["personal_email"]
        if jinjer_company and jinjer_company.casefold() not in {a.casefold() for a in ledger_company}:
            mismatches.append(
                f"{employee_id} {ledger_name}: 社用が台帳とjinjerで不一致"
                f"（台帳: {'; '.join(ledger_company) or 'なし'} / jinjer: {jinjer_company}）")
        if jinjer_personal and jinjer_personal.casefold() not in {a.casefold() for a in ledger_personal}:
            mismatches.append(
                f"{employee_id} {ledger_name}: 個人が台帳とjinjerで不一致"
                f"（台帳: {'; '.join(ledger_personal) or 'なし'} / jinjer: {jinjer_personal}）")

    additions.sort(key=lambda item: item["id"])
    retirees.sort(key=lambda item: item["id"])
    missing_in_jinjer.sort(key=lambda item: item["id"])
    return {
        "additions": additions,
        "retirees": retirees,
        "missing_in_jinjer": missing_in_jinjer,
        "mismatches": mismatches,
    }


def _append_log(log_path: Path, rows: list[list[str]]) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    is_new = not log_path.exists()
    encoding = "utf-8-sig" if is_new else "utf-8"
    with open(log_path, "a", encoding=encoding, newline="") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(LOG_HEADERS)
        writer.writerows(rows)


def _find_header_row(sheet: Any) -> int:
    for row in range(1, 31):
        value = str(sheet.Cells(row, 2).Value or "").strip()
        if value in ("社員番号", "従業員番号"):
            return row
    raise RuntimeError("台帳シートの社員番号ヘッダー（B列）が見つかりません")


def apply_ledger_update(
    xlsm_path: str | Path,
    additions: list[dict[str, Any]],
    retiree_ids: list[str],
    *,
    backup_dir: str | Path | None = None,
    log_dir: str | Path | None = None,
) -> dict[str, Any]:
    """バックアップを作ってから Excel COM で台帳を書き換える。

    - 削除は下の行から行う（行番号ズレ防止）。追加は最終データ行の直後に B〜F だけ書く。
    - ブックが開かれていて読み取り専用でしか開けない場合は中断する。
    """
    xlsm_path = Path(xlsm_path)
    if not xlsm_path.exists():
        raise FileNotFoundError(f"メール台帳が見つかりません: {xlsm_path}")
    if not additions and not retiree_ids:
        return {"added": 0, "deleted": 0, "backup_path": "", "log_path": ""}

    backup_dir = Path(backup_dir) if backup_dir else xlsm_path.parent / "_backup"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"{xlsm_path.stem}_{stamp}{xlsm_path.suffix}"
    shutil.copy2(xlsm_path, backup_path)

    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(xlsm_path), UpdateLinks=0)
        if workbook.ReadOnly:
            raise RuntimeError(
                "台帳が他の人に開かれていて読み取り専用になっています。"
                "全員が閉じてから再実行してください（変更は加えていません）")
        sheet = workbook.Worksheets(LEDGER_SHEET)
        header_row = _find_header_row(sheet)
        last_row = sheet.Cells(sheet.Rows.Count, 2).End(XL_UP).Row

        delete_targets = []
        retiree_set = {str(item) for item in retiree_ids}
        deleted_rows: list[list[str]] = []
        for row in range(header_row + 1, last_row + 1):
            row_id = str(sheet.Cells(row, 2).Value or "").strip()
            if row_id.endswith(".0"):
                row_id = row_id[:-2]
            if row_id in retiree_set:
                delete_targets.append(row)
                deleted_rows.append([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "削除（退職）",
                    row_id, str(sheet.Cells(row, 3).Value or "").strip(), f"元{row}行目"])
        for row in sorted(delete_targets, reverse=True):
            sheet.Rows(row).Delete()

        last_row = sheet.Cells(sheet.Rows.Count, 2).End(XL_UP).Row
        added_rows: list[list[str]] = []
        for item in additions:
            last_row += 1
            sheet.Cells(last_row, 2).NumberFormat = "@"
            sheet.Cells(last_row, 2).Value = item["id"]
            sheet.Cells(last_row, 3).Value = item["name"]
            if item.get("company_email"):
                sheet.Cells(last_row, 4).Value = item["company_email"]
            if item.get("personal_email"):
                sheet.Cells(last_row, 6).Value = item["personal_email"]
            added_rows.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "追加（新入社員）",
                item["id"], item["name"],
                f"社用={item.get('company_email') or 'なし'} 個人={item.get('personal_email') or 'なし'}"])

        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()

    log_path = Path(log_dir or Config.MAIL_OUTPUT_DIR) / LOG_FILENAME
    _append_log(log_path, deleted_rows + added_rows)
    return {
        "added": len(added_rows),
        "deleted": len(deleted_rows),
        "backup_path": str(backup_path),
        "log_path": str(log_path),
    }
