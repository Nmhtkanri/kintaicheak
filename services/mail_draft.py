# -*- coding: utf-8 -*-
"""メール下書きモード — 一覧表 × テンプレート × メール台帳から Outlook 下書きを一括作成する。

2026-07-28 決定（Z:\\API連携\\docs\\システム全体マップと機能追加ルール.md 3-1）:
- 作るのは「下書きまで」。直接送信の機能はこのモジュールに存在しない（送信は人が Outlook で行う）。
- 差し込みは {{列名}}。列が無い・値が空欄の人は「要確認」にして下書きを作らない。
- 宛先はメール台帳（B=社員番号 / C=氏名 / D=社用 / E=就業先 / F=個人）と社員番号で突合。
  To=社用優先、就業先・個人は本人BCC。台帳に無い人・氏名相違・アドレス無しは自動で対象外。
- テンプレートは共有フォルダの JSON（Config.MAIL_TEMPLATES_JSON）。exe 再ビルド不要で追加・修正できる。
- Outlook COM (classic Outlook) に触るのは create_drafts() だけ。他は純ロジックでテスト可能。

移植元: Z:\\API連携\\create_paid_leave_outlook_mail.py（有休案内専用ツール）の宛先突合・検証ロジック。
"""
from __future__ import annotations

import csv
import json
import os
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook

from config import Config

ADDRESS_SHEET_NAMES = ("メール送信", "メール送信シート")
ID_HEADER_CANDIDATES = ("社員番号", "従業員番号")
NAME_HEADER_CANDIDATES = ("氏名", "名前", "社員名", "従業員名", "スタッフ名")
ADDRESS_SPLITTER = re.compile(r"[;,\n、]+")
EMAIL_PATTERN = re.compile(r"^[^@\s;]+@[^@\s;]+\.[^@\s;]+$")
PLACEHOLDER_PATTERN = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
IMPORTANCE_LEVELS = {"normal": 1, "high": 2}
STATUS_OK = "OK"
STATUS_NG = "要確認"
LOG_FILENAME = "下書き作成ログ.csv"
LOG_HEADERS = ["処理日時", "社員番号", "氏名", "To", "BCC", "件名", "結果"]


# ---------------------------------------------------------------------------
# 正規化・検証（移植）
# ---------------------------------------------------------------------------

def normalize_employee_id(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_name(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or ""))
    return re.sub(r"[\s\u3000]+", "", text).casefold()


def split_addresses(value: Any) -> tuple[str, ...]:
    parts = [part.strip() for part in ADDRESS_SPLITTER.split(str(value or ""))]
    unique: list[str] = []
    seen: set[str] = set()
    for part in parts:
        if not part:
            continue
        key = part.casefold()
        if key not in seen:
            unique.append(part)
            seen.add(key)
    return tuple(unique)


def invalid_addresses(addresses: Iterable[str]) -> list[str]:
    return [address for address in addresses if not EMAIL_PATTERN.fullmatch(address)]


def value_to_text(value: Any) -> str:
    """差し込み用の文字列化。Excel由来の 3.0 → 3、日付 → 2027年3月31日 に整える。"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return f"{value.year}年{value.month}月{value.day}日"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, float):
        return f"{round(value, 2):g}"
    return str(value).strip()


# ---------------------------------------------------------------------------
# 一覧表（Excel/CSV）の読み込み
# ---------------------------------------------------------------------------

def _rows_from_csv(path: Path) -> list[list[Any]]:
    for encoding in ("utf-8-sig", "cp932"):
        try:
            with open(path, "r", encoding=encoding, newline="") as f:
                return [row for row in csv.reader(f)]
        except UnicodeDecodeError:
            continue
    raise ValueError(f"CSVの文字コードを判定できません（UTF-8/Shift-JIS以外）: {path}")


def _rows_from_excel(path: Path) -> list[list[Any]]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook.worksheets[0]
        return [list(row) for row in sheet.iter_rows(values_only=True)]
    finally:
        workbook.close()


def load_recipients_table(path: str | Path) -> tuple[list[str], list[dict[str, Any]], str, str | None]:
    """一覧表を読み、(ヘッダー一覧, 行dictのリスト, 社員番号列名, 氏名列名orNone) を返す。

    ヘッダー行は先頭30行から「社員番号 / 従業員番号」を含む行を探す。
    同名ヘッダーが複数ある場合は右側の列が勝つ（差し込みには一意な列名を推奨）。
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"一覧表が見つかりません: {path}")
    suffix = path.suffix.lower()
    if suffix == ".csv":
        raw_rows = _rows_from_csv(path)
    elif suffix in (".xlsx", ".xlsm"):
        raw_rows = _rows_from_excel(path)
    else:
        raise ValueError(f"一覧表は .xlsx / .xlsm / .csv を指定してください: {path.name}")

    header_index = None
    for index, row in enumerate(raw_rows[:30]):
        cells = {str(cell or "").strip() for cell in row}
        if cells & set(ID_HEADER_CANDIDATES):
            header_index = index
            break
    if header_index is None:
        raise ValueError(
            "一覧表に「社員番号」（または従業員番号）のヘッダー行が見つかりません（先頭30行を確認）")

    header_row = raw_rows[header_index]
    headers = [str(cell or "").strip() for cell in header_row]
    id_key = next(h for h in headers if h in ID_HEADER_CANDIDATES)
    name_key = next((h for h in headers if h in NAME_HEADER_CANDIDATES), None)

    rows: list[dict[str, Any]] = []
    for raw in raw_rows[header_index + 1:]:
        row_map: dict[str, Any] = {}
        for column, header in enumerate(headers):
            if not header:
                continue
            row_map[header] = raw[column] if column < len(raw) else None
        if normalize_employee_id(row_map.get(id_key)):
            rows.append(row_map)
    if not rows:
        raise ValueError("一覧表にデータ行がありません（社員番号が入った行が必要です）")
    return [h for h in headers if h], rows, id_key, name_key


# ---------------------------------------------------------------------------
# メール台帳（移植: B=社員番号 / C=氏名 / D=社用 / E=就業先 / F=個人）
# ---------------------------------------------------------------------------

def load_address_book(path: str | Path) -> dict[str, list[dict[str, Any]]]:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"メール台帳が見つかりません: {path}")
    workbook = load_workbook(path, data_only=True, read_only=False)
    try:
        sheet_name = next((name for name in ADDRESS_SHEET_NAMES if name in workbook.sheetnames), None)
        if sheet_name is None:
            raise ValueError(
                f"メール台帳に「{' / '.join(ADDRESS_SHEET_NAMES)}」シートがありません")
        sheet = workbook[sheet_name]
        header_row = None
        for row_number in range(1, min(sheet.max_row, 30) + 1):
            if str(sheet.cell(row_number, 2).value or "").strip() in ID_HEADER_CANDIDATES:
                header_row = row_number
                break
        if header_row is None:
            raise ValueError("メール台帳の社員番号ヘッダー（B列）を特定できません")

        entries: dict[str, list[dict[str, Any]]] = {}
        for row_number in range(header_row + 1, sheet.max_row + 1):
            employee_id = normalize_employee_id(sheet.cell(row_number, 2).value)
            if not employee_id:
                continue
            entries.setdefault(employee_id, []).append({
                "employee_id": employee_id,
                "name": str(sheet.cell(row_number, 3).value or "").strip(),
                "company": split_addresses(sheet.cell(row_number, 4).value),
                "client": split_addresses(sheet.cell(row_number, 5).value),
                "personal": split_addresses(sheet.cell(row_number, 6).value),
            })
        return entries
    finally:
        workbook.close()


def _address_signature(entry: dict[str, Any]) -> tuple:
    return (entry["company"], entry["client"], entry["personal"])


def _deduplicate_addresses(*groups: Iterable[str]) -> tuple[str, ...]:
    result: list[str] = []
    seen: set[str] = set()
    for group in groups:
        for address in group:
            key = address.casefold()
            if key not in seen:
                result.append(address)
                seen.add(key)
    return tuple(result)


def select_recipients(entry: dict[str, Any]) -> tuple[tuple[str, ...], tuple[str, ...], str]:
    """To=社用優先（無ければ就業先→個人）、残りは本人BCC。内訳ラベルも返す。"""
    if entry["company"]:
        to_addresses, to_label = entry["company"], "社用"
        other_groups = (entry["client"], entry["personal"])
    elif entry["client"]:
        to_addresses, to_label = entry["client"], "就業先"
        other_groups = (entry["personal"],)
    else:
        to_addresses, to_label = entry["personal"], "個人"
        other_groups = ()

    bcc_addresses = _deduplicate_addresses(*other_groups)
    to_keys = {address.casefold() for address in to_addresses}
    bcc_addresses = tuple(a for a in bcc_addresses if a.casefold() not in to_keys)
    bcc_keys = {address.casefold() for address in bcc_addresses}
    bcc_labels = []
    if any(a.casefold() in bcc_keys for a in entry["client"]):
        bcc_labels.append("就業先")
    if any(a.casefold() in bcc_keys for a in entry["personal"]):
        bcc_labels.append("個人")
    breakdown = f"To:{to_label}"
    if bcc_labels:
        breakdown += " / BCC:" + "・".join(bcc_labels)
    return to_addresses, bcc_addresses, breakdown


# ---------------------------------------------------------------------------
# 差し込み
# ---------------------------------------------------------------------------

def extract_placeholders(*texts: str) -> list[str]:
    found: list[str] = []
    for text in texts:
        for match in PLACEHOLDER_PATTERN.finditer(text or ""):
            key = match.group(1).strip()
            if key not in found:
                found.append(key)
    return found


def render_template_text(text: str, row: dict[str, Any], headers: set[str]) -> tuple[str, list[str]]:
    """{{列名}} を row の値で置換する。列なし・空欄は missing に列挙し、置換しない。"""
    missing: list[str] = []

    def _substitute(match: re.Match) -> str:
        key = match.group(1).strip()
        if key not in headers:
            missing.append(f"{{{{{key}}}}}（列がありません）")
            return match.group(0)
        rendered = value_to_text(row.get(key))
        if rendered == "":
            missing.append(f"{{{{{key}}}}}（空欄）")
        return rendered

    return PLACEHOLDER_PATTERN.sub(_substitute, text or ""), missing


# ---------------------------------------------------------------------------
# 計画づくり（純ロジックの中心）
# ---------------------------------------------------------------------------

def build_mail_plans(
    rows: list[dict[str, Any]],
    id_key: str,
    name_key: str | None,
    headers: list[str],
    address_book: dict[str, list[dict[str, Any]]],
    template: dict[str, Any],
) -> list[dict[str, Any]]:
    importance = str(template.get("importance") or "normal")
    if importance not in IMPORTANCE_LEVELS:
        raise ValueError(f"重要度は normal / high を指定してください: {importance}")
    # to_only: Toだけに送る（既定・2026-07-29に谷津さん指定で変更） / bcc: 就業先・個人を本人BCCに入れる
    bcc_mode = str(template.get("bcc_mode") or "to_only")
    if bcc_mode not in ("bcc", "to_only"):
        bcc_mode = "to_only"
    cc_addresses = split_addresses(template.get("cc"))
    bad_cc = invalid_addresses(cc_addresses)
    if bad_cc:
        raise ValueError(f"CCのアドレス形式が不正です: {', '.join(bad_cc)}")

    header_set = set(headers)
    plans: list[dict[str, Any]] = []
    for row in rows:
        employee_id = normalize_employee_id(row.get(id_key))
        table_name = value_to_text(row.get(name_key)) if name_key else ""
        problems: list[str] = []
        notes: list[str] = []

        entries = address_book.get(employee_id, [])
        entry = None
        if not entries:
            problems.append("社員番号に対応するメール台帳行がありません")
        elif len(entries) == 1:
            entry = entries[0]
        else:
            signatures = {_address_signature(item) for item in entries}
            if len(signatures) > 1:
                problems.append("同じ社員番号に異なるメールアドレスが複数あります")
            else:
                entry = entries[0]
                notes.append("メール台帳に同一社員番号の重複行があります")

        to_addresses: tuple[str, ...] = ()
        bcc_addresses: tuple[str, ...] = ()
        breakdown = ""
        if entry is not None:
            if table_name and entry["name"] and normalize_name(table_name) != normalize_name(entry["name"]):
                problems.append(f"氏名相違（一覧表: {table_name} / 台帳: {entry['name']}）")
            all_addresses = _deduplicate_addresses(
                entry["company"], entry["client"], entry["personal"])
            if invalid_addresses(all_addresses):
                problems.append("メールアドレス形式エラー")
            if not all_addresses:
                problems.append("利用できるメールアドレスがありません")
            elif not problems:
                to_addresses, bcc_addresses, breakdown = select_recipients(entry)
                if bcc_mode == "to_only":
                    bcc_addresses = ()
                    breakdown = breakdown.split(" / BCC:")[0]

        subject, missing_subject = render_template_text(template.get("subject", ""), row, header_set)
        body, missing_body = render_template_text(template.get("body", ""), row, header_set)
        for item in missing_subject + missing_body:
            message = f"差し込みできません: {item}"
            if message not in problems:
                problems.append(message)

        plans.append({
            "employee_id": employee_id,
            "name": table_name or (entry["name"] if entry else ""),
            "to": list(to_addresses),
            "bcc": list(bcc_addresses),
            "cc": list(cc_addresses),
            "breakdown": breakdown,
            "subject": subject,
            "body": body,
            "importance": IMPORTANCE_LEVELS[importance],
            "status": STATUS_NG if problems else STATUS_OK,
            "issues": problems + notes,
        })
    return plans


def build_plans_for(
    table_path: str | Path,
    address_book_path: str | Path,
    template: dict[str, Any],
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """一覧表と台帳を読み、(plans, meta) を返す。ルートからの入口。"""
    headers, rows, id_key, name_key = load_recipients_table(table_path)
    address_book = load_address_book(address_book_path)
    plans = build_mail_plans(rows, id_key, name_key, headers, address_book, template)
    ok_count = sum(plan["status"] == STATUS_OK for plan in plans)
    meta = {
        "columns": headers,
        "id_column": id_key,
        "name_column": name_key,
        "placeholders": extract_placeholders(template.get("subject", ""), template.get("body", "")),
        "counts": {"total": len(plans), "ok": ok_count, "warn": len(plans) - ok_count},
    }
    return plans, meta


# ---------------------------------------------------------------------------
# テンプレート（共有フォルダの JSON・再ビルド不要）
# ---------------------------------------------------------------------------

def load_templates(path: str | Path) -> list[dict[str, Any]]:
    path = Path(path)
    if not path.exists():
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as exc:
        raise ValueError(f"テンプレートJSONを読み取れません（{path}）: {exc}") from exc
    templates = data.get("templates") if isinstance(data, dict) else None
    if not isinstance(templates, list):
        raise ValueError(f"テンプレートJSONの形式が不正です（templates配列が必要）: {path}")
    return [t for t in templates if isinstance(t, dict) and str(t.get("name", "")).strip()]


def save_template(path: str | Path, template: dict[str, Any], *, delete: bool = False) -> list[dict[str, Any]]:
    path = Path(path)
    name = str(template.get("name", "")).strip()
    if not name:
        raise ValueError("テンプレート名を入力してください")
    templates = load_templates(path) if path.exists() else []
    templates = [t for t in templates if str(t.get("name", "")).strip() != name]
    if not delete:
        record = {
            "name": name,
            "subject": str(template.get("subject", "")),
            "body": str(template.get("body", "")),
            "cc": str(template.get("cc", "")).strip(),
            "bcc_mode": str(template.get("bcc_mode", "to_only")),
            "importance": str(template.get("importance", "normal")),
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        if record["importance"] not in IMPORTANCE_LEVELS:
            record["importance"] = "normal"
        if record["bcc_mode"] not in ("bcc", "to_only"):
            record["bcc_mode"] = "bcc"
        templates.append(record)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"version": 1, "templates": templates}
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)
    return templates


# ---------------------------------------------------------------------------
# Outlook 下書き作成（COM に触るのはここだけ。送信機能は存在しない）
# ---------------------------------------------------------------------------

class OutlookMailer:
    def __init__(self) -> None:
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "Outlook連携に必要な pywin32 がありません（pip install pywin32）") from exc
        try:
            self.application = win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:
            raise RuntimeError(f"Outlook を起動できません（classic Outlook が必要）: {exc}") from exc

    def create_draft(self, *, to: str, cc: str, bcc: str, subject: str, body: str, importance: int) -> None:
        mail = self.application.CreateItem(0)
        mail.To = to
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        mail.Subject = subject
        mail.Body = body
        mail.Importance = importance
        mail.Save()


def _append_log(log_path: Path, rows: list[list[str]]) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    is_new = not log_path.exists()
    # 新規作成時だけ BOM 付き UTF-8（谷津さんの Excel が BOM 無しを文字化けさせるため）。
    encoding = "utf-8-sig" if is_new else "utf-8"
    with open(log_path, "a", encoding=encoding, newline="") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(LOG_HEADERS)
        writer.writerows(rows)


def create_drafts(
    plans: list[dict[str, Any]],
    *,
    only_ids: list[str],
    log_dir: str | Path | None = None,
    mailer: Any | None = None,
) -> dict[str, Any]:
    """選択された OK 行だけ Outlook 下書きを作成し、ログを残す。

    - status が要確認の行は選択されていても作らない（skipped に数える）。
    - 3件連続で失敗したら中断する（Outlook 側の異常を疑う）。
    """
    selected = {str(item) for item in only_ids}
    targets = [p for p in plans if p["employee_id"] in selected and p["status"] == STATUS_OK]
    skipped = len([p for p in plans if p["employee_id"] in selected]) - len(targets)
    if not targets:
        return {"processed": 0, "skipped": skipped, "failed": 0, "results": [], "log_path": ""}

    log_dir = Path(log_dir or Config.MAIL_OUTPUT_DIR)
    log_path = log_dir / LOG_FILENAME

    own_com = mailer is None
    if own_com:
        try:
            import pythoncom  # type: ignore
            pythoncom.CoInitialize()
        except ImportError:
            pythoncom = None
    try:
        mailer = mailer or OutlookMailer()
        processed = 0
        failed = 0
        consecutive_failures = 0
        results: list[dict[str, str]] = []
        log_rows: list[list[str]] = []
        for plan in targets:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            to = "; ".join(plan["to"])
            bcc = "; ".join(plan["bcc"])
            try:
                mailer.create_draft(
                    to=to,
                    cc="; ".join(plan["cc"]),
                    bcc=bcc,
                    subject=plan["subject"],
                    body=plan["body"],
                    importance=plan["importance"],
                )
                result_text = "下書き作成済"
                processed += 1
                consecutive_failures = 0
            except Exception as exc:
                result_text = f"エラー: {exc}"[:250]
                failed += 1
                consecutive_failures += 1
            results.append({"employee_id": plan["employee_id"], "name": plan["name"],
                            "result": result_text})
            log_rows.append([now, plan["employee_id"], plan["name"], to, bcc,
                             plan["subject"], result_text])
            if consecutive_failures >= 3:
                _append_log(log_path, log_rows)
                raise RuntimeError(
                    "Outlook の下書き作成が3件連続で失敗したため中断しました。"
                    f"ログを確認してください: {log_path}")
        _append_log(log_path, log_rows)
        return {"processed": processed, "skipped": skipped, "failed": failed,
                "results": results, "log_path": str(log_path)}
    finally:
        if own_com:
            try:
                import pythoncom  # type: ignore
                pythoncom.CoUninitialize()
            except Exception:
                pass
