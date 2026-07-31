"""シフト表の氏名 → jinjer 従業員ID の読み替え表（エイリアス）

同姓が複数いる氏名（例: 吉田 英伸 / 吉田 拓矢）は、事故防止のため
`build_name_to_id_map` が**わざと登録しない**。そのためシフト表が姓だけで
氏名を持っている場合、その人は毎回 ID 未解決＝「未分類」CSV に落ちる。

現場ごとに「この表の"吉田"は誰か」は決まっているので、**シフト表の系統
（source）ごと**に読み替え表を外部CSVで持つ。全社共通にはしない
（別の現場の"吉田"が別人でも壊れないようにするため）。

CSV フォーマット（1行目はヘッダー / BOM付きUTF-8・CP932 どちらも可）:

    シフト表氏名,従業員ID,備考
    吉田,2025007,KDX 夜勤シフト表の「吉田」は吉田 拓矢さん

置き場所は共有フォルダ（Config.SCHEDULE_NAME_ALIAS_DIR）。経理モードの
マッピング表と同じ方式で、直せば exe 再ビルドなしに次回実行から効く。
"""

from __future__ import annotations

import csv
import logging
import os
import re

logger = logging.getLogger(__name__)

# source（シフト表の系統）→ CSV ファイル名
ALIAS_CSV_BY_SOURCE = {
    "kdx": "スケジュール氏名エイリアス_KDX.csv",
}

_NAME_COLUMNS = ("シフト表氏名", "氏名", "名前")
_ID_COLUMNS = ("従業員ID", "社員番号", "ID")

# 自社社員の社員番号は 20YY 始まり。5/6/9 始まり（派遣・テスト番号）は
# 給与計算の対象外なのでスケジュール投入先としても認めない。
_VALID_EMPLOYEE_ID_RE = re.compile(r"^20\d{5}$")


def _read_rows(csv_path: str) -> list[dict]:
    """BOM付きUTF-8 → CP932 の順で読む（Excel で保存し直しても壊れないように）"""
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "cp932"):
        try:
            with open(csv_path, encoding=encoding, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError as e:
            last_error = e
            continue
    raise ValueError(f"文字コードを判定できません（UTF-8/CP932 で読めません）: {last_error}")


def _pick(row: dict, columns) -> str:
    for col in columns:
        if col in row and row[col] is not None:
            v = str(row[col]).strip()
            if v:
                return v
    return ""


def load_employee_aliases(csv_path: str) -> dict[str, str]:
    """エイリアスCSV → {シフト表氏名: 従業員ID}

    ファイルが無い場合は空辞書（エイリアス無しで従来通り動く）。
    行が壊れている場合はその行だけ捨てて warning を出す（全体は止めない）。

    Raises:
        ValueError: ファイルはあるが文字コード・列構成が読めない
    """
    if not csv_path or not os.path.exists(csv_path):
        return {}

    rows = _read_rows(csv_path)
    aliases: dict[str, str] = {}
    for i, row in enumerate(rows, start=2):  # 1行目はヘッダー
        if not isinstance(row, dict):
            continue
        name = _pick(row, _NAME_COLUMNS)
        emp_id = _pick(row, _ID_COLUMNS)
        if not name and not emp_id:
            continue  # 空行
        if not name or not emp_id:
            logger.warning("%s の %d 行目: 氏名または従業員IDが空のためスキップ", csv_path, i)
            continue
        if not _VALID_EMPLOYEE_ID_RE.fullmatch(emp_id):
            logger.warning(
                "%s の %d 行目: 従業員ID %r は自社社員の形式(20YYNNN)ではないためスキップ",
                csv_path, i, emp_id,
            )
            continue
        if name in aliases and aliases[name] != emp_id:
            logger.warning(
                "%s の %d 行目: 氏名 %r が重複（%s → %s で上書き）",
                csv_path, i, name, aliases[name], emp_id,
            )
        aliases[name] = emp_id
    return aliases


def alias_csv_path(source: str, alias_dir: str) -> str:
    """source に対応するエイリアスCSVのパス（対象外の source なら空文字）"""
    filename = ALIAS_CSV_BY_SOURCE.get(str(source or "").strip().lower())
    if not filename or not alias_dir:
        return ""
    return os.path.join(alias_dir, filename)


def load_aliases_for_source(source: str, alias_dir: str) -> tuple[dict[str, str], str]:
    """source（"kdx" 等）に対応するエイリアスを読む

    Returns:
        (aliases, warning)  warning は読み込みに失敗した場合の説明文（正常時は空文字）
    """
    path = alias_csv_path(source, alias_dir)
    if not path:
        return {}, ""
    try:
        aliases = load_employee_aliases(path)
    except (OSError, ValueError) as e:
        logger.warning("氏名エイリアス表の読み込みに失敗 %s: %s", path, e)
        return {}, f"氏名エイリアス表を読めませんでした（{os.path.basename(path)}）: {e}"
    if aliases:
        logger.info("氏名エイリアス表を適用: source=%s %d件 (%s)", source, len(aliases), path)
    return aliases, ""


def apply_aliases(name_to_id: dict[str, str], aliases: dict[str, str]) -> dict[str, str]:
    """氏名→IDマップにエイリアスを重ねた**新しい辞書**を返す（元は書き換えない）

    エイリアスを後勝ちにすることで、同姓のため未登録だった姓（"吉田"）を
    明示指定の従業員IDへ確定させる。
    """
    if not aliases:
        return name_to_id
    merged = dict(name_to_id or {})
    merged.update(aliases)
    return merged
