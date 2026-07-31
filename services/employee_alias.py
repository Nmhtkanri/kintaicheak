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
    "kddi_oyama": "スケジュール氏名エイリアス_KDDI小山.csv",
}

# 対象者リスト。**他社の方が同じ表に載っている系統**だけに置く。
# ファイルがある系統では、ここに載っている人だけを取り込む（載っていない行は捨てる）。
# UAL勤務管理表（KDDI小山）は他社の「小島」が当社の小島さん(2024044)に名前一致して
# しまうため、このリストが無いと他人のスケジュールを投入する事故になる。
ROSTER_CSV_BY_SOURCE = {
    "kddi_oyama": "スケジュール対象者_KDDI小山.csv",
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


def normalize_name_key(name) -> str:
    """氏名の突合キー（前後と内部の空白を除去）"""
    return re.sub(r"[\s　]+", "", str(name or "").strip())


def load_roster(csv_path: str) -> tuple[set[str], dict[str, str]]:
    """対象者リストCSV → (氏名キーの集合, {氏名: 従業員ID})

    従業員ID 列は**任意**。空欄なら氏名だけ登録し、IDは通常の照合に任せる
    （同姓が複数いる場合は画面の候補プルダウンで選ばせる）。

    Raises:
        ValueError: ファイルはあるが文字コード・列構成が読めない
    """
    names: set[str] = set()
    ids: dict[str, str] = {}
    if not csv_path or not os.path.exists(csv_path):
        return names, ids

    for i, row in enumerate(_read_rows(csv_path), start=2):
        if not isinstance(row, dict):
            continue
        name = _pick(row, _NAME_COLUMNS)
        if not name:
            continue
        names.add(normalize_name_key(name))
        emp_id = _pick(row, _ID_COLUMNS)
        if not emp_id:
            continue  # ID未確定でも対象者としては有効
        if not _VALID_EMPLOYEE_ID_RE.fullmatch(emp_id):
            logger.warning(
                "%s の %d 行目: 従業員ID %r は自社社員の形式(20YYNNN)ではないため無視",
                csv_path, i, emp_id,
            )
            continue
        ids[name] = emp_id
    return names, ids


def roster_csv_path(source: str, alias_dir: str) -> str:
    """source に対応する対象者リストCSVのパス（対象外の source なら空文字）"""
    filename = ROSTER_CSV_BY_SOURCE.get(str(source or "").strip().lower())
    if not filename or not alias_dir:
        return ""
    return os.path.join(alias_dir, filename)


def load_roster_for_source(
    source: str, alias_dir: str
) -> tuple[set[str] | None, dict[str, str], str]:
    """source に対応する対象者リストを読む

    Returns:
        (names, ids, warning)
        names   … 対象者の氏名キー集合。**リストを持たない系統では None**（＝絞り込まない）
        ids     … 氏名→従業員ID（リストのID列にあるぶんだけ）
        warning … 読み込みに失敗した場合の説明文（正常時は空文字）
    """
    path = roster_csv_path(source, alias_dir)
    if not path:
        return None, {}, ""
    if not os.path.exists(path):
        # リストを持つべき系統なのにファイルが無い＝他社の人を取り込む危険がある
        return None, {}, (
            f"対象者リストが見つかりません（{os.path.basename(path)}）。"
            "この勤務表には他社の方も含まれるため、当社社員だけに絞り込めていません。"
        )
    try:
        names, ids = load_roster(path)
    except (OSError, ValueError) as e:
        logger.warning("対象者リストの読み込みに失敗 %s: %s", path, e)
        return None, {}, f"対象者リストを読めませんでした（{os.path.basename(path)}）: {e}"
    if not names:
        return None, {}, f"対象者リストが空です（{os.path.basename(path)}）。絞り込みをしていません。"
    logger.info("対象者リストを適用: source=%s %d名 (%s)", source, len(names), path)
    return names, ids, ""


def filter_employees_by_roster(
    employees: list[dict], roster_names: "set[str] | None"
) -> tuple[list[dict], list[str]]:
    """対象者リストに載っている従業員だけを残す

    Returns:
        (kept, excluded_names)  roster_names が None のときは絞り込まない
    """
    if roster_names is None:
        return list(employees or []), []
    kept: list[dict] = []
    excluded: list[str] = []
    for emp in employees or []:
        if not isinstance(emp, dict):
            continue
        name = emp.get("name") or ""
        if normalize_name_key(name) in roster_names:
            kept.append(emp)
        else:
            excluded.append(name or "(名無し)")
    return kept, excluded


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
