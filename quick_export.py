"""quick_export.py — 5月本番 MVP jinjer アップロード用 全件CSV生成

入力:
  --diff       : quick_compare.py が出力した 差異一覧xlsx（人間判断埋め込み済み）
  --jinjer-dir : 元の jinjer ダウンロードCSV 群のフォルダ（CP932, 194列）
  --output     : 出力する jinjer アップロード用CSV のパス（CP932）

処理:
  1. 差異一覧の「人間判断=承認」行を集める
  2. 元 jinjer CSV を全件読み（行順保持）
  3. 承認行に応じて該当セルを上書き:
     - 出勤差異 → 出勤1（22列目）
     - 退勤差異 → 退勤1（23列目）
     - 休憩差異 → 手入力休憩1 / 手入力復帰1 / 手入力休憩時間 があれば該当列へ反映
     - 総労働時間差異 → 警告のみ、上書きしない
  4. 全件をそのままアップロード用CSV として書き出す

実績確定済みの扱い:
  実績確定済（180列=TRUE）は「本人が打刻申請を確定した状態」を示すだけで、
  勤怠が正しいことを保証しない。原則として請求勤怠の値が正であり、管理部の
  「人間判断=承認」があれば実績確定済でも上書きする。
  実績確定済の件数は参考表示としてサマリに出すが、上書きの可否判断には使わない。

dry-run（既定）: 承認件数・上書き予定件数を表示するだけ
--execute      : CSV を書き出す
"""

from __future__ import annotations

import argparse
import csv
import io
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, date, time
from pathlib import Path
from typing import Any

import pandas as pd


# ===== jinjer CSV の上書き対象列（ヘッダー名で照合） =====
JINJER_COL_PUNCH_IN = "出勤1"
JINJER_COL_PUNCH_OUT = "退勤1"
JINJER_COL_BREAK_1_START = "休憩1"
JINJER_COL_BREAK_1_END = "復帰1"
JINJER_COL_BREAK_TOTAL = "休憩時間"
JINJER_COL_EMP_ID = "*従業員ID"
JINJER_COL_DATE = "*年月日"

# 差異種別
DIFF_KIND_PUNCH_IN = "出勤"
DIFF_KIND_PUNCH_OUT = "退勤"
DIFF_KIND_BREAK = "休憩"
DIFF_KIND_TOTAL = "総労働時間"

JUDGE_APPROVE = "承認"
JUDGE_REJECT = "却下"
JUDGE_HOLD = "保留"
JUDGMENTS = {JUDGE_APPROVE, JUDGE_REJECT, JUDGE_HOLD}

# 「人間判断」を入力すべきなのに、誤って判断（承認/却下/保留）が書き込まれやすい列。
# 本来これらの列には時刻・数値・メモが入るため、判断キーワードが入っていたら
# 「人間判断」列の入力ミスとみなして回収する（先頭ほど優先）。
MISPLACED_JUDGE_COLS = [
    "打刻修正", "手入力修正値", "手入力休憩1", "手入力復帰1", "手入力休憩時間",
    "自動修正提案値", "判断メモ",
]


# ----------------------------------------------------------------------
# データクラス
# ----------------------------------------------------------------------

@dataclass
class ApprovedRow:
    emp_id: str
    target_date_iso: str  # YYYY-MM-DD
    kind: str
    auto_fix_value: str
    manual_fix_value: str
    manual_break_start: str
    manual_break_end: str
    manual_break_total: str
    name: str  # 表示用
    warn_level: str
    source_diff_row_id: int


@dataclass
class Stats:
    total_diff_rows: int = 0
    approved: int = 0
    rejected: int = 0
    held: int = 0
    pending: int = 0
    # 差異種別ごとの承認内訳
    approved_punch_in: int = 0
    approved_punch_out: int = 0
    approved_break: int = 0
    approved_total: int = 0
    # 実際の上書き処理結果
    overwritten_punch_in: int = 0
    overwritten_punch_out: int = 0
    skipped_break: int = 0  # 承認されたが上書きしなかった
    skipped_total: int = 0
    overwritten_break_start: int = 0
    overwritten_break_end: int = 0
    overwritten_break_total: int = 0
    not_matched: int = 0  # jinjer 行が見つからない承認行
    overwritten_finalized: int = 0  # 上書きしたうち実績確定済みだった件数（参考表示）
    recovered_misplaced: int = 0  # 「人間判断」列以外に入っていた判断を回収した件数
    misplaced_by_col: dict[str, int] = field(default_factory=dict)  # 列名→誤入力件数
    warnings: list[str] = field(default_factory=list)


# ----------------------------------------------------------------------
# ユーティリティ
# ----------------------------------------------------------------------

def normalize_date_iso(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    s = str(value).strip()
    if not s or s.lower() in ("nan", "nat"):
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date().isoformat()
    except ValueError:
        pass
    try:
        return datetime.strptime(s, "%Y/%m/%d").date().isoformat()
    except ValueError:
        pass
    return None


def iso_to_jinjer_date(date_iso: str) -> str:
    """'2026-05-12' → '2026/5/12'（jinjer の 0パディングなし形式）"""
    d = datetime.strptime(date_iso, "%Y-%m-%d").date()
    return f"{d.year}/{d.month}/{d.day}"


_HHMM_RE = re.compile(r"^(\d{1,3}):(\d{2})")


def _hhmm_to_minutes(value: Any) -> int | None:
    """'H:MM' / 'HH:MM'（24時超表記含む）を分に変換。不正値は None。"""
    if value is None:
        return None
    m = _HHMM_RE.match(str(value).strip())
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def to_overnight_punch_out(punch_in: Any, punch_out: str) -> str:
    """退勤が翌朝に跨ぐ場合、jinjer インポート用の 24時超表記へ変換する。

    出勤 > 退勤（夜勤で翌日にずれ込み）のときだけ退勤に 24h を足す（08:15→32:15）。
    すでに 24時超表記なら out >= in となり再変換しない（冪等）。出退勤の判定が
    できないときは元の退勤値をそのまま返す。
    """
    if not punch_out:
        return punch_out
    in_min = _hhmm_to_minutes(punch_in)
    out_min = _hhmm_to_minutes(punch_out)
    if in_min is None or out_min is None or out_min >= in_min:
        return punch_out
    adjusted = out_min + 24 * 60
    return f"{adjusted // 60:02d}:{adjusted % 60:02d}"


def clean_excel_text(value: Any) -> str:
    """Excel由来の値を、jinjer汎用データCSVへ書く文字列に正規化する。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return ""
    return s


# ----------------------------------------------------------------------
# 入力読み込み
# ----------------------------------------------------------------------

def load_approved_rows(diff_xlsx: Path, stats: Stats) -> list[ApprovedRow]:
    """差異一覧xlsx の「差異一覧」シートから 人間判断=承認 の行を抽出。"""
    if not diff_xlsx.exists():
        raise FileNotFoundError(f"差異一覧xlsx が見つかりません: {diff_xlsx}")

    xl = pd.ExcelFile(diff_xlsx)
    if "差異一覧" not in xl.sheet_names:
        if "突合結果" in xl.sheet_names:
            raise ValueError(
                "選択されたファイルは手順1の勤怠突合結果xlsxです。"
                "手順3には、手順2で生成して人間判断を入力した差異一覧xlsxを指定してください。"
            )
        sheets = "、".join(xl.sheet_names)
        raise ValueError(
            "差異一覧シートが見つかりません。"
            "手順3には、手順2で生成して人間判断を入力した差異一覧xlsxを指定してください。"
            f" このファイルのシート: {sheets}"
        )

    df = pd.read_excel(xl, sheet_name="差異一覧", dtype=object)
    stats.total_diff_rows = len(df)

    approved: list[ApprovedRow] = []
    for _, row in df.iterrows():
        judge = str(row.get("人間判断") or "").strip()

        # 判断が「人間判断」列に無い場合、別列への入力ミスを回収する。
        # 承認/却下/保留 は時刻・メモには本来現れない閉じた語彙なので、
        # 他の列に厳密一致で入っていれば判断の置き場所間違いとみなして拾う。
        misplaced_col: str | None = None
        if judge not in JUDGMENTS:
            for col in MISPLACED_JUDGE_COLS:
                v = str(row.get(col) or "").strip()
                if v in JUDGMENTS:
                    judge = v
                    misplaced_col = col
                    stats.recovered_misplaced += 1
                    stats.misplaced_by_col[col] = stats.misplaced_by_col.get(col, 0) + 1
                    break

        if judge == JUDGE_APPROVE:
            stats.approved += 1
        elif judge == JUDGE_REJECT:
            stats.rejected += 1
            continue
        elif judge == JUDGE_HOLD:
            stats.held += 1
            continue
        else:
            stats.pending += 1
            continue

        # 判断を回収した列の値は「承認」等の文字列なので、時刻として書き込まないよう除外する。
        def _field(col: str) -> str:
            if col == misplaced_col:
                return ""
            v = clean_excel_text(row.get(col))
            # 念のため、どの入力欄にも判断キーワードが時刻として紛れ込まないよう除外
            return "" if v in JUDGMENTS else v

        emp_id = str(row.get("従業員ID") or "").strip()
        date_iso = normalize_date_iso(row.get("対象日付"))
        kind = str(row.get("差異種別") or "").strip()
        auto_fix = _field("自動修正提案値")
        # 列名は「打刻修正」（新）。旧フォーマットの「手入力修正値」も後方互換で読む。
        manual_fix = _field("打刻修正") or _field("手入力修正値")
        manual_break_start = _field("手入力休憩1")
        manual_break_end = _field("手入力復帰1")
        manual_break_total = _field("手入力休憩時間")
        name = str(row.get("氏名") or "").strip()
        warn_level = str(row.get("警告レベル") or "").strip()
        row_id_raw = row.get("行ID")
        try:
            row_id = int(row_id_raw) if row_id_raw is not None else 0
        except (ValueError, TypeError):
            row_id = 0

        if not emp_id or not date_iso or not kind:
            stats.warnings.append(
                f"行ID={row_id_raw} 承認だが必須フィールド欠落 (emp_id={emp_id}, date={date_iso}, kind={kind})"
            )
            continue

        approved.append(ApprovedRow(
            emp_id=emp_id, target_date_iso=date_iso, kind=kind,
            auto_fix_value=auto_fix,
            manual_fix_value=manual_fix,
            manual_break_start=manual_break_start,
            manual_break_end=manual_break_end,
            manual_break_total=manual_break_total,
            name=name, warn_level=warn_level,
            source_diff_row_id=row_id,
        ))

        # 種別別集計
        if kind == DIFF_KIND_PUNCH_IN:
            stats.approved_punch_in += 1
        elif kind == DIFF_KIND_PUNCH_OUT:
            stats.approved_punch_out += 1
        elif kind == DIFF_KIND_BREAK:
            stats.approved_break += 1
        elif kind == DIFF_KIND_TOTAL:
            stats.approved_total += 1

    # 判断が「人間判断」列以外に入っていたものを回収した場合、目立つ警告を残す。
    if stats.recovered_misplaced:
        detail = "、".join(
            f"『{col}』列 {cnt}件" for col, cnt in stats.misplaced_by_col.items()
        )
        stats.warnings.insert(
            0,
            f"⚠️ 判断（承認/却下/保留）が「人間判断」列ではなく {detail} に入力されていました。"
            f"合計 {stats.recovered_misplaced} 件を自動で回収して処理しました。"
            "次回は必ず「人間判断」列（プルダウンのある列）に入力してください。",
        )

    return approved


def _csv_input_files(path: Path) -> list[Path]:
    if path.is_file():
        return [path]
    return sorted([p for p in path.glob("*.csv") if not p.name.startswith("~$")])


def load_jinjer_csvs(jinjer_dir: Path) -> tuple[list[str], list[list[str]]]:
    """jinjer ダウンロードCSV を全件 CP932 で読む。
    複数ファイルがある場合は縦結合（ヘッダーは先頭ファイルのものを採用、列数一致を要求）。
    CSVファイル単体の指定にも対応する。

    Returns:
        (headers, rows) — rows は 1行 = list[str]
    """
    if not jinjer_dir.exists():
        raise FileNotFoundError(f"jinjer CSVまたはフォルダが見つかりません: {jinjer_dir}")

    csv_paths = _csv_input_files(jinjer_dir)
    if not csv_paths:
        raise RuntimeError(f"jinjer CSV が見つかりません: {jinjer_dir}")

    all_headers: list[str] | None = None
    all_rows: list[list[str]] = []
    for path in csv_paths:
        with open(path, encoding="cp932", newline="") as f:
            reader = csv.reader(f)
            headers = next(reader)
            if all_headers is None:
                all_headers = headers
            else:
                if len(headers) != len(all_headers):
                    raise RuntimeError(
                        f"jinjer CSV の列数が一致しません: {path.name} は {len(headers)} 列、"
                        f"先頭ファイルは {len(all_headers)} 列"
                    )
            for row in reader:
                # 列数を揃える（足りなければ空文字で埋める）
                if len(row) < len(all_headers):
                    row = row + [""] * (len(all_headers) - len(row))
                elif len(row) > len(all_headers):
                    row = row[: len(all_headers)]
                all_rows.append(row)
        print(f"[info] jinjer CSV 読込: {path.name} (累計 {len(all_rows)} 行)")

    assert all_headers is not None
    return all_headers, all_rows


def build_jinjer_row_index(
    headers: list[str], rows: list[list[str]]
) -> dict[tuple[str, str], int]:
    """(従業員ID, 日付ISO) → rows のインデックス"""
    try:
        emp_col_idx = headers.index(JINJER_COL_EMP_ID)
        date_col_idx = headers.index(JINJER_COL_DATE)
    except ValueError as e:
        raise RuntimeError(f"jinjer CSV に必須列が見つかりません: {e}")

    index: dict[tuple[str, str], int] = {}
    duplicates: list[tuple[str, str]] = []
    for i, row in enumerate(rows):
        emp_id = (row[emp_col_idx] or "").strip()
        date_iso = normalize_date_iso(row[date_col_idx])
        if not emp_id or not date_iso:
            continue
        key = (emp_id, date_iso)
        if key in index:
            duplicates.append(key)
        index[key] = i
    if duplicates:
        print(f"[warn] (従業員ID, 年月日) の重複が {len(duplicates)} 件ありました（最後の行を採用）")
    return index


# ----------------------------------------------------------------------
# CSV 上書き
# ----------------------------------------------------------------------

def apply_approved_rows(
    headers: list[str],
    rows: list[list[str]],
    row_index: dict[tuple[str, str], int],
    approved: list[ApprovedRow],
    stats: Stats,
) -> None:
    """承認行に応じて rows を in-place で上書きする。

    人間判断=承認 のみが上書きの条件。実績確定状況は上書き可否の判定材料にしない
    （実績確定済 = 本人が打刻申請を確定しただけで、勤怠が正しいことを保証しない。
    原則として請求勤怠が正しいため、管理部の判断＝承認なら実績確定済でも上書きする）。
    """
    required_cols = [JINJER_COL_PUNCH_IN, JINJER_COL_PUNCH_OUT]
    missing = [col for col in required_cols if col not in headers]
    if missing:
        raise RuntimeError(f"jinjer CSV に上書き対象列がありません: {', '.join(missing)}")
    punch_in_col = headers.index(JINJER_COL_PUNCH_IN)
    punch_out_col = headers.index(JINJER_COL_PUNCH_OUT)
    break_start_col = headers.index(JINJER_COL_BREAK_1_START) if JINJER_COL_BREAK_1_START in headers else None
    break_end_col = headers.index(JINJER_COL_BREAK_1_END) if JINJER_COL_BREAK_1_END in headers else None
    break_total_col = headers.index(JINJER_COL_BREAK_TOTAL) if JINJER_COL_BREAK_TOTAL in headers else None

    # 実績確定状況列があれば、上書きしたうち何件が実績確定済だったかを集計する（参考表示のみ）
    finalized_col = headers.index("実績確定状況") if "実績確定状況" in headers else None

    for app in approved:
        # 休憩は請求勤怠から自動推定せず、人間が差異一覧に入力した欄だけ反映する。
        if app.kind == DIFF_KIND_BREAK:
            key = (app.emp_id, app.target_date_iso)
            idx = row_index.get(key)
            if idx is None:
                stats.not_matched += 1
                stats.warnings.append(
                    f"行ID={app.source_diff_row_id} jinjer CSV に該当行なし "
                    f"(emp={app.emp_id} date={app.target_date_iso} {app.name})"
                )
                continue

            wrote = False
            if app.manual_break_start:
                if break_start_col is None:
                    stats.warnings.append(f"行ID={app.source_diff_row_id} 汎用データに '{JINJER_COL_BREAK_1_START}' 列がありません")
                else:
                    rows[idx][break_start_col] = app.manual_break_start
                    stats.overwritten_break_start += 1
                    wrote = True
            if app.manual_break_end:
                if break_end_col is None:
                    stats.warnings.append(f"行ID={app.source_diff_row_id} 汎用データに '{JINJER_COL_BREAK_1_END}' 列がありません")
                else:
                    rows[idx][break_end_col] = app.manual_break_end
                    stats.overwritten_break_end += 1
                    wrote = True
            if app.manual_break_total:
                if break_total_col is None:
                    stats.warnings.append(f"行ID={app.source_diff_row_id} 汎用データに '{JINJER_COL_BREAK_TOTAL}' 列がありません")
                else:
                    rows[idx][break_total_col] = app.manual_break_total
                    stats.overwritten_break_total += 1
                    wrote = True

            if not wrote:
                stats.skipped_break += 1
                stats.warnings.append(
                    f"行ID={app.source_diff_row_id} 休憩差異が承認されましたが、手入力休憩欄が空のため反映しませんでした "
                    f"(emp={app.emp_id} date={app.target_date_iso} {app.name})"
                )
            continue
        if app.kind == DIFF_KIND_TOTAL:
            stats.skipped_total += 1
            stats.warnings.append(
                f"行ID={app.source_diff_row_id} 総労働時間差異が承認されましたが、自動反映はスキップしました "
                f"(emp={app.emp_id} date={app.target_date_iso} {app.name})"
            )
            continue

        if app.kind not in (DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT):
            stats.warnings.append(
                f"行ID={app.source_diff_row_id} 未知の差異種別 '{app.kind}' を無視 "
                f"(emp={app.emp_id} date={app.target_date_iso})"
            )
            continue

        # jinjer 行を引く
        key = (app.emp_id, app.target_date_iso)
        idx = row_index.get(key)
        if idx is None:
            stats.not_matched += 1
            stats.warnings.append(
                f"行ID={app.source_diff_row_id} jinjer CSV に該当行なし "
                f"(emp={app.emp_id} date={app.target_date_iso} {app.name})"
            )
            continue

        # 上書き（実績確定済かどうかに関わらず、承認されたものは上書きする）
        if app.kind == DIFF_KIND_PUNCH_IN:
            rows[idx][punch_in_col] = app.manual_fix_value or app.auto_fix_value
            stats.overwritten_punch_in += 1
        elif app.kind == DIFF_KIND_PUNCH_OUT:
            new_out = app.manual_fix_value or app.auto_fix_value
            # 夜勤で翌朝退勤の場合、jinjer は 24時超表記でないとインポートできないため
            # 現在の出勤1と突き合わせて 08:15→32:15 のように補正する（冪等・安全網）。
            new_out = to_overnight_punch_out(rows[idx][punch_in_col], new_out)
            rows[idx][punch_out_col] = new_out
            stats.overwritten_punch_out += 1

        # 参考集計: 上書き対象が実績確定済だった件数
        if finalized_col is not None:
            finalized = (rows[idx][finalized_col] or "").strip().upper()
            if finalized == "TRUE":
                stats.overwritten_finalized += 1


# ----------------------------------------------------------------------
# 出力
# ----------------------------------------------------------------------

def write_csv(output_path: Path, headers: list[str], rows: list[list[str]]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    # CRLF + CP932
    with open(output_path, "w", encoding="cp932", newline="") as f:
        writer = csv.writer(f, lineterminator="\r\n")
        writer.writerow(headers)
        writer.writerows(rows)


def print_summary(stats: Stats, output_path: Path, dry_run: bool, total_rows: int) -> None:
    mode = "[dry-run]" if dry_run else "[execute]"
    print()
    print(f"{mode} ===== 差異一覧の判断状況 =====")
    print(f"  総差異件数 : {stats.total_diff_rows}")
    print(f"  承認       : {stats.approved}")
    print(f"  却下       : {stats.rejected}")
    print(f"  保留       : {stats.held}")
    print(f"  未判断     : {stats.pending}")
    if stats.recovered_misplaced:
        detail = "、".join(f"{col}={cnt}" for col, cnt in stats.misplaced_by_col.items())
        print(f"  ※ 判断の列ミスを回収: {stats.recovered_misplaced} 件 ({detail})")
    print()
    print(f"{mode} ===== 承認の種別内訳 =====")
    print(f"  出勤差異   : {stats.approved_punch_in} → 出勤1 列に上書き")
    print(f"  退勤差異   : {stats.approved_punch_out} → 退勤1 列に上書き")
    print(f"  休憩差異   : {stats.approved_break} → 手入力欄がある場合のみ休憩列へ上書き")
    print(f"  総労働時間 : {stats.approved_total} → 警告のみ、上書きしない")
    print()
    print(f"{mode} ===== 実際の上書き結果 =====")
    print(f"  出勤1 上書き         : {stats.overwritten_punch_in}")
    print(f"  退勤1 上書き         : {stats.overwritten_punch_out}")
    print(f"  休憩1 上書き         : {stats.overwritten_break_start}")
    print(f"  復帰1 上書き         : {stats.overwritten_break_end}")
    print(f"  休憩時間 上書き      : {stats.overwritten_break_total}")
    print(f"  休憩スキップ         : {stats.skipped_break}")
    print(f"  総労働スキップ       : {stats.skipped_total}")
    print(f"  jinjer行 未マッチ    : {stats.not_matched}")
    print(f"  (うち実績確定済を上書き: {stats.overwritten_finalized} 件 / 参考)")
    print()
    print(f"{mode} 出力CSV総行数: {total_rows + 1} 行 (ヘッダー含む)")
    if dry_run:
        print(f"{mode} 出力先（書き出されません）: {output_path}")
        print(f"{mode} 本番書き出しは --execute を指定してください")
    else:
        print(f"[done] 出力完了: {output_path}")

    if stats.warnings:
        print()
        print(f"{mode} ===== 警告 ({len(stats.warnings)} 件) =====")
        for w in stats.warnings[:30]:
            print(f"  - {w}")
        if len(stats.warnings) > 30:
            print(f"  ... ほか {len(stats.warnings) - 30} 件")


# ----------------------------------------------------------------------
# 共通実行関数（CLI / Flask 共用）
# ----------------------------------------------------------------------

@dataclass
class ExportResult:
    ok: bool
    output_path: Path
    dry_run: bool
    stats: Stats = field(default_factory=Stats)
    total_jinjer_rows: int = 0
    error: str = ""


def run_quick_export(
    diff_xlsx: Path,
    jinjer_dir: Path,
    output_path: Path,
    dry_run: bool,
    log_func=print,
) -> ExportResult:
    """差異一覧xlsx + jinjer CSV → アップロード用CSV を生成。CLI と Web UI から共用。"""
    stats = Stats()
    result = ExportResult(ok=False, output_path=output_path, dry_run=dry_run, stats=stats)

    log_func(f"[start] 差異一覧: {diff_xlsx}")
    log_func(f"[start] jinjer CSV/フォルダ: {jinjer_dir}")
    log_func(f"[start] 出力先: {output_path}")
    log_func(f"[start] モード: {'dry-run' if dry_run else 'execute'}")

    try:
        approved = load_approved_rows(diff_xlsx, stats)
    except Exception as e:
        msg = f"差異一覧読み込み失敗: {e}"
        log_func(f"[error] {msg}")
        result.error = msg
        return result
    log_func(f"[info] 承認行 {len(approved)} 件 / 総差異 {stats.total_diff_rows} 件")

    try:
        headers, rows = load_jinjer_csvs(jinjer_dir)
    except Exception as e:
        msg = f"jinjer CSV 読み込み失敗: {e}"
        log_func(f"[error] {msg}")
        result.error = msg
        return result
    log_func(f"[info] jinjer 行 {len(rows)} 件 / 列数 {len(headers)}")
    result.total_jinjer_rows = len(rows)

    try:
        row_index = build_jinjer_row_index(headers, rows)
    except Exception as e:
        msg = f"jinjer 行インデックス構築失敗: {e}"
        log_func(f"[error] {msg}")
        result.error = msg
        return result
    log_func(f"[info] (従業員ID, 年月日) インデックス {len(row_index)} 件")

    apply_approved_rows(headers, rows, row_index, approved, stats)

    if not dry_run:
        try:
            write_csv(output_path, headers, rows)
        except Exception as e:
            msg = f"CSV 書き出し失敗: {e}"
            log_func(f"[error] {msg}")
            result.error = msg
            return result

    print_summary(stats, output_path, dry_run, len(rows))
    result.ok = True
    return result


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="5月本番 MVP jinjer アップロード用 全件CSV生成")
    p.add_argument("--diff", required=True, help="差異一覧xlsx（人間判断埋め込み済み）")
    p.add_argument("--jinjer-dir", required=True, help="元 jinjer ダウンロードCSV のフォルダ")
    p.add_argument("--output", required=True, help="出力 CSV パス (CP932)")
    p.add_argument("--execute", action="store_true", help="実際に CSV を書き出す（既定は dry-run）")
    return p.parse_args()


def _unquote_path(s: str) -> str:
    """前後の空白と、前後が同じクォートで囲まれているときだけ1組を除去する。"""
    s = (s or "").strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ('"', "'"):
        s = s[1:-1].strip()
    return s


def main() -> int:
    args = parse_args()
    result = run_quick_export(
        diff_xlsx=Path(_unquote_path(args.diff)),
        jinjer_dir=Path(_unquote_path(args.jinjer_dir)),
        output_path=Path(_unquote_path(args.output)),
        dry_run=not args.execute,
    )
    return 0 if result.ok else 1


if __name__ == "__main__":
    sys.exit(main())
