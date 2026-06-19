"""quick_compare.py — 5月本番 MVP 差異一覧生成スクリプト

入力:
  --kintai-dir : 既存 kintai-checker が出力した突合結果xlsx 群のフォルダ
  --jinjer-dir : jinjer 管理画面からダウンロードした「汎用データ（まるめ適用後）」CSV 群のフォルダ
  --output     : 出力する差異一覧xlsx のパス

出力:
  差異一覧_<YYYY-MM>.xlsx
    - サマリ
    - 差異一覧（人間判断プルダウン・警告レベル付き）
    - 取込ログ

設計書: docs/PLAN_5月本番_3営業日MVP.md  /  docs/DESIGN_月次マスター_P0_P3.md
"""

from __future__ import annotations

import argparse
import csv
import sys
import re
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule


# ===== jinjer CSV の列マップ（1-indexed、ヘッダー名で照合する想定だが定数として控えておく） =====
JINJER_HEADERS = {
    "name": "名前",
    "emp_id": "*従業員ID",
    "date": "*年月日",
    "punch_in_1": "出勤1",
    "punch_out_1": "退勤1",
    "break_1_start": "休憩1",
    "break_1_end": "復帰1",
    "finalized": "実績確定状況",
    "total_work": "総労働時間",
    "actual_work": "実労働時間",
    "break_total": "休憩時間",
}

# ===== 警告レベル =====
LEVEL_DANGER = "DANGER"
LEVEL_WARN = "WARN"
LEVEL_INFO = "INFO"

# ===== 警告閾値（後で調整可能） =====
PUNCH_DIFF_WARN_MIN = 120  # 出勤/退勤差分が 120分以上で WARN
LONG_WORK_HOURS = 10       # 総労働時間 10h 超で WARN
SHORT_BREAK_AT_LONG_WORK = 6  # 勤務 6h 超で休憩 0:00 なら DANGER（労基法的に必要）
OVER_BREAK_HOURS = 2       # 休憩 2h 超で WARN

# ===== 出力xlsx 列定義 =====
# 「人間判断」を「自動修正提案値」の直後に置く。判断列を提案値のすぐ隣にすることで、
# ユーザーが手入力欄（時刻を入れる列）へ 請求勤怠/jinjer勤怠/保留 を誤入力するのを防ぐ。
# 判断値: 「請求勤怠」=請求勤怠を正としてjinjerへ書き戻す / 「jinjer勤怠」=jinjerを正（書き戻さない） / 「保留」。
# 手入力欄（任意・上級者向け）は右側へ寄せる。
# ※ quick_export 側は列名で読むため、列順を変えても後方互換は保たれる。
# 横スクロールを減らすため「判断に必要な列」を左へ集約する。
# 識別3列(氏名/対象日付/差異種別)はウィンドウ枠固定。コメント2列は判断材料なので人間判断の左へ。
# 手入力・参考(予定/休日休暇/ID/トレーサビリティ)は右へ寄せる。
# ※ quick_export は列名で読むため、列順を変えても後方互換は保たれる。
DIFF_COLUMNS = [
    # 識別（ウィンドウ枠固定）
    "氏名", "対象日付", "差異種別",
    # すぐ見る（差異の中身）
    "請求勤怠値", "jinjer値", "差分(分)",
    # 確認区分（要確認/自動採用/自動OK/参考のみ）。深刻さは色＋警告理由で表す（警告レベル列は廃止）
    "確認区分",
    # 判断の根拠（人間判断の左にまとめる）。有休も判断材料なので予定/休日休暇もここに置く。
    "打刻時コメント",       # 汎用データ#96「打刻時コメント」より（出勤:/退勤: 両方）
    "打刻修正時コメント",   # 申請データCSVの「理由」より
    "警告理由",
    "出勤予定", "退勤予定", "休憩予定", "休日休暇名1", "休日休暇名1：種別",
    # 判断（入力）
    "人間判断", "判断メモ",
    # 反映（手入力・普段使わない）
    "自動修正提案値", "打刻修正", "手入力休憩1", "手入力復帰1", "手入力休憩時間",
    # 参考（右端・普段見ない）。実績確定/ID/トレーサビリティ
    "実績確定状況", "従業員ID", "行ID", "元突合結果ファイル",
]

# 汎用データから転記する予定・有休 列のキャノニカル名 → (候補ヘッダー, 完全一致のみか)
# 有休系は完全一致のみ（部分一致だと「有休」が「AM有休」「PM有休」を誤ヒットするため）。
JINJER_EXTRA_COLUMNS: dict[str, tuple[list[str], bool]] = {
    "出勤予定": (["出勤予定時刻", "出勤予定"], False),
    "退勤予定": (["退勤予定時刻", "退勤予定"], False),
    "休憩予定": (["休憩予定時間", "休憩予定"], False),
    "有休": (["有休"], True),
    "AM有休": (["AM有休"], True),
    "PM有休": (["PM有休"], True),
    # 休日休暇名1 は「休日休暇名1：種別」の部分文字列なので完全一致のみ（誤ヒット防止）。
    "休日休暇名1": (["休日休暇名1"], True),
    "休日休暇名1：種別": (["休日休暇名1：種別"], True),
    # 打刻時コメント（汎用データ#96）。"出勤: ○○ , 退勤: ○○" 形式で出勤退勤両方を含む。
    "打刻時コメント": (["打刻時コメント"], True),
}


def resolve_jinjer_extra_columns(columns) -> dict[str, str]:
    """汎用データの実ヘッダーから、予定/有休 のキャノニカル名→実ヘッダー名を解決する。

    完全一致を優先し、exact_only=False の列のみ部分一致でフォールバックする。
    見つからない列は dict に含めない（呼び出し側で空欄転記＋警告ログにする）。
    """
    cols = [str(c).strip() for c in columns]
    resolved: dict[str, str] = {}
    for canonical, (candidates, exact_only) in JINJER_EXTRA_COLUMNS.items():
        found = None
        for cand in candidates:  # 完全一致
            if cand in cols:
                found = cand
                break
        if found is None and not exact_only:  # 部分一致フォールバック
            for cand in candidates:
                for col in cols:
                    if cand in col:
                        found = col
                        break
                if found:
                    break
        if found:
            resolved[canonical] = found
    return resolved

DIFF_KIND_PUNCH_IN = "出勤"
DIFF_KIND_PUNCH_OUT = "退勤"
DIFF_KIND_BREAK = "休憩"
DIFF_KIND_TOTAL = "総労働時間"

# 既存 kintai-checker の services パスを通す（氏名正規化を借りる場合用）
KINTAI_CHECKER_ROOT = Path(__file__).resolve().parent
if str(KINTAI_CHECKER_ROOT) not in sys.path:
    sys.path.insert(0, str(KINTAI_CHECKER_ROOT))

# トリアージ（要確認/自動採用/自動OK の分類）
from services.triage import (  # noqa: E402
    classify as triage_classify,
    TRIAGE_NEEDS_CHECK,
    TRIAGE_AUTO_KINTAI,
    TRIAGE_AUTO_OK,
    TRIAGE_INFO_ONLY,
    TRIAGE_ORDER,
)


# ----------------------------------------------------------------------
# データクラス
# ----------------------------------------------------------------------

@dataclass
class DiffRow:
    row_id: int
    emp_id: str
    name: str
    target_date: str  # YYYY-MM-DD
    kind: str
    kintai_value: str
    jinjer_value: str
    diff_minutes: str  # 文字列で持つ（空欄表現のため）
    warn_level: str
    warn_reason: str
    auto_fix_value: str
    finalized: str  # 実績確定状況（参考情報。"TRUE"/"FALSE"/空）
    source_file: str
    # 汎用データから転記する予定・有休（参考表示。差異判定には使わない）
    sched_in: str = ""
    sched_out: str = ""
    sched_break: str = ""
    yukyu: str = ""
    am_yukyu: str = ""
    pm_yukyu: str = ""
    # 汎用データの休日休暇名1 / 休日休暇名1：種別（参考表示）
    holiday_name1: str = ""
    holiday_name1_type: str = ""
    # 打刻時コメント（汎用データ#96。出勤:/退勤: 両方を含む。判断材料）
    punch_comment: str = ""
    # 打刻修正時コメント（打刻修正申請の従業員コメント等。同一 emp/date の全差異行に併記）
    jinjer_stamp_comment: str = ""
    # トリアージ区分（要確認/自動採用/自動OK）と既定の人間判断
    triage: str = ""
    judge_default: str = ""


@dataclass
class LogEntry:
    severity: str  # INFO/WARN/ERROR
    message: str
    source: str = ""


# ----------------------------------------------------------------------
# 時刻ユーティリティ
# ----------------------------------------------------------------------

_TIME_RE = re.compile(r"^(\d{1,3}):(\d{2})(?::\d{2})?$")


def parse_hhmm(value: Any) -> int | None:
    """'H:MM' / 'HH:MM' / 'HH:MM:SS' を分に変換。空欄や不正値は None。
    24時超表記（25:30 等）はそのまま分換算する（時刻 ≠ 時間長 の用途で使い分け）。
    """
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    m = _TIME_RE.match(s)
    if not m:
        return None
    h = int(m.group(1))
    mm = int(m.group(2))
    return h * 60 + mm


def format_minutes_as_hhmm(minutes: int | None) -> str:
    if minutes is None:
        return ""
    h, m = divmod(int(minutes), 60)
    return f"{h}:{m:02d}"


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return ""
    return s


def normalize_kintai_result_columns(df: pd.DataFrame) -> pd.DataFrame:
    """新しい「請求勤怠」ヘッダーを内部処理用の旧キーへ寄せる。

    既存ユーザーがすでに出力済みの「勤務表」ヘッダーの差異一覧も読めるよう、
    内部キーは当面「勤務表_*」を維持する。
    """
    aliases = {
        "請求勤怠_出勤": "勤務表_出勤",
        "請求勤怠_出勤時刻": "勤務表_出勤時刻",
        "請求勤怠_退勤": "勤務表_退勤",
        "請求勤怠_退勤時刻": "勤務表_退勤時刻",
        "請求勤怠_総労働": "勤務表_総労働",
        "請求勤怠_総労働時間": "勤務表_総労働時間",
        "請求勤怠_総労働時間(分)": "勤務表_総労働時間(分)",
        "請求勤怠_総労働(分)": "勤務表_総労働(分)",
        "請求勤怠_実働": "勤務表_実働時間",
        "請求勤怠_実働時間": "勤務表_実働時間",
        "請求勤怠_コメント": "勤務表_コメント",
    }
    rename_map = {
        src: dst
        for src, dst in aliases.items()
        if src in df.columns and dst not in df.columns
    }
    if rename_map:
        return df.rename(columns=rename_map)
    return df


def elapsed_minutes_from_values(start: Any, end: Any) -> int | None:
    """出勤・退勤値から経過分を計算する。日跨ぎは翌日退勤として扱う。"""
    start_min = parse_hhmm(start)
    end_min = parse_hhmm(end)
    if start_min is None or end_min is None:
        return None
    if end_min < start_min:
        end_min += 24 * 60
    return end_min - start_min


def to_jinjer_overnight_punch_out(punch_in: Any, punch_out: Any) -> str:
    """日跨ぎ勤務の退勤時刻を jinjer インポート用の 24時超表記へ変換する。

    jinjer は「年月日」を勤務開始日として扱うため、翌朝退勤は 24時を超える表記
    （例: 翌 08:15 → 32:15）でないとインポートできない。請求勤怠の出勤 > 退勤
    （= 退勤が翌日にずれ込んだ夜勤）のときだけ、退勤に 24h を足した 'HH:MM' を返す。

    日跨ぎでない・出退勤のどちらかが空/不正なときは、退勤値をそのまま返す。
    すでに 24時超表記（32:15 等）なら out >= in となり再変換されない（冪等）。
    """
    out_clean = clean_cell(punch_out)
    in_min = parse_hhmm(punch_in)
    out_min = parse_hhmm(punch_out)
    if in_min is None or out_min is None or out_min >= in_min:
        return out_clean
    adjusted = out_min + 24 * 60
    h, m = divmod(adjusted, 60)
    return f"{h:02d}:{m:02d}"


def first_present(row: pd.Series, candidates: list[str]) -> Any:
    for col in candidates:
        if col not in row:
            continue
        value = row.get(col)
        if value is None:
            continue
        s = str(value).strip()
        if s and s.lower() not in ("nan", "none"):
            return value
    return None


def kintai_total_minutes(krow: pd.Series) -> tuple[int | None, str]:
    """勤怠突合結果行から勤務表側の総労働時間を取得する。

    請求勤怠ファイル記載の正味労働時間（勤務表_実働時間）があれば最優先で使う。
    無ければ拘束時間ベースの総労働列、それも無ければ勤務表_出勤/退勤から計算する。
    """
    explicit = first_present(krow, [
        # 請求勤怠ファイル記載の正味労働（休憩控除後）を最優先。
        "勤務表_実働時間",
        "勤務表_実働(分)",
        "勤務表_総労働",
        "勤務表_総労働時間",
        "勤務表_総労働時間(分)",
        "勤務表_総労働(分)",
    ])
    explicit_min = parse_hhmm(explicit)
    if explicit_min is not None:
        return explicit_min, format_minutes_as_hhmm(explicit_min)
    if explicit is not None:
        try:
            numeric = int(float(str(explicit)))
            return numeric, format_minutes_as_hhmm(numeric)
        except (TypeError, ValueError):
            pass

    start = first_present(krow, ["勤務表_出勤", "勤務表_出勤時刻"])
    end = first_present(krow, ["勤務表_退勤", "勤務表_退勤時刻"])
    total = elapsed_minutes_from_values(start, end)
    return total, format_minutes_as_hhmm(total)


def normalize_date_iso(value: Any) -> str | None:
    """日付値を 'YYYY-MM-DD' に正規化。jinjer の 'YYYY/M/D' にも対応。"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    s = str(value).strip()
    if not s or s.lower() in ("nan", "nat"):
        return None
    # ISO
    try:
        return datetime.strptime(s, "%Y-%m-%d").date().isoformat()
    except ValueError:
        pass
    # YYYY/M/D
    try:
        return datetime.strptime(s, "%Y/%m/%d").date().isoformat()
    except ValueError:
        pass
    # YYYY/MM/DD
    for fmt in ("%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return None


# ----------------------------------------------------------------------
# 入力読み込み
# ----------------------------------------------------------------------

KINTAI_RESULT_SHEET_CANDIDATES = ["突合結果", "突合"]


def _input_files(path: Path, patterns: list[str]) -> list[Path]:
    if path.is_file():
        return [path]

    files: list[Path] = []
    for pattern in patterns:
        files.extend(sorted(path.glob(pattern)))
    return files


def _parse_jp_date(value: Any) -> str | None:
    """'2026年05月07日' / 'YYYY-MM-DD' / 'YYYY/M/D' を 'YYYY-MM-DD' に正規化。"""
    s = str(value or "").strip()
    if not s:
        return None
    m = re.match(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
        except ValueError:
            return None
    return normalize_date_iso(value)


def load_stamp_correction_reasons(
    path: Path, logs: list[LogEntry]
) -> dict[tuple[str, str], list[dict]]:
    """jinjer「申請データ（打刻修正申請）」CSV を読み、(従業員ID, 日付ISO)→[{type,method,comment}] を返す。

    打刻修正申請の「理由」列が従業員コメント。理由が空の申請は除外する。
    同一(emp,date)に複数申請があれば全て保持し、フォーマット時に結合する。
    """
    result: dict[tuple[str, str], list[dict]] = {}
    if not path.exists():
        logs.append(LogEntry("WARN", f"申請データCSVが見つかりません: {path}"))
        return result

    total = 0
    for f in _input_files(path, ["*.csv"]):
        if f.name.startswith("~$"):
            continue
        df = None
        for enc in ("cp932", "utf-8-sig", "utf-8"):
            try:
                df = pd.read_csv(f, encoding=enc, dtype=object)
                break
            except UnicodeDecodeError:
                continue
            except Exception as e:
                logs.append(LogEntry("WARN", f"申請データCSV読込失敗: {f.name}: {e}"))
                df = None
                break
        if df is None:
            continue

        cols = {str(c).strip(): c for c in df.columns}
        emp_col, date_col, reason_col = cols.get("従業員ID"), cols.get("日付"), cols.get("理由")
        status_col = cols.get("ステータス")
        if not (emp_col and date_col and reason_col):
            logs.append(LogEntry(
                "WARN", f"申請データCSVに必要列(従業員ID/日付/理由)がありません: {f.name}"))
            continue

        for _, row in df.iterrows():
            reason = str(row.get(reason_col) or "").strip()
            if not reason or reason.lower() == "nan":
                continue
            emp_id = str(row.get(emp_col) or "").strip()
            date_iso = _parse_jp_date(row.get(date_col))
            if not emp_id or not date_iso:
                continue
            status = str(row.get(status_col) or "").strip() if status_col else ""
            comment = f"{reason}（{status}）" if status and status.lower() != "nan" else reason
            result.setdefault((emp_id, date_iso), []).append(
                {"type": "", "method": "", "comment": comment}
            )
            total += 1

    logs.append(LogEntry(
        "INFO", f"申請データCSVから打刻修正理由 {total} 件（{len(result)} (emp,date)）を読込"))
    return result


def load_kintai_results(kintai_dir: Path, logs: list[LogEntry]) -> pd.DataFrame:
    """突合結果xlsx を読んで縦結合する。ファイル単体指定にも対応する。"""
    if not kintai_dir.exists():
        logs.append(LogEntry("ERROR", f"突合結果xlsxまたはフォルダが見つかりません: {kintai_dir}"))
        return pd.DataFrame()

    frames = []
    xlsx_files = _input_files(kintai_dir, ["*.xlsx"])
    for xlsx in xlsx_files:
        # 一時ファイル除外
        if xlsx.name.startswith("~$"):
            continue
        try:
            xl = pd.ExcelFile(xlsx)
        except Exception as e:
            logs.append(LogEntry("ERROR", f"xlsx を開けません: {xlsx.name}: {e}"))
            continue
        sheet = None
        for cand in KINTAI_RESULT_SHEET_CANDIDATES:
            if cand in xl.sheet_names:
                sheet = cand
                break
        if sheet is None:
            sheet = xl.sheet_names[0]
            logs.append(LogEntry("WARN", f"突合結果シートが見つからないため先頭シート '{sheet}' を使用: {xlsx.name}"))
        try:
            df = pd.read_excel(xlsx, sheet_name=sheet, dtype=object)
        except Exception as e:
            logs.append(LogEntry("ERROR", f"シート読み込み失敗 {xlsx.name}:{sheet}: {e}"))
            continue
        df = normalize_kintai_result_columns(df)
        df["_source_file"] = xlsx.name
        frames.append(df)
        logs.append(LogEntry("INFO", f"突合結果読込: {xlsx.name} ({len(df)} 行)"))

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def load_jinjer_csvs(jinjer_dir: Path, logs: list[LogEntry]) -> pd.DataFrame:
    """jinjer CSV 群を CP932 で読んで縦結合する。ファイル単体指定にも対応する。"""
    if not jinjer_dir.exists():
        logs.append(LogEntry("ERROR", f"jinjer CSVまたはフォルダが見つかりません: {jinjer_dir}"))
        return pd.DataFrame()

    frames = []
    for path in _input_files(jinjer_dir, ["*.csv", "*.xlsx"]):
        if path.name.startswith("~$"):
            continue
        try:
            if path.suffix.lower() == ".csv":
                df = pd.read_csv(path, encoding="cp932", dtype=object)
            else:
                df = pd.read_excel(path, dtype=object)
        except Exception as e:
            logs.append(LogEntry("ERROR", f"jinjer データ読込失敗 {path.name}: {e}"))
            continue
        df["_jinjer_source"] = path.name
        frames.append(df)
        logs.append(LogEntry("INFO", f"jinjer 読込: {path.name} ({len(df)} 行)"))

    if not frames:
        return pd.DataFrame()
    # 必須列確認
    combined = pd.concat(frames, ignore_index=True)
    required = [JINJER_HEADERS["emp_id"], JINJER_HEADERS["date"]]
    for col in required:
        if col not in combined.columns:
            logs.append(LogEntry("ERROR", f"jinjer CSV に必須列 '{col}' がありません"))
            return pd.DataFrame()
    return combined


# ----------------------------------------------------------------------
# 突合
# ----------------------------------------------------------------------

def build_jinjer_index(jinjer_df: pd.DataFrame) -> dict[tuple[str, str], dict]:
    """(従業員ID, 日付ISO) -> jinjer 行 dict"""
    index: dict[tuple[str, str], dict] = {}
    emp_col = JINJER_HEADERS["emp_id"]
    date_col = JINJER_HEADERS["date"]
    for _, row in jinjer_df.iterrows():
        emp_id = str(row.get(emp_col) or "").strip()
        date_iso = normalize_date_iso(row.get(date_col))
        if not emp_id or not date_iso:
            continue
        index[(emp_id, date_iso)] = row.to_dict()
    return index


def build_name_to_emp_map(jinjer_df: pd.DataFrame) -> dict[str, str]:
    """氏名 → 従業員ID マップ。jinjer CSV の「名前」列から作る。
    姓名のスペース有無バリエーションも登録。同姓同名は最初の1人を採用（後で警告）。
    """
    name_col = JINJER_HEADERS["name"]
    emp_col = JINJER_HEADERS["emp_id"]
    result: dict[str, str] = {}
    if name_col not in jinjer_df.columns:
        return result
    seen_ids = set()
    for _, row in jinjer_df.iterrows():
        emp_id = str(row.get(emp_col) or "").strip()
        name = str(row.get(name_col) or "").strip()
        if not emp_id or not name or emp_id in seen_ids:
            continue
        seen_ids.add(emp_id)
        # 表記バリエーション
        variants = {
            name,
            name.replace(" ", ""),
            name.replace("　", ""),
            name.replace(" ", "").replace("　", ""),
        }
        for v in variants:
            if v and v not in result:
                result[v] = emp_id
    return result


def resolve_emp_id(name: Any, name_map: dict[str, str]) -> str | None:
    if not name:
        return None
    s = str(name).strip()
    if not s:
        return None
    for key in [s, s.replace(" ", ""), s.replace("　", ""), s.replace(" ", "").replace("　", "")]:
        if key in name_map:
            return name_map[key]
    return None


def to_int_diff(value: Any) -> int | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return None


def clean_punch_comment(value: Any) -> str:
    """汎用データ「打刻時コメント」(#96) の値を整形する。

    形式は "出勤: X , 退勤: Y" で、中身が無いと "出勤:  , 退勤:  ," のようなゴミ文字列に
    なるため、中身のあるラベルだけ残す。何も無ければ空文字を返す。
    """
    s = clean_cell(value)
    if not s:
        return ""
    parts = []
    found_label = False
    for m in re.finditer(r"(出勤|退勤)\s*[:：]\s*([^,]*)", s):
        found_label = True
        body = m.group(2).strip()
        if body:
            parts.append(f"{m.group(1)}: {body}")
    if parts:
        return " / ".join(parts)
    if found_label:
        return ""  # ラベルはあるが中身なし → ゴミなので空
    return s  # 想定外フォーマットはそのまま


def format_stamp_comments(items: list[dict]) -> str:
    """打刻コメント list（{type,method,comment}）を1セルぶんの文字列に整形する。

    例: "出勤[打刻修正申請] KDX出社 / 退勤[PC] 私用のため早退"
    """
    parts = []
    for it in items or []:
        comment = str(it.get("comment") or "").strip()
        if not comment:
            continue
        stype = str(it.get("type") or "").strip()
        method = str(it.get("method") or "").strip()
        label = "".join([stype, f"[{method}]" if method else ""])
        parts.append(f"{label} {comment}".strip() if label else comment)
    return " / ".join(parts)


def compute_diffs(
    kintai_df: pd.DataFrame,
    jinjer_index: dict[tuple[str, str], dict],
    name_map: dict[str, str],
    logs: list[LogEntry],
    extra_cols: dict[str, str] | None = None,
    stamp_comments: dict[tuple[str, str], list[dict]] | None = None,
) -> list[DiffRow]:
    """突合結果xlsx の各行に対し、4種の差異を生成して返す。

    stamp_comments: ``{(従業員ID, 日付ISO): [{type,method,comment}, ...]}``。
    打刻修正申請の従業員コメント等を、同一 emp/date の各差異行に併記する。
    """
    rows: list[DiffRow] = []
    next_id = 1
    extra_cols = extra_cols or {}
    stamp_comments = stamp_comments or {}

    seen_emp_date: set[tuple[str, str]] = set()

    for _, krow in kintai_df.iterrows():
        name = str(krow.get("氏名") or "").strip()
        date_iso = normalize_date_iso(krow.get("日付"))
        if not name or not date_iso:
            continue

        emp_id = resolve_emp_id(name, name_map)
        if not emp_id:
            logs.append(LogEntry("WARN", f"氏名 → 従業員ID 解決失敗: {name} ({krow.get('_source_file', '')})"))
            continue

        source_file = str(krow.get("_source_file") or "")
        jrow = jinjer_index.get((emp_id, date_iso))
        seen_emp_date.add((emp_id, date_iso))

        # 実績確定状況（参考表示のみ、警告レベル判定には使わない）
        finalized = ""
        if jrow is not None:
            finalized = str(jrow.get(JINJER_HEADERS["finalized"]) or "").strip()

        # 汎用データから予定・有休を転記（同一 emp/date のすべての差異行に併記する参考情報）。
        # 列が解決できない / jinjer 行が無いときは空欄。
        def _extra(canonical: str) -> str:
            if jrow is None:
                return ""
            header = extra_cols.get(canonical)
            return clean_cell(jrow.get(header)) if header else ""

        extra = {
            "sched_in": _extra("出勤予定"),
            "sched_out": _extra("退勤予定"),
            "sched_break": _extra("休憩予定"),
            "yukyu": _extra("有休"),
            "am_yukyu": _extra("AM有休"),
            "pm_yukyu": _extra("PM有休"),
            "holiday_name1": _extra("休日休暇名1"),
            "holiday_name1_type": _extra("休日休暇名1：種別"),
            "punch_comment": clean_punch_comment(_extra("打刻時コメント")),
            "jinjer_stamp_comment": format_stamp_comments(stamp_comments.get((emp_id, date_iso))),
        }

        # ----- 出勤差異 -----
        k_in = clean_cell(krow.get("勤務表_出勤"))
        j_in = clean_cell(krow.get("jinjer_出勤"))
        diff_in = to_int_diff(krow.get("出勤差分(分)"))
        if diff_in is not None and diff_in != 0:
            level, reason = classify_punch_diff(diff_in, "in")
            rows.append(DiffRow(
                row_id=next_id, emp_id=emp_id, name=name, target_date=date_iso,
                kind=DIFF_KIND_PUNCH_IN,
                kintai_value=k_in, jinjer_value=j_in, diff_minutes=str(diff_in),
                warn_level=level, warn_reason=reason,
                auto_fix_value=k_in,
                finalized=finalized,
                source_file=source_file,
                **extra,
            ))
            next_id += 1
        elif diff_in is None and k_in and not j_in:
            rows.append(DiffRow(
                row_id=next_id, emp_id=emp_id, name=name, target_date=date_iso,
                kind=DIFF_KIND_PUNCH_IN,
                kintai_value=k_in, jinjer_value="", diff_minutes="",
                warn_level=LEVEL_WARN,
                warn_reason="jinjer出勤なし / 請求勤怠側に時刻あり",
                auto_fix_value=k_in,
                finalized=finalized,
                source_file=source_file,
                **extra,
            ))
            next_id += 1

        # ----- 退勤差異 -----
        k_out = clean_cell(krow.get("勤務表_退勤"))
        j_out = clean_cell(krow.get("jinjer_退勤"))
        diff_out = to_int_diff(krow.get("退勤差分(分)"))
        if diff_out is not None and diff_out != 0:
            level, reason = classify_punch_diff(diff_out, "out")
            rows.append(DiffRow(
                row_id=next_id, emp_id=emp_id, name=name, target_date=date_iso,
                kind=DIFF_KIND_PUNCH_OUT,
                kintai_value=k_out, jinjer_value=j_out, diff_minutes=str(diff_out),
                warn_level=level, warn_reason=reason,
                auto_fix_value=to_jinjer_overnight_punch_out(k_in, k_out),
                finalized=finalized,
                source_file=source_file,
                **extra,
            ))
            next_id += 1
        elif diff_out is None and k_out and not j_out:
            rows.append(DiffRow(
                row_id=next_id, emp_id=emp_id, name=name, target_date=date_iso,
                kind=DIFF_KIND_PUNCH_OUT,
                kintai_value=k_out, jinjer_value="", diff_minutes="",
                warn_level=LEVEL_WARN,
                warn_reason="jinjer退勤なし / 請求勤怠側に時刻あり",
                auto_fix_value=to_jinjer_overnight_punch_out(k_in, k_out),
                finalized=finalized,
                source_file=source_file,
                **extra,
            ))
            next_id += 1

        # ----- 総労働時間差異 -----
        # 「休憩」の突合は廃止（総労働時間が正味で突合されるため不要）。
        if jrow is not None:
            k_total_min, k_total = kintai_total_minutes(krow)
            j_total = str(jrow.get(JINJER_HEADERS["total_work"]) or "").strip()
            j_total_min = parse_hhmm(j_total)
            if k_total_min is not None and j_total_min is not None and k_total_min != j_total_min:
                diff_total = k_total_min - j_total_min
                level, reason = classify_total_work_diff(diff_total, jrow)
                rows.append(DiffRow(
                    row_id=next_id, emp_id=emp_id, name=name, target_date=date_iso,
                    kind=DIFF_KIND_TOTAL,
                    kintai_value=k_total, jinjer_value=j_total, diff_minutes=str(diff_total),
                    warn_level=level, warn_reason=reason,
                    auto_fix_value="",  # 総労働時間は自動反映しない
                    finalized=finalized,
                    source_file=source_file,
                    **extra,
                ))
                next_id += 1
            else:
                total_warn = classify_total_work(jrow)
                if total_warn:
                    level, reason, j_total = total_warn
                    rows.append(DiffRow(
                        row_id=next_id, emp_id=emp_id, name=name, target_date=date_iso,
                        kind=DIFF_KIND_TOTAL,
                        kintai_value=k_total or "-", jinjer_value=j_total, diff_minutes="",
                        warn_level=level, warn_reason=reason,
                        auto_fix_value="",  # 総労働時間は自動反映しない
                        finalized=finalized,
                        source_file=source_file,
                        **extra,
                    ))
                    next_id += 1

    # トリアージ: 各差異行を 要確認/自動採用/自動OK に分類し、既定の人間判断を付ける
    for r in rows:
        r.triage, r.judge_default = triage_classify(
            kind=r.kind,
            warn_level=r.warn_level,
            punch_comment=r.punch_comment,
            stamp_comment=r.jinjer_stamp_comment,
            kintai_value=r.kintai_value,
            holiday_name1=r.holiday_name1,
            holiday_name1_type=r.holiday_name1_type,
        )

    return rows


def classify_punch_diff(diff_minutes: int, side: str) -> tuple[str, str]:
    """出勤/退勤の差分行の警告レベル判定。
    side: 'in' or 'out'

    注意: 実績確定状況（jinjer 側で本人が打刻申請を確定した状態）は警告レベル判定に
    使わない。実績確定済は「本人が確定済み」という意味であって「勤怠が正しい」という
    意味ではないため、それを理由に上書きをブロックしたり DANGER 表示すべきではない。
    実績確定状況は差異一覧シートに参考列として表示するだけ。
    """
    reasons: list[str] = []
    level = LEVEL_INFO

    abs_diff = abs(diff_minutes)
    if abs_diff >= PUNCH_DIFF_WARN_MIN:
        level = LEVEL_WARN
        reasons.append(f"{side}差分 {abs_diff}分 (>={PUNCH_DIFF_WARN_MIN})")

    if not reasons:
        reasons.append("通常差異")
    return level, " / ".join(reasons)


def classify_break(jrow: dict) -> tuple[str, str, str] | None:
    """jinjer 側の休憩時間に問題があるか判定。
    戻り値: (level, reason, jinjer_break_str) or None（警告不要）
    """
    j_break_str = str(jrow.get(JINJER_HEADERS["break_total"]) or "").strip()
    j_total_str = str(jrow.get(JINJER_HEADERS["total_work"]) or "").strip()
    j_break_min = parse_hhmm(j_break_str)
    j_total_min = parse_hhmm(j_total_str)

    reasons: list[str] = []
    level = LEVEL_INFO

    # ① 勤務 6h 超で休憩 0 → 労基法違反疑い (DANGER)
    if j_break_min == 0 and j_total_min is not None and j_total_min >= SHORT_BREAK_AT_LONG_WORK * 60:
        level = LEVEL_DANGER
        reasons.append(f"休憩 0:00 だが総労働 {format_minutes_as_hhmm(j_total_min)} (労基法違反疑い)")

    # ② 休憩 2h 超 → WARN
    if j_break_min is not None and j_break_min > OVER_BREAK_HOURS * 60:
        if level == LEVEL_INFO:
            level = LEVEL_WARN
        reasons.append(f"休憩 {format_minutes_as_hhmm(j_break_min)} 過剰の疑い")

    if not reasons:
        return None
    return level, " / ".join(reasons), j_break_str


def classify_total_work(jrow: dict) -> tuple[str, str, str] | None:
    """jinjer 側の総労働時間に問題があるか判定。
    戻り値: (level, reason, jinjer_total_str) or None
    """
    j_total_str = str(jrow.get(JINJER_HEADERS["total_work"]) or "").strip()
    j_punch_in = str(jrow.get(JINJER_HEADERS["punch_in_1"]) or "").strip()
    j_total_min = parse_hhmm(j_total_str)

    reasons: list[str] = []
    level = LEVEL_INFO

    if j_total_min is None:
        return None

    # ① 長時間労働
    if j_total_min > LONG_WORK_HOURS * 60:
        level = LEVEL_WARN
        reasons.append(f"総労働 {format_minutes_as_hhmm(j_total_min)} (>{LONG_WORK_HOURS}h)")

    # ② 出勤打刻あり / 総労働 0:00 → 計算不能警告
    if j_total_min == 0 and j_punch_in:
        level = LEVEL_DANGER
        reasons.append("出勤打刻あり/総労働 0:00 (集計不整合)")

    if not reasons:
        return None
    return level, " / ".join(reasons), j_total_str


def classify_total_work_diff(diff_minutes: int, jrow: dict) -> tuple[str, str]:
    """勤務表とjinjerの総労働時間差異を判定。"""
    abs_diff = abs(diff_minutes)
    reasons = [f"請求勤怠総労働とjinjer総労働の差分 {abs_diff}分"]
    level = LEVEL_INFO
    if abs_diff >= PUNCH_DIFF_WARN_MIN:
        level = LEVEL_WARN

    j_total_str = str(jrow.get(JINJER_HEADERS["total_work"]) or "").strip()
    j_punch_in = str(jrow.get(JINJER_HEADERS["punch_in_1"]) or "").strip()
    j_total_min = parse_hhmm(j_total_str)
    if j_total_min == 0 and j_punch_in:
        level = LEVEL_DANGER
        reasons.append("出勤打刻あり/総労働 0:00 (集計不整合)")
    elif j_total_min is not None and j_total_min > LONG_WORK_HOURS * 60 and level == LEVEL_INFO:
        level = LEVEL_WARN
        reasons.append(f"総労働 {format_minutes_as_hhmm(j_total_min)} (>{LONG_WORK_HOURS}h)")

    return level, " / ".join(reasons)


# ----------------------------------------------------------------------
# Excel 出力
# ----------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
# 「人間判断」列ヘッダー強調（入力すべき列だと分かるように）
JUDGE_HEADER_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
LEVEL_FILL = {
    LEVEL_DANGER: PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    LEVEL_WARN: PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
    LEVEL_INFO: PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
}
# トリアージ区分の色（要確認を目立たせ、自動は淡色に）
TRIAGE_FILL = {
    TRIAGE_NEEDS_CHECK: PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    TRIAGE_AUTO_OK: PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"),
    TRIAGE_AUTO_KINTAI: PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    TRIAGE_INFO_ONLY: PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"),
}


def write_excel(output_path: Path, diff_rows: list[DiffRow], logs: list[LogEntry], month_label: str) -> None:
    # トリアージ区分で並べ替え（要確認を上へ。同区分内は元の順序を保持＝安定ソート）
    diff_rows = sorted(diff_rows, key=lambda r: TRIAGE_ORDER.get(r.triage, 9))
    needs_cnt = sum(1 for r in diff_rows if r.triage == TRIAGE_NEEDS_CHECK)
    auto_ok_cnt = sum(1 for r in diff_rows if r.triage == TRIAGE_AUTO_OK)
    auto_kintai_cnt = sum(1 for r in diff_rows if r.triage == TRIAGE_AUTO_KINTAI)
    info_only_cnt = sum(1 for r in diff_rows if r.triage == TRIAGE_INFO_ONLY)

    wb = Workbook()
    # サマリ
    ws_sum = wb.active
    ws_sum.title = "サマリ"
    summary_data = [
        ("対象月", month_label),
        ("総差異件数", len(diff_rows)),
        ("★要確認 件数（人が見る）", needs_cnt),
        ("自動OK(jinjer勤怠) 件数", auto_ok_cnt),
        ("自動採用(請求勤怠) 件数", auto_kintai_cnt),
        ("参考のみ 件数（判断不要）", info_only_cnt),
        ("出勤差異件数", sum(1 for r in diff_rows if r.kind == DIFF_KIND_PUNCH_IN)),
        ("退勤差異件数", sum(1 for r in diff_rows if r.kind == DIFF_KIND_PUNCH_OUT)),
        ("総労働時間差異件数", sum(1 for r in diff_rows if r.kind == DIFF_KIND_TOTAL)),
        ("DANGER 件数", sum(1 for r in diff_rows if r.warn_level == LEVEL_DANGER)),
        ("WARN 件数", sum(1 for r in diff_rows if r.warn_level == LEVEL_WARN)),
        ("INFO 件数", sum(1 for r in diff_rows if r.warn_level == LEVEL_INFO)),
        ("生成日時", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]
    for r_idx, (label, value) in enumerate(summary_data, start=1):
        c1 = ws_sum.cell(row=r_idx, column=1, value=label)
        c1.font = Font(bold=True)
        ws_sum.cell(row=r_idx, column=2, value=value)
    ws_sum.column_dimensions["A"].width = 22
    ws_sum.column_dimensions["B"].width = 28

    # 差異一覧
    ws = wb.create_sheet("差異一覧")
    for col_idx, header in enumerate(DIFF_COLUMNS, start=1):
        c = ws.cell(row=1, column=col_idx, value=header)
        # 「人間判断」列は入力すべき列だと一目で分かるよう、ヘッダーを目立たせる
        c.fill = JUDGE_HEADER_FILL if header == "人間判断" else HEADER_FILL
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.auto_filter.ref = f"A1:{get_column_letter(len(DIFF_COLUMNS))}1"
    # 識別3列(氏名/対象日付/差異種別=A〜C)とヘッダー行を固定。右に行っても誰の行か見失わない。
    ws.freeze_panes = "D2"

    # データ行（列名→値の対応で書き込み、列順の変更に強くする）
    for r_idx, drow in enumerate(diff_rows, start=2):
        row_values = {
            "行ID": drow.row_id,
            "従業員ID": drow.emp_id,
            "氏名": drow.name,
            "対象日付": drow.target_date,
            "出勤予定": drow.sched_in,
            "退勤予定": drow.sched_out,
            "休憩予定": drow.sched_break,
            "有休": drow.yukyu,
            "AM有休": drow.am_yukyu,
            "PM有休": drow.pm_yukyu,
            "休日休暇名1": drow.holiday_name1,
            "休日休暇名1：種別": drow.holiday_name1_type,
            "差異種別": drow.kind,
            "請求勤怠値": drow.kintai_value,
            "jinjer値": drow.jinjer_value,
            "差分(分)": drow.diff_minutes,
            "確認区分": drow.triage,
            "警告理由": drow.warn_reason,
            "打刻時コメント": drow.punch_comment,
            "打刻修正時コメント": drow.jinjer_stamp_comment,
            "自動修正提案値": drow.auto_fix_value,
            "人間判断": drow.judge_default,   # トリアージの既定値を事前入力（要確認は空欄）
            "判断メモ": "",
            "打刻修正": "",        # 出勤/退勤の提案値を人間が上書きしたい場合のみ
            "手入力休憩1": "",     # 休憩差異を承認して汎用データに反映する場合のみ
            "手入力復帰1": "",
            "手入力休憩時間": "",
            "実績確定状況": drow.finalized,  # 参考表示
            "元突合結果ファイル": drow.source_file,
        }
        for c_idx, header in enumerate(DIFF_COLUMNS, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=row_values.get(header, ""))
            cell.alignment = Alignment(vertical="center")
        # 深刻さ（DANGER/WARN/INFO）は「警告理由」セルの色で表す（警告レベル列は廃止）
        fill = LEVEL_FILL.get(drow.warn_level)
        if fill:
            ws.cell(row=r_idx, column=DIFF_COLUMNS.index("警告理由") + 1).fill = fill
        # 確認区分セルに色塗り（要確認を目立たせる）
        tfill = TRIAGE_FILL.get(drow.triage)
        if tfill:
            ws.cell(row=r_idx, column=DIFF_COLUMNS.index("確認区分") + 1).fill = tfill

    # データ検証プルダウン: 人間判断
    judge_col_letter = get_column_letter(DIFF_COLUMNS.index("人間判断") + 1)
    if diff_rows:
        dv = DataValidation(
            type="list",
            formula1='"請求勤怠,jinjer勤怠,保留"',
            allow_blank=True,
            showErrorMessage=True,
            errorTitle="不正な値",
            error="請求勤怠 / jinjer勤怠 / 保留 のいずれかを選択してください",
        )
        ws.add_data_validation(dv)
        judge_range = f"{judge_col_letter}2:{judge_col_letter}{1 + len(diff_rows)}"
        dv.add(judge_range)

        # 人間判断の選択値に応じてセル色を変える（条件付き書式。選んだ瞬間に反映される）
        # 保留 → 背景赤・太字白 ／ jinjer勤怠 → 背景黄・太字黒
        ws.conditional_formatting.add(judge_range, CellIsRule(
            operator="equal", formula=['"保留"'],
            fill=PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"),
            font=Font(bold=True, color="FFFFFF"),
        ))
        ws.conditional_formatting.add(judge_range, CellIsRule(
            operator="equal", formula=['"jinjer勤怠"'],
            fill=PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"),
            font=Font(bold=True, color="000000"),
        ))

    # 列幅（列名で指定。列順を変えても崩れない）
    width_map = {
        "行ID": 6, "従業員ID": 12, "氏名": 16, "対象日付": 12,
        "出勤予定": 10, "退勤予定": 10, "休憩予定": 10, "有休": 8, "AM有休": 8, "PM有休": 8,
        "休日休暇名1": 14, "休日休暇名1：種別": 12,
        "差異種別": 10, "請求勤怠値": 10, "jinjer値": 10, "差分(分)": 8,
        "確認区分": 14, "警告理由": 40, "自動修正提案値": 14,
        "人間判断": 12, "判断メモ": 28,
        "打刻時コメント": 40, "打刻修正時コメント": 40,
        "打刻修正": 14, "手入力休憩1": 14, "手入力復帰1": 14, "手入力休憩時間": 14,
        "実績確定状況": 12, "元突合結果ファイル": 32,
    }
    for i, header in enumerate(DIFF_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width_map.get(header, 14)

    # 取込ログ
    ws_log = wb.create_sheet("取込ログ")
    ws_log.cell(row=1, column=1, value="severity").font = Font(bold=True)
    ws_log.cell(row=1, column=2, value="message").font = Font(bold=True)
    ws_log.cell(row=1, column=3, value="source").font = Font(bold=True)
    for r_idx, entry in enumerate(logs, start=2):
        ws_log.cell(row=r_idx, column=1, value=entry.severity)
        ws_log.cell(row=r_idx, column=2, value=entry.message)
        ws_log.cell(row=r_idx, column=3, value=entry.source)
    ws_log.column_dimensions["A"].width = 10
    ws_log.column_dimensions["B"].width = 80
    ws_log.column_dimensions["C"].width = 32

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


# ----------------------------------------------------------------------
# 共通実行関数（CLI / Flask 共用）
# ----------------------------------------------------------------------

@dataclass
class CompareResult:
    ok: bool
    output_path: Path
    diff_count: int = 0
    danger_count: int = 0
    warn_count: int = 0
    info_count: int = 0
    kintai_rows_read: int = 0
    jinjer_rows_read: int = 0
    name_map_size: int = 0
    error: str = ""
    logs: list[LogEntry] = field(default_factory=list)


def run_quick_compare(
    kintai_dir: Path,
    jinjer_dir: Path,
    output_path: Path,
    month_label: str,
    log_func=print,
    application_csv: Path | None = None,
) -> CompareResult:
    """突合結果xlsx + jinjer CSV → 差異一覧xlsx を生成する。CLI と Web UI から共用。

    application_csv: jinjer「申請データ（打刻修正申請）」CSV のパス（任意）。
    指定すると「打刻修正時コメント」列にその「理由」を併記する。未指定なら attendances API
    の打刻コメントにフォールバックする。
    """
    logs: list[LogEntry] = []
    result = CompareResult(ok=False, output_path=output_path, logs=logs)

    log_func(f"[start] 勤怠突合結果xlsx/フォルダ: {kintai_dir}")
    log_func(f"[start] jinjer CSV/フォルダ: {jinjer_dir}")
    log_func(f"[start] 出力先: {output_path}")

    kintai_df = load_kintai_results(kintai_dir, logs)
    if kintai_df.empty:
        msg = "突合結果xlsx が読めませんでした"
        log_func(f"[error] {msg}")
        result.error = msg
        write_excel(output_path, [], logs, month_label)
        return result

    jinjer_df = load_jinjer_csvs(jinjer_dir, logs)
    if jinjer_df.empty:
        msg = "jinjer CSV が読めませんでした"
        log_func(f"[error] {msg}")
        result.error = msg
        write_excel(output_path, [], logs, month_label)
        return result

    result.kintai_rows_read = len(kintai_df)
    result.jinjer_rows_read = len(jinjer_df)
    log_func(f"[info] 突合結果 {len(kintai_df)} 行 / jinjer {len(jinjer_df)} 行 読み込み完了")

    name_map = build_name_to_emp_map(jinjer_df)
    jinjer_index = build_jinjer_index(jinjer_df)
    result.name_map_size = len(name_map)
    log_func(f"[info] 氏名→ID マップ {len(name_map)} 件 / jinjer index {len(jinjer_index)} 件")

    # 汎用データから転記する予定・有休 列を解決（見つからない列は空欄転記＋ログ）
    extra_cols = resolve_jinjer_extra_columns(jinjer_df.columns)
    missing_extra = [k for k in JINJER_EXTRA_COLUMNS if k not in extra_cols]
    if missing_extra:
        logs.append(LogEntry(
            "INFO",
            f"汎用データに次の列が見つからないため空欄で転記します: {', '.join(missing_extra)}",
        ))
    log_func(f"[info] 予定/有休 転記列の解決: {extra_cols}")

    # 「打刻修正時コメント」列の元データ。申請データCSVが指定されていればその「理由」を最優先。
    # 未指定なら attendances API の打刻コメントにフォールバック。いずれも取得できなくても続行。
    stamp_comments: dict[tuple[str, str], list[dict]] = {}
    if application_csv is not None:
        stamp_comments = load_stamp_correction_reasons(application_csv, logs)
        log_func(f"[info] 申請データCSVから打刻修正理由: {len(stamp_comments)} (emp,date) 件")
    else:
        checked_emp_ids = sorted({
            eid for n in kintai_df["氏名"].dropna().unique()
            if (eid := resolve_emp_id(str(n).strip(), name_map))
        })
        if checked_emp_ids and re.fullmatch(r"\d{4}-\d{2}", str(month_label or "").strip()):
            try:
                from services.jinjer_api_client import fetch_stamp_comments
                stamp_comments = fetch_stamp_comments(checked_emp_ids, str(month_label).strip())
                log_func(f"[info] jinjer 打刻コメント取得(API): {len(stamp_comments)} (emp,date) 件")
            except Exception as e:  # 認証失敗・通信不可など。差異一覧は続行する。
                logs.append(LogEntry("WARN", f"jinjer 打刻コメント取得に失敗（コメント欄は空で続行）: {e}"))
                log_func(f"[warn] jinjer 打刻コメント取得に失敗: {e}")
        else:
            logs.append(LogEntry(
                "INFO",
                f"打刻コメント取得をスキップ（申請データCSV未指定 / 対象月={month_label} / 対象ID数={len(checked_emp_ids)}）",
            ))

    diff_rows = compute_diffs(kintai_df, jinjer_index, name_map, logs, extra_cols, stamp_comments)
    result.diff_count = len(diff_rows)
    log_func(f"[info] 差異・警告 合計 {len(diff_rows)} 件")

    by_level = {LEVEL_DANGER: 0, LEVEL_WARN: 0, LEVEL_INFO: 0}
    for r in diff_rows:
        by_level[r.warn_level] = by_level.get(r.warn_level, 0) + 1
    result.danger_count = by_level[LEVEL_DANGER]
    result.warn_count = by_level[LEVEL_WARN]
    result.info_count = by_level[LEVEL_INFO]
    log_func(f"[info] DANGER={result.danger_count} / WARN={result.warn_count} / INFO={result.info_count}")

    write_excel(output_path, diff_rows, logs, month_label)
    log_func(f"[done] 出力完了: {output_path}")
    result.ok = True
    return result


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="5月本番 MVP 差異一覧生成")
    p.add_argument("--month", required=True, help="対象月 YYYY-MM (例: 2026-05)")
    p.add_argument("--kintai-dir", required=True, help="突合結果xlsx 格納フォルダ")
    p.add_argument("--jinjer-dir", required=True, help="jinjer 汎用データCSV 格納フォルダ")
    p.add_argument("--application-csv", default="", help="jinjer 申請データ(打刻修正申請) CSV（任意）")
    p.add_argument("--output", required=True, help="出力 xlsx パス")
    return p.parse_args()


def _unquote_path(s: str) -> str:
    """前後の空白と、前後が同じクォートで囲まれているときだけ1組を除去する。"""
    s = (s or "").strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ('"', "'"):
        s = s[1:-1].strip()
    return s


def main() -> int:
    args = parse_args()
    app_csv = _unquote_path(args.application_csv)
    result = run_quick_compare(
        kintai_dir=Path(_unquote_path(args.kintai_dir)),
        jinjer_dir=Path(_unquote_path(args.jinjer_dir)),
        output_path=Path(_unquote_path(args.output)),
        month_label=args.month,
        application_csv=Path(app_csv) if app_csv else None,
    )
    return 0 if result.ok else 1


if __name__ == "__main__":
    sys.exit(main())
