"""多年度横並びレイアウトのシフト表 xlsx と、シフトルール xlsx の構造化パーサ

以下の2種類の xlsx を Claude API を使わず確定的に解析する。

1. シフトルール (legend) xlsx
   行ごとに | 記号 | 開始時間 | 終了時間 | 休憩開始 | 休憩終了 |
   末尾に「休(m/d) | 振休」「休 | 休暇」といった休扱い行が続く。

2. シフト表 xlsx (多年度横並び)
   - 1〜2 行目: 休マーカー / 日付行（Excel シリアル値が並ぶ）
   - 3 行目: 曜日（土,日,月,...）
   - 4 行目以降: 氏名 + シフトパターン列 + 各日のシフト記号
   - 同じ「4/1 〜 3/31」または「1/1 〜 12/31」のパターンが複数年度分
     横方向に積まれており、年ラベルは無い。日付セルだけは全年度とも
     2026 シリアルが入っている（コピーで上書きされた状態）が、
     **3 行目の曜日は元の年度のまま**なので、target_year の曜日と一致
     するセクションが「本物の」当該年度データである。

target_year/target_month に対応するセクションを曜日マッチで自動選定し、
{name, shifts:[{date, code}]} 形式の従業員データを返す。"""

from __future__ import annotations

import calendar
import logging
import os
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Iterable

import openpyxl

logger = logging.getLogger(__name__)


# 曜日 → 漢字 1 文字（Mon=0..Sun=6）
_WD_KANJI = ["月", "火", "水", "木", "金", "土", "日"]

# Excel epoch (1900 leap-year bug 込み)
_EXCEL_EPOCH = date(1899, 12, 30)


# =============================================================================
# 共通ユーティリティ
# =============================================================================

def _to_date(v) -> date | None:
    """セル値を date に変換できれば返す（数値シリアル / datetime / date のみ対応）"""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        try:
            n = int(v)
        except Exception:
            return None
        if n < 1 or n > 60_000:
            return None
        try:
            return _EXCEL_EPOCH + timedelta(days=n)
        except Exception:
            return None
    return None


def _to_hhmm(v) -> str | None:
    """セル値を "HH:MM" 文字列に変換"""
    if v is None:
        return None
    if isinstance(v, time):
        return v.strftime("%H:%M")
    if isinstance(v, datetime):
        return v.strftime("%H:%M")
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        # "07:00:00" / "7:00" 等を切り詰める
        parts = s.split(":")
        if len(parts) >= 2:
            try:
                h = int(parts[0])
                m = int(parts[1])
                return f"{h:02d}:{m:02d}"
            except ValueError:
                return None
        return None
    return None


def _clean_str(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


# =============================================================================
# 1. シフトルール xlsx パーサ
# =============================================================================

# 「休扱い」とみなすラベル（凡例に開始/終了時刻が無い行用）
_OFF_LABEL_PATTERNS = ("振休", "休暇", "公休", "休日", "代休", "明け")


def parse_legend_xlsx(filepath: str) -> list[dict]:
    """シフトルール xlsx を解析して凡例リストを返す

    Returns:
        [{code, label, start_time, end_time, break_minutes, is_off}, ...]
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    legend: list[dict] = []
    seen_codes: set[str] = set()

    for row in rows:
        if not row:
            continue
        code = _clean_str(row[0])
        if not code:
            continue
        # ヘッダ行は除外
        if code in ("記号", "コード", "シフト記号"):
            continue

        col1 = _clean_str(row[1]) if len(row) > 1 else ""

        # 開始/終了時刻が文字列ラベルの場合（休系）
        start_time = _to_hhmm(row[1]) if len(row) > 1 else None
        end_time = _to_hhmm(row[2]) if len(row) > 2 else None

        if start_time is None and end_time is None:
            # 「休」「休(m/d)」などの休扱い行
            if any(p in col1 for p in _OFF_LABEL_PATTERNS) or code.startswith("休"):
                if code in seen_codes:
                    continue
                seen_codes.add(code)
                legend.append({
                    "code": code,
                    "label": col1 or code,
                    "start_time": "",
                    "end_time": "",
                    "break_minutes": 0,
                    "is_off": True,
                })
            continue

        if not start_time or not end_time:
            continue

        # 休憩は「開始」「終了」が直接入っている前提で分換算
        break_minutes = 0
        b_start = _to_hhmm(row[3]) if len(row) > 3 else None
        b_end = _to_hhmm(row[4]) if len(row) > 4 else None
        if b_start and b_end:
            try:
                sh, sm = map(int, b_start.split(":"))
                eh, em = map(int, b_end.split(":"))
                diff = (eh * 60 + em) - (sh * 60 + sm)
                if diff > 0:
                    break_minutes = diff
            except ValueError:
                pass

        # 重複コードは最初の登場分のみ採用
        if code in seen_codes:
            logger.warning("シフトルール内で重複コード '%s' を検出（後続行をスキップ）", code)
            continue
        seen_codes.add(code)

        legend.append({
            "code": code,
            "label": code,  # シフトルールには label 列が無いのでコード自身を使う
            "start_time": start_time,
            "end_time": end_time,
            "break_minutes": break_minutes,
            "is_off": False,
        })

    return legend


def is_legend_xlsx(filepath: str) -> bool:
    """シフトルール xlsx かどうかをファイル内容から判定する

    判定条件:
      - 単一シート
      - 行数が比較的少ない（<= 30）
      - 1 列目に A/B/C... のような 1-2 文字コードが並ぶ
      - 2-3 列目に時刻型セルがある
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    except Exception:
        return False

    if len(rows) > 30:
        return False

    code_count = 0
    time_count = 0
    for row in rows:
        if not row or row[0] is None:
            continue
        code = _clean_str(row[0])
        if 1 <= len(code) <= 3 and any(c.isalpha() or c == "休" for c in code):
            code_count += 1
        if len(row) > 1 and isinstance(row[1], time):
            time_count += 1
        elif len(row) > 1 and isinstance(row[1], datetime):
            time_count += 1
    return code_count >= 3 and time_count >= 3


# =============================================================================
# 2. 多年度横並び シフト表 xlsx パーサ
# =============================================================================

@dataclass
class _Section:
    """日付行を連続セクション（次の日に増分する範囲）に分けた1ブロック"""
    index: int           # 1-based
    start_col: int       # 1-based
    end_col: int         # 1-based, inclusive
    dates: dict[int, date]  # col_idx (1-based) -> date


def _find_date_row(rows: list[tuple]) -> int:
    """日付行の 0-indexed 行番号を返す。ヒューリスティック:
    日付シリアルに変換できるセルが最も多い行を選ぶ。
    """
    best_idx = -1
    best_count = 0
    for i, row in enumerate(rows[:10]):  # 上位 10 行のみ
        if not row:
            continue
        cnt = sum(1 for v in row if _to_date(v) is not None)
        if cnt > best_count:
            best_count = cnt
            best_idx = i
    if best_count < 30:  # 最低 30 セルくらいは欲しい
        raise ValueError("日付行が見つかりません（30セル以上の日付シリアル列が無い）")
    return best_idx


def _split_date_sections(date_row: tuple) -> list[_Section]:
    """日付行を連続日付のセクションに分割する。日付が前のセルの翌日でないところで切る。"""
    sections: list[_Section] = []
    sec_start_col: int | None = None
    sec_dates: dict[int, date] = {}
    last_date: date | None = None

    for i, v in enumerate(date_row):
        col_1based = i + 1
        d = _to_date(v)
        if d is None:
            # 文字列「2/29」など → 連続性を切らずにスキップ
            continue
        if last_date is None or (d - last_date).days != 1:
            # 新セクション開始
            if sec_dates:
                sections.append(_Section(
                    index=len(sections) + 1,
                    start_col=sec_start_col,
                    end_col=last_col,
                    dates=sec_dates,
                ))
            sec_start_col = col_1based
            sec_dates = {}
        sec_dates[col_1based] = d
        last_date = d
        last_col = col_1based

    if sec_dates:
        sections.append(_Section(
            index=len(sections) + 1,
            start_col=sec_start_col,
            end_col=last_col,
            dates=sec_dates,
        ))

    return sections


def _score_section_weekdays(
    section: _Section,
    weekday_row: tuple,
    target_year: int,
    target_month: int,
) -> tuple[int, int]:
    """セクション内 target_year/month の曜日が weekday_row の表記と一致する数を返す

    Returns:
        (matched, total)
    """
    matched = 0
    total = 0
    for col_1based, d in section.dates.items():
        if d.year != target_year or d.month != target_month:
            continue
        total += 1
        wd_actual = _WD_KANJI[d.weekday()]
        cell = weekday_row[col_1based - 1] if col_1based - 1 < len(weekday_row) else None
        wd_excel = _clean_str(cell)
        if wd_excel == wd_actual:
            matched += 1
    return matched, total


def _pick_target_section(
    sections: list[_Section],
    weekday_row: tuple,
    target_year: int,
    target_month: int,
) -> _Section | None:
    """target_year/month の曜日が最も一致するセクションを返す

    タイブレイク: rightmost (start_col が大きい) を優先（直近年度の想定）
    """
    candidates: list[tuple[float, int, _Section]] = []
    for sec in sections:
        # target year/month の日が 1 件でも含まれていなければスキップ
        contains = any(
            d.year == target_year and d.month == target_month
            for d in sec.dates.values()
        )
        if not contains:
            continue
        matched, total = _score_section_weekdays(sec, weekday_row, target_year, target_month)
        ratio = (matched / total) if total else 0.0
        candidates.append((ratio, sec.start_col, sec))

    if not candidates:
        return None

    # ratio 降順 → start_col 降順 (rightmost) でソート
    candidates.sort(key=lambda x: (-x[0], -x[1]))
    return candidates[0][2]


def _find_employee_rows(rows: list[tuple], data_start_row: int) -> list[int]:
    """従業員行の 0-indexed リストを返す

    判定: data_start_row 以降で col 0 が空でなく、コード値（D など）が
    複数個出現している行。
    """
    result: list[int] = []
    for i in range(data_start_row, len(rows)):
        row = rows[i]
        if not row:
            continue
        name = _clean_str(row[0])
        if not name:
            continue
        # コード値が 5 個以上ある行を従業員行とみなす（パターン列だけ埋まっている行を除外）
        code_count = 0
        for v in row[2:]:  # col 0 = name, col 1 = "シフトパターン" 想定
            s = _clean_str(v)
            if s and s != "シフトパターン":
                code_count += 1
                if code_count >= 5:
                    break
        if code_count >= 5:
            result.append(i)
    return result


def parse_multi_year_shift_xlsx(
    filepath: str,
    target_year: int,
    target_month: int,
) -> dict:
    """多年度横並びレイアウトのシフト表 xlsx を解析

    Args:
        filepath: シフト表 xlsx のパス
        target_year, target_month: 抽出対象（必須）

    Returns:
        {
          "year": int,
          "month": int,
          "filename": str,
          "employees": [{"name": str, "shifts": [{"date": "YYYY-MM-DD", "code": str}]}],
          "off_markers": list[str],
          "section_info": {...},  # デバッグ用: 採用したセクション
        }

    Raises:
        ValueError: レイアウトが該当しない / 該当セクションが見つからない場合
    """
    if not target_year or not target_month:
        raise ValueError("target_year / target_month は必須です")

    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("空のシートです")

    # 日付行と曜日行を特定
    date_row_idx = _find_date_row(rows)
    weekday_row_idx = date_row_idx + 1
    if weekday_row_idx >= len(rows):
        raise ValueError("曜日行が見つかりません")

    date_row = rows[date_row_idx]
    weekday_row = rows[weekday_row_idx]

    # 日付セクション分割
    sections = _split_date_sections(date_row)
    if not sections:
        raise ValueError("日付セクションが分割できません")

    target_section = _pick_target_section(sections, weekday_row, target_year, target_month)
    if target_section is None:
        raise ValueError(
            f"{target_year}年{target_month}月 を含むセクションが見つかりません "
            f"(検出セクション数: {len(sections)})"
        )

    # 採用セクションが本当に当該年度データかチェック
    matched, total = _score_section_weekdays(
        target_section, weekday_row, target_year, target_month
    )
    if total > 0 and matched / total < 0.5:
        logger.warning(
            "採用セクション section%d の曜日マッチ率が低いです (%d/%d)",
            target_section.index, matched, total,
        )

    # 従業員行を抽出
    employee_row_indices = _find_employee_rows(rows, weekday_row_idx + 1)
    if not employee_row_indices:
        raise ValueError("従業員行が見つかりません")

    days_in_month = calendar.monthrange(target_year, target_month)[1]
    target_dates = {
        col: d
        for col, d in target_section.dates.items()
        if d.year == target_year and d.month == target_month
    }

    employees: list[dict] = []
    for ridx in employee_row_indices:
        row = rows[ridx]
        name = _clean_str(row[0])
        if not name:
            continue
        # 「シフトパターン」のラベル列は無視
        shifts = []
        for col_1based, d in sorted(target_dates.items(), key=lambda x: x[1]):
            cell = row[col_1based - 1] if col_1based - 1 < len(row) else None
            code = _clean_str(cell)
            shifts.append({"date": d.isoformat(), "code": code})
        employees.append({"name": name, "shifts": shifts})

    return {
        "year": target_year,
        "month": target_month,
        "filename": os.path.basename(filepath),
        "employees": employees,
        "off_markers": ["休"],  # 「休」を含むセルはデフォルトで休扱い
        "section_info": {
            "section_index": target_section.index,
            "start_col": target_section.start_col,
            "end_col": target_section.end_col,
            "weekday_match_ratio": (matched / total) if total else 0.0,
            "weekday_matched": matched,
            "weekday_total": total,
            "total_sections": len(sections),
        },
    }


def is_multi_year_shift_xlsx(filepath: str) -> bool:
    """多年度横並びレイアウトに該当するか判定する

    判定条件:
      - 1 シート
      - 上位 5 行内に「日付シリアルが 100 セル以上ある行」が 1 つある
      - その下に曜日（土,日,月,...）が並ぶ行がある
      - 列方向の幅が 200 以上ある（短い1ヶ月シートではない）
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=1, max_row=5, values_only=True))
    except Exception:
        return False

    if not rows:
        return False
    max_width = max((len(r) for r in rows if r), default=0)
    if max_width < 200:
        return False

    has_long_date_row = False
    for row in rows:
        if not row:
            continue
        cnt = sum(1 for v in row if _to_date(v) is not None)
        if cnt >= 100:
            has_long_date_row = True
            break
    return has_long_date_row


# =============================================================================
# 3. 高レベルエントリ: 複数ファイルから「凡例 + 従業員シフト」を組み立てる
# =============================================================================

def parse_structured_files(
    filepaths: Iterable[str],
    target_year: int,
    target_month: int,
) -> tuple[list[dict], list[str]] | None:
    """複数のアップロードパスを sniff し、構造化解析できる xlsx だけを処理する

    多年度シフト表 .xlsx が 1 つも見つからなければ何もせず None を返す。
    見つかった場合は、そのシフト表 + マッチするシフトルール .xlsx をまとめて
    構造化パースし、consumed パス（後続の Claude フォールバックから除外
    すべきパス）と合わせて返す。

    Args:
        filepaths: アップロードされた timesheet 系ファイルのパス
        target_year, target_month: 抽出対象

    Returns:
        (sheets, consumed_paths) — sheets は code_sheets と同形式
        該当する多年度シフト表が無い場合は None
    """
    paths = list(filepaths)
    if not paths:
        return None

    legend_paths: list[str] = []
    shift_paths: list[str] = []
    for p in paths:
        if not p.lower().endswith((".xlsx", ".xls")):
            continue
        try:
            if is_multi_year_shift_xlsx(p):
                shift_paths.append(p)
            elif is_legend_xlsx(p):
                legend_paths.append(p)
        except Exception as e:
            logger.warning("xlsx sniff 失敗 %s: %s", p, e)

    if not shift_paths:
        return None

    # 凡例をマージ（複数渡された場合は重複コードを除外しつつ全件統合）
    legend: list[dict] = []
    seen: set[str] = set()
    for lp in legend_paths:
        try:
            for entry in parse_legend_xlsx(lp):
                if entry["code"] in seen:
                    continue
                seen.add(entry["code"])
                legend.append(entry)
        except Exception as e:
            logger.warning("シフトルール %s の解析に失敗: %s", lp, e)

    # シフト表ごとにシートを作る
    sheets: list[dict] = []
    consumed: list[str] = []
    for sp in shift_paths:
        try:
            result = parse_multi_year_shift_xlsx(sp, target_year, target_month)
        except Exception as e:
            logger.warning("シフト表 %s の構造化解析に失敗: %s", sp, e)
            continue
        sheets.append({
            "mode": "code",
            "filename": result["filename"],
            "legend": legend,
            "employees": result["employees"],
            "off_markers": result["off_markers"],
            "year": result["year"],
            "month": result["month"],
            "section_info": result.get("section_info"),
        })
        consumed.append(sp)

    if not sheets:
        return None

    # 構造化解析が成功したシフト表があれば、シフトルール .xlsx も消費扱いにする
    consumed.extend(legend_paths)
    return sheets, consumed
