"""シフト記号 → 時刻の解決ロジック（純粋関数）

Claude APIで抽出した「凡例」と「シフト表（記号入り）」を受け取り、
既存の matcher.py に流せる DataFrame 形式に変換する。

主な責務:
1. 記号 → 出退勤時刻 のマッピング解決
2. 「明け」「振」「休」など休日扱い記号のスキップ
3. 連続するシフトの統合（例: 田村さん 4日 16:30-24:00 + 5日 24:00-翌09:00 → 4日 16:30-33:00）
"""

from __future__ import annotations

import re
import logging
from datetime import datetime, date, time, timedelta
from typing import Iterable

import pandas as pd

logger = logging.getLogger(__name__)


# matcher.py / 既存パイプラインと完全一致させる列構成
RECORD_COLUMNS = ["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"]

# デフォルトで「休扱い」とみなす記号・空欄パターン
DEFAULT_OFF_MARKERS = {"", "—", "ー", "-", "休", "公休", "週休", "公", "×", "✕", "\u3000"}


def _parse_time_str(value):
    """'12:30' '25:00' '33:30' などを datetime.time に変換（>=24h は %24 で正規化）

    既存の timesheet_parser._parse_time_str と整合性を取る。
    """
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() == "null":
        return None
    m = re.match(r"^(\d{1,2}):(\d{2})$", s)
    if not m:
        return None
    hour = int(m.group(1))
    minute = int(m.group(2))
    if minute > 59:
        return None
    # >= 24:00 は次日扱いとして %24 で正規化
    hour = hour % 24
    try:
        return time(hour, minute)
    except ValueError:
        return None


def _parse_date_str(value):
    """'2026-04-01' などを date に変換"""
    if value is None:
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def normalize_legend(raw_legend) -> dict:
    """凡例リスト → {code: entry} dict に正規化

    Args:
        raw_legend: list of {"code": str, "label": str, "start_time": str|None,
                              "end_time": str|None, "break_minutes": int,
                              "is_off": bool, ...}

    Returns:
        {code(str): {label, start_time(time|None), end_time(time|None),
                      break_minutes(int), is_off(bool)}}
    """
    legend = {}
    if not raw_legend:
        return legend

    for entry in raw_legend:
        if not isinstance(entry, dict):
            continue
        code = entry.get("code")
        if code is None:
            continue
        code = str(code).strip()
        if code == "":
            continue

        is_off = bool(entry.get("is_off"))
        start = _parse_time_str(entry.get("start_time"))
        end = _parse_time_str(entry.get("end_time"))
        if start is None and end is None:
            # 時刻が両方無い ＝ 暗黙的に休扱い
            is_off = True

        try:
            break_minutes = int(entry.get("break_minutes") or 0)
        except (TypeError, ValueError):
            break_minutes = 0

        legend[code] = {
            "label": entry.get("label") or code,
            "start_time": start,
            "end_time": end,
            "break_minutes": break_minutes,
            "is_off": is_off,
        }
    return legend


def _is_off_code(code, legend, off_markers):
    """その記号は休日扱いか"""
    if code is None:
        return True
    c = str(code).strip()
    if c in off_markers:
        return True
    entry = legend.get(c)
    if entry and entry.get("is_off"):
        return True
    return False


def _times_equal_at_midnight(t1, t2):
    """両方とも 24:00 / 00:00 を表現していて、結合の境目になっているか"""
    if t1 is None or t2 is None:
        return False
    return t1 == time(0, 0) and t2 == time(0, 0)


def _merge_consecutive_overnight(records: list[dict]) -> list[dict]:
    """連続日のシフトを統合する

    例: 田村さん
      day 4: 16:30 - 24:00 (4)
      day 5: 24:00 - 09:00 (a)  ← 翌日扱いの a
      → 統合: day 4: 16:30 - 09:00（実態は 33:00）

    判定条件:
      - 同一氏名
      - 日付が連続（diff = 1日）
      - record N の 退勤時刻 が 00:00（=24:00 を %24 した結果）
      - record N+1 の 出勤時刻 が 00:00（=24:00 を %24 した結果）

    Args:
        records: 各 dict は RECORD_COLUMNS を持つ

    Returns:
        統合済みリスト（in-placeではなく新リスト）
    """
    if len(records) < 2:
        return list(records)

    # 氏名→ソート済みリスト に分ける
    by_name: dict[str, list[dict]] = {}
    for rec in records:
        by_name.setdefault(rec.get("氏名", ""), []).append(rec)
    for name, recs in by_name.items():
        recs.sort(key=lambda r: r.get("日付") or date.min)

    merged_all: list[dict] = []
    for name, recs in by_name.items():
        i = 0
        while i < len(recs):
            current = dict(recs[i])
            # 次のレコードと結合できるか確認
            while i + 1 < len(recs):
                nxt = recs[i + 1]
                d_cur = current.get("日付")
                d_nxt = nxt.get("日付")
                if not (isinstance(d_cur, date) and isinstance(d_nxt, date)):
                    break
                if (d_nxt - d_cur) != timedelta(days=1):
                    break
                if not _times_equal_at_midnight(
                    current.get("退勤時刻"), nxt.get("出勤時刻")
                ):
                    break
                # 結合: 退勤時刻 を nxt の退勤に置き換える
                current["退勤時刻"] = nxt.get("退勤時刻")
                # コメントも結合（あれば）
                merged_comments = [
                    c for c in [current.get("コメント"), nxt.get("コメント")] if c
                ]
                current["コメント"] = " / ".join(merged_comments) if merged_comments else None
                i += 1  # 次のレコードを消費
            merged_all.append(current)
            i += 1

    # 元の DataFrame に近い順序で並び替え（氏名 → 日付）
    merged_all.sort(key=lambda r: (r.get("氏名", ""), r.get("日付") or date.min))
    return merged_all


def resolve_shifts(
    legend_raw,
    employees,
    *,
    off_markers: Iterable[str] | None = None,
    source_label: str = "勤務表",
    merge_overnight: bool = True,
) -> pd.DataFrame:
    """凡例 × シフト表 → matcher.py 用の DataFrame に解決する

    Args:
        legend_raw: list of legend entries（normalize_legend 入力）
        employees: list of {
            "name": str,
            "shifts": [{"date": "YYYY-MM-DD", "code": str, "comment": str|None}, ...]
        }
        off_markers: 追加で休扱いとみなす記号セット
        source_label: 出力 DataFrame の "データソース" カラム値
        merge_overnight: True の場合、連続日の統合（24:00跨ぎ）を行う

    Returns:
        pd.DataFrame（columns = RECORD_COLUMNS）
    """
    legend = normalize_legend(legend_raw)
    markers = set(DEFAULT_OFF_MARKERS)
    if off_markers:
        for m in off_markers:
            if m is None:
                continue
            markers.add(str(m).strip())

    records: list[dict] = []
    if not employees:
        return pd.DataFrame(columns=RECORD_COLUMNS)

    for emp in employees:
        if not isinstance(emp, dict):
            continue
        name = emp.get("name") or emp.get("employee_name") or "不明"
        shifts = emp.get("shifts") or []
        for shift in shifts:
            if not isinstance(shift, dict):
                continue
            d = _parse_date_str(shift.get("date"))
            if d is None:
                continue
            raw_code = shift.get("code")
            code = str(raw_code).strip() if raw_code is not None else ""

            if _is_off_code(code, legend, markers):
                continue  # 休日はレコードを作らない

            entry = legend.get(code)
            if entry is None:
                # 凡例に無い記号はスキップ（ログに残す）
                logger.info("凡例に無い記号をスキップ: name=%s date=%s code=%r", name, d, code)
                continue

            start = entry.get("start_time")
            end = entry.get("end_time")
            if start is None and end is None:
                continue

            label = entry.get("label") or code
            comment_parts = [f"[{code}={label}]"]
            shift_comment = shift.get("comment")
            if shift_comment:
                comment_parts.append(str(shift_comment))
            comment = " ".join(comment_parts) if comment_parts else None

            records.append({
                "氏名": name,
                "日付": d,
                "出勤時刻": start,
                "退勤時刻": end,
                "コメント": comment,
                "データソース": source_label,
            })

    if merge_overnight:
        records = _merge_consecutive_overnight(records)

    if not records:
        return pd.DataFrame(columns=RECORD_COLUMNS)

    return pd.DataFrame(records, columns=RECORD_COLUMNS)
