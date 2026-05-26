"""jinjer スケジュール雛形マッチング & 新規雛形 CSV 生成

- 既存の `スケジュール雛形一覧.csv` を読み込み、解決済みの記号→時刻と照合する
- 完全一致しない記号があれば「新規雛形候補」として CSV を生成する
  → ユーザーはそのまま jinjer にインポートできる
"""

from __future__ import annotations

import csv
import logging
import os
import re
from datetime import time
from typing import Iterable

logger = logging.getLogger(__name__)


# jinjer スケジュール雛形 CSV のヘッダー定義（_2026-04-27 の新フォーマットに準拠）
# ※新規雛形 CSV を書き出すときはこのヘッダーで出力する（jinjer の現行インポート仕様）
TEMPLATE_CSV_HEADERS = [
    "No",
    "＊スケジュール雛形名",
    "略称(3文字以内)",
    "＊スケジュール雛形ID",
    "表示順",
    "半休ID",
    "＊出勤時間(0:00~47:59)",
    "＊退勤時間(0:00~47:59)",
    "休憩開始時間1(0:00~47:59)",
    "復帰時間1(0:00~47:59)",
    "休憩時間2(0:00~47:59)",
    "復帰時間2(0:00~47:59)",
    "休憩時間3(0:00~47:59)",
    "復帰時間3(0:00~47:59)",
    "休憩時間4(0:00~47:59)",
    "復帰時間4(0:00~47:59)",
    "休憩時間5(0:00~47:59)",
    "復帰時間5(0:00~47:59)",
    "スケジュール外休憩開始時間(0:00~47:59)",
    "スケジュール外復帰時間(0:00~47:59)",
]

# 旧フォーマット（_2026-04-24 以前）→ 新フォーマット への canonical キー対応表
# 既存雛形 CSV を読むときに、どちらのフォーマットでも _tpl_get で値が拾えるようにする
TEMPLATE_HEADER_ALIASES = {
    "＊スケジュール雛形名": ["勤怠スケジュール名称"],
    "＊スケジュール雛形ID": ["勤怠スケジュール名ID"],
    "表示順": ["予定区分"],
    "半休ID": ["休日ID"],
    "＊出勤時間(0:00~47:59)": ["出勤時刻(0:00~47:59)"],
    "＊退勤時間(0:00~47:59)": ["退勤時刻(0:00~47:59)"],
    "スケジュール外休憩開始時間(0:00~47:59)": ["スケジュール前休憩開始時間(0:00~47:59)"],
    "スケジュール外復帰時間(0:00~47:59)": ["スケジュール前復帰時間(0:00~47:59)"],
}


def _tpl_get(tpl: dict, canonical_key: str, default: str = "") -> str:
    """雛形 dict から canonical キーで値を取る。新旧フォーマットどちらでも拾えるようエイリアスも見る。"""
    val = tpl.get(canonical_key)
    if val is not None and str(val).strip() != "":
        return val
    for alias in TEMPLATE_HEADER_ALIASES.get(canonical_key, []):
        v = tpl.get(alias)
        if v is not None and str(v).strip() != "":
            return v
    return default


def _time_to_str(t):
    """time → 'HH:MM:SS'。None なら空文字"""
    if t is None:
        return ""
    if isinstance(t, str):
        # 既に文字列ならそのまま（HH:MM 形式想定）
        if len(t.split(":")) == 2:
            return t + ":00"
        return t
    if isinstance(t, time):
        return t.strftime("%H:%M:%S")
    return ""


def _normalize_time_str(s: str) -> str:
    """雛形CSV内の様々な時刻表記を 'H:MM' 形式に揃える"""
    if not s:
        return ""
    s = s.strip()
    # "8:00:00" → "8:00"
    parts = s.split(":")
    if len(parts) >= 2:
        try:
            h = int(parts[0])
            m = int(parts[1])
            return f"{h}:{m:02d}"
        except ValueError:
            return s
    return s


def _clock_minutes(value: str) -> int | None:
    """'H:MM' / 'HH:MM' を 24h 超表記のまま分へ変換する。"""
    if not value:
        return None
    parts = str(value).strip().split(":")
    if len(parts) < 2:
        return None
    try:
        hour = int(parts[0])
        minute = int(parts[1])
    except ValueError:
        return None
    if hour < 0 or minute < 0 or minute > 59:
        return None
    return hour * 60 + minute


def _minutes_to_time_str(minutes: int) -> str:
    return f"{minutes // 60}:{minutes % 60:02d}"


def canonicalize_overnight_times(start_time, end_time) -> tuple[str, str]:
    """画像読み取りの深夜退勤を jinjer の 24h 超表記へ寄せる。"""
    start = _normalize_time_str(_time_to_str(start_time))
    end = _normalize_time_str(_time_to_str(end_time))
    start_minutes = _clock_minutes(start)
    end_minutes = _clock_minutes(end)
    if start_minutes is None or end_minutes is None:
        return start, end

    while end_minutes < start_minutes:
        end_minutes += 24 * 60
    return _minutes_to_time_str(start_minutes), _minutes_to_time_str(end_minutes)


def suggest_template_id(code: str) -> str:
    """新規雛形 CSV と月次 CSV で共有する ID 候補を返す。"""
    return re.sub(r"[^A-Za-z0-9]", "", str(code or "").strip()) or "X"


def load_jinjer_templates(csv_path: str) -> list[dict]:
    """jinjer スケジュール雛形 CSV を読み込む

    Args:
        csv_path: CSV ファイルパス（Shift_JIS / CP932 想定）

    Returns:
        list of dict（各行）。読めない場合は空リスト。
    """
    if not csv_path or not os.path.exists(csv_path):
        logger.warning("jinjer 雛形 CSV が見つかりません: %s", csv_path)
        return []

    rows = []
    # まず CP932 を試し、ダメなら UTF-8
    for encoding in ("cp932", "utf-8-sig", "utf-8"):
        try:
            with open(csv_path, "r", encoding=encoding, newline="") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    rows.append(row)
            logger.info("雛形 CSV 読み込み: %s 件 (encoding=%s)", len(rows), encoding)
            return rows
        except UnicodeDecodeError:
            continue
        except Exception as e:
            logger.error("雛形 CSV 読み込みエラー: %s", e)
            return []

    logger.error("雛形 CSV のエンコーディングを判別できませんでした")
    return []


def find_matching_template(
    start_time: str | time,
    end_time: str | time,
    break_minutes: int,
    templates: list[dict],
) -> dict | None:
    """指定された出退勤時刻に完全一致する雛形を探す

    マッチ条件:
      - 出勤時刻 完全一致
      - 退勤時刻 完全一致
      （現状は 休憩は照合に含めない: 休憩時間表記が雛形によって異なるため）

    Returns:
        マッチした雛形 dict、無ければ None
    """
    target_start, target_end = canonicalize_overnight_times(start_time, end_time)
    if not target_start or not target_end:
        return None

    for tpl in templates:
        tpl_start = _normalize_time_str(_tpl_get(tpl, "＊出勤時間(0:00~47:59)"))
        tpl_end = _normalize_time_str(_tpl_get(tpl, "＊退勤時間(0:00~47:59)"))
        if tpl_start == target_start and tpl_end == target_end:
            return tpl
    return None


def match_legend_to_templates(
    legend: list[dict],
    csv_path: str,
) -> dict:
    """凡例 × 既存雛形 を照合し、マッチ結果と新規候補を返す

    Args:
        legend: shift_resolver.normalize_legend に渡す形式の凡例リスト
        csv_path: jinjer 雛形 CSV のパス

    Returns:
        {
          "matched":   [{code, label, template_no, template_name, start, end, ...}],
          "unmatched": [{code, label, start, end, break_minutes}],
          "templates_total": int,
        }
    """
    templates = load_jinjer_templates(csv_path)
    matched, unmatched = [], []

    for entry in legend:
        if not isinstance(entry, dict):
            continue
        if entry.get("is_off"):
            continue  # 休扱いは雛形マッチ対象外
        code = entry.get("code")
        if not code:
            continue
        start = entry.get("start_time")
        end = entry.get("end_time")
        if not start or not end:
            continue

        start_norm, end_norm = canonicalize_overnight_times(start, end)
        tpl = find_matching_template(start, end, entry.get("break_minutes", 0), templates)
        if tpl:
            matched.append({
                "code": code,
                "label": entry.get("label") or code,
                "start_time": start_norm,
                "end_time": end_norm,
                "template_no": tpl.get("No"),
                "template_name": _tpl_get(tpl, "＊スケジュール雛形名"),
                "template_id": _tpl_get(tpl, "＊スケジュール雛形ID"),
            })
        else:
            unmatched.append({
                "code": code,
                "label": entry.get("label") or code,
                "start_time": start_norm,
                "end_time": end_norm,
                "break_minutes": entry.get("break_minutes", 0),
            })

    return {
        "matched": matched,
        "unmatched": unmatched,
        "templates_total": len(templates),
    }


def _next_template_no(existing_templates: list[dict]) -> int:
    """既存雛形の最大 No+1 を返す"""
    max_no = 0
    for tpl in existing_templates:
        try:
            n = int(tpl.get("No") or 0)
            if n > max_no:
                max_no = n
        except (ValueError, TypeError):
            pass
    return max_no + 1


def _format_csv_time(value) -> str:
    """雛形 CSV に書き込む時刻フォーマット ('H:MM:SS')"""
    s = _time_to_str(value)
    return s


def _calc_break_window(start_str: str, break_minutes: int) -> tuple[str, str]:
    """休憩 1 の開始/終了時刻を簡易計算する

    出勤から 3〜4 時間後を昼休憩にするのが既存雛形の慣例。
    安全策として「出勤+3:30」スタートで break_minutes 分とる。
    """
    if not start_str or break_minutes <= 0:
        return "", ""
    try:
        h, m, *_ = start_str.split(":")
        start_min = int(h) * 60 + int(m)
        break_start = start_min + 3 * 60 + 30   # 出勤 +3:30
        break_end = break_start + break_minutes
        return (
            f"{break_start // 60}:{break_start % 60:02d}:00",
            f"{break_end // 60}:{break_end % 60:02d}:00",
        )
    except ValueError:
        return "", ""


def generate_new_templates_csv(
    unmatched: list[dict],
    csv_path_existing: str,
    output_path: str,
) -> dict:
    """マッチしなかった記号から新規雛形 CSV を生成する

    Args:
        unmatched: match_legend_to_templates の "unmatched" リスト
        csv_path_existing: 既存雛形 CSV (No の重複回避用)
        output_path: 出力先 CSV パス

    Returns:
        {"path": output_path, "rows": list[dict], "count": int}
    """
    if not unmatched:
        return {"path": None, "rows": [], "count": 0}

    existing = load_jinjer_templates(csv_path_existing)
    next_no = _next_template_no(existing)

    rows = []
    used_ids = {_tpl_get(tpl, "＊スケジュール雛形ID") for tpl in existing}
    for entry in unmatched:
        code = str(entry.get("code") or "").strip()
        label = entry.get("label") or code
        start, end = canonicalize_overnight_times(
            entry.get("start_time") or "",
            entry.get("end_time") or "",
        )
        break_minutes = int(entry.get("break_minutes") or 0)

        # ID 重複回避: code をベースに連番を振る
        base_id = suggest_template_id(code)
        candidate = base_id
        suffix = 1
        while candidate in used_ids:
            suffix += 1
            candidate = f"{base_id}{suffix}"
        used_ids.add(candidate)

        # 略称: 3 文字以内
        short_name = code[:3] if code else label[:3]

        b_start, b_end = _calc_break_window(_format_csv_time(start), break_minutes)

        row = {h: "" for h in TEMPLATE_CSV_HEADERS}
        row["No"] = str(next_no)
        row["＊スケジュール雛形名"] = label
        row["略称(3文字以内)"] = short_name
        row["＊スケジュール雛形ID"] = candidate
        row["表示順"] = "9998"     # 既存運用に倣う
        row["半休ID"] = "1"
        row["＊出勤時間(0:00~47:59)"] = _format_csv_time(start)
        row["＊退勤時間(0:00~47:59)"] = _format_csv_time(end)
        row["休憩開始時間1(0:00~47:59)"] = b_start
        row["復帰時間1(0:00~47:59)"] = b_end
        rows.append(row)
        next_no += 1

    # 出力先ディレクトリを確保
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    # jinjer は CP932 想定なので CP932 で書き出す
    with open(output_path, "w", encoding="cp932", newline="", errors="replace") as f:
        writer = csv.DictWriter(f, fieldnames=TEMPLATE_CSV_HEADERS)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)

    logger.info("新規雛形 CSV 出力: %s (%d 件)", output_path, len(rows))
    return {"path": output_path, "rows": rows, "count": len(rows)}
