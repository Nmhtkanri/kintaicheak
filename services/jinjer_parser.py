import pandas as pd
import re
from datetime import datetime, time
from config import Config


# 氏名から除去する注記パターン（例：「（例）」「(例)」など）
_NAME_ANNOTATION_RE = re.compile(r'[（(][^）)]*[）)]')


def _clean_name(name):
    """氏名から（例）などの注記を除去し、前後の空白を除く"""
    name = _NAME_ANNOTATION_RE.sub('', str(name))
    return name.strip()


def _find_column(df_columns, candidates):
    """
    候補カラム名リストから実際のカラム名を探す。
    完全一致を優先し、次に部分一致を試みる。
    """
    cols = [str(c).strip() for c in df_columns]
    # 1. 完全一致
    for candidate in candidates:
        for col in cols:
            if candidate == col:
                return col
    # 2. 部分一致
    for candidate in candidates:
        for col in cols:
            if candidate in col:
                return col
    return None


def _parse_time(value):
    """
    時刻文字列を datetime.time に変換。
    対応形式: HH:MM / H:MM / HH:MM:SS / H:MM:SS（秒は切り捨て）
    深夜跨ぎ（25:30:00 など、時 >= 24）は 24 で割った余りで処理。
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    value_str = str(value).strip()
    if not value_str or value_str in ("nan", "None"):
        return None

    # HH:MM:SS または H:MM:SS（秒付き）
    m = re.match(r'^(\d{1,2}):(\d{2}):\d{2}$', value_str)
    if m:
        hour, minute = int(m.group(1)), int(m.group(2))
        if hour >= 24:
            hour = hour % 24
        try:
            return time(hour, minute)
        except ValueError:
            return None

    # HH:MM または H:MM
    m = re.match(r'^(\d{1,2}):(\d{2})$', value_str)
    if m:
        hour, minute = int(m.group(1)), int(m.group(2))
        if hour >= 24:
            hour = hour % 24
        try:
            return time(hour, minute)
        except ValueError:
            return None

    return None


def _parse_duration_minutes(value):
    """HH:MM[:SS] の継続時間を分へ変換する（24時間超も保持）。"""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    value_str = str(value).strip()
    if not value_str or value_str in ("nan", "None"):
        return None
    m = re.fullmatch(r"(\d{1,3}):(\d{2})(?::(\d{2}))?", value_str)
    if not m:
        return None
    hours, minutes, seconds = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
    if minutes >= 60 or seconds >= 60:
        return None
    total = hours * 60 + minutes + round(seconds / 60)
    return total if total > 0 else None


def _parse_date(value):
    """日付を datetime.date に変換"""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, datetime):
        return value.date()
    value_str = str(value).strip()
    if not value_str or value_str in ("nan", "None"):
        return None
    for fmt in ["%Y/%m/%d", "%Y-%m-%d", "%m/%d/%Y", "%Y年%m月%d日"]:
        try:
            return datetime.strptime(value_str, fmt).date()
        except ValueError:
            continue
    return None


def _merge_comments(*values):
    """複数のコメント値を結合。空・nan は除外し、「 / 」で連結"""
    parts = []
    for v in values:
        if v and str(v).strip() not in ("", "nan", "None"):
            parts.append(str(v).strip())
    return " / ".join(parts) if parts else None


def _read_csv_auto_encoding(filepath, **kwargs):
    """utf-8 → cp932 → utf-8-sig の順にエンコーディングを試して CSV を読み込む"""
    encodings = ["utf-8", "cp932", "utf-8-sig"]
    last_error = None
    for enc in encodings:
        try:
            return pd.read_csv(filepath, encoding=enc, **kwargs)
        except (UnicodeDecodeError, UnicodeError) as e:
            last_error = e
    raise ValueError(
        f"CSVの文字コードを自動判定できませんでした（試行: {encodings}）: {last_error}"
    )


def parse_jinjer_csv(filepath):
    """
    jinjer CSVを解析して統一 DataFrame を返す。

    Returns:
        DataFrame with columns: 氏名, 日付, 出勤時刻, 退勤時刻, コメント,
        データソース, 総労働時間(分)
    """
    # ヘッダー行を自動検出（氏名候補カラムが含まれる行を探す）
    raw = _read_csv_auto_encoding(filepath, header=None, dtype=str)

    header_row = 0
    all_name_candidates = Config.JINJER_COLUMN_MAPPING["氏名"]
    for i, row in raw.iterrows():
        row_values = [str(v).strip() for v in row.values]
        if any(candidate in row_values for candidate in all_name_candidates):
            header_row = i
            break

    df = _read_csv_auto_encoding(filepath, header=header_row, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    mapping = Config.JINJER_COLUMN_MAPPING
    col_name    = _find_column(df.columns, mapping["氏名"])
    col_date    = _find_column(df.columns, mapping["日付"])
    col_start   = _find_column(df.columns, mapping["出勤時刻"])
    col_end     = _find_column(df.columns, mapping["退勤時刻"])
    col_total   = "総労働時間" if "総労働時間" in df.columns else None
    col_scheduled_start = _find_column(df.columns, ["出勤予定時刻"])
    col_scheduled_end = _find_column(df.columns, ["退勤予定時刻"])
    col_comment = _find_column(df.columns, mapping["コメント"])
    col_comment2 = _find_column(df.columns, mapping.get("コメント2", []))

    if col_name is None or col_date is None:
        raise ValueError(
            f"必須カラム（氏名・日付）が見つかりません。"
            f"検出されたカラム: {list(df.columns[:20])}..."
        )

    records = []
    for _, row in df.iterrows():
        raw_name = row.get(col_name, "") if col_name else ""
        name = _clean_name(raw_name)

        date_val = _parse_date(row.get(col_date)) if col_date else None
        start    = _parse_time(row.get(col_start)) if col_start else None
        end      = _parse_time(row.get(col_end))   if col_end   else None

        c1 = row.get(col_comment) if col_comment else None
        c2 = row.get(col_comment2) if col_comment2 else None
        comment = _merge_comments(c1, c2)

        scheduled_start = _parse_time(row.get(col_scheduled_start)) if col_scheduled_start else None
        scheduled_end = _parse_time(row.get(col_scheduled_end)) if col_scheduled_end else None

        # スキップ条件: 氏名なし / 日付なし / 実績も予定もない（休日行など）
        if not name or name in ("nan", "None") or date_val is None:
            continue
        if start is None and end is None and scheduled_start is None and scheduled_end is None:
            continue

        records.append({
            "氏名": name,
            "日付": date_val,
            "出勤時刻": start,
            "退勤時刻": end,
            "コメント": comment,
            "データソース": "jinjer",
            "総労働時間(分)": _parse_duration_minutes(row.get(col_total)) if col_total else None,
        })

    return pd.DataFrame(records, columns=["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース", "総労働時間(分)"])
