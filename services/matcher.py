import re
import unicodedata
from datetime import datetime, date, time, timedelta
import pandas as pd


def normalize_name(name):
    """氏名の正規化"""
    if not name or pd.isna(name):
        return ""
    name = str(name)
    name = unicodedata.normalize("NFKC", name)
    name = name.strip()
    name = name.replace("\u3000", " ")  # 全角スペース→半角スペース
    name = re.sub(r'[,\u3001\uff0c]', '', name)  # カンマ類を除去
    name = re.sub(r'\s+', '', name)    # スペース全除去
    return name


def normalize_name_keys(name):
    """氏名の照合候補キーを返す（SAPの「姓, 名」と逆順登録を吸収）"""
    if not name or pd.isna(name):
        return [""]

    raw = unicodedata.normalize("NFKC", str(name)).strip().replace("\u3000", " ")
    keys = [normalize_name(raw)]

    comma_parts = [p.strip() for p in re.split(r'[,\u3001\uff0c]', raw) if p.strip()]
    if len(comma_parts) == 2:
        keys.append(normalize_name("".join(reversed(comma_parts))))

    space_parts = [p.strip() for p in re.split(r'\s+', raw) if p.strip()]
    if len(space_parts) == 2:
        keys.append(normalize_name("".join(reversed(space_parts))))

    return list(dict.fromkeys(k for k in keys if k))


def _is_probable_staff_code(name):
    """画像解析が氏名の代わりに拾いやすいスタッフコードかを判定する。"""
    if not name or pd.isna(name):
        return False
    raw = unicodedata.normalize("NFKC", str(name)).strip()
    return bool(re.fullmatch(r"[A-Z0-9_-]{2,12}", raw))


def _to_minutes(t):
    """datetime.timeを分に変換"""
    if t is None or (not isinstance(t, time)):
        return None
    return t.hour * 60 + t.minute


def time_diff_minutes(t1, t2):
    """2つのtimeの差分を分単位で返す（t1 - t2）"""
    m1 = _to_minutes(t1)
    m2 = _to_minutes(t2)
    if m1 is None or m2 is None:
        return None
    return m1 - m2


def _is_midnight(value):
    return isinstance(value, time) and value.hour == 0 and value.minute == 0


def _next_date(value):
    if isinstance(value, datetime):
        return value.date() + timedelta(days=1)
    if isinstance(value, date):
        return value + timedelta(days=1)
    return None


def _split_jinjer_overnight_rows(jinjer_df, timesheet_df):
    """SAP側が日跨ぎ勤務を2日分に分けている場合、jinjer側も突合用に2分割する。"""
    if jinjer_df.empty or timesheet_df.empty:
        return jinjer_df

    sheet_keys = set()
    for _, row in timesheet_df.iterrows():
        key = row.get("氏名_normalized")
        work_date = row.get("日付")
        if not key or pd.isna(work_date):
            continue
        if _is_midnight(row.get("勤務表_退勤時刻")):
            sheet_keys.add((key, work_date, "first"))
        if _is_midnight(row.get("勤務表_出勤時刻")):
            sheet_keys.add((key, work_date, "second"))

    rows = []
    for _, row in jinjer_df.iterrows():
        start = row.get("jinjer_出勤時刻")
        end = row.get("jinjer_退勤時刻")
        work_date = row.get("日付")
        next_work_date = _next_date(work_date)
        is_overnight = (
            isinstance(start, time)
            and isinstance(end, time)
            and _to_minutes(start) > _to_minutes(end)
            and next_work_date is not None
        )
        should_split = is_overnight and (
            (row.get("氏名_normalized"), work_date, "first") in sheet_keys
            and (row.get("氏名_normalized"), next_work_date, "second") in sheet_keys
        )

        if not should_split:
            rows.append(row.to_dict())
            continue

        first = row.to_dict()
        first["jinjer_退勤時刻"] = time(0, 0)
        rows.append(first)

        second = row.to_dict()
        second["日付"] = next_work_date
        second["jinjer_出勤時刻"] = time(0, 0)
        rows.append(second)

    return pd.DataFrame(rows, columns=jinjer_df.columns)


def _work_segment_key(start, end):
    """同じ日付内の夜勤前半/後半が混ざらないようにする内部キー。"""
    if _is_midnight(start) and not _is_midnight(end):
        return "after_midnight"
    if _is_midnight(end) and not _is_midnight(start):
        return "before_midnight"
    return "normal"


def judge(row, threshold_minutes=10):
    """
    各レコードの判定を行う

    Returns:
        tuple(判定, 詳細)
    """
    jinjer_start = row.get("jinjer_出勤時刻")
    jinjer_end = row.get("jinjer_退勤時刻")
    sheet_start = row.get("勤務表_出勤時刻")
    sheet_end = row.get("勤務表_退勤時刻")

    def _clean(v):
        if isinstance(v, time):
            return v
        try:
            if pd.isna(v):
                return None
        except (TypeError, ValueError):
            pass
        return None

    jinjer_start = _clean(jinjer_start)
    jinjer_end = _clean(jinjer_end)
    sheet_start = _clean(sheet_start)
    sheet_end = _clean(sheet_end)

    has_jinjer_row = pd.notna(row.get("jinjer_氏名")) and str(row.get("jinjer_氏名")) not in ["", "nan", "None"]

    # 1. 片方にしかデータがない場合
    if jinjer_start is None and jinjer_end is None and sheet_start is None and sheet_end is None:
        if has_jinjer_row:
            return "データ欠損", "jinjer実績なし・勤務表側に勤怠なし"
        return "データ欠損", "jinjer側にデータなし"
    if jinjer_start is None and jinjer_end is None:
        if has_jinjer_row:
            return "データ欠損", "jinjer実績なし・勤務表側で勤怠あり"
        return "データ欠損", "jinjer側にデータなし"
    if sheet_start is None and sheet_end is None:
        return "データ欠損", "勤務表側にデータなし"

    # 2. 差分計算
    start_diff = time_diff_minutes(sheet_start, jinjer_start)
    end_diff = time_diff_minutes(sheet_end, jinjer_end)

    start_diff_abs = abs(start_diff) if start_diff is not None else 0
    end_diff_abs = abs(end_diff) if end_diff is not None else 0
    max_diff = max(start_diff_abs, end_diff_abs)

    # 3. 判定
    if max_diff <= threshold_minutes:
        return "OK", ""

    has_comment = bool(
        (row.get("jinjer_コメント") and str(row.get("jinjer_コメント")) not in ["", "nan", "None"]) or
        (row.get("勤務表_コメント") and str(row.get("勤務表_コメント")) not in ["", "nan", "None"])
    )

    comment = row.get("jinjer_コメント") or row.get("勤務表_コメント") or ""

    if has_comment:
        return "要確認", f"差分{max_diff}分 / コメント: {comment}"
    else:
        return "NG", f"差分{max_diff}分 / コメントなし"


def match(jinjer_df, timesheet_df, threshold_minutes=10):
    """
    jinjer DataFrameと勤務表DataFrameを突合する

    突合対象は勤務表に含まれる社員のみ。
    jinjer側にのみ存在する社員は「未提出リスト」として返す。

    Returns:
        tuple(result_df, unsubmitted_names)
        - result_df: 突合結果 DataFrame
        - unsubmitted_names: 勤務表未提出の社員名リスト
    """
    # 氏名正規化
    jinjer_df = jinjer_df.copy()
    timesheet_df = timesheet_df.copy()
    jinjer_df["氏名_normalized"] = jinjer_df["氏名"].apply(normalize_name)
    timesheet_df["氏名_normalized"] = timesheet_df["氏名"].apply(normalize_name)
    jinjer_name_keys = set(jinjer_df["氏名_normalized"].unique())

    def choose_sheet_key(name):
        for key in normalize_name_keys(name):
            if key in jinjer_name_keys:
                return key
        keys = normalize_name_keys(name)
        return keys[0] if keys else ""

    timesheet_df["氏名_normalized"] = timesheet_df["氏名"].apply(choose_sheet_key)

    # OCR/画像解析で勤務表上部のスタッフコードだけが氏名として拾われる場合がある。
    # 単一人物同士で曖昧さがないときだけ、勤務表側コードをjinjer側の氏名に寄せる。
    jinjer_unique_keys = [k for k in dict.fromkeys(jinjer_df["氏名_normalized"].tolist()) if k]
    jinjer_unique_names = [n for n in dict.fromkeys(jinjer_df["氏名"].tolist()) if pd.notna(n) and str(n).strip()]
    sheet_unique_names = [n for n in dict.fromkeys(timesheet_df["氏名"].tolist()) if pd.notna(n) and str(n).strip()]
    sheet_unique_keys = [k for k in dict.fromkeys(timesheet_df["氏名_normalized"].tolist()) if k]
    if (
        len(jinjer_unique_keys) == 1
        and len(jinjer_unique_names) == 1
        and len(sheet_unique_names) == 1
        and len(sheet_unique_keys) == 1
        and sheet_unique_keys[0] not in jinjer_name_keys
        and _is_probable_staff_code(sheet_unique_names[0])
    ):
        timesheet_df["氏名_normalized"] = jinjer_unique_keys[0]
        timesheet_df["氏名"] = jinjer_unique_names[0]

    # 勤務表に含まれる社員の正規化名セット
    sheet_names = set(timesheet_df["氏名_normalized"].unique())

    # jinjer側にのみ存在する社員（勤務表未提出）を抽出
    jinjer_names = set(jinjer_df["氏名_normalized"].unique())
    unsubmitted_normalized = jinjer_names - sheet_names

    # 未提出社員の元の氏名を取得（重複なし・順序保持）
    unsubmitted_rows = jinjer_df[jinjer_df["氏名_normalized"].isin(unsubmitted_normalized)]
    unsubmitted_names = list(dict.fromkeys(unsubmitted_rows["氏名"].tolist()))

    # jinjer側を勤務表の社員のみにフィルタリング
    jinjer_filtered = jinjer_df[jinjer_df["氏名_normalized"].isin(sheet_names)]

    # カラムリネーム
    jinjer_renamed = jinjer_filtered.rename(columns={
        "氏名": "jinjer_氏名",
        "出勤時刻": "jinjer_出勤時刻",
        "退勤時刻": "jinjer_退勤時刻",
        "コメント": "jinjer_コメント",
    })

    sheet_renamed = timesheet_df.rename(columns={
        "氏名": "勤務表_氏名",
        "出勤時刻": "勤務表_出勤時刻",
        "退勤時刻": "勤務表_退勤時刻",
        "コメント": "勤務表_コメント",
    })

    jinjer_renamed = _split_jinjer_overnight_rows(jinjer_renamed, sheet_renamed)
    jinjer_renamed["突合区分"] = jinjer_renamed.apply(
        lambda r: _work_segment_key(r.get("jinjer_出勤時刻"), r.get("jinjer_退勤時刻")),
        axis=1,
    )
    sheet_renamed["突合区分"] = sheet_renamed.apply(
        lambda r: _work_segment_key(r.get("勤務表_出勤時刻"), r.get("勤務表_退勤時刻")),
        axis=1,
    )

    # outer join（氏名_normalized + 日付 + 突合区分）— 対象は勤務表社員のみ
    merged = pd.merge(
        jinjer_renamed,
        sheet_renamed,
        on=["氏名_normalized", "日付", "突合区分"],
        how="outer",
        suffixes=("", "_sheet")
    )

    # 氏名の決定（どちらかから取得）
    merged["氏名"] = merged.apply(
        lambda r: r.get("jinjer_氏名") if pd.notna(r.get("jinjer_氏名")) and str(r.get("jinjer_氏名")) not in ["nan", ""]
        else r.get("勤務表_氏名", "不明"),
        axis=1
    )

    # 差分計算
    def calc_diff(row, col_sheet, col_jinjer):
        d = time_diff_minutes(row.get(col_sheet), row.get(col_jinjer))
        return abs(d) if d is not None else None

    merged["出勤差分(分)"] = merged.apply(lambda r: calc_diff(r, "勤務表_出勤時刻", "jinjer_出勤時刻"), axis=1)
    merged["退勤差分(分)"] = merged.apply(lambda r: calc_diff(r, "勤務表_退勤時刻", "jinjer_退勤時刻"), axis=1)

    # 判定
    results = merged.apply(lambda r: judge(r, threshold_minutes), axis=1)
    merged["判定"] = [r[0] for r in results]
    merged["詳細"] = [r[1] for r in results]

    result_columns = [
        "氏名",
        "日付",
        "勤務表_出勤時刻",
        "jinjer_出勤時刻",
        "出勤差分(分)",
        "勤務表_退勤時刻",
        "jinjer_退勤時刻",
        "退勤差分(分)",
        "判定",
        "詳細",
        "jinjer_コメント",
        "勤務表_コメント",
    ]

    # 存在しないカラムを補完
    for col in result_columns:
        if col not in merged.columns:
            merged[col] = None

    result = merged[result_columns].sort_values(["氏名", "日付"]).reset_index(drop=True)
    return result, unsubmitted_names
