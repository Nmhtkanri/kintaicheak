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
    # ローマ字氏名の表記ゆれ（MAHARJAN / Maharjan 等）を吸収するため大文字小文字を畳む。
    # 漢字・かなには影響しない。
    name = name.casefold()
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


def _levenshtein(a, b):
    """2文字列の編集距離（挿入/削除/置換）を返す。"""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cost = 0 if ca == cb else 1
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost))
        prev = cur
    return prev[-1]


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


def elapsed_minutes(start, end):
    """出勤から退勤までの経過分を返す。日跨ぎは翌日退勤として扱う。"""
    start_min = _to_minutes(start)
    end_min = _to_minutes(end)
    if start_min is None or end_min is None:
        return None
    if end_min < start_min:
        end_min += 24 * 60
    return end_min - start_min


def format_duration(minutes):
    if minutes is None or pd.isna(minutes):
        return ""
    h, m = divmod(int(minutes), 60)
    return f"{h}:{m:02d}"


def _is_midnight(value):
    return isinstance(value, time) and value.hour == 0 and value.minute == 0


# 「特記」列の値。quick_compare（差異一覧）が部分一致で解釈して行の出し方を変える。
# jinjer は夜勤を「勤務開始日の1行・退勤24時超表記(例 33:30)」で記録するため、
# 書き戻しも必ず開始日の行に24時超表記で行う——というのがこの一連の特記の前提。
TOKKI_MONTH_END_OVERNIGHT = "月末日跨ぎ勤務(請求側の翌月分未取得のため退勤・総労働は翌月確認)"
TOKKI_PREV_MONTH_TAIL_MATCHED = "前月末夜勤の後半(前月末日で突合)"
TOKKI_PREV_MONTH_TAIL_UNMATCHED = (
    "前月末夜勤の後半(jinjer側は前月末日行に記録・自動書戻し不可。"
    "jinjerダウンロード範囲に前月末日を含めると突合可能)"
)
TOKKI_JINJER_SPLIT = (
    "jinjer側2行分割登録(退勤の自動書戻し不可。"
    "修正する場合はjinjerを開始日1行の24時超表記へ手動統合)"
)


def _as_date(value):
    """日付値を datetime.date に正規化する（不明な型は None）。"""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def _dedup_timesheet_rows(timesheet_df):
    """請求勤怠側の完全重複行を除去する。

    Fieldglassのレポートは同一の勤務明細を複数回出力することがあり、そのままだと
    差異一覧に同じ差異が重複して並ぶ。氏名・日付・出退勤・実働・コメントがすべて
    同じ行だけを1行に畳む（中抜け等の正当な複数勤務は時刻が異なるため残る）。
    """
    if timesheet_df.empty:
        return timesheet_df
    subset = [
        c for c in ("氏名_normalized", "日付", "出勤時刻", "退勤時刻", "総労働時間(分)", "コメント")
        if c in timesheet_df.columns
    ]
    return timesheet_df.drop_duplicates(subset=subset, keep="first").reset_index(drop=True)


def _merge_midnight_split_rows(timesheet_df):
    """「深夜0時割り」の2行（D日 X〜24:00 ＋ D+1日 0:00〜Y）を勤務開始日の1行に結合する。

    SAP Fieldglass等は日跨ぎ夜勤を暦日で2行に分割する。またjinjer側にも、深夜0時で
    2行に分けて登録された夜勤が存在する（例: 6/2 17:00〜24:00 ＋ 6/3 0:00〜7:00）。
    突合は「勤務開始日の1行（退勤は24時超相当）」の形に両側を揃えて行うため、
    日付=D・出勤=X・退勤=Y（翌日時刻）・実働=両行の合計 に結合する。
    曖昧なケース（同日に複数候補）は結合しない。

    Returns:
        tuple(df, merged_keys)
        - merged_keys: 結合が起きた (氏名_normalized, 開始日) の集合
    """
    if timesheet_df.empty:
        return timesheet_df, set()

    df = timesheet_df.reset_index(drop=True)
    first_halves = {}   # (氏名key, 日付) -> [行番号] 退勤がちょうど0:00（前半）
    second_halves = {}  # (氏名key, 日付) -> [行番号] 出勤がちょうど0:00（後半）
    for idx, row in df.iterrows():
        key = row.get("氏名_normalized")
        d = _as_date(row.get("日付"))
        if not key or d is None:
            continue
        start = row.get("出勤時刻")
        end = row.get("退勤時刻")
        if isinstance(start, time) and not _is_midnight(start) and _is_midnight(end):
            first_halves.setdefault((key, d), []).append(idx)
        if isinstance(end, time) and not _is_midnight(end) and _is_midnight(start):
            second_halves.setdefault((key, d), []).append(idx)

    dropped = set()
    merged_keys = set()
    has_total = "総労働時間(分)" in df.columns
    for (key, d), f_idxs in first_halves.items():
        s_idxs = second_halves.get((key, d + timedelta(days=1)), [])
        if len(f_idxs) != 1 or len(s_idxs) != 1:
            continue
        fi, si = f_idxs[0], s_idxs[0]
        if fi in dropped or si in dropped:
            continue
        merged_keys.add((key, d))
        df.at[fi, "退勤時刻"] = df.at[si, "退勤時刻"]
        if has_total:
            t1, t2 = df.at[fi, "総労働時間(分)"], df.at[si, "総労働時間(分)"]
            try:
                total = int(t1) + int(t2) if pd.notna(t1) and pd.notna(t2) else None
            except (TypeError, ValueError):
                total = None
            df.at[fi, "総労働時間(分)"] = total
        comments = [
            str(c).strip()
            for c in (df.at[fi, "コメント"], df.at[si, "コメント"])
            if pd.notna(c) and str(c).strip() and str(c).strip().lower() not in ("nan", "none")
        ]
        df.at[fi, "コメント"] = " / ".join(dict.fromkeys(comments)) or None
        dropped.add(si)

    if not dropped:
        return df, merged_keys
    return df.drop(index=list(dropped)).reset_index(drop=True), merged_keys


def _pair_cross_midnight_rows(merged, sheet_min_date, sheet_max_date):
    """突合区分の食い違いで別行になった夜勤を再ペアリングする。

    対象は2パターン:
      1. 同日で 請求側「X〜24:00」（打切り）× jinjer側「夜勤で翌朝まで」。
         1行に結合し、日付が請求勤怠の最終日なら「月末日跨ぎ」の特記を付ける
         （退勤・総労働の差分は呼び出し側で落とし、翌月確認に回す）。
      2. 月初の「0:00〜Y」の孤児行（前月末に始まった勤務の後半）。jinjer側は前月末日の
         行に記録されるため、前月末日のjinjerのみ行があれば日付=前月末日で結合し
         （書き戻しが開始日の行へ24時超表記で行える形）、無ければ「自動書戻し不可」の
         特記を付ける。
    """
    if merged.empty:
        return merged

    def _has(v):
        return pd.notna(v) and str(v).strip() not in ("", "nan", "None")

    df = merged.reset_index(drop=True)
    jonly = {}  # (氏名key, 日付) -> [行番号] jinjerのみ行
    sonly = {}  # (氏名key, 日付) -> [行番号] 勤務表のみ行
    for idx, row in df.iterrows():
        key = row.get("氏名_normalized")
        d = _as_date(row.get("日付"))
        if not key or d is None:
            continue
        has_j = _has(row.get("jinjer_氏名"))
        has_s = _has(row.get("勤務表_氏名"))
        if has_j and not has_s:
            jonly.setdefault((key, d), []).append(idx)
        elif has_s and not has_j:
            sonly.setdefault((key, d), []).append(idx)

    jinjer_side_cols = [c for c in df.columns if str(c).startswith("jinjer_")]
    if "データソース" in df.columns:
        jinjer_side_cols.append("データソース")

    dropped = set()

    # パターン1: 同日の 勤務表(X〜24:00) × jinjer(夜勤・翌朝まで) を結合
    for (key, d), s_idxs in sonly.items():
        j_idxs = [ji for ji in jonly.get((key, d), []) if ji not in dropped]
        if len(s_idxs) != 1 or len(j_idxs) != 1:
            continue
        si, ji = s_idxs[0], j_idxs[0]
        if si in dropped:
            continue
        s_start = df.at[si, "勤務表_出勤時刻"]
        s_end = df.at[si, "勤務表_退勤時刻"]
        j_start = df.at[ji, "jinjer_出勤時刻"]
        j_end = df.at[ji, "jinjer_退勤時刻"]
        if not (isinstance(s_start, time) and not _is_midnight(s_start) and _is_midnight(s_end)):
            continue
        if not (
            isinstance(j_start, time) and isinstance(j_end, time)
            and _to_minutes(j_end) < _to_minutes(j_start)
        ):
            continue
        for col in jinjer_side_cols:
            df.at[si, col] = df.at[ji, col]
        if sheet_max_date is not None and d == sheet_max_date and not _is_midnight(j_end):
            df.at[si, "特記"] = TOKKI_MONTH_END_OVERNIGHT
        dropped.add(ji)

    # 月末の打切り行（jinjer側に打刻がなくペアにならなかったもの）にも特記を付ける。
    # 退勤24:00は勤務がそこで終わった保証がないため、自動書き戻しの対象にさせない。
    for (key, d), s_idxs in sonly.items():
        if sheet_max_date is None or d != sheet_max_date:
            continue
        for si in s_idxs:
            if si in dropped or _has(df.at[si, "jinjer_氏名"]):
                continue
            s_start = df.at[si, "勤務表_出勤時刻"]
            s_end = df.at[si, "勤務表_退勤時刻"]
            if isinstance(s_start, time) and not _is_midnight(s_start) and _is_midnight(s_end):
                df.at[si, "特記"] = TOKKI_MONTH_END_OVERNIGHT

    # パターン2: 月初の「0:00〜Y」の孤児行を前月末日のjinjer行へ付け替える
    for (key, d), s_idxs in sonly.items():
        if sheet_min_date is None or d != sheet_min_date:
            continue
        for si in s_idxs:
            if si in dropped or _has(df.at[si, "jinjer_氏名"]):
                continue
            s_start = df.at[si, "勤務表_出勤時刻"]
            s_end = df.at[si, "勤務表_退勤時刻"]
            if not (_is_midnight(s_start) and isinstance(s_end, time) and not _is_midnight(s_end)):
                continue
            picked = None
            cand = [ji for ji in jonly.get((key, d - timedelta(days=1)), []) if ji not in dropped]
            if len(cand) == 1:
                j_start = df.at[cand[0], "jinjer_出勤時刻"]
                j_end = df.at[cand[0], "jinjer_退勤時刻"]
                if (
                    isinstance(j_start, time) and isinstance(j_end, time)
                    and _to_minutes(j_end) < _to_minutes(j_start)
                ):
                    picked = cand[0]
            if picked is not None:
                ji = picked
                # 前半（出勤〜24:00）は前月分のチェックで確認済みのため、出勤は差0で埋める。
                df.at[ji, "勤務表_氏名"] = df.at[si, "勤務表_氏名"]
                df.at[ji, "勤務表_出勤時刻"] = df.at[ji, "jinjer_出勤時刻"]
                df.at[ji, "勤務表_退勤時刻"] = s_end
                if "勤務表_実働(分)" in df.columns:
                    df.at[ji, "勤務表_実働(分)"] = None  # 後半のみの実働のため総労働は突合しない
                if "勤務表_コメント" in df.columns:
                    df.at[ji, "勤務表_コメント"] = df.at[si, "勤務表_コメント"]
                if "データソース_sheet" in df.columns:
                    df.at[ji, "データソース_sheet"] = df.at[si, "データソース_sheet"]
                df.at[ji, "特記"] = TOKKI_PREV_MONTH_TAIL_MATCHED
                dropped.add(si)
            else:
                df.at[si, "特記"] = TOKKI_PREV_MONTH_TAIL_UNMATCHED

    if not dropped:
        return df
    return df.drop(index=list(dropped)).reset_index(drop=True)


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
    # 月末日跨ぎ（請求側が24:00打切り・翌月分未取得）は退勤・総労働を判定に使わない。
    month_end_truncated = "月末日跨ぎ" in str(row.get("特記") or "")
    start_diff = time_diff_minutes(sheet_start, jinjer_start)
    end_diff = None if month_end_truncated else time_diff_minutes(sheet_end, jinjer_end)

    total_diff = None if month_end_truncated else row.get("総労働差分(分)")
    start_diff_abs = abs(start_diff) if start_diff is not None else 0
    end_diff_abs = abs(end_diff) if end_diff is not None else 0
    total_diff_abs = abs(total_diff) if total_diff is not None and pd.notna(total_diff) else 0
    max_diff = max(start_diff_abs, end_diff_abs, total_diff_abs)

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
    # 請求勤怠ファイル記載の正味労働時間（分）。パーサーが付けない経路（シフト記号等）でも
    # 後段で参照できるよう、無ければ None 列を補完しておく。
    if "総労働時間(分)" not in timesheet_df.columns:
        timesheet_df["総労働時間(分)"] = None
    # 「特記」列（月末跨ぎ・Fieldglass時刻なし等の注記）も同様に補完する。
    if "特記" not in timesheet_df.columns:
        timesheet_df["特記"] = ""
    timesheet_df["特記"] = timesheet_df["特記"].fillna("")
    jinjer_df["氏名_normalized"] = jinjer_df["氏名"].apply(normalize_name)
    timesheet_df["氏名_normalized"] = timesheet_df["氏名"].apply(normalize_name)
    jinjer_name_keys = set(jinjer_df["氏名_normalized"].unique())
    jinjer_key_to_name = {
        key: name
        for key, name in (
            (row.get("氏名_normalized"), row.get("氏名"))
            for _, row in jinjer_df.iterrows()
        )
        if key and pd.notna(name) and str(name).strip()
    }

    def choose_unique_partial_key(key):
        if not key or len(key) < 2 or key in jinjer_name_keys:
            return None
        candidates = [
            jinjer_key
            for jinjer_key in jinjer_name_keys
            if jinjer_key and (jinjer_key.startswith(key) or key.startswith(jinjer_key))
        ]
        return candidates[0] if len(candidates) == 1 else None

    def choose_fuzzy_key(key):
        """OCRの1文字誤読を吸収する近似一致。

        スキャン勤務表のOCRは氏名の漢字を1文字だけ読み違えることがある
        （例: 矢野瑞穂 → 矢野瑞也）。完全一致・部分一致のどちらも外れたときに限り、
        jinjer側に「編集距離1で他に候補がない」氏名が1つだけある場合だけ採用する。
        誤マッチ防止のため、両者とも4文字以上のフルネームに限定する
        （田中/田口のような短い姓だけの誤一致を避ける）。
        """
        if not key or len(key) < 4 or key in jinjer_name_keys:
            return None
        candidates = [
            jinjer_key
            for jinjer_key in jinjer_name_keys
            if jinjer_key and len(jinjer_key) >= 4 and _levenshtein(key, jinjer_key) <= 1
        ]
        return candidates[0] if len(candidates) == 1 else None

    def choose_sheet_key(name):
        for key in normalize_name_keys(name):
            if key in jinjer_name_keys:
                return key
            partial_key = choose_unique_partial_key(key)
            if partial_key:
                return partial_key
        # 完全一致・部分一致が無いときだけ、OCR誤読を想定した近似一致を試す
        for key in normalize_name_keys(name):
            fuzzy_key = choose_fuzzy_key(key)
            if fuzzy_key:
                return fuzzy_key
        keys = normalize_name_keys(name)
        return keys[0] if keys else ""

    timesheet_df["氏名_normalized"] = timesheet_df["氏名"].apply(choose_sheet_key)
    timesheet_df["氏名"] = timesheet_df.apply(
        lambda r: jinjer_key_to_name.get(r.get("氏名_normalized"), r.get("氏名")),
        axis=1,
    )

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

    # 請求勤怠側の完全重複行を除去し、深夜0時割りの2行を勤務開始日1本に結合する
    # （jinjerは夜勤を開始日の1行・24時超表記で記録するため、同じ形に揃えて突合する）。
    timesheet_df = _dedup_timesheet_rows(timesheet_df)
    timesheet_df, _ = _merge_midnight_split_rows(timesheet_df)

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

    # jinjer側にも深夜0時割りで2行登録された夜勤が存在する（例: 17:00〜24:00 ＋ 翌日0:00〜7:00）。
    # 突合用に開始日1本へ結合する。該当した勤務は「特記」を付け、退勤の自動書き戻しを止める
    # （実データのjinjerは2行のままなので、開始日行だけに24時超の退勤を書くと二重計上になる）。
    jinjer_filtered, jinjer_merged_keys = _merge_midnight_split_rows(jinjer_filtered)

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
        "総労働時間(分)": "勤務表_実働(分)",
    })

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

    # 突合区分の食い違いで別行になった夜勤（月末打切り・前月末後半の孤児行）を再ペアリング
    sheet_dates = [d for d in (_as_date(v) for v in sheet_renamed["日付"]) if d is not None]
    sheet_min_date = min(sheet_dates) if sheet_dates else None
    sheet_max_date = max(sheet_dates) if sheet_dates else None
    merged = _pair_cross_midnight_rows(merged, sheet_min_date, sheet_max_date)
    if "特記" not in merged.columns:
        merged["特記"] = ""
    merged["特記"] = merged["特記"].fillna("")

    # jinjer側を結合した勤務に特記を付ける（退勤の自動書き戻し抑止用）
    if jinjer_merged_keys:
        def _mark_jinjer_split(row):
            tokki = str(row.get("特記") or "").strip()
            if (row.get("氏名_normalized"), _as_date(row.get("日付"))) in jinjer_merged_keys:
                return f"{tokki} / {TOKKI_JINJER_SPLIT}" if tokki else TOKKI_JINJER_SPLIT
            return tokki
        merged["特記"] = merged.apply(_mark_jinjer_split, axis=1)

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

    merged["勤務表_総労働時間(分)"] = merged.apply(
        lambda r: elapsed_minutes(r.get("勤務表_出勤時刻"), r.get("勤務表_退勤時刻")),
        axis=1,
    )
    merged["jinjer_総労働時間(分)"] = merged.apply(
        lambda r: elapsed_minutes(r.get("jinjer_出勤時刻"), r.get("jinjer_退勤時刻")),
        axis=1,
    )
    merged["総労働差分(分)"] = merged.apply(
        lambda r: (
            abs(r.get("勤務表_総労働時間(分)") - r.get("jinjer_総労働時間(分)"))
            if pd.notna(r.get("勤務表_総労働時間(分)")) and pd.notna(r.get("jinjer_総労働時間(分)"))
            else None
        ),
        axis=1,
    )
    merged["勤務表_総労働時間"] = merged["勤務表_総労働時間(分)"].apply(format_duration)
    merged["jinjer_総労働時間"] = merged["jinjer_総労働時間(分)"].apply(format_duration)

    # 請求勤怠ファイルに日別の正味労働時間（実働）が記載されていればそれを保持する。
    # 手順1の「総労働時間」表示は拘束時間(退勤−出勤)のまま据え置き（jinjer側も拘束のため整合）、
    # 差異一覧(quick_compare)はこの正味列を優先して正味同士で突合する。
    if "勤務表_実働(分)" not in merged.columns:
        merged["勤務表_実働(分)"] = None
    merged["勤務表_実働時間"] = merged["勤務表_実働(分)"].apply(format_duration)

    # 月末日跨ぎ（請求側が24:00で打切り）の行は退勤・総労働を突合しない（翌月確認）。
    # 請求の24:00は暦日割りの都合であって実際の退勤ではないため、差分として扱わない。
    month_end_mask = merged["特記"].astype(str).str.contains("月末日跨ぎ", na=False)
    if month_end_mask.any():
        merged.loc[month_end_mask, "退勤差分(分)"] = None
        merged.loc[month_end_mask, "総労働差分(分)"] = None

    # 判定
    results = merged.apply(lambda r: judge(r, threshold_minutes), axis=1)
    merged["判定"] = [r[0] for r in results]
    merged["詳細"] = [r[1] for r in results]

    # 特記がある行は詳細に注記を足し、月末跨ぎ等の確認が必要なものはOKでも要確認へ引き上げる。
    def _apply_tokki(row):
        tokki = str(row.get("特記") or "").strip()
        if not tokki or tokki.lower() == "nan":
            return row["判定"], row["詳細"]
        detail = f"{row['詳細']} / {tokki}" if row["詳細"] else tokki
        judgement = row["判定"]
        # 前月末突合済み・jinjer2行分割（差異なしなら問題なし）はOKのまま据え置く
        keep_ok = TOKKI_PREV_MONTH_TAIL_MATCHED in tokki or (
            TOKKI_JINJER_SPLIT in tokki and tokki.strip(" /") == TOKKI_JINJER_SPLIT
        )
        if judgement == "OK" and not keep_ok:
            judgement = "要確認"
        return judgement, detail

    tokki_results = merged.apply(_apply_tokki, axis=1)
    merged["判定"] = [r[0] for r in tokki_results]
    merged["詳細"] = [r[1] for r in tokki_results]

    result_columns = [
        "氏名",
        "日付",
        "勤務表_出勤時刻",
        "jinjer_出勤時刻",
        "出勤差分(分)",
        "勤務表_退勤時刻",
        "jinjer_退勤時刻",
        "退勤差分(分)",
        "勤務表_総労働時間",
        "jinjer_総労働時間",
        "総労働差分(分)",
        "勤務表_実働時間",
        "判定",
        "詳細",
        "jinjer_コメント",
        "勤務表_コメント",
        "特記",
    ]

    # 存在しないカラムを補完
    for col in result_columns:
        if col not in merged.columns:
            merged[col] = None

    result = merged[result_columns].sort_values(["氏名", "日付"]).reset_index(drop=True)
    return result, unsubmitted_names
