"""KDX（ユニアデックス作成）勤務シフト表 PDF の構造化パーサ

「(エヌエム・ヒューマテック様)YYYY年M月分勤務シフト表.pdf」を Claude を経由せず
確定的に解析する。この表のセル記号は勤務時間ではなく**休憩時間帯コード**である点が特殊:

  A1〜A6 … 日勤（9:00〜17:30）の休憩1h の時間帯（例 A5=13:30〜14:30）
  C1〜C6 … 夜勤（16:30〜34:00）の休憩2枠・計2h の時間帯（例 C4=23:00〜24:00/5:00〜6:00）
  ／     … 公休（希望休は背景色違いのみで記号は同じ）
  ー     … 夜勤明け（前日16:30開始の勤務が当日10:00まで続く日。開始日1行ルールにより
            スケジュール行は作らない＝「休み」扱い）
  有     … 有給（全日）

勤務時間は表に書かれていないため固定値（KDX_DAY_* / KDX_NIGHT_*）を凡例に補い、
既存の雛形時刻マッチ（find_matching_template）で jinjer 雛形へ自動解決させる。
休憩の時間帯はコードごとに異なるが、jinjer 雛形は「出退勤時刻が一致し休憩は合計が
同じであればよい」運用（2026-07-22 谷津確定）のため、雛形の代表値に吸収させる。
休憩コードの内訳は label に残して確認画面で見えるようにする。
"""

from __future__ import annotations

import calendar
import logging
import re
import unicodedata
from datetime import date

logger = logging.getLogger(__name__)

# シフト表の系統識別子。氏名エイリアス表の適用範囲を KDX に限るために使う。
KDX_SOURCE = "kdx"

# 勤務時間の固定値（表に記載がないため。変更時はここを直す）
KDX_DAY_START = "9:00"
KDX_DAY_END = "17:30"
KDX_DAY_BREAK_MIN = 60
KDX_NIGHT_START = "16:30"
KDX_NIGHT_END = "34:00"      # 24時超表記（翌10:00）。雛形マッチは生文字列のまま行われる
KDX_NIGHT_BREAK_MIN = 120

# sniff / タイトル
_TITLE_RE = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月度?\s*勤務スケジュール表")
_SNIFF_KEYWORDS = ("勤務スケジュール表", "休憩取得予定時間")

# 凡例「A1：11:30～12:30」「C1：20:00～21:00／2:00～3:00」
_LEGEND_RE = re.compile(
    r"([AC][1-6])：(\d{1,2}:\d{2})～(\d{1,2}:\d{2})(?:／(\d{1,2}:\d{2})～(\d{1,2}:\d{2}))?"
)

# 明け・公休のバリアント（PDF実データは「ー」=U+30FC、「／」=U+FF0F）
_AKE_VARIANTS = {"ー", "-", "−", "—", "―"}
_OFF_VARIANTS = {"／", "/"}
_YUKYU_CODE = "有"
_AKE_CODE = "ー"
_OFF_CODE = "／"

_WEEKDAY_KANJI = ["月", "火", "水", "木", "金", "土", "日"]


def _norm_code(text: str) -> str | None:
    """word テキスト → シフト記号へ正規化。記号でなければ None。

    A/C コードは全角英数の月にも耐えるよう NFKC で半角へ寄せる。
    ー／ のバリアント（半角ハイフン等）は代表字へ寄せる。
    """
    s = str(text or "").strip()
    if not s:
        return None
    if s in _AKE_VARIANTS:
        return _AKE_CODE
    if s in _OFF_VARIANTS:
        return _OFF_CODE
    if s == _YUKYU_CODE:
        return s
    n = unicodedata.normalize("NFKC", s)
    if re.fullmatch(r"[AC][1-6]", n):
        return n
    return None


def _group_words_into_rows(words: list[dict], tol: float = 2.0) -> list[dict]:
    """word を top 座標でグルーピングして行にする（tol pt 以内は同一行）。"""
    rows: list[dict] = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        for row in rows:
            if abs(row["top"] - w["top"]) < tol:
                row["words"].append(w)
                break
        else:
            rows.append({"top": w["top"], "words": [w]})
    return rows


def _find_date_columns(rows: list[dict]) -> tuple[dict[int, float], float] | None:
    """「日 付 1 2 ... 31」行を探し、{day: x中心} と行topを返す。

    数字 word が 20 個以上並ぶ行を日付ヘッダーとみなす（1〜31 の連番確認つき）。
    """
    for row in rows:
        day_x: dict[int, float] = {}
        for w in row["words"]:
            t = str(w["text"]).strip()
            if t.isdigit():
                d = int(t)
                if 1 <= d <= 31 and d not in day_x:
                    day_x[d] = (w["x0"] + w["x1"]) / 2
        if len(day_x) >= 20:
            days = sorted(day_x)
            if days == list(range(days[0], days[-1] + 1)) and days[0] == 1:
                return day_x, row["top"]
    return None


def _check_weekday_row(
    rows: list[dict],
    day_x: dict[int, float],
    date_row_top: float,
    year: int,
    month: int,
    snap_tol: float,
) -> tuple[int, int]:
    """日付行の直下にある曜日行をカレンダーと突合する。

    Returns:
        (matched, total)  曜日行が見つからなければ (0, 0)
    """
    candidates = [
        r for r in rows
        if 0 < r["top"] - date_row_top < 20
        and sum(1 for w in r["words"] if str(w["text"]).strip() in _WEEKDAY_KANJI) >= 10
    ]
    if not candidates:
        return 0, 0
    row = candidates[0]
    days_in_month = calendar.monthrange(year, month)[1]
    matched = total = 0
    for w in row["words"]:
        t = str(w["text"]).strip()
        if t not in _WEEKDAY_KANJI:
            continue
        x = (w["x0"] + w["x1"]) / 2
        day = _snap_to_day(x, day_x, snap_tol)
        if day is None or day > days_in_month:
            continue
        total += 1
        if _WEEKDAY_KANJI[date(year, month, day).weekday()] == t:
            matched += 1
    return matched, total


def _snap_to_day(x_center: float, day_x: dict[int, float], tol: float) -> int | None:
    """x 中心座標を最近傍の日付列へスナップする（tol を超えたら None）。"""
    best_day, best_dist = None, None
    for d, dx in day_x.items():
        dist = abs(x_center - dx)
        if best_dist is None or dist < best_dist:
            best_day, best_dist = d, dist
    if best_day is None or best_dist is None or best_dist > tol:
        return None
    return best_day


def _extract_legend_breaks(text: str) -> dict[str, str]:
    """ページ全文から休憩凡例を抽出する → {code: "11:30～12:30" / "20:00～21:00/2:00～3:00"}"""
    result: dict[str, str] = {}
    for m in _LEGEND_RE.finditer(text or ""):
        code = m.group(1)
        desc = f"{m.group(2)}～{m.group(3)}"
        if m.group(4):
            desc += f"/{m.group(4)}～{m.group(5)}"
        result.setdefault(code, desc)
    return result


def build_kdx_legend(break_desc: dict[str, str], seen_codes: set[str]) -> list[dict]:
    """固定勤務時間 ＋ PDF凡例の休憩内訳から legend リストを組み立てる。

    凡例に載っている全コード＋表に実際に出現したコードを対象にする
    （どちらか片方にしか無くても落とさない）。
    """
    codes = sorted(
        {c for c in break_desc} | {c for c in seen_codes if re.fullmatch(r"[AC][1-6]", c)}
    )
    legend: list[dict] = []
    for code in codes:
        desc = break_desc.get(code, "")
        if code.startswith("A"):
            label = f"KDX日勤(休憩{desc})" if desc else "KDX日勤"
            legend.append({
                "code": code, "label": label,
                "start_time": KDX_DAY_START, "end_time": KDX_DAY_END,
                "break_minutes": KDX_DAY_BREAK_MIN, "is_off": False,
            })
        else:
            label = f"KDX夜勤(休憩{desc})" if desc else "KDX夜勤"
            legend.append({
                "code": code, "label": label,
                "start_time": KDX_NIGHT_START, "end_time": KDX_NIGHT_END,
                "break_minutes": KDX_NIGHT_BREAK_MIN, "is_off": False,
            })
    # 「明」を含む label により exporter の明け判定（→"休み"）が効く
    legend.append({"code": _AKE_CODE, "label": "夜勤明け", "start_time": None,
                   "end_time": None, "break_minutes": 0, "is_off": True})
    legend.append({"code": _OFF_CODE, "label": "公休", "start_time": None,
                   "end_time": None, "break_minutes": 0, "is_off": True})
    if _YUKYU_CODE in seen_codes:
        # 全日有休は exporter 側で「一般」雛形になる（_is_full_day_paid_leave が
        # is_off 判定より先に効く）。is_off=True にして雛形時刻マッチの対象から外す
        # （False だと suggest_template_id のゴミ ID "X" が code_to_tpl に入る）。
        legend.append({"code": _YUKYU_CODE, "label": "有給休暇", "start_time": None,
                       "end_time": None, "break_minutes": 0, "is_off": True})
    return legend


def parse_kdx_words(
    words: list[dict],
    page_text: str,
    *,
    filename: str,
    target_year: "int | None",
    target_month: "int | None",
) -> dict:
    """word リスト（pdfplumber extract_words 互換）→ code_sheet 形式へ解析する純関数。

    target_year/month が None（画面の対象年月が未入力）の場合は、
    PDFタイトルの年月をそのまま採用する（曜日整合チェックは常に行う）。

    Raises:
        ValueError: タイトル年月が対象と不一致 / 日付ヘッダーが見つからない 等
    """
    rows = _group_words_into_rows(words)

    # --- タイトルから年月（対象年月の指定があり不一致なら誤投入防止のため中止）---
    m = _TITLE_RE.search(page_text or "")
    if not m:
        raise ValueError(f"{filename}: タイトルの年月（YYYY年M月度勤務スケジュール表）が見つかりません")
    year, month = int(m.group(1)), int(m.group(2))
    if target_year is not None and target_month is not None and (
            year != target_year or month != target_month):
        raise ValueError(
            f"{filename}: シフト表の年月 {year}年{month}月 が対象 {target_year}年{target_month}月 と一致しません")

    # --- 日付ヘッダー列 ---
    found = _find_date_columns(rows)
    if not found:
        raise ValueError(f"{filename}: 日付ヘッダー行（1〜31）が見つかりません")
    day_x, date_row_top = found
    days_in_month = calendar.monthrange(year, month)[1]

    xs = [day_x[d] for d in sorted(day_x)]
    col_pitch = (xs[-1] - xs[0]) / (len(xs) - 1) if len(xs) > 1 else 18.0
    snap_tol = col_pitch * 0.6
    left_edge = xs[0] - col_pitch      # 日付1列より左＝氏名エリア
    right_edge = xs[-1] + col_pitch    # 日付31列より右＝「日勤/夜勤シフト」ラベル

    # --- 曜日行の整合チェック（年月違いの誤読み安全弁）---
    wd_matched, wd_total = _check_weekday_row(rows, day_x, date_row_top, year, month, snap_tol)
    if wd_total >= 10 and wd_matched < wd_total - 2:
        raise ValueError(
            f"{filename}: 曜日行がカレンダーと一致しません（{wd_matched}/{wd_total}）")

    # --- 記号行の抽出 ---
    employees: list[dict] = []
    seen_codes: set[str] = set()
    for row in rows:
        if row["top"] <= date_row_top:
            continue
        day_codes: dict[int, str] = {}
        for w in row["words"]:
            x = (w["x0"] + w["x1"]) / 2
            if not (left_edge <= x <= right_edge):
                continue  # 右端の「日勤/シフト」等のラベルを除外
            code = _norm_code(w["text"])
            if code is None:
                continue
            day = _snap_to_day(x, day_x, snap_tol)
            if day is None or day > days_in_month:
                continue
            if day in day_codes:
                logger.warning("%s: top=%.1f 日%d に複数記号（%s/%s）→先勝ち",
                               filename, row["top"], day, day_codes[day], code)
                continue
            day_codes[day] = code
        if len(day_codes) < 5:
            continue  # 記号行ではない（凡例・祝日ラベル等）

        # 氏名: 氏名エリア（日付1列より左）にあり、この記号行に最も近い行の word
        name = _find_name_for_row(rows, row["top"], left_edge, date_row_top)
        if not name:
            logger.warning("%s: top=%.1f の記号行に対応する氏名が見つかりません（スキップ）",
                           filename, row["top"])
            continue

        shifts = []
        for d in range(1, days_in_month + 1):
            code = day_codes.get(d, "")
            if code:
                seen_codes.add(code)
            shifts.append({"date": date(year, month, d).isoformat(), "code": code})
        employees.append({"name": name, "shifts": shifts})

    if not employees:
        raise ValueError(f"{filename}: 従業員のシフト行を1件も抽出できませんでした")

    # --- 凡例（休憩内訳）---
    break_desc = _extract_legend_breaks(page_text)
    legend = build_kdx_legend(break_desc, seen_codes)

    return {
        "filename": filename,
        "year": year,
        "month": month,
        "legend": legend,
        "employees": employees,
        "off_markers": [_OFF_CODE, _AKE_CODE],
        # シフト表の系統。氏名エイリアス表（services/employee_alias.py）を
        # この系統に限って適用するために使う。
        "source": KDX_SOURCE,
        "section_info": {
            "section_index": None,
            "weekday_matched": wd_matched,
            "weekday_total": wd_total,
        },
    }


def _find_name_for_row(
    rows: list[dict],
    symbol_top: float,
    left_edge: float,
    date_row_top: float,
    max_dist: float = 14.0,
) -> str:
    """記号行 top に最も近い行の氏名エリア word を氏名として返す。

    KDX表は氏名セル（縦中央）と記号（下寄り）のベースラインが数pt ずれるため、
    ±max_dist pt の範囲で最近傍の行から拾う。同一行に複数 word あれば x0 順で連結。
    """
    best_name, best_dist = "", None
    for row in rows:
        if row["top"] <= date_row_top:
            continue
        dist = abs(row["top"] - symbol_top)
        if dist > max_dist:
            continue
        name_words = [w for w in sorted(row["words"], key=lambda w: w["x0"])
                      if w["x1"] <= left_edge and _norm_code(w["text"]) is None]
        if not name_words:
            continue
        name = "".join(str(w["text"]).strip() for w in name_words)
        if not name or name in ("日付", "日", "付", "曜日"):
            continue
        if best_dist is None or dist < best_dist:
            best_name, best_dist = name, dist
    return best_name


def is_kdx_shift_pdf(filepath: str) -> bool:
    """KDX（ユニアデックス）勤務シフト表 PDF かを軽量判定する。"""
    if not str(filepath).lower().endswith(".pdf"):
        return False
    try:
        import pdfplumber
        with pdfplumber.open(filepath) as pdf:
            if not pdf.pages:
                return False
            text = pdf.pages[0].extract_text() or ""
    except Exception as e:
        logger.warning("KDX sniff 失敗 %s: %s", filepath, e)
        return False
    return all(k in text for k in _SNIFF_KEYWORDS) and bool(_TITLE_RE.search(text))


def parse_kdx_shift_pdf(
    filepath: str,
    target_year: "int | None" = None,
    target_month: "int | None" = None,
) -> dict:
    """KDX勤務シフト表 PDF → code_sheet 形式（parse_structured_files と同形）。

    target_year/month 未指定（None）の場合はPDFタイトルの年月を採用する。
    """
    import os
    import pdfplumber

    with pdfplumber.open(filepath) as pdf:
        page = pdf.pages[0]
        words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
        page_text = page.extract_text() or ""

    result = parse_kdx_words(
        words, page_text,
        filename=os.path.basename(filepath),
        target_year=target_year, target_month=target_month,
    )
    logger.info(
        "KDXシフト表を構造化解析: %s → %d年%d月 従業員%d人 凡例%d個",
        result["filename"], result["year"], result["month"],
        len(result["employees"]), len(result["legend"]),
    )
    return result


# =============================================================================
# AI読み取り経路への凡例の強制適用（2026-09 の文字なしPDF対策）
# =============================================================================

# AI が返しうる表記のうち _norm_code() が拾えないもの → 代表記号。
# 2026-09 のシフト表は文字がアウトライン化されていて構造化パースが使えず、
# セルの文字列は Claude の画像読み取りが返した文字列になる。表記ゆれを吸収する。
_EXTRA_CODE_ALIASES = {
    "休": _OFF_CODE, "公": _OFF_CODE, "公休": _OFF_CODE, "希望休": _OFF_CODE,
    "明": _AKE_CODE, "明け": _AKE_CODE, "夜勤明け": _AKE_CODE,
    "有休": _YUKYU_CODE, "有給": _YUKYU_CODE, "年休": _YUKYU_CODE,
    "有給休暇": _YUKYU_CODE,
}

_FILENAME_YEAR_RE = re.compile(r"(20\d{2})\s*年")
_FILENAME_MONTH_RE = re.compile(r"(\d{1,2})\s*月")


def normalize_kdx_code(text) -> "str | None":
    """KDX表のセル文字列 → 代表記号。記号として解釈できなければ None。

    構造化パース用の `_norm_code()` を土台に、AI読み取り経路で出うる表記ゆれ
    （小文字 `a1`、「休」「明け」「有給」といった語）まで寄せる。
    """
    s = str(text or "").strip()
    if not s:
        return None
    code = _norm_code(s)
    if code:
        return code
    upper = unicodedata.normalize("NFKC", s).upper()
    code = _norm_code(upper)
    if code:
        return code
    return _EXTRA_CODE_ALIASES.get(s) or _EXTRA_CODE_ALIASES.get(upper)


def _year_month_from_filename(filename: str) -> tuple["int | None", "int | None"]:
    """ファイル名から年月を拾う（例「KDXオペチーム9月シフト.pdf」→ (None, 9)）。"""
    name = str(filename or "")
    ym = _FILENAME_YEAR_RE.search(name)
    mm = _FILENAME_MONTH_RE.search(name)
    year = int(ym.group(1)) if ym else None
    month = int(mm.group(1)) if mm else None
    if month is not None and not (1 <= month <= 12):
        month = None
    return year, month


def _as_int(value) -> "int | None":
    try:
        return int(value) if value not in (None, "") else None
    except (TypeError, ValueError):
        return None


def force_kdx_code_sheet(
    parsed: dict,
    *,
    filename: str,
    target_year: "int | None" = None,
    target_month: "int | None" = None,
) -> dict:
    """AI読み取り結果（code モード）に KDX の固定凡例を強制適用する。

    2026-09 の勤務シフト表（`KDXオペチーム9月シフト.pdf`）は文字がすべて
    アウトライン化されたベクターPDFで、pdfplumber でも pypdf でも1文字も
    抽出できない（chars=0 / rect=5084 / curve=1521）。そのため
    `is_kdx_shift_pdf()` が False になり構造化パースに乗らず、Claude の
    画像読み取りへ回る。AI は記号グリッドは読めるが「A1＝9:00〜17:30」という
    凡例を持たないので、`build_legend_to_template_name()` が
    `suggest_template_id("A1") == "A1"` に落ち、**生記号がそのまま jinjer
    スケジュールCSVのセルに書かれる**（2026-09 実例）。

    この関数は AI が付けた凡例を捨て、`build_kdx_legend()` の固定凡例
    （A系→9:00〜17:30 / C系→16:30〜34:00）に置き換える。記号の読み取り結果
    （employees）だけを使い、時刻は一切 AI に決めさせない。

    Args:
        parsed: `parse_timesheet_smart()` が返す code モードの dict
        filename: 画面に出す表示名
        target_year, target_month: 画面で指定された対象年月（最優先）

    Returns:
        `parse_kdx_shift_pdf()` と同形の code_sheet ＋ `unknown_codes`
    """
    # --- 年月: 画面指定 > ファイル名 > AI読み取り値 ---
    fn_year, fn_month = _year_month_from_filename(filename)
    year = target_year if target_year is not None else (
        fn_year if fn_year is not None else _as_int(parsed.get("year")))
    month = target_month if target_month is not None else (
        fn_month if fn_month is not None else _as_int(parsed.get("month")))

    days_in_month = None
    if year is not None and month is not None and 1 <= month <= 12:
        days_in_month = calendar.monthrange(year, month)[1]

    # --- 記号の正規化。読めない記号は捨てず unknown_codes に積む ---
    seen_codes: set[str] = set()
    unknown_codes: list[str] = []
    employees: list[dict] = []
    for emp in parsed.get("employees") or []:
        if not isinstance(emp, dict):
            continue
        shifts: list[dict] = []
        for shift in emp.get("shifts") or []:
            if not isinstance(shift, dict):
                continue
            raw = str(shift.get("code") or "").strip()
            code = normalize_kdx_code(raw)
            if code:
                seen_codes.add(code)
            elif raw:
                code = raw            # 捨てない。凡例確認画面で直せるように残す
                if raw not in unknown_codes:
                    unknown_codes.append(raw)
            else:
                code = ""
            new_shift = dict(shift)
            new_shift["code"] = code
            # 年月が確定しているなら日付もその月に揃える（確認画面の表示用。
            # 日の割り当ては exporter が .day しか見ないので結果は変わらない）
            iso = _realign_shift_date(shift.get("date"), year, month, days_in_month)
            if iso is not None:
                new_shift["date"] = iso
            shifts.append(new_shift)
        employees.append({"name": emp.get("name") or "", "shifts": shifts})

    legend = build_kdx_legend({}, seen_codes)

    return {
        "filename": filename,
        "year": year,
        "month": month,
        "legend": legend,
        "employees": employees,
        "off_markers": [_OFF_CODE, _AKE_CODE],
        "source": KDX_SOURCE,
        "unknown_codes": unknown_codes,
        "section_info": {
            "section_index": None,
            # AI読み取り経路なので曜日行の突合はできない（0/0 で「未検証」を表す）
            "weekday_matched": 0,
            "weekday_total": 0,
        },
    }


def _realign_shift_date(raw_date, year, month, days_in_month) -> "str | None":
    """シフトの日付を対象年月へ揃える。揃えられなければ None（元のまま）。"""
    if year is None or month is None or days_in_month is None:
        return None
    try:
        day = date.fromisoformat(str(raw_date).strip()).day
    except (ValueError, TypeError, AttributeError):
        return None
    if not (1 <= day <= days_in_month):
        return None
    return date(year, month, day).isoformat()
