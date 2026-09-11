# -*- coding: utf-8 -*-
"""健康診断申込: Google スプレッドシートのシート名・列名・スキーマ版の正本。

Hub はここに書いた列名でシートを検証してから読む。Apps Script 側も同じ列並びで
書く約束（段階③で Code.gs の HEADERS 定数と突き合わせるテストを置く）。

列は「先頭からこの順で完全一致」。末尾に列を足すのは自由（Hub は先頭N列だけ見る）。
"""

from __future__ import annotations

import datetime as _dt

# 2027.2: 対象者に「年度末年齢」（Hub が jinjer の生年月日から計算）を足し、年齢で健診種別を制限する
# 2027.3: 対象者に「性別」（jinjer の personal.gender）を足し、男性には婦人科検診を出さない。被扶養者の申込は廃止（列は残す）
SCHEMA_VERSION = "2027.3"

SHEET_SETTINGS = "設定"
SHEET_OPTIONS = "選択肢"
SHEET_TARGETS = "対象者"
SHEET_RESPONSES = "回答"
SHEET_AUDIT = "監査ログ"
ALL_SHEETS = (SHEET_SETTINGS, SHEET_OPTIONS, SHEET_TARGETS, SHEET_RESPONSES, SHEET_AUDIT)

# Hub が追記してよいシート。回答・選択肢・設定には決して書かない。
WRITABLE_SHEETS = frozenset({SHEET_TARGETS, SHEET_AUDIT})

SETTINGS_HEADERS = ("キー", "値", "備考")
OPTION_HEADERS = ("区分", "コード", "表示名", "有効", "並び順", "別名", "備考")
TARGET_HEADERS = (
    # 1〜15列: Hub が対象者登録で書く
    "年度", "社員番号", "氏名", "社用メール", "在籍区分",
    "前年度情報元", "前年度健診機関コード", "前年度健診機関名",
    "前年度健診種別コード", "前年度健診種別名", "前年度追加検査", "前年度健診機関(原文)",
    "登録日時", "登録者", "年度末年齢", "性別",
    # 17列〜: Apps Script（案内送信・回答受付）が更新する
    "トークンハッシュ", "送信日時", "送信回数", "初回アクセス日時",
    "申込状態", "受付番号", "回答版", "回答日時", "備考",
)
TARGET_HUB_COLUMNS = 16

# 年齢による健診種別の制限（Apps Script の画面と Hub の検証で同じ規則を使う）
#   年度末（翌年3月31日）時点の満年齢で、34歳以下は定期健康診断だけ、35歳以上は人間ドックA/B/Cから必ず選ぶ。
#   年度末年齢が空（jinjer に生年月日が無い）なら制限しない。
EXAM_TYPE_REGULAR = "10"
EXAM_TYPES_DOCK = ("11", "12", "13")
DOCK_AGE_FROM = 35


# 性別と追加検査（婦人科検診）。性別は jinjer の personal.gender.name をそのまま持つ（男性／女性。空＝不明）
GENDER_MALE = "男性"
EXTRA_GYN = "GYN"


def gyn_allowed(gender) -> bool:
    """婦人科検診を出してよいか。男性だけ出さない（空・不明は出す）。"""
    return str(gender or "").strip() != GENDER_MALE


def fiscal_year_end(fiscal_year: int) -> _dt.date:
    """年度末＝翌年3月31日。"""
    return _dt.date(int(fiscal_year) + 1, 3, 31)


def age_on(birth_date: str | None, on: _dt.date) -> int | None:
    """生年月日（yyyy-mm-dd）から、その日時点の満年齢。読めなければ None。"""
    s = str(birth_date or "").strip().replace("/", "-")
    try:
        b = _dt.date.fromisoformat(s[:10])
    except ValueError:
        return None
    age = on.year - b.year - ((on.month, on.day) < (b.month, b.day))
    return age if age >= 0 else None


def allowed_exam_types(age) -> tuple[str, ...] | None:
    """年度末年齢で選べる健診種別コード。None は制限なし（年齢不明）。"""
    try:
        n = int(str(age).strip())
    except (TypeError, ValueError):
        return None
    return EXAM_TYPES_DOCK if n >= DOCK_AGE_FROM else (EXAM_TYPE_REGULAR,)


def age_band_label(age) -> str:
    allowed = allowed_exam_types(age)
    if allowed is None:
        return "年齢不明（制限なし）"
    return f"{DOCK_AGE_FROM}歳以上（人間ドックA/B/C）" if allowed == EXAM_TYPES_DOCK else f"{DOCK_AGE_FROM - 1}歳以下（定期健康診断）"
RESPONSE_HEADERS = (
    "回答日時", "受付番号", "年度", "社員番号", "回答版", "氏名", "社用メール",
    "申込区分", "健診機関コード", "健診機関名", "その他医療機関名",
    "健診種別コード", "健診種別名", "追加検査", "その他健診予定日",
    "被扶養者申込", "続柄", "被扶養者氏名", "備考", "トークンハッシュ", "回答元",
)
AUDIT_HEADERS = ("日時", "イベント", "実行元", "実行者", "年度", "社員番号", "詳細")

HEADERS_BY_SHEET: dict[str, tuple[str, ...]] = {
    SHEET_SETTINGS: SETTINGS_HEADERS,
    SHEET_OPTIONS: OPTION_HEADERS,
    SHEET_TARGETS: TARGET_HEADERS,
    SHEET_RESPONSES: RESPONSE_HEADERS,
    SHEET_AUDIT: AUDIT_HEADERS,
}

# 設定シートに必ず要るキー（値の妥当性は使う側で見る）
REQUIRED_SETTING_KEYS = (
    "スキーマ版", "年度", "前年度",
    "受付開始", "受付終了", "受診期間開始", "受診期間終了", "回答受付",
)
SETTING_SCHEMA_VERSION = "スキーマ版"
SETTING_FISCAL_YEAR = "年度"
SETTING_PREVIOUS_YEAR = "前年度"

# 対象者「申込状態」の値
STATUS_UNSENT = "未送信"
STATUS_SENT = "送信済"
STATUS_ANSWERED = "回答済"
STATUS_REANSWER = "再回答待ち"
STATUS_INVALID = "無効"
TARGET_STATUSES = (STATUS_UNSENT, STATUS_SENT, STATUS_ANSWERED, STATUS_REANSWER, STATUS_INVALID)

# 前年度情報元
SOURCE_HISTORY = "履歴"
SOURCE_CURRENT = "現在値"
SOURCE_NONE = "なし"

# 申込区分
KIND_SAME = "same"
KIND_CHANGE = "change"

# 監査ログの実行元
ACTOR_HUB = "Hub"
ACTOR_APPS_SCRIPT = "AppsScript"


class SchemaError(ValueError):
    """シート構成や設定値が Hub の想定と食い違うときの例外（画面に文言をそのまま出す）。"""


def _cell(value) -> str:
    return "" if value is None else str(value).strip()


def assert_writable(sheet: str) -> None:
    """書込先が許可シートでなければ例外。Gateway の append_rows が必ず通す。"""
    if sheet not in WRITABLE_SHEETS:
        raise SchemaError(
            f"「{sheet}」シートへの書き込みは許可されていません（書けるのは "
            + "・".join(sorted(WRITABLE_SHEETS)) + " だけ）")


def verify_headers(sheet: str, header_row: list | None) -> list[str]:
    """1行目が想定の列並びかを見て、食い違いを文言のリストで返す（空なら合格）。"""
    expected = HEADERS_BY_SHEET.get(sheet)
    if expected is None:
        return [f"Hub が知らないシートです: {sheet}"]
    actual = [_cell(v) for v in (header_row or [])]
    errors: list[str] = []
    missing = [name for i, name in enumerate(expected) if i >= len(actual) or not actual[i]]
    for i, name in enumerate(expected):
        if i >= len(actual) or not actual[i]:
            continue
        if actual[i] != name:
            errors.append(f"{sheet}シートの{i + 1}列目が想定外です: 「{actual[i]}」（期待: 「{name}」）")
    if missing:
        errors.append(f"{sheet}シートに列がありません: " + "、".join(missing))
    return errors


def split_header(values: list[list] | None) -> tuple[list[str], list[list[str]]]:
    """read_values の戻り（1行目=ヘッダー）を (ヘッダー, データ行) に分ける。"""
    if not values:
        return [], []
    header = [_cell(v) for v in values[0]]
    return header, [[_cell(v) for v in row] for row in values[1:]]


def rows_to_dicts(header: list[str] | tuple[str, ...], rows: list[list]) -> list[dict[str, str]]:
    """データ行を {列名: 値} にする。

    Sheets API は末尾の空セルを切り落として返すので、ヘッダー長までパディングする。
    ヘッダーより長い行の余りは捨てる（末尾に列を足しても Hub は先頭N列しか見ない）。
    全セル空の行は飛ばす。
    """
    names = list(header)
    out: list[dict[str, str]] = []
    for row in rows:
        cells = [_cell(v) for v in row]
        if not any(cells):
            continue
        cells = (cells + [""] * len(names))[:len(names)]
        out.append(dict(zip(names, cells)))
    return out


def settings_to_kv(rows: list[list]) -> dict[str, str]:
    """設定シートのデータ行（キー/値/備考）を {キー: 値} にする。空キーは無視、重複は後勝ち。"""
    kv: dict[str, str] = {}
    for row in rows:
        cells = [_cell(v) for v in row]
        if not cells or not cells[0]:
            continue
        kv[cells[0]] = cells[1] if len(cells) > 1 else ""
    return kv


def verify_settings(settings_kv: dict[str, str], expected_year: int) -> list[str]:
    """設定シートの必須キー・スキーマ版・年度を検査する。"""
    errors: list[str] = []
    missing = [k for k in REQUIRED_SETTING_KEYS if not settings_kv.get(k)]
    if missing:
        errors.append("設定シートに値がありません: " + "、".join(missing))
    version = settings_kv.get(SETTING_SCHEMA_VERSION, "")
    if version and version != SCHEMA_VERSION:
        errors.append(f"設定シートのスキーマ版が違います: 「{version}」（Hub は {SCHEMA_VERSION}）")
    year = settings_kv.get(SETTING_FISCAL_YEAR, "")
    if year and year != str(expected_year):
        errors.append(f"設定シートの年度が違います: 「{year}」（年度設定JSONは {expected_year}）")
    return errors


def verify_workbook(sheet_titles: list[str], headers: dict[str, list],
                    settings_kv: dict[str, str], expected_year: int) -> list[str]:
    """ブック全体の構成検査。個人情報の行を読む前に必ず通す。

    Args:
        sheet_titles: ブック内のシート名一覧
        headers: {シート名: 1行目のセル}
        settings_kv: 設定シートの {キー: 値}
        expected_year: 年度設定JSONの年度
    Returns:
        食い違いの文言リスト（空なら合格）
    """
    errors: list[str] = []
    present = set(sheet_titles or [])
    for name in ALL_SHEETS:
        if name not in present:
            errors.append(f"シートがありません: {name}")
    for name in ALL_SHEETS:
        if name in present:
            errors.extend(verify_headers(name, headers.get(name)))
    if SHEET_SETTINGS in present and not any(e.startswith(f"{SHEET_SETTINGS}シート") for e in errors):
        errors.extend(verify_settings(settings_kv, expected_year))
    return errors
