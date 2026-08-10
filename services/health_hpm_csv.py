"""HPM取込用CSV（302列・CP932）を組み立てて書き出す

**血圧について守ること（このモジュールの存在理由）**

    1回目と2回目は別の列（50/51 と 52/53）へ原票のまま出す。
    平均を計算しない。1回分しかなければ2回目は空欄のままにする。
    このモジュールに血圧の算術は一切書かない。値は文字列のまま素通しする。

**判定について守ること**

    原票の判定A〜GはHPMのマスタコードとは体系が違うので転記しない。
    判定列（183〜197）は常に空欄。`source_judgement` はここから参照しない。

**書き出しについて守ること**

    共有フォルダのCSVがExcelで開き直されて 6名→4名 に行が消え、受診番号の
    先頭ゼロも落ちた事故がある。そのため
      - 全文をCP932で strict エンコードしてから 'xb' で一括書き込み（上書き不可）
      - 書いた後にバイトから読み直して、行数・列数・全セルを突き合わせる
    をセットで行う。
"""

from __future__ import annotations

import csv
import io
import logging
import os
import re
import unicodedata
from datetime import date

from services.health_hpm_excel import Issue
from services.health_hpm_master import BP_EXPECTED_COLS, JUDGEMENT_COLS, TOTAL_COLS
from services.health_hpm_match import gender_to_hpm

logger = logging.getLogger(__name__)

# 識別列（0始まり）。行を組み立てる側が埋める。
COL_NAME = 6
COL_KANA = 7
COL_BIRTH = 8
COL_GENDER = 9
COL_AGE = 10
COL_COURSE = 18
COL_EXAM_DATE = 19
COL_EXAM_NO = 20
COL_VENUE = 21
COL_LOCATION = 23

ENCODING = "cp932"
LINE_TERMINATOR = "\r\n"

# 会計年度の始まり（4月）
FISCAL_START_MONTH = 4

_INVALID_FILENAME_RE = re.compile(r'[\\/:*?"<>|]')


class OutputExistsError(FileExistsError):
    """同名ファイルがある。上書きは絶対にしない。"""


# ---------------------------------------------------------------------------
# 小物
# ---------------------------------------------------------------------------

def fiscal_year(d: date) -> int:
    """会計年度（4月始まり）。2027-03-31 は 2026年度。"""
    return d.year if d.month >= FISCAL_START_MONTH else d.year - 1


def format_date_yyyymmdd(d: date | None) -> str:
    return d.strftime("%Y%m%d") if d else ""


# 全角カタカナ → 半角カナ。NFDで濁点を分解してから1文字ずつ置き換える
# （ガ → カ + U+3099 → ｶ + ﾞ）。NFKCは半角化の向きが逆なので使えない。
_KANA_FULL = ("アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホ"
              "マミムメモヤユヨラリルレロワヲンァィゥェォッャュョー・「」、。")
_KANA_HALF = ("ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎ"
              "ﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｧｨｩｪｫｯｬｭｮｰ･｢｣､｡")
_KANA_TABLE = dict(zip(_KANA_FULL, _KANA_HALF))
_COMBINING_VOICED = "゙"
_COMBINING_SEMI_VOICED = "゚"
_HIRAGANA_START, _HIRAGANA_END = 0x3041, 0x3096


def zen_to_han_kana(text: str) -> str:
    """カナ氏名をHPMの表記（半角カナ）に直す。ひらがなはカタカナに寄せる。"""
    if not text:
        return ""
    out: list[str] = []
    for ch in unicodedata.normalize("NFD", str(text)):
        if _HIRAGANA_START <= ord(ch) <= _HIRAGANA_END:
            ch = chr(ord(ch) + 0x60)  # ひらがな → カタカナ
        if ch == _COMBINING_VOICED:
            out.append("ﾞ")
        elif ch == _COMBINING_SEMI_VOICED:
            out.append("ﾟ")
        elif ch in _KANA_TABLE:
            out.append(_KANA_TABLE[ch])
        elif ch == "　":
            out.append(" ")
        elif ch == "ヴ":
            out.append("ｳﾞ")
        else:
            out.append(ch)
    return "".join(out)


def sanitize_filename_part(text: str) -> str:
    text = _INVALID_FILENAME_RE.sub("_", str(text or ""))
    return re.sub(r"\s+", "_", text).strip("_")


# ---------------------------------------------------------------------------
# 行の組み立て
# ---------------------------------------------------------------------------

def _method_of(metric, rules) -> tuple[object, Issue | None]:
    """検査方式が複数ある項目（HBs/HCV等）で、どの列に出すかを決める。

    原票に方式が明記されているときだけ出す。推測して近い列へ入れない。
    """
    if len(rules) == 1:
        return rules[0], None

    haystack = " ".join([metric.item, metric.original_display, metric.category])
    hits = [r for r in rules if r.method and r.method in haystack]
    if len(hits) == 1:
        return hits[0], None

    methods = "、".join(r.method or "指定なし" for r in rules)
    return None, Issue(
        "warning", "METHOD_UNKNOWN",
        f"{metric.category}/{metric.item}: 検査方式（{methods}）が原票から判別できないため"
        "出力しません。方式が分かる場合は整形済Excelの項目名に方式を入れてください",
    )


def build_person_row(person, employee, course, institution, master,
                     header: list[str] | None = None) -> tuple[list[str], list[Issue]]:
    """1名分の302列。定義した列だけ埋め、それ以外は空欄のまま。"""
    from services.health_hpm_match import age_at

    header = header or master.header
    row = [""] * TOTAL_COLS
    issues: list[Issue] = []

    # --- 識別列（jinjerの登録内容を正とする） ---
    row[COL_NAME] = employee.name
    row[COL_KANA] = zen_to_han_kana(employee.kana)
    row[COL_BIRTH] = format_date_yyyymmdd(employee.birth_date)
    row[COL_GENDER] = gender_to_hpm(employee.gender)
    age = age_at(employee.birth_date, person.exam_date)
    row[COL_AGE] = str(age) if age is not None else ""
    row[COL_COURSE] = course.hpm_value
    row[COL_EXAM_DATE] = format_date_yyyymmdd(person.exam_date)
    row[COL_EXAM_NO] = person.exam_no
    row[COL_VENUE] = master.venue_code
    row[COL_LOCATION] = institution.location_code

    # --- 検査値 ---
    unmapped: list[str] = []
    for metric in person.metrics:
        if metric.category == "血圧" and metric.occurrence not in (1, 2):
            continue  # 3回目以降はHPMへ出さない（警告はExcel解析側で出している）

        rules = master.rules_for(metric.category, metric.item, metric.occurrence)
        if not rules:
            unmapped.append(f"{metric.category}/{metric.item}")
            continue

        rule, issue = _method_of(metric, rules)
        if issue is not None:
            issues.append(Issue(issue.level, issue.code, f"{person.name}: {issue.message}"))
            continue
        if rule is None:
            continue

        row[rule.hpm_col] = metric.value

    if unmapped:
        issues.append(Issue(
            "warning", "UNMAPPED_ITEM",
            f"{person.name}: 変換マスタに無い項目は出力しません（{'、'.join(sorted(set(unmapped)))}）",
        ))

    # --- 最後の砦: 判定列と血圧列を検算する ---
    for col in JUDGEMENT_COLS:
        if row[col]:
            raise AssertionError(
                f"判定列 {col}（{header[col]}）に値が入りました。"
                "原票判定はHPMへ転記しない決まりです"
            )
    _assert_blood_pressure(person, row, header)

    if len(row) != TOTAL_COLS:
        raise AssertionError(f"列数が {len(row)} です（{TOTAL_COLS} でなければなりません）")
    return row, issues


def _assert_blood_pressure(person, row, header) -> None:
    """出来上がった行の血圧が、原票の値そのままか確かめる。

    平均・複製・取り違えが混ざっていないかを、書き出す直前にもう一度見る。
    """
    bp = person.blood_pressure()
    for (item, occurrence), col in BP_EXPECTED_COLS.items():
        key = "sys" if item == "収縮期血圧" else "dia"
        expected = bp.get(occurrence, {}).get(key, "")
        if row[col] != expected:
            raise AssertionError(
                f"{person.name}: {header[col]} が {row[col]!r} になっています。"
                f"原票の値は {expected!r} です（平均・複製はしない決まりです）"
            )


def build_csv_rows(resolved, master) -> tuple[list[list[str]], list[Issue]]:
    """(ヘッダー行 + 人数分の行, 警告)。resolved は (person, employee, course, institution)。"""
    rows: list[list[str]] = [list(master.header)]
    issues: list[Issue] = []
    for person, employee, course, institution in resolved:
        row, row_issues = build_person_row(person, employee, course, institution, master)
        rows.append(row)
        issues.extend(row_issues)
    return rows, issues


# ---------------------------------------------------------------------------
# 書き出しと検証
# ---------------------------------------------------------------------------

def check_cp932(rows: list[list[str]], header: list[str] | None = None) -> list[Issue]:
    """CP932にできない文字を書き出す前に洗い出す。

    strict で書きながら例外を出すと中途半端なファイルが残るので、必ず先に全量見る。
    （実測でだめだったもの: ㎗ / µ(U+00B5) / ² / ⁴）
    """
    issues: list[Issue] = []
    header = header or (rows[0] if rows else [])
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            if not value:
                continue
            try:
                str(value).encode(ENCODING)
            except UnicodeEncodeError as e:
                bad = str(value)[e.start:e.end]
                column = header[c] if c < len(header) else ""
                where = "ヘッダー" if r == 0 else f"{r}行目"
                issues.append(Issue(
                    "error", "CP932_UNENCODABLE",
                    f"{where} の列{c}（{column}）の {value!r} に、"
                    f"CP932にできない文字 {bad!r} (U+{ord(bad[0]):04X}) があります",
                ))
    return issues


def render_csv_text(rows: list[list[str]]) -> str:
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator=LINE_TERMINATOR,
                        quoting=csv.QUOTE_MINIMAL)
    writer.writerows(rows)
    return buffer.getvalue()


def write_hpm_csv(path: str, rows: list[list[str]]) -> int:
    """302列CSVを書く。BOMなし・CRLF・CP932。既存ファイルは上書きしない。"""
    for i, row in enumerate(rows):
        if len(row) != TOTAL_COLS:
            raise ValueError(
                f"{i}行目が {len(row)} 列です（{TOTAL_COLS} 列でなければなりません）"
            )

    text = render_csv_text(rows)
    data = text.encode(ENCODING, errors="strict")  # 置換は禁止

    if os.path.exists(path):
        raise OutputExistsError(f"同名のファイルがあります（上書きはしません）: {path}")
    try:
        # 'x' はファイルシステム側で上書きを不可能にする（事前チェックとの二重防衛）
        with open(path, "xb") as f:
            f.write(data)
    except FileExistsError as e:
        raise OutputExistsError(f"同名のファイルがあります（上書きはしません）: {path}") from e
    return len(data)


def verify_written_csv(path: str, expected_rows: list[list[str]]) -> list[str]:
    """書いたファイルをバイトから読み直して突き合わせる。差異の一覧を返す（空ならOK）。"""
    problems: list[str] = []
    raw = open(path, "rb").read()

    if raw[:3] == b"\xef\xbb\xbf":
        problems.append("BOMが付いています")
    if not raw.endswith(LINE_TERMINATOR.encode("ascii")):
        problems.append("末尾がCRLFではありません")
    try:
        text = raw.decode(ENCODING)
    except UnicodeDecodeError as e:
        problems.append(f"CP932として読み直せません: {e}")
        return problems

    actual = list(csv.reader(io.StringIO(text, newline="")))
    if len(actual) != len(expected_rows):
        problems.append(
            f"行数が違います（書いたはず {len(expected_rows)} / ファイル {len(actual)}）"
        )

    for i, (want, got) in enumerate(zip(expected_rows, actual)):
        if len(got) != TOTAL_COLS:
            problems.append(f"{i}行目の列数が {len(got)} です（{TOTAL_COLS} のはず）")
            continue
        for c, (a, b) in enumerate(zip(want, got)):
            if a != b:
                problems.append(f"{i}行目 列{c}: 書いたはず {a!r} / ファイル {b!r}")
                if len(problems) > 40:
                    problems.append("（差異が多いため以降は省略）")
                    return problems
    return problems


# ---------------------------------------------------------------------------
# 保存先
# ---------------------------------------------------------------------------

def default_output_dir(exam_dates: list[date], base: str) -> str:
    """{base}\\{年度}\\{年度}年度健康診断受診者結果\\CSV格納"""
    if not exam_dates:
        raise ValueError("受診日がありません")
    year = fiscal_year(min(exam_dates))
    return os.path.join(base, str(year), f"{year}年度健康診断受診者結果", "CSV格納")


def mixed_fiscal_years(exam_dates: list[date]) -> list[int]:
    return sorted({fiscal_year(d) for d in exam_dates})


def short_institution_name(master, institution) -> str:
    """ファイル名に使う短い機関名。別名があればいちばん短いものを使う。"""
    aliases = [a for a, official in master.aliases.items()
               if official == institution.name]
    if aliases:
        return min(aliases, key=len)
    return institution.name


def default_output_filename(master, institutions, exam_dates: list[date],
                            person_count: int) -> str:
    """既存の命名に合わせる: HPM取込用_同友会_20260701-0703_6名.csv"""
    names = sorted({i.name for i in institutions})
    if len(names) == 1:
        label = short_institution_name(master, next(iter(institutions)))
    else:
        label = "複数機関"
    start, end = min(exam_dates), max(exam_dates)
    span = format_date_yyyymmdd(start)
    if end != start:
        span = f"{span}-{end.strftime('%m%d')}"
    return f"HPM取込用_{sanitize_filename_part(label)}_{span}_{person_count}名.csv"
