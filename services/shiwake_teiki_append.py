# -*- coding: utf-8 -*-
"""jinjer 仕訳データCSV を補う「通勤定期代」行の生成（イレギュラー救済処置）。

定期代の人は毎月申請しないのが正常なので、申請が無いまま仕訳に載らず支給漏れになる。
2026-08-07 に谷津さんが手作業で拾った流れ（29名31行・562,570円）をそのまま機能にしたもの。
**定常機能ではない**。通勤費申請なしリストを見て、必要なときだけ回す。

入力はファイルでもフォルダでもよい（2026-08-12 谷津さん依頼）。追加で計上した分が
別CSVになるパターンがあるため、フォルダなら中のCSVを全部読み、
**どれかに定期代が計上済みの人はスキップする**＝二重計上が構造的に起きない。
過去にこのツールが出したCSVがフォルダに入っていても同じ仕組みで冪等になる。

出力は**追記行だけの新しいCSV**（ヘッダー＋追記行のみ。2026-08-12 谷津さん決定）。
元ファイルは読み取り専用で一切書かない。統合一覧表はフォルダの中のCSVを全部読むので、
できたCSVを jinjer経費フォルダに追加するだけでよい。以前の「元CSV＋追記の別名コピー」は、
元とコピーが両方フォルダに残ると同じデータを二重に取り込むため廃止した。

手作業で起きた事故を構造的に防ぐ:
  1. Excelを経由したせいで仕訳No.の先頭ゼロが落ちた（00000240 → 240）
     → このモジュールは最初から最後まで文字列だけを扱い、数値化を一切しない。
  2. 追記のついでに既存行が2行ほど書き換わっていた
     → 既存ファイルには書き込みで触れない（standalone出力）。
"""
from __future__ import annotations

import csv
import io
import re
from calendar import monthrange
from dataclasses import dataclass, field
from pathlib import Path

from services.kotsuhi_seisa import (
    COMMUTE_MONTHLY_LIMIT,
    CommuteMaster,
    build_no_commute_rows,
    load_seisa_inputs,
)

ENCODING = "cp932"
LINE_TERMINATOR = "\r\n"
SHIWAKE_COLS = 33
KIND_PASS = "通勤定期代"

# 33列の位置（jinjer 仕訳データCSV。統合一覧表の1〜33列と同じ並び）
C_EMP, C_NAME, C_APPDATE, C_TOTAL, C_MEMO_REQ = 0, 1, 2, 3, 4
C_USEDATE, C_TRANS = 5, 6
C_SUBTOTAL, C_FARE, C_ROUNDTRIP = 12, 13, 14
C_MEMO_LINE = 19
C_ENTRY_DATE, C_ENTRY_DATE8 = 20, 21
C_TAX_DR, C_TAX_CR = 22, 23
C_SHIWAKE_NO, C_ENTRY_TYPE, C_COMPANY = 24, 25, 26
C_PAYMENT = 29
C_BOARD, C_ALIGHT, C_ROUTE = 30, 31, 32

# 支給間隔＝毎月 が定期代。マスタ側の呼び名
INTERVAL_MONTHLY = "毎月"
# 追記行に固定で入れる値（2026-08-07 の手作業成果物から確定）
ROUNDTRIP_ONEWAY = "片道"
TAX_DR, TAX_CR = "10", "0"
PAYMENT_METHOD = "従業員立替"
# 対象にする判定。金額0の行（支給漏れの疑い・勤怠実績なしでマスタ無し）は追記しても意味がない
JUDGE_PAYABLE = "マスタから支給"


class ShiwakeError(ValueError):
    """仕訳CSVが想定と違う／雛形が見つからない等、続けると危ないもの。"""


@dataclass
class TeikiAppendPreview:
    ok: bool = False
    rows: list = field(default_factory=list)        # 追記する行（画面表示用の辞書）
    skipped: list = field(default_factory=list)     # 対象外 {社員番号, 氏名, 理由}
    warnings: list = field(default_factory=list)
    source_rows: int = 0                            # 読み込んだ全CSVのデータ行数合計
    source_files: list = field(default_factory=list)  # [{名前, 行数}] 読み込んだCSVの一覧
    booked_count: int = 0                           # 仕訳データに計上済みでスキップした人数
    ref_shiwake_no: str = ""                        # 既存行から引き継ぐ仕訳No.（文字列のまま）
    ref_company: str = ""
    ref_source: str = ""                            # 仕訳No.をどのファイルから取ったか
    append_count: int = 0
    append_total: int = 0
    error: str = ""


def month_bounds(month: str) -> tuple[str, str, str]:
    """YYYY-MM から (利用日, 計上日, 計上日8桁) を作る。表記は jinjer の出力に合わせる。"""
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", month.strip())
    if not m:
        raise ShiwakeError(f"対象月は YYYY-MM 形式で指定してください: {month}")
    y, mo = int(m.group(1)), int(m.group(2))
    if not 1 <= mo <= 12:
        raise ShiwakeError(f"対象月の月が不正です: {month}")
    last = monthrange(y, mo)[1]
    return f"{y}/{mo}/1", f"{y}/{mo}/{last}", f"{y}{mo:02d}{last:02d}"


@dataclass
class ShiwakeFile:
    """読み込んだ仕訳データCSV1つ分。"""
    name: str                      # ファイル名（エラー・スキップ理由の表示用）
    header: list
    body: list                     # 空行を除いたデータ行


def _load_one_shiwake_csv(path: Path) -> ShiwakeFile:
    """仕訳データCSVを1つ読む。

    BOM付き・33列でない・CP932で読めない、のいずれもここで止める。
    黙って別形式を受け入れると、列がずれた行を給与の元データに混ぜてしまう。
    エラーには必ずファイル名を入れる（フォルダ指定だとどれが悪いか分からないため）。
    """
    raw = path.read_bytes()
    if raw[:3] == b"\xef\xbb\xbf":
        raise ShiwakeError(
            f"仕訳データCSVにBOMが付いています（jinjerからの生CSVを使ってください）: {path.name}")
    try:
        text = raw.decode(ENCODING)
    except UnicodeDecodeError as e:
        raise ShiwakeError(f"仕訳データCSVをCP932として読めません（{e}）: {path.name}") from e
    rows = list(csv.reader(io.StringIO(text, newline="")))
    if not rows:
        raise ShiwakeError(f"仕訳データCSVが空です: {path.name}")
    header = rows[0]
    if len(header) != SHIWAKE_COLS:
        raise ShiwakeError(
            f"仕訳データCSVの列数が {len(header)} です"
            f"（{SHIWAKE_COLS} 列でなければなりません）: {path.name}")
    body = [r for r in rows[1:] if any(str(c).strip() for c in r)]
    return ShiwakeFile(name=path.name, header=header, body=body)


def load_shiwake_sources(path: "str | Path") -> list[ShiwakeFile]:
    """仕訳データCSVを読む。ファイルなら1つ、フォルダなら中の *.csv を全部。

    追加で計上した分が別CSVになるパターンがあるため（例: 8/6本体＋8/7追加計上分）、
    フォルダごと渡せば全部を突合対象にできる（2026-08-12 谷津さん依頼）。
    並びはファイル名ソート＝統合一覧表のフォルダ読み（keihi_summary）と同じ。
    """
    p = Path(path)
    if not p.exists():
        raise ShiwakeError(f"仕訳データCSVが見つかりません: {p}")
    files = sorted(p.glob("*.csv")) if p.is_dir() else [p]
    if not files:
        raise ShiwakeError(f"フォルダにCSVファイルがありません: {p}")
    return [_load_one_shiwake_csv(f) for f in files]


def booked_pass_members(files: list[ShiwakeFile], month: str) -> dict:
    """いずれかのCSVに**対象月の通勤定期代**が既に計上されている人 {社員番号: ファイル名}。

    追加計上分のCSVや、このツールが過去に出したCSVがフォルダに入っていても、
    ここで拾ってスキップするので二重計上にならない（＝再実行しても冪等）。
    """
    from services.kotsuhi_seisa import year_month
    booked: dict = {}
    for f in files:
        for r in f.body:
            if len(r) > C_TRANS and r[C_TRANS] == KIND_PASS \
                    and year_month(r[C_USEDATE]) == month:
                emp = str(r[C_EMP] or "").strip()
                if emp:
                    booked.setdefault(emp, f.name)
    return booked


def find_teiki_reference(files: list[ShiwakeFile]) -> tuple[str, str, str]:
    """既存の「通勤定期代」行から (仕訳No., 企業名, 取得元ファイル名) を**文字列のまま**取る。

    仕訳No. は月ごと・出力ごとに変わる（8/6版=239 / 8/7版=240）ので固定値にできない。
    ファイル名ソートの**最後（=最新）から遡って**最初に見つかった定期代行を採用する。
    全ファイルに無ければ止める（雛形が無い月に当て推量で書くと、取込側で別仕訳になる）。
    """
    for f in reversed(files):
        for r in f.body:
            if len(r) > C_COMPANY and r[C_TRANS] == KIND_PASS:
                no = str(r[C_SHIWAKE_NO] or "").strip()
                company = str(r[C_COMPANY] or "").strip()
                if no:
                    return no, company, f.name
    raise ShiwakeError(
        "仕訳データCSVに「通勤定期代」の既存行がありません。"
        "仕訳No.・企業名を引き継げないため中止しました"
        "（対象月の定期代申請が1件も無いファイルの可能性があります）。")


def build_teiki_append_rows(no_commute_rows: list[dict], master: CommuteMaster,
                            month: str, limit_exempt: dict | None = None,
                            monthly_limit: int = COMMUTE_MONTHLY_LIMIT,
                            booked: dict | None = None,
                            ) -> tuple[list[dict], list[dict], list[str]]:
    """通勤費申請なしリストから追記対象を組み立てる。(追記行, スキップ, 警告) を返す。

    booked は仕訳データに定期代が計上済みの人 {社員番号: ファイル名}
    （booked_pass_members の戻り値）。該当者は理由付きでスキップする。

    1人が複数経路を持つ場合は**経路ごとに1行**出し、備考に「経路N本のうちM本目」を入れる。
    乗継の合算なのか勤務地の選択制なのかはマスタから判別できないため、
    合算して1行にすると選択制の人を二重支給しかねない。人が見て消せる形にする。
    """
    limit_exempt = limit_exempt or {}
    booked = booked or {}
    rows: list[dict] = []
    skipped: list[dict] = []
    warnings: list[str] = []

    for src in no_commute_rows:
        emp = str(src.get("社員番号") or "")
        name = str(src.get("氏名") or "")
        kubun = str(src.get("区分") or "")
        judge = str(src.get("判定") or "")

        if kubun != KIND_PASS:
            skipped.append({"社員番号": emp, "氏名": name,
                            "理由": f"区分が{kubun or '未設定'}（通勤定期代のみ追記します）"})
            continue
        if judge != JUDGE_PAYABLE:
            skipped.append({"社員番号": emp, "氏名": name,
                            "理由": f"判定が「{judge}」（精査で確定してから追記してください）"})
            continue
        if emp in booked:
            # 追加計上分のCSVや過去の追記出力に既に載っている＝追記したら二重計上になる
            skipped.append({"社員番号": emp, "氏名": name,
                            "理由": f"仕訳データに計上済み（{booked[emp]}）"})
            continue

        legs = master.legs(emp, INTERVAL_MONTHLY)
        if not legs:
            skipped.append({"社員番号": emp, "氏名": name,
                            "理由": "マスタに毎月支給の経路が無い（支給額0）"})
            continue

        other = master.other_interval_total(emp, INTERVAL_MONTHLY)
        if other:
            warnings.append(
                f"⚠️ {emp} {name}: マスタに毎日支給の経路も {other:,}円 あります。"
                f"追記するのは毎月支給分だけです。")

        total_monthly = sum(m["支給金額"] for m in legs)
        over_limit = total_monthly > monthly_limit and emp not in limit_exempt
        no_attendance = not src.get("出勤日数")

        for i, leg in enumerate(legs, start=1):
            flags = []
            if str(leg.get("利用交通機関") or "") == "車":
                # 既存の通勤定期代は全員が公共交通機関。車は費目自体が違う可能性がある
                flags.append("車通勤")
            if len(legs) > 1:
                flags.append(f"経路{len(legs)}本のうち{i}本目")
            if no_attendance:
                flags.append("勤怠実績0日")
            if over_limit:
                flags.append(f"上限{monthly_limit:,}円超（免除リスト外）")
            rows.append({
                "社員番号": emp, "氏名": name,
                "経路No": leg.get("経路No"),
                "乗車場所": str(leg.get("出発") or ""),
                "降車場所": str(leg.get("到着") or ""),
                "利用交通機関": str(leg.get("利用交通機関") or ""),
                "金額": int(leg["支給金額"]),
                "要確認": "、".join(flags),
                "備考(明細)": _memo(month, flags),
            })
    return rows, skipped, warnings


def _memo(month: str, flags: list[str]) -> str:
    y, mo = month.split("-")
    base = f"通勤費申請なし・マスタから支給（{int(y)}年{int(mo)}月）"
    return base + ("　※要確認:" + "、".join(flags) if flags else "")


def render_shiwake_rows(preview_rows: list[dict], month: str,
                        ref_shiwake_no: str, ref_company: str) -> list[list[str]]:
    """追記行を33列のCSV行へ整形する。数値化は一切しない（全部文字列）。"""
    use_date, entry_date, entry_date8 = month_bounds(month)
    out: list[list[str]] = []
    for r in preview_rows:
        row = [""] * SHIWAKE_COLS
        row[C_EMP] = str(r["社員番号"])
        row[C_NAME] = str(r["氏名"])
        row[C_APPDATE] = ""                      # 申請が無いので空
        amount = str(int(r["金額"]))
        row[C_TOTAL] = amount
        row[C_USEDATE] = use_date
        row[C_TRANS] = KIND_PASS
        row[C_SUBTOTAL] = amount
        row[C_FARE] = amount
        row[C_ROUNDTRIP] = ROUNDTRIP_ONEWAY      # 定期代は往復＝片道・小計＝マスタ支給額
        row[C_MEMO_LINE] = r["備考(明細)"]
        row[C_ENTRY_DATE] = entry_date
        row[C_ENTRY_DATE8] = entry_date8
        row[C_TAX_DR], row[C_TAX_CR] = TAX_DR, TAX_CR
        row[C_SHIWAKE_NO] = ref_shiwake_no
        row[C_ENTRY_TYPE] = "計上仕訳"
        row[C_COMPANY] = ref_company
        row[C_PAYMENT] = PAYMENT_METHOD
        row[C_BOARD] = r["乗車場所"]
        row[C_ALIGHT] = r["降車場所"]
        row[C_ROUTE] = ""                        # 既存の定期代行と同じく空欄
        out.append(row)
    return out


def check_cp932(rows: list[list[str]]) -> list[str]:
    """CP932にできない文字を書き出す前に洗い出す（中途半端なファイルを残さない）。"""
    problems: list[str] = []
    for i, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            if not value:
                continue
            try:
                str(value).encode(ENCODING)
            except UnicodeEncodeError as e:
                bad = str(value)[e.start:e.end]
                problems.append(
                    f"追記{i}行目 列{c}: {value!r} にCP932で書けない文字 "
                    f"{bad!r} (U+{ord(bad[0]):04X}) があります")
    return problems


def write_standalone_shiwake_csv(header: list, new_rows: list[list[str]],
                                 out_path: "str | Path") -> list[str]:
    """ヘッダー＋追記行だけの新しいCSVを書き、読み直して検証する。戻り値は問題一覧（空ならOK）。

    既存ファイルには一切書かない。できたCSVを jinjer経費フォルダに**追加**すれば、
    統合一覧表がフォルダごと読むので既存分と一緒に取り込まれる（8/7追加計上CSVと同じ扱い）。
    """
    out_path = Path(out_path)
    if len(header) != SHIWAKE_COLS:
        raise ShiwakeError(f"ヘッダーが {len(header)} 列です（{SHIWAKE_COLS} 列のはず）")
    for i, row in enumerate(new_rows, start=1):
        if len(row) != SHIWAKE_COLS:
            raise ShiwakeError(f"追記{i}行目が {len(row)} 列です（{SHIWAKE_COLS} 列のはず）")
    problems = check_cp932([list(header)] + new_rows)
    if problems:
        return problems

    buf = io.StringIO()
    w = csv.writer(buf, lineterminator=LINE_TERMINATOR, quoting=csv.QUOTE_MINIMAL)
    w.writerow(header)
    w.writerows(new_rows)
    data = buf.getvalue().encode(ENCODING, errors="strict")   # 置換は禁止
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(data)

    # 読み直して中身がそのまま入っていることを確かめる（このCSVは給与の元データになる）
    written = out_path.read_bytes()
    if written[:3] == b"\xef\xbb\xbf":
        problems.append("BOMが付いています")
    if not written.endswith(b"\r\n"):
        problems.append("末尾がCRLFではありません")
    try:
        rows = list(csv.reader(io.StringIO(written.decode(ENCODING), newline="")))
    except UnicodeDecodeError as e:
        problems.append(f"書いたファイルをCP932として読み直せません: {e}")
        return problems
    want = [list(header)] + [list(r) for r in new_rows]
    if len(rows) != len(want):
        problems.append(f"行数が合いません（書いたはず {len(want)} / ファイル {len(rows)}）")
    for i, (a, b) in enumerate(zip(want, rows)):
        if a != b:
            problems.append(f"{i}行目が書いた内容と違います")
    return problems


def _collect(kotsuhi_csv, check_xlsx, shiwake_csv, month, target_list=None,
             excluded_list=None, limit_exempt_list=None, monthly_limit=COMMUTE_MONTHLY_LIMIT,
             log_func=print):
    """プレビューと生成で完全に同じ計算を通すための共通部分。"""
    files = load_shiwake_sources(shiwake_csv)
    ref_no, ref_company, ref_source = find_teiki_reference(files)
    booked = booked_pass_members(files, month)
    total_rows = sum(len(f.body) for f in files)
    log_func(f"[info] 仕訳データCSV: {len(files)}ファイル・計{total_rows}行"
             f"（仕訳No. {ref_no}＝{ref_source} から / 定期代の計上済み {len(booked)}名）")

    inputs = load_seisa_inputs(Path(kotsuhi_csv), Path(check_xlsx), month,
                               target_list=target_list, excluded_list=excluded_list,
                               limit_exempt_list=limit_exempt_list)
    no_commute = build_no_commute_rows(inputs.details, inputs.idx, inputs.master,
                                       inputs.workdays, inputs.target_ids, inputs.excluded)
    log_func(f"[info] 通勤費申請なし: {len(no_commute)}名")
    rows, skipped, warnings = build_teiki_append_rows(
        no_commute, inputs.master, month, inputs.limit_exempt, monthly_limit, booked=booked)
    return files, ref_no, ref_company, ref_source, booked, rows, skipped, warnings


def run_teiki_shiwake_preview(kotsuhi_csv, check_xlsx, shiwake_csv, month,
                              target_list=None, excluded_list=None,
                              limit_exempt_list=None,
                              monthly_limit: int = COMMUTE_MONTHLY_LIMIT,
                              log_func=print) -> TeikiAppendPreview:
    """追記内容を計算して返す（ファイルは1つも作らない）。"""
    res = TeikiAppendPreview()
    try:
        files, ref_no, ref_company, ref_source, booked, rows, skipped, warnings = _collect(
            kotsuhi_csv, check_xlsx, shiwake_csv, month, target_list,
            excluded_list, limit_exempt_list, monthly_limit, log_func)
    except ShiwakeError as e:
        res.error = str(e)
        log_func(f"[error] {e}")
        return res
    except Exception as e:  # noqa: BLE001
        res.error = f"追記内容の計算に失敗しました: {e}"
        log_func(f"[error] {res.error}")
        return res

    res.ok = True
    res.rows, res.skipped, res.warnings = rows, skipped, warnings
    res.source_files = [{"名前": f.name, "行数": len(f.body)} for f in files]
    res.source_rows = sum(len(f.body) for f in files)
    res.booked_count = len(booked)
    res.ref_shiwake_no, res.ref_company, res.ref_source = ref_no, ref_company, ref_source
    res.append_count = len(rows)
    res.append_total = sum(r["金額"] for r in rows)
    log_func(f"[info] 追記対象: {res.append_count}行 / 計 {res.append_total:,}円"
             f"（対象外 {len(skipped)}名）")
    return res


def run_teiki_shiwake_generate(kotsuhi_csv, check_xlsx, shiwake_csv, month, out_path,
                               expected_count: int, expected_total: int,
                               target_list=None, excluded_list=None,
                               limit_exempt_list=None,
                               monthly_limit: int = COMMUTE_MONTHLY_LIMIT,
                               log_func=print) -> TeikiAppendPreview:
    """プレビューと同じ計算をやり直し、追記行だけの新CSVを書く。

    件数・合計がプレビューと1つでも違えば書かずに止める
    （承認が進んで入力が変わった＝人が見ていない行が混ざる）。
    """
    res = TeikiAppendPreview()
    try:
        files, ref_no, ref_company, ref_source, booked, rows, skipped, warnings = _collect(
            kotsuhi_csv, check_xlsx, shiwake_csv, month, target_list,
            excluded_list, limit_exempt_list, monthly_limit, log_func)
    except ShiwakeError as e:
        res.error = str(e)
        log_func(f"[error] {e}")
        return res
    except Exception as e:  # noqa: BLE001
        res.error = f"追記内容の計算に失敗しました: {e}"
        log_func(f"[error] {res.error}")
        return res

    res.rows, res.skipped, res.warnings = rows, skipped, warnings
    res.source_files = [{"名前": f.name, "行数": len(f.body)} for f in files]
    res.source_rows = sum(len(f.body) for f in files)
    res.booked_count = len(booked)
    res.ref_shiwake_no, res.ref_company, res.ref_source = ref_no, ref_company, ref_source
    res.append_count = len(rows)
    res.append_total = sum(r["金額"] for r in rows)

    if not rows:
        res.error = "追記する行がありません（通勤費申請なしリストに定期代の対象がいません）。"
        log_func(f"[error] {res.error}")
        return res
    if (res.append_count, res.append_total) != (expected_count, expected_total):
        res.error = (
            "プレビューの内容と一致しません"
            f"（プレビュー {expected_count}行/{expected_total:,}円 → 今回 "
            f"{res.append_count}行/{res.append_total:,}円）。"
            "承認が進んで入力が変わった可能性があります。もう一度プレビューからやり直してください。")
        log_func(f"[error] {res.error}")
        return res

    # ヘッダーは読み込んだ先頭ファイルのものをそのまま使う（全ファイル33列検証済み）
    problems = write_standalone_shiwake_csv(
        files[0].header, render_shiwake_rows(rows, month, ref_no, ref_company), out_path)
    if problems:
        Path(out_path).unlink(missing_ok=True)
        res.error = "追記ファイルの検証に失敗したため出力を取り消しました:\n" + "\n".join(problems)
        log_func(f"[error] {res.error}")
        return res

    res.ok = True
    log_func(f"[info] 追記CSVを作りました: {out_path}"
             f"（{res.append_count}行 / 計 {res.append_total:,}円）")
    return res
