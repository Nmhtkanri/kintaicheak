"""jinjer給与へのインポート行生成とAPI投入（経費マクロ移植 P2）

集計結果（keihi_classify の bucket）から jinjer給与インポートCSV（11列・谷津さん確定ヘッダー）を作り、
`POST /v1/jinji-imports`（メニュー種別=給与計算、テンプレート「経費APIインポート用」id=44450）で投入する。
※旧テンプレ「経費インポート用」id=37047 は「立替金」裸列が無く計上漏れになるため廃止（2026-07-17 谷津確認）。

列マッピング（live 集計シート実測 2026-07-16。マクロは dRY列振り分け→Module2 の2段変換）:
  - 夜間当番手当       ← F列相当 = 夜間当番(D) + RINK(E)   ※dRYがF値をR/Sペアに入れるため合算値になる
  - 定常外業務対応手当  ← 画面からの手入力（マクロでは「仕訳データ」K/L列 → R/S・T/Uペア）
  - 支給過不足調整     ← **ファイル取込**（1〜30名規模になるため。CSV/xlsxの「社員番号,金額」。負数可）
  - 非課税通勤費       ← 交通費(H)
  - 立替金（顧客請求分）← 顧客請求分(G)
  - 立替金            ← 非課税精算(立替金)(I)
  - その他            ← その他(J)
  - その他手当         ← 画面からの手入力（マクロでは集計シート R/T へ直接手入力）
  - 現物支給          ← 画面からの手入力
  ※テレワーク手当(K) は現行マクロでもインポートCSVに乗らない（dRYでK廃止）。金額があれば警告表示する。

定常外業務対応手当・その他手当・現物支給は経費4ソースから導けない「その月に手で決める」
金額なので画面の手入力欄から受け取る（`parse_manual_allowances`）。支給過不足調整は
対象者が1〜30名規模になり得るためファイル取込にする（`load_allowance_file`）。マクロ側は R/S・T/U の2枠しか無く、
夜間当番手当＋通信手当＋定常外業務対応手当 の3つが揃うと定常外が
「スキップ(両スロット使用済み)」で消える不具合があるが、本実装は項目ごとに独立の
フィールドを持つため取りこぼさない（通信手当はCSV対象外なので枠を食わない）。

投入は人間の最終チェック（プレビューHTML＋CSVダウンロード）を経てから行う（確認後実行）。
"""

from __future__ import annotations

import csv as _csv
import re as _re
import time as _time
from dataclasses import dataclass, field
from pathlib import Path

from services.keihi_classify import B_YAKAN, B_RINK, B_TRANS, B_ETC, B_TW, B_BILL, B_NONTAX

# 谷津さん確定の正式ヘッダー（2026-07-16。jinjerテンプレート「経費インポート用」と一致させること）
IMPORT_HEADERS = [
    "社員番号", "氏名", "夜間当番手当", "定常外業務対応手当", "支給過不足調整",
    "非課税通勤費", "立替金（顧客請求分）", "立替金", "その他", "その他手当", "現物支給",
]
# インポート行が持つ支給項目（社員番号・氏名以外）。テンプレ列名との照合対象。
ITEM_KEYS = [
    "夜間当番手当", "定常外業務対応手当", "支給過不足調整", "非課税通勤費",
    "立替金（顧客請求分）", "立替金", "その他", "その他手当", "現物支給", "テレワーク手当",
]
# jinjer側テンプレート（GET /v1/master/jinji-import-templates で実測）
DEFAULT_TEMPLATE_ID = "44450"   # 「経費APIインポート用」menu=payroll_calculation（11列・立替金裸列あり・2026-07-17検証済）


def _norm_col(s) -> str:
    """列名の照合キー。空白・全角スペース・先頭の必須マーク「*」を除去。"""
    return _re.sub(r"[\s　*＊]", "", "" if s is None else str(s)).strip()


def _match_column(col: str) -> "str | None":
    """テンプレの列名 → インポート行のどのキーに対応するか。空/不明は None。"""
    n = _norm_col(col)
    if n == "":
        return None
    if n in ("社員番号", "従業員番号"):
        return "社員番号"
    if n in ("氏名", "名前"):
        return "氏名"
    for k in ITEM_KEYS:
        if _norm_col(k) == n:
            return k
    return None


def read_template_header(path: "str | Path") -> list[str]:
    """jinjerからダウンロードした空テンプレCSVの先頭行（列名の並び）を返す。

    列位置対応のインポートに合わせるため、空セルもそのまま（"" として）保持する。
    """
    path = Path(path)
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            with open(path, encoding=enc, newline="") as f:
                header = next(_csv.reader(f))
            return [("" if c is None else str(c)).strip() for c in header]
        except (UnicodeDecodeError, StopIteration):
            continue
    raise ValueError(f"テンプレートCSVを読めませんでした: {path}")


def _parse_amount(s) -> "float | None":
    """金額文字列を数値化する（VBA ValJP 相当）。全角・¥・円・カンマ・(負数) に対応。"""
    import unicodedata as _ud
    t = _ud.normalize("NFKC", "" if s is None else str(s)).strip()
    for ch in ("¥", "￥", "円", ",", " ", "　"):
        t = t.replace(ch, "")
    if t == "":
        return None
    if len(t) >= 2 and t.startswith("(") and t.endswith(")"):
        t = "-" + t[1:-1]
    try:
        return float(t)
    except ValueError:
        return None


def parse_manual_allowances(text: "str | None") -> "tuple[dict[str, int], list[str]]":
    """「社員番号,金額」形式の手入力テキストを {社員番号: 金額} にする。

    定常外業務対応手当・その他手当は経費から導けない手決めの金額なので画面入力を受ける。
    区切りは カンマ / タブ / 全角カンマ（Excelの2列をそのままコピペできる）。
    同一社員の複数行は加算（マクロ「仕訳データ振り分け」と同じ挙動）。
    空行・「#」開始行・ヘッダー行は無視。読めなかった行は errors に理由を返す。
    """
    from services.keihi_summary import excel_coerce, in_company_scope
    out: dict[str, int] = {}
    errors: list[str] = []
    for lineno, raw in enumerate((text or "").splitlines(), start=1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        # タブがあればタブ区切り（Excelの2列コピペ）。無ければカンマ区切り。
        # 金額の桁区切りカンマ（8,000）で割れるため、2列目以降は結合してから数値化する。
        parts = line.split("\t") if "\t" in line else _re.split(r"[,、，]", line)
        if len(parts) < 2:
            errors.append(f"{lineno}行目: 「社員番号,金額」の形式で入力してください → {line!r}")
            continue
        emp = excel_coerce(parts[0].strip())    # "02026012" → "2026012"
        amt = _parse_amount("".join(p.strip() for p in parts[1:]))
        is_id = bool(_re.fullmatch(r"\d+", emp))
        if amt is None:
            if not is_id:
                continue    # 「社員番号,金額」等のヘッダー行 → 無視
            errors.append(f"{lineno}行目: 金額を数値として読めません → {parts[1].strip()!r}")
            continue
        if not is_id:
            errors.append(f"{lineno}行目: 社員番号が数字ではありません → {parts[0].strip()!r}")
            continue
        if not in_company_scope(emp):
            errors.append(f"{lineno}行目: {emp} は給与計算対象外（5/6/9始まり）のため除外しました")
            continue
        if amt == 0:
            continue        # 0 は入れても結果が変わらないので無視
        out[emp] = out.get(emp, 0) + int(round(amt))
    return out, errors


# 手当ファイルの列名候補（ヘッダーから自動判別する）
_ID_HEADERS = ("社員番号", "従業員番号", "社員no", "社員ｎｏ", "empno", "employee_id")
_AMT_HEADERS = ("金額", "支給額", "調整額", "支給過不足調整", "現物支給", "amount")


def load_allowance_file(path: "str | Path", label: str = "手当") -> "tuple[dict[str, int], list[str]]":
    """「社員番号・金額」の2列を持つ CSV/xlsx を読み、{社員番号: 金額} にする。

    人数が多い項目（支給過不足調整は1〜30名規模）向け。ヘッダー行の列名から
    社員番号列・金額列を自動判別し、見つからなければ先頭2列を使う。
    金額は負数可（過払いの戻し等）。同一社員の複数行は加算。
    """
    path = Path(path)
    rows: list[tuple] = []
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xltx"):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            rows = [r for r in ws.iter_rows(values_only=True) if r is not None]
        finally:
            wb.close()
    else:
        for enc in ("cp932", "utf-8-sig", "utf-8"):
            try:
                with open(path, encoding=enc, newline="") as f:
                    rows = [tuple(r) for r in _csv.reader(f)]
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError(f"{label}ファイルの文字コードを判別できませんでした: {path}")

    rows = [r for r in rows if any(c is not None and str(c).strip() for c in r)]
    if not rows:
        return {}, [f"{label}ファイルにデータがありません: {path}"]

    # ヘッダーから列位置を判別（無ければ先頭2列）
    head = [_norm_col(c).lower() for c in rows[0]]
    id_col = next((i for i, h in enumerate(head) if h in _ID_HEADERS), None)
    amt_col = next((i for i, h in enumerate(head) if h in _AMT_HEADERS), None)
    if id_col is None or amt_col is None:
        id_col, amt_col = 0, 1
        body = rows if _parse_amount(rows[0][1] if len(rows[0]) > 1 else None) is not None else rows[1:]
    else:
        body = rows[1:]

    text = "\n".join(
        f"{r[id_col]}\t{r[amt_col]}"
        for r in body
        if len(r) > max(id_col, amt_col)
    )
    got, errors = parse_manual_allowances(text)
    return got, [f"{label}ファイル {e}" for e in errors]


def build_import_rows(by_id: dict, emp_names: dict, roster_names: "dict | None" = None,
                      teijo: "dict | None" = None, sonota: "dict | None" = None,
                      genbutsu: "dict | None" = None, kabusoku: "dict | None" = None,
                      ) -> tuple[list[dict], list[str]]:
    """集計（数字ID→7バケット）＋手入力手当からインポート行と警告リストを返す。

    行は 20YY 始まりの社員のみ・社員番号昇順（5/6/9始まりは給与計算対象外）。
    teijo/sonota は `parse_manual_allowances` の戻り（{社員番号: 金額}）。
    **経費が無く手当だけの社員も行を出す**（マクロの「A列に無い社員は無言スキップ」で
    起きた計上漏れ事故と同じ轍を踏まないため）。
    """
    from services.keihi_summary import in_company_scope
    roster_names = roster_names or {}
    teijo = teijo or {}
    sonota = sonota or {}
    genbutsu = genbutsu or {}
    kabusoku = kabusoku or {}
    rows: list[dict] = []
    warnings: list[str] = []
    for emp_id in sorted(set(by_id) | set(teijo) | set(sonota) | set(genbutsu) | set(kabusoku)):
        if not in_company_scope(emp_id):
            continue
        v = by_id.get(emp_id) or [0.0] * 7
        name = emp_names.get(emp_id) or roster_names.get(emp_id, "")
        rows.append({
            "社員番号": emp_id,
            "氏名": name,
            "夜間当番手当": int(v[B_YAKAN] + v[B_RINK]),   # F列相当（夜間＋RINK）
            "定常外業務対応手当": int(teijo.get(emp_id, 0)),
            "支給過不足調整": int(kabusoku.get(emp_id, 0)),
            "非課税通勤費": int(v[B_TRANS]),
            "立替金（顧客請求分）": int(v[B_BILL]),
            "立替金": int(v[B_NONTAX]),
            "その他": int(v[B_ETC]),
            "その他手当": int(sonota.get(emp_id, 0)),
            "現物支給": int(genbutsu.get(emp_id, 0)),
            "テレワーク手当": int(v[B_TW]),
        })

    # 手入力の社員番号が在籍者一覧に無ければ入力ミスの可能性 → 警告（行は出す）
    for emp_id in sorted(set(teijo) | set(sonota) | set(genbutsu) | set(kabusoku)):
        if in_company_scope(emp_id) and roster_names and emp_id not in roster_names:
            warnings.append(
                f"⚠️ 手入力手当: 社員番号 {emp_id} が在籍者一覧にありません（入力ミスの可能性）。")
    return rows, warnings


def _escape_csv(s: str) -> str:
    """Module2.EscapeCSV: 文字列は必ずダブルクォート囲み・内部の " は "" に。"""
    return '"' + str(s).replace('"', '""') + '"'


def check_template_coverage(rows: list[dict], template_header: list[str]) -> list[str]:
    """テンプレに列が無いのに金額が発生している支給項目を警告として返す。

    位置対応インポートでは、テンプレに列が無い項目はサイレントに捨てられて計上漏れになる
    （今回の 川口さん立替金12,177円・夜間当番手当30,000円の事故がこれ）。事前に検知する。
    """
    covered = {_match_column(c) for c in template_header}
    warnings: list[str] = []
    for key in ITEM_KEYS:
        if key in covered:
            continue
        total = sum(int(r.get(key, 0) or 0) for r in rows)
        n = sum(1 for r in rows if int(r.get(key, 0) or 0))
        if total:
            warnings.append(
                f"⚠️ テンプレに「{key}」の列がありません → {n}名・計{total:,}円が"
                f"取り込まれず計上漏れになります。jinjer側テンプレートに「{key}」列を追加してください。")
    return warnings


def render_import_csv(rows: list[dict], template_header: "list[str] | None" = None) -> bytes:
    """インポートCSVを Shift-JIS バイト列で返す（Module2 と同形式: 文字列のみクォート）。

    template_header を渡すと、その列名の並び（空セルも保持）に合わせて列を組み立てる。
    jinjer のテンプレインポートは列名でなく**列位置**で対応付けるため、テンプレの並びに
    追従させることが必須（並びがズレると値が別項目に入る）。未指定時は IMPORT_HEADERS。
    """
    header = template_header if template_header is not None else IMPORT_HEADERS
    lines = [",".join(header)]   # ヘッダーは列名をそのまま（テンプレの見出しを忠実に）
    for r in rows:
        cells = []
        for col in header:
            key = _match_column(col)
            if key in ("社員番号", "氏名"):
                cells.append(_escape_csv(r.get(key, "")))
            elif key in ITEM_KEYS:
                cells.append(str(int(r.get(key, 0) or 0)))
            else:
                cells.append("")   # 空列・不明列はそのまま空（テンプレのスキップ列）
        lines.append(",".join(cells))
    return ("\r\n".join(lines) + "\r\n").encode("cp932")


def write_import_csv(rows: list[dict], out_path: "str | Path",
                     template_header: "list[str] | None" = None) -> Path:
    """インポートCSVをファイルに書き出す（人間チェック用ダウンロード）。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(render_import_csv(rows, template_header))
    return out_path


# ----------------------------------------------------------------------
# jinjer API 投入（確認後実行）
# ----------------------------------------------------------------------

@dataclass
class PayrollImportResult:
    ok: bool
    month: str = ""
    file_name: str = ""
    status: str = ""      # "1"=全行成功 / "2"=一部失敗 / "0"=処理待ち・不明
    total_rows: int = 0
    failed_rows: int = 0
    import_id: str = ""
    error: str = ""
    logs: list = field(default_factory=list)


def post_payroll_import(
    client,
    month: str,
    csv_bytes: bytes,
    file_name: str,
    template_id: str = DEFAULT_TEMPLATE_ID,
    apply_formulas_off: bool = True,
    poll_seconds: int = 180,
    log_func=print,
) -> PayrollImportResult:
    """給与計算インポートを `POST /v1/jinji-imports` で投入し、完了をポーリングする。

    Args:
        client: JinjerClient（認証済み）
        month: 処理月 "YYYY-MM"（salary_setting.executed_on）
        csv_bytes: Shift-JIS のインポートCSV
        file_name: インポートファイル名
        template_id: 入力テンプレートID（既定 44450「経費APIインポート用」）
        apply_formulas_off: 「編集した項目は計算式を適用しない」を有効にするか
        poll_seconds: 完了待ちの最大秒数（バッチ処理のため時間がかかる）
    """
    result = PayrollImportResult(ok=False, month=month, file_name=file_name)

    def log(msg):
        result.logs.append(msg)
        log_func(msg)

    try:
        data = client.post_jinji_import(
            csv_bytes=csv_bytes, file_name=file_name, template_id=template_id,
            executed_on=month, apply_formulas_off=apply_formulas_off,
        )
        result.import_id = str(data.get("id") or "")
        log(f"[info] jinji-imports 投入予約 OK: {file_name} "
            f"(処理月 {month} / テンプレ {template_id} / id={result.import_id or '不明'})")
    except Exception as e:  # noqa: BLE001
        result.error = f"投入予約に失敗しました: {e}"
        log(f"[error] {result.error}")
        return result

    # ポーリング。jinji-imports の GET には status フィールドが無く、
    # number_of_failed_rows / number_of_total_rows で成否を判定する（2026-07-16 実測。
    # updated_at が入ればバッチ処理完了とみなす）。
    if not result.import_id:
        result.ok = True
        log("[warn] 投入予約は成功しましたが id が取得できず完了確認をスキップします。"
            "jinjer画面のインポート履歴をご確認ください。")
        return result

    waited = 0
    interval = 10
    while waited < poll_seconds:
        _time.sleep(interval)
        waited += interval
        try:
            item = client.find_jinji_import(result.import_id)
        except Exception as e:  # noqa: BLE001
            log(f"[warn] 状況確認に失敗（リトライ）: {e}")
            continue
        if item and str(item.get("updated_at") or "").strip():
            result.total_rows = int(item.get("number_of_total_rows") or 0)
            result.failed_rows = int(item.get("number_of_failed_rows") or 0)
            if result.failed_rows == 0:
                result.status = "1"
                result.ok = True
                log(f"[done] インポート成功: {result.total_rows}行すべて取込（{waited}秒）。"
                    "jinjer給与画面で金額をご確認ください。")
            else:
                result.status = "2"
                result.ok = False
                result.error = (
                    f"{result.total_rows}行中 {result.failed_rows}行が失敗しました。"
                    "通知メールとjinjer画面のインポート履歴でエラー行をご確認ください。")
                log(f"[error] {result.error}")
            return result
        log(f"[info] 処理待ち... ({waited}秒)")

    result.status = "0"
    result.ok = True  # 予約自体は成功（完了待ちタイムアウトは失敗扱いにしない）
    log("[warn] 完了確認がタイムアウトしました（投入予約は成功）。"
        "後ほどjinjer画面のインポート履歴か通知メールをご確認ください。")
    return result
