# -*- coding: utf-8 -*-
r"""請求書モードが見に行くフォルダを、画面から一覧・編集するサービス（2026-08-20 谷津さん決定）

画面では3分類で見せる。

  ① 勤怠フォルダ       … 勤怠フォルダ ＋ 勤怠ファイル（* が使える）
  ② 請求書作成Excel     … 請求書Excel ＋ シート名
  ③ 請求書格納フォルダ   … 出力フォルダ ＋ 出力ファイル名（提出用PDFの置き場）
                        ＋ freee取込CSVが請求書PDFを探しに行く会社フォルダ

③に2種類が同居するのは、**同じ場所を違う粒度で指しているから**。アクシスの
「,提出データ」は①で作った提出用PDFの置き場であり、同時に②が請求書PDFを
探しに行く先でもある。別々の画面に分けると片方だけ直して食い違うので、
1つのタブに「人単位の出力先」と「会社単位の探索ルート」として並べる。

実体は既存の2本のCSVのまま。壊さないよう、列は後方互換で足すだけにする。

  請求書モード_PDF作成設定.csv   … 1人1行（①②③の人単位ぶん）＋「対象」列を追加
  請求書モード_対象フォルダ.csv  … 1社1行（③の会社単位ぶん）＋「取引先」列を追加

追加した列は、無くても今までどおり動く（対象＝空なら対象、取引先＝空なら
フォルダ名から作る）。知らない列が入っていても消さずに書き戻す。

書き込みは許可ユーザーのみ。共有exeは5名が使うので、ここが壊れると
**別の人の勤怠を綴じた請求書**ができうる。判定と許可リストの読み方は
SAP台帳と同じものを使い回す（services.sap_import_ledger）。
"""

from __future__ import annotations

import csv
import datetime as dt
import hashlib
import io
import logging
import re
import shutil
import unicodedata
from pathlib import Path
from typing import Any, Iterable, Sequence

# 書き込み許可の判定は台帳と同じものを使う（許可リストCSVのパスは引数で渡す）。
from services.sap_import_ledger import can_write, current_user, load_writers  # noqa: F401
from services.invoice_pdf import expand, fiscal_year  # noqa: F401

logger = logging.getLogger(__name__)


class InvoiceFoldersError(ValueError):
    """設定の読み書きができないときに投げる。"""


class InvoiceFoldersConflict(RuntimeError):
    """自分が画面を開いた後に、他の人が同じCSVを更新していた。"""


# 1人1行。この順で書き戻す。「対象」は今回追加（無い既存CSVでも読める）。
PEOPLE_COLUMNS = [
    "対象", "取引先", "氏名",
    "請求書Excel", "シート名",
    "勤怠フォルダ", "勤怠ファイル",
    "出力フォルダ", "出力ファイル名",
]

# 1社1行。「取引先」は表示用に今回追加。
ROOT_COLUMNS = ["対象", "取引先", "フォルダパス"]

# 「対象」列を外す言葉。invoice_mode.load_target_roots と同じ判定にそろえる。
_OFF_TOKENS = {"0", "false", "no", "対象外", "無効"}

# 判定の重さ。画面のバッジ色に対応する。
LEVEL_OK = "ok"       # そのまま作れる
LEVEL_WARN = "warn"   # 作れるが人が見たほうがよい／作成時に止まる
LEVEL_STOP = "stop"   # materialが無い。直さないと作れない
LEVEL_INFO = "info"   # 参考情報（確認していない等）
LEVEL_OFF = "off"     # 対象外（チェックを外している）


# ----------------------------------------------------------------------
# 正規化
# ----------------------------------------------------------------------

def _nfkc(value: Any) -> str:
    return unicodedata.normalize("NFKC", str(value or "")).strip()


def is_on(value: Any) -> bool:
    """「対象」列の判定。空欄は対象とみなす（列が無い既存CSVと同じ扱い）。"""
    return _nfkc(value).casefold() not in _OFF_TOKENS


def _on_text(flag: bool) -> str:
    return "1" if flag else "0"


def _partner_from_path(path: str) -> str:
    """会社フォルダのパスから表示用の取引先名を作る。

    「Z:\\NetMarks以外(常駐）\\アクシスITパートナーズ（細川・渡会）」→
    「アクシスITパートナーズ」。カッコの中は在籍者名なので落とす。
    """
    name = Path(str(path).strip().strip('"')).name
    return re.sub(r"[（(][^）)]*[）)]\s*$", "", name).strip() or name


# ----------------------------------------------------------------------
# CSVの読み書き
# ----------------------------------------------------------------------

def _read_rows(path: Path) -> list[dict[str, str]]:
    """文字コードを判定して読む。BOM付きUTF-8 → UTF-8 → CP932 の順で試す。"""
    if not path.exists():
        return []
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp932"):
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        reader = csv.DictReader(io.StringIO(text))
        return [{(k.lstrip("\ufeff") if k else k): (v or "")
                 for k, v in row.items() if k is not None}
                for row in reader]
    raise InvoiceFoldersError(f"CSVの文字コードを判定できません: {path}")


def signature(path: "str | Path") -> str:
    """中身のハッシュ。NASでは mtime の粒度が粗いので中身で見る。"""
    p = Path(path)
    if not p.exists():
        return ""
    return hashlib.sha256(p.read_bytes()).hexdigest()


def _backup(path: Path) -> "Path | None":
    """書き換え前のバックアップ。設定が壊れると請求書が作れなくなるので必ず取る。"""
    if not path.exists():
        return None
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak_dir = path.parent / "_backup"
    bak_dir.mkdir(parents=True, exist_ok=True)
    bak = bak_dir / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, bak)
    return bak


def _write_rows(path: Path, columns: Sequence[str], rows: Sequence[dict[str, str]]) -> None:
    """BOM付きUTF-8で書く（谷津さんがExcelで開いても文字化けしない）。

    同じフォルダへ一時ファイルを作ってから置き換える。共有フォルダなので、
    書いている途中の半端なCSVを他の人が読む状態を作らない。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(columns), extrasaction="ignore",
                                lineterminator="\r\n")
        writer.writeheader()
        writer.writerows(rows)
    tmp.replace(path)


def _split_known(row: dict[str, str], columns: Sequence[str]) -> tuple[dict[str, str], dict[str, str]]:
    """既知の列と、それ以外（人が手で足した列）に分ける。"""
    known = {c: (row.get(c) or "").strip() for c in columns}
    extra = {k: v for k, v in row.items()
             if k and k not in columns and (v or "").strip()}
    return known, extra


# ----------------------------------------------------------------------
# 読み込み
# ----------------------------------------------------------------------

def load_people(settings_csv: "str | Path") -> list[dict[str, Any]]:
    """PDF作成設定（1人1行）を画面用に読む。"""
    rows: list[dict[str, Any]] = []
    for raw in _read_rows(Path(settings_csv)):
        if not (raw.get("氏名") or "").strip():
            continue
        known, extra = _split_known(raw, PEOPLE_COLUMNS)
        known["対象"] = is_on(raw.get("対象", ""))
        known["_extra"] = extra
        rows.append(known)
    return rows


def load_roots(roots_csv: "str | Path") -> list[dict[str, Any]]:
    """freee取込CSVの探索ルート（1社1行）を画面用に読む。

    CSVが無い場合は invoice_mode の既定28件を出す。画面から保存すれば
    以後はCSVが正になる（コード内の既定は触らない）。
    """
    path = Path(roots_csv)
    rows: list[dict[str, Any]] = []
    if path.exists():
        for raw in _read_rows(path):
            folder = (raw.get("フォルダパス") or raw.get("Path") or raw.get("path") or "")
            folder = str(folder).strip().strip('"')
            if not folder:
                continue
            _, extra = _split_known(raw, ROOT_COLUMNS + ["Path", "path", "有効"])
            rows.append({
                "対象": is_on(raw.get("対象", raw.get("有効", ""))),
                "取引先": (raw.get("取引先") or "").strip() or _partner_from_path(folder),
                "フォルダパス": folder,
                "_extra": extra,
            })
        return rows

    from services.invoice_mode import DEFAULT_TARGET_ROOTS
    return [{"対象": True, "取引先": _partner_from_path(f), "フォルダパス": f, "_extra": {}}
            for f in DEFAULT_TARGET_ROOTS]


def load_all(settings_csv: "str | Path", roots_csv: "str | Path",
             writers_csv: "str | Path") -> dict[str, Any]:
    """画面が最初に読む一式。"""
    writable, write_message = can_write(writers_csv, label="フォルダ設定")
    return {
        "people": load_people(settings_csv),
        "roots": load_roots(roots_csv),
        "settings_csv": str(settings_csv),
        "roots_csv": str(roots_csv),
        "writers_csv": str(writers_csv),
        "user": current_user(),
        "writable": writable,
        "write_message": write_message,
        "signatures": {
            "people": signature(settings_csv),
            "roots": signature(roots_csv),
        },
    }


# ----------------------------------------------------------------------
# 保存
# ----------------------------------------------------------------------

def validate(people: Sequence[dict[str, Any]],
             roots: Sequence[dict[str, Any]]) -> list[str]:
    """保存前の検査。ここで止めないと、あとで別人の勤怠を綴じかねない。"""
    errors: list[str] = []
    seen: set[tuple[str, str]] = set()
    for i, row in enumerate(people, 1):
        partner = _nfkc(row.get("取引先"))
        name = _nfkc(row.get("氏名"))
        if not name:
            errors.append(f"勤怠フォルダ {i}行目: 氏名が空です")
            continue
        if not partner:
            errors.append(f"{name}: 取引先が空です")
        key = (partner.casefold(), name.casefold())
        if key in seen:
            errors.append(f"{partner} {name}: 同じ取引先・氏名の行が2つあります")
        seen.add(key)
        if not bool(row.get("対象", True)):
            continue    # 対象外の行は空欄のままでも保存できる
        for column, label in (("請求書Excel", "請求書Excel"), ("シート名", "シート名"),
                              ("勤怠フォルダ", "勤怠フォルダ"), ("勤怠ファイル", "勤怠ファイル名"),
                              ("出力フォルダ", "出力フォルダ"), ("出力ファイル名", "出力ファイル名")):
            if not (row.get(column) or "").strip():
                errors.append(f"{partner} {name}: {label}が空です")

    root_seen: set[str] = set()
    for i, row in enumerate(roots, 1):
        folder = (row.get("フォルダパス") or "").strip()
        if not folder:
            errors.append(f"請求書格納フォルダ（会社単位）{i}行目: フォルダパスが空です")
            continue
        key = folder.casefold().rstrip("\\/")
        if key in root_seen:
            errors.append(f"{folder}: 同じフォルダパスの行が2つあります")
        root_seen.add(key)
    return errors


def save_all(settings_csv: "str | Path", roots_csv: "str | Path", *,
             people: Sequence[dict[str, Any]], roots: Sequence[dict[str, Any]],
             signatures: dict[str, str] | None = None,
             force: bool = False) -> dict[str, Any]:
    """2本のCSVを書き戻す。

    画面を開いた後に他の人が直していたら、上書きせず中止する（force で強制）。
    書く前に必ずバックアップを取る。
    """
    settings_path, roots_path = Path(settings_csv), Path(roots_csv)
    signatures = signatures or {}
    if not force:
        for label, path, key in (("PDF作成設定", settings_path, "people"),
                                 ("対象フォルダ", roots_path, "roots")):
            before = signatures.get(key)
            if before is not None and signature(path) != before:
                raise InvoiceFoldersConflict(
                    f"{label}のCSVが他の人によって更新されています（{path}）。"
                    "上書きせず中止しました。画面を開き直してから、もう一度直してください。")

    errors = validate(people, roots)
    if errors:
        raise InvoiceFoldersError(" / ".join(errors))

    people_extra: list[str] = []
    people_rows: list[dict[str, str]] = []
    for row in people:
        out = {c: str(row.get(c) or "").strip() for c in PEOPLE_COLUMNS}
        out["対象"] = _on_text(bool(row.get("対象", True)))
        for k, v in (row.get("_extra") or {}).items():
            out[k] = str(v)
            if k not in people_extra:
                people_extra.append(k)
        people_rows.append(out)

    root_extra: list[str] = []
    root_rows: list[dict[str, str]] = []
    for row in roots:
        folder = str(row.get("フォルダパス") or "").strip()
        out = {
            "対象": _on_text(bool(row.get("対象", True))),
            "取引先": str(row.get("取引先") or "").strip() or _partner_from_path(folder),
            "フォルダパス": folder,
        }
        for k, v in (row.get("_extra") or {}).items():
            out[k] = str(v)
            if k not in root_extra:
                root_extra.append(k)
        root_rows.append(out)

    backups: list[str] = []
    for path, columns, rows in ((settings_path, PEOPLE_COLUMNS + people_extra, people_rows),
                                (roots_path, ROOT_COLUMNS + root_extra, root_rows)):
        bak = _backup(path)
        if bak:
            backups.append(str(bak))
        _write_rows(path, columns, rows)

    return {
        "backups": backups,
        "people_saved": len(people_rows),
        "roots_saved": len(root_rows),
        "signatures": {"people": signature(settings_path), "roots": signature(roots_path)},
    }


# ----------------------------------------------------------------------
# 「今月の状態」チェック
# ----------------------------------------------------------------------

def _state(level: str, text: str, detail: str = "") -> dict[str, str]:
    return {"level": level, "text": text, "detail": detail}


def _target_date(month: str) -> dt.date:
    matched = re.fullmatch(r"(\d{4})-(\d{2})", _nfkc(month))
    if not matched:
        raise InvoiceFoldersError("対象月は YYYY-MM 形式で指定してください")
    m = int(matched.group(2))
    if not 1 <= m <= 12:
        raise InvoiceFoldersError("対象月が不正です")
    return dt.date(int(matched.group(1)), m, 1)


def check_kintai(row: dict[str, Any], target: dt.date) -> dict[str, str]:
    """勤怠PDFが1件に決まるか。0件・複数件はどちらも作成が止まる。"""
    folder = Path(expand(row.get("勤怠フォルダ") or "", target))
    pattern = expand(row.get("勤怠ファイル") or "", target)
    if not str(folder):
        return _state(LEVEL_STOP, "未設定")
    if not pattern:
        return _state(LEVEL_STOP, "ファイル名が未設定")
    try:
        if not folder.exists():
            return _state(LEVEL_STOP, "フォルダなし", str(folder))
        found = sorted(p for p in folder.glob(pattern) if p.is_file())
    except OSError as exc:
        return _state(LEVEL_STOP, "フォルダを読めません", str(exc))
    if not found:
        return _state(LEVEL_STOP, "0件", f"{folder}\\{pattern} に合うPDFがありません")
    if len(found) > 1:
        return _state(LEVEL_WARN, f"候補{len(found)}件",
                      "どれを綴じるか決められないので作成は止まります: "
                      + " / ".join(p.name for p in found))
    return _state(LEVEL_OK, "1件", found[0].name)


def check_output(row: dict[str, Any], target: dt.date) -> dict[str, str]:
    """出力先フォルダがあるか、同名PDFが既にないか。"""
    folder = Path(expand(row.get("出力フォルダ") or "", target))
    name = expand(row.get("出力ファイル名") or "", target)
    if not name:
        return _state(LEVEL_STOP, "ファイル名が未設定")
    try:
        if not folder.exists():
            return _state(LEVEL_STOP, "フォルダなし", str(folder))
        if (folder / name).exists():
            return _state(LEVEL_WARN, "同名あり", f"{name} が既にあります（上書きしません）")
    except OSError as exc:
        return _state(LEVEL_STOP, "フォルダを読めません", str(exc))
    return _state(LEVEL_OK, "作れます", str(folder / name))


def check_root(row: dict[str, Any]) -> dict[str, str]:
    """会社フォルダがあるか、配下に「提出データ」があるか。

    請求書PDFが何件見つかるかまでは見ない。それは②の「PDFを検索して確認」が
    28フォルダを再帰で掘って出すもので、ここで二重にやると保存前のチェックが
    毎回数十秒かかる。ここは打ち間違いを見つけるための軽い確認にとどめる。
    """
    folder = Path((row.get("フォルダパス") or "").strip())
    if not str(folder) or str(folder) == ".":
        return _state(LEVEL_STOP, "未設定")
    try:
        if not folder.exists():
            return _state(LEVEL_STOP, "フォルダなし", str(folder))
        bases = [c for c in folder.iterdir()
                 if c.is_dir() and "提出データ" in _nfkc(c.name)]
    except OSError as exc:
        return _state(LEVEL_STOP, "フォルダを読めません", str(exc))
    if not bases:
        return _state(LEVEL_INFO, "提出データなし",
                      "直下に「提出データ」フォルダが無いので、会社フォルダ全体を探します")
    return _state(LEVEL_OK, f"提出データ{len(bases)}個",
                  " / ".join(b.name for b in bases))


class _SheetReader:
    """請求書Excelのシート名を読む。Excelが使えない環境では黙って諦める。

    .xls があるため openpyxl では読めず、Excel COM が要る。共有exeを普段の
    PCで動かす前提だが、チェックのために画面が落ちるのは困るので、
    失敗したら「ファイルはある・シートは未確認」に倒す。
    """

    def __init__(self) -> None:
        self._excel = None
        self._com_ready = False
        self._failed = False

    def __enter__(self) -> "_SheetReader":
        return self

    def __exit__(self, *exc_info) -> None:
        self.close()

    def _ensure(self):
        if self._excel is not None or self._failed:
            return self._excel
        try:
            import pythoncom  # type: ignore
            import win32com.client as win32  # type: ignore
            pythoncom.CoInitialize()
            self._com_ready = True
            excel = win32.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            self._excel = excel
        except Exception:
            logger.info("Excelを起動できないためシート名の確認を省きます", exc_info=True)
            self._failed = True
        return self._excel

    def sheet_names(self, workbook: Path) -> "list[str] | None":
        excel = self._ensure()
        if excel is None:
            return None
        book = None
        try:
            book = excel.Workbooks.Open(str(workbook), False, True)   # 読み取り専用
            return [s.Name for s in book.Worksheets]
        except Exception:
            logger.info("シート名を読めません: %s", workbook, exc_info=True)
            return None
        finally:
            if book is not None:
                try:
                    book.Close(False)
                except Exception:
                    pass

    def close(self) -> None:
        if self._excel is not None:
            try:
                self._excel.Quit()
            except Exception:
                pass
            self._excel = None
        if self._com_ready:
            try:
                import pythoncom  # type: ignore
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_ready = False


def check_excel(row: dict[str, Any], target: dt.date,
                reader: "_SheetReader | None" = None) -> dict[str, str]:
    """請求書Excelがあるか、対象月のシートがあるか。

    前月シートをコピーして作る運用なので「今月のシートをまだ作っていない」が
    実際に起きる。作成の直前ではなく、ここで気づけるようにする。
    """
    workbook = Path(expand(row.get("請求書Excel") or "", target))
    sheet = expand(row.get("シート名") or "", target)
    if not workbook.name:
        return _state(LEVEL_STOP, "未設定")
    try:
        if not workbook.exists():
            return _state(LEVEL_STOP, "ファイルなし", str(workbook))
    except OSError as exc:
        return _state(LEVEL_STOP, "ファイルを読めません", str(exc))
    if not sheet:
        return _state(LEVEL_STOP, "シート名が未設定")
    if reader is None:
        return _state(LEVEL_INFO, "ファイルあり", f"シート「{sheet}」は未確認")
    names = reader.sheet_names(workbook)
    if names is None:
        return _state(LEVEL_INFO, "ファイルあり", f"シート「{sheet}」は未確認")
    # シート名の末尾に空白が入っているブックがある（'26年7月 '）ので詰めて比べる。
    if any(n.strip() == sheet.strip() for n in names):
        return _state(LEVEL_OK, "シートあり", sheet)
    return _state(LEVEL_WARN, "シートなし",
                  f"「{sheet}」がありません（あるのは {', '.join(names)}）")


def check(month: str, *, people: Sequence[dict[str, Any]],
          roots: Sequence[dict[str, Any]], scope: str = "all",
          read_sheets: bool = True) -> dict[str, Any]:
    """対象月について、各行が今どうなっているかを返す。

    scope で見るものを絞る。画面はタブごとに呼ぶので、Excelを起動するのは
    「請求書作成Excel」タブを開いたときだけになる。
    """
    target = _target_date(month)
    want = lambda key: scope in ("all", key)          # noqa: E731
    result: dict[str, Any] = {"month": month, "scope": scope,
                              "people": [], "roots": []}

    reader: "_SheetReader | None" = None
    try:
        if read_sheets and want("excel"):
            reader = _SheetReader()
        for index, row in enumerate(people):
            entry: dict[str, Any] = {"index": index}
            if not bool(row.get("対象", True)):
                off = _state(LEVEL_OFF, "対象外", "チェックを外しているので作成しません")
                for key in ("kintai", "excel", "output"):
                    if want(key):
                        entry[key] = off
            else:
                if want("kintai"):
                    entry["kintai"] = check_kintai(row, target)
                if want("excel"):
                    entry["excel"] = check_excel(row, target, reader)
                if want("output"):
                    entry["output"] = check_output(row, target)
            result["people"].append(entry)
    finally:
        if reader is not None:
            reader.close()

    if want("output"):
        for index, row in enumerate(roots):
            if not bool(row.get("対象", True)):
                state = _state(LEVEL_OFF, "対象外", "チェックを外しているので探しません")
            else:
                state = check_root(row)
            result["roots"].append({"index": index, "root": state})
    return result


def summarize(states: Iterable[dict[str, Any]]) -> dict[str, int]:
    """バッジの数を数える（画面の見出しに出す用）。"""
    counts = {LEVEL_OK: 0, LEVEL_WARN: 0, LEVEL_STOP: 0, LEVEL_INFO: 0, LEVEL_OFF: 0}
    for state in states:
        level = (state or {}).get("level")
        if level in counts:
            counts[level] += 1
    return counts
