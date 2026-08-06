# -*- coding: utf-8 -*-
"""SAP経費 取込済み費用シート台帳（2026-08-06 谷津さん決定）

SAP Fieldglass の月次経費CSVは **前々月分が再掲されて落ちてくる**。従来は「前月の生CSV」を
過去データに指定して費用シートIDで突合していたが、この方式は 2026-08-06 に破綻した:

  - 実際に取り込んだのは 7月\\経費確認精査データ\\SAP経費\\...(4)_重複除外済.csv
  - しかし旧ツールはファイル名に「重複除外済」を含むCSVを過去データから自動スキップする
  - SAP元データ フォルダに入っていた (5).csv は **7月に使わなかった別ダウンロード**で、
    22件の費用シートIDが欠落していた
  → 過去に (5) を渡すと除外0行になり、285,642円（13名/93行）が二重支給になりかけた

そこで「どのファイルで取り込んだか」を人間が思い出す必要をなくし、**取り込んだ費用シートを
台帳に記録して、翌月はファイルではなく台帳と突合する**方式へ移行した。

判定は3段階（2026-08-06 谷津さん確認）:

  1. 台帳と **明細まで完全一致** → 自動除外（確実に支給済み）
  2. **費用シートIDは一致するが明細が違う** → 要確認（自動除外も自動取込もしない）
  3. 該当なし → そのまま取込

2 を作ったのは、SAPが同じ費用シートIDのまま **金額を訂正して再発行する**ため。
2026-07 の実データでは、6月ファイルで全行5,000円（仮の値）だった費用シートが
7月ファイルで230円・242円…の実額に直っていた例、当番手当が税抜2,500円⇄税込2,750円で
付け替わっていた例が27件あった。ID単位だけで弾くと「遅れて申請された前々月分」を
落として未払いになり、明細単位だけで弾くと訂正再掲を二重計上する。両方を分けて出す。

台帳の状態は 暫定 / 確定 の2段階。**判定に使うのは確定分のみ**。
除外を実行した時点では暫定で記録し、jinjerへ取り込んだ後に確定させる。
暫定のまま残っている月があれば呼び出し側で警告する（confirm 忘れの検知）。

書き込みは許可ユーザーのみ（既定: 谷津さん・平良さん）。共有exeを使う5名全員が
台帳を書き換えられる状態を避けるため。許可リストは外部CSVなので exe 再ビルド不要で追加できる。
"""

from __future__ import annotations

import csv
import datetime as dt
import getpass
import hashlib
import logging
import os
import shutil
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

logger = logging.getLogger(__name__)

ID_COLUMN = "費用シート ID"

# 台帳の列。1行＝SAP生CSVの1明細。
LEDGER_COLUMNS = [
    "状態",             # 暫定 / 確定
    "取込年月",          # YYYY-MM（給与計算の月。例: 8月給与で7月経費を精算 → 2026-08）
    ID_COLUMN,
    "姓", "名",
    "費用エントリ日",     # YYYY/M/D（元CSVの表記のまま保持）
    "費用合計",
    "業者名",
    "説明",
    "元CSV",            # 取込に使ったCSVのファイル名（証跡）
    "記録日時", "記録者",
    "確定日時", "確定者",
]

STATUS_PROVISIONAL = "暫定"
STATUS_CONFIRMED = "確定"

# 明細の同一性を決める列。金額訂正の再発行を「別物」と判定させるため金額まで含める。
_ROW_KEY_COLUMNS = (ID_COLUMN, "姓", "名", "費用エントリ日", "費用合計", "業者名", "説明")


# ----------------------------------------------------------------------
# 正規化
# ----------------------------------------------------------------------

def _norm(value) -> str:
    """前後空白と全角空白を落とした文字列。"""
    return str(value if value is not None else "").strip().replace("　", " ").strip()


def _norm_date(value) -> str:
    """2026/6/1 も 2026/06/01 も同じキーになるよう YYYY/M/D へ揃える。"""
    s = _norm(value).replace("-", "/")
    parts = [p for p in s.split("/") if p]
    if len(parts) == 3:
        try:
            return f"{int(parts[0])}/{int(parts[1])}/{int(parts[2])}"
        except ValueError:
            pass
    return s


def _norm_amount(value) -> str:
    """"2,096" も "2096" も "2096.0" も同じキーになるよう整数文字列へ揃える。"""
    s = _norm(value).replace(",", "")
    if not s:
        return ""
    try:
        return str(int(round(float(s))))
    except ValueError:
        return s


def sheet_id_of(row: dict, id_column: str = ID_COLUMN) -> str:
    """費用シートIDの照合キー。空IDは "" を返し、判定に使わない。"""
    return _norm(row.get(id_column, ""))


def row_key(row: dict, id_column: str = ID_COLUMN) -> tuple:
    """明細1行の同一性キー（費用シートID＋氏名＋日付＋金額＋業者名＋説明）。

    山田さんの実例のように、同じ費用シートID・同じ日・同じ金額でも業者名と説明だけが
    違う明細が並ぶことがあるため、業者名・説明まで含めないと別明細を潰してしまう。
    """
    return (
        sheet_id_of(row, id_column),
        _norm(row.get("姓", "")),
        _norm(row.get("名", "")),
        _norm_date(row.get("費用エントリ日", "")),
        _norm_amount(row.get("費用合計", "")),
        _norm(row.get("業者名", "")),
        _norm(row.get("説明", "")),
    )


# ----------------------------------------------------------------------
# 台帳の読み書き
# ----------------------------------------------------------------------

def _signature(path: Path) -> str:
    """台帳ファイルの内容シグネチャ（同時書き込み検知用）。無ければ空文字。

    NAS では mtime の粒度が粗く当てにならないので中身のハッシュを使う。
    """
    if not path.exists():
        return ""
    return hashlib.sha256(path.read_bytes()).hexdigest()


@dataclass
class Ledger:
    path: Path
    rows: list[dict] = field(default_factory=list)
    # 読み込んだ時点のシグネチャ。保存直前に読み直して変わっていたら、
    # 自分が読んでから誰かが台帳を更新したということなので上書きしない。
    loaded_signature: str = ""

    @property
    def confirmed_rows(self) -> list[dict]:
        return [r for r in self.rows if _norm(r.get("状態")) == STATUS_CONFIRMED]

    @property
    def provisional_months(self) -> list[str]:
        """暫定のまま残っている取込年月（新しい順）。確定忘れの検知に使う。"""
        months = {_norm(r.get("取込年月")) for r in self.rows
                  if _norm(r.get("状態")) == STATUS_PROVISIONAL}
        return sorted((m for m in months if m), reverse=True)

    def months(self) -> list[str]:
        return sorted({_norm(r.get("取込年月")) for r in self.rows if _norm(r.get("取込年月"))})

    def confirmed_row_counter(self) -> Counter:
        """確定分の明細キー → 件数。同一明細が複数あるケースを取りこぼさないため件数で持つ。"""
        return Counter(row_key(r) for r in self.confirmed_rows)

    def confirmed_sheet_ids(self) -> set[str]:
        return {sheet_id_of(r) for r in self.confirmed_rows} - {""}

    def sheet_id_months(self) -> dict[str, str]:
        """費用シートID → 最後に確定した取込年月（要確認リストの表示用）。"""
        out: dict[str, str] = {}
        for r in self.confirmed_rows:
            sid = sheet_id_of(r)
            if sid:
                out[sid] = max(out.get(sid, ""), _norm(r.get("取込年月")))
        return out


def load_ledger(path: "str | Path") -> Ledger:
    """台帳CSVを読む。無ければ空の台帳を返す（初回運用のため例外にしない）。"""
    p = Path(path).expanduser()
    if not p.exists():
        logger.info("SAP台帳がまだありません（新規作成されます）: %s", p)
        return Ledger(path=p, rows=[], loaded_signature="")
    data = p.read_bytes()
    encoding = "utf-8-sig" if data.startswith(b"\xef\xbb\xbf") else "utf-8"
    try:
        data.decode(encoding)
    except UnicodeDecodeError:
        encoding = "cp932"
    with p.open("r", encoding=encoding, newline="") as f:
        reader = csv.DictReader(f)
        fieldnames = [(n.lstrip("﻿") if n else n) for n in (reader.fieldnames or [])]
        rows = [{(k.lstrip("﻿") if k else k): (v or "") for k, v in r.items() if k is not None}
                for r in reader]
    # 壊れた/別物のファイルを「空の台帳」として黙って受け入れると、除外が一切効かなくなり
    # 二重支給に直結する。必須列が無ければ読めなかったものとして落とす。
    required = {"状態", "取込年月", ID_COLUMN}
    missing = required - set(fieldnames)
    if missing:
        raise ValueError(
            f"台帳CSVの形式が不正です（列 {'・'.join(sorted(missing))} がありません）: {p}")
    return Ledger(path=p, rows=rows,
                  loaded_signature=hashlib.sha256(data).hexdigest())


def _backup(path: Path) -> "Path | None":
    """書き換え前のバックアップ。台帳が壊れると給与が狂うので必ず取る。"""
    if not path.exists():
        return None
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak_dir = path.parent / "_台帳バックアップ"
    bak_dir.mkdir(parents=True, exist_ok=True)
    bak = bak_dir / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, bak)
    return bak


class LedgerConflictError(RuntimeError):
    """自分が読み込んでから他の人が台帳を更新していた（＝上書きすると相手の変更が消える）。"""


def save_ledger(ledger: Ledger, force: bool = False) -> "Path | None":
    """台帳CSVを書き出す（BOM付きUTF-8＝谷津さんのExcelで文字化けしない）。戻り値はバックアップ先。

    書き込みは「読む→直す→書く」なので、谷津さんと平良さんの操作が数秒重なると
    後から書いた方が相手の変更を消す。たとえば片方が「2026-07を確定」した直後に
    もう片方が経費統合を実行すると**確定が消え、翌月の除外が効かなくなる**。
    そこで保存の直前にファイルを読み直し、読み込んだ時点から変わっていたら中止する。

    force=True で強制上書き（相手の変更は失われる。通常は使わない）。
    """
    if not force and _signature(ledger.path) != ledger.loaded_signature:
        raise LedgerConflictError(
            f"台帳が他の人によって更新されています（{ledger.path}）。"
            "あなたが読み込んだ後に変更されたため、上書きせず中止しました。"
            "画面を開き直して、もう一度実行してください。"
        )
    bak = _backup(ledger.path)
    ledger.path.parent.mkdir(parents=True, exist_ok=True)
    with ledger.path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=LEDGER_COLUMNS, extrasaction="ignore",
                           lineterminator="\r\n")
        w.writeheader()
        w.writerows(ledger.rows)
    ledger.loaded_signature = _signature(ledger.path)  # 連続保存できるよう更新
    return bak


# ----------------------------------------------------------------------
# 書き込み許可ユーザー
# ----------------------------------------------------------------------

def current_user() -> str:
    """Windows のログオンユーザー名。"""
    return os.environ.get("USERNAME") or os.environ.get("USER") or getpass.getuser()


def load_writers(path: "str | Path") -> list[dict]:
    """書き込み許可ユーザーCSV（列: ユーザー名, 表示名, 備考）を読む。"""
    p = Path(path).expanduser()
    if not p.exists():
        return []
    data = p.read_bytes()
    encoding = "utf-8-sig" if data.startswith(b"\xef\xbb\xbf") else "utf-8"
    try:
        data.decode(encoding)
    except UnicodeDecodeError:
        encoding = "cp932"
    with p.open("r", encoding=encoding, newline="") as f:
        return [{(k.lstrip("﻿") if k else k): (v or "") for k, v in r.items() if k is not None}
                for r in csv.DictReader(f)]


def can_write(writers_csv: "str | Path", user: "str | None" = None) -> tuple[bool, str]:
    """(書き込み可否, 画面に出す理由) を返す。

    許可リストが読めない場合は **書き込み不可** に倒す。台帳は給与に直結するため、
    設定ミスで全員が書ける状態になる方が危ない。
    """
    who = user or current_user()
    rows = load_writers(writers_csv)
    if not rows:
        return False, (f"書き込み許可リストが読めません（{writers_csv}）。"
                       f"現在のユーザー: {who}")
    allowed = {_norm(r.get("ユーザー名", "")).lower() for r in rows} - {""}
    if _norm(who).lower() in allowed:
        return True, f"書き込み可（ユーザー: {who}）"
    names = "・".join(_norm(r.get("表示名") or r.get("ユーザー名")) for r in rows)
    return False, (f"台帳への書き込みは {names} のみです。現在のユーザー: {who}"
                   f"（追加する場合は {writers_csv} に1行足してください）")


# ----------------------------------------------------------------------
# 3段階判定
# ----------------------------------------------------------------------

@dataclass
class LedgerPlan:
    """当月SAP生CSVを台帳と突合した結果。"""
    kept_rows: list[dict] = field(default_factory=list)      # 取込対象
    excluded_rows: list[dict] = field(default_factory=list)  # 明細まで一致＝支給済み
    review_rows: list[dict] = field(default_factory=list)    # IDのみ一致＝要確認
    ledger_month_count: int = 0
    ledger_row_count: int = 0

    @property
    def current_row_count(self) -> int:
        return len(self.kept_rows) + len(self.excluded_rows) + len(self.review_rows)

    def summary_dict(self) -> dict:
        return {
            "current_rows": self.current_row_count,
            "kept_rows": len(self.kept_rows),
            "excluded_rows": len(self.excluded_rows),
            "review_rows": len(self.review_rows),
            "ledger_rows": self.ledger_row_count,
            "ledger_months": self.ledger_month_count,
        }


def classify_rows(current_rows: Iterable[dict], ledger: Ledger,
                  id_column: str = ID_COLUMN) -> LedgerPlan:
    """当月の各行を「取込 / 除外 / 要確認」の3つに振り分ける。

    - 明細キーが確定台帳にある（件数まで見る） → 除外
    - 費用シートIDだけ確定台帳にある            → 要確認
    - どちらも無い                              → 取込
    """
    remaining = ledger.confirmed_row_counter()
    known_ids = ledger.confirmed_sheet_ids()
    id_months = ledger.sheet_id_months()

    plan = LedgerPlan(
        ledger_month_count=len(ledger.months()),
        ledger_row_count=len(ledger.confirmed_rows),
    )
    for row in current_rows:
        sid = sheet_id_of(row, id_column)
        key = row_key(row, id_column)
        if sid and remaining.get(key, 0) > 0:
            remaining[key] -= 1
            out = dict(row)
            out["判定"] = "除外（支給済み）"
            out["理由"] = "台帳と明細まで一致"
            out["台帳の取込年月"] = id_months.get(sid, "")
            plan.excluded_rows.append(out)
        elif sid and sid in known_ids:
            out = dict(row)
            out["判定"] = "要確認"
            out["理由"] = "費用シートIDは台帳にあるが明細が違う（金額訂正の再発行の可能性）"
            out["台帳の取込年月"] = id_months.get(sid, "")
            plan.review_rows.append(out)
        else:
            plan.kept_rows.append(dict(row))
    return plan


REVIEW_EXTRA_COLUMNS = ["判定", "理由", "台帳の取込年月"]


# ----------------------------------------------------------------------
# 記録（暫定）と確定
# ----------------------------------------------------------------------

def append_provisional(ledger: Ledger, rows: Iterable[dict], import_month: str,
                       source_csv: "str | Path", user: "str | None" = None,
                       id_column: str = ID_COLUMN) -> int:
    """取込対象の明細を「暫定」で台帳に追加する。同じ取込年月の暫定は入れ替える。

    やり直し（同じ月の再実行）で行が二重に積まれないよう、**同一取込年月の暫定行は
    先に取り除いてから**追加する。確定済みの行には触れない。
    戻り値は追加した行数。
    """
    month = _norm(import_month)
    who = user or current_user()
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    src = Path(source_csv).name if source_csv else ""

    ledger.rows = [r for r in ledger.rows
                   if not (_norm(r.get("取込年月")) == month
                           and _norm(r.get("状態")) == STATUS_PROVISIONAL)]
    added = 0
    for row in rows:
        if not sheet_id_of(row, id_column):
            continue  # 費用シートIDが無い行は判定に使えないので台帳にも入れない
        ledger.rows.append({
            "状態": STATUS_PROVISIONAL,
            "取込年月": month,
            ID_COLUMN: sheet_id_of(row, id_column),
            "姓": _norm(row.get("姓", "")),
            "名": _norm(row.get("名", "")),
            "費用エントリ日": _norm(row.get("費用エントリ日", "")),
            "費用合計": _norm(row.get("費用合計", "")),
            "業者名": _norm(row.get("業者名", "")),
            "説明": _norm(row.get("説明", "")),
            "元CSV": src,
            "記録日時": now,
            "記録者": who,
            "確定日時": "",
            "確定者": "",
        })
        added += 1
    return added


def confirm_month(ledger: Ledger, import_month: str, user: "str | None" = None) -> int:
    """指定した取込年月の暫定行を「確定」にする。戻り値は確定した行数。"""
    month = _norm(import_month)
    who = user or current_user()
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    n = 0
    for r in ledger.rows:
        if _norm(r.get("取込年月")) == month and _norm(r.get("状態")) == STATUS_PROVISIONAL:
            r["状態"] = STATUS_CONFIRMED
            r["確定日時"] = now
            r["確定者"] = who
            n += 1
    return n


def unconfirm_month(ledger: Ledger, import_month: str) -> int:
    """確定を暫定に戻す（取込をやり直すとき用）。戻り値は戻した行数。"""
    month = _norm(import_month)
    n = 0
    for r in ledger.rows:
        if _norm(r.get("取込年月")) == month and _norm(r.get("状態")) == STATUS_CONFIRMED:
            r["状態"] = STATUS_PROVISIONAL
            r["確定日時"] = ""
            r["確定者"] = ""
            n += 1
    return n


def default_import_month(today: "dt.date | None" = None) -> str:
    """既定の取込年月＝実行日の年月（8月に動かす＝8月給与＝7月経費の精算）。"""
    d = today or dt.date.today()
    return f"{d.year}-{d.month:02d}"
