# -*- coding: utf-8 -*-
"""jinjer給与へ投入した経費インポートの内容を月ごとに保持する台帳。

**なぜ要るか**（2026-08-10 谷津さん依頼）:
投入したあとにイレギュラー経費が出てくると、統合一覧表を4ソースから作り直し、
経路突合レビューもやり直してからでないと入れ直せず、とても大変だった。
投入した中身を持っておけば「前回分＋追加分」を組み直して再投入するだけで済む。

**方式は上書き**（2026-08-10 谷津さん決定）:
差分だけを送るのではなく、前回分と追加分を合わせた**全行**を作り直して再投入する。
jinjer は同じ月・同じ社員を再投入すると加算ではなく置き換えになる（実績から確認済み）。
差分計算が要らないぶん、二重計上の事故が起きにくい。

**ファイルポインタではなく行データを持つ理由**:
インポートCSVの既定名は固定（経費統合一覧表_jinjerインポート.csv）で、次に統合を実行すると
上書きされてしまう。「どのファイルを投入したか」を覚えても中身は後から変わる。

同時書き込みの守り方は SAP台帳（services/sap_import_ledger.py）と同じ。
NAS では mtime が当てにならないので中身の SHA-256 で比べる。
"""
from __future__ import annotations

import csv
import datetime as dt
import getpass
import hashlib
import io
import logging
import os
import shutil
from dataclasses import dataclass, field
from pathlib import Path

from services.keihi_payroll_import import ITEM_KEYS, _match_column

logger = logging.getLogger(__name__)

# 台帳の列。1行＝1社員の投入内容（その月に投入した全項目）
BASE_COLUMNS = ["対象月", "社員番号", "氏名"]
META_COLUMNS = ["元CSV", "テンプレID", "投入結果", "投入日時", "投入者"]
LEDGER_COLUMNS = BASE_COLUMNS + list(ITEM_KEYS) + META_COLUMNS

# 投入結果。jinjer の status: "1"=全成功 / "0"=タイムアウト（画面で要確認）
RESULT_OK = "成功"
RESULT_TIMEOUT = "タイムアウト（jinjer画面で要確認）"
_STATUS_TO_RESULT = {"1": RESULT_OK, "0": RESULT_TIMEOUT}


def _norm(v) -> str:
    return ("" if v is None else str(v)).strip()


def _to_int(v) -> int:
    s = _norm(v).replace(",", "")
    if not s:
        return 0
    try:
        return int(round(float(s)))
    except ValueError:
        return 0


def current_user() -> str:
    """Windows のログオンユーザー名。"""
    return os.environ.get("USERNAME") or os.environ.get("USER") or getpass.getuser()


def _signature(path: Path) -> str:
    """台帳ファイルの内容シグネチャ（同時書き込み検知用）。無ければ空文字。"""
    if not path.exists():
        return ""
    return hashlib.sha256(path.read_bytes()).hexdigest()


class LedgerConflictError(RuntimeError):
    """自分が読み込んでから他の人が台帳を更新していた（＝上書きすると相手の記録が消える）。"""


@dataclass
class ImportLedger:
    path: Path
    rows: list = field(default_factory=list)
    # 読み込んだ時点のシグネチャ。保存直前に読み直して変わっていたら上書きしない。
    loaded_signature: str = ""

    def months(self) -> list:
        return sorted({_norm(r.get("対象月")) for r in self.rows if _norm(r.get("対象月"))},
                      reverse=True)


def month_rows(ledger: ImportLedger, month: str) -> list:
    """その月に投入した行だけを返す。"""
    m = _norm(month)
    return [r for r in ledger.rows if _norm(r.get("対象月")) == m]


def months_summary(ledger: ImportLedger) -> list:
    """画面に出す月ごとの状態（新しい順）。"""
    out = []
    for m in ledger.months():
        rows = month_rows(ledger, m)
        stamps = sorted(_norm(r.get("投入日時")) for r in rows)
        results = {_norm(r.get("投入結果")) for r in rows}
        out.append({
            "month": m,
            "rows": len(rows),
            "last_at": stamps[-1] if stamps else "",
            "last_result": "／".join(sorted(x for x in results if x)),
            "last_csv": _norm(rows[-1].get("元CSV")) if rows else "",
        })
    return out


def load_ledger(path: "str | Path") -> ImportLedger:
    """台帳CSVを読む。無ければ空の台帳を返す（初回運用のため例外にしない）。"""
    p = Path(path).expanduser()
    if not p.exists():
        logger.info("経費インポート投入台帳がまだありません（新規作成されます）: %s", p)
        return ImportLedger(path=p, rows=[], loaded_signature="")
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
    # 壊れた/別物のファイルを「空の台帳」として黙って受け入れると、追加投入の土台が
    # 空になり前回分が消えたまま投入されてしまう。必須列が無ければ読めなかったものとして落とす。
    missing = {"対象月", "社員番号"} - set(fieldnames)
    if missing:
        raise ValueError(
            f"投入台帳CSVの形式が不正です（列 {'・'.join(sorted(missing))} がありません）: {p}")
    return ImportLedger(path=p, rows=rows,
                        loaded_signature=hashlib.sha256(data).hexdigest())


def _backup(path: Path) -> "Path | None":
    """書き換え前のバックアップ。台帳が狂うと再投入の中身が狂うので必ず取る。"""
    if not path.exists():
        return None
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak_dir = path.parent / "_台帳バックアップ"
    bak_dir.mkdir(parents=True, exist_ok=True)
    bak = bak_dir / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, bak)
    return bak


def save_ledger(ledger: ImportLedger, force: bool = False) -> "Path | None":
    """台帳CSVを書き出す（BOM付きUTF-8＝Excelで文字化けしない）。戻り値はバックアップ先。

    保存の直前にファイルを読み直し、読み込んだ時点から変わっていたら中止する。
    force=True で強制上書き（相手の記録は失われる。通常は使わない）。
    """
    if not force and _signature(ledger.path) != ledger.loaded_signature:
        raise LedgerConflictError(
            f"投入台帳が他の人によって更新されています（{ledger.path}）。"
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
# 投入済みCSVの読み取り
# ----------------------------------------------------------------------

def parse_import_csv(csv_bytes: bytes) -> tuple[list, list]:
    """投入したインポートCSVを {社員番号, 氏名, 各項目: 金額} の行に戻す。(行, 警告)。

    列の並びは jinjer のテンプレートに合わせて可変なので、見出しを
    keihi_payroll_import._match_column で正規のキーへ引き直す（未知の列は無視）。
    """
    warnings: list = []
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            text = csv_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        return [], ["投入CSVの文字コードを判別できませんでした（台帳には記録していません）"]

    table = list(csv.reader(io.StringIO(text, newline="")))
    if not table:
        return [], ["投入CSVが空でした（台帳には記録していません）"]

    header = table[0]
    keys = [_match_column(c) for c in header]
    if "社員番号" not in keys:
        return [], ["投入CSVに社員番号の列が見つかりませんでした（台帳には記録していません）"]

    rows: list = []
    for raw in table[1:]:
        if not any(_norm(c) for c in raw):
            continue
        rec: dict = {}
        for i, key in enumerate(keys):
            if not key or i >= len(raw):
                continue
            rec[key] = _norm(raw[i]) if key in ("社員番号", "氏名") else _to_int(raw[i])
        if not rec.get("社員番号"):
            continue
        rows.append({"社員番号": rec["社員番号"], "氏名": rec.get("氏名", ""),
                     **{k: rec.get(k, 0) for k in ITEM_KEYS}})
    if not rows:
        warnings.append("投入CSVに明細行がありませんでした（台帳には記録していません）")
    return rows, warnings


def replace_month(ledger: ImportLedger, month: str, rows: list, source_csv: str,
                  template_id: str, status: str, user: "str | None" = None) -> int:
    """その月の記録を丸ごと入れ替える。戻り値は記録した行数。

    投入は毎回「全行を作り直して上書き」なので、台帳も差分ではなく全入替にする。
    こうしておけば台帳の中身は常に jinjer の最新状態と一致する。
    """
    m = _norm(month)
    who = user or current_user()
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result = _STATUS_TO_RESULT.get(_norm(status), _norm(status))
    ledger.rows = [r for r in ledger.rows if _norm(r.get("対象月")) != m]
    for r in rows:
        rec = {"対象月": m, "社員番号": _norm(r.get("社員番号")), "氏名": _norm(r.get("氏名"))}
        for k in ITEM_KEYS:
            rec[k] = _to_int(r.get(k))
        rec.update({"元CSV": source_csv, "テンプレID": _norm(template_id),
                    "投入結果": result, "投入日時": stamp, "投入者": who})
        ledger.rows.append(rec)
    return len(rows)


def record_submission(path: "str | Path", month: str, csv_bytes: bytes, file_name: str,
                      template_id: str, status: str,
                      user: "str | None" = None) -> tuple[bool, list]:
    """投入が済んだあとに台帳へ記録する。(記録できたか, 警告) を返す。

    **例外を外に出さない**。投入自体はもう終わっているので、記録に失敗しても
    画面上は失敗扱いにしない。ただし黙って諦めると「台帳に入ったつもり」で
    翌週の追加投入が前回分を落としてしまうので、必ず警告として返す。
    """
    rows, warnings = parse_import_csv(csv_bytes)
    if not rows:
        return False, warnings
    for attempt in (1, 2):
        try:
            ledger = load_ledger(path)
            replace_month(ledger, month, rows, file_name, template_id, status, user)
            save_ledger(ledger)
            return True, warnings
        except LedgerConflictError:
            if attempt == 1:
                continue    # 誰かが同時に書いた。読み直してもう一度だけ試す
            warnings.append(
                "投入台帳が同時に更新されていたため記録できませんでした。"
                "追加投入を使う前に、もう一度この月を投入し直してください。")
            return False, warnings
        except Exception as e:  # noqa: BLE001 — 記録の失敗で投入結果を失敗にはしない
            logger.exception("投入台帳への記録に失敗")
            warnings.append(f"投入台帳へ記録できませんでした（{e}）。追加投入は使えません。")
            return False, warnings
    return False, warnings


# ----------------------------------------------------------------------
# 追加投入（前回分＋追加分）
# ----------------------------------------------------------------------

def merge_addon(base_rows: list, manual_items: dict) -> tuple[list, list, list]:
    """前回の投入行に追加分を足す。(投入用の行, 画面用のプレビュー行, 警告)。

    manual_items は {項目: {社員番号: 金額}}（イレギュラー経費の入力と同じ形）。
    同じ社員・同じ項目は加算する。マイナスを入れれば取り消しになる。
    """
    manual_items = manual_items or {}
    warnings: list = []
    merged: dict = {}
    order: list = []
    for r in base_rows:
        emp = _norm(r.get("社員番号"))
        if not emp:
            continue
        if emp not in merged:
            order.append(emp)
            merged[emp] = {"社員番号": emp, "氏名": _norm(r.get("氏名")),
                           **{k: 0 for k in ITEM_KEYS}}
        for k in ITEM_KEYS:
            merged[emp][k] += _to_int(r.get(k))

    added: dict = {}
    for item, per_emp in manual_items.items():
        if item not in ITEM_KEYS:
            warnings.append(f"⚠️ 「{item}」は投入できる項目ではないため無視しました")
            continue
        for emp, amount in (per_emp or {}).items():
            emp = _norm(emp)
            if not emp:
                continue
            if emp not in merged:
                order.append(emp)
                merged[emp] = {"社員番号": emp, "氏名": "", **{k: 0 for k in ITEM_KEYS}}
            merged[emp][item] += _to_int(amount)
            added.setdefault(emp, {})[item] = added.setdefault(emp, {}).get(item, 0) + _to_int(amount)

    base_ids = {_norm(r.get("社員番号")) for r in base_rows}
    rows, preview = [], []
    for emp in order:
        rec = merged[emp]
        rows.append(rec)
        add = added.get(emp) or {}
        if emp not in base_ids:
            kubun = "新規"
        elif add:
            kubun = "追加あり"
        else:
            kubun = "前回のみ"
        preview.append({**rec, "区分": kubun,
                        "追加内容": "、".join(f"{k} {v:+,}円" for k, v in add.items())})
    return rows, preview, warnings


def resolve_names(ledger: ImportLedger, emp_ids, roster_fetch=None) -> tuple[dict, list]:
    """社員番号→氏名を埋める。台帳の全月を先に見て、足りないときだけ roster を1回取る。

    jinjer API のレート制限（429）はテナント単位なので、要らない取得はしない。
    取れなくても投入は続けられる（氏名は表示用で、jinjer 側は社員番号で紐づく）。
    """
    warnings: list = []
    names: dict = {}
    for r in ledger.rows:
        emp, name = _norm(r.get("社員番号")), _norm(r.get("氏名"))
        if emp and name:
            names.setdefault(emp, name)
    unknown = [e for e in emp_ids if _norm(e) and not names.get(_norm(e))]
    if unknown and roster_fetch is not None:
        try:
            fetched = roster_fetch() or {}
            for emp in unknown:
                got = _norm(fetched.get(_norm(emp)))
                if got:
                    names[_norm(emp)] = got
        except Exception as e:  # noqa: BLE001 — 氏名が空でも投入はできる
            warnings.append(f"⚠️ 氏名を取得できませんでした（{e}）。氏名は空欄のままにします。")
    still = [e for e in emp_ids if not names.get(_norm(e))]
    if still:
        warnings.append(
            "⚠️ 台帳に無い社員番号があります（氏名は空欄で投入します）: " + "、".join(still))
    return names, warnings
