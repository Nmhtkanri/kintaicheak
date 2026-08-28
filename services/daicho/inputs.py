# -*- coding: utf-8 -*-
"""入力ファイルの解決と鮮度チェック（build と画面で共有し、glob 定義の二重管理を防ぐ）。

input/ 直下から「mtime 最新の1本」を自動選択する。パターンを変えるときはここだけ直す。
ハブ画面用の check_freshness()・quarter_status() もここ（入力と成果物を見るだけで何も書かない）。
"""
from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path

from . import config

# build が input/ から自動選択する glob
PATTERN_TC = "TCnmht*.csv"
PATTERN_CPI = "CPInmht*.csv"
PATTERN_ROSTER = "従業員一覧*.xlsx"
PATTERN_FG_WO = "*WorkOrder*.csv"
PATTERN_FG_DETAILS = "*fieldglass_details*.json"
# 2026-08-28〜: Fieldglass レポートスケジュールでメール添付が届く新形式（四半期ごと）
PATTERN_FG_UAL_REPORT = "*ユニアデックス*業務内容*.xlsx"
PATTERN_FG_ERICSSON = "派遣元管理台帳作成用*.csv"


def _glob_newest(folder: Path, pattern: str) -> list[Path]:
    # Excel が開いている間のロックファイル（~$〜.xlsx）は拾わない
    files = [p for p in folder.glob(pattern) if not p.name.startswith("~$")]
    return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)


def newest(folder: Path, pattern: str) -> Path:
    """folder 直下で pattern に合う mtime 最新の1本。無ければ FileNotFoundError。"""
    files = _glob_newest(folder, pattern)
    if not files:
        raise FileNotFoundError(f"入力が見つかりません: {folder}\\{pattern}")
    return files[0]


def newest_or_none(folder: Path, pattern: str) -> Path | None:
    """newest と同じ選び方で、無ければ None（Fieldglass 系は任意入力のため）。"""
    files = _glob_newest(folder, pattern)
    return files[0] if files else None


# =============================================================================
# 鮮度チェック（①）と四半期ステータス（ステッパー用）
# =============================================================================

# ファイル名から「データの日付」を取る。NASへのコピーで mtime は汚れるため名前優先。
_STAMP14 = re.compile(r"(\d{14})\.csv$", re.IGNORECASE)         # TCnmht/CPInmht の DL 時刻
_DATE_YMD = re.compile(r"(\d{4}-\d{2}-\d{2})")                   # 従業員一覧_2026-08-21.xlsx
_DATE_8 = re.compile(r"_(\d{8})")                                # fieldglass_details_20260828.json


def default_quarter(today: dt.date | None = None) -> str:
    """直前の（＝直近で締まった）四半期。1〜3月なら前年Q4。"""
    today = today or dt.date.today()
    qn = (today.month - 1) // 3 + 1
    return f"{today.year - 1}Q4" if qn == 1 else f"{today.year}Q{qn - 1}"


def _mtime_dt(p: Path) -> dt.datetime:
    return dt.datetime.fromtimestamp(p.stat().st_mtime)


def _fmt(d: dt.datetime | dt.date | None) -> str:
    if d is None:
        return ""
    if isinstance(d, dt.datetime):
        return d.strftime("%Y-%m-%d %H:%M")
    return d.strftime("%Y-%m-%d")


def _row(key: str, label: str, **kw) -> dict:
    base = {"key": key, "label": label, "filename": "", "date": "", "date_source": "",
            "in_quarter": None, "verdict": "ok", "note": ""}
    base.update(kw)
    return base


def _name_stamp14(p: Path) -> dt.datetime | None:
    m = _STAMP14.search(p.name)
    if not m:
        return None
    try:
        return dt.datetime.strptime(m.group(1), "%Y%m%d%H%M%S")
    except ValueError:
        return None


def check_freshness(quarter: str, quick: bool = False) -> dict:
    """入力ファイルごとの鮮度と対象期の該当件数を返す（読むだけ・何も書かない）。

    quick=True は存在と日付だけ（ステッパー用。CSVのパースを省く）。
    verdict: ok / warn / missing / info。overall: ok / warn / missing。
    """
    q = quarter.upper()
    q_start, q_end = config.quarter_range(q)
    today = dt.date.today()
    rows: list[dict] = []

    # --- e-staffing TC / CPI（必須。ファイル名の14桁=DL時刻で判定）---
    tc = newest_or_none(config.INPUT_DIR, PATTERN_TC)
    cpi = newest_or_none(config.INPUT_DIR, PATTERN_CPI)
    for key, label, p in (("tc", "e-staffing 契約データ (TC)", tc),
                          ("cpi", "e-staffing 契約書・通知書 (CPI)", cpi)):
        if p is None:
            rows.append(_row(key, label, verdict="missing",
                             note=f"input に {PATTERN_TC if key == 'tc' else PATTERN_CPI} がありません"))
            continue
        stamp = _name_stamp14(p)
        if stamp is not None:
            date, source = stamp, "ファイル名"
        else:
            date, source = _mtime_dt(p), "更新日時"
        if date.date() >= q_end:
            rows.append(_row(key, label, filename=p.name, date=_fmt(date), date_source=source))
        else:
            rows.append(_row(key, label, filename=p.name, date=_fmt(date), date_source=source,
                             verdict="warn",
                             note=f"対象期が締まる前（{_fmt(date)}時点）のダウンロードです。期末後に取り直すと確定分が入ります"))
    if not quick and tc is not None and cpi is not None:
        try:
            from .estaffing import contracts_in_quarter, load_contracts
            contracts, _warns = load_contracts(tc, cpi)
            rows[0]["in_quarter"] = len(contracts_in_quarter(contracts, q_start, q_end))
            rows[1]["note"] = (rows[1]["note"] + "／" if rows[1]["note"] else "") + "期内件数はTCと結合して左の行に表示"
        except Exception as e:  # noqa: BLE001 — 件数は補助情報。読めなくても日付判定は返す
            rows[0]["note"] = (rows[0]["note"] + "／" if rows[0]["note"] else "") + f"期内件数を数えられませんでした: {e}"

    # --- 従業員一覧 xlsx（必須。名前の YYYY-MM-DD）---
    roster = newest_or_none(config.INPUT_DIR, PATTERN_ROSTER)
    if roster is None:
        rows.append(_row("roster", "jinjer 従業員一覧 (xlsx)", verdict="missing",
                         note=f"input に {PATTERN_ROSTER} がありません"))
    else:
        m = _DATE_YMD.search(roster.name)
        if m:
            date = dt.datetime.strptime(m.group(1), "%Y-%m-%d")
            source = "ファイル名"
        else:
            date, source = _mtime_dt(roster), "更新日時"
        age = (today - date.date()).days
        rows.append(_row("roster", "jinjer 従業員一覧 (xlsx)", filename=roster.name,
                         date=_fmt(date.date()), date_source=source,
                         verdict="ok" if age <= 60 else "warn",
                         note="" if age <= 60 else f"{age}日前のものです（新入社員が落ちる可能性）"))

    # --- jinjer API 人マスタキャッシュ（中の fetched_at）---
    cache = config.INPUT_DIR / "jinjer_api_roster.json"
    if not cache.exists():
        rows.append(_row("api_cache", "jinjer API 人マスタ（退職者込み）", verdict="warn",
                         note="キャッシュなし。buildの「jinjer人マスタをAPIから取り直す」で作成できます"))
    else:
        fetched = ""
        try:
            fetched = json.loads(cache.read_text(encoding="utf-8")).get("fetched_at", "")
        except Exception:  # noqa: BLE001
            pass
        age = None
        try:
            age = (today - dt.datetime.strptime(fetched[:10], "%Y-%m-%d").date()).days
        except ValueError:
            pass
        rows.append(_row("api_cache", "jinjer API 人マスタ（退職者込み）", filename=cache.name,
                         date=fetched, date_source="fetched_at",
                         verdict="ok" if (age is not None and age <= 30) else "warn",
                         note="" if (age is not None and age <= 30)
                         else "30日より古いキャッシュです。buildの「jinjer人マスタをAPIから取り直す」で更新できます"))

    # --- Fieldglass（新レポートがあればそれが正。無ければ旧2段構えを表示）---
    ual = newest_or_none(config.INPUT_DIR, PATTERN_FG_UAL_REPORT)
    if ual is not None:
        m = _DATE_8.search(ual.name)
        u_date = dt.datetime.strptime(m.group(1), "%Y%m%d").date() if m else _mtime_dt(ual).date()
        row = _row("fg_ual_report", "FG新レポート（ユニアデックス）", filename=ual.name,
                   date=_fmt(u_date), date_source="ファイル名" if m else "更新日時",
                   note="旧 WO CSV＋詳細JSON の代わりにこれ1本を読みます")
        if not quick:
            try:
                from .fieldglass import workorders_in_quarter
                from .fieldglass_report import load_ual_report
                wos, _s, _w = load_ual_report(ual)
                n = len(workorders_in_quarter(wos, q_start, q_end))
                row["in_quarter"] = n
                if n == 0:
                    row["verdict"] = "warn"
                    row["note"] = "対象期に重なるWOが0件（期のズレたレポートの疑い）"
            except Exception as e:  # noqa: BLE001
                row["verdict"] = "warn"
                row["note"] = f"レポートを読めませんでした: {e}"
        rows.append(row)
    else:
        fg = newest_or_none(config.INPUT_DIR, PATTERN_FG_WO)
        if fg is None:
            rows.append(_row("fg_wo", "Fieldglass WorkOrder CSV", verdict="warn",
                             note="見つかりません（Fieldglass 分は台帳に入りません）"))
        else:
            row = _row("fg_wo", "Fieldglass WorkOrder CSV", filename=fg.name,
                       date=_fmt(_mtime_dt(fg)), date_source="更新日時")
            if not quick:
                try:
                    from .fieldglass import load_workorders, workorders_in_quarter
                    n = len(workorders_in_quarter(load_workorders(fg), q_start, q_end))
                    row["in_quarter"] = n
                    if n == 0:
                        row["verdict"] = "warn"
                        row["note"] = "対象期に重なるWOが0件（期のズレたレポートの疑い）"
                except Exception as e:  # noqa: BLE001
                    row["note"] = f"WO件数を数えられませんでした: {e}"
            rows.append(row)

        det = newest_or_none(config.INPUT_DIR, PATTERN_FG_DETAILS)
        if det is None:
            rows.append(_row("fg_details", "SAP詳細 JSON (fieldglass_details)", verdict="warn",
                             note="見つかりません（派遣先責任者等は引き継ぎ値になります）"))
        else:
            m = _DATE_8.search(det.name)
            d_date = dt.datetime.strptime(m.group(1), "%Y%m%d").date() if m else _mtime_dt(det).date()
            row = _row("fg_details", "SAP詳細 JSON (fieldglass_details)", filename=det.name,
                       date=_fmt(d_date), date_source="ファイル名" if m else "更新日時")
            if fg is not None and d_date < _mtime_dt(fg).date():
                row["verdict"] = "warn"
                row["note"] = "WorkOrder CSV より古い詳細です（取得日のズレ。2段構えの片方だけ更新の疑い）"
            rows.append(row)

    # --- FG新レポート（エリクソン。直接契約3人の現行値上書きに使う）---
    eric = newest_or_none(config.INPUT_DIR, PATTERN_FG_ERICSSON)
    if eric is None:
        rows.append(_row("fg_ericsson", "FG新レポート（エリクソン）", verdict="info",
                         note="無し（エリクソン分は手入力マスタの値のまま）"))
    else:
        m = _DATE_8.search(eric.name)
        e_date = dt.datetime.strptime(m.group(1), "%Y%m%d").date() if m else _mtime_dt(eric).date()
        rows.append(_row("fg_ericsson", "FG新レポート（エリクソン）", filename=eric.name,
                         date=_fmt(e_date), date_source="ファイル名" if m else "更新日時",
                         verdict="info", note="直接契約（エリクソン3人）の責任者・就業時間を現行値に更新"))

    # --- 直接契約マスタ（固定名。自動＋手入力の結合で期内件数）---
    from .direct import AUTO_CSV, MANUAL_CSV
    d_count = None
    if not quick:
        try:
            from .direct import load_master, rows_in_quarter
            d_count = len(rows_in_quarter(load_master(), q_start, q_end))
        except Exception:  # noqa: BLE001
            d_count = None
    for key, label, p in (("direct_auto", "直接契約マスタ（自動）", AUTO_CSV),
                          ("direct_manual", "直接契約マスタ（手入力）", MANUAL_CSV)):
        if not p.exists():
            rows.append(_row(key, label, verdict="warn",
                             note="ありません（この分の直接契約は台帳に入りません）"))
        else:
            rows.append(_row(key, label, filename=p.name, date=_fmt(_mtime_dt(p)),
                             date_source="更新日時", verdict="info",
                             in_quarter=d_count if key == "direct_auto" else None,
                             note="自動＋手入力の結合後の期内件数" if key == "direct_auto" and d_count is not None else ""))

    # --- テンプレート（無ければ build が自動生成）---
    if config.TEMPLATE_XLSX.exists():
        rows.append(_row("template", "台帳テンプレート", filename=config.TEMPLATE_XLSX.name,
                         date=_fmt(_mtime_dt(config.TEMPLATE_XLSX)), date_source="更新日時", verdict="info"))
    else:
        rows.append(_row("template", "台帳テンプレート", verdict="info",
                         note="無ければ build が旧フォームから自動生成します"))

    required_missing = any(r["verdict"] == "missing" for r in rows)
    has_warn = any(r["verdict"] == "warn" for r in rows)
    overall = "missing" if required_missing else ("warn" if has_warn else "ok")
    return {"quarter": q, "label": config.quarter_label(q), "inputs": rows, "overall": overall}


def quarter_status(quarter: str | None = None) -> dict:
    """ステッパー用の四半期ステータス（見るだけ・何も書かない）。

    steps: freshness（quickの全体判定）/ build（xlsxの有無）/ pdf（期内PDF枚数）/
           attach（進捗JSONの要約）。due は 1・4・7・10月に直前四半期が未作成のとき。
    """
    q = (quarter or default_quarter()).upper()
    label = config.quarter_label(q)          # 不正な四半期はここで ValueError
    today = dt.date.today()

    freshness = {"overall": "unknown"}
    try:
        freshness = {"overall": check_freshness(q, quick=True)["overall"]}
    except Exception:  # noqa: BLE001 — NASが読めなくてもステッパーは返す
        pass

    xlsx = config.OUTPUT_DIR / f"派遣元管理台帳_{q}.xlsx"
    build = {"exists": xlsx.exists()}
    if build["exists"]:
        build["file"] = xlsx.name
        build["mtime"] = _fmt(_mtime_dt(xlsx))
        warn_csv = config.OUTPUT_DIR / f"派遣元管理台帳_{q}_警告.csv"
        if warn_csv.exists():
            try:
                n_lines = sum(1 for _ in warn_csv.open("r", encoding="utf-8-sig"))
                build["n_warn"] = max(n_lines - 1, 0)
            except OSError:
                pass

    pdf_suffix = f"{label[:-1]}分.pdf"       # '2025年7-9月期' → '*_2025年7-9月分.pdf'
    try:
        pdf_count = sum(1 for _ in config.PDF_ROOT.glob(f"*/*_{pdf_suffix}"))
    except OSError:
        pdf_count = None
    pdf = {"count": pdf_count}

    attach = {"state": "none"}
    if config.ATTACH_PROGRESS_JSON.exists():
        try:
            data = json.loads(config.ATTACH_PROGRESS_JSON.read_text(encoding="utf-8"))
            attach = {k: data.get(k) for k in ("state", "done", "skip", "total", "quarter",
                                               "started_at", "start_at", "finished_at", "message")}
        except Exception:  # noqa: BLE001 — NASの書き込み途中で壊れて見えても前回表示のままにする
            attach = {"state": "unknown"}

    prev_q = default_quarter(today)
    due = (today.month in (1, 4, 7, 10)
           and not (config.OUTPUT_DIR / f"派遣元管理台帳_{prev_q}.xlsx").exists())
    return {"quarter": q, "label": label, "due": due, "due_quarter": prev_q,
            "steps": {"freshness": freshness, "build": build, "pdf": pdf, "attach": attach}}
