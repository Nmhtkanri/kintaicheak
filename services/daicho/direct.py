# -*- coding: utf-8 -*-
"""直接契約（e-staffing・Fieldglass 以外の紙/Excel契約）の第3ソース。

「直接契約マスタ」CSV（1契約1行）を読み、e-staffing の Contract と同じ形に写して
records パイプラインへ流す。マスタは2枚:
  - input\\直接契約マスタ_自動.csv … extract_direct が契約書から再生成する（手で直さない）
  - input\\直接契約マスタ_手入力.csv … スキャン等から手で起こした行（保護される）
同一キー（氏名×開始×終了×派遣先）は手入力が勝つ。
"""
from __future__ import annotations

import csv
import datetime as dt
import re
from pathlib import Path

from .config import INPUT_DIR, overlaps
from .estaffing import Contract, parse_date

AUTO_CSV = INPUT_DIR / "直接契約マスタ_自動.csv"
MANUAL_CSV = INPUT_DIR / "直接契約マスタ_手入力.csv"
_CMD = "事業所の名称及び所在地その他派遣就業場所"

# マスタの列（自動・手入力とも同じ）
COLUMNS = [
    "出所ファイル", "様式", "派遣先名称", "契約番号", "氏名", "氏名カナ", "性別",
    "派遣期間開始", "派遣期間終了", "就業場所名称", "就業場所住所", "組織単位",
    "業務内容", "責任の程度", "就業曜日", "就業時間", "休憩時間", "実労働時間", "休日", "時間外労働",
    "安全及び衛生", "便宜供与",
    "派遣先責任者_部署", "派遣先責任者_役職", "派遣先責任者_氏名", "派遣先責任者_TEL",
    "指揮命令者_部署", "指揮命令者_役職", "指揮命令者_氏名", "指揮命令者_TEL",
    "苦情申出先_部署", "苦情申出先_役職", "苦情申出先_氏名", "苦情申出先_TEL",
    "派遣元責任者_部署", "派遣元責任者_役職", "派遣元責任者_氏名", "派遣元責任者_TEL",
    "派遣元苦情_部署", "派遣元苦情_役職", "派遣元苦情_氏名", "派遣元苦情_TEL",
    "健康保険", "厚生年金", "雇用保険", "雇用形態", "限定の別", "事業所抵触日", "備考",
]


def _read(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as fp:
        return [dict(r) for r in csv.DictReader(fp)]


def _key(row: dict) -> tuple:
    return (re.sub(r"[\s　]+", "", row.get("氏名", "")), row.get("派遣期間開始", ""),
            row.get("派遣期間終了", ""), row.get("派遣先名称", ""))


def load_master(auto_csv: Path | str = AUTO_CSV, manual_csv: Path | str = MANUAL_CSV) -> list[dict]:
    rows = {_key(r): r for r in _read(Path(auto_csv))}
    for r in _read(Path(manual_csv)):        # 手入力が上書き
        rows[_key(r)] = {**r, "様式": (r.get("様式") or "手入力")}
    return [r for r in rows.values() if r.get("氏名") and r.get("派遣期間開始")]


_TIME_RE = re.compile(r"^\d{1,2}:\d{2}$")


def _split_range(s: str) -> tuple[str, str]:
    """'9:00～17:30' のような時刻2つの単純な形だけ分割する。
    シフト制の説明文（内側に～を含む）は分割せず開始側にそのまま入れ、台帳へ原文で流す。"""
    s = (s or "").strip()
    if not s:
        return "", ""
    m = re.split(r"[～〜]", s)
    if len(m) == 2 and _TIME_RE.match(m[0].strip()) and _TIME_RE.match(m[1].strip()):
        return m[0].strip(), m[1].strip()
    return s, ""


def contract_from_row(row: dict) -> Contract:
    """マスタ1行 → e-staffing と同じキーを持つ仮想 Contract。空欄は空のまま
    （協定対象・36協定・保険などの穴は build 側の標準値補完と jinjer 属性で埋まる）。"""
    g = lambda k: (row.get(k) or "").strip()
    cpi: dict[str, str] = {}
    tc: dict[str, str] = {}
    cpi["労働者氏名"] = g("氏名")
    parts = re.split(r"[\s　]+", g("氏名"), maxsplit=1)
    tc["スタッフ姓（日本語）"] = parts[0] if parts else ""
    tc["スタッフ名（日本語）"] = parts[1] if len(parts) > 1 else ""
    tc["スタッフ姓（カナ）"] = re.split(r"[\s　]+", g("氏名カナ"), maxsplit=1)[0] if g("氏名カナ") else ""
    cpi["派遣先企業 名称"] = g("派遣先名称")
    tc["就業先企業名"] = g("派遣先名称")
    cpi["派遣期間 開始日"] = g("派遣期間開始")
    cpi["派遣期間 終了日"] = g("派遣期間終了")
    tc["契約開始日"] = g("派遣期間開始")
    tc["契約終了日"] = g("派遣期間終了")
    cpi[f"{_CMD} 事業所の名称"] = g("就業場所名称")
    cpi[f"{_CMD} 事業所の所在地及び就業場所"] = g("就業場所住所")
    tc["組織単位"] = g("組織単位")
    cpi["業務内容"] = g("業務内容")
    cpi["責任の程度"] = g("責任の程度")
    cpi["勤務日"] = g("就業曜日")
    cpi["休日"] = g("休日")
    s, e = _split_range(g("就業時間"))
    cpi["就業時間 開始時間"], cpi["就業時間 終了時間"] = s, e
    cpi["就業時間 就業時間"] = g("実労働時間")
    bs, be = _split_range(g("休憩時間"))
    cpi["休憩時間1 開始時間"], cpi["休憩時間1 終了時間"] = bs, be
    cpi["36協定1 時間外労働、休日労働"] = g("時間外労働")
    tc["安全及び衛生"] = g("安全及び衛生")
    if g("便宜供与"):
        tc["便宜供与：その他1"] = g("便宜供与")
    for src, dst in (("派遣先責任者", "派遣先責任者"), ("苦情申出先", "派遣先苦情申出先"),
                     ("派遣元責任者", "派遣元責任者"), ("派遣元苦情", "派遣元苦情申出先")):
        cpi[f"{dst} 部署"] = g(f"{src}_部署")
        cpi[f"{dst} 役職"] = g(f"{src}_役職")
        cpi[f"{dst} 氏名"] = g(f"{src}_氏名")
        cpi[f"{dst} TEL"] = g(f"{src}_TEL")
    cpi[f"{_CMD} 指揮命令者部署"] = g("指揮命令者_部署")
    cpi[f"{_CMD} 指揮命令者役職"] = g("指揮命令者_役職")
    cpi[f"{_CMD} 指揮命令者氏名"] = g("指揮命令者_氏名")
    cpi[f"{_CMD} 部署TEL"] = g("指揮命令者_TEL") or g("派遣先責任者_TEL")
    for kind in ("健康保険", "厚生年金", "雇用保険"):
        if g(kind):
            cpi[kind] = g(kind)
    emp = g("雇用形態")
    if emp:
        cpi["派遣元での雇用形態"] = "無期雇用契約" if "無期" in emp else "有期雇用契約"
    tc["事業所抵触日"] = g("事業所抵触日")
    tc["契約確定日"] = ""
    no = g("契約番号") or f"{g('派遣先名称')}_{g('派遣期間開始')}"
    return Contract(contract_no=f"DC-{no}", tc=tc, cpi=cpi)


def rows_in_quarter(rows: list[dict], q_start: dt.date, q_end: dt.date) -> list[dict]:
    hit = []
    for r in rows:
        s = parse_date((r.get("派遣期間開始") or "").strip())
        e = parse_date((r.get("派遣期間終了") or "").strip())
        if s and e and overlaps(s, e, q_start, q_end):
            hit.append(r)
    hit.sort(key=lambda r: (r.get("氏名", ""), r.get("派遣期間開始", "")))
    return hit
