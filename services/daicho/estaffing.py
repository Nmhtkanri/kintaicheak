# -*- coding: utf-8 -*-
"""e-staffing の2種類のCSVを読み、契約Noで結合して Contract にする。

- TCnmht*.csv  … 契約データ（222列）。契約確定日・組織単位・組織の長の職名・事業所抵触日・
                  法定休日・便宜供与・安全衛生・保険の対象外理由（フラグ）などが入る。
- CPInmht*.csv … 契約書・通知書データ（205列）。責任の程度・協定対象か否か・個人抵触日・
                  保険の文言（有/無(加入対象外)）・勤務日/休日/就業時間の文言が入る。

列は**ヘッダー名**で引く（列位置は世代で変わる。旧ブックの位置は信用しない）。
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import re
from dataclasses import dataclass, field
from pathlib import Path

SHIFT_NOTE = "別途シフト表に定める"


def read_csv(path: Path | str, encoding: str = "cp932") -> tuple[list[str], list[dict[str, str]]]:
    """CSVを辞書行で返す。重複ヘッダーは2個目以降に #2, #3 を付けて区別する。"""
    raw = Path(path).read_bytes()
    text = raw.decode(encoding)
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return [], []
    header: list[str] = []
    seen: dict[str, int] = {}
    for name in rows[0]:
        name = name.strip().lstrip("\ufeff")
        n = seen.get(name, 0) + 1
        seen[name] = n
        header.append(name if n == 1 else f"{name}#{n}")
    records = []
    for r in rows[1:]:
        if not any(c.strip() for c in r):
            continue
        d = {h: (r[i].strip() if i < len(r) else "") for i, h in enumerate(header)}
        records.append(d)
    return header, records


def parse_date(s: str | None) -> dt.date | None:
    if not s:
        return None
    s = s.strip()
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y%m%d", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def split_time(s: str) -> tuple[str, bool]:
    """'09:00 別途シフト表に定める' → ('09:00', True)。'--:--' は ('', False)。"""
    s = (s or "").strip()
    shift = SHIFT_NOTE in s
    s = s.replace(SHIFT_NOTE, "").strip()
    if s in ("", "--:--"):
        return "", shift
    m = re.match(r"^(\d{1,2}):(\d{2})", s)
    return (f"{int(m.group(1)):02d}:{m.group(2)}" if m else s), shift


@dataclass
class Contract:
    contract_no: str
    tc: dict[str, str] = field(default_factory=dict)
    cpi: dict[str, str] = field(default_factory=dict)

    # --- 取り出しヘルパー（無い列は空文字） ---
    def t(self, name: str) -> str:
        return self.tc.get(name, "")

    def c(self, name: str) -> str:
        return self.cpi.get(name, "")

    @property
    def start(self) -> dt.date | None:
        return parse_date(self.c("派遣期間 開始日")) or parse_date(self.t("契約開始日"))

    @property
    def end(self) -> dt.date | None:
        return parse_date(self.c("派遣期間 終了日")) or parse_date(self.t("契約終了日"))

    @property
    def worker_name(self) -> str:
        n = self.c("労働者氏名")
        if n:
            return n
        return f"{self.t('スタッフ姓（日本語）')} {self.t('スタッフ名（日本語）')}".strip()

    @property
    def worker_sei_mei(self) -> tuple[str, str]:
        sei, mei = self.t("スタッフ姓（日本語）"), self.t("スタッフ名（日本語）")
        if sei or mei:
            return sei, mei
        parts = re.split(r"[\s\u3000]+", self.worker_name, maxsplit=1)
        return (parts[0], parts[1] if len(parts) > 1 else "")

    @property
    def client_name(self) -> str:
        return self.c("派遣先企業 名称") or self.t("契約先企業名") or self.t("就業先企業名")


def load_contracts(tc_path: Path | str, cpi_path: Path | str) -> tuple[list[Contract], list[str]]:
    """契約データ(TC)と契約書・通知書データ(CPI)を契約Noで結合する。

    戻り値: (契約リスト, 警告リスト)。どちらか片方にしか無い契約は警告に出し、
    TC側があれば CPI 空のまま残す（CPI側だけのものは落とす）。
    """
    warnings: list[str] = []
    _, tc_rows = read_csv(tc_path)
    _, cpi_rows = read_csv(cpi_path)
    tc_key = "契約No"
    cpi_key = "e-staffing契約No"
    if tc_rows and tc_key not in tc_rows[0]:
        raise ValueError(f"契約データ(TC)に『{tc_key}』列がありません: {tc_path}")
    if cpi_rows and cpi_key not in cpi_rows[0]:
        raise ValueError(f"契約書・通知書データ(CPI)に『{cpi_key}』列がありません: {cpi_path}")
    cpi_by = {r[cpi_key]: r for r in cpi_rows}
    contracts: list[Contract] = []
    seen: set[str] = set()
    for r in tc_rows:
        no = r[tc_key]
        if no in seen:
            warnings.append(f"契約データに契約No {no} が重複しています（後の行を無視）")
            continue
        seen.add(no)
        cpi = cpi_by.pop(no, None)
        if cpi is None:
            warnings.append(f"契約No {no}: 契約書・通知書データ(CPI)に無い → 責任の程度・協定対象などが空になります")
            cpi = {}
        else:
            ts, te = parse_date(r.get("契約開始日")), parse_date(r.get("契約終了日"))
            cs, ce = parse_date(cpi.get("派遣期間 開始日")), parse_date(cpi.get("派遣期間 終了日"))
            if ts and cs and (ts, te) != (cs, ce):
                warnings.append(f"契約No {no}: 契約データの期間 {ts}～{te} と契約書・通知書データの派遣期間 {cs}～{ce} が違う（契約書側を採用）")
        contracts.append(Contract(contract_no=no, tc=r, cpi=cpi))
    for no in cpi_by:
        warnings.append(f"契約No {no}: 契約書・通知書データ(CPI)にだけ存在（契約データ側に無いため台帳には載せません）")
    return contracts, warnings


def contracts_in_quarter(contracts: list[Contract], q_start: dt.date, q_end: dt.date) -> list[Contract]:
    from .config import overlaps
    hit = [c for c in contracts if overlaps(c.start, c.end, q_start, q_end)]
    hit.sort(key=lambda c: (c.worker_name, c.start or dt.date.min, c.contract_no))
    return hit
