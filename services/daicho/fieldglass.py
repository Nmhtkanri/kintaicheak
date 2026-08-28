# -*- coding: utf-8 -*-
"""SAP Fieldglass の Work Order 一覧CSV（サプライヤー側レポート）を読む。

2026年4月にユニアデックスの派遣管理が e-staffing → Fieldglass に移った分の入力。
レポートの列（13列・UTF-8 BOM）:
  求人情報 ID / 求人情報タイトル / 事業単位 / コストセンター / 応募者の ID / 応募者/スタッフ /
  メインドキュメント ID / 改訂番号 / 勤務地 / 作業オーダー開始日 / 作業オーダー終了日 /
  作業オーダー: スタッフのスーパーバイザ / 作業オーダー ID

構造の癖:
- 行は WO × 改訂 × コストセンター で重複する（期間などは改訂内で同一）
- **改訂ごとに期間が違う**（四半期ごとに改訂される）。最新改訂＝今の期間なので、
  過去の四半期を作るときは「対象期間に重なる改訂」を選ぶこと
- 氏名は漢字の「姓, 名」形式。**一部は「名, 姓」と逆に登録されている**ため、両方向で名寄せする
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from .config import overlaps
from .roster import normalize_name


def _parse_date(s: str) -> dt.date | None:
    s = (s or "").strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def name_candidates(raw: str) -> list[str]:
    """『姓, 名』→ 正規化キー（カンマ除去）。逆順（名, 姓）の可能性も返す。"""
    parts = [p for p in re.split(r"[,、，]", unicodedata.normalize("NFKC", raw or "")) if p.strip()]
    cands = [normalize_name("".join(parts))] if parts else [normalize_name(raw)]
    if len(parts) == 2:
        cands.append(normalize_name(parts[1] + parts[0]))
    return [c for c in dict.fromkeys(cands) if c]


@dataclass
class WorkOrder:
    wo_id: str
    revision: int
    staff_raw: str            # 「姓, 名」のまま
    staff_fg_id: str
    start: dt.date | None
    end: dt.date | None       # None = 継続中（終了日未定）
    site: str                 # 勤務地
    business_unit: str        # 事業単位
    supervisors: list[str] = field(default_factory=list)
    job_title: str = ""
    job_posting_id: str = ""
    cost_centers: list[str] = field(default_factory=list)

    @property
    def staff_display(self) -> str:
        parts = [p.strip() for p in re.split(r"[,、，]", self.staff_raw) if p.strip()]
        return " ".join(parts) if parts else self.staff_raw


def load_workorders(path: Path | str) -> list[WorkOrder]:
    raw = Path(path).read_bytes()
    text = raw.decode("utf-8-sig")
    rows = list(csv.DictReader(io.StringIO(text)))
    merged: dict[tuple[str, str, int], WorkOrder] = {}
    for r in rows:
        wo_id = (r.get("作業オーダー ID") or "").strip()
        if not wo_id:
            continue
        rev = int((r.get("改訂番号") or "0").strip() or 0)
        fg_id = (r.get("応募者の ID") or "").strip()
        key = (wo_id, fg_id, rev)
        wo = merged.get(key)
        if wo is None:
            wo = WorkOrder(
                wo_id=wo_id, revision=rev, staff_raw=(r.get("応募者/スタッフ") or "").strip(),
                staff_fg_id=fg_id,
                start=_parse_date(r.get("作業オーダー開始日", "")),
                end=_parse_date(r.get("作業オーダー終了日", "")),
                site=(r.get("勤務地") or "").strip(),
                business_unit=(r.get("事業単位") or "").strip(),
                job_title=(r.get("求人情報タイトル") or "").strip(),
                job_posting_id=(r.get("求人情報 ID") or "").strip(),
            )
            merged[key] = wo
        sup = (r.get("作業オーダー: スタッフのスーパーバイザ") or "").strip()
        if sup and sup not in wo.supervisors:
            wo.supervisors.append(sup)
        cc = (r.get("コストセンター") or "").strip()
        if cc and cc not in wo.cost_centers:
            wo.cost_centers.append(cc)
    return sorted(merged.values(), key=lambda w: (w.staff_raw, w.wo_id, w.revision))


# WO に無く、直近の e-staffing 契約から引き継ぐ項目（台帳の備考・警告に明記する）
FG_CARRIED_FIELDS = "派遣先責任者・苦情申出先・責任の程度・業務内容・就業時間・休憩・休日"
_CMD = "事業所の名称及び所在地その他派遣就業場所"

_DOW = ["月", "火", "水", "木", "金", "土", "日"]


def parse_workdays(s: str) -> tuple[str, str]:
    """『月曜 勤務-火曜 勤務-…-日曜 非勤務』→ (就業曜日 '月 火 …', 休日 '土 日')。"""
    work, off = [], []
    for seg in (s or "").split("-"):
        seg = seg.strip()
        m = re.match(r"^([月火水木金土日])曜?\s*(勤務|非勤務)$", seg)
        if not m:
            continue
        (work if m.group(2) == "勤務" else off).append(m.group(1))
    order = {d: i for i, d in enumerate(_DOW)}
    work.sort(key=lambda d: order[d])
    off.sort(key=lambda d: order[d])
    return " ".join(work), " ".join(off)


def load_details(json_path: Path | str) -> tuple[dict, dict]:
    """Codex が Fieldglass の Job Posting/WO 詳細から取得した JSON を読む。

    戻り値: (staffId → entry, workOrderId → entry)。entry の主なキー:
    clientResponsible / complaintRecipient（氏名・役職・部署・電話）、responsibilityDegree、
    businessContent、workdaysHolidays(+Notes)、workStart/End、breakStart/End。
    値が空の項目は上書きに使わない（引き継ぎ値を保持する）。
    """
    import json

    data = json.loads(Path(json_path).read_text(encoding="utf-8"))
    by_staff = {e["staffId"]: e for e in data if e.get("staffId")}
    by_wo = {e["workOrderId"]: e for e in data if e.get("workOrderId")}
    return by_staff, by_wo


def detail_has_schedule(detail: dict | None) -> bool:
    return bool(detail and (detail.get("workStart") or "").strip())


# 引き継ぎ元契約が無い新規配属者に、e-staffing 全体の標準値（最頻値）で補う項目。
# 谷津さん指示（2026-08-28）:「e-staffing の過去データをもとに作成して構わない」
_DEFAULT_CPI_KEYS = (
    "36協定1 時間外労働、休日労働", "協定対象派遣労働者に該当するか否かの別",
    "健康保険", "厚生年金", "雇用保険",
    "派遣元責任者 部署", "派遣元責任者 TEL", "派遣元責任者 氏名",
    "派遣元苦情申出先 部署", "派遣元苦情申出先 TEL", "派遣元苦情申出先 氏名",
    "派遣元企業 名称", "派遣元企業 派遣許可番号", "派遣元企業 事業所名", "派遣元企業 事業所所在地",
)
_DEFAULT_TC_KEYS = ("安全及び衛生", "法定休日", "事業所抵触日")


def derive_defaults(contracts, q_start: dt.date, q_end: dt.date, client_hint: str = "ユニアデックス") -> dict:
    """e-staffing 契約群から会社共通の標準値を最頻値で作る（対象期に重なる契約を優先）。"""
    from collections import Counter

    from .config import overlaps

    pool = [c for c in contracts if overlaps(c.start, c.end, q_start, q_end)] or list(contracts)
    client_pool = [c for c in pool if client_hint in (c.client_name or "")] or pool

    def mode(cands, getter) -> str:
        counter = Counter(v for v in (getter(c) for c in cands) if v)
        return counter.most_common(1)[0][0] if counter else ""

    cpi = {k: mode(pool, lambda c, k=k: c.c(k)) for k in _DEFAULT_CPI_KEYS}
    tc = {k: mode(pool, lambda c, k=k: c.t(k)) for k in _DEFAULT_TC_KEYS}
    # 事業所抵触日は派遣先ごとの値なので、同じ派遣先の契約に絞った最頻値を使う
    tc["事業所抵触日"] = mode(client_pool, lambda c: c.t("事業所抵触日"))
    return {"cpi": cpi, "tc": tc}


def apply_defaults(contract, defaults: dict) -> list[str]:
    """空欄の項目だけ標準値で埋める。埋めた項目名を返す（警告表示用）。"""
    filled: list[str] = []
    for key, val in defaults.get("cpi", {}).items():
        if val and not contract.cpi.get(key):
            contract.cpi[key] = val
            filled.append(key)
    for key, val in defaults.get("tc", {}).items():
        if val and not contract.tc.get(key):
            contract.tc[key] = val
            filled.append(key)
    return filled


def fill_from_person(contract, person, wo_start: dt.date | None) -> None:
    """新規配属者の雇用形態・期間制限の対象外理由・個人抵触日を jinjer の人属性から埋める。"""
    if person is None or not person.employment_type:
        return
    mukei = person.employment_type in ("正社員", "役員")
    if not contract.cpi.get("派遣元での雇用形態"):
        contract.cpi["派遣元での雇用形態"] = "無期雇用契約" if mukei else "有期雇用契約"
    age = person.age_at(wo_start)
    reasons = []
    if mukei:
        reasons.append("無期雇用派遣労働者")
    if age is not None and age >= 60:
        reasons.append("60歳以上派遣労働者")
    if reasons:
        if not contract.cpi.get("期間制限の対象外理由"):
            contract.cpi["期間制限の対象外理由"] = " / ".join(reasons)
        if not contract.cpi.get("個人抵触日"):
            contract.cpi["個人抵触日"] = "--/--/-- (期間制限の対象外)"
    elif not contract.cpi.get("個人抵触日") and wo_start is not None:
        # 有期かつ60歳未満: 受け入れ開始から3年（正確な起算は組織単位での受入開始日＝要確認）
        contract.cpi["個人抵触日"] = f"{wo_start.year + 3}/{wo_start.month:02d}/{wo_start.day:02d}（受入開始から3年・要確認）"


def contract_from_workorder(wo: WorkOrder, src, q_end: dt.date, person=None, detail: dict | None = None):
    """WO を Contract の形に写す。src（直近の e-staffing 契約）があれば土台にして、
    期間・組織単位・指揮命令者を WO の現行値で、責任者・業務内容・就業時間などを
    SAP 詳細（detail）で置き換える。detail の空欄は引き継ぎ値のまま。"""
    from .estaffing import Contract

    tc = dict(src.tc) if src is not None else {}
    cpi = dict(src.cpi) if src is not None else {}
    start = wo.start
    end = wo.end or q_end
    s = start.strftime("%Y/%m/%d") if start else ""
    e = end.strftime("%Y/%m/%d")
    cpi["派遣期間 開始日"] = s
    cpi["派遣期間 終了日"] = e
    tc["契約開始日"] = s
    tc["契約終了日"] = e
    tc["契約確定日"] = ""
    if wo.business_unit:
        tc["組織単位"] = wo.business_unit
    if wo.supervisors:
        cpi[f"{_CMD} 指揮命令者部署"] = wo.business_unit or cpi.get(f"{_CMD} 指揮命令者部署", "")
        cpi[f"{_CMD} 指揮命令者役職"] = ""
        cpi[f"{_CMD} 指揮命令者氏名"] = "、".join(wo.supervisors)
    if detail:
        cr = detail.get("clientResponsible") or {}
        if (cr.get("name") or "").strip():
            cpi["派遣先責任者 部署"] = (cr.get("department") or "").strip()
            cpi["派遣先責任者 役職"] = (cr.get("title") or "").strip()
            cpi["派遣先責任者 氏名"] = (cr.get("name") or "").strip()
            cpi["派遣先責任者 TEL"] = (cr.get("phone") or "").strip()
        co = detail.get("complaintRecipient") or {}
        if (co.get("name") or "").strip():
            cpi["派遣先苦情申出先 部署"] = (co.get("department") or "").strip()
            cpi["派遣先苦情申出先 役職"] = (co.get("title") or "").strip()
            cpi["派遣先苦情申出先 氏名"] = (co.get("name") or "").strip()
            cpi["派遣先苦情申出先 TEL"] = (co.get("phone") or "").strip()
        if (detail.get("responsibilityDegree") or "").strip():
            cpi["責任の程度"] = detail["responsibilityDegree"].strip()
        if (detail.get("businessContent") or "").strip():
            cpi["業務内容"] = detail["businessContent"].strip()
        workdays, offdays = parse_workdays(detail.get("workdaysHolidays", ""))
        if workdays:
            cpi["勤務日"] = workdays
        if offdays:
            cpi["休日"] = offdays
        notes = (detail.get("workdaysHolidaysNotes") or "").strip()
        if notes:
            tc["休日（その他）"] = notes
        if (detail.get("workStart") or "").strip():
            cpi["就業時間 開始時間"] = detail["workStart"].strip()
            cpi["就業時間 終了時間"] = (detail.get("workEnd") or "").strip()
            cpi["就業時間 就業時間"] = ""      # 引き継ぎの実働時間は時刻と食い違いうるので消す
        if (detail.get("breakStart") or "").strip():
            cpi["休憩時間1 開始時間"] = detail["breakStart"].strip()
            cpi["休憩時間1 終了時間"] = (detail.get("breakEnd") or "").strip()
            for n in ("1 時間", "2 開始時間", "2 終了時間", "2 時間", "3 開始時間", "3 終了時間", "3 時間"):
                cpi[f"休憩時間{n}"] = ""
    if person is not None:
        cpi["労働者氏名"] = person.name
        tc["スタッフ姓（日本語）"] = person.sei
        tc["スタッフ名（日本語）"] = person.mei
    elif not src:
        cpi["労働者氏名"] = wo.staff_display
    if not cpi.get("派遣先企業 名称"):
        cpi["派遣先企業 名称"] = "ユニアデックス株式会社"
    return Contract(contract_no=f"FG-{wo.wo_id}(改訂{wo.revision})", tc=tc, cpi=cpi)


def workorders_in_quarter(wos: list[WorkOrder], q_start: dt.date, q_end: dt.date) -> list[WorkOrder]:
    """対象四半期に重なる改訂を、(WO, スタッフ) ごとに1つ選ぶ（重なる中で最大の改訂番号）。

    終了日が空の改訂は「継続中」= 期末まで続くとみなす。
    """
    best: dict[tuple[str, str], WorkOrder] = {}
    for w in wos:
        end = w.end or q_end
        if not overlaps(w.start, end, q_start, q_end):
            continue
        key = (w.wo_id, w.staff_fg_id)
        cur = best.get(key)
        if cur is None or w.revision > cur.revision:
            best[key] = w
    return sorted(best.values(), key=lambda w: (w.staff_raw, w.wo_id))
