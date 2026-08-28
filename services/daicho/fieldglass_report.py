# -*- coding: utf-8 -*-
"""SAP Fieldglass の「レポートスケジュール」で届く新形式（2026-08-28 自動化）を読む。

四半期ごと（10/15・1/15・4/15・7/15 9:00）に谷津さん＋kanri宛てメールで添付が届き、
input/ へ保存するだけが収集作業（Codex のブラウザ転記＝details JSON は廃止方向）。

- ユニアデックス「派遣元管理台帳」…… XLSX。1行目タイトル・2行目ヘッダー・66列。
  行は WO×改訂×コストセンターで重複。旧 WO CSV＋details JSON の2段構えを1本で置き換える。
  ⚠ 終了日列は「最新の作業オーダー終了日」＝改訂共通の最新値。過去の改訂の終了日は
  「次の改訂の開始日−1日」で復元しないと、過去期の台帳の契約期間が伸びてしまう。
- エリクソン「派遣元管理台帳作成用」…… CSV 47列・3人。期間・WO列が無い（足すと行が
  全滅する既知の罠）ため台帳の行は従来どおり直接契約マスタ（手入力）から作り、
  このレポートは責任者・就業時間などの**現行値の上書きソース**として使う。
  重複ヘッダーが6組あり原則「最初の1個」を採用（Notes 系だけは両方を拾う＝
  片側に変形労働時間制の記述が入るため）。事業所抵触日・苦情申出先の氏名/部署は
  レポートに無い＝手入力マスタの値を維持する。
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import re
import unicodedata
from pathlib import Path

from .fieldglass import WorkOrder

# ---------------------------------------------------------------------------
# 共通ヘルパー
# ---------------------------------------------------------------------------

def _as_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.time):
        return v.strftime("%H:%M")
    s = str(v).strip()
    return "" if s in ("-", "－") else s          # レポートの「該当なし」表記は空扱い


def _as_date(v) -> dt.date | None:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = _as_str(v)
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def _as_time(v) -> str:
    """'9:00'・datetime.time・'09:00:00' → 'H:MM'（e-staffing と同じ表記のまま渡す）。"""
    if isinstance(v, dt.time):
        return f"{v.hour}:{v.minute:02d}"
    s = _as_str(v)
    m = re.match(r"^(\d{1,2}):(\d{2})", s)
    return f"{int(m.group(1))}:{m.group(2)}" if m else s


def _split5(v) -> dict:
    """『氏名;役職;部署;TEL;Email』のセミコロン5点を dict に。全角空白は保つ。"""
    parts = [p.strip() for p in _as_str(v).split(";")]
    parts += [""] * (5 - len(parts))
    return {"name": parts[0], "title": parts[1], "department": parts[2],
            "phone": parts[3], "email": parts[4]}


# ---------------------------------------------------------------------------
# ユニアデックス XLSX
# ---------------------------------------------------------------------------

# 日本語列（DTS＝別バイヤーのカスタム項目。ユニアデックス行では空、岡崎さんの行だけ値が入る）
_DTS_FALLBACK = {
    # entryキー: (英語列から作る関数の代わりに使う日本語列名たち)
    "responsibilityDegree": "従事する業務に伴う責任の程度",
    "workdaysHolidays": "就業日",
    "workdaysHolidaysNotes": "就業日・時間・時間外労働および休日労働に関する備考",
    "orgChief": "組織の長の職名",
    "orgUnit": "組織単位の名称",
    "training": "業務に必要な能力を付与するための教育訓練に関する事項",
}


def load_ual_report(path: Path | str) -> tuple[list[WorkOrder], dict, dict]:
    """ユニアデックスの新レポート1本 → (WorkOrder一覧, staffId→entry, workOrderId→entry)。

    entry は旧 details JSON と同じキー（clientResponsible / complaintRecipient /
    responsibilityDegree / businessContent / workdaysHolidays(+Notes) / workStart/End /
    breakStart/End / jobPostingId）に、レポートで増えた分（supervisor / orgUnit / orgChief /
    siteTenure / workOffice / workAddress / conveniences / agreementTarget / training /
    monthlyStdLower/Upper）を足した形。空の値は入れない（空は上書きに使わない既存規約）。
    """
    import openpyxl

    wb = openpyxl.load_workbook(Path(path), read_only=True, data_only=True)
    try:
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        next(rows)                       # 1行目=レポートタイトル
        header = next(rows)              # 2行目=ヘッダー
        idx: dict[str, int] = {}
        for i, h in enumerate(header):
            name = _as_str(h)
            if name and name not in idx:  # 最初の1個を採用
                idx[name] = i

        def col(row, name: str):
            i = idx.get(name)
            return row[i] if (i is not None and i < len(row)) else None

        merged: dict[tuple[str, str, int], WorkOrder] = {}
        entries: dict[tuple[str, str, int], dict] = {}
        for row in rows:
            wo_id = _as_str(col(row, "作業オーダー ID"))
            if not wo_id:
                continue
            try:
                rev = int(float(_as_str(col(row, "改訂番号")) or 0))   # '0.0' 表記
            except ValueError:
                rev = 0
            fg_id = _as_str(col(row, "応募者の ID"))
            key = (wo_id, fg_id, rev)
            wo = merged.get(key)
            if wo is None:
                wo = WorkOrder(
                    wo_id=wo_id, revision=rev,
                    staff_raw=_as_str(col(row, "応募者/スタッフ")),
                    staff_fg_id=fg_id,
                    start=_as_date(col(row, "作業オーダー開始日")),
                    # ⚠「最新の」終了日＝改訂共通。後で次改訂の開始日-1日に切り詰める
                    end=_as_date(col(row, "最新の作業オーダー終了日"))
                        or _as_date(col(row, "作業オーダー終了日")),
                    site=_as_str(col(row, "勤務地")),
                    business_unit=_as_str(col(row, "事業単位")),
                    job_title=_as_str(col(row, "求人情報タイトル")),
                    job_posting_id=_as_str(col(row, "求人情報 ID")),
                )
                merged[key] = wo
                entries[key] = _ual_entry(row, col)
            sup_name = (entries[key].get("supervisor") or {}).get("name", "")
            if sup_name and sup_name not in wo.supervisors:
                wo.supervisors.append(sup_name)
            cc = _as_str(col(row, "コストセンター"))
            if cc and cc not in wo.cost_centers:
                wo.cost_centers.append(cc)
    finally:
        wb.close()

    wos = sorted(merged.values(), key=lambda w: (w.staff_raw, w.wo_id, w.revision))
    _truncate_revision_ends(wos)

    # 詳細は最大改訂（＝現行値）のものを採用。旧 details JSON と同じ「現在のスナップショット」
    det_by_staff: dict[str, dict] = {}
    det_by_wo: dict[str, dict] = {}
    best_rev: dict[str, int] = {}
    for (wo_id, fg_id, rev), entry in entries.items():
        if fg_id and rev >= best_rev.get(f"s:{fg_id}", -1):
            best_rev[f"s:{fg_id}"] = rev
            det_by_staff[fg_id] = entry
        if rev >= best_rev.get(f"w:{wo_id}", -1):
            best_rev[f"w:{wo_id}"] = rev
            det_by_wo[wo_id] = entry
    return wos, det_by_staff, det_by_wo


def _ual_entry(row, col) -> dict:
    """1行ぶんの詳細列 → entry dict（空値はキーごと入れない）。"""
    e: dict = {}

    def put(key: str, val: str) -> None:
        if (val or "").strip():
            e[key] = val.strip()

    put("jobPostingId", _as_str(col(row, "求人情報 ID")))
    cr = _split5(col(row, "Client Side Responsible Person Info (Name, Title, Dept, Tel, Email)"))
    if not cr["name"]:      # DTS行は日本語列に入る
        cr = {"name": _as_str(col(row, "派遣先責任者の氏名")),
              "title": _as_str(col(row, "派遣先責任者の役職")),
              "department": _as_str(col(row, "派遣先責任者の部署")),
              "phone": _as_str(col(row, "派遣先責任者の電話番号")), "email": ""}
    if cr["name"]:
        e["clientResponsible"] = cr
    co = _split5(col(row, "Complaints Handling Representative Info (Name, Title, Dept, Tel, Email)"))
    if co["name"]:
        e["complaintRecipient"] = co
    sup = _split5(col(row, "Supervisor Info (Name, Title, Dept, Tel, Email)"))
    if sup["name"]:
        e["supervisor"] = sup

    put("responsibilityDegree", _as_str(col(row, "Level of Responsibility"))
        or _as_str(col(row, _DTS_FALLBACK["responsibilityDegree"])))
    put("businessContent", _as_str(col(row, "Description of Work (Please be as descriptive as possible)"))
        or _as_str(col(row, "業務内容")))
    put("workdaysHolidays", _as_str(col(row, "Typical Working days(new)"))
        or _as_str(col(row, _DTS_FALLBACK["workdaysHolidays"])))

    # シフト系の求人は Work Hours が空で、Shift Pattern / remarks / confirmation timing に入る
    notes = [_as_str(col(row, "Typical Working days remarks"))
             or _as_str(col(row, _DTS_FALLBACK["workdaysHolidaysNotes"]))]
    shift_pattern = _as_str(col(row, "Shift Pattern"))
    if shift_pattern:
        notes.append(f"シフト: {shift_pattern}")
    timing = _as_str(col(row, "Shift confirmation timing"))
    if timing:
        notes.append(f"シフト確定時期: {timing}")
    put("workdaysHolidaysNotes", "／".join(n for n in notes if n))

    put("workStart", _as_time(col(row, "Work Hours Start Time")))
    put("workEnd", _as_time(col(row, "Work Hours End Time")))
    put("breakStart", _as_time(col(row, "Break Start Time")))
    put("breakEnd", _as_time(col(row, "Break End Time")))

    div = _as_str(col(row, "Division Name UAL"))
    if div:
        unit, _, chief = div.partition(";")
        put("orgUnit", unit)
        put("orgChief", chief)
    else:
        put("orgUnit", _as_str(col(row, _DTS_FALLBACK["orgUnit"])))
        put("orgChief", _as_str(col(row, _DTS_FALLBACK["orgChief"])))

    tenure = _as_date(col(row, "Site tenure"))
    if tenure:
        e["siteTenure"] = f"{tenure.year}/{tenure.month}/{tenure.day}"

    loc = _as_str(col(row, "Work Location UAL"))
    if loc:
        parts = loc.split(";", 2)
        put("workOffice", parts[0] if parts else "")
        put("workAddress", parts[1] if len(parts) > 1 else "")
        put("conveniences", (parts[2] if len(parts) > 2 else "").replace(";", "、"))
    if "workAddress" not in e:
        put("workAddress", _as_str(col(row, "Site address")))

    agree = _as_str(col(row, "Workers subject to Labor Agreement or not"))
    if agree:
        e["agreementTarget"] = "該当する" if "労使協定" in agree else agree
    put("training", _as_str(col(row, "Matters about education and training to provide necessary capabilities for the work"))
        or _as_str(col(row, _DTS_FALLBACK["training"])))
    put("monthlyStdLower", _as_str(col(row, "Monthly ST : Lower Limit")))
    put("monthlyStdUpper", _as_str(col(row, "Monthly ST : Upper Limit")))
    return e


def _truncate_revision_ends(wos: list[WorkOrder]) -> None:
    """「最新の終了日」しか無いため、同じ (WO, スタッフ) の改訂列で
    各改訂の終了日を「次の改訂の開始日−1日」まで切り詰める（最終改訂はそのまま）。"""
    from collections import defaultdict

    groups: dict[tuple[str, str], list[WorkOrder]] = defaultdict(list)
    for w in wos:
        groups[(w.wo_id, w.staff_fg_id)].append(w)
    for revs in groups.values():
        revs.sort(key=lambda w: (w.start or dt.date.min, w.revision))
        for cur, nxt in zip(revs, revs[1:]):
            if cur.end and nxt.start:
                bound = nxt.start - dt.timedelta(days=1)
                if cur.end > bound:
                    cur.end = bound


# ---------------------------------------------------------------------------
# エリクソン CSV（直接契約マスタへの現行値上書き）
# ---------------------------------------------------------------------------

def _name_tokens(raw: str) -> frozenset[str]:
    """'Ohta, Takuya'・'MAHARJAN RAMITA' → 正規化トークン集合（長音 oh/oo/ou/uu を圧縮）。"""
    s = unicodedata.normalize("NFKC", raw or "").lower()
    tokens = [t for t in re.split(r"[,\s、，]+", s) if t]
    out = []
    for t in tokens:
        t = re.sub(r"oh(?=[bcdfghjklmnpqrstvwxyz])", "o", t)   # Ohta → ota
        t = t.replace("oo", "o").replace("ou", "o").replace("uu", "u")
        out.append(t)
    return frozenset(out)


def load_ericsson_report(path: Path | str) -> list[dict]:
    """エリクソンの新レポート → 1人1dict のリスト（DictReaderは重複ヘッダーで潰れるため不使用）。"""
    raw = Path(path).read_bytes()
    enc = "utf-8-sig" if raw.startswith(b"\xef\xbb\xbf") else "utf-8"
    try:
        text = raw.decode(enc)
    except UnicodeDecodeError:
        text = raw.decode("cp932")
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return []
    header = rows[0]
    idx: dict[str, int] = {}
    notes_idx: list[int] = []
    for i, h in enumerate(header):
        name = (h or "").strip()
        if name == "Notes on breaks, holidays or work hours":
            notes_idx.append(i)          # Notes だけは全列拾う（片側に変形労働時間制の記述）
        if name and name not in idx:     # ほかは最初の1個を採用（重複6組・値は同一）
            idx[name] = i

    out: list[dict] = []
    for r in rows[1:]:
        def g(name: str) -> str:
            i = idx.get(name)
            return _as_str(r[i]) if (i is not None and i < len(r)) else ""

        name = g("Job Seeker")
        if not name:
            continue
        notes = []
        for i in notes_idx:
            v = _as_str(r[i]) if i < len(r) else ""
            if v and v not in notes and v != "特に無し":
                notes.append(v)
        sup = g("Supervisor level 1")
        sup = re.sub(r"\([^)]*\)\s*$", "", sup).strip()      # 'Stephen Li(email)' → 'Stephen Li'
        out.append({
            "name": name,
            "tokens": _name_tokens(name),
            "client_responsible": {
                "name": g("Client Side Responsible Person"),
                "title": g("Client Side Responsible Person Position"),
                "department": g("Client Side Responsible Person Department"),
                "phone": g("Client Side Responsible Person Telephone Number"),
            },
            "complaint_title": g("Complaints Handling Representative Position"),
            "complaint_phone": g("Complaints Handling Representative Telephone Number"),
            "supervisor_name": sup,
            "responsibility": g("Level of Responsibility"),
            "work_start": _as_time(g("Work Hours Start Time")),
            "work_end": _as_time(g("Work Hours End Time")),
            "break_start": _as_time(g("Break Start Time")),
            "break_end": _as_time(g("Break End Time")),
            "workdays": g("Typical Working days"),
            "non_workdays": g("Typical Non-working days"),
            "age_60_plus": g("Is the workers age greater than or equal to 60 years old?"),
            "location": g("Worker Location (at client site)"),
            "notes": notes,
        })
    return out


def apply_ericsson_report(report_rows: list[dict], master_rows: list[dict],
                          source_name: str = "") -> tuple[int, list[str]]:
    """エリクソンの直接契約マスタ行へレポートの現行値を上書きする（行dictを直接書き換え）。

    - 突合キー: レポートの Job Seeker（ローマ字）と、マスタの 氏名 または 氏名カナ の
      正規化トークン集合の一致（氏名カナにはFG表記のローマ字を入れておく運用）
    - 就業時間・休憩・曜日・休日は**固定時間の契約のみ**上書き（マスタの就業時間に
      「シフト」を含む行は変形労働時間制＝マスタの記述を維持し、Notes を備考に足すだけ）
    - 苦情申出先は 役職・TEL のみ（氏名・部署はレポートに無い＝手入力を維持）
    - 各行に row["_エリクソン注記"] を残す（build 側で台帳の警告に載せる）

    戻り値: (上書きした行数, 突合できなかったマスタ行の氏名リスト)
    """
    updated = 0
    unmatched: list[str] = []
    label = f"エリクソンFGレポート（{source_name}）" if source_name else "エリクソンFGレポート"
    for row in master_rows:
        if "エリクソン" not in (row.get("派遣先名称") or ""):
            continue
        cand = _name_tokens(row.get("氏名", "")) | _name_tokens(row.get("氏名カナ", ""))
        hit = next((r for r in report_rows if r["tokens"] and r["tokens"] <= cand), None)
        if hit is None:
            unmatched.append(row.get("氏名", ""))
            row["_エリクソン注記"] = (f"{label}に氏名を突合できず → 手入力の値のまま"
                                  "（氏名カナ列にFG表記のローマ字を入れると突合できます）")
            continue

        changed: list[str] = []

        def put(col: str, val: str) -> None:
            val = (val or "").strip()
            if val and val != (row.get(col) or "").strip():
                row[col] = val
                changed.append(col)

        cr = hit["client_responsible"]
        put("派遣先責任者_氏名", cr["name"])
        put("派遣先責任者_役職", cr["title"])
        put("派遣先責任者_部署", cr["department"])
        put("派遣先責任者_TEL", cr["phone"])
        put("苦情申出先_役職", hit["complaint_title"])
        put("苦情申出先_TEL", hit["complaint_phone"])
        put("指揮命令者_氏名", hit["supervisor_name"])
        put("責任の程度", hit["responsibility"])
        is_shift = "シフト" in (row.get("就業時間") or "")
        if not is_shift:
            if hit["work_start"] and hit["work_end"]:
                put("就業時間", f"{hit['work_start']}～{hit['work_end']}")
            if hit["break_start"] and hit["break_end"]:
                put("休憩時間", f"{hit['break_start']}～{hit['break_end']}")
            put("就業曜日", hit["workdays"])
            put("休日", hit["non_workdays"])
        extra = "／".join(hit["notes"])
        if extra and extra not in (row.get("備考") or ""):
            row["備考"] = "／".join(x for x in ((row.get("備考") or "").strip(), extra) if x)
        if changed:
            updated += 1
            row["_エリクソン注記"] = (f"{label}の現行値で更新: {'・'.join(changed)}"
                                  + ("（シフト制のため就業時間・曜日は手入力を維持）" if is_shift else ""))
        else:
            row["_エリクソン注記"] = f"{label}と一致（更新なし）"
    return updated, unmatched
