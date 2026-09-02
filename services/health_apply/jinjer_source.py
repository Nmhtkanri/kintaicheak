# -*- coding: utf-8 -*-
"""健康診断申込: jinjer から対象者の氏名・社用メール・在籍区分・前年度の健診内容を取る（GET のみ）。

前年度内容の取り方（優先順）:
  1. カスタム項目 menu_id=15「健康診断履歴」（項目追加（横）型・年度ごとに1レコード）から
     年度 == 前年度 のレコード
  2. 無ければ menu_id=14「健康診断」（項目羅列型・1人1レコード）の現在値
  3. どちらも無ければ「なし」
どちらも `/v1/employees/custom-items` の1回の取得に含まれる（services.keiri_api の get_custom_items）。
選択肢ID→名前は `/v1/master/custom-menus`（get_custom_menus）。

レート制限はテナント共有なので逐次・50件チャンクで呼ぶ。並行させない。
"""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field

MENU_HISTORY = "15"
MENU_CURRENT = "14"
HISTORY_ITEM_YEAR = "年度"
HISTORY_ITEM_ORG = "健診機関"
HISTORY_ITEM_CONTENT = "健診内容"
HISTORY_ITEM_EXAM_DATE = "受診日"
CURRENT_ITEM_ORG = "検診機関"          # menu 14 は「検診」表記（jinjer 実測）
CURRENT_ITEM_DATE = "受診日"
CURRENT_ITEM_OPTIONS = "希望したオプション"

SOURCE_HISTORY = "履歴"
SOURCE_CURRENT = "現在値"
SOURCE_NONE = "なし"

_ID_SPLIT = re.compile(r"[,;、；\s]+")


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _dig(obj, *keys, default=""):
    cur = obj
    for k in keys:
        if not isinstance(cur, dict):
            return default
        cur = cur.get(k)
    return default if cur is None else cur


@dataclass
class EmployeeProfile:
    employee_id: str
    name: str
    email: str
    enrollment_id: str
    enrollment_name: str
    retirement_date: str = ""

    def as_dict(self) -> dict:
        return asdict(self)


@dataclass
class PreviousRaw:
    """jinjer から取ったままの前年度情報（コードに寄せる前）。"""
    source: str = SOURCE_NONE
    year: str = ""
    institution_text: str = ""
    content_labels: list[str] = field(default_factory=list)
    exam_date: str = ""
    notes: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return asdict(self)


# ----------------------------------------------------------------------
# 社員一覧
# ----------------------------------------------------------------------

def profiles_from_employees(employees: list[dict]) -> dict[str, EmployeeProfile]:
    out: dict[str, EmployeeProfile] = {}
    for emp in employees or []:
        if not isinstance(emp, dict):
            continue
        emp_id = _text(emp.get("id") or emp.get("employee_id"))
        if not emp_id:
            continue
        last = _text(_dig(emp, "company", "last_name"))
        first = _text(_dig(emp, "company", "first_name"))
        out[emp_id] = EmployeeProfile(
            employee_id=emp_id,
            name=f"{last} {first}".strip(),
            email=_text(_dig(emp, "company", "email")),
            enrollment_id=_text(_dig(emp, "company", "enrollment_classification", "id")),
            enrollment_name=_text(_dig(emp, "company", "enrollment_classification", "name")),
            retirement_date=_text(_dig(emp, "company", "retirement_date")),
        )
    return out


def fetch_profiles(client, employee_ids: list[str]) -> dict[str, EmployeeProfile]:
    """社員一覧を全員分（退職者含む）取り、必要な社員番号だけ返す。"""
    wanted = {str(i) for i in employee_ids}
    profiles = profiles_from_employees(client.get_employees(only_active=False))
    return {k: v for k, v in profiles.items() if k in wanted}


# ----------------------------------------------------------------------
# 選択肢ID → 名前
# ----------------------------------------------------------------------

def option_names_from_menus(menus: list[dict]) -> dict[str, dict[str, str]]:
    """/v1/master/custom-menus → {menu_id: {option_id: 選択肢名}}（選択肢を持つ項目だけ）。"""
    out: dict[str, dict[str, str]] = {}
    for menu in menus or []:
        if not isinstance(menu, dict):
            continue
        menu_id = _text(menu.get("id"))
        names = out.setdefault(menu_id, {})
        for item in menu.get("customize_item") or []:
            for opt in (item or {}).get("options") or []:
                oid = _text((opt or {}).get("id"))
                if oid:
                    names[oid] = _text(opt.get("name"))
    return out


def fetch_option_names(client) -> dict[str, dict[str, str]]:
    return option_names_from_menus(client.get_custom_menus())


# ----------------------------------------------------------------------
# 前年度の健診内容
# ----------------------------------------------------------------------

def _value_ids(value) -> list[str]:
    """チェックボックス項目の value（配列 or 文字列）を選択肢IDのリストにする。"""
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        ids = [_text(v) for v in value]
    else:
        ids = _ID_SPLIT.split(_text(value))
    return [i for i in ids if i]


def _labels(ids: list[str], names: dict[str, str], notes: list[str]) -> list[str]:
    labels = []
    for oid in ids:
        name = names.get(oid)
        if name:
            labels.append(name)
        else:
            labels.append(oid)
            notes.append(f"jinjer の選択肢ID {oid} の名前が分かりません")
    return labels


def _menus(person: dict | None) -> list[dict]:
    menus = (person or {}).get("customize_menu") or []
    if isinstance(menus, dict):
        menus = [menus]
    return [m for m in menus if isinstance(m, dict)]


def history_records(person: dict | None) -> list[dict[str, object]]:
    """menu 15 のレコード一覧（{項目名: 値} の dict）。customize_item[].id がレコード番号。"""
    grouped: dict[str, dict] = {}
    for menu in _menus(person):
        if _text(menu.get("id")) != MENU_HISTORY:
            continue
        for item in menu.get("customize_item") or []:
            rec_id = _text((item or {}).get("id"))
            if not rec_id:
                continue
            grouped.setdefault(rec_id, {})[_text(item.get("title"))] = item.get("value")
    return list(grouped.values())


def current_record(person: dict | None) -> dict[str, object] | None:
    """menu 14（1人1レコード）の {項目名: 値}。無ければ None。"""
    for menu in _menus(person):
        if _text(menu.get("id")) != MENU_CURRENT:
            continue
        rec: dict[str, object] = {}
        for item in menu.get("customize_item") or []:
            rec[_text((item or {}).get("title"))] = item.get("value")
        return rec
    return None


def _norm_date(value) -> str:
    s = _text(value)
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else s


def previous_from_custom_items(person: dict | None, previous_year: int,
                               option_names: dict[str, dict[str, str]]) -> PreviousRaw:
    notes: list[str] = []
    wanted = str(previous_year)
    records = [r for r in history_records(person) if _text(r.get(HISTORY_ITEM_YEAR)) == wanted]
    if records:
        if len(records) > 1:
            records.sort(key=lambda r: _norm_date(r.get(HISTORY_ITEM_EXAM_DATE)))
            notes.append(f"健康診断履歴に {wanted} 年度が {len(records)} 件あり、受診日が最新のものを使いました")
        rec = records[-1]
        labels = _labels(_value_ids(rec.get(HISTORY_ITEM_CONTENT)), option_names.get(MENU_HISTORY, {}), notes)
        return PreviousRaw(SOURCE_HISTORY, wanted, _text(rec.get(HISTORY_ITEM_ORG)), labels,
                           _norm_date(rec.get(HISTORY_ITEM_EXAM_DATE)), notes)
    cur = current_record(person)
    if cur and (_text(cur.get(CURRENT_ITEM_ORG)) or _value_ids(cur.get(CURRENT_ITEM_OPTIONS))):
        labels = _labels(_value_ids(cur.get(CURRENT_ITEM_OPTIONS)), option_names.get(MENU_CURRENT, {}), notes)
        notes.append(f"健康診断履歴に {wanted} 年度が無いため「健康診断」の現在値を使いました")
        return PreviousRaw(SOURCE_CURRENT, "", _text(cur.get(CURRENT_ITEM_ORG)), labels,
                           _norm_date(cur.get(CURRENT_ITEM_DATE)), notes)
    return PreviousRaw(SOURCE_NONE, "", "", [], "", notes)


def fetch_previous(client, employee_ids: list[str], previous_year: int,
                   option_names: dict[str, dict[str, str]]) -> dict[str, PreviousRaw]:
    ids = [str(i) for i in employee_ids]
    raw = client.get_custom_items(ids) if ids else {}
    return {emp: previous_from_custom_items(raw.get(emp), previous_year, option_names) for emp in ids}
