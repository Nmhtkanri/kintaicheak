# -*- coding: utf-8 -*-
"""jinjer API から退職者込みの人マスタを取る（読み取りのみ・書き込みなし）。

従業員一覧 xlsx は在籍者中心なので、過去四半期の台帳では退職者が氏名で見つからず
社員番号・生年月日が空になる。ここで `/v1/employees` を在籍区分フィルタなしで取り、
台帳に使う項目だけ（社員番号・氏名・生年月日・性別・雇用区分・入社日・在籍区分・退職日）を
`Z:\\派遣元管理台帳\\input\\jinjer_api_roster.json` にキャッシュして、xlsx の人マスタに合流させる。

クライアントは本アプリの services/jinjer_api_client.JinjerClient を使う（2026-08-28 ハブ移設）。
"""
from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path

from .config import INPUT_DIR
from .roster import Person, Roster

CACHE_PATH = INPUT_DIR / "jinjer_api_roster.json"
_ENROLLMENT = {"0": "在籍", "1": "退職", "2": "休職"}


def _dig(record, *paths):
    """'company.joined_on' 形式のパス候補を順に試し、最初に見つかった非空値を返す。"""
    for path in paths:
        cur = record
        ok = True
        for key in path.split("."):
            if isinstance(cur, dict) and key in cur:
                cur = cur[key]
            else:
                ok = False
                break
        if ok and cur not in (None, ""):
            return cur
    return ""


def _as_text(value) -> str:
    if isinstance(value, dict):
        return str(value.get("name") or value.get("id") or "")
    return str(value or "")


def _to_date(value) -> dt.date | None:
    s = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def fetch_employees_raw() -> list[dict]:
    """在籍区分フィルタなし（＝退職者・休職者含む全件）で従業員一覧を取る。"""
    from services.jinjer_api_client import JinjerClient

    client = JinjerClient()
    client.authenticate()
    return client.get_employees(only_active=False)


def _to_person(emp: dict) -> Person | None:
    emp_id = str(emp.get("id", "")).strip()
    if not re.match(r"^20\d{2}", emp_id):
        return None   # 20YY始まりのみ自社社員（5/6/9始まりは派遣・テストで対象外）
    enrollment = _as_text(_dig(emp, "company.enrollment_classification", "enrollment_classification"))
    status = _ENROLLMENT.get(enrollment, enrollment)
    return Person(
        emp_id=emp_id,
        sei=_as_text(_dig(emp, "company.last_name", "last_name")).strip(),
        mei=_as_text(_dig(emp, "company.first_name", "first_name")).strip(),
        birth=_to_date(_dig(emp, "personal.date_of_birth", "personal.birthday", "date_of_birth")),
        sex=_as_text(_dig(emp, "personal.gender", "gender")).replace("性", ""),
        employment_type=_as_text(_dig(emp, "company.employment_classification", "employment_classification")),
        joined=_to_date(_dig(emp, "company.joined_on", "joined_on")),
        status=status,
        retired=_to_date(_dig(emp, "company.retirement_date", "company.resigned_on", "retired_on")),
    )


def _person_to_dict(p: Person) -> dict:
    return {
        "emp_id": p.emp_id, "sei": p.sei, "mei": p.mei,
        "birth": p.birth.isoformat() if p.birth else "",
        "sex": p.sex, "employment_type": p.employment_type,
        "joined": p.joined.isoformat() if p.joined else "",
        "status": p.status,
        "retired": p.retired.isoformat() if p.retired else "",
    }


def _person_from_dict(d: dict) -> Person:
    return Person(
        emp_id=d.get("emp_id", ""), sei=d.get("sei", ""), mei=d.get("mei", ""),
        birth=_to_date(d.get("birth")), sex=d.get("sex", ""),
        employment_type=d.get("employment_type", ""),
        joined=_to_date(d.get("joined")), status=d.get("status", ""),
        retired=_to_date(d.get("retired")),
    )


def refresh_cache(cache_path: Path | str = CACHE_PATH) -> list[Person]:
    """APIから取得して台帳用の最小項目だけキャッシュする（個人情報＝リポジトリ外に置く）。"""
    raw = fetch_employees_raw()
    people = [p for p in (_to_person(e) for e in raw) if p is not None]
    payload = {
        "fetched_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "note": "daicho用 人マスタキャッシュ（退職者込み・台帳に使う項目のみ）",
        "count": len(people),
        "people": [_person_to_dict(p) for p in people],
    }
    path = Path(cache_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    return people


def load_cache(cache_path: Path | str = CACHE_PATH) -> tuple[list[Person] | None, str]:
    path = Path(cache_path)
    if not path.exists():
        return None, ""
    payload = json.loads(path.read_text(encoding="utf-8"))
    people = [_person_from_dict(d) for d in payload.get("people", [])]
    return people, payload.get("fetched_at", "")


def merge_into_roster(roster: Roster, people: list[Person]) -> str:
    """API の人を xlsx 由来の Roster に合流させる。

    - xlsx に居ない人（退職者など）は追加
    - 居る人は空いている項目（生年月日・性別・雇用区分・入社日・在籍区分・退職日）だけ埋める
      （氏名・勤務先・所属グループは xlsx 側を優先）
    """
    added = filled = 0
    for p in people:
        cur = roster.people.get(p.emp_id)
        if cur is None:
            roster.people[p.emp_id] = p
            added += 1
            continue
        before = (cur.birth, cur.sex, cur.employment_type, cur.joined, cur.status, cur.retired)
        cur.birth = cur.birth or p.birth
        cur.sex = cur.sex or p.sex
        cur.employment_type = cur.employment_type or p.employment_type
        cur.joined = cur.joined or p.joined
        cur.status = cur.status or p.status
        cur.retired = cur.retired if cur.retired is not None else p.retired
        if before != (cur.birth, cur.sex, cur.employment_type, cur.joined, cur.status, cur.retired):
            filled += 1
    # 氏名索引を作り直す
    roster.by_name.clear()
    for p in roster.people.values():
        roster.by_name.setdefault(p.key, []).append(p)
    return f"jinjer API 人マスタ: {len(people)}人を合流（追加 {added}・補完 {filled}）"
