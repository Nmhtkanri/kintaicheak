"""健診受診者と jinjer 社員の照合（純ロジック・API呼び出しなし）

HPMへ出す氏名・カナ・生年月日・性別は、健診結果Excelの表記ではなく
**jinjer の登録内容**を正とする。健診機関ごとに氏名表記が揺れるため。

自動で確定するのは「氏名が一意に一致 かつ 性別一致 かつ 受診日時点の満年齢が
Excelの年齢と一致」のときだけ。ひとつでも外れたら人に選ばせる。
同姓同名を先勝ちで確定して別人のIDを書いた事故（スケジュール取込・2026-07）が
あるため、迷ったら自動確定しない側に倒す。
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import date, datetime

from services.matcher import normalize_name

logger = logging.getLogger(__name__)

# 自社社員の社員番号は 20YY 始まりの7桁。5/6/9 始まり（派遣・テスト番号）は
# 健診の転記対象にしない（kotsuhi_seisa.is_company_employee と同じ考え方）。
COMPANY_EMPLOYEE_RE = re.compile(r"20\d{5}")

STATUS_OK = "ok"
STATUS_SELECT = "select"

# 氏名の連結。HPMの実CSVは漢字氏名が全角スペース、カナ氏名が半角スペース区切り。
KANJI_SEPARATOR = "　"
KANA_SEPARATOR = " "


@dataclass(frozen=True)
class JinjerCandidate:
    employee_id: str
    last_name: str = ""
    first_name: str = ""
    last_kana: str = ""
    first_kana: str = ""
    birth_date: date | None = None
    gender: str = ""  # "男性" / "女性"

    @property
    def name(self) -> str:
        return KANJI_SEPARATOR.join(p for p in (self.last_name, self.first_name) if p)

    @property
    def kana(self) -> str:
        return KANA_SEPARATOR.join(p for p in (self.last_kana, self.first_kana) if p)

    def as_dict(self) -> dict:
        """画面に渡す形。"""
        return {
            "employee_id": self.employee_id,
            "name": self.name,
            "kana": self.kana,
            "birth_date": self.birth_date.isoformat() if self.birth_date else "",
            "gender": self.gender,
        }


@dataclass
class MatchResult:
    status: str = STATUS_SELECT
    employee_id: str = ""
    candidates: list[JinjerCandidate] = field(default_factory=list)
    reasons: list[str] = field(default_factory=list)

    @property
    def is_ok(self) -> bool:
        return self.status == STATUS_OK and bool(self.employee_id)


def _safe_get(data, *keys):
    """jinjer のレスポンスは階層が深く null も混ざるので、途中で切れても落ちないように取る。"""
    current = data
    for key in keys:
        if not isinstance(current, dict):
            return None
        current = current.get(key)
        if current is None:
            return None
    return current


def _parse_birth(value) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip().split("T")[0].split(" ")[0]
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def is_company_employee(employee_id) -> bool:
    return bool(COMPANY_EMPLOYEE_RE.fullmatch(str(employee_id or "").strip()))


def build_candidates(employees: list[dict]) -> list[JinjerCandidate]:
    """jinjer の get_employees() の生dictから照合候補を作る。

    自社形式でない社員番号は捨てる（派遣・テスト番号にHPMの結果を紐付けない）。
    """
    out: list[JinjerCandidate] = []
    for emp in employees or []:
        employee_id = str(_safe_get(emp, "id") or "").strip()
        if not is_company_employee(employee_id):
            continue
        out.append(JinjerCandidate(
            employee_id=employee_id,
            last_name=str(_safe_get(emp, "company", "last_name") or "").strip(),
            first_name=str(_safe_get(emp, "company", "first_name") or "").strip(),
            last_kana=str(_safe_get(emp, "company", "last_name_phonetic") or "").strip(),
            first_kana=str(_safe_get(emp, "company", "first_name_phonetic") or "").strip(),
            birth_date=_parse_birth(_safe_get(emp, "personal", "date_of_birth")),
            gender=str(_safe_get(emp, "personal", "gender", "name") or "").strip(),
        ))
    return out


# 健診の原票は戸籍どおりの異体字で印字されることがあり、jinjer は常用字体で
# 登録されている（2026-08-10 実例: 原票「髙橋」× jinjer「高橋」）。字が違うだけで
# 手動選択に落ちるのは手間なので、**照合キーを作るときだけ**常用字体へ寄せる。
# 画面表示とCSV出力はそれぞれの登録どおりのまま（原票は原票、氏名はjinjer）。
# 全社共通の services/matcher.py は他モードにも効くので触らない。
_ITAIJI = str.maketrans({
    "髙": "高", "﨑": "崎", "濵": "浜", "濱": "浜",
    "邊": "辺", "邉": "辺", "齋": "斎", "齊": "斉",
    "冨": "富", "廣": "広", "德": "徳", "橫": "横",
})


def fold_itaiji(text: str) -> str:
    return str(text or "").translate(_ITAIJI)


def match_key(text: str) -> str:
    """氏名の照合キー。空白除去などの共通正規化に、異体字寄せを重ねる。"""
    return normalize_name(fold_itaiji(text))


def age_at(birth: date | None, on: date | None) -> int | None:
    """受診日時点の満年齢。誕生日が来ていなければ1つ引く。"""
    if birth is None or on is None:
        return None
    years = on.year - birth.year
    if (on.month, on.day) < (birth.month, birth.day):
        years -= 1
    return years


def gender_matches(excel_gender: str, jinjer_gender: str) -> bool:
    """「男性」と「男」を同じものとして扱う。どちらか空なら判断しない（False）。"""
    a = (excel_gender or "").strip()
    b = (jinjer_gender or "").strip()
    if not a or not b:
        return False
    return a[0] == b[0]


def gender_to_hpm(jinjer_gender: str) -> str:
    """HPMの性別欄は「男」「女」の1文字。"""
    text = (jinjer_gender or "").strip()
    if text.startswith("男"):
        return "男"
    if text.startswith("女"):
        return "女"
    return ""


def match_person(name: str, gender: str, age: int | None, exam_date: date | None,
                 candidates: list[JinjerCandidate]) -> MatchResult:
    """1名分の照合。全部そろったときだけ ok、それ以外は人に選ばせる。"""
    key = match_key(name)
    hits = [c for c in candidates if match_key(c.name) == key]

    if not hits:
        # 姓だけで一致する人も候補には出す（選ぶのは人）
        loose = [c for c in candidates
                 if key and (match_key(c.last_name) == key
                             or key in match_key(c.name))]
        return MatchResult(
            status=STATUS_SELECT,
            candidates=loose,
            reasons=[f"jinjerに「{name}」と一致する社員が見つかりません"],
        )

    if len(hits) > 1:
        return MatchResult(
            status=STATUS_SELECT,
            candidates=hits,
            reasons=[f"同姓同名が{len(hits)}名います（社員番号: "
                     f"{'、'.join(c.employee_id for c in hits)}）。どちらか選んでください"],
        )

    candidate = hits[0]
    reasons: list[str] = []

    if not gender_matches(gender, candidate.gender):
        reasons.append(
            f"性別が一致しません（健診結果: {gender or '空欄'} / "
            f"jinjer: {candidate.gender or '空欄'}）"
        )

    jinjer_age = age_at(candidate.birth_date, exam_date)
    if candidate.birth_date is None:
        reasons.append("jinjerに生年月日が登録されていません")
    elif age is None:
        reasons.append("健診結果に年齢がありません")
    elif jinjer_age != age:
        reasons.append(
            f"受診日時点の年齢が一致しません（健診結果: {age}歳 / "
            f"生年月日から: {jinjer_age}歳）"
        )

    if reasons:
        return MatchResult(status=STATUS_SELECT, candidates=hits, reasons=reasons)

    return MatchResult(status=STATUS_OK, employee_id=candidate.employee_id,
                       candidates=hits)


def find_candidate(employee_id: str, candidates: list[JinjerCandidate]) -> JinjerCandidate | None:
    employee_id = str(employee_id or "").strip()
    for candidate in candidates:
        if candidate.employee_id == employee_id:
            return candidate
    return None


def validate_selection(employee_id: str, candidates: list[JinjerCandidate]) -> JinjerCandidate:
    """CSV生成時の再検証。画面から来たIDを信用せず在籍一覧に当てる。"""
    employee_id = str(employee_id or "").strip()
    if not employee_id:
        raise ValueError("社員番号が選ばれていません")
    if not is_company_employee(employee_id):
        raise ValueError(
            f"社員番号 {employee_id} は自社の形式（20YY＋3桁）ではありません"
        )
    candidate = find_candidate(employee_id, candidates)
    if candidate is None:
        raise ValueError(
            f"社員番号 {employee_id} はjinjerの在籍者一覧にありません"
        )
    return candidate
