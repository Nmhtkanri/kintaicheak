# -*- coding: utf-8 -*-
"""健康診断HPM: 健診申込の「選択肢」シート（Google）の機関・種別・追加検査を画面へ追記する。

HPM取込用CSVの健診機関・健診種別は、これまで変換マスタ（Excel）の登録分だけだった。
2026-09-09 から、健診申込モードが正本にしている Google「選択肢」シートの内容も
プルダウンへ **追記** する（谷津さん依頼）。

約束事:
- 変換マスタが正。同じ場所コード（または同じ機関名）の機関はシート側を載せない
  （同友会のようにコース名テキストを HPM 出力値にしている設定を守るため）
- 「その他」（OTHER）は場所コードが無いので載せない
- 種別（10〜15）は全機関の既存コースの後ろに追記する。HPM出力値が同じものは重複させない
- シートの「有効」は申込画面で選ばせるかどうかの旗。健診結果は無効機関からも届くので
  HPM側では有効・無効とも選べるようにし、無効は表示名で分かるようにする
- 追加検査（婦人科検診）は 302列に該当する列が無い。CSVには書かず、画面の既定
  （女性はON）とログ・警告だけに残す（暫定。HPM側の列が決まったら書く場所を足す）
- シートが読めなくても HPM モードは止めない（変換マスタだけで動き、理由を画面に出す）
"""

from __future__ import annotations

from dataclasses import dataclass, field

from services.health_hpm_excel import Issue
from services.health_hpm_master import (
    Course,
    HpmMaster,
    Institution,
    courses_of,
    find_course,
    resolve_institution,
)

SOURCE_MASTER = "master"
SOURCE_SHEET = "sheet"

# 選択肢シートに載っていない・読めないときでも出す追加検査（暫定の既定）
FALLBACK_EXTRA_CODE = "GYN"
FALLBACK_EXTRA_NAME = "婦人科検診"

# 場所コードが無い予約語（健診申込の「その他」）
OTHER_INSTITUTION_CODE = "OTHER"

INACTIVE_SUFFIX = "（申込では無効）"
FEMALE = "女性"


@dataclass(frozen=True)
class SheetInstitution:
    code: str
    name: str
    active: bool = True
    order: int = 9999
    note: str = ""


@dataclass(frozen=True)
class SheetExamType:
    code: str
    name: str
    active: bool = True
    order: int = 9999


@dataclass(frozen=True)
class SheetExtra:
    code: str
    name: str


@dataclass
class SheetOptions:
    """選択肢シートから取った分。error が空でなければ読めなかった（fallback）。"""

    institutions: list[SheetInstitution] = field(default_factory=list)
    exam_types: list[SheetExamType] = field(default_factory=list)
    extras: list[SheetExtra] = field(default_factory=list)
    source: str = ""
    error: str = ""

    @property
    def loaded(self) -> bool:
        return not self.error

    def institution_by_code(self, code: str) -> SheetInstitution | None:
        code = str(code or "").strip()
        if not code:
            return None
        for inst in self.institutions:
            if inst.code == code:
                return inst
        return None

    def institution_by_name(self, name: str) -> SheetInstitution | None:
        name = str(name or "").strip()
        if not name:
            return None
        for inst in self.institutions:
            if inst.name == name:
                return inst
        return None

    def exam_type_by_code(self, code: str) -> SheetExamType | None:
        code = str(code or "").strip()
        for et in self.exam_types:
            if et.code == code:
                return et
        return None

    def extra_by_code(self, code: str) -> SheetExtra | None:
        code = str(code or "").strip()
        for ex in self.extras:
            if ex.code == code:
                return ex
        return None

    def as_dict(self) -> dict:
        return {
            "loaded": self.loaded,
            "source": self.source,
            "error": self.error,
            "counts": {"institutions": len(self.institutions),
                       "exam_types": len(self.exam_types),
                       "extras": len(self.extras)},
            "extras": [{"code": e.code, "name": e.name} for e in self.extras],
        }


def fallback_options(reason: str) -> SheetOptions:
    """シートが読めないときの分。追加検査だけは暫定の既定（婦人科検診）を持たせる。"""
    return SheetOptions(
        extras=[SheetExtra(FALLBACK_EXTRA_CODE, FALLBACK_EXTRA_NAME)],
        error=str(reason or "選択肢シートを読めませんでした"),
    )


def from_catalog(catalog, source: str = "") -> SheetOptions:
    """health_apply.options.OptionCatalog → SheetOptions。

    OptionCatalog の区分名（機関／種別／追加検査）に依存するが、ここで一度だけ読み替えて
    HPM 側の残りのコードは OptionCatalog を知らずに済むようにする。
    """
    from services.health_apply.options import (
        KIND_EXAM_TYPE,
        KIND_EXTRA,
        KIND_INSTITUTION,
        normalize_extra_name,
    )

    institutions = [
        SheetInstitution(code=o.code, name=o.name, active=o.active, order=o.order, note=o.note)
        for o in catalog.of_kind(KIND_INSTITUTION, active_only=False)
        if o.code != OTHER_INSTITUTION_CODE
    ]
    exam_types = [
        SheetExamType(code=o.code, name=o.name, active=o.active, order=o.order)
        for o in catalog.of_kind(KIND_EXAM_TYPE, active_only=False)
    ]
    extras = [
        SheetExtra(code=o.code, name=normalize_extra_name(o.name))
        for o in catalog.of_kind(KIND_EXTRA, active_only=True)
    ]
    if not extras:
        extras = [SheetExtra(FALLBACK_EXTRA_CODE, FALLBACK_EXTRA_NAME)]

    # 有効を先、無効を後ろ。同じ群の中は並び順
    institutions.sort(key=lambda i: (not i.active, i.order, i.name))
    exam_types.sort(key=lambda e: (not e.active, e.order, e.name))
    return SheetOptions(institutions=institutions, exam_types=exam_types,
                        extras=extras, source=source)


# ---------------------------------------------------------------------------
# 画面用の合成
# ---------------------------------------------------------------------------

def _generic_courses(options: SheetOptions, existing_values: set[str]) -> list[dict]:
    out = []
    for et in options.exam_types:
        if et.code in existing_values:
            continue
        out.append({"display_name": et.name + ("" if et.active else INACTIVE_SUFFIX),
                    "hpm_value": et.code, "source": SOURCE_SHEET})
    return out


def merged_institutions(master: HpmMaster, options: SheetOptions) -> list[dict]:
    """プルダウン用。変換マスタ → 選択肢シート（重複除外）の順。

    値（value）は変換マスタが機関名、シートが場所コード。generate 側の
    resolve_institution_choice がその順で引く。
    """
    out: list[dict] = []
    master_codes: set[str] = set()
    master_names: set[str] = set()
    for name in sorted(master.institutions):
        inst = master.institutions[name]
        master_codes.add(inst.location_code)
        master_names.add(inst.name)
        courses = [{"display_name": c.display_name, "hpm_value": c.hpm_value,
                    "source": SOURCE_MASTER} for c in courses_of(master, inst.name)]
        courses += _generic_courses(options, {c["hpm_value"] for c in courses})
        out.append({
            "value": inst.name,
            "name": inst.name,
            "location_code": inst.location_code,
            "hpm_confirmed": inst.hpm_confirmed,
            "note": inst.note,
            "source": SOURCE_MASTER,
            "active": True,
            "courses": courses,
        })

    for sheet_inst in options.institutions:
        if sheet_inst.code in master_codes or sheet_inst.name in master_names:
            continue
        out.append({
            "value": sheet_inst.code,
            "name": sheet_inst.name + ("" if sheet_inst.active else INACTIVE_SUFFIX),
            "location_code": sheet_inst.code,
            "hpm_confirmed": True,   # hpm.txt 由来のコードなのでHPM側にある前提
            "note": sheet_inst.note,
            "source": SOURCE_SHEET,
            "active": sheet_inst.active,
            "courses": _generic_courses(options, set()),
        })
    return out


# ---------------------------------------------------------------------------
# 画面から戻った選択値の解決（generate 側）
# ---------------------------------------------------------------------------

def resolve_institution_choice(master: HpmMaster, options: SheetOptions,
                               value: str) -> Institution | None:
    """変換マスタ（機関名・別名）→ シート（場所コード）→ シート（機関名）の順で引く。"""
    value = str(value or "").strip()
    if not value:
        return None
    inst = resolve_institution(master, value)
    if inst is not None:
        return inst
    sheet_inst = options.institution_by_code(value) or options.institution_by_name(
        value[:-len(INACTIVE_SUFFIX)] if value.endswith(INACTIVE_SUFFIX) else value)
    if sheet_inst is None:
        return None
    return Institution(
        name=sheet_inst.name,
        location_code=sheet_inst.code,
        hpm_confirmed=True,
        note=f"健診申込の選択肢シート由来{'' if sheet_inst.active else '（申込では無効）'}",
    )


def resolve_course_choice(master: HpmMaster, options: SheetOptions,
                          institution: Institution, hpm_value: str) -> Course | None:
    """変換マスタのコース → シートの種別（コード一致）の順で引く。"""
    hpm_value = str(hpm_value or "").strip()
    if not hpm_value:
        return None
    if institution.name in master.institutions:
        course = find_course(master, institution.name, hpm_value)
        if course is not None:
            return course
    et = options.exam_type_by_code(hpm_value)
    if et is None:
        return None
    return Course(institution=institution.name, display_name=et.name, hpm_value=et.code)


def resolve_extras(options: SheetOptions, codes) -> tuple[list[SheetExtra], list[str]]:
    """画面から戻った追加検査コードを検証する。(有効な分, 不明なコード)。"""
    found: list[SheetExtra] = []
    unknown: list[str] = []
    for raw in (codes or []):
        code = str(raw or "").strip()
        if not code:
            continue
        ex = options.extra_by_code(code)
        if ex is None:
            unknown.append(code)
        elif ex not in found:
            found.append(ex)
    return found, unknown


def default_extra_codes(options: SheetOptions, *genders: str) -> list[str]:
    """暫定の既定: いずれかの性別が女性なら、追加検査を全部ON（今は婦人科検診だけ）。

    画面（script.js の hhIsFemale）と同じ規則。サーバー側はテストと将来の再利用のために持つ。
    """
    if any(str(g or "").strip() == FEMALE for g in genders):
        return [e.code for e in options.extras]
    return []


def extras_issue(selected: list[tuple[str, list[SheetExtra]]]) -> Issue | None:
    """追加検査がある人を警告にまとめる（CSVには出していないことを明記）。"""
    lines = []
    for person_name, extras in selected:
        if extras:
            lines.append(f"{person_name}（{'・'.join(e.name for e in extras)}）")
    if not lines:
        return None
    return Issue(
        "warning", "EXTRA_NOT_IN_CSV",
        "追加検査は HPM の302列に該当する列が無いため CSV には出していません（暫定）: "
        + "、".join(lines),
    )
