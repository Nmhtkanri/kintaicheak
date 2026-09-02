# -*- coding: utf-8 -*-
"""健康診断申込: 「対象者」「回答」シートを突き合わせて集計・検証する（純粋関数）。

Google からもらった行（dict）だけを入力にし、ネットワークにもファイルにも触らない。
判定の考え方:
- 取込不可（error）は「このままでは次工程（jinjer登録）に使えない」行。
  重複回答／対象外社員／年度不一致／選択肢不明／対象者と回答の状態不一致／
  「前年度と同じ」なのに前年度情報が無い、が該当。
- 最新回答は **回答版の最大** で決める（回答日時では決めない。時計ずれや手入力に強くする）。
- 「初回アクセス日時」はメールのセキュリティ検査でも付くので、本人が見た証拠にはしない。
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

from services.health_apply import schema as S
from services.health_apply.options import (
    GYN_CODE, KIND_EXAM_TYPE, KIND_EXTRA, KIND_INSTITUTION, KIND_RELATIONSHIP,
    OTHER_INSTITUTION_CODE, OptionCatalog, normalize_extra_name,
)

BUCKETS = ("unsent", "sent_not_accessed", "accessed_only", "answered",
           "reanswer_pending", "invalid", "error")
BUCKET_LABELS = {
    "unsent": "未送信",
    "sent_not_accessed": "送信済み・未アクセス",
    "accessed_only": "URLが開かれた記録あり・未回答",
    "answered": "回答済み",
    "reanswer_pending": "再回答待ち",
    "invalid": "無効",
    "error": "取込不可",
}
ENROLLMENT_LABELS = {"0": "在籍", "1": "退職", "2": "休職"}

_EXTRA_SEPARATORS = (";", "；", "、", ",")


@dataclass
class Issue:
    level: str      # "error" | "warning"
    code: str
    message: str

    def as_dict(self) -> dict:
        return {"level": self.level, "code": self.code, "message": self.message}


@dataclass
class ResponseView:
    employee_id: str
    name: str
    bucket: str = "unsent"
    target: dict = field(default_factory=dict)
    latest: dict | None = None
    history_count: int = 0
    issues: list[Issue] = field(default_factory=list)

    @property
    def importable(self) -> bool:
        return not any(i.level == "error" for i in self.issues)

    def as_dict(self) -> dict:
        return {
            "employee_id": self.employee_id,
            "name": self.name,
            "bucket": self.bucket,
            "bucket_label": BUCKET_LABELS.get(self.bucket, self.bucket),
            "importable": self.importable,
            "target": self.target,
            "latest": self.latest,
            "history_count": self.history_count,
            "issues": [i.as_dict() for i in self.issues],
        }


def split_codes(text: str | None) -> list[str]:
    s = str(text or "")
    for sep in _EXTRA_SEPARATORS[1:]:
        s = s.replace(sep, _EXTRA_SEPARATORS[0])
    return [c.strip() for c in s.split(_EXTRA_SEPARATORS[0]) if c.strip()]


def _int_or_none(value: str) -> int | None:
    try:
        return int(str(value or "").strip())
    except ValueError:
        return None


# ----------------------------------------------------------------------
# 対象者
# ----------------------------------------------------------------------

def index_targets(target_rows: list[dict], fiscal_year: int) -> tuple[dict[str, dict], list[Issue]]:
    """{社員番号: 行} と、シート全体の指摘（別年度の行・社員番号重複）を返す。"""
    index: dict[str, dict] = {}
    issues: list[Issue] = []
    other_years = 0
    for row in target_rows:
        year = str(row.get("年度", "")).strip()
        emp = str(row.get("社員番号", "")).strip()
        if year != str(fiscal_year):
            other_years += 1
            continue
        if not emp:
            issues.append(Issue("error", "target_blank_id", "対象者シートに社員番号が空の行があります"))
            continue
        if emp in index:
            issues.append(Issue("error", "target_duplicate",
                                f"対象者シートに社員番号 {emp} が2行以上あります（先の行だけを使います）"))
            continue
        index[emp] = row
    if other_years:
        issues.append(Issue("warning", "target_other_year",
                            f"対象者シートに {fiscal_year} 年度以外の行が {other_years} 行あります（無視します）"))
    return index, issues


def _option_payload(catalog: OptionCatalog, kind: str, code: str, raw_name: str = "") -> dict:
    opt = catalog.lookup(kind, code)
    return {
        "code": code,
        "name": opt.name if opt else (raw_name or catalog.display(kind, code)),
        "raw_name": raw_name,
        "known": opt is not None,
        "active": bool(opt and opt.active),
    }


def target_payload(row: dict, catalog: OptionCatalog) -> dict:
    """画面に出す対象者行（Apps Script 管轄の列も含む）。"""
    enrollment = str(row.get("在籍区分", "")).strip()
    extras = [_option_payload(catalog, KIND_EXTRA, c) for c in split_codes(row.get("前年度追加検査"))]
    return {
        "email": row.get("社用メール", ""),
        "enrollment": enrollment,
        "enrollment_label": ENROLLMENT_LABELS.get(enrollment, enrollment),
        "previous": {
            "source": row.get("前年度情報元", "") or S.SOURCE_NONE,
            "institution": _option_payload(catalog, KIND_INSTITUTION, row.get("前年度健診機関コード", ""),
                                           row.get("前年度健診機関名", "")),
            "exam_type": _option_payload(catalog, KIND_EXAM_TYPE, row.get("前年度健診種別コード", ""),
                                         row.get("前年度健診種別名", "")),
            "extras": extras,
            "raw": row.get("前年度健診機関(原文)", ""),
        },
        "registered_at": row.get("登録日時", ""),
        "registered_by": row.get("登録者", ""),
        "sent_at": row.get("送信日時", ""),
        "sent_count": row.get("送信回数", ""),
        "first_access_at": row.get("初回アクセス日時", ""),
        "status": row.get("申込状態", "") or S.STATUS_UNSENT,
        "receipt": row.get("受付番号", ""),
        "version": row.get("回答版", ""),
        "answered_at": row.get("回答日時", ""),
        "note": row.get("備考", ""),
    }


# ----------------------------------------------------------------------
# 回答
# ----------------------------------------------------------------------

def parse_response(row: dict, catalog: OptionCatalog) -> tuple[dict, list[Issue]]:
    """回答1行を画面用 dict にし、その行だけで分かる指摘を返す。"""
    issues: list[Issue] = []
    kind = str(row.get("申込区分", "")).strip()
    if kind not in (S.KIND_SAME, S.KIND_CHANGE):
        issues.append(Issue("error", "bad_kind", f"申込区分が想定外です: 「{kind}」"))

    inst_code = str(row.get("健診機関コード", "")).strip()
    inst = _option_payload(catalog, KIND_INSTITUTION, inst_code, row.get("健診機関名", ""))
    if not inst["known"]:
        issues.append(Issue("error", "unknown_institution", f"健診機関コードが選択肢にありません: 「{inst_code or '空'}」"))
    elif not inst["active"]:
        issues.append(Issue("warning", "inactive_institution", f"健診機関「{inst['name']}」は選択肢で無効になっています"))
    elif inst["raw_name"] and inst["raw_name"] != inst["name"]:
        issues.append(Issue("warning", "institution_name_differs",
                            f"回答の健診機関名「{inst['raw_name']}」が選択肢の表示名「{inst['name']}」と違います"))
    other_name = str(row.get("その他医療機関名", "")).strip()
    if inst_code == OTHER_INSTITUTION_CODE and not other_name:
        issues.append(Issue("warning", "other_without_name", "「その他」なのに医療機関名が空です"))

    type_code = str(row.get("健診種別コード", "")).strip()
    exam_type = _option_payload(catalog, KIND_EXAM_TYPE, type_code, row.get("健診種別名", ""))
    if not exam_type["known"]:
        issues.append(Issue("error", "unknown_exam_type", f"健診種別コードが選択肢にありません: 「{type_code or '空'}」"))
    elif not exam_type["active"]:
        issues.append(Issue("warning", "inactive_exam_type", f"健診種別「{exam_type['name']}」は選択肢で無効になっています"))

    extras = []
    for code in split_codes(row.get("追加検査")):
        opt = catalog.lookup(KIND_EXTRA, code) or catalog.resolve_name(KIND_EXTRA, normalize_extra_name(code))
        if opt is None:
            issues.append(Issue("error", "unknown_extra", f"追加検査が選択肢にありません: 「{code}」"))
            extras.append({"code": code, "name": code, "known": False, "active": False})
        else:
            extras.append({"code": opt.code, "name": opt.name, "known": True, "active": opt.active})

    dependent_requested = str(row.get("被扶養者申込", "")).strip() in ("1", "TRUE", "true", "希望する")
    relationship = str(row.get("続柄", "")).strip()
    dependent_name = str(row.get("被扶養者氏名", "")).strip()
    if dependent_requested:
        if not relationship:
            issues.append(Issue("warning", "dependent_without_relationship", "被扶養者を希望しているのに続柄が空です"))
        elif catalog.lookup(KIND_RELATIONSHIP, relationship) is None \
                and catalog.resolve_name(KIND_RELATIONSHIP, relationship) is None:
            issues.append(Issue("error", "unknown_relationship", f"続柄が選択肢にありません: 「{relationship}」"))
        if not dependent_name:
            issues.append(Issue("warning", "dependent_without_name", "被扶養者を希望しているのに氏名が空です"))

    latest = {
        "answered_at": row.get("回答日時", ""),
        "receipt": str(row.get("受付番号", "")).strip(),
        "version": str(row.get("回答版", "")).strip(),
        "kind": kind,
        "kind_label": "前年度と同じ" if kind == S.KIND_SAME else ("変更" if kind == S.KIND_CHANGE else kind),
        "institution": inst,
        "other_institution": other_name,
        "exam_type": exam_type,
        "extras": extras,
        "other_date": row.get("その他健診予定日", ""),
        "dependent": {"requested": dependent_requested, "relationship": relationship, "name": dependent_name},
        "remarks": row.get("備考", ""),
        "source": row.get("回答元", ""),
        "name": row.get("氏名", ""),
        "email": row.get("社用メール", ""),
    }
    return latest, issues


def validate_responses(target_index: dict[str, dict], response_rows: list[dict],
                       catalog: OptionCatalog, fiscal_year: int) -> tuple[list[ResponseView], list[Issue]]:
    """対象者と回答を突き合わせ、社員ごとの ResponseView と、シート全体の指摘を返す。"""
    workbook_issues: list[Issue] = []
    by_emp: dict[str, list[tuple[dict, dict, list[Issue]]]] = defaultdict(list)
    receipts: dict[str, list[str]] = defaultdict(list)

    for row in response_rows:
        emp = str(row.get("社員番号", "")).strip()
        year = str(row.get("年度", "")).strip()
        latest, issues = parse_response(row, catalog)
        if year != str(fiscal_year):
            issues.append(Issue("error", "year_mismatch", f"回答の年度が違います: 「{year}」（対象: {fiscal_year}）"))
        if _int_or_none(latest["version"]) is None:
            issues.append(Issue("error", "bad_version", f"回答版が数値ではありません: 「{latest['version'] or '空'}」"))
        if latest["receipt"]:
            receipts[latest["receipt"]].append(emp)
        if not emp:
            workbook_issues.append(Issue("error", "response_blank_id",
                                         f"回答シートに社員番号が空の行があります（受付番号 {latest['receipt'] or '無し'}）"))
            continue
        by_emp[emp].append((row, latest, issues))

    dup_receipts = {r for r, emps in receipts.items() if len(emps) > 1}

    views: list[ResponseView] = []
    for emp, target in target_index.items():
        views.append(_build_view(emp, target, by_emp.get(emp, []), catalog, dup_receipts))
    for emp, items in by_emp.items():
        if emp in target_index:
            continue
        view = _build_view(emp, None, items, catalog, dup_receipts)
        view.issues.insert(0, Issue("error", "not_a_target", f"対象者シートに無い社員番号 {emp} からの回答です"))
        view.bucket = "error"
        views.append(view)
    return views, workbook_issues


def _build_view(emp: str, target: dict | None, items: list, catalog: OptionCatalog,
                dup_receipts: set[str]) -> ResponseView:
    issues: list[Issue] = []
    name = (target or {}).get("氏名", "") if target else (items[0][1]["name"] if items else "")
    view = ResponseView(employee_id=emp, name=name,
                        target=target_payload(target, catalog) if target else {})

    # 回答版の重複・受付番号の重複
    seen_versions: dict[str, int] = defaultdict(int)
    for _, latest, _ in items:
        seen_versions[latest["version"]] += 1
    for version, n in seen_versions.items():
        if n > 1:
            issues.append(Issue("error", "duplicate_version", f"回答版 {version or '空'} の回答が {n} 行あります"))
    for _, latest, _ in items:
        if latest["receipt"] in dup_receipts:
            issues.append(Issue("error", "duplicate_receipt", f"受付番号 {latest['receipt']} が他の行と重複しています"))
            break

    # 最新回答＝回答版の最大
    latest = None
    if items:
        ordered = sorted(items, key=lambda it: (_int_or_none(it[1]["version"]) or -1, it[1]["answered_at"]))
        _, latest, row_issues = ordered[-1]
        issues.extend(row_issues)
        # 旧版の error も残す（重複や年度違いは一覧で見えないと困る）
        for _, _, old_issues in ordered[:-1]:
            issues.extend(i for i in old_issues if i.level == "error" and i.code in ("year_mismatch", "bad_version"))
    view.latest = latest
    view.history_count = len(items)

    if target:
        status = str(target.get("申込状態", "")).strip() or S.STATUS_UNSENT
        enrollment = str(target.get("在籍区分", "")).strip()
        if enrollment and enrollment != "0":
            issues.append(Issue("warning", "not_enrolled",
                                f"在籍区分が「{ENROLLMENT_LABELS.get(enrollment, enrollment)}」です"))
        if status == S.STATUS_ANSWERED and latest is None:
            issues.append(Issue("error", "status_answered_without_response", "対象者は「回答済」なのに回答シートに回答がありません"))
        if latest is not None and status == S.STATUS_UNSENT:
            issues.append(Issue("error", "status_unsent_with_response", "対象者は「未送信」なのに回答があります"))
        if latest is not None:
            t_receipt = str(target.get("受付番号", "")).strip()
            t_version = str(target.get("回答版", "")).strip()
            if t_receipt != latest["receipt"]:
                issues.append(Issue("error", "receipt_mismatch",
                                    f"対象者の受付番号「{t_receipt or '空'}」と最新回答「{latest['receipt'] or '空'}」が一致しません"))
            elif t_version != latest["version"]:
                issues.append(Issue("error", "version_mismatch",
                                    f"対象者の回答版「{t_version or '空'}」と最新回答「{latest['version'] or '空'}」が一致しません"))
            if latest["kind"] == S.KIND_SAME and (target.get("前年度情報元", "") or S.SOURCE_NONE) == S.SOURCE_NONE:
                issues.append(Issue("error", "same_without_previous", "「前年度と同じ」なのに対象者に前年度情報がありません"))
        view.bucket = bucket_of(target, latest, issues)
    view.issues = issues
    return view


def bucket_of(target: dict, latest: dict | None, issues: list[Issue]) -> str:
    if any(i.level == "error" for i in issues):
        return "error"
    status = str(target.get("申込状態", "")).strip() or S.STATUS_UNSENT
    if status == S.STATUS_INVALID:
        return "invalid"
    if status == S.STATUS_REANSWER:
        return "reanswer_pending"
    if latest is not None:
        return "answered"
    if str(target.get("初回アクセス日時", "")).strip():
        return "accessed_only"
    if str(target.get("送信日時", "")).strip() or status == S.STATUS_SENT:
        return "sent_not_accessed"
    return "unsent"


def summarize(views: list[ResponseView], target_count: int | None = None) -> dict[str, int]:
    counts = {b: 0 for b in BUCKETS}
    for v in views:
        counts[v.bucket] = counts.get(v.bucket, 0) + 1
    counts["targets"] = target_count if target_count is not None else sum(1 for v in views if v.target)
    return counts


def sort_for_display(views: list[ResponseView]) -> list[ResponseView]:
    order = {b: i for i, b in enumerate(("error", "reanswer_pending", "answered", "accessed_only",
                                          "sent_not_accessed", "unsent", "invalid"))}
    return sorted(views, key=lambda v: (order.get(v.bucket, 99), v.employee_id))


def build_report(target_rows: list[dict], response_rows: list[dict],
                 catalog: OptionCatalog, fiscal_year: int) -> dict:
    """ルートがそのまま jsonify できる形にまとめる。"""
    index, target_issues = index_targets(target_rows, fiscal_year)
    views, response_issues = validate_responses(index, response_rows, catalog, fiscal_year)
    ordered = sort_for_display(views)
    return {
        "counts": summarize(views, target_count=len(index)),
        "rows": [v.as_dict() for v in ordered],
        "workbook_issues": [i.as_dict() for i in (target_issues + response_issues)],
    }
