# -*- coding: utf-8 -*-
"""健康診断申込: 対象者登録の差分計画（純粋関数）。

貼り付けられた社員番号 → jinjer の情報 → 既存「対象者」シートとの突合 → add / unchanged / conflict /
blocked に分ける。conflict か blocked が1件でもあれば登録できない（全体停止）。
同じ貼り付けの再実行は全件 unchanged になるので冪等。

確認語は "REGISTER {年度} {追加件数}"。プレビューと commit の間にシートが変わっていないかは
plan_fingerprint（追加行の内容の sha256）で見張る。
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from dataclasses import dataclass, field

from services.health_apply import schema as S
from services.health_apply.jinjer_source import EmployeeProfile, PreviousRaw
from services.health_apply.options import KIND_EXAM_TYPE, KIND_EXTRA, KIND_INSTITUTION, OptionCatalog
from services.health_apply.responses import ENROLLMENT_LABELS, Issue, split_codes

EMPLOYEE_ID_RE = re.compile(r"^20\d{5}$")   # 自社社員は 20 始まり7桁
EMAIL_DOMAIN = "@nmht.co.jp"               # 社用メールのドメイン。緩めるならここだけ

ACTION_ADD = "add"
ACTION_UNCHANGED = "unchanged"
ACTION_CONFLICT = "conflict"
ACTION_BLOCKED = "blocked"
ACTION_LABELS = {ACTION_ADD: "追加", ACTION_UNCHANGED: "変更なし", ACTION_CONFLICT: "競合", ACTION_BLOCKED: "登録不可"}

# 既存行との比較に使う列（これが同値なら「変更なし」）
KEY_COLUMNS = ("氏名", "社用メール", "前年度情報元", "前年度健診機関コード", "前年度健診種別コード", "前年度追加検査")


@dataclass
class PreviousExam:
    source: str = S.SOURCE_NONE
    year: str = ""
    institution_code: str = ""
    institution_name: str = ""
    institution_raw: str = ""
    exam_type_code: str = ""
    exam_type_name: str = ""
    extra_codes: tuple[str, ...] = ()
    extra_names: tuple[str, ...] = ()
    content_labels: tuple[str, ...] = ()
    exam_date: str = ""
    notes: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "source": self.source, "year": self.year,
            "institution": {"code": self.institution_code, "name": self.institution_name, "raw": self.institution_raw},
            "exam_type": {"code": self.exam_type_code, "name": self.exam_type_name},
            "extras": [{"code": c, "name": n} for c, n in zip(self.extra_codes, self.extra_names)],
            "content_labels": list(self.content_labels),
            "exam_date": self.exam_date,
            "notes": list(self.notes),
        }


@dataclass
class Candidate:
    employee_id: str
    name: str = ""
    email: str = ""
    enrollment_id: str = ""
    enrollment_name: str = ""
    retirement_date: str = ""
    previous: PreviousExam = field(default_factory=PreviousExam)
    issues: list[Issue] = field(default_factory=list)

    @property
    def blocked(self) -> bool:
        return any(i.level == "error" for i in self.issues)

    def as_dict(self) -> dict:
        return {
            "employee_id": self.employee_id, "name": self.name, "email": self.email,
            "enrollment": self.enrollment_id,
            "enrollment_label": self.enrollment_name or ENROLLMENT_LABELS.get(self.enrollment_id, self.enrollment_id),
            "retirement_date": self.retirement_date,
            "previous": self.previous.as_dict(),
            "issues": [i.as_dict() for i in self.issues],
        }


@dataclass
class PlanRow:
    candidate: Candidate
    action: str
    reasons: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        d = self.candidate.as_dict()
        d["action"] = self.action
        d["action_label"] = ACTION_LABELS.get(self.action, self.action)
        d["reasons"] = list(self.reasons)
        return d


@dataclass
class TargetPlan:
    rows: list[PlanRow]
    fiscal_year: int
    input_issues: list[Issue] = field(default_factory=list)

    def by_action(self, action: str) -> list[PlanRow]:
        return [r for r in self.rows if r.action == action]

    def counts(self) -> dict[str, int]:
        return {
            "input": len(self.rows),
            "add": len(self.by_action(ACTION_ADD)),
            "unchanged": len(self.by_action(ACTION_UNCHANGED)),
            "conflict": len(self.by_action(ACTION_CONFLICT)),
            "blocked": len(self.by_action(ACTION_BLOCKED)),
            "warnings": sum(1 for r in self.rows if any(i.level == "warning" for i in r.candidate.issues)),
            "input_errors": sum(1 for i in self.input_issues if i.level == "error"),
        }

    def can_commit(self) -> bool:
        c = self.counts()
        return c["conflict"] == 0 and c["blocked"] == 0 and c["add"] > 0

    def as_dict(self) -> dict:
        return {
            "fiscal_year": self.fiscal_year,
            "counts": self.counts(),
            "can_commit": self.can_commit(),
            "confirm_phrase": confirm_phrase(self.fiscal_year, self.counts()["add"]),
            "input_issues": [i.as_dict() for i in self.input_issues],
            "rows": [r.as_dict() for r in self.rows],
        }


# ----------------------------------------------------------------------
# 入力
# ----------------------------------------------------------------------

def parse_employee_ids(text: str | None) -> tuple[list[str], list[Issue]]:
    """貼り付けテキスト → 社員番号のリスト（順序を保ち、重複は1件に）。形式外は Issue。"""
    s = unicodedata.normalize("NFKC", str(text or ""))
    ids: list[str] = []
    issues: list[Issue] = []
    seen: set[str] = set()
    for line_no, line in enumerate(s.splitlines(), start=1):
        for token in re.split(r"[,\s;、；\t]+", line):
            token = token.strip()
            if not token:
                continue
            if not EMPLOYEE_ID_RE.match(token):
                issues.append(Issue("error", "bad_employee_id",
                                    f"{line_no}行目 「{token}」: 20始まり7桁の社員番号ではありません"))
                continue
            if token in seen:
                issues.append(Issue("warning", "duplicate_employee_id",
                                    f"{line_no}行目 「{token}」: 重複しているので1件にまとめました"))
                continue
            seen.add(token)
            ids.append(token)
    return ids, issues


# ----------------------------------------------------------------------
# 前年度情報をコードへ寄せる
# ----------------------------------------------------------------------

def resolve_previous(raw: PreviousRaw | None, catalog: OptionCatalog) -> PreviousExam:
    if raw is None or raw.source == S.SOURCE_NONE:
        return PreviousExam(source=S.SOURCE_NONE, notes=list(raw.notes) if raw else [])
    notes = list(raw.notes)
    inst = catalog.resolve_name(KIND_INSTITUTION, raw.institution_text) if raw.institution_text else None
    types = []
    extras = []
    unresolved = []
    for label in raw.content_labels:
        t = catalog.resolve_name(KIND_EXAM_TYPE, label)
        if t is not None:
            types.append(t)
            continue
        e = catalog.resolve_name(KIND_EXTRA, label)
        if e is not None:
            extras.append(e)
            continue
        unresolved.append(label)

    ok = True
    if inst is None:
        ok = False
        notes.append(f"前年度の健診機関「{raw.institution_text or '空'}」を選択肢に寄せられません（選択肢シートの別名に足すと解決します）")
    if not types:
        ok = False
        notes.append("前年度の健診内容から健診種別を決められません" + (f"（未対応: {'、'.join(unresolved)}）" if unresolved else ""))
    elif len(types) > 1:
        ok = False
        notes.append("前年度の健診内容に健診種別が複数あります: " + "、".join(t.name for t in types))
    elif unresolved:
        notes.append("選択肢に寄せられなかった健診内容: " + "、".join(unresolved))

    if not ok:
        return PreviousExam(source=S.SOURCE_NONE, year=raw.year, institution_raw=raw.institution_text,
                            content_labels=tuple(raw.content_labels), exam_date=raw.exam_date, notes=notes)
    seen = set()
    extra_codes, extra_names = [], []
    for e in extras:
        if e.code not in seen:
            seen.add(e.code)
            extra_codes.append(e.code)
            extra_names.append(e.name)
    return PreviousExam(
        source=raw.source, year=raw.year,
        institution_code=inst.code, institution_name=inst.name, institution_raw=raw.institution_text,
        exam_type_code=types[0].code, exam_type_name=types[0].name,
        extra_codes=tuple(extra_codes), extra_names=tuple(extra_names),
        content_labels=tuple(raw.content_labels), exam_date=raw.exam_date, notes=notes,
    )


# ----------------------------------------------------------------------
# 候補と差分
# ----------------------------------------------------------------------

def build_candidates(employee_ids: list[str], profiles: dict[str, EmployeeProfile],
                     previous_raw: dict[str, PreviousRaw], catalog: OptionCatalog) -> list[Candidate]:
    out: list[Candidate] = []
    for emp in employee_ids:
        profile = profiles.get(emp)
        if profile is None:
            out.append(Candidate(employee_id=emp, issues=[Issue("error", "not_in_jinjer", "jinjer にこの社員番号がありません")]))
            continue
        issues: list[Issue] = []
        email = profile.email.strip()
        if not email:
            issues.append(Issue("error", "email_missing", "jinjer に社用メールがありません"))
        elif not email.lower().endswith(EMAIL_DOMAIN):
            issues.append(Issue("error", "email_domain", f"社用メールが {EMAIL_DOMAIN} ではありません: {email}"))
        if profile.enrollment_id and profile.enrollment_id != "0":
            label = profile.enrollment_name or ENROLLMENT_LABELS.get(profile.enrollment_id, profile.enrollment_id)
            issues.append(Issue("warning", "not_enrolled",
                                f"在籍区分が「{label}」です" + (f"（退職日 {profile.retirement_date}）" if profile.retirement_date else "")))
        prev = resolve_previous(previous_raw.get(emp), catalog)
        for note in prev.notes:
            issues.append(Issue("warning", "previous_note", note))
        if prev.source == S.SOURCE_NONE:
            issues.append(Issue("warning", "no_previous", "前年度情報が無いので、本人は「前年度と同じ」を選べません"))
        out.append(Candidate(employee_id=emp, name=profile.name, email=email,
                             enrollment_id=profile.enrollment_id, enrollment_name=profile.enrollment_name,
                             retirement_date=profile.retirement_date, previous=prev, issues=issues))
    return out


def candidate_key_values(c: Candidate) -> dict[str, str]:
    return {
        "氏名": c.name,
        "社用メール": c.email,
        "前年度情報元": c.previous.source,
        "前年度健診機関コード": c.previous.institution_code,
        "前年度健診種別コード": c.previous.exam_type_code,
        "前年度追加検査": ";".join(c.previous.extra_codes),
    }


def plan_targets(candidates: list[Candidate], existing_targets: list[dict], responses: list[dict],
                 fiscal_year: int, input_issues: list[Issue] | None = None) -> TargetPlan:
    year = str(fiscal_year)
    existing = {}
    for row in existing_targets:
        if str(row.get("年度", "")).strip() == year:
            existing.setdefault(str(row.get("社員番号", "")).strip(), row)
    responded = {str(r.get("社員番号", "")).strip() for r in responses if str(r.get("年度", "")).strip() == year}

    rows: list[PlanRow] = []
    for c in candidates:
        if c.blocked:
            rows.append(PlanRow(c, ACTION_BLOCKED, [i.message for i in c.issues if i.level == "error"]))
            continue
        row = existing.get(c.employee_id)
        if row is None:
            rows.append(PlanRow(c, ACTION_ADD))
            continue
        reasons: list[str] = []
        status = str(row.get("申込状態", "")).strip()
        if status == S.STATUS_ANSWERED or str(row.get("受付番号", "")).strip() or c.employee_id in responded:
            reasons.append("既に回答があります（対象者の内容を変えると回答との整合が崩れます）")
        wanted = candidate_key_values(c)
        for col in KEY_COLUMNS:
            old = str(row.get(col, "")).strip()
            new = wanted[col]
            if col == "前年度追加検査":
                old = ";".join(split_codes(old))
            if old != new:
                reasons.append(f"{col}: シート「{old or '空'}」→ jinjer「{new or '空'}」")
        if reasons:
            rows.append(PlanRow(c, ACTION_CONFLICT, reasons))
            continue
        old_enrollment = str(row.get("在籍区分", "")).strip()
        if old_enrollment and c.enrollment_id and old_enrollment != c.enrollment_id:
            c.issues.append(Issue("warning", "enrollment_changed",
                                  f"在籍区分がシート「{old_enrollment}」→ jinjer「{c.enrollment_id}」に変わっています（登録済みなので触りません）"))
        rows.append(PlanRow(c, ACTION_UNCHANGED, ["登録済み・同じ内容"]))
    return TargetPlan(rows=rows, fiscal_year=fiscal_year, input_issues=list(input_issues or []))


def confirm_phrase(fiscal_year: int, n_add: int) -> str:
    return f"REGISTER {fiscal_year} {n_add}"


def build_target_rows(plan: TargetPlan, user: str, now_iso: str) -> list[list[str]]:
    """追加分を対象者シートの列順（Hub 管轄の先頭14列）で返す。15列目以降は Apps Script が書く。"""
    rows: list[list[str]] = []
    for r in plan.by_action(ACTION_ADD):
        c = r.candidate
        p = c.previous
        values = {
            "年度": str(plan.fiscal_year),
            "社員番号": c.employee_id,
            "氏名": c.name,
            "社用メール": c.email,
            "在籍区分": c.enrollment_id,
            "前年度情報元": p.source,
            "前年度健診機関コード": p.institution_code,
            "前年度健診機関名": p.institution_name,
            "前年度健診種別コード": p.exam_type_code,
            "前年度健診種別名": p.exam_type_name,
            "前年度追加検査": ";".join(p.extra_codes),
            "前年度健診機関(原文)": p.institution_raw,
            "登録日時": now_iso,
            "登録者": user,
        }
        rows.append([values.get(name, "") for name in S.TARGET_HEADERS[:S.TARGET_HUB_COLUMNS]])
    return rows


def plan_fingerprint(plan: TargetPlan) -> str:
    """追加行の中身だけから作る指紋（順序に依らない）。commit 時の再計算と比べる。"""
    items = sorted(
        "\t".join([c.employee_id] + [candidate_key_values(c)[k] for k in KEY_COLUMNS])
        for c in (r.candidate for r in plan.by_action(ACTION_ADD))
    )
    h = hashlib.sha256()
    h.update(str(plan.fiscal_year).encode("utf-8"))
    for item in items:
        h.update(b"\n")
        h.update(item.encode("utf-8"))
    return h.hexdigest()
