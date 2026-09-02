# -*- coding: utf-8 -*-
"""健康診断申込: 監査ログ行の組み立て（Hub が書く分だけ）。"""

from __future__ import annotations

from services.health_apply import schema as S


def _row(now_iso: str, event: str, user: str, fiscal_year: int, employee_id: str, detail: str) -> list[str]:
    values = {
        "日時": now_iso, "イベント": event, "実行元": S.ACTOR_HUB, "実行者": user,
        "年度": str(fiscal_year), "社員番号": employee_id, "詳細": detail,
    }
    return [values[name] for name in S.AUDIT_HEADERS]


def target_register_rows(user: str, fiscal_year: int, added: list[dict], now_iso: str, fingerprint: str) -> list[list[str]]:
    """REGISTER_BATCH 1行 ＋ 追加した社員ごとに REGISTER 1行。

    added は PlanRow.as_dict() の形（employee_id / previous{source, institution{code}, exam_type{code}, extras[]}）。
    """
    rows = [_row(now_iso, "REGISTER_BATCH", user, fiscal_year, "",
                 f"{len(added)}名を対象者へ追記 fingerprint={fingerprint[:12]}")]
    for item in added:
        prev = item.get("previous", {}) or {}
        detail = (f"前年度情報元={prev.get('source', '')} "
                  f"機関={(prev.get('institution') or {}).get('code', '')} "
                  f"種別={(prev.get('exam_type') or {}).get('code', '')} "
                  f"追加検査={';'.join((e or {}).get('code', '') for e in (prev.get('extras') or []))}")
        rows.append(_row(now_iso, "REGISTER", user, fiscal_year, str(item.get("employee_id", "")), detail.strip()))
    return rows
