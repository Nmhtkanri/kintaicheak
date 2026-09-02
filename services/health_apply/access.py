# -*- coding: utf-8 -*-
"""健康診断申込: 利用許可ユーザー（画面とAPIの両方をこれで止める）。

許可CSV（Config.HEALTH_APPLY_ALLOWED_USERS_CSV。列: ユーザー名, 表示名, 備考）に
Windows ログオンユーザーが載っているときだけ使える。CSV が読めなければ **誰も使えない**
（フェイルクローズ）。SAP台帳の書き込み許可（services.sap_import_ledger）と同じ考え方。
"""

from __future__ import annotations

import unicodedata

from services.sap_import_ledger import current_user, load_writers

LABEL = "健診申込モード"


def _norm(s: str) -> str:
    return unicodedata.normalize("NFKC", str(s or "")).strip().casefold()


def check_access(csv_path: str, user: str | None = None) -> tuple[bool, str]:
    """(利用可否, 画面に出す理由) を返す。"""
    who = user or current_user()
    rows = load_writers(csv_path)
    if not rows:
        return False, (f"{LABEL}の利用許可リストが読めません（{csv_path}）。"
                       f"現在のユーザー: {who}")
    allowed = {_norm(r.get("ユーザー名", "")) for r in rows} - {""}
    if _norm(who) in allowed:
        return True, f"利用可（ユーザー: {who}）"
    names = "・".join(str(r.get("表示名") or r.get("ユーザー名") or "").strip() for r in rows)
    return False, (f"{LABEL}を使えるのは {names} のみです。現在のユーザー: {who}"
                   f"（追加する場合は {csv_path} に1行足してください）")
