# -*- coding: utf-8 -*-
"""健康診断申込: 対象者プレビューの一時保持（管理者PCのローカル・JSON・2時間で消える）。

社員番号・氏名・メールを含むので共有フォルダには置かない（Config.HEALTH_APPLY_SESSION_DIR は
%LOCALAPPDATA% の絶対パス。launcher.py が作業フォルダを NAS へ chdir しても相対パスを使わない）。
健診HPMの _health_session_*（pickle・health_ 接頭辞・8時間）とは別物として持つ。
"""

from __future__ import annotations

import datetime as _dt
import json
import os
import re
import time
import uuid

SESSION_PREFIX = "happly_"
SESSION_ID_RE = re.compile(r"^happly_[0-9a-f]{32}$")


def new_session_id() -> str:
    return SESSION_PREFIX + uuid.uuid4().hex


def session_path(directory: str, session_id: str) -> str | None:
    """ID を検証してからパスにする（../ などを弾く）。"""
    if not SESSION_ID_RE.match(str(session_id or "")):
        return None
    return os.path.join(directory, f"{session_id}.json")


def save_preview(directory: str, payload: dict) -> str:
    os.makedirs(directory, exist_ok=True)
    session_id = new_session_id()
    data = dict(payload)
    data["session_id"] = session_id
    data["saved_at"] = _dt.datetime.now().isoformat(timespec="seconds")
    with open(session_path(directory, session_id), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    return session_id


def load_preview(directory: str, session_id: str, ttl_hours: float) -> dict | None:
    """期限内なら中身を返す。期限切れは消して None。不正IDや無いものも None。"""
    path = session_path(directory, session_id)
    if path is None or not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return None
    saved_at = str(data.get("saved_at", ""))
    try:
        age = (_dt.datetime.now() - _dt.datetime.fromisoformat(saved_at)).total_seconds()
    except ValueError:
        age = float("inf")
    if age > float(ttl_hours) * 3600:
        drop_preview(directory, session_id)
        return None
    return data


def drop_preview(directory: str, session_id: str) -> None:
    path = session_path(directory, session_id)
    if path is None:
        return
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


def sweep_previews(directory: str, ttl_hours: float) -> int:
    """期限切れのプレビューを消す（happly_ で始まる JSON だけ触る）。"""
    try:
        names = os.listdir(directory)
    except OSError:
        return 0
    limit = float(ttl_hours) * 3600
    now = time.time()
    removed = 0
    for name in names:
        if not (name.startswith(SESSION_PREFIX) and name.endswith(".json")):
            continue
        path = os.path.join(directory, name)
        try:
            if now - os.path.getmtime(path) > limit:
                os.remove(path)
                removed += 1
        except OSError:
            continue
    return removed
