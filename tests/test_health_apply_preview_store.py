# -*- coding: utf-8 -*-
"""健康診断申込: 対象者プレビューの一時保持（ローカル・JSON・期限切れで消える）。"""

import datetime as dt
import json
import os
import time

from services.health_apply import preview_store as P


def test_save_and_load_roundtrip(tmp_path):
    d = tmp_path / "sessions"
    sid = P.save_preview(str(d), {"fiscal_year": 2027, "rows": [{"employee_id": "2099001", "name": "試験 太郎"}]})
    assert P.SESSION_ID_RE.match(sid)
    assert (d / f"{sid}.json").exists()
    data = P.load_preview(str(d), sid, 2)
    assert data["fiscal_year"] == 2027 and data["rows"][0]["name"] == "試験 太郎"
    assert data["session_id"] == sid and data["saved_at"]
    raw = (d / f"{sid}.json").read_text(encoding="utf-8")
    assert "試験 太郎" in raw   # ensure_ascii=False


def test_expired_preview_is_removed_on_load(tmp_path):
    d = tmp_path / "sessions"
    sid = P.save_preview(str(d), {"x": 1})
    path = d / f"{sid}.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    data["saved_at"] = (dt.datetime.now() - dt.timedelta(hours=3)).isoformat(timespec="seconds")
    path.write_text(json.dumps(data), encoding="utf-8")
    assert P.load_preview(str(d), sid, 2) is None
    assert not path.exists()


def test_bad_ids_and_missing_files_return_none(tmp_path):
    d = tmp_path / "sessions"
    d.mkdir()
    assert P.load_preview(str(d), "../etc/passwd", 2) is None
    assert P.load_preview(str(d), "health_" + "a" * 32, 2) is None
    assert P.load_preview(str(d), "happly_" + "a" * 32, 2) is None
    assert P.session_path(str(d), "happly_zz") is None
    (d / ("happly_" + "b" * 32 + ".json")).write_text("{not json", encoding="utf-8")
    assert P.load_preview(str(d), "happly_" + "b" * 32, 2) is None


def test_drop_and_sweep_only_touch_own_files(tmp_path):
    d = tmp_path / "sessions"
    sid = P.save_preview(str(d), {"x": 1})
    old = P.save_preview(str(d), {"x": 2})
    other = d / ("health_" + "c" * 32 + ".pkl")
    other.write_bytes(b"x")
    stale = time.time() - 3 * 3600
    os.utime(d / f"{old}.json", (stale, stale))
    os.utime(other, (stale, stale))
    assert P.sweep_previews(str(d), 2) == 1
    assert not (d / f"{old}.json").exists()
    assert other.exists() and (d / f"{sid}.json").exists()
    P.drop_preview(str(d), sid)
    assert not (d / f"{sid}.json").exists()
    P.drop_preview(str(d), sid)              # 2回目も落ちない
    P.drop_preview(str(d), "bad id")
    assert P.sweep_previews(str(tmp_path / "nope"), 2) == 0
