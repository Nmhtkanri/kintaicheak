# -*- coding: utf-8 -*-
"""jinjer添付ジョブ本体（デタッチ子プロセスで動く。2026-08-28 谷津さん決定）。

「今すぐ」「今夜」ともハブが launcher.py --daicho-attach（exe）／daicho_attach_run.py
（python）を DETACHED_PROCESS で起動する＝ハブの黒い窓を閉じても添付は走り続ける。
1期≈165枚×実測約60秒/枚≈3時間かかるため、リクエストスレッドやdaemonスレッドでは持たない。

- 進捗: NAS の 添付進捗.json へ tmp→os.replace のアトミック置換（標報投入と同じ型。
  ハブ画面・別PC・翌朝のハブ再起動後からも読める）
- 開始待機（--start-at HH:MM）: 子プロセス内で待つ（schtasks は使わない）
- NASロック: 取得は子プロセス側（「今夜」の予約が日中からロックを握らないため）。
  jinjer のレート制限はテナント単位なので、標報投入ロックが生きている間は待つ
- キャンセル: 添付キャンセル.flag を見て、予約中は即・実行中は次の1件から止まる
- スリープ抑止: SetThreadExecutionState（実行中はPCを眠らせない。画面は消えてよい）
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
import time
from pathlib import Path

from . import config

# 自ロックの stale 判定。実行3時間＋余裕（標報投入は2時間だが、こちらは長丁場）
LOCK_MAX_AGE_HOURS = 5
# 標報投入ロックの stale 判定（shaho_writer.acquire_lock の max_age_hours と同じ2時間）
SHAHO_LOCK_MAX_AGE_HOURS = 2

_ES_CONTINUOUS = 0x80000000
_ES_SYSTEM_REQUIRED = 0x00000001


class HakenAttachError(RuntimeError):
    pass


def _now() -> dt.datetime:
    return dt.datetime.now()


def _current_user() -> str:
    import getpass

    return os.environ.get("USERNAME") or os.environ.get("USER") or getpass.getuser()


def _keep_awake(on: bool) -> None:
    """実行中はシステムスリープを抑止する（ディスプレイは消えてよい）。失敗しても続行。"""
    try:
        import ctypes

        flags = _ES_CONTINUOUS | (_ES_SYSTEM_REQUIRED if on else 0)
        ctypes.windll.kernel32.SetThreadExecutionState(flags)
    except Exception:  # noqa: BLE001
        pass


# ---------------------------------------------------------------------------
# 進捗ファイル
# ---------------------------------------------------------------------------

def write_progress(payload: dict) -> None:
    """tmpに書いて os.replace（アトミック置換。NASの書き込み遅延・読みかけ対策）。"""
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    tmp = config.ATTACH_PROGRESS_JSON.with_name(config.ATTACH_PROGRESS_JSON.name + ".tmp")
    payload = {**payload, "updated_at": _now().isoformat(timespec="seconds")}
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, config.ATTACH_PROGRESS_JSON)


def read_progress() -> dict:
    """進捗を読む。無ければ {'state': 'none'}。壊れて読めない瞬間は 'unknown'。"""
    path = config.ATTACH_PROGRESS_JSON
    if not path.exists():
        return {"state": "none"}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {"state": "unknown"}


def request_cancel() -> None:
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    config.ATTACH_CANCEL_FLAG.write_text(
        f"{_now():%Y-%m-%d %H:%M:%S} {_current_user()}", encoding="utf-8")


def _cancel_requested() -> bool:
    return config.ATTACH_CANCEL_FLAG.exists()


def _clear_cancel() -> None:
    try:
        config.ATTACH_CANCEL_FLAG.unlink()
    except OSError:
        pass


# ---------------------------------------------------------------------------
# ロック（自ロック＝Config.HAKEN_ATTACH_LOCK_FILE、標報投入ロックも見る）
# ---------------------------------------------------------------------------

def _lock_paths() -> tuple[Path, Path]:
    from config import Config

    return Path(Config.HAKEN_ATTACH_LOCK_FILE), Path(Config.SHAHO_IMPORT_LOCK_FILE)


def lock_info(path: Path, max_age_hours: float) -> dict | None:
    """有効な（staleでない）ロックの中身。無ければ None。読めない壊れロックは stale 扱い。"""
    if not path.exists():
        return None
    try:
        held = json.loads(path.read_text(encoding="utf-8"))
        started = dt.datetime.fromisoformat(held.get("started_at"))
    except (OSError, ValueError, TypeError):
        return None
    if (_now() - started).total_seconds() >= max_age_hours * 3600:
        return None
    return held


def acquire_lock(count: int) -> Path:
    """自ロックを取る。別PCの有効ロックがあれば HakenAttachError（フェイルクローズ）。"""
    own, _shaho = _lock_paths()
    held = lock_info(own, LOCK_MAX_AGE_HOURS)
    if held:
        raise HakenAttachError(
            f"いま別の添付ジョブが動いています（{held.get('user', '不明')} が "
            f"{held.get('started_at', '?')} に開始・{held.get('count', '?')}件）。"
            "終わってから実行してください")
    own.parent.mkdir(parents=True, exist_ok=True)
    own.write_text(json.dumps({"user": _current_user(), "started_at": _now().isoformat(),
                               "count": count, "pid": os.getpid()}, ensure_ascii=False),
                   encoding="utf-8")
    return own


def release_lock() -> None:
    own, _shaho = _lock_paths()
    try:
        own.unlink()
    except OSError:
        pass


def own_lock_info() -> dict | None:
    """自ロックが有効ならその中身（ハブ画面用）。"""
    own, _shaho = _lock_paths()
    return lock_info(own, LOCK_MAX_AGE_HOURS)


def shaho_lock_message() -> str | None:
    """標報投入が実行中ならその説明文、いなければ None。"""
    _own, shaho = _lock_paths()
    held = lock_info(shaho, SHAHO_LOCK_MAX_AGE_HOURS)
    if held is None:
        return None
    return (f"標報投入が実行中です（{held.get('user', '不明')} が "
            f"{held.get('started_at', '?')} に開始）。429がテナント単位のため終了を待ちます")


# ---------------------------------------------------------------------------
# プレビュー（dry-run。書き込みなし）と verify
# ---------------------------------------------------------------------------

def preview(employees: list[str] | None = None, limit: int | None = None) -> dict:
    """scan_pdfs＋jinjer側レコードで「何件添付することになるか」を数える（GETのみ）。"""
    from collections import defaultdict

    from . import jinjer_attach as ja

    jobs, skipped = ja.scan_pdfs(employees=set(employees) if employees else None)
    if limit:
        jobs = jobs[:limit]
    emp_ids = sorted({emp for emp, *_ in jobs})
    plan: list[dict] = []
    already = 0
    if jobs:
        client = ja._client(25.0)
        records = ja.fetch_records(client, emp_ids)
        groups: dict[tuple[str, str], list] = defaultdict(list)
        for emp, folder, pdf, date in jobs:
            groups[(emp, ja._norm_date(date))].append((emp, folder, pdf, date))
        for (emp, date_n), job_list in groups.items():
            rows = [r for r in records.get(emp, []) if r["date"] == date_n]
            attached_n = sum(1 for r in rows if r["attached"])
            has_empty = any(not r["attached"] for r in rows)
            already += min(attached_n, len(job_list))
            for emp_, folder, pdf, date in job_list[min(attached_n, len(job_list)):]:
                plan.append({"emp": emp_, "folder": folder, "pdf": pdf.name, "date": date,
                             "action": "既存レコードに添付" if has_empty else "レコード作成→添付"})
    return {"total": len(jobs), "people": len(emp_ids), "already": already,
            "to_attach": len(plan), "plan": plan,
            "skipped": list(skipped),
            # 実測 約60秒/枚（POST+PATCHの2書き込み。2026-08-28 のログ812行より）
            "eta_minutes": len(plan)}


def verify_data(employees: list[str] | None = None) -> dict:
    """PDFフォルダと jinjer 側レコードを突き合わせ、未反映一覧を返す（GETのみ）。"""
    from collections import defaultdict

    from . import jinjer_attach as ja

    jobs, _skipped = ja.scan_pdfs(employees=set(employees) if employees else None)
    emp_ids = sorted({emp for emp, *_ in jobs})
    ok = 0
    missing: list[dict] = []
    if jobs:
        client = ja._client(25.0)
        records = ja.fetch_records(client, emp_ids)
        groups: dict[tuple[str, str], list] = defaultdict(list)
        for emp, folder, pdf, date in jobs:
            groups[(emp, ja._norm_date(date))].append((emp, pdf, date))
        for (emp, date_n), job_list in groups.items():
            attached_n = sum(1 for r in records.get(emp, [])
                             if r["date"] == date_n and r["attached"])
            ok += min(attached_n, len(job_list))
            for emp_, pdf, date in job_list[min(attached_n, len(job_list)):]:
                missing.append({"emp": emp_, "pdf": pdf.name, "date": date})
    return {"total": len(jobs), "ok": ok, "missing": missing}


# ---------------------------------------------------------------------------
# ジョブ本体（デタッチ子プロセス／フォールバック時はスレッドからも呼ばれる）
# ---------------------------------------------------------------------------

def _parse_start_at(start_at: str | None) -> dt.datetime | None:
    """'22:00' → 今日のその時刻（すでに過ぎていれば None＝即時開始）。"""
    if not start_at:
        return None
    try:
        h, m = (int(x) for x in start_at.strip().split(":", 1))
        target = _now().replace(hour=h, minute=m, second=0, microsecond=0)
    except (ValueError, AttributeError):
        raise HakenAttachError(f"開始時刻の形式が不正です: {start_at!r}（例: 22:00）")
    return target if target > _now() else None


def run_attach_job(execute: bool = True, employees: list[str] | None = None,
                   limit: int | None = None, interval: float = 25.0,
                   start_at: str | None = None) -> int:
    """添付ジョブ1本ぶん。戻り値は終了コード（0=正常/キャンセル、1=エラー）。"""
    from . import jinjer_attach as ja

    base = {"user": _current_user(), "pid": os.getpid(),
            "employees": employees or [], "limit": limit,
            "started_at": _now().isoformat(timespec="seconds"),
            "start_at": start_at or ""}
    _keep_awake(True)
    locked = False
    try:
        # --- 開始待機（「今夜回す」。60秒ごとにキャンセルを見る）---
        target = _parse_start_at(start_at)
        if target is not None:
            write_progress({**base, "state": "scheduled",
                            "message": f"{target:%H:%M} に開始予定（PCの電源は入れたままに）"})
            while _now() < target:
                if _cancel_requested():
                    _clear_cancel()
                    write_progress({**base, "state": "cancelled",
                                    "finished_at": _now().isoformat(timespec="seconds"),
                                    "message": "開始前に予約を取り消しました"})
                    return 0
                time.sleep(min(60, max((target - _now()).total_seconds(), 1)))

        # --- 標報投入が動いていれば待つ（テナント単位429の相互防衛・15分間隔）---
        while True:
            msg = shaho_lock_message()
            if msg is None:
                break
            write_progress({**base, "state": "scheduled", "message": msg})
            for _ in range(15):
                if _cancel_requested():
                    _clear_cancel()
                    write_progress({**base, "state": "cancelled",
                                    "finished_at": _now().isoformat(timespec="seconds"),
                                    "message": "待機中に取り消しました"})
                    return 0
                time.sleep(60)

        # --- 自ロック ---
        acquire_lock(count=limit or 0)
        locked = True

        # --- 実行 ---
        progress = {"done": 0, "skip": 0, "total": 0, "current": ""}

        def _on_progress(done: int, skip: int, total: int, current: str) -> None:
            progress.update(done=done, skip=skip, total=total, current=current)
            write_progress({**base, "state": "running", **progress})

        write_progress({**base, "state": "running", **progress,
                        "message": "対象PDFと jinjer 側レコードを確認しています"})
        done, skip, total = ja.run(
            employees=employees, dry_run=not execute, write_interval=interval,
            limit=limit, on_progress=_on_progress, should_stop=_cancel_requested)

        cancelled = _cancel_requested()
        _clear_cancel()
        state = "cancelled" if cancelled else "done"
        write_progress({**base, "state": state, "done": done, "skip": skip, "total": total,
                        "finished_at": _now().isoformat(timespec="seconds"),
                        "message": (f"添付 {done} / 添付済みスキップ {skip} / 対象 {total}"
                                    + ("（キャンセルで途中終了。再実行で続きから）" if cancelled else
                                       "。反映は非同期（最大35分）なので後で verify を"))})
        return 0
    except Exception as e:  # noqa: BLE001 — 子プロセスの最後の砦。状態を残して終わる
        write_progress({**base, "state": "error",
                        "finished_at": _now().isoformat(timespec="seconds"),
                        "message": f"{type(e).__name__}: {e}"})
        return 1
    finally:
        if locked:
            release_lock()
        _keep_awake(False)


def main_cli(argv: list[str] | None = None) -> int:
    """デタッチ子プロセスの入口（launcher.py --daicho-attach / daicho_attach_run.py）。"""
    import argparse

    from dotenv import load_dotenv

    load_dotenv(Path(__file__).resolve().parents[2] / ".env")
    args = [a for a in list(argv or sys.argv)[1:] if a != "--daicho-attach"]
    ap = argparse.ArgumentParser(prog="daicho-attach", description="派遣台帳PDFのjinjer添付ジョブ")
    ap.add_argument("--execute", action="store_true", help="実書き込み（無指定はdry-run）")
    ap.add_argument("--employee", action="append", help="社員番号で絞る（複数可）")
    ap.add_argument("--limit", type=int, help="先頭N件だけ処理")
    ap.add_argument("--interval", type=float, default=25.0, help="書き込み間隔秒")
    ap.add_argument("--start-at", help="この時刻まで待って開始（例: 22:00）")
    ns = ap.parse_args(args)
    return run_attach_job(execute=ns.execute, employees=ns.employee, limit=ns.limit,
                          interval=ns.interval, start_at=ns.start_at)
