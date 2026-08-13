# -*- coding: utf-8 -*-
"""スケジュール開始時刻のピンポイント編集（確認ゲート付きの「書く」機能）。

「社員番号, 日付, 新しい開始時刻」を複数行入れると、その日のスケジュール開始だけを
jinjer に反映する。雛形は使わず、実証済みの時刻直書き（汎用データCSV →
POST /v1/kintai-imports 種別5）で書く。2026-08-13 谷津さん依頼。

階層は「書く」（暴走すると jinjer の従業員スケジュールが書き換わる）。
ガード: dry-run 既定 → 差分プレビュー → fingerprint 一致時のみ実行 → 実行ログと
変更前スナップショットを必ず保存（kinou-tsuika-rule の必須ガード）。

jinjer 側の4つの罠と対処（すべて実測済み。詳細は memory: jinjer_api_cheatsheet）:
  1. スケジュール書込は日単位の丸ごと置き換え → 開始だけ送ると休憩予定が消える。
     必ず work-schedules で現状を読み、開始だけ差し替えて
     出勤予定＋退勤予定＋休憩予定のフルセットを送る（本モジュールの核心）。
  2. 休暇レコードがある日はサイレント無視 → requested-day-offs で事前検知して行を弾く。
  3. 書式（YYYY/M/D・秒なし時刻・cp932）→ schedule_import_runner の実証済み部品を再利用。
  4. executor に勤怠管理者ロールが無いとサイレント破棄 → 既定は未指定（マスタ扱い）。

打刻グループの罠: グループ移動者は旧グループの予定が残骸として残るため、
所属履歴から対象日時点のグループを解決してから work-schedules を絞る
（schedule_import_runner と同じ流儀。石下さん・木村さんの実害の再発防止）。
"""
from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable

from services.jinjer_api_client import (
    JinjerAPIError,
    JinjerClient,
    fetch_attendance_groups_at,
    fetch_employee_id_map,
)
from services.kintai_import_runner import poll_import_status, t2m
from services.schedule_import_runner import (
    GENERIC_IMPORT_HEADER,
    build_import_rows,
    plan_fingerprint,
    rows_to_csv_bytes,
)

# 1回の編集で受け付ける上限。ピンポイント修正用途なので小さく絞る
# （大量に書くならスケジュールアップロード機能を使う）。
MAX_EDITS = 200

_DATE_RE = re.compile(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$")
_TIME_RE = re.compile(r"^(\d{1,2}):(\d{2})(?::\d{2})?$")


def fmt_hm(minutes: int) -> str:
    """分 → 'H:MM'（0埋めなし・秒なし。jinjer インポートが受ける形）。"""
    return f"{minutes // 60}:{minutes % 60:02d}"


def parse_edit_lines(text: str) -> tuple[list[dict], list[str]]:
    """「社員番号, 日付, 新開始時刻」の複数行テキストをパースする（純粋関数）。

    区切りはカンマ・タブ・空白のどれでも可。日付は 2026/8/1・2026-08-01 の両対応。
    時刻は H:MM（秒付きは秒を落とす）。開始時刻は当日の時計として 0:00〜23:59 に限定
    （24時超の開始は前日行の領分なので誤入力として弾く）。
    戻り値: (edits, errors)。edits の要素は {emp, date_iso, new_start, new_min}。
    """
    edits: list[dict] = []
    errors: list[str] = []
    seen: set[tuple[str, str]] = set()
    for ln, raw in enumerate((text or "").splitlines(), start=1):
        line = raw.strip().lstrip("﻿")
        if not line:
            continue
        parts = [p for p in re.split(r"[,\t　 ]+", line) if p]
        if len(parts) != 3:
            errors.append(f"{ln}行目: 「社員番号, 日付, 新開始時刻」の3項目で入力してください: {raw!r}")
            continue
        emp, date_s, time_s = parts
        if not re.fullmatch(r"\d{5,}", emp):
            errors.append(f"{ln}行目: 社員番号が数字ではありません: {emp!r}")
            continue
        md = _DATE_RE.match(date_s)
        if not md:
            errors.append(f"{ln}行目: 日付は 2026/8/1 か 2026-08-01 の形式で入力してください: {date_s!r}")
            continue
        try:
            date_iso = datetime(int(md.group(1)), int(md.group(2)), int(md.group(3))).strftime("%Y-%m-%d")
        except ValueError:
            errors.append(f"{ln}行目: 実在しない日付です: {date_s!r}")
            continue
        mt = _TIME_RE.match(time_s)
        if not mt:
            errors.append(f"{ln}行目: 時刻は H:MM で入力してください: {time_s!r}")
            continue
        new_min = int(mt.group(1)) * 60 + int(mt.group(2))
        if not (0 <= new_min <= 23 * 60 + 59):
            errors.append(f"{ln}行目: 開始時刻は 0:00〜23:59 で入力してください（24時超の開始は前日行の扱い）: {time_s!r}")
            continue
        key = (emp, date_iso)
        if key in seen:
            errors.append(f"{ln}行目: 同じ社員×日付が重複しています: {emp} {date_iso}")
            continue
        seen.add(key)
        edits.append({"emp": emp, "date_iso": date_iso,
                      "new_start": fmt_hm(new_min), "new_min": new_min})
    if len(edits) > MAX_EDITS:
        errors.append(f"一度に編集できるのは {MAX_EDITS} 行までです（{len(edits)} 行入力）。"
                      "大量に書き換える場合はスケジュールアップロード機能を使ってください。")
        edits = []
    return edits, errors


# プレビュー行の「状態」値
ST_CHANGE = "変更"
ST_SAME = "変更なし（スキップ）"


def build_plan(
    edits: list[dict],
    *,
    schedules: dict[tuple[str, str], dict[str, dict]],
    day_offs: dict[tuple[str, str], dict[str, str]],
    names: dict[str, str],
    groups: dict[tuple[str, str], tuple[str, str]],
) -> tuple[list[dict], list[dict], list[str]]:
    """編集指示＋現状 → (投入プラン, プレビュー行, エラー) を作る（純粋関数）。

    schedules: {(emp, month): {date_iso: {start,end,breaks,...}}}（現グループで絞った現状）
    day_offs : {(emp, month): {date_iso: 休暇説明}}
    names    : {emp: 姓}（汎用データの氏名列は姓のみ照合）
    groups   : {(emp, date_iso): (打刻グループID, 名前)}
    """
    plan: list[dict] = []
    preview: list[dict] = []
    errors: list[str] = []
    for e in edits:
        emp, date_iso = e["emp"], e["date_iso"]
        month = date_iso[:7]
        label = f"{emp} {names.get(emp, '')} {date_iso}".strip()

        def prev_row(status: str, cur: dict | None = None) -> dict:
            cur = cur or {}
            return {
                "従業員ID": emp, "氏名": names.get(emp, ""), "日付": date_iso,
                "現在の開始": cur.get("start", ""), "新しい開始": e["new_start"],
                "終了(維持)": cur.get("end", ""),
                "休憩(維持)": " / ".join(f"{s}-{x}" for s, x in cur.get("breaks", [])),
                "状態": status,
            }

        if emp not in names:
            errors.append(f"{label}: jinjer の在籍従業員に見つかりません（退職者・番号誤り）")
            continue
        off = (day_offs.get((emp, month)) or {}).get(date_iso)
        if off:
            # 罠2: 休暇日はエラーにならず黙って無反映になるため、送信前に弾く
            errors.append(f"{label}: 休暇レコードがあるため書き込めません（{off}）。"
                          "jinjer は休暇日のスケジュール書込をサイレントに無視します")
            continue
        cur = (schedules.get((emp, month)) or {}).get(date_iso)
        if not cur or not cur.get("start") or not cur.get("end"):
            errors.append(f"{label}: この日のスケジュールがありません（新規作成は対象外。"
                          "スケジュールアップロード機能を使ってください）")
            continue
        end_min = t2m(cur["end"])
        if end_min is not None and e["new_min"] >= end_min:
            errors.append(f"{label}: 新しい開始 {e['new_start']} が退勤予定 {cur['end']} 以降です")
            continue
        cur_min = t2m(cur["start"])
        if cur_min == e["new_min"]:
            preview.append(prev_row(ST_SAME, cur))
            continue
        gid, _gname = groups.get((emp, date_iso), ("", ""))
        if not gid:
            errors.append(f"{label}: 打刻グループを解決できませんでした（所属履歴が空）")
            continue
        # 罠1対策の核心: 開始だけ差し替え、終了・休憩は現状をそのまま送る
        plan.append({
            "emp": emp, "name": names.get(emp, ""), "date_iso": date_iso,
            "start": e["new_start"], "end": cur["end"],
            "breaks": list(cur.get("breaks") or []),
            "store_id": gid, "holiday": "", "template_id": "",
        })
        preview.append(prev_row(ST_CHANGE, cur))
    return plan, preview, errors


@dataclass
class ScheduleStartEditResult:
    ok: bool = True
    dry_run: bool = True
    preview: list[dict] = field(default_factory=list)
    fingerprint: str = ""
    errors: list[str] = field(default_factory=list)
    log: list[str] = field(default_factory=list)
    change_count: int = 0
    import_status: str = ""       # "1"=成功 "2"=失敗 ""=未実行/タイムアウト
    snapshot_path: str = ""       # 変更前スナップショット＋実行ログCSV
    verify_ng: list[str] = field(default_factory=list)


def _collect_current_state(edits: list[dict], client: JinjerClient, log: Callable[[str], None]):
    """API から現状（氏名・打刻グループ・スケジュール・休暇）を集める。"""
    _n2i, names, _amb = fetch_employee_id_map()
    groups: dict[tuple[str, str], tuple[str, str]] = {}
    for date_iso in sorted({e["date_iso"] for e in edits}):
        emps = sorted({e["emp"] for e in edits if e["date_iso"] == date_iso and e["emp"] in names})
        if emps:
            got = fetch_attendance_groups_at(emps, date_iso)
            for emp in emps:
                groups[(emp, date_iso)] = got.get(emp, ("", ""))
    schedules: dict[tuple[str, str], dict[str, dict]] = {}
    day_offs: dict[tuple[str, str], dict[str, str]] = {}
    for emp, month in sorted({(e["emp"], e["date_iso"][:7]) for e in edits if e["emp"] in names}):
        # 同一 (emp, month) に複数日付があってもグループはほぼ同じ。最初の該当日のグループで絞る
        gid = next((groups.get((emp, e["date_iso"]), ("", ""))[0]
                    for e in edits if e["emp"] == emp and e["date_iso"].startswith(month)), "")
        schedules[(emp, month)] = client.get_work_schedules(emp, month, store_id=gid)
        day_offs[(emp, month)] = client.get_requested_day_offs(emp, month)
    log(f"現状取得: 従業員 {len({e['emp'] for e in edits})} 名 / "
        f"月 {len({e['date_iso'][:7] for e in edits})} 種")
    return names, groups, schedules, day_offs


def _write_snapshot(output_dir: Path, plan: list[dict], preview: list[dict],
                    import_status: str, file_name: str) -> Path:
    """変更前スナップショット＋実行記録を BOM付きUTF-8 CSV で残す（Excelで開ける）。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = output_dir / f"スケジュール開始編集_実行ログ_{ts}.csv"
    cols = ["従業員ID", "氏名", "日付", "現在の開始", "新しい開始", "終了(維持)",
            "休憩(維持)", "状態", "インポート結果", "投入ファイル名"]
    with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in preview:
            w.writerow({**{c: r.get(c, "") for c in cols},
                        "インポート結果": import_status, "投入ファイル名": file_name})
    return path


def run_schedule_start_edit(
    text: str,
    *,
    dry_run: bool = True,
    expected_fingerprint: str = "",
    executor_id: str = "",
    output_dir: Path | str = "outputs",
    client: JinjerClient | None = None,
    log_func: Callable[[str], None] | None = None,
    poll_func: Callable = poll_import_status,
) -> ScheduleStartEditResult:
    """入力テキスト → プレビュー（dry_run=True）／投入（dry_run=False）。

    実行時は expected_fingerprint（プレビューが返した値）と、jinjer から取り直した
    現状で再計算した fingerprint が一致しない限り送信しない（承認後に jinjer 側が
    変わった場合の保険。schedule_import_runner と同じゲート）。
    """
    result = ScheduleStartEditResult(dry_run=dry_run)

    def log(msg: str) -> None:
        result.log.append(msg)
        if log_func:
            log_func(msg)

    edits, errors = parse_edit_lines(text)
    result.errors.extend(errors)
    if not edits:
        if not errors:
            result.errors.append("入力がありません。")
        result.ok = False
        return result

    client = client or JinjerClient()
    try:
        names, groups, schedules, day_offs = _collect_current_state(edits, client, log)
    except JinjerAPIError as exc:
        result.ok = False
        result.errors.append(f"jinjer からの現状取得に失敗しました: {exc}")
        return result

    plan, preview, plan_errors = build_plan(
        edits, schedules=schedules, day_offs=day_offs, names=names, groups=groups)
    result.preview = preview
    result.errors.extend(plan_errors)
    result.change_count = len(plan)
    result.fingerprint = plan_fingerprint(plan)

    if plan_errors:
        # 一部だけ書けても混乱のもとなので、エラーが1件でもあれば全体を止める
        # （直して再実行してもらう。書けた行と書けない行が混ざる方が事故に近い）
        result.ok = False
        log(f"エラー {len(plan_errors)} 件のため送信しません（全行そろってから実行してください）。")
        return result
    if not plan:
        result.ok = False
        result.errors.append("変更対象がありません（すべて変更なし）。")
        return result
    if dry_run:
        log(f"dry-run: {len(plan)} 行が書き込み対象です。内容を確認して実行してください。")
        return result

    if not expected_fingerprint or expected_fingerprint != result.fingerprint:
        result.ok = False
        result.errors.append(
            "プレビュー後に jinjer 側のスケジュールが変わっています（fingerprint 不一致）。"
            "もう一度プレビューから確認してください。")
        return result

    rows = build_import_rows(plan)
    csv_bytes = rows_to_csv_bytes(GENERIC_IMPORT_HEADER, rows)
    file_name = f"スケジュール開始編集_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    try:
        client.post_kintai_import(csv_bytes, file_name, executor_id or None)
    except JinjerAPIError as exc:
        result.ok = False
        result.errors.append(f"kintai-imports への投入に失敗しました: {exc}")
        return result
    log(f"投入予約 OK: {file_name}（{len(rows)} 行）。反映を待っています…")
    status = poll_func(client, file_name, log)
    result.import_status = status

    # 反映検証: 開始時刻が実際に変わったかを取り直して突き合わせる
    if status == "1":
        for p in plan:
            month = p["date_iso"][:7]
            try:
                after = client.get_work_schedules(p["emp"], month, store_id=p["store_id"])
            except JinjerAPIError:
                after = {}
            got = (after.get(p["date_iso"]) or {}).get("start", "")
            if t2m(got) != t2m(p["start"]):
                result.verify_ng.append(
                    f"{p['emp']} {p['name']} {p['date_iso']}: 開始が {got or '空'} のまま"
                    f"（期待 {p['start']}）")
        if result.verify_ng:
            result.ok = False
            log(f"⚠️ 反映検証 NG {len(result.verify_ng)} 件。jinjer 画面で確認してください。")
        else:
            log(f"反映検証 OK: {len(plan)} 行すべて新しい開始時刻になっています。")
    else:
        result.ok = False
        result.errors.append(
            "インポートが成功しませんでした"
            f"（status={status or 'タイムアウト/キュー未出現'}）。"
            "executor の勤怠管理者ロール・同時実行（1件まで）を確認してください。")

    result.snapshot_path = str(_write_snapshot(
        Path(output_dir), plan, preview, status or "未確定", file_name))
    log(f"実行ログ: {result.snapshot_path}")
    return result
