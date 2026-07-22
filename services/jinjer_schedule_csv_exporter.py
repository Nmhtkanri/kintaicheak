"""jinjer スケジュール一括登録 CSV のエクスポータ

シフト記号モード（凡例＋シフト表）の解析結果から、
jinjer の「CSV登録用テンプレート」フォーマットの CSV を出力する。

CSV フォーマット:
  行1: "{年}年", "{月}月", 1, 2, ..., 31           ← 日付（ヘッダー）
  行2: "氏名", "従業員ID", 月, 火, 水, ...        ← 曜日
  行3〜: 氏名, 従業員ID, 値, 値, ...              ← 各従業員1行

各セルの値:
  - "明" / "明け休" を含むコード  → "休み" （明け休）
  - 凡例で is_off=True のコード   → 土曜=所、日曜=法、平日=所（A案）
  - 凡例の休扱い記号 / 空欄         → 同上（A案）
  - 凡例にあって時刻定義あり        → jinjer 雛形ID
  - 雛形にマッチしない            → 新規雛形 CSV と同じ ID 候補
"""

from __future__ import annotations

import calendar
import csv
import logging
import os
import re
from datetime import date, time
from typing import Iterable

from services.jinjer_template_matcher import (
    load_jinjer_templates,
    find_matching_template,
    suggest_template_id,
    _tpl_get,
)
from services.shift_resolver import normalize_legend, DEFAULT_OFF_MARKERS

logger = logging.getLogger(__name__)

# 曜日 (Mon=0 ... Sun=6) → 1文字漢字
WEEKDAY_KANJI = ["月", "火", "水", "木", "金", "土", "日"]

# 全日有休（有 / 有休 など）は休扱い（所/法）にせず jinjer の「一般」雛形(9:00~17:30)を入れる。
# 半休（AM有休 / PM有休 等）は対象外＝従来通り休扱いのまま。
FULL_DAY_PAID_LEAVE_CODES = {"有", "有休", "有給", "有給休暇", "年休", "年次有給休暇", "全有"}
_HALF_DAY_LEAVE_MARKERS = ("AM", "PM", "am", "pm", "ＡＭ", "ＰＭ", "午前", "午後", "半")
GENERAL_TEMPLATE_NAME = "一般"
GENERAL_TEMPLATE_START = "9:00"
GENERAL_TEMPLATE_END = "17:30"
AKE_REST_VALUE = "休み"


def _is_full_day_paid_leave(code: str, label: str = "") -> bool:
    """全日有休（有 / 有休 など）を表す記号か。

    AM/PM/午前/午後/半 を含む半休はここでは False（従来通り休扱い）。
    記号・ラベルのどちらかが全日有休と完全一致すれば True。
    """
    tokens = [str(code or "").strip(), str(label or "").strip()]
    joined = "".join(t for t in tokens if t)
    if any(h in joined for h in _HALF_DAY_LEAVE_MARKERS):
        return False
    return any(t in FULL_DAY_PAID_LEAVE_CODES for t in tokens if t)


def _resolve_general_template_id(templates: list[dict]) -> str:
    """jinjer の「一般」(9:00~17:30) 雛形ID を解決する。

    名前が「一般」の雛形を優先し、無ければ 9:00-17:30 の時刻一致でフォールバックする。
    見つからなければ空文字（→ 有休セルは従来通り休扱いになる）。
    """
    if not templates:
        return ""
    for t in templates:
        if str(_tpl_get(t, "＊スケジュール雛形名") or "").strip() == GENERAL_TEMPLATE_NAME:
            tid = str(_tpl_get(t, "＊スケジュール雛形ID") or "").strip()
            if tid:
                return tid
    tpl = find_matching_template(GENERAL_TEMPLATE_START, GENERAL_TEMPLATE_END, 0, templates)
    if tpl:
        tid = str(_tpl_get(tpl, "＊スケジュール雛形ID") or "").strip()
        if tid:
            return tid
    return ""


def _build_raw_legend_times(raw_legend) -> dict:
    """{code: (start_raw_str, end_raw_str, break_minutes)} を作る

    legend_normalized では HH:MM が % 24 されて 33:00 → 09:00 になってしまうため、
    深夜跨ぎシフトの統合処理用に生の HH:MM 文字列（24時超表記を保持）を別途保持する。
    """
    result = {}
    for entry in raw_legend or []:
        if not isinstance(entry, dict):
            continue
        code = str(entry.get("code") or "").strip()
        if not code:
            continue
        start_raw = str(entry.get("start_time") or "").strip() or None
        end_raw = str(entry.get("end_time") or "").strip() or None
        try:
            br = int(entry.get("break_minutes") or 0)
        except (TypeError, ValueError):
            br = 0
        result[code] = (start_raw, end_raw, br)
    return result


def _detect_employee_overnight_merges(
    day_map: dict[int, str],
    legend_normalized: dict,
    raw_times: dict,
    days_in_month: int,
) -> tuple[dict[int, dict], set[int]]:
    """1人分の day_map をスキャンして、深夜跨ぎで連続する2日のシフト統合候補を検出する

    判定条件（shift_resolver._merge_consecutive_overnight と整合）:
      - day N の凡例: end_time が 24:00（正規化後 00:00）
      - day N+1 の凡例: start_time が 24:00（正規化後 00:00）
      - 両方とも is_off=False で凡例にエントリあり

    Returns:
        merges: {day_n: merge_info}
        consumed_days: day N+1 として吸収された日の集合（それらの日は休扱いにする）
    """
    merges: dict[int, dict] = {}
    consumed: set[int] = set()
    midnight = time(0, 0)
    d = 1
    while d <= days_in_month:
        if d in consumed:
            d += 1
            continue
        c1 = day_map.get(d, "")
        e1 = legend_normalized.get(c1) if c1 else None
        if (
            not e1
            or e1.get("is_off")
            or e1.get("end_time") != midnight
            or e1.get("start_time") is None
        ):
            d += 1
            continue
        if d + 1 > days_in_month:
            d += 1
            continue
        c2 = day_map.get(d + 1, "")
        e2 = legend_normalized.get(c2) if c2 else None
        if (
            not e2
            or e2.get("is_off")
            or e2.get("start_time") != midnight
            or e2.get("end_time") is None
        ):
            d += 1
            continue
        s1_raw, _, b1 = raw_times.get(c1, (None, None, 0))
        _, e2_raw, b2 = raw_times.get(c2, (None, None, 0))
        if not s1_raw or not e2_raw:
            d += 1
            continue
        label1 = e1.get("label") or c1
        label2 = e2.get("label") or c2
        merges[d] = {
            "code1": c1,
            "code2": c2,
            "label1": label1,
            "label2": label2,
            "merged_code": f"{c1}+{c2}",
            "merged_label": f"{label1}+{label2}",
            "merged_start": s1_raw,
            "merged_end": e2_raw,
            "merged_break": b1 + b2,
        }
        consumed.add(d + 1)
        d += 2
    return merges, consumed


def _resolve_merged_cell_value(
    merge_info: dict,
    templates: list[dict],
) -> tuple[str, dict | None]:
    """統合シフトのセル値を決定し、未マッチなら新規雛形候補を返す

    Returns:
        (cell_value, unmatched_entry_or_None)
        - 既存雛形にマッチ → (雛形ID, None)   ← jinjer CSV インポートは ID 必須
        - マッチしない    → (merged_label, 新規雛形候補 dict)
    """
    start_raw = merge_info["merged_start"]
    end_raw = merge_info["merged_end"]
    label = merge_info["merged_label"]
    code = merge_info["merged_code"]
    break_minutes = merge_info["merged_break"]

    if templates and start_raw and end_raw:
        tpl = find_matching_template(start_raw, end_raw, break_minutes, templates)
        if tpl:
            tpl_id = _tpl_get(tpl, "＊スケジュール雛形ID")
            if tpl_id:
                return (tpl_id, None)
            # ID 列が空でも月次 CSV には名前でなく ID 候補を書く。
            return (suggest_template_id(code), None)

    unmatched_entry = {
        "code": code,
        "label": label,
        "start_time": start_raw or "",
        "end_time": end_raw or "",
        "break_minutes": break_minutes,
    }
    return (suggest_template_id(code), unmatched_entry)


_PAREN_NOTE_RE = re.compile(r"[（(][^（()）]*[)）]")


def _name_variants(name: str) -> list[str]:
    """氏名の照合バリエーション（元表記 → 空白除去 → 括弧注記除去 → 両方）。

    KDXシフト表の「亘（わたり）」のようなルビ・注記括弧を落として照合できるようにする。
    """
    raw = str(name or "").strip()
    variants: list[str] = []
    for base in (raw, _PAREN_NOTE_RE.sub("", raw)):
        for v in (base, re.sub(r"[\s　]+", "", base)):
            v = v.strip()
            if v and v not in variants:
                variants.append(v)
    return variants


def resolve_employee_id(name: str, name_to_id: dict[str, str]) -> str:
    """勤務表の氏名 → jinjer 従業員ID を解決する（失敗は空文字）。

    解決順:
      1. 完全一致（元表記 / 空白除去 / 括弧注記除去）
      2. 前方一致: jinjer 側の氏名キー（空白除去）が勤務表名で始まり、
         候補IDがちょうど1人のときだけ採用（「市川正」→「市川正人」）。
         3文字未満は誤マッチ防止のため前方一致しない（姓のみは 1 の専用キーで拾う）。

    name_to_id は build_name_to_id_map の出力（同姓など曖昧キーは含まれない）を想定。
    """
    if not name_to_id:
        return ""
    variants = _name_variants(name)

    for v in variants:
        if v in name_to_id:
            return name_to_id[v]

    stripped_map: dict[str, set[str]] = {}
    for k, eid in name_to_id.items():
        ks = re.sub(r"[\s　]+", "", str(k))
        if ks:
            stripped_map.setdefault(ks, set()).add(eid)
    for v in variants:
        ids = stripped_map.get(v)
        if ids and len(ids) == 1:
            return next(iter(ids))

    for v in variants:
        if len(v) < 3:
            continue
        ids = {eid for ks, id_set in stripped_map.items()
               if ks.startswith(v) for eid in id_set}
        if len(ids) == 1:
            return next(iter(ids))
    return ""


def annotate_unresolved_name(name: str, ambiguous_names: dict) -> str:
    """ID未解決の氏名に、同姓複数などの「なぜ引けないか」の注記を付ける（警告表示用）。"""
    for v in _name_variants(name):
        hits = (ambiguous_names or {}).get(v)
        if hits:
            cands = " / ".join(f"{full}({eid})" for eid, full in hits)
            return (f"{name}（同じ氏名の候補が複数いるため自動確定できません: {cands}"
                    f" — CSVの従業員ID列に正しいIDを入力してください）")
    return name


def _is_ake_code(code: str, label: str = "") -> bool:
    """「明け休」を意味する記号か"""
    if not code:
        return False
    c = str(code).strip()
    if "明" in c:
        return True
    if label and "明" in str(label):
        return True
    return False


def _off_value_for_weekday(weekday: int) -> str:
    """A案: 土曜=所、日曜=法、それ以外=所"""
    if weekday == 6:  # 日
        return "法"
    return "所"  # それ以外（土含む）→所


def build_legend_to_template_name(
    legend: list[dict],
    template_csv_path: str,
) -> dict[str, str]:
    """凡例コード → jinjer 雛形ID の辞書を作る

    jinjer のスケジュール CSV インポートは **雛形ID** を期待する（雛形名ではない）。
    雛形にマッチした場合は ＊スケジュール雛形ID（例: "K1", "1"）を返す。
    マッチしない場合は新規雛形 CSV と同じ ID 候補にフォールバックする。
    月次スケジュール CSV は jinjer の雛形IDを要求するため、凡例ラベルを
    そのまま入れるとアップロード時に弾かれる。

    Args:
        legend: shift_legend_parser から得られた凡例リスト
        template_csv_path: jinjer 雛形 CSV のパス

    Returns:
        {code: 雛形ID 文字列}
    """
    # NOTE: normalize_legend は >=24h を %24 で丸めるため "33:30" が "09:30" に化けて
    # 雛形マッチが失敗する。ここでは生の HH:MM 文字列のまま find_matching_template に渡す。
    templates = load_jinjer_templates(template_csv_path) if template_csv_path else []

    code_to_name: dict[str, str] = {}
    for entry in legend or []:
        if not isinstance(entry, dict):
            continue
        if entry.get("is_off"):
            continue  # 休扱いは別ロジック
        code = str(entry.get("code") or "").strip()
        if not code:
            continue
        # UI のプルダウンで雛形を明示選択した場合は、時刻マッチより優先してそのIDを使う
        chosen_id = str(entry.get("template_id") or "").strip()
        if chosen_id:
            code_to_name[code] = chosen_id
            continue
        start_raw = entry.get("start_time")
        end_raw = entry.get("end_time")
        label = entry.get("label") or code
        try:
            br = int(entry.get("break_minutes") or 0)
        except (TypeError, ValueError):
            br = 0

        if start_raw and end_raw and templates:
            tpl = find_matching_template(start_raw, end_raw, br, templates)
            logger.debug(
                "build_legend_to_template_name: code=%s start=%r end=%r break=%s -> id=%s name=%s",
                code, start_raw, end_raw, br,
                _tpl_get(tpl, "＊スケジュール雛形ID") if tpl else None,
                _tpl_get(tpl, "＊スケジュール雛形名") if tpl else None,
            )
            if tpl:
                tpl_id = _tpl_get(tpl, "＊スケジュール雛形ID")
                if tpl_id:
                    code_to_name[code] = tpl_id
                    continue
                # ID 列が空でも月次 CSV には名前でなく ID 候補を書く。
                code_to_name[code] = suggest_template_id(code)
                continue
        code_to_name[code] = suggest_template_id(code)

    return code_to_name


def _parse_iso_date(s) -> date | None:
    if not s:
        return None
    if isinstance(s, date):
        return s
    try:
        return date.fromisoformat(str(s).strip())
    except (ValueError, TypeError):
        return None


def _build_employee_day_map(employee: dict, days_in_month: int = 31) -> dict[int, str]:
    """1人分の shifts → {day(int): code(str)} に変換

    shifts は「表の左端から1日ずつ、空欄も code:'' として順番に記録したリスト」
    という契約（shift_legend_parser のプロンプト参照）。画像から年月が読めないと
    Claude は date を null で返すため、日付が無いシフトは**並び順（index+1）で日を
    割り当てる**。日付が明示されているシフトはそちらを優先し、日付なしシフトは
    既に埋まっていない日にだけ位置ベースで補完する。

    こうしないと、日付が null のシフトが全部捨てられ、その日が休扱いデフォルト
    （所/法）に化けてスケジュールが反映されない。
    """
    result: dict[int, str] = {}
    shifts = employee.get("shifts") or []

    # 1) 明示的な日付を持つシフトを優先で配置（こちらが正）
    dated_days: set[int] = set()
    dateless: list[tuple[int, str]] = []
    for idx, shift in enumerate(shifts):
        if not isinstance(shift, dict):
            continue
        code = str(shift.get("code") or "").strip()
        d = _parse_iso_date(shift.get("date"))
        if d is not None:
            if 1 <= d.day <= days_in_month:
                result[d.day] = code
                dated_days.add(d.day)
        else:
            dateless.append((idx, code))

    # 2) 日付が無いシフトは並び順（1始まりの日）で補完（既存の日は上書きしない）
    for idx, code in dateless:
        day = idx + 1
        if 1 <= day <= days_in_month and day not in dated_days:
            result[day] = code

    return result


def _resolve_cell_value(
    code: str,
    day_obj: date,
    legend_normalized: dict,
    code_to_template_name: dict[str, str],
    off_markers: set[str],
    general_template_id: str = "",
) -> str:
    """1セルの最終的な書き込み値を決定する"""
    weekday = day_obj.weekday()  # Mon=0 ... Sun=6
    off_default = _off_value_for_weekday(weekday)

    code = (code or "").strip()
    label = ""
    entry = legend_normalized.get(code) if code else None
    if entry:
        label = entry.get("label") or ""

    # 0) 全日有休（有 / 有休） → 一般(9:00~17:30) 雛形ID。
    #    休扱い(所/法)より優先する。半休(AM/PM)は対象外＝下の休扱いへ流れる。
    if general_template_id and _is_full_day_paid_leave(code, label):
        return general_template_id

    # 1) 明け休 → "休み"
    if _is_ake_code(code, label):
        return AKE_REST_VALUE

    # 2) 凡例で is_off の場合
    if entry and entry.get("is_off"):
        return off_default

    # 3) 空欄 / 休扱い記号
    if not code or code in off_markers:
        return off_default

    # 4) 凡例にマッチ → 雛形名
    if code in code_to_template_name:
        return code_to_template_name[code]

    # 5) 凡例にも雛形にも無い → 記号そのまま
    return code


def export_jinjer_schedule_csv(
    legend: list[dict],
    employees: list[dict],
    year: int,
    month: int,
    name_to_id: dict[str, str],
    output_path: str,
    *,
    template_csv_path: str = "",
    off_markers: Iterable[str] | None = None,
    id_to_official_name: dict[str, str] | None = None,
) -> dict:
    """jinjer スケジュール登録 CSV を書き出す

    Args:
        legend: 凡例（shift_legend_parser の出力）
        employees: 従業員ごとの shifts（同上）
        year, month: 対象年月
        name_to_id: 氏名 → 従業員ID の辞書
        output_path: 出力 CSV パス
        template_csv_path: jinjer 雛形 CSV（マッチ用）
        off_markers: 追加の休扱い記号

    Returns:
        {
          "path": str,
          "rows": int,
          "missing_ids": list[str],   # ID 取得できなかった氏名
          "year": int, "month": int,
        }
    """
    if year is None or month is None:
        raise ValueError("year / month は必須です（凡例レビュー時に確認してください）")

    days_in_month = calendar.monthrange(year, month)[1]
    day_objs = [date(year, month, d) for d in range(1, days_in_month + 1)]

    legend_normalized = normalize_legend(legend)
    code_to_template_name = build_legend_to_template_name(legend, template_csv_path)
    raw_times = _build_raw_legend_times(legend)
    templates = load_jinjer_templates(template_csv_path) if template_csv_path else []
    # 全日有休セルに入れる「一般」(9:00~17:30) 雛形ID（無ければ空＝従来通り休扱い）
    general_template_id = _resolve_general_template_id(templates)

    markers = set(DEFAULT_OFF_MARKERS)
    if off_markers:
        for m in off_markers:
            if m is None:
                continue
            markers.add(str(m).strip())

    # ----- ヘッダー行を組み立て -----
    header1 = [f"{year}年", f"{month}月"] + [str(d.day) for d in day_objs]
    header2 = ["氏名", "従業員ID"] + [WEEKDAY_KANJI[d.weekday()] for d in day_objs]

    # ----- 各従業員の行を組み立て -----
    rows: list[list[str]] = []
    missing_ids: list[str] = []
    merge_log: list[dict] = []
    merged_unmatched: list[dict] = []
    merged_unmatched_seen: set[tuple] = set()

    id_to_official_name = id_to_official_name or {}

    for emp in employees:
        if not isinstance(emp, dict):
            continue
        name = (emp.get("name") or "不明").strip()
        emp_id = resolve_employee_id(name, name_to_id)
        if not emp_id:
            missing_ids.append(name)

        # jinjer のスケジュールインポートは 氏名+従業員ID の組で照合するため、
        # 氏名カラムは jinjer 登録の「姓のみ」に揃える（公式サンプル CSV 仕様）。
        # ID が引けなかった場合のみ、勤務表の表記をそのまま残す。
        official_name = id_to_official_name.get(emp_id) if emp_id else ""
        display_name = official_name or name

        day_map = _build_employee_day_map(emp, days_in_month)

        # 深夜跨ぎ統合の検出（同一従業員内のみ）
        emp_merges, consumed_days = _detect_employee_overnight_merges(
            day_map, legend_normalized, raw_times, days_in_month
        )

        cells: list[str] = []
        for d in day_objs:
            day_num = d.day
            if day_num in emp_merges:
                # 統合シフトの先頭日 → 統合後の雛形名を書き込む
                m = emp_merges[day_num]
                value, unmatched_entry = _resolve_merged_cell_value(m, templates)
                cells.append(value)
                if unmatched_entry:
                    key = (
                        unmatched_entry["code"],
                        unmatched_entry["start_time"],
                        unmatched_entry["end_time"],
                    )
                    if key not in merged_unmatched_seen:
                        merged_unmatched_seen.add(key)
                        merged_unmatched.append(unmatched_entry)
                merge_log.append({
                    "name": name,
                    "day_n": day_num,
                    "day_n_plus_1": day_num + 1,
                    "code1": m["code1"],
                    "code2": m["code2"],
                    "label1": m["label1"],
                    "label2": m["label2"],
                    "merged_label": m["merged_label"],
                    "merged_start": m["merged_start"],
                    "merged_end": m["merged_end"],
                    "cell_value": value,
                })
                logger.info(
                    "深夜跨ぎ統合: %s %d日(%s)+%d日(%s) → %s-%s [%s]",
                    name, day_num, m["code1"], day_num + 1, m["code2"],
                    m["merged_start"], m["merged_end"], value,
                )
            elif day_num in consumed_days:
                # day N+1 として吸収された日 → 休扱い（曜日に応じて 所/法）
                cells.append(_off_value_for_weekday(d.weekday()))
            else:
                code_for_day = day_map.get(day_num, "")
                value = _resolve_cell_value(
                    code_for_day,
                    d,
                    legend_normalized,
                    code_to_template_name,
                    markers,
                    general_template_id,
                )
                cells.append(value)

        rows.append([display_name, emp_id] + cells)

    # ----- CP932 で書き出し -----
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="cp932", newline="", errors="replace") as f:
        writer = csv.writer(f)
        writer.writerow(header1)
        writer.writerow(header2)
        for r in rows:
            writer.writerow(r)

    logger.info(
        "jinjer スケジュール CSV 出力: %s (%d 件 / ID欠落 %d 件 / 統合 %d 件)",
        output_path, len(rows), len(missing_ids), len(merge_log),
    )

    return {
        "path": output_path,
        "rows": len(rows),
        "missing_ids": missing_ids,
        "year": year,
        "month": month,
        "merges": merge_log,
        "merged_unmatched": merged_unmatched,
    }


# =============================================================================
# 打刻グループ別の自動分割エクスポート
# =============================================================================

# Windows ファイル名で禁止される文字
_WIN_INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def _sanitize_filename_part(s) -> str:
    """ファイル名に使える形に整形する（Windows 禁止文字を除去）"""
    if s is None:
        return "未分類"
    cleaned = _WIN_INVALID_FILENAME_CHARS.sub("_", str(s))
    cleaned = re.sub(r"\s+", "_", cleaned).strip("._ ")
    return cleaned or "未分類"


def _resolve_employee_id_for_group(
    emp: dict,
    name_to_id: dict[str, str],
) -> str:
    """resolve_employee_id と同じ流儀で従業員ID を引く（マッチに失敗したら空文字）"""
    if not isinstance(emp, dict):
        return ""
    name = (emp.get("name") or "").strip()
    if not name:
        return ""
    return resolve_employee_id(name, name_to_id)


def export_jinjer_schedule_csv_split(
    legend: list[dict],
    employees: list[dict],
    year: int,
    month: int,
    name_to_id: dict[str, str],
    attendance_group_map: dict[str, tuple[str, str]],
    output_dir: str,
    *,
    template_csv_path: str = "",
    off_markers: Iterable[str] | None = None,
    id_to_official_name: dict[str, str] | None = None,
    filename_prefix: str = "jinjerスケジュール",
    filename_hash: str = "",
) -> dict:
    """打刻グループ別に CSV を自動分割して書き出す

    jinjer の月次スケジュール一括登録 CSV は **アップロード先打刻グループに所属しない
    従業員を 1 行でも含むと「全行エラー」で全否定される** ため、対象月時点の打刻グループ
    ごとにファイルを分けて出力する。

    Args:
        legend, employees, year, month, name_to_id, off_markers,
        template_csv_path, id_to_official_name:
            ``export_jinjer_schedule_csv`` と同義
        attendance_group_map: ``{employee_id: (group_id, group_name)}``
            （``services.jinjer_api_client.fetch_attendance_groups_at`` の戻り値）
        output_dir: 出力ディレクトリ
        filename_prefix: 出力ファイル名のプレフィックス
        filename_hash: 末尾に付けるハッシュ（衝突回避用、任意）

    Returns:
        {
          "csv_files": [
            {"path", "filename", "rows", "year", "month",
             "attendance_group_id", "attendance_group_name", "missing_ids", "merges"},
            ...
          ],
          "missing_ids": list[str],            # 全グループ合算
          "merges": list[dict],                # 全グループ合算
          "merged_unmatched": list[dict],      # 全グループ合算
          "ungrouped": list[str],              # 打刻グループが取れなかった従業員氏名
        }
    """
    # 従業員を打刻グループ毎にバケツに分ける
    buckets: dict[tuple[str, str], list[dict]] = {}
    ungrouped: list[str] = []

    for emp in employees or []:
        if not isinstance(emp, dict):
            continue
        emp_id = _resolve_employee_id_for_group(emp, name_to_id)
        if emp_id and emp_id in attendance_group_map:
            gid, gname = attendance_group_map[emp_id]
        else:
            gid, gname = ("", "")

        if not gid:
            # ID が引けなかった or 打刻グループ未設定
            ungrouped.append(emp.get("name") or "(名無し)")
            # それでも CSV には出したいので "未分類" バケツに入れる
            buckets.setdefault(("", "未分類"), []).append(emp)
            continue

        buckets.setdefault((gid, gname), []).append(emp)

    csv_files: list[dict] = []
    all_missing_ids: list[str] = []
    all_merges: list[dict] = []
    all_merged_unmatched: list[dict] = []

    os.makedirs(output_dir, exist_ok=True)

    # 打刻グループ id 昇順に出力（"" は先頭になる→未分類が先）
    sorted_keys = sorted(buckets.keys(), key=lambda k: (k[0] == "", k[0], k[1]))

    for gid, gname in sorted_keys:
        bucket_employees = buckets[(gid, gname)]
        if not bucket_employees:
            continue

        safe_name = _sanitize_filename_part(gname or gid or "未分類")
        suffix = f"_{filename_hash}" if filename_hash else ""
        out_filename = f"{filename_prefix}_{safe_name}_{year}{month:02d}{suffix}.csv"
        out_path = os.path.join(output_dir, out_filename)

        result = export_jinjer_schedule_csv(
            legend=legend,
            employees=bucket_employees,
            year=year,
            month=month,
            name_to_id=name_to_id,
            output_path=out_path,
            template_csv_path=template_csv_path,
            off_markers=off_markers,
            id_to_official_name=id_to_official_name,
        )

        csv_files.append({
            "path": result["path"],
            "filename": os.path.basename(result["path"]),
            "rows": result["rows"],
            "year": result["year"],
            "month": result["month"],
            "attendance_group_id": gid,
            "attendance_group_name": gname or "未分類",
            "missing_ids": result.get("missing_ids", []),
            "merges": result.get("merges", []),
        })
        all_missing_ids.extend(result.get("missing_ids", []))
        all_merges.extend(result.get("merges", []))
        all_merged_unmatched.extend(result.get("merged_unmatched", []))

        logger.info(
            "打刻グループ別CSV出力: gid=%s name=%s rows=%d -> %s",
            gid, gname, result["rows"], out_path,
        )

    return {
        "csv_files": csv_files,
        "missing_ids": all_missing_ids,
        "merges": all_merges,
        "merged_unmatched": all_merged_unmatched,
        "ungrouped": ungrouped,
    }
