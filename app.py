import os
import re
import json
import uuid
import logging
import pickle
from flask import Flask, render_template, request, jsonify, send_from_directory, Response
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

from config import Config
from services.jinjer_parser import parse_jinjer_csv
from services.timesheet_parser import (
    parse_timesheet_smart,
    resolve_code_mode_to_df,
)
from services.matcher import match
from services.excel_exporter import export_to_excel
from services.jinjer_template_matcher import (
    match_legend_to_templates,
    generate_new_templates_csv,
    load_jinjer_templates,
    _tpl_get,
)
from services.jinjer_api_client import (
    fetch_employee_id_map,
    fetch_attendance_groups_at,
    JinjerAPIError,
)
from services.jinjer_schedule_csv_exporter import (
    export_jinjer_schedule_csv,
    export_jinjer_schedule_csv_split,
)
from services.multi_year_shift_parser import parse_structured_files

import threading

# 月次集約 MVP（quick_compare / quick_export）
from pathlib import Path as _Path
from quick_compare import run_quick_compare
from quick_export import run_quick_export
from services.kintai_import_runner import run_api_import
from services.schedule_import_runner import run_schedule_api_import
from services.batch_runner import run_batch_compare
from services.expense_check import run_telework_export
from services.keihi_summary import run_keihi_integration

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config.from_object(Config)

# アップロード/出力フォルダ作成
os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)
os.makedirs(Config.OUTPUT_FOLDER, exist_ok=True)
os.makedirs(Config.SHIFT_SESSION_FOLDER, exist_ok=True)

# APIキーチェック
if not os.environ.get("ANTHROPIC_API_KEY"):
    logger.warning(
        "ANTHROPIC_API_KEY が設定されていません。"
        ".env ファイルに ANTHROPIC_API_KEY=your_key_here を設定してください。"
    )


def allowed_file(filename, file_type):
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    return ext in Config.ALLOWED_EXTENSIONS.get(file_type, set())


def format_time(t):
    if t is None:
        return ""
    try:
        return t.strftime("%H:%M")
    except Exception:
        return str(t) if str(t) not in ["nan", "None"] else ""


def format_date(d):
    if d is None:
        return ""
    try:
        return d.strftime("%Y/%m/%d")
    except Exception:
        return str(d)


# =============================================================================
# セッション保存（凡例レビュー → resolve のため、サーバー側に一時保持する）
# =============================================================================

def _save_session(session_id: str, payload: dict):
    """セッションデータを pickle で保存（jinjer_df + 各シート凡例）"""
    path = os.path.join(Config.SHIFT_SESSION_FOLDER, f"{session_id}.pkl")
    with open(path, "wb") as f:
        pickle.dump(payload, f)


def _load_session(session_id: str) -> dict | None:
    path = os.path.join(Config.SHIFT_SESSION_FOLDER, f"{session_id}.pkl")
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return pickle.load(f)


def _drop_session(session_id: str):
    path = os.path.join(Config.SHIFT_SESSION_FOLDER, f"{session_id}.pkl")
    _safe_remove(path)


def _apply_single_supplemental_legend(code_sheets: list[dict]) -> list[dict]:
    """別画像の凡例を、凡例なしのコード表へ安全に補う。"""
    legend_sources = [sheet for sheet in code_sheets if sheet.get("legend")]
    legendless_employee_sheets = [
        sheet
        for sheet in code_sheets
        if sheet.get("employees") and not sheet.get("legend")
    ]
    if len(legend_sources) != 1 or not legendless_employee_sheets:
        return code_sheets

    source = legend_sources[0]
    source_legend = [
        dict(entry)
        for entry in source.get("legend") or []
        if isinstance(entry, dict)
    ]
    source_markers = [
        str(marker).strip()
        for marker in source.get("off_markers") or []
        if marker is not None
    ]
    if not source_legend:
        return code_sheets

    for sheet in legendless_employee_sheets:
        sheet["legend"] = [dict(entry) for entry in source_legend]
        sheet["off_markers"] = list(dict.fromkeys([
            *[str(marker).strip() for marker in sheet.get("off_markers") or [] if marker is not None],
            *source_markers,
        ]))
        logger.info(
            "別画像の凡例を補完: schedule=%s legend=%s",
            sheet.get("filename"),
            source.get("filename"),
        )

    if not source.get("employees"):
        return [sheet for sheet in code_sheets if sheet is not source]
    return code_sheets


# =============================================================================
# ルート
# =============================================================================

@app.route("/")
def index():
    api_key_set = bool(os.environ.get("ANTHROPIC_API_KEY"))
    return render_template("index.html", api_key_set=api_key_set)


@app.route("/upload", methods=["POST"])
def upload():
    """ファイルアップロード & 解析（SSE対応）

    モード:
      - mode=match     (デフォルト) jinjer 突合モード
      - mode=csv_export                jinjer インポート用 CSV 出力モード（jinjer CSV 不要）

    モード判定で:
      - direct モード（時刻直書き）→ そのまま突合まで実行（match モードのみ）
      - code   モード（シフト記号）→ 凡例＋シフト表を返してユーザー確認待ち

    全ファイルが direct モードの場合のみ "done" を返す。
    1つでも code モードがあれば、すべて "code_review_needed" を返してフロントに引き継ぐ。
    """
    mode = request.form.get("mode", "match")
    jinjer_file = request.files.get("jinjer_csv")
    timesheet_files = request.files.getlist("timesheet_files")
    threshold = int(request.form.get("threshold", Config.DEFAULT_THRESHOLD_MINUTES))

    # 多年度横並びレイアウトを構造化パースする際に使う対象年月（任意）
    def _safe_int(v, default=None):
        try:
            return int(v) if v not in (None, "") else default
        except (TypeError, ValueError):
            return default
    target_year = _safe_int(request.form.get("target_year"))
    target_month = _safe_int(request.form.get("target_month"))

    errors = []
    # match モードは jinjer CSV 必須、csv_export モードは任意
    if mode == "match":
        if not jinjer_file or jinjer_file.filename == "":
            errors.append("jinjer CSVファイルが選択されていません")
        elif not allowed_file(jinjer_file.filename, "jinjer"):
            errors.append("jinjer ファイルはCSV形式のみ対応しています")
    else:  # csv_export
        if jinjer_file and jinjer_file.filename and not allowed_file(jinjer_file.filename, "jinjer"):
            errors.append("jinjer ファイルはCSV形式のみ対応しています")

    if not timesheet_files or all(f.filename == "" for f in timesheet_files):
        errors.append("請求勤怠ファイルが選択されていません")
    else:
        for f in timesheet_files:
            if f.filename and not allowed_file(f.filename, "timesheet"):
                errors.append(f"{f.filename} は未対応の形式です（対応: xlsx, xls, xlsb, pdf, png, jpg, jpeg）")

    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    # --- リクエストコンテキストが切れる前にファイルを保存 ---
    jinjer_path = None
    if jinjer_file and jinjer_file.filename:
        jinjer_path = os.path.join(Config.UPLOAD_FOLDER, f"jinjer_{uuid.uuid4().hex}.csv")
        jinjer_file.save(jinjer_path)

    saved_timesheet_paths = []
    valid_timesheet_files = [f for f in timesheet_files if f.filename]
    for ts_file in valid_timesheet_files:
        ts_path = os.path.join(Config.UPLOAD_FOLDER, f"ts_{uuid.uuid4().hex}_{ts_file.filename}")
        ts_file.save(ts_path)
        saved_timesheet_paths.append((ts_path, ts_file.filename))

    def generate():
        nonlocal saved_timesheet_paths
        try:
            jinjer_df = pd.DataFrame()
            if jinjer_path:
                yield _sse_event("progress", {"message": "jinjer CSVを解析中..."})
                try:
                    jinjer_df = parse_jinjer_csv(jinjer_path)
                except Exception as e:
                    yield _sse_event("error", {"message": f"jinjer CSV解析エラー: {str(e)}"})
                    return
                finally:
                    _safe_remove(jinjer_path)

                yield _sse_event("progress", {"message": f"jinjer CSV解析完了: {len(jinjer_df)}件"})

            # 各請求勤怠ファイルをモード別に解析
            direct_dfs: list[pd.DataFrame] = []
            code_sheets: list[dict] = []  # 凡例レビュー対象
            total = len(saved_timesheet_paths)

            # ----- 構造化パース（多年度横並び xlsx + シフトルール xlsx 専用 fast path）-----
            # target_year/month が指定されており、アップロードされた xlsx の中に
            # 多年度シフト表があれば、Claude を経由せず確定的に解析する。
            # 該当しないファイルや該当しないレイアウトはそのまま Claude フォールバックへ。
            if target_year and target_month and saved_timesheet_paths:
                try:
                    yield _sse_event("progress", {"message": "シフト表を構造化解析中..."})
                    structured_result = parse_structured_files(
                        [p for p, _ in saved_timesheet_paths],
                        target_year, target_month,
                    )
                except Exception as e:
                    logger.warning("構造化パース失敗、Claudeフォールバックします: %s", e)
                    structured_result = None

                if structured_result:
                    structured_sheets, consumed_paths = structured_result
                    consumed_set = set(consumed_paths)
                    for sheet in structured_sheets:
                        code_sheets.append({
                            "filename": sheet["filename"],
                            "year": sheet["year"],
                            "month": sheet["month"],
                            "legend": sheet["legend"],
                            "employees": sheet["employees"],
                            "off_markers": sheet["off_markers"],
                        })
                        sec = sheet.get("section_info") or {}
                        yield _sse_event("progress", {
                            "message": (
                                f"構造化解析完了: {sheet['filename']} "
                                f"→ {sheet['year']}年{sheet['month']}月 "
                                f"凡例 {len(sheet['legend'])}個 / "
                                f"従業員 {len(sheet['employees'])}人 "
                                f"(セクション{sec.get('section_index')} "
                                f"曜日マッチ {sec.get('weekday_matched', 0)}/"
                                f"{sec.get('weekday_total', 0)})"
                            )
                        })
                    # 消費されたパスは Claude フォールバックから除外し、tmp も削除
                    new_paths = []
                    for p, fn in saved_timesheet_paths:
                        if p in consumed_set:
                            _safe_remove(p)
                        else:
                            new_paths.append((p, fn))
                    saved_timesheet_paths = new_paths
                    total = len(saved_timesheet_paths)

            # モードに応じた表示名（スケジュールアップロード時は「スケジュール」）
            parse_noun = "スケジュール" if mode == "csv_export" else "請求勤怠"
            parse_failures = []
            for idx, (ts_path, ts_filename) in enumerate(saved_timesheet_paths, start=1):
                yield _sse_event("progress", {
                    "message": f"{parse_noun}を解析中... ({idx}/{total}: {ts_filename})"
                })
                try:
                    parsed = parse_timesheet_smart(ts_path)
                except Exception as e:
                    logger.error(f"請求勤怠解析エラー ({ts_filename}): {e}")
                    parse_failures.append(f"{ts_filename}: {str(e)}")
                    yield _sse_event("progress", {
                        "message": f"スキップ ({idx}/{total}): {ts_filename} - {str(e)}"
                    })
                    _safe_remove(ts_path)
                    continue

                # NOTE: ループ内では request mode (mode) を上書きせず別変数で保持
                parsed_mode = parsed.get("mode")
                if parsed_mode == "direct":
                    df = parsed.get("df")
                    if df is not None and not df.empty:
                        direct_dfs.append(df)
                        yield _sse_event("progress", {
                            "message": f"解析完了 ({idx}/{total}): {ts_filename} → 時刻直書き {len(df)}件"
                        })
                    else:
                        parse_failures.append(f"{ts_filename}: 解析結果が0件でした")
                        yield _sse_event("progress", {
                            "message": f"データなし ({idx}/{total}): {ts_filename}"
                        })
                elif parsed_mode == "code":
                    code_sheets.append({
                        "filename": ts_filename,
                        "year": parsed.get("year"),
                        "month": parsed.get("month"),
                        "legend": parsed.get("legend", []),
                        "employees": parsed.get("employees", []),
                        "off_markers": parsed.get("off_markers", []),
                    })
                    yield _sse_event("progress", {
                        "message": f"解析完了 ({idx}/{total}): {ts_filename} → 記号式（凡例 {len(parsed.get('legend', []))}個 / 従業員 {len(parsed.get('employees', []))}人）"
                    })
                else:
                    parse_failures.append(f"{ts_filename}: 不明な解析モード {parsed_mode}")
                    yield _sse_event("progress", {
                        "message": f"スキップ ({idx}/{total}): {ts_filename} - 不明なモード"
                    })

                _safe_remove(ts_path)

            if not direct_dfs and not code_sheets:
                detail = " / ".join(parse_failures) if parse_failures else "詳細理由を取得できませんでした"
                yield _sse_event("error", {
                    "message": f"{parse_noun}の解析に成功したファイルがありませんでした。詳細: {detail}"
                })
                return

            code_sheets = _apply_single_supplemental_legend(code_sheets)

            # CSV変換モードでは記号式（code）のみが対象
            if mode == "csv_export" and not code_sheets:
                yield _sse_event("error", {"message": "CSV変換モードには記号式の勤務表が必要です（時刻直書きのファイルは対象外）"})
                return

            # --- 分岐 ---
            if code_sheets:
                # 1つでも記号式があれば、レビュー画面で全部まとめて確認させる
                # (direct モードのものは直接 resolve せずセッションに退避)
                session_id = uuid.uuid4().hex
                payload = {
                    "mode": mode,
                    "jinjer_df": jinjer_df,
                    "direct_dfs": direct_dfs,
                    "code_sheets": code_sheets,
                    "threshold": threshold,
                }
                _save_session(session_id, payload)

                # 雛形マッチング（既存雛形があるかチェック→UI に表示）
                template_csv = Config.get_jinjer_template_csv_path()
                for sheet in code_sheets:
                    sheet["template_match"] = match_legend_to_templates(
                        sheet["legend"], template_csv
                    )

                # プルダウン用の jinjer 雛形一覧（名称・時刻つき）。記号→雛形を手動選択できる。
                available_templates = []
                for t in load_jinjer_templates(template_csv):
                    tid = _tpl_get(t, "＊スケジュール雛形ID")
                    if not tid:
                        continue
                    available_templates.append({
                        "id": tid,
                        "name": _tpl_get(t, "＊スケジュール雛形名") or tid,
                        "abbr": _tpl_get(t, "略称(3文字以内)"),
                        "start": _tpl_get(t, "＊出勤時間(0:00~47:59)"),
                        "end": _tpl_get(t, "＊退勤時間(0:00~47:59)"),
                    })

                yield _sse_event("code_review_needed", {
                    "session_id": session_id,
                    "mode": mode,
                    "code_sheets": code_sheets,
                    "available_templates": available_templates,
                })
                return

            # direct モードのみ（match モード時のみ到達）→ 既存フロー通りに突合まで
            yield _sse_event("progress", {"message": "突合処理中..."})
            timesheet_df = pd.concat(direct_dfs, ignore_index=True)
            warnings = _build_match_warnings(jinjer_df, timesheet_df)
            result_df, unsubmitted_names = match(jinjer_df, timesheet_df, threshold)

            yield _sse_event("progress", {"message": "Excelファイルを生成中..."})
            excel_path = export_to_excel(result_df, threshold, unsubmitted_names=unsubmitted_names)
            excel_filename = os.path.basename(excel_path)

            yield _sse_event("done", _build_done_payload(
                result_df, excel_filename, unsubmitted_names, warnings=warnings
            ))

        except Exception as e:
            logger.exception(f"処理エラー: {e}")
            yield _sse_event("error", {"message": f"処理中にエラーが発生しました: {str(e)}"})

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/resolve_and_match", methods=["POST"])
def resolve_and_match():
    """凡例レビュー後の最終突合実行

    POST body (JSON):
      {
        "session_id": "...",
        "sheets": [
          {
            "filename": "...",
            "legend": [{code, label, start_time, end_time, break_minutes, is_off}, ...],
            "off_markers": [...],
            "employees": [...]      // upload 時のものをそのまま返す
          }
        ]
      }
    """
    payload_in = request.get_json(force=True, silent=True) or {}
    session_id = payload_in.get("session_id")
    sheets_in = payload_in.get("sheets") or []

    session = _load_session(session_id) if session_id else None
    if not session:
        return jsonify({"success": False, "errors": ["セッションが見つかりません。最初からやり直してください。"]}), 400

    use_sse = request.headers.get("Accept") == "text/event-stream"

    def generate():
        try:
            jinjer_df = session["jinjer_df"]
            direct_dfs = session.get("direct_dfs") or []
            threshold = session.get("threshold", Config.DEFAULT_THRESHOLD_MINUTES)

            yield _sse_event("progress", {"message": "シフト記号を時刻に変換中..."})

            resolved_dfs: list[pd.DataFrame] = list(direct_dfs)
            unmatched_total: list[dict] = []

            for sheet in sheets_in:
                legend = sheet.get("legend") or []
                off_markers = sheet.get("off_markers") or []
                employees = sheet.get("employees") or []
                df = resolve_code_mode_to_df(
                    legend, employees, off_markers=off_markers,
                    source_label=sheet.get("filename") or "勤務表",
                )
                if not df.empty:
                    resolved_dfs.append(df)

                # 雛形マッチ→未マッチを集約
                template_csv = Config.get_jinjer_template_csv_path()
                tm = match_legend_to_templates(legend, template_csv)
                unmatched_total.extend(tm.get("unmatched", []))

            if not resolved_dfs:
                yield _sse_event("error", {"message": "解決後のレコードが空です"})
                return

            yield _sse_event("progress", {"message": "突合処理中..."})
            timesheet_df = pd.concat(resolved_dfs, ignore_index=True)
            warnings = _build_match_warnings(jinjer_df, timesheet_df)
            result_df, unsubmitted_names = match(jinjer_df, timesheet_df, threshold)

            yield _sse_event("progress", {"message": "Excelファイルを生成中..."})
            excel_path = export_to_excel(result_df, threshold, unsubmitted_names=unsubmitted_names)
            excel_filename = os.path.basename(excel_path)

            # 新規雛形 CSV 生成（未マッチ記号があった場合のみ）
            new_template_filename = None
            new_template_count = 0
            if unmatched_total:
                # 重複を除く（code, start, end）
                seen = set()
                unique_unmatched = []
                for u in unmatched_total:
                    key = (u.get("code"), u.get("start_time"), u.get("end_time"))
                    if key in seen:
                        continue
                    seen.add(key)
                    unique_unmatched.append(u)

                if unique_unmatched:
                    out_path = os.path.join(
                        Config.OUTPUT_FOLDER,
                        f"新規雛形_{uuid.uuid4().hex[:8]}.csv",
                    )
                    gen = generate_new_templates_csv(
                        unique_unmatched,
                        template_csv,
                        out_path,
                    )
                    if gen.get("count", 0) > 0:
                        new_template_filename = os.path.basename(gen["path"])
                        new_template_count = gen["count"]

            done = _build_done_payload(result_df, excel_filename, unsubmitted_names, warnings=warnings)
            done["new_template_filename"] = new_template_filename
            done["new_template_count"] = new_template_count
            yield _sse_event("done", done)

            _drop_session(session_id)

        except Exception as e:
            logger.exception(f"resolve_and_match エラー: {e}")
            yield _sse_event("error", {"message": f"処理中にエラーが発生しました: {str(e)}"})

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/export_jinjer_csv", methods=["POST"])
def export_jinjer_csv():
    """凡例レビュー後、jinjer インポート用 CSV を生成する（CSV変換モード）

    POST body (JSON):
      {
        "session_id": "...",
        "sheets": [{filename, legend, off_markers, employees, year, month}, ...]
      }
    """
    payload_in = request.get_json(force=True, silent=True) or {}
    session_id = payload_in.get("session_id")
    sheets_in = payload_in.get("sheets") or []

    session = _load_session(session_id) if session_id else None
    if not session:
        return jsonify({"success": False, "errors": ["セッションが見つかりません。最初からやり直してください。"]}), 400

    def generate():
        try:
            yield _sse_event("progress", {"message": "jinjer API から従業員一覧を取得中..."})
            try:
                name_to_id, id_to_official_name = fetch_employee_id_map()
            except JinjerAPIError as e:
                logger.warning("jinjer API 失敗: %s", e)
                yield _sse_event("progress", {
                    "message": f"⚠️ jinjer API 取得失敗: {e}（従業員IDは空欄で出力します）"
                })
                name_to_id = {}
                id_to_official_name = {}

            yield _sse_event("progress", {"message": f"取得完了: {len(name_to_id)}件の氏名→IDマップ"})

            output_files: list[dict] = []
            all_missing_ids: list[str] = []
            all_merges: list[dict] = []
            new_template_filename = None
            new_template_count = 0

            for sheet in sheets_in:
                filename = sheet.get("filename") or "勤務表"
                legend = sheet.get("legend") or []
                off_markers = sheet.get("off_markers") or []
                employees = sheet.get("employees") or []
                year = sheet.get("year")
                month = sheet.get("month")

                if year is None or month is None:
                    yield _sse_event("error", {
                        "message": f"{filename} に対象年月の情報がありません。凡例レビューで年月を指定してください。"
                    })
                    return

                year_i = int(year)
                month_i = int(month)
                target_date = f"{year_i:04d}-{month_i:02d}-01"

                yield _sse_event("progress", {
                    "message": f"CSV生成中: {filename} ({year_i}年{month_i}月)"
                })

                # この勤務表に登場する従業員の ID を集める（同姓ヒット回避: 厳密マッチ→ストリップ）
                emp_ids_for_sheet: list[str] = []
                emp_id_seen: set[str] = set()
                for emp in employees:
                    if not isinstance(emp, dict):
                        continue
                    name = (emp.get("name") or "").strip()
                    if not name:
                        continue
                    eid = name_to_id.get(name) or ""
                    if not eid:
                        import re as _re
                        stripped = _re.sub(r"\s+", "", name)
                        for k, v in name_to_id.items():
                            if _re.sub(r"\s+", "", k) == stripped:
                                eid = v
                                break
                    if eid and eid not in emp_id_seen:
                        emp_id_seen.add(eid)
                        emp_ids_for_sheet.append(eid)

                # 打刻グループ判定（対象月初時点）
                attendance_group_map: dict[str, tuple[str, str]] = {}
                if emp_ids_for_sheet:
                    yield _sse_event("progress", {
                        "message": f"  ▸ 打刻グループを判定中（{len(emp_ids_for_sheet)}名 / 基準日 {target_date}）..."
                    })
                    try:
                        attendance_group_map = fetch_attendance_groups_at(
                            emp_ids_for_sheet, target_date
                        )
                    except JinjerAPIError as e:
                        logger.warning("affiliations 取得失敗: %s", e)
                        yield _sse_event("progress", {
                            "message": f"  ⚠️ 打刻グループ取得失敗（単一ファイル出力にフォールバック）: {e}"
                        })
                        attendance_group_map = {}

                # 打刻グループ別に CSV を分割出力
                hash_suffix = uuid.uuid4().hex[:6]
                try:
                    split_result = export_jinjer_schedule_csv_split(
                        legend=legend,
                        employees=employees,
                        year=year_i,
                        month=month_i,
                        name_to_id=name_to_id,
                        attendance_group_map=attendance_group_map,
                        output_dir=Config.OUTPUT_FOLDER,
                        template_csv_path=Config.get_jinjer_template_csv_path(),
                        off_markers=off_markers,
                        id_to_official_name=id_to_official_name,
                        filename_hash=hash_suffix,
                    )
                except Exception as e:
                    logger.exception("CSV出力エラー: %s", e)
                    yield _sse_event("error", {"message": f"CSV出力失敗 ({filename}): {e}"})
                    return

                # ファイルごとに output_files に追加
                for cf in split_result.get("csv_files", []):
                    output_files.append({
                        "filename": cf["filename"],
                        "rows": cf["rows"],
                        "year": cf["year"],
                        "month": cf["month"],
                        "source": filename,
                        "attendance_group_id": cf.get("attendance_group_id", ""),
                        "attendance_group_name": cf.get("attendance_group_name", ""),
                    })
                    yield _sse_event("progress", {
                        "message": (
                            f"  ▸ 出力: {cf['filename']} "
                            f"[{cf.get('attendance_group_name') or '未分類'}] "
                            f"{cf['rows']}人分"
                        )
                    })

                all_missing_ids.extend(split_result.get("missing_ids", []))

                # 打刻グループが引けなかった従業員を警告として流す
                ungrouped = split_result.get("ungrouped", [])
                if ungrouped:
                    yield _sse_event("progress", {
                        "message": (
                            f"  ⚠️ 打刻グループが判定できなかった従業員 ({len(ungrouped)}名): "
                            f"{', '.join(ungrouped[:10])}"
                            + (" …" if len(ungrouped) > 10 else "")
                            + " → \"未分類\" CSV にまとめて出力しました。jinjer 画面で個別に登録してください。"
                        )
                    })

                # 深夜跨ぎ統合のログを SSE に流す
                for m in split_result.get("merges", []):
                    enriched = dict(m)
                    enriched["source"] = filename
                    enriched["year"] = year_i
                    enriched["month"] = month_i
                    all_merges.append(enriched)
                    yield _sse_event("progress", {
                        "message": (
                            f"  ▸ 深夜跨ぎ統合: {m['name']} "
                            f"{month_i}/{m['day_n']}({m['code1']}) + "
                            f"{month_i}/{m['day_n_plus_1']}({m['code2']}) "
                            f"→ {m['merged_start']}-{m['merged_end']} [{m['cell_value']}]"
                        )
                    })

                # 未マッチ雛形を集約（凡例ベース + 統合シフトベース）
                template_csv = Config.get_jinjer_template_csv_path()
                tm = match_legend_to_templates(legend, template_csv)
                unmatched = list(tm.get("unmatched", []))
                merged_unmatched = split_result.get("merged_unmatched") or []

                # 重複排除（code, start, end）でマージ
                seen = {(u.get("code"), u.get("start_time"), u.get("end_time")) for u in unmatched}
                for mu in merged_unmatched:
                    key = (mu.get("code"), mu.get("start_time"), mu.get("end_time"))
                    if key in seen:
                        continue
                    seen.add(key)
                    unmatched.append(mu)

                if unmatched:
                    new_path = os.path.join(
                        Config.OUTPUT_FOLDER,
                        f"新規雛形_{uuid.uuid4().hex[:8]}.csv",
                    )
                    gen = generate_new_templates_csv(
                        unmatched, template_csv, new_path
                    )
                    if gen.get("count", 0) > 0:
                        new_template_filename = os.path.basename(gen["path"])
                        new_template_count += gen["count"]

            # 重複除去
            unique_missing = sorted(set(all_missing_ids))

            yield _sse_event("csv_export_done", {
                "csv_files": output_files,
                "missing_ids": unique_missing,
                "merges": all_merges,
                "new_template_filename": new_template_filename,
                "new_template_count": new_template_count,
            })

            _drop_session(session_id)

        except Exception as e:
            logger.exception(f"export_jinjer_csv エラー: {e}")
            yield _sse_event("error", {"message": f"処理中にエラーが発生しました: {str(e)}"})

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


def _date_range_text(df):
    if df is None or df.empty or "日付" not in df.columns:
        return None
    dates = sorted({d for d in df["日付"].dropna().tolist() if pd.notna(d)})
    if not dates:
        return None
    return f"{format_date(dates[0])}〜{format_date(dates[-1])}"


def _build_match_warnings(jinjer_df, timesheet_df):
    warnings = []
    if (
        jinjer_df is None
        or timesheet_df is None
        or jinjer_df.empty
        or timesheet_df.empty
        or "日付" not in jinjer_df.columns
        or "日付" not in timesheet_df.columns
    ):
        return warnings

    jinjer_dates = {d for d in jinjer_df["日付"].dropna().tolist() if pd.notna(d)}
    timesheet_dates = {d for d in timesheet_df["日付"].dropna().tolist() if pd.notna(d)}
    if jinjer_dates and timesheet_dates and not (jinjer_dates & timesheet_dates):
        jinjer_range = _date_range_text(jinjer_df) or "不明"
        timesheet_range = _date_range_text(timesheet_df) or "不明"
        warnings.append(
            "jinjer CSVと請求勤怠の日付範囲が一致していません。"
            f"jinjer CSV: {jinjer_range}、請求勤怠: {timesheet_range}。"
            "同じ対象月のファイルを選択してください。"
        )
    return warnings


def _build_done_payload(result_df, excel_filename, unsubmitted_names, warnings=None):
    counts = result_df["判定"].value_counts().to_dict()
    summary = {
        "total": len(result_df),
        "ok": counts.get("OK", 0),
        "ng": counts.get("NG", 0),
        "caution": counts.get("要確認", 0),
        "missing": counts.get("データ欠損", 0),
    }

    ng_rows = result_df[result_df["判定"].isin(["NG", "要確認"])].copy()
    table_data = []
    for _, row in ng_rows.iterrows():
        table_data.append({
            "name": str(row["氏名"]) if row["氏名"] else "",
            "date": format_date(row["日付"]),
            "sheet_start": format_time(row["勤務表_出勤時刻"]),
            "jinjer_start": format_time(row["jinjer_出勤時刻"]),
            "start_diff": int(row["出勤差分(分)"]) if pd.notna(row.get("出勤差分(分)")) else "",
            "sheet_end": format_time(row["勤務表_退勤時刻"]),
            "jinjer_end": format_time(row["jinjer_退勤時刻"]),
            "end_diff": int(row["退勤差分(分)"]) if pd.notna(row.get("退勤差分(分)")) else "",
            "sheet_total_work": str(row.get("勤務表_総労働時間") or ""),
            "jinjer_total_work": str(row.get("jinjer_総労働時間") or ""),
            "total_work_diff": int(row["総労働差分(分)"]) if pd.notna(row.get("総労働差分(分)")) else "",
            "judgment": str(row["判定"]),
            "detail": str(row["詳細"]) if row["詳細"] else "",
        })

    return {
        "summary": summary,
        "table": table_data,
        "excel_filename": excel_filename,
        "unsubmitted": unsubmitted_names,
        "warnings": warnings or [],
    }


@app.route("/download/<filename>")
def download(filename):
    """結果Excelファイルのダウンロード"""
    safe_name = os.path.basename(filename)
    return send_from_directory(os.path.abspath(Config.OUTPUT_FOLDER), safe_name, as_attachment=True)


# =============================================================================
# 月次集約 MVP — quick_compare / quick_export
# =============================================================================

@app.route("/quick_compare", methods=["POST"])
def route_quick_compare():
    """突合結果xlsx 群 + jinjer CSV 群 → 差異一覧xlsx 生成

    フォーム:
      - kintai_dir : 突合結果xlsx ファイルまたは格納フォルダの絶対パス
      - jinjer_dir : jinjer 汎用データCSV ファイルまたは格納フォルダの絶対パス
      - month      : YYYY-MM
      - output_filename : 任意、未指定なら自動生成
    """
    kintai_dir_str = _clean_path_input(request.form.get("kintai_dir"))
    jinjer_dir_str = _clean_path_input(request.form.get("jinjer_dir"))
    application_csv_str = _clean_path_input(request.form.get("application_csv"))
    month_label = (request.form.get("month") or "").strip()
    output_filename = (request.form.get("output_filename") or "").strip()
    # 自動修正提案値（採用ラベル）の許容しきい値。未指定なら既定値。
    try:
        threshold = int(request.form.get("threshold", Config.DEFAULT_THRESHOLD_MINUTES))
    except (TypeError, ValueError):
        threshold = Config.DEFAULT_THRESHOLD_MINUTES

    errors = []
    if not kintai_dir_str:
        errors.append("手順1で出力した勤怠突合結果xlsx、またはその保存フォルダのパスを入力してください")
    if not jinjer_dir_str:
        errors.append("jinjer 汎用データCSV、またはその保存フォルダのパスを入力してください")
    if not application_csv_str:
        errors.append("jinjer 申請データ（打刻修正申請）CSVのパスを入力してください"
                      "（必須：打刻修正の申請理由を差異一覧に載せるため）")
    if not month_label:
        errors.append("対象月（YYYY-MM）を入力してください")

    kintai_dir = _Path(kintai_dir_str) if kintai_dir_str else None
    jinjer_dir = _Path(jinjer_dir_str) if jinjer_dir_str else None
    application_csv = _Path(application_csv_str) if application_csv_str else None
    if kintai_dir and not kintai_dir.exists():
        errors.append(f"勤怠突合結果xlsxまたはフォルダが見つかりません: {kintai_dir}")
    elif kintai_dir and kintai_dir.is_file() and kintai_dir.suffix.lower() != ".xlsx":
        errors.append(f"勤怠突合結果は .xlsx ファイルを指定してください: {kintai_dir}")
    if jinjer_dir and not jinjer_dir.exists():
        errors.append(f"jinjer 汎用データCSVまたはフォルダが見つかりません: {jinjer_dir}")
    elif jinjer_dir and jinjer_dir.is_file() and jinjer_dir.suffix.lower() not in [".csv", ".xlsx"]:
        errors.append(f"jinjer 汎用データは .csv または .xlsx ファイルを指定してください: {jinjer_dir}")
    # 申請データCSVは必須（2026-07-09〜）。存在もチェックする。
    if application_csv and not application_csv.exists():
        errors.append(f"申請データCSVまたはフォルダが見つかりません: {application_csv}")

    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    if not output_filename:
        ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"差異一覧_{month_label}_{ts}.xlsx"
    output_filename = os.path.basename(output_filename)  # 安全化
    output_filename = _ensure_extension(output_filename, ".xlsx")
    output_path = _Path(os.path.abspath(os.path.join(Config.OUTPUT_FOLDER, output_filename)))

    log_lines: list[str] = []
    def _log(msg: str) -> None:
        log_lines.append(msg)
        logger.info(msg)

    try:
        result = run_quick_compare(
            kintai_dir=kintai_dir,
            jinjer_dir=jinjer_dir,
            output_path=output_path,
            month_label=month_label,
            log_func=_log,
            application_csv=application_csv,
            threshold_minutes=threshold,
        )
    except Exception as e:
        logger.exception("quick_compare failed")
        return jsonify({"success": False, "errors": [str(e)], "log": log_lines}), 500

    payload = {
        "success": result.ok,
        "download_url": f"/download/{output_filename}" if result.ok else None,
        "output_filename": output_filename,
        "stats": {
            "diff_count": result.diff_count,
            "danger_count": result.danger_count,
            "warn_count": result.warn_count,
            "info_count": result.info_count,
            "kintai_rows_read": result.kintai_rows_read,
            "jinjer_rows_read": result.jinjer_rows_read,
            "name_map_size": result.name_map_size,
        },
        "logs": [{"severity": e.severity, "message": e.message, "source": e.source} for e in result.logs],
        "console": log_lines,
    }
    if not result.ok:
        payload["errors"] = [result.error] if result.error else ["差異一覧生成に失敗しました"]
        return jsonify(payload), 500
    return jsonify(payload)


@app.route("/batch_compare", methods=["POST"])
def route_batch_compare():
    """請求勤怠フォルダ + jinjer 汎用データ + 申請データ → 突合→差異一覧 を一括生成（機能3）

    フォーム:
      - timesheet_dir   : 請求勤怠フォルダ（または単一ファイル）の絶対パス
      - jinjer_dir      : jinjer 汎用データCSV（ファイル or フォルダ）
      - application_csv : jinjer 申請データ（打刻修正申請）CSV（必須）
      - month           : YYYY-MM
      - output_filename : 任意
    手順1(突合)を内部で実行し、手順2(差異一覧＋トリアージ)まで一括で行う。
    申請データCSVは打刻修正の「申請理由」を差異一覧に載せるために必須
    （2026-07-09 谷津さん指定。入れ忘れると判断材料が欠けるため）。
    """
    timesheet_dir_str = _clean_path_input(request.form.get("timesheet_dir"))
    jinjer_dir_str = _clean_path_input(request.form.get("jinjer_dir"))
    application_csv_str = _clean_path_input(request.form.get("application_csv"))
    month_label = (request.form.get("month") or "").strip()
    output_filename = (request.form.get("output_filename") or "").strip()
    # 自動修正提案値（採用ラベル）の許容しきい値。突合と同じしきい値を使う。
    try:
        threshold = int(request.form.get("threshold", Config.DEFAULT_THRESHOLD_MINUTES))
    except (TypeError, ValueError):
        threshold = Config.DEFAULT_THRESHOLD_MINUTES

    errors = []
    if not timesheet_dir_str:
        errors.append("請求勤怠フォルダ（またはファイル）のパスを入力してください")
    if not jinjer_dir_str:
        errors.append("jinjer 汎用データCSV（またはフォルダ）のパスを入力してください")
    if not application_csv_str:
        errors.append("jinjer 申請データ（打刻修正申請）CSVのパスを入力してください"
                      "（必須：打刻修正の申請理由を差異一覧に載せるため）")
    if not month_label:
        errors.append("対象月（YYYY-MM）を入力してください")

    timesheet_dir = _Path(timesheet_dir_str) if timesheet_dir_str else None
    jinjer_dir = _Path(jinjer_dir_str) if jinjer_dir_str else None
    application_csv = _Path(application_csv_str) if application_csv_str else None
    if timesheet_dir and not timesheet_dir.exists():
        errors.append(f"請求勤怠フォルダ/ファイルが見つかりません: {timesheet_dir}")
    if jinjer_dir and not jinjer_dir.exists():
        errors.append(f"jinjer 汎用データが見つかりません: {jinjer_dir}")
    if application_csv and not application_csv.exists():
        errors.append(f"申請データCSVが見つかりません: {application_csv}")
    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    if not output_filename:
        ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"差異一覧_{month_label}_{ts}.xlsx"
    output_filename = _ensure_extension(os.path.basename(output_filename), ".xlsx")
    output_path = _Path(os.path.abspath(os.path.join(Config.OUTPUT_FOLDER, output_filename)))

    log_lines: list[str] = []
    def _log(msg: str) -> None:
        log_lines.append(msg)
        logger.info(msg)

    try:
        result, skipped, unsubmitted = run_batch_compare(
            timesheet_dir=timesheet_dir,
            jinjer_dir=jinjer_dir,
            output_path=output_path,
            month_label=month_label,
            application_csv=application_csv,
            threshold_minutes=threshold,
            log_func=_log,
        )
    except Exception as e:
        logger.exception("batch_compare failed")
        return jsonify({"success": False, "errors": [str(e)], "log": log_lines}), 500

    payload = {
        "success": result.ok,
        "download_url": f"/download/{output_filename}" if result.ok else None,
        "output_filename": output_filename,
        "stats": {
            "diff_count": result.diff_count,
            "danger_count": result.danger_count,
            "warn_count": result.warn_count,
            "info_count": result.info_count,
            "skipped_count": len(skipped),
            "unsubmitted_count": len(unsubmitted),
        },
        "skipped": [{"file": n, "reason": r} for n, r in skipped[:50]],
        "unsubmitted": unsubmitted[:100],
        "logs": [{"severity": e.severity, "message": e.message, "source": e.source} for e in result.logs],
        "console": log_lines,
    }
    if not result.ok:
        payload["errors"] = [result.error] if result.error else ["一括生成に失敗しました"]
        return jsonify(payload), 500
    return jsonify(payload)


@app.route("/quick_export", methods=["POST"])
def route_quick_export():
    """差異一覧xlsx + jinjer CSV → アップロード用CSV 生成

    multipart:
      - diff_file       : 人間判断埋め込み済みの差異一覧xlsx（ファイルアップロード）
      - jinjer_dir      : 元 jinjer CSV ファイルまたはフォルダの絶対パス
      - output_filename : 任意
      - execute         : "1" なら本実行、未指定/0 は dry-run
    """
    diff_file = request.files.get("diff_file")
    jinjer_dir_str = _clean_path_input(request.form.get("jinjer_dir"))
    output_filename = (request.form.get("output_filename") or "").strip()
    execute_flag = request.form.get("execute") == "1"
    dry_run = not execute_flag

    errors = []
    if not diff_file or not diff_file.filename:
        errors.append("差異一覧xlsx を選択してください")
    if not jinjer_dir_str:
        errors.append("jinjer CSV、またはその保存フォルダのパスを入力してください")

    jinjer_dir = _Path(jinjer_dir_str) if jinjer_dir_str else None
    if jinjer_dir and not jinjer_dir.exists():
        errors.append(f"jinjer CSVまたはフォルダが見つかりません: {jinjer_dir}")
    elif jinjer_dir and jinjer_dir.is_file() and jinjer_dir.suffix.lower() != ".csv":
        errors.append(f"jinjer CSVは .csv ファイルを指定してください: {jinjer_dir}")

    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    diff_path = _Path(os.path.join(Config.UPLOAD_FOLDER, f"diff_{uuid.uuid4().hex}.xlsx"))
    diff_file.save(diff_path)

    if not output_filename:
        ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"jinjer_upload_{ts}.csv"
    output_filename = os.path.basename(output_filename)
    output_filename = _ensure_extension(output_filename, ".csv")
    output_path = _Path(os.path.abspath(os.path.join(Config.OUTPUT_FOLDER, output_filename)))

    log_lines: list[str] = []
    def _log(msg: str) -> None:
        log_lines.append(msg)
        logger.info(msg)

    try:
        result = run_quick_export(
            diff_xlsx=diff_path,
            jinjer_dir=jinjer_dir,
            output_path=output_path,
            dry_run=dry_run,
            log_func=_log,
        )
    except Exception as e:
        logger.exception("quick_export failed")
        _safe_remove(str(diff_path))
        return jsonify({"success": False, "errors": [str(e)], "log": log_lines}), 500
    finally:
        # アップロードした差異一覧xlsxは一時ファイル。書き出し成功後に削除
        pass

    stats = result.stats
    payload = {
        "success": result.ok,
        "dry_run": dry_run,
        "download_url": (f"/download/{output_filename}" if result.ok and not dry_run else None),
        "output_filename": output_filename if not dry_run else None,
        "stats": {
            "total_diff_rows": stats.total_diff_rows,
            "approved": stats.approved,
            "rejected": stats.rejected,
            "held": stats.held,
            "pending": stats.pending,
            "approved_punch_in": stats.approved_punch_in,
            "approved_punch_out": stats.approved_punch_out,
            "approved_break": stats.approved_break,
            "approved_total": stats.approved_total,
            "overwritten_punch_in": stats.overwritten_punch_in,
            "overwritten_sched_in": stats.overwritten_sched_in,
            "overwritten_punch_out": stats.overwritten_punch_out,
            "overwritten_break_start": stats.overwritten_break_start,
            "overwritten_break_end": stats.overwritten_break_end,
            "overwritten_break_total": stats.overwritten_break_total,
            "skipped_break": stats.skipped_break,
            "skipped_total": stats.skipped_total,
            "not_matched": stats.not_matched,
            "overwritten_finalized": stats.overwritten_finalized,
            "recovered_misplaced": stats.recovered_misplaced,
            "total_jinjer_rows": result.total_jinjer_rows,
        },
        "warnings": stats.warnings[:200],
        "console": log_lines,
    }
    _safe_remove(str(diff_path))

    if not result.ok:
        payload["errors"] = [result.error] if result.error else ["CSV 生成に失敗しました"]
        return jsonify(payload), 500
    return jsonify(payload)


# ----------------------------------------------------------------------
# 手順③ 後工程: jinjer への API 直接投入（kintai-imports）
#   画面の汎用データインポートがスケジュール列をサイレントに反映しない
#   不具合(2026-07-09)の恒久回避。docs/PLAN_手順3_API直接投入.md 参照。
#   投入はバッチ処理で数分かかるためスレッドで実行し、ジョブIDでポーリングする。
#   jinjer側の同時予約は1件までのため、実行中ジョブがあれば新規投入を拒否する。
# ----------------------------------------------------------------------
_api_import_jobs: dict = {}
_api_import_guard = threading.Lock()


@app.route("/api_import", methods=["POST"])
def route_api_import():
    """quick_export で生成済みのアップロード用CSVを API で jinjer へ投入する。

    フォーム:
      - csv_filename : OUTPUT_FOLDER 内のアップロード用CSVファイル名
      - execute      : "1" なら本投入、未指定/0 は dry-run（ガードと件数のみ）
      - month        : 検証対象月 YYYY-MM（省略時はCSVから自動判定）
    """
    csv_filename = os.path.basename((request.form.get("csv_filename") or "").strip())
    execute_flag = request.form.get("execute") == "1"
    month = (request.form.get("month") or "").strip()

    if not csv_filename:
        return jsonify({"success": False, "errors": ["CSVファイル名が指定されていません"]}), 400
    csv_path = _Path(os.path.abspath(os.path.join(Config.OUTPUT_FOLDER, csv_filename)))
    if not csv_path.is_file():
        return jsonify({"success": False,
                        "errors": [f"アップロード用CSVが見つかりません: {csv_filename}。先に手順③の本実行でCSVを生成してください"]}), 400

    with _api_import_guard:
        if any(not j["done"] for j in _api_import_jobs.values()):
            return jsonify({"success": False,
                            "errors": ["別のAPI投入が実行中です（jinjer側も同時予約1件の制限があります）。完了を待ってください"]}), 409
        job_id = uuid.uuid4().hex
        job = {"done": False, "ok": False, "log": [], "result": None}
        _api_import_jobs[job_id] = job

    def _worker():
        try:
            r = run_api_import(
                upload_csv=csv_path,
                output_dir=_Path(Config.OUTPUT_FOLDER),
                executor_id=Config.JINJER_IMPORT_EXECUTOR_ID,
                dry_run=not execute_flag,
                month=month,
                log_func=lambda m: job["log"].append(m),
            )
            report_name = os.path.basename(r.report_path) if r.report_path else ""
            job["ok"] = r.ok
            job["result"] = {
                "dry_run": r.dry_run,
                "total_rows": r.total_rows,
                "submitted_rows": r.submitted_rows,
                "excluded_count": len(r.excluded),
                "verified_ok": r.verified_ok,
                "verified_ng": r.verified_ng,
                "report_url": f"/download/{report_name}" if report_name else None,
            }
        except Exception as e:  # noqa: BLE001 — ジョブ内の失敗はログで返す
            logger.exception("api_import failed")
            job["log"].append(f"[ERROR] {e}")
        finally:
            job["done"] = True

    threading.Thread(target=_worker, daemon=True).start()
    logger.info("api_import job %s 開始 (csv=%s execute=%s)", job_id, csv_filename, execute_flag)
    return jsonify({"success": True, "job_id": job_id})


@app.route("/api_import_status/<job_id>")
def route_api_import_status(job_id):
    job = _api_import_jobs.get(job_id)
    if not job:
        return jsonify({"success": False, "errors": ["ジョブが見つかりません"]}), 404
    return jsonify({
        "success": True,
        "done": job["done"],
        "ok": job["ok"],
        "log": job["log"][-200:],
        "result": job["result"],
    })


# ----------------------------------------------------------------------
# スケジュールアップロードモード後工程: グリッドCSVを API でスケジュール投入
#   画面の月次スケジュールインポートが一部従業員にサイレントに反映されない
#   問題(2026-07: 11名で実測)の恒久回避。docs/PLAN_スケジュールAPI投入.md 参照。
#   ジョブ辞書・ロックは手順③(/api_import)と共有し、jinjer側の同時予約1件
#   制限がモード横断で効くようにする。ステータス取得も /api_import_status を共用。
# ----------------------------------------------------------------------
@app.route("/schedule_api_import", methods=["POST"])
def route_schedule_api_import():
    """生成済みグリッドCSVを差分プラン化して API で jinjer へ投入する。

    フォーム:
      - csv_filenames : OUTPUT_FOLDER 内のグリッドCSVファイル名のJSON配列
      - month         : 対象月 YYYY-MM（必須）
      - execute       : "1" なら投入、未指定/0 は差分プレビュー（dry-run）
      - fingerprint   : execute=1 のとき必須。差分プレビューが返した値
    """
    try:
        names = json.loads(request.form.get("csv_filenames") or "[]")
    except json.JSONDecodeError:
        return jsonify({"success": False, "errors": ["csv_filenames が不正です"]}), 400
    csv_filenames = [os.path.basename(str(n or "").strip()) for n in names]
    csv_filenames = [n for n in csv_filenames if n]
    month = (request.form.get("month") or "").strip()
    execute_flag = request.form.get("execute") == "1"
    fingerprint = (request.form.get("fingerprint") or "").strip()

    if not csv_filenames:
        return jsonify({"success": False, "errors": ["グリッドCSVが指定されていません"]}), 400
    if not re.match(r"^\d{4}-\d{2}$", month):
        return jsonify({"success": False, "errors": ["対象月(YYYY-MM)が不正です"]}), 400
    if execute_flag and not fingerprint:
        return jsonify({"success": False,
                        "errors": ["投入には差分プレビューのfingerprintが必要です。"
                                   "先に①差分プレビューを実行してください"]}), 400
    csv_paths = []
    for n in csv_filenames:
        p = _Path(os.path.abspath(os.path.join(Config.OUTPUT_FOLDER, n)))
        if not p.is_file():
            return jsonify({"success": False,
                            "errors": [f"グリッドCSVが見つかりません: {n}。"
                                       "先にスケジュールCSVを作成してください"]}), 400
        csv_paths.append(p)

    with _api_import_guard:
        if any(not j["done"] for j in _api_import_jobs.values()):
            return jsonify({"success": False,
                            "errors": ["別のAPI投入が実行中です（jinjer側も同時予約1件の"
                                       "制限があります）。完了を待ってください"]}), 409
        job_id = uuid.uuid4().hex
        job = {"done": False, "ok": False, "log": [], "result": None}
        _api_import_jobs[job_id] = job

    def _worker():
        try:
            r = run_schedule_api_import(
                grid_csvs=csv_paths,
                output_dir=_Path(Config.OUTPUT_FOLDER),
                month=month,
                executor_id=Config.JINJER_IMPORT_EXECUTOR_ID,
                dry_run=not execute_flag,
                expected_fingerprint=fingerprint if execute_flag else "",
                log_func=lambda m: job["log"].append(m),
            )
            report_name = os.path.basename(r.report_path) if r.report_path else ""
            job["ok"] = r.ok
            job["result"] = {
                "mode": "schedule",
                "dry_run": r.dry_run,
                "month": r.month,
                "plan_rows": r.plan_rows,
                "matched_rows": r.matched_rows,
                "plan_preview": [
                    {"emp": p["emp"], "name": p["name"], "date": p["date_iso"],
                     "youbi": p["youbi"], "kind": p["kind"], "cell": p["cell"],
                     "new": f"{p['start']}-{p['end']}",
                     "breaks": " ".join(f"{s}-{e}" for s, e in p["breaks"]) or "(なし)",
                     "cur": p["cur"]}
                    for p in r.plan[:300]
                ],
                "manual": r.manual[:300],
                "manual_count": len(r.manual),
                "fingerprint": r.fingerprint,
                "submitted_rows": r.submitted_rows,
                "verified_ok": r.verified_ok,
                "verified_ng": r.verified_ng,
                "report_url": f"/download/{report_name}" if report_name else None,
            }
        except Exception as e:  # noqa: BLE001 — ジョブ内の失敗はログで返す
            logger.exception("schedule_api_import failed")
            job["log"].append(f"[ERROR] {e}")
        finally:
            job["done"] = True

    threading.Thread(target=_worker, daemon=True).start()
    logger.info("schedule_api_import job %s 開始 (files=%s month=%s execute=%s)",
                job_id, csv_filenames, month, execute_flag)
    return jsonify({"success": True, "job_id": job_id})


@app.route("/expense_telework", methods=["POST"])
def route_expense_telework():
    """経費チェック: 指定月のテレワーク日数・出社日数を jinjer API から集計して Excel 出力。

    フォーム:
      - month           : YYYY-MM
      - output_filename : 任意
      - no_commute_xlsx : 任意。手動整備の「通勤費申請なし」リストExcelのパス。
                          指定すると整備済み「通勤費申請なし」シートを追加する。
    在籍者全員の勤怠を1名ずつ取得するため数分かかる。
    """
    month_label = (request.form.get("month") or "").strip()
    output_filename = (request.form.get("output_filename") or "").strip()
    commute_csv_str = _clean_path_input(request.form.get("commute_csv"))
    no_commute_str = _clean_path_input(request.form.get("no_commute_xlsx"))

    errors = []
    if not re.fullmatch(r"\d{4}-\d{2}", month_label):
        errors.append("対象月は YYYY-MM 形式で入力してください（例: 2026-05）")
    commute_csv = _Path(commute_csv_str) if commute_csv_str else None
    if commute_csv and not commute_csv.exists():
        errors.append(f"通勤費CSVが見つかりません: {commute_csv}")
    no_commute_xlsx = _Path(no_commute_str) if no_commute_str else None
    if no_commute_xlsx and not no_commute_xlsx.exists():
        errors.append(f"通勤費申請なしリストが見つかりません: {no_commute_xlsx}")
    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    if not output_filename:
        y, m = month_label.split("-")
        output_filename = f"テレワーク出社日数_{y}年{int(m):02d}月.xlsx"
    output_filename = _ensure_extension(os.path.basename(output_filename), ".xlsx")
    output_path = _Path(os.path.abspath(os.path.join(Config.OUTPUT_FOLDER, output_filename)))

    log_lines: list[str] = []
    def _log(msg: str) -> None:
        log_lines.append(msg)
        logger.info(msg)

    try:
        result = run_telework_export(
            month=month_label, output_path=output_path, log_func=_log,
            commute_csv=commute_csv, no_commute_xlsx=no_commute_xlsx,
        )
    except Exception as e:
        logger.exception("expense_telework failed")
        return jsonify({"success": False, "errors": [str(e)], "console": log_lines}), 500

    payload = {
        "success": result.ok,
        "download_url": f"/download/{output_filename}" if result.ok else None,
        "output_filename": output_filename if result.ok else None,
        "stats": {
            "employee_count": result.employee_count,
            "telework_total": result.telework_total,
            "no_data_count": result.no_data_count,
            "commute_count": result.commute_count,
            "no_commute_kept": result.no_commute_kept,
            "no_commute_removed": result.no_commute_removed,
        },
        "console": log_lines,
    }
    if not result.ok:
        payload["errors"] = [result.error] if result.error else ["テレワーク集計に失敗しました"]
        return jsonify(payload), 500
    return jsonify(payload)


@app.route("/expense_integration", methods=["POST"])
def route_expense_integration():
    """経費統合一覧表の生成（経費マクロ移植 P1a）。

    4ソース（jinjer / e-staffing / SAP / freee）の生CSVパスを受け取り、前処理・社員番号照合して
    34列の経費統合一覧表を作り、経路突合チェックシートと合わせて Excel 出力する。
    指定されたソースだけ取り込む（未指定は「未取込」）。

    フォーム:
      - jinjer_csv / estaffing_csv / sap_csv / freee_csv : 各生CSVパス（いずれか1つ以上必須）
      - output_filename : 任意
      - route_check     : "1"/"0"（既定 "1"。通勤経路は jinjer API から取得）
    """
    sources = {
        "jinjer_csv": _clean_path_input(request.form.get("jinjer_csv")),
        "estaffing_csv": _clean_path_input(request.form.get("estaffing_csv")),
        "sap_csv": _clean_path_input(request.form.get("sap_csv")),
        "freee_csv": _clean_path_input(request.form.get("freee_csv")),
    }
    output_filename = (request.form.get("output_filename") or "").strip()
    route_check = (request.form.get("route_check") or "1").strip() != "0"
    classify = (request.form.get("classify") or "1").strip() != "0"
    keywords_file_str = _clean_path_input(request.form.get("keywords_file"))
    import_template_str = _clean_path_input(request.form.get("import_template_csv"))

    errors = []
    paths: dict = {}
    for key, val in sources.items():
        if not val:
            paths[key] = None
            continue
        p = _Path(val)
        if not p.exists():
            errors.append(f"{key} が見つかりません: {p}")
        paths[key] = p
    keywords_file = _Path(keywords_file_str) if keywords_file_str else None
    if keywords_file and not keywords_file.exists():
        errors.append(f"キーワード設定ファイルが見つかりません: {keywords_file}")
    import_template_csv = _Path(import_template_str) if import_template_str else None
    if import_template_csv and not import_template_csv.exists():
        errors.append(f"インポートテンプレCSVが見つかりません: {import_template_csv}")
    if not any(paths.values()):
        errors.append("少なくとも1つのソースCSV（jinjer / e-staffing / SAP / freee）を指定してください。")
    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    if not output_filename:
        output_filename = "経費統合一覧表.xlsx"
    output_filename = _ensure_extension(os.path.basename(output_filename), ".xlsx")
    output_path = _Path(os.path.abspath(os.path.join(Config.OUTPUT_FOLDER, output_filename)))

    log_lines: list[str] = []
    def _log(msg: str) -> None:
        log_lines.append(msg)
        logger.info(msg)

    try:
        result = run_keihi_integration(
            output_path=output_path,
            jinjer_csv=paths["jinjer_csv"], estaffing_csv=paths["estaffing_csv"],
            sap_csv=paths["sap_csv"], freee_csv=paths["freee_csv"],
            route_check=route_check, classify=classify, keywords_file=keywords_file,
            import_template_csv=import_template_csv, log_func=_log,
        )
    except Exception as e:
        logger.exception("expense_integration failed")
        return jsonify({"success": False, "errors": [str(e)], "console": log_lines}), 500

    payload = {
        "success": result.ok,
        "download_url": f"/download/{output_filename}" if result.ok else None,
        "output_filename": output_filename if result.ok else None,
        "import_csv_url": f"/download/{result.import_csv_name}" if result.import_csv_name else None,
        "import_csv_name": result.import_csv_name or None,
        "import_preview": result.import_preview,
        "import_warnings": result.import_warnings,
        "stats": {
            "integrated_rows": result.integrated_rows,
            "source_counts": result.source_counts,
            "unmatched_emp": result.unmatched_emp,
            "route_summary": result.route_summary,
            "classify_summary": result.classify_summary,
        },
        "console": log_lines,
    }
    if not result.ok:
        payload["errors"] = [result.error] if result.error else ["統合一覧表の生成に失敗しました"]
        return jsonify(payload), 500
    return jsonify(payload)


@app.route("/expense_payroll_import", methods=["POST"])
def route_expense_payroll_import():
    """集計済みインポートCSVを jinjer給与へAPIインポートする（確認後実行）。

    フォーム:
      - import_csv_name : /expense_integration が outputs に出力したインポートCSVのファイル名
      - month           : 処理月 YYYY-MM（jinjer給与計算の対象月）
      - template_id     : 任意（既定 44450「経費APIインポート用」）
      - confirmed       : "1" 必須（プレビュー確認済みの明示）
    """
    from services.keihi_payroll_import import DEFAULT_TEMPLATE_ID, post_payroll_import
    from services.jinjer_api_client import JinjerClient, JinjerAPIError

    import_csv_name = os.path.basename((request.form.get("import_csv_name") or "").strip())
    month = (request.form.get("month") or "").strip()
    template_id = (request.form.get("template_id") or "").strip() or DEFAULT_TEMPLATE_ID
    confirmed = (request.form.get("confirmed") or "").strip() == "1"

    errors = []
    if not confirmed:
        errors.append("プレビュー確認済みフラグがありません（画面から実行してください）。")
    if not re.fullmatch(r"\d{4}-\d{2}", month):
        errors.append("処理月は YYYY-MM 形式で入力してください（例: 2026-07）")
    csv_path = _Path(os.path.abspath(os.path.join(Config.OUTPUT_FOLDER, import_csv_name))) if import_csv_name else None
    if not csv_path or not csv_path.exists():
        errors.append(f"インポートCSVが見つかりません: {import_csv_name}（先に統合一覧表を生成してください）")
    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    log_lines: list[str] = []
    def _log(msg: str) -> None:
        log_lines.append(msg)
        logger.info(msg)

    try:
        client = JinjerClient()
        client.authenticate()
    except JinjerAPIError as e:
        return jsonify({"success": False, "errors": [f"jinjer 認証に失敗しました: {e}"]}), 500

    try:
        result = post_payroll_import(
            client=client, month=month, csv_bytes=csv_path.read_bytes(),
            file_name=import_csv_name, template_id=template_id, log_func=_log,
        )
    except Exception as e:
        logger.exception("expense_payroll_import failed")
        return jsonify({"success": False, "errors": [str(e)], "console": log_lines}), 500

    payload = {
        "success": result.ok,
        "status": result.status,
        "month": month,
        "file_name": import_csv_name,
        "console": log_lines,
    }
    if not result.ok:
        payload["errors"] = [result.error] if result.error else ["インポート投入に失敗しました"]
        return jsonify(payload), 500
    return jsonify(payload)


@app.route("/api/status")
def api_status():
    return jsonify({"status": "ok", "api_key_set": bool(os.environ.get("ANTHROPIC_API_KEY"))})


def _sse_event(event_type, data):
    return f"event: {event_type}\ndata: {json.dumps(data, ensure_ascii=False, default=str)}\n\n"


def _safe_remove(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def _clean_path_input(s):
    """前後の空白と、貼り付け時に付くダブル/シングルクォートを除去する。

    Windows の「パスのコピー」は `"Z:\\...\\汎用データ.csv"` のように前後に
    ダブルクォートが付く。そのまま Path に渡すと存在しないパス扱いになる。

    - ダブルクォート(")は Windows のパスに使えない文字なので、前後にあれば
      **無条件に**除去する（片側だけ・複数個でもOK。例: `"Z:\\a` も剥がす）。
    - シングルクォート(')はフォルダ名に使える（例: O'Brien）ため、前後が対で
      囲まれているときだけ1組だけ剥がす。
    """
    s = (s or "").strip()
    # 前後のダブルクォートを全て除去（" はパスに使えないので安全）
    s = s.strip('"').strip()
    # 対のシングルクォートで囲まれている場合のみ1組剥がす
    if len(s) >= 2 and s[0] == s[-1] == "'":
        s = s[1:-1].strip()
    return s


def _ensure_extension(filename, extension):
    root, ext = os.path.splitext(filename)
    if not ext:
        return filename + extension
    if ext.lower() != extension:
        return root + extension
    return filename


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
