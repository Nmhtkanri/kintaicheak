import os
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
        errors.append("勤務表ファイルが選択されていません")
    else:
        for f in timesheet_files:
            if f.filename and not allowed_file(f.filename, "timesheet"):
                errors.append(f"{f.filename} は未対応の形式です（対応: xlsx, xls, pdf, png, jpg, jpeg）")

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

            # 各勤務表をモード別に解析
            direct_dfs: list[pd.DataFrame] = []
            code_sheets: list[dict] = []  # 凡例レビュー対象
            total = len(saved_timesheet_paths)

            for idx, (ts_path, ts_filename) in enumerate(saved_timesheet_paths, start=1):
                yield _sse_event("progress", {
                    "message": f"勤務表を解析中... ({idx}/{total}: {ts_filename})"
                })
                try:
                    parsed = parse_timesheet_smart(ts_path)
                except Exception as e:
                    logger.error(f"勤務表解析エラー ({ts_filename}): {e}")
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
                    yield _sse_event("progress", {
                        "message": f"スキップ ({idx}/{total}): {ts_filename} - 不明なモード"
                    })

                _safe_remove(ts_path)

            if not direct_dfs and not code_sheets:
                yield _sse_event("error", {"message": "勤務表の解析に成功したファイルがありませんでした"})
                return

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
                template_csv = Config.JINJER_TEMPLATE_CSV_PATH
                for sheet in code_sheets:
                    sheet["template_match"] = match_legend_to_templates(
                        sheet["legend"], template_csv
                    )

                yield _sse_event("code_review_needed", {
                    "session_id": session_id,
                    "mode": mode,
                    "code_sheets": code_sheets,
                })
                return

            # direct モードのみ（match モード時のみ到達）→ 既存フロー通りに突合まで
            yield _sse_event("progress", {"message": "突合処理中..."})
            timesheet_df = pd.concat(direct_dfs, ignore_index=True)
            result_df, unsubmitted_names = match(jinjer_df, timesheet_df, threshold)

            yield _sse_event("progress", {"message": "Excelファイルを生成中..."})
            excel_path = export_to_excel(result_df, threshold, unsubmitted_names=unsubmitted_names)
            excel_filename = os.path.basename(excel_path)

            yield _sse_event("done", _build_done_payload(
                result_df, excel_filename, unsubmitted_names
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
                tm = match_legend_to_templates(legend, Config.JINJER_TEMPLATE_CSV_PATH)
                unmatched_total.extend(tm.get("unmatched", []))

            if not resolved_dfs:
                yield _sse_event("error", {"message": "解決後のレコードが空です"})
                return

            yield _sse_event("progress", {"message": "突合処理中..."})
            timesheet_df = pd.concat(resolved_dfs, ignore_index=True)
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
                        Config.JINJER_TEMPLATE_CSV_PATH,
                        out_path,
                    )
                    if gen.get("count", 0) > 0:
                        new_template_filename = os.path.basename(gen["path"])
                        new_template_count = gen["count"]

            done = _build_done_payload(result_df, excel_filename, unsubmitted_names)
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
                        template_csv_path=Config.JINJER_TEMPLATE_CSV_PATH,
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
                tm = match_legend_to_templates(legend, Config.JINJER_TEMPLATE_CSV_PATH)
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
                        unmatched, Config.JINJER_TEMPLATE_CSV_PATH, new_path
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


def _build_done_payload(result_df, excel_filename, unsubmitted_names):
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
            "judgment": str(row["判定"]),
            "detail": str(row["詳細"]) if row["詳細"] else "",
        })

    return {
        "summary": summary,
        "table": table_data,
        "excel_filename": excel_filename,
        "unsubmitted": unsubmitted_names,
    }


@app.route("/download/<filename>")
def download(filename):
    """結果Excelファイルのダウンロード"""
    safe_name = os.path.basename(filename)
    return send_from_directory(Config.OUTPUT_FOLDER, safe_name, as_attachment=True)


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


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
