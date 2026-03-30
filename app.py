import os
import json
import uuid
import logging
from flask import Flask, render_template, request, jsonify, send_from_directory, Response
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

from config import Config
from services.jinjer_parser import parse_jinjer_csv
from services.timesheet_parser import parse_timesheet
from services.matcher import match
from services.excel_exporter import export_to_excel

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config.from_object(Config)

# アップロード/出力フォルダ作成
os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)
os.makedirs(Config.OUTPUT_FOLDER, exist_ok=True)

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


@app.route("/")
def index():
    api_key_set = bool(os.environ.get("ANTHROPIC_API_KEY"))
    return render_template("index.html", api_key_set=api_key_set)


@app.route("/upload", methods=["POST"])
def upload():
    """ファイルアップロード & 突合チェック実行（SSE対応）"""
    use_sse = request.headers.get("Accept") == "text/event-stream"

    jinjer_file = request.files.get("jinjer_csv")
    timesheet_files = request.files.getlist("timesheet_files")
    threshold = int(request.form.get("threshold", Config.DEFAULT_THRESHOLD_MINUTES))

    errors = []
    if not jinjer_file or jinjer_file.filename == "":
        errors.append("jinjer CSVファイルが選択されていません")
    elif not allowed_file(jinjer_file.filename, "jinjer"):
        errors.append("jinjer ファイルはCSV形式のみ対応しています")

    if not timesheet_files or all(f.filename == "" for f in timesheet_files):
        errors.append("勤務表ファイルが選択されていません")
    else:
        for f in timesheet_files:
            if f.filename and not allowed_file(f.filename, "timesheet"):
                errors.append(f"{f.filename} は未対応の形式です（対応: xlsx, xls, pdf, png, jpg, jpeg）")

    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    # --- リクエストコンテキストが有効なうちにファイルをディスクに保存する ---
    # SSE generator 内でファイルオブジェクトにアクセスすると
    # "read of closed file" エラーになるため、ここで先に保存する
    jinjer_path = os.path.join(Config.UPLOAD_FOLDER, f"jinjer_{uuid.uuid4().hex}.csv")
    jinjer_file.save(jinjer_path)

    saved_timesheet_paths = []  # (保存パス, 元ファイル名) のリスト
    valid_timesheet_files = [f for f in timesheet_files if f.filename]
    for ts_file in valid_timesheet_files:
        ts_path = os.path.join(Config.UPLOAD_FOLDER, f"ts_{uuid.uuid4().hex}_{ts_file.filename}")
        ts_file.save(ts_path)
        saved_timesheet_paths.append((ts_path, ts_file.filename))

    def generate():
        try:
            yield _sse_event("progress", {"message": "jinjer CSVを解析中..."})

            # jinjer CSV解析
            try:
                jinjer_df = parse_jinjer_csv(jinjer_path)
            except Exception as e:
                yield _sse_event("error", {"message": f"jinjer CSV解析エラー: {str(e)}"})
                return
            finally:
                _safe_remove(jinjer_path)

            yield _sse_event("progress", {"message": f"jinjer CSV解析完了: {len(jinjer_df)}件"})

            # 勤務表ファイル解析
            all_timesheets = []
            total = len(saved_timesheet_paths)

            for idx, (ts_path, ts_filename) in enumerate(saved_timesheet_paths, start=1):
                yield _sse_event("progress", {
                    "message": f"勤務表を解析中... ({idx}/{total}: {ts_filename})"
                })
                try:
                    ts_df = parse_timesheet(ts_path)
                    all_timesheets.append(ts_df)
                    yield _sse_event("progress", {
                        "message": f"解析完了 ({idx}/{total}): {ts_filename} → {len(ts_df)}件"
                    })
                except Exception as e:
                    logger.error(f"勤務表解析エラー ({ts_filename}): {e}")
                    yield _sse_event("progress", {
                        "message": f"スキップ ({idx}/{total}): {ts_filename} - {str(e)}"
                    })
                finally:
                    _safe_remove(ts_path)

            if not all_timesheets:
                yield _sse_event("error", {"message": "勤務表の解析に成功したファイルがありませんでした"})
                return

            yield _sse_event("progress", {"message": "突合処理中..."})
            timesheet_df = pd.concat(all_timesheets, ignore_index=True)
            result_df, unsubmitted_names = match(jinjer_df, timesheet_df, threshold)

            yield _sse_event("progress", {"message": "Excelファイルを生成中..."})
            excel_path = export_to_excel(result_df, threshold, unsubmitted_names=unsubmitted_names)
            excel_filename = os.path.basename(excel_path)

            # サマリー集計
            counts = result_df["判定"].value_counts().to_dict()
            summary = {
                "total": len(result_df),
                "ok": counts.get("OK", 0),
                "ng": counts.get("NG", 0),
                "caution": counts.get("要確認", 0),
                "missing": counts.get("データ欠損", 0),
            }

            # NG・要確認レコードをテーブル用に変換
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

            yield _sse_event("done", {
                "summary": summary,
                "table": table_data,
                "excel_filename": excel_filename,
                "unsubmitted": unsubmitted_names,
            })

        except Exception as e:
            logger.exception(f"処理エラー: {e}")
            yield _sse_event("error", {"message": f"処理中にエラーが発生しました: {str(e)}"})

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/download/<filename>")
def download(filename):
    """結果Excelファイルのダウンロード"""
    # パストラバーサル対策
    safe_name = os.path.basename(filename)
    return send_from_directory(Config.OUTPUT_FOLDER, safe_name, as_attachment=True)


@app.route("/api/status")
def api_status():
    return jsonify({"status": "ok", "api_key_set": bool(os.environ.get("ANTHROPIC_API_KEY"))})


def _sse_event(event_type, data):
    return f"event: {event_type}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


def _safe_remove(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
