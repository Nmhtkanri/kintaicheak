import anthropic
import base64
import csv
import io
import json
import time
import logging
import os
import re
from datetime import datetime, date, time as dt_time
import pandas as pd
from config import Config
from services.shift_legend_parser import (
    parse_with_legend_extraction,
    to_legend_dict_for_ui,
)
from services.shift_resolver import resolve_shifts

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """あなたは勤務表データを解析する専門家です。
与えられた勤務表から、以下の情報を抽出してJSON形式で返してください。

出力形式:
{
  "employee_name": "氏名（漢字）",
  "records": [
    {
      "date": "YYYY-MM-DD",
      "start_time": "HH:MM",
      "end_time": "HH:MM",
      "comment": "備考があれば記載、なければnull"
    }
  ]
}

ルール:
- 日付は必ず YYYY-MM-DD 形式にする
- 時刻は必ず HH:MM（24時間制）にする
- 休日・休暇の行は含めない（出退勤時刻がある行のみ抽出）
- 深夜勤務で日付をまたぐ場合、退勤時刻は翌日の時刻として "25:00" のように表記する
- 氏名が見つからない場合は "employee_name": "不明" とする
- 複数人分のデータがある場合は、配列で返す:
  [{"employee_name": "...", "records": [...]}, ...]
- JSONのみを返し、それ以外のテキストは含めないこと
"""


def _excel_to_text(filepath):
    """ExcelファイルをCSV風テキストに変換"""
    import openpyxl
    wb = openpyxl.load_workbook(filepath, data_only=True)
    lines = []
    for sheet in wb.worksheets:
        lines.append(f"=== シート: {sheet.title} ===")
        for row in sheet.iter_rows(values_only=True):
            row_vals = [str(v) if v is not None else "" for v in row]
            if any(v.strip() for v in row_vals):
                lines.append(",".join(row_vals))
    return "\n".join(lines)


def _pdf_to_text_or_bytes(filepath):
    """PDFからテキスト抽出。テキストが少ない場合はバイナリを返す"""
    import pdfplumber
    text_pages = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text_pages.append(text)
    full_text = "\n".join(text_pages)
    if len(full_text.strip()) < 100:
        # テキストが少ない → PDFバイナリとして返す
        with open(filepath, "rb") as f:
            return None, f.read()
    return full_text, None


def _pdf_first_page_to_png_bytes(filepath):
    """文字を持たないPDFを画像解析に回すため、1ページ目をPNG化する。"""
    try:
        import pypdfium2 as pdfium

        pdf = pdfium.PdfDocument(filepath)
        if len(pdf) == 0:
            return None
        page = pdf[0]
        bitmap = page.render(scale=4.0)
        image = bitmap.to_pil().convert("RGB")
        buf = io.BytesIO()
        image.save(buf, format="PNG")
        return buf.getvalue()
    except Exception as e:
        logger.warning("PDFの画像化に失敗しました: %s", e)
        return None


def _parse_with_claude(file_content, file_type, media_type=None):
    """Claude APIで勤務表を構造化する（リトライ付き）"""
    client = anthropic.Anthropic()

    if file_type == "text":
        messages = [{
            "role": "user",
            "content": f"以下の勤務表データを解析してください:\n\n{file_content}"
        }]
    elif file_type == "image":
        b64_data = base64.standard_b64encode(file_content).decode("utf-8")
        messages = [{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": media_type or "image/png",
                        "data": b64_data
                    }
                },
                {"type": "text", "text": "この勤務表を解析してください。"}
            ]
        }]
    elif file_type == "pdf":
        b64_data = base64.standard_b64encode(file_content).decode("utf-8")
        messages = [{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": b64_data
                    }
                },
                {"type": "text", "text": "この勤務表を解析してください。"}
            ]
        }]
    else:
        raise ValueError(f"未対応のfile_type: {file_type}")

    for attempt in range(3):
        try:
            response = client.messages.create(
                model=Config.ANTHROPIC_MODEL,
                max_tokens=Config.ANTHROPIC_MAX_TOKENS,
                system=SYSTEM_PROMPT,
                messages=messages
            )
            result_text = response.content[0].text.strip()
            # JSONブロック抽出
            json_match = re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", result_text)
            if json_match:
                result_text = json_match.group(1)
            return json.loads(result_text)
        except json.JSONDecodeError as e:
            logger.warning(f"JSON解析失敗 (試行{attempt+1}/3): {e}")
            if attempt < 2:
                # プロンプトを補強して再送
                messages[-1]["content"] = messages[-1]["content"] if isinstance(messages[-1]["content"], str) else messages[-1]["content"]
                if isinstance(messages[-1]["content"], list):
                    messages[-1]["content"][-1]["text"] = "この勤務表を解析してください。必ずJSONのみを返し、説明文は不要です。"
                time.sleep(2)
        except anthropic.RateLimitError:
            logger.warning(f"レート制限 (試行{attempt+1}/3)、3秒後にリトライ...")
            time.sleep(3)
        except anthropic.APIError as e:
            raise RuntimeError(f"Claude APIエラー: {e}")

    raise RuntimeError("Claude APIからの応答をJSONとして解析できませんでした")


def _normalize_records(claude_result, source_label="勤務表"):
    """Claude APIの結果を統一DataFrameに変換"""
    records = []

    def process_entry(entry):
        name = entry.get("employee_name", "不明")
        for rec in entry.get("records", []):
            try:
                d = datetime.strptime(rec["date"], "%Y-%m-%d").date()
            except (ValueError, KeyError):
                continue

            start = _parse_time_str(rec.get("start_time"))
            end = _parse_time_str(rec.get("end_time"))
            comment = rec.get("comment")
            if comment == "null":
                comment = None

            if start is None and end is None:
                continue

            records.append({
                "氏名": name,
                "日付": d,
                "出勤時刻": start,
                "退勤時刻": end,
                "コメント": comment,
                "データソース": source_label,
            })

    if isinstance(claude_result, list):
        for entry in claude_result:
            process_entry(entry)
    else:
        process_entry(claude_result)

    return pd.DataFrame(records, columns=["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"])


def _parse_time_str(value):
    """HH:MM文字列をdatetime.timeに変換（25:00等の深夜跨ぎ対応）"""
    if not value or value == "null":
        return None
    match = re.match(r"^(\d{1,2}):(\d{2})$", str(value).strip())
    if match:
        hour, minute = int(match.group(1)), int(match.group(2))
        if hour >= 24:
            hour = hour % 24
        try:
            return dt_time(hour, minute)
        except ValueError:
            return None
    return None


def _parse_excel_time(value):
    """Excelセルの時刻値を datetime.time に変換"""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, dt_time):
        return value
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if hasattr(value, "to_pydatetime"):
        try:
            return value.to_pydatetime().time().replace(second=0, microsecond=0)
        except Exception:
            pass
    if hasattr(value, "total_seconds"):
        total_minutes = int(value.total_seconds() // 60)
        return dt_time((total_minutes // 60) % 24, total_minutes % 60)

    value_str = str(value).strip()
    if not value_str or value_str in ("nan", "None", "NaT"):
        return None

    match = re.match(r"^(\d{1,2}):(\d{2})(?::\d{2})?$", value_str)
    if match:
        hour, minute = int(match.group(1)), int(match.group(2))
        try:
            return dt_time(hour % 24, minute)
        except ValueError:
            return None

    return None


def _parse_excel_date(value):
    """Excelセルの日付値を datetime.date に変換"""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if hasattr(value, "to_pydatetime"):
        try:
            return value.to_pydatetime().date()
        except Exception:
            pass

    value_str = str(value).strip()
    if not value_str or value_str in ("nan", "None", "NaT"):
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(value_str, fmt).date()
        except ValueError:
            continue
    return None


def _read_text_auto_encoding(filepath):
    """テキストファイルを文字コード自動判定で読み込む"""
    last_error = None
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            with open(filepath, "r", encoding=encoding, newline="") as f:
                return f.read()
        except UnicodeDecodeError as e:
            last_error = e
    raise ValueError(f"テキストの文字コードを判別できませんでした: {last_error}")


def _read_csv_auto_encoding(filepath):
    """CSVを文字コード自動判定で読み込む"""
    last_error = None
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return pd.read_csv(filepath, encoding=encoding, dtype=object)
        except UnicodeDecodeError as e:
            last_error = e
    raise ValueError(f"CSVの文字コードを判別できませんでした: {last_error}")


def _sap_timesheet_df_from_table(df):
    """SAP/Fieldglass の勤怠取得フォーマット表を統一 DataFrame 化する"""
    required = {"スタッフ", "出勤時刻", "終了時刻", "時間エントリ日"}
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if not required.issubset(set(df.columns)):
        return None

    rows = []
    for _, row in df.iterrows():
        name = row.get("スタッフ")
        if name is None or str(name).strip() in ("", "nan", "None"):
            continue

        work_date = _parse_excel_date(row.get("時間エントリ日"))
        start = _parse_excel_time(row.get("出勤時刻"))
        end = _parse_excel_time(row.get("終了時刻"))
        if work_date is None or (start is None and end is None):
            continue

        rows.append({
            "氏名": str(name).strip(),
            "日付": work_date,
            "出勤時刻": start,
            "退勤時刻": end,
            "コメント": None,
            "データソース": "勤務表",
        })

    if not rows:
        return None

    return pd.DataFrame(rows, columns=["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"])


def _parse_sap_timesheet_file(filepath):
    """SAP/Fieldglass の勤怠取得フォーマットを直接 DataFrame 化する"""
    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".csv":
        try:
            return _sap_timesheet_df_from_table(_read_csv_auto_encoding(filepath))
        except Exception:
            return None

    if ext not in (".xlsx", ".xls"):
        return None

    try:
        xls = pd.ExcelFile(filepath)
    except Exception:
        return None

    frames = []
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=object)
        except Exception:
            continue

        parsed = _sap_timesheet_df_from_table(df)
        if parsed is not None:
            frames.append(parsed)

    if not frames:
        return None

    result = pd.concat(frames, ignore_index=True)
    return result[["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"]]


def _parse_estaffing_timesheet_text(filepath):
    """e-staffing の勤怠処理ツール用テキストを直接 DataFrame 化する"""
    if os.path.splitext(filepath)[1].lower() != ".txt":
        return None

    try:
        text = _read_text_auto_encoding(filepath)
    except Exception:
        return None

    rows = []
    current_name = None
    for fields in csv.reader(text.splitlines(), delimiter="\t"):
        if not fields:
            continue

        row_type = (fields[0] or "").strip()
        if row_type == "H":
            current_name = fields[5].strip() if len(fields) > 5 else None
            continue

        if row_type != "D" or not current_name:
            continue

        work_date = _parse_excel_date(fields[1] if len(fields) > 1 else None)
        start = _parse_excel_time(fields[3] if len(fields) > 3 else None)
        end = _parse_excel_time(fields[4] if len(fields) > 4 else None)
        if work_date is None or (start is None and end is None):
            continue

        comment = None
        if len(fields) > 7 and str(fields[7]).strip():
            comment = str(fields[7]).strip()

        rows.append({
            "氏名": current_name,
            "日付": work_date,
            "出勤時刻": start,
            "退勤時刻": end,
            "コメント": comment,
            "データソース": "勤務表",
        })

    if not rows:
        return None

    return pd.DataFrame(rows, columns=["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"])


def _estaffing_csv_df_from_table(df):
    """e-staffing の請求勤怠 CSV を統一 DataFrame 化する"""
    required = {"スタッフ氏名", "就業年月日", "開始時刻", "終了時刻"}
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if not required.issubset(set(df.columns)):
        return None

    rows = []
    for _, row in df.iterrows():
        name = row.get("スタッフ氏名")
        if name is None or str(name).strip() in ("", "nan", "None"):
            continue

        work_date = _parse_excel_date(row.get("就業年月日"))
        start = _parse_excel_time(row.get("開始時刻"))
        end = _parse_excel_time(row.get("終了時刻"))
        if work_date is None or (start is None and end is None):
            continue

        comment = None
        raw_comment = row.get("備考コメント")
        if raw_comment is not None and str(raw_comment).strip() not in ("", "nan", "None"):
            comment = str(raw_comment).strip()

        rows.append({
            "氏名": str(name).strip(),
            "日付": work_date,
            "出勤時刻": start,
            "退勤時刻": end,
            "コメント": comment,
            "データソース": "勤務表",
        })

    if not rows:
        return None

    return pd.DataFrame(rows, columns=["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"])


def _parse_estaffing_timesheet_csv(filepath):
    """e-staffing の請求勤怠 CSV を直接 DataFrame 化する"""
    if os.path.splitext(filepath)[1].lower() != ".csv":
        return None

    try:
        return _estaffing_csv_df_from_table(_read_csv_auto_encoding(filepath))
    except Exception:
        return None


def _parse_known_timesheet_file(filepath):
    """既知フォーマットをAI解析せずに直接読む"""
    for parser in (
        _parse_sap_timesheet_file,
        _parse_estaffing_timesheet_csv,
        _parse_estaffing_timesheet_text,
    ):
        df = parser(filepath)
        if df is not None:
            return df
    return None


def parse_timesheet(filepath, progress_callback=None):
    """
    勤務表ファイル（Excel/PDF/画像）を解析して統一DataFrameを返す

    progress_callback: 進捗通知用コールバック (任意)

    Note: 後方互換のため "direct" モード相当の挙動のみ。
          記号式シフトの判別が必要な場合は parse_timesheet_smart() を使う。
    """
    ext = os.path.splitext(filepath)[1].lower()
    direct_df = _parse_known_timesheet_file(filepath)
    if direct_df is not None:
        if progress_callback:
            progress_callback(f"勤怠データを直接解析しました: {len(direct_df)}件")
        return direct_df

    if ext in (".xlsx", ".xls"):
        if progress_callback:
            progress_callback("Excelファイルをテキスト変換中...")
        text = _excel_to_text(filepath)
        result = _parse_with_claude(text, "text")

    elif ext in (".csv", ".txt"):
        if progress_callback:
            progress_callback("テキストファイルを解析中...")
        text = _read_text_auto_encoding(filepath)
        result = _parse_with_claude(text, "text")

    elif ext == ".pdf":
        if progress_callback:
            progress_callback("PDFを解析中...")
        text, pdf_bytes = _pdf_to_text_or_bytes(filepath)
        if text:
            result = _parse_with_claude(text, "text")
        elif image_bytes := _pdf_first_page_to_png_bytes(filepath):
            result = _parse_with_claude(image_bytes, "image", media_type="image/png")
        else:
            result = _parse_with_claude(pdf_bytes, "pdf")

    elif ext in (".png", ".jpg", ".jpeg"):
        if progress_callback:
            progress_callback("画像ファイルを解析中...")
        media_type_map = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}
        with open(filepath, "rb") as f:
            img_bytes = f.read()
        result = _parse_with_claude(img_bytes, "image", media_type=media_type_map[ext])

    else:
        raise ValueError(f"未対応のファイル形式: {ext}")

    return _normalize_records(result)


def _load_file_for_claude(filepath):
    """ファイルから (content, file_type, media_type) を返す"""
    ext = os.path.splitext(filepath)[1].lower()

    if ext in (".xlsx", ".xls"):
        return _excel_to_text(filepath), "text", None

    if ext in (".csv", ".txt"):
        return _read_text_auto_encoding(filepath), "text", None

    if ext == ".pdf":
        text, pdf_bytes = _pdf_to_text_or_bytes(filepath)
        if text:
            return text, "text", None
        if image_bytes := _pdf_first_page_to_png_bytes(filepath):
            return image_bytes, "image", "image/png"
        return pdf_bytes, "pdf", None

    if ext in (".png", ".jpg", ".jpeg"):
        media_type_map = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}
        with open(filepath, "rb") as f:
            return f.read(), "image", media_type_map[ext]

    raise ValueError(f"未対応のファイル形式: {ext}")


def parse_timesheet_smart(filepath, progress_callback=None):
    """
    勤務表ファイルを自動でモード判定して解析する

    - direct モード: そのまま DataFrame を返す（既存 parse_timesheet と同じ）
    - code   モード: 凡例＋シフト表（生）を返す → ユーザー確認後に shift_resolver で解決

    Returns:
        {
          "mode": "direct" | "code",
          "df": pd.DataFrame,        # direct モードのときのみ
          "legend": [...],           # code モードのときのみ
          "employees": [...],        # code モードのときのみ
          "off_markers": [...],      # code モードのときのみ
          "year": int|None,
          "month": int|None,
          "filename": str,
        }
    """
    if progress_callback:
        progress_callback(f"ファイルを解析中... ({os.path.basename(filepath)})")

    direct_df = _parse_known_timesheet_file(filepath)
    if direct_df is not None:
        return {
            "mode": "direct",
            "df": direct_df,
            "filename": os.path.basename(filepath),
        }

    file_content, file_type, media_type = _load_file_for_claude(filepath)
    fallback_df = None
    try:
        result = parse_with_legend_extraction(file_content, file_type, media_type)
    except Exception:
        logger.exception("凡例対応の勤務表解析に失敗しました。時刻直書き専用解析へフォールバックします。")
        result = {"mode": "direct", "data": []}
        try:
            fallback_df = _normalize_records(_parse_with_claude(file_content, file_type, media_type))
        except Exception:
            logger.exception("時刻直書き専用解析へのフォールバックにも失敗しました。")
            raise

    mode = result.get("mode")
    filename = os.path.basename(filepath)

    if mode == "direct":
        # 既存の direct モード形式に変換して DataFrame 化
        df = _normalize_records(result.get("data") or [])
        if df.empty and fallback_df is None:
            try:
                fallback_df = _normalize_records(_parse_with_claude(file_content, file_type, media_type))
            except Exception:
                logger.exception("空データのため時刻直書き専用解析へフォールバックしましたが失敗しました。")
        if df.empty and fallback_df is not None and not fallback_df.empty:
            df = fallback_df
        if df.empty:
            raise ValueError("AI解析は完了しましたが、出退勤時刻のある勤務行を抽出できませんでした")
        return {
            "mode": "direct",
            "df": df,
            "filename": filename,
        }

    if mode == "code":
        ui_data = to_legend_dict_for_ui(result)
        return {
            "mode": "code",
            "legend": ui_data.get("legend", []),
            "employees": ui_data.get("employees", []),
            "off_markers": ui_data.get("off_markers", []),
            "year": ui_data.get("year"),
            "month": ui_data.get("month"),
            "filename": filename,
        }

    raise RuntimeError(f"未対応の解析モード: {mode}")


def resolve_code_mode_to_df(legend, employees, off_markers=None, source_label="勤務表"):
    """凡例レビュー後の resolve（app.py から呼ばれる）

    Args:
        legend: list of dict（UI で編集された後の凡例）
        employees: list of dict（{name, shifts: [{date, code}]}）
        off_markers: 追加で休扱いとみなす記号セット
        source_label: DataFrame の "データソース" カラム値

    Returns:
        pd.DataFrame（matcher.py に流せる形式）
    """
    return resolve_shifts(legend, employees, off_markers=off_markers, source_label=source_label)
