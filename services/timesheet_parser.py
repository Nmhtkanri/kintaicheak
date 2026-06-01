import anthropic
import base64
import csv
import io
import json
import time
import logging
import os
import re
import tempfile
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
    workbook_path = filepath
    converted_path = None
    if os.path.splitext(filepath)[1].lower() == ".xlsb":
        converted_path = _convert_xlsb_to_xlsx(filepath)
        if converted_path is None:
            raise ValueError("xlsbをxlsxに変換できませんでした")
        workbook_path = converted_path

    try:
        wb = openpyxl.load_workbook(workbook_path, data_only=True)
        lines = []
        for sheet in wb.worksheets:
            lines.append(f"=== シート: {sheet.title} ===")
            for row in sheet.iter_rows(values_only=True):
                row_vals = [str(v) if v is not None else "" for v in row]
                if any(v.strip() for v in row_vals):
                    lines.append(",".join(row_vals))
        return "\n".join(lines)
    finally:
        if converted_path:
            try:
                os.remove(converted_path)
            except OSError:
                pass


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


# 画像化したPDFページの長辺の上限(px)。
# Anthropic APIは長辺1568pxを超える画像を内部で縮小するため、これ以上大きくしても
# 精度は上がらず、base64後のサイズだけ膨らんでリクエスト上限(32MB)を超えて413になる。
_PDF_RENDER_MAX_LONG_EDGE = 2000


def _pdf_first_page_to_png_bytes(filepath):
    """文字を持たないPDFを画像解析に回すため、1ページ目をPNG化する。

    レンダリング倍率は固定せず、ページ実寸から長辺が _PDF_RENDER_MAX_LONG_EDGE 以内に
    収まるよう動的に決める。高解像度スキャンPDFでも画像が肥大化せず、Claude APIの
    リクエストサイズ上限(32MB)・画像最大辺(8000px)を超えない。
    """
    try:
        import pypdfium2 as pdfium

        pdf = pdfium.PdfDocument(filepath)
        if len(pdf) == 0:
            return None
        page = pdf[0]
        width_pt, height_pt = page.get_size()  # scale=1.0 相当のピクセル寸法
        native_long_edge = max(width_pt, height_pt)
        if native_long_edge <= 0:
            scale = 4.0
        else:
            # 長辺を _PDF_RENDER_MAX_LONG_EDGE に合わせる倍率。小さなPDFは拡大、
            # 巨大スキャンは縮小する。拡大しすぎを防ぐため上限は4.0倍。
            scale = min(4.0, _PDF_RENDER_MAX_LONG_EDGE / native_long_edge)
        bitmap = page.render(scale=scale)
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


def _convert_xlsb_to_xlsx(filepath):
    """Excel COMでxlsbを一時xlsxへ変換する。失敗時はNoneを返す。"""
    if os.path.splitext(filepath)[1].lower() != ".xlsb":
        return None

    fd, converted_path = tempfile.mkstemp(prefix="kintai_xlsb_", suffix=".xlsx")
    os.close(fd)

    excel = None
    workbook = None
    try:
        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(os.path.abspath(filepath), ReadOnly=True)
        workbook.SaveAs(os.path.abspath(converted_path), FileFormat=51)
        return converted_path
    except Exception:
        logger.exception("xlsbからxlsxへの一時変換に失敗しました: %s", filepath)
        try:
            os.remove(converted_path)
        except OSError:
            pass
        return None
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _extract_itone_name_from_filename(filepath):
    stem = os.path.splitext(os.path.basename(filepath))[0]
    match = re.search(r"(.+?)さん", stem)
    if match:
        return match.group(1).strip()
    match = re.search(r"・([^・()]+)\)_\d{6}$", stem)
    if match:
        return match.group(1).strip()
    return None


def _parse_itone_dispatch_timesheet_file(filepath):
    """IToneの派遣労働者勤務報告書を直接DataFrame化する。"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in (".xlsx", ".xlsb"):
        return None

    workbook_path = filepath
    converted_path = None
    if ext == ".xlsb":
        converted_path = _convert_xlsb_to_xlsx(filepath)
        if converted_path is None:
            return None
        workbook_path = converted_path

    try:
        import openpyxl

        wb = openpyxl.load_workbook(workbook_path, data_only=True)
        if "派遣労働者勤務報告書" not in wb.sheetnames:
            return None

        ws = wb["派遣労働者勤務報告書"]
        if str(ws["B12"].value or "").strip() != "日付":
            return None

        employee_name = ws["U6"].value or _extract_itone_name_from_filename(filepath)
        if not employee_name:
            return None
        employee_name = str(employee_name).strip()

        rows = []
        for row_idx in range(13, ws.max_row + 1):
            work_date = _parse_excel_date(ws.cell(row_idx, 2).value)  # B
            if work_date is None:
                continue

            start = _parse_excel_time(ws.cell(row_idx, 24).value)  # X
            end = _parse_excel_time(ws.cell(row_idx, 29).value)  # AC
            if start is None and end is None:
                continue

            comment = ws.cell(row_idx, 12).value or ws.cell(row_idx, 8).value  # L / H
            if comment is not None:
                comment = str(comment).strip() or None

            rows.append({
                "氏名": employee_name,
                "日付": work_date,
                "出勤時刻": start,
                "退勤時刻": end,
                "コメント": comment,
                "データソース": "勤務表",
            })

        if not rows:
            return None

        return pd.DataFrame(rows, columns=["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"])
    except Exception:
        logger.exception("ITone派遣労働者勤務報告書の解析に失敗しました: %s", filepath)
        return None
    finally:
        if converted_path:
            try:
                os.remove(converted_path)
            except OSError:
                pass


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


def _extract_fieldglass_name_from_filename(filepath):
    stem = os.path.splitext(os.path.basename(filepath))[0]
    match = re.search(r"timesheet_[A-Za-z]+[0-9]+(.+)$", stem, flags=re.IGNORECASE)
    if not match:
        return None

    tail = match.group(1)
    tail = tail.strip(" _,.-")
    if not tail:
        return None

    if re.search(r"[\u3040-\u30ff\u3400-\u9fff]", tail):
        return tail.replace("さん", "").strip()

    tail = tail.replace("_", " ").replace(",", " ").strip()
    tail = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", tail)
    tail = re.sub(r"\s+", " ", tail).strip()
    return tail.upper() if tail else None


def _parse_fieldglass_worker_name(text):
    match = re.search(r"\bWorker\s+([^(\n]+)\(", text)
    if not match:
        return None

    raw = match.group(1).strip()
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if len(parts) == 2:
        raw = f"{parts[1]} {parts[0]}"
    if re.fullmatch(r"[A-Za-z][A-Za-z\s.'-]*", raw):
        return raw.upper()
    return raw


def _fieldglass_date_for_month_day(month, day, period_start, period_end):
    for year in range(period_start.year - 1, period_end.year + 2):
        try:
            candidate = date(year, month, day)
        except ValueError:
            continue
        if period_start <= candidate <= period_end:
            return candidate
    return None


def _fieldglass_dates_from_day_line(line, period_start, period_end):
    dates = []
    for month, day, _weekday in re.findall(r"(\d{1,2})-(\d{2})\s+([A-Za-z]{3})", line):
        dates.append(_fieldglass_date_for_month_day(int(month), int(day), period_start, period_end))
    return dates


def _fieldglass_pad_left(values, dates):
    if len(values) >= len(dates):
        return values[:len(dates)]
    leading_outside_period = 0
    for d in dates:
        if d is None:
            leading_outside_period += 1
        else:
            break
    pad_count = min(len(dates) - len(values), leading_outside_period)
    return (["-"] * pad_count + values)[:len(dates)]


def _fieldglass_time_values(line, prefix, dates):
    body = line[len(prefix):].strip()
    values = re.findall(r"\d{1,2}:\d{2}|-", body)
    return _fieldglass_pad_left(values, dates)


def _fieldglass_duration_values(line, dates):
    body = line[len("Total"):].strip()
    values = [
        "-" if value == "-" else value
        for value in re.findall(r"-|\d+h\s+\d+m", body)
    ]
    in_period_count = sum(1 for d in dates if d is not None)
    if len(values) in (len(dates) + 1, in_period_count + 1):
        values = values[:-1]
    return _fieldglass_pad_left(values, dates)


def _is_zero_fieldglass_duration(value):
    if not value or value == "-":
        return True
    match = re.fullmatch(r"(\d+)h\s+(\d+)m", value.strip())
    return bool(match and int(match.group(1)) == 0 and int(match.group(2)) == 0)


def _parse_fieldglass_pdf_timesheet(filepath):
    """SAP Fieldglass の Time Sheet PDF を直接 DataFrame 化する"""
    if os.path.splitext(filepath)[1].lower() != ".pdf":
        return None

    try:
        text, pdf_bytes = _pdf_to_text_or_bytes(filepath)
    except Exception:
        return None
    if not text or "Time Sheet" not in text or "Time in/time out" not in text:
        return None

    period_match = re.search(r"\bPeriod\s+(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})", text)
    if not period_match:
        return None
    period_start = datetime.strptime(period_match.group(1), "%Y-%m-%d").date()
    period_end = datetime.strptime(period_match.group(2), "%Y-%m-%d").date()

    name = _extract_fieldglass_name_from_filename(filepath) or _parse_fieldglass_worker_name(text)
    if not name:
        return None

    rows = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        if line != "Time in/time out":
            continue

        day_line = None
        time_in_line = None
        time_out_line = None
        total_line = None
        for candidate in lines[idx + 1: idx + 12]:
            if candidate.startswith("Day "):
                day_line = candidate
            elif candidate.startswith("Time In "):
                time_in_line = candidate
            elif candidate.startswith("Time Out "):
                time_out_line = candidate
            elif candidate.startswith("Total "):
                total_line = candidate
                break

        if not day_line or not time_in_line or not time_out_line:
            continue

        dates = _fieldglass_dates_from_day_line(day_line, period_start, period_end)
        start_values = _fieldglass_time_values(time_in_line, "Time In", dates)
        end_values = _fieldglass_time_values(time_out_line, "Time Out", dates)
        total_values = _fieldglass_duration_values(total_line, dates) if total_line else [""] * len(dates)

        for work_date, start_value, end_value, total_value in zip(dates, start_values, end_values, total_values):
            if work_date is None:
                continue

            start = _parse_time_str(start_value)
            end = _parse_time_str(end_value)
            if start is None and end is None:
                continue
            if _is_zero_fieldglass_duration(total_value):
                continue

            rows.append({
                "氏名": name,
                "日付": work_date,
                "出勤時刻": start,
                "退勤時刻": end,
                "コメント": None,
                "データソース": "勤務表",
            })

    if not rows:
        return None

    return pd.DataFrame(rows, columns=["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"])


def _parse_known_timesheet_file(filepath):
    """既知フォーマットをAI解析せずに直接読む"""
    for parser in (
        _parse_itone_dispatch_timesheet_file,
        _parse_sap_timesheet_file,
        _parse_estaffing_timesheet_csv,
        _parse_estaffing_timesheet_text,
        _parse_fieldglass_pdf_timesheet,
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

    if ext in (".xlsx", ".xls", ".xlsb"):
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

    if ext in (".xlsx", ".xls", ".xlsb"):
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
