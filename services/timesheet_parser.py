import anthropic
import base64
import json
import time
import logging
import re
from datetime import datetime, date, time as dt_time
import pandas as pd
from config import Config

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


def parse_timesheet(filepath, progress_callback=None):
    """
    勤務表ファイル（Excel/PDF/画像）を解析して統一DataFrameを返す

    progress_callback: 進捗通知用コールバック (任意)
    """
    import os
    ext = os.path.splitext(filepath)[1].lower()

    if ext in (".xlsx", ".xls"):
        if progress_callback:
            progress_callback("Excelファイルをテキスト変換中...")
        text = _excel_to_text(filepath)
        result = _parse_with_claude(text, "text")

    elif ext == ".pdf":
        if progress_callback:
            progress_callback("PDFを解析中...")
        text, pdf_bytes = _pdf_to_text_or_bytes(filepath)
        if text:
            result = _parse_with_claude(text, "text")
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
