"""シフト記号式の勤務表を Claude API で解析するモジュール

役割:
1. 勤務表のモード判定（時刻直書き / 記号式 / 混在）
2. 記号式の場合: 凡例＋シフト表 を構造化抽出
3. 結果を shift_resolver に渡せる形式で返す
"""

from __future__ import annotations

import base64
import json
import logging
import re
import time as _time
from typing import Any

import anthropic

from config import Config

logger = logging.getLogger(__name__)


# モード判定 + 抽出を 1 回の Claude 呼び出しで行うシステムプロンプト
LEGEND_SYSTEM_PROMPT = """あなたは勤務シフト表を解析する専門家です。
与えられた勤務表は以下のいずれかのモードで書かれています:

- "direct" : 出退勤時刻が直接書かれている (例: 9:00-17:30)
- "code"   : 記号やコードで書かれていて、凡例（記号→時刻の対応表）が同じファイルに含まれている
             (例: B = 12:30-21:00, ● = 16:30-翌1:00)

このうち "code" モードの場合、以下の JSON を返してください。

出力形式:
{
  "mode": "code",
  "year": 2026,                       // シフト表の対象年（読み取れる場合）。不明なら null
  "month": 4,                         // シフト表の対象月（読み取れる場合）。不明なら null
  "legend": [
    {
      "code": "B",                    // 表中で使われている記号そのまま
      "label": "B勤",                 // 凡例に書かれている種別名
      "start_time": "12:30",          // HH:MM 形式
      "end_time": "21:00",            // HH:MM 形式 (深夜跨ぎは "25:00" "33:00" のように 24h 超で表現)
      "break_minutes": 60,            // 休憩時間（分）。記載が無ければ 0
      "is_off": false                 // 休日扱いの記号なら true
    },
    {
      "code": "明",
      "label": "明け休",
      "is_off": true                  // 休扱いなので時刻は不要
    }
  ],
  "off_markers": ["", "—", "休", "×"],  // 表中で「休み」を意味する記号や空欄パターン
  "employees": [
    {
      "name": "小嶋桃子",              // 漢字氏名（不明なら "不明"）
      "shifts": [
        {"date": "2026-04-01", "code": "B"},
        {"date": "2026-04-02", "code": "B"},
        {"date": "2026-04-03", "code": ""},   // 空欄もそのまま記録
        ...
      ]
    }
  ]
}

"direct" モードの場合は以下の形式で返してください（既存の単純な勤務表向け）:
{
  "mode": "direct",
  "data": [
    {
      "employee_name": "氏名",
      "records": [
        {"date": "YYYY-MM-DD", "start_time": "HH:MM", "end_time": "HH:MM", "comment": null}
      ]
    }
  ]
}

ルール:
- 日付は必ず YYYY-MM-DD 形式
- 時刻は必ず HH:MM (24時間制)。深夜跨ぎは "25:00" "33:00" のように 24h+ で表記
- 凡例に記載のある記号はすべて legend に列挙する
- "明" "明け" のような「明け休」を意味する記号は is_off=true で必ず登録する
- 表中の空欄は "code": "" として記録する（休扱いの判定は off_markers と is_off で行う）
- mode の判定:
  - 凡例（記号→時刻の対応表）が表内・表下・別シートにある  → "code"
  - 各セルに "9:00-17:30" のような時刻文字列が直接ある       → "direct"
  - 両方ある場合は "code" 優先（凡例があれば必ず "code" 扱い）
- JSON のみを返し、それ以外のテキスト・コードフェンス・説明は一切含めない
"""


def _build_messages(file_content, file_type: str, media_type: str | None = None):
    """Claude API に投げる messages を構築する"""
    if file_type == "text":
        return [{
            "role": "user",
            "content": f"以下の勤務表データを解析してください:\n\n{file_content}",
        }]

    if file_type == "image":
        b64 = base64.standard_b64encode(file_content).decode("utf-8")
        return [{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": media_type or "image/png",
                        "data": b64,
                    },
                },
                {"type": "text", "text": "この勤務表を解析してください。"},
            ],
        }]

    if file_type == "pdf":
        b64 = base64.standard_b64encode(file_content).decode("utf-8")
        return [{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": b64,
                    },
                },
                {"type": "text", "text": "この勤務表を解析してください。"},
            ],
        }]

    raise ValueError(f"未対応のfile_type: {file_type}")


def _extract_json(text: str) -> dict | None:
    """Claude の応答から JSON を取り出す（コードフェンス対応）"""
    text = text.strip()
    fence = re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", text)
    if fence:
        text = fence.group(1).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        logger.warning("JSON パース失敗: %s", e)
        return None


def parse_with_legend_extraction(
    file_content,
    file_type: str,
    media_type: str | None = None,
    max_retries: int = 3,
) -> dict[str, Any]:
    """凡例＋シフト表を抽出する（モード判定込み）

    Args:
        file_content: text の場合は str、image/pdf の場合は bytes
        file_type: "text" | "image" | "pdf"
        media_type: image の場合の MIME タイプ（"image/png" 等）

    Returns:
        {"mode": "code"|"direct", ...}  失敗時は RuntimeError
    """
    client = anthropic.Anthropic()
    messages = _build_messages(file_content, file_type, media_type)

    last_error = None
    for attempt in range(max_retries):
        try:
            response = client.messages.create(
                model=Config.ANTHROPIC_MODEL,
                max_tokens=Config.ANTHROPIC_MAX_TOKENS,
                system=LEGEND_SYSTEM_PROMPT,
                messages=messages,
            )
            result_text = response.content[0].text
            data = _extract_json(result_text)
            if data is None:
                logger.warning("試行 %d: JSON 抽出失敗", attempt + 1)
                last_error = "JSON parse failed"
                # プロンプトを強化して再送
                if isinstance(messages[-1]["content"], list):
                    messages[-1]["content"][-1]["text"] = (
                        "この勤務表を解析してください。必ず JSON のみを返し、"
                        "前後に説明文やコードフェンスを付けないでください。"
                    )
                _time.sleep(2)
                continue

            mode = data.get("mode")
            if mode not in ("code", "direct"):
                logger.warning("不明な mode: %r", mode)
                last_error = f"unknown mode: {mode}"
                _time.sleep(2)
                continue

            return data

        except anthropic.RateLimitError:
            logger.warning("レート制限 (試行 %d/%d)", attempt + 1, max_retries)
            _time.sleep(3)
            last_error = "rate limit"
        except anthropic.APIError as e:
            raise RuntimeError(f"Claude API エラー: {e}") from e

    raise RuntimeError(f"勤務表の解析に失敗しました: {last_error}")


def to_legend_dict_for_ui(data: dict) -> dict:
    """code モードの結果を UI 表示用に整形する

    Returns:
        {
          "year": int|None, "month": int|None,
          "legend": [...],
          "off_markers": [...],
          "employees": [...]   ← UI には見せないが server-side セッションに保持
        }
    """
    if data.get("mode") != "code":
        return {}
    return {
        "year": data.get("year"),
        "month": data.get("month"),
        "legend": data.get("legend") or [],
        "off_markers": data.get("off_markers") or [],
        "employees": data.get("employees") or [],
    }
