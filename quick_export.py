"""quick_export.py — 5月本番 MVP jinjer アップロード用 全件CSV生成

入力:
  --diff       : quick_compare.py が出力した 差異一覧xlsx（人間判断埋め込み済み）
  --jinjer-dir : 元の jinjer ダウンロードCSV 群のフォルダ（CP932, 194列）
  --output     : 出力する jinjer アップロード用CSV のパス（CP932）

処理:
  1. 差異一覧の「人間判断=承認」行を集める
  2. 元 jinjer CSV を全件読み（行順保持）
  3. 承認行に応じて該当セルを上書き:
     - 出勤差異 → 出勤1（22列目）
     - 退勤差異 → 退勤1（23列目）
     - 手入力休憩1 / 手入力復帰1 / 手入力休憩時間 → 差異種別（出勤/退勤/休憩/総労働時間）を
       問わず、入力があれば (従業員, 日付) 単位で休憩1/復帰1/休憩時間へ反映（保留の行は除く）
     - 手入力セルがExcelの経過時間(timedelta: 31:00等の24時超入力)や時刻(time)で
       保存されていても、jinjerが受ける H:MM 表記（24時超はそのまま31:00形式）へ正規化
     - 総労働時間差異 → 手入力休憩が無い場合は警告のみ、上書きしない
  4. このツールが値を変更した (従業員, 日付) の行だけをアップロード用CSVとして書き出す
     - 無変更行を送り返すと、jinjer側の既存不整合（休暇残数不足・勤務状況フラグの矛盾
       など、jinjer自身がエクスポートした行でも再インポートで弾かれるもの）でエラー
       通知が汚れるだけのため出力しない
     - 人間判断=保留 の日は変更されないので自然に出力から外れる（保留=「この日は
       ツールで触らない」。DL後にjinjer画面で手修正した日を保留にする運用のため、
       古い値の再インポートによる手修正の巻き戻しを防ぐ）
     - 打刻を書く日は 勤務状況(未打刻/欠勤)フラグをクリアし、出勤予定時刻の同期は
       退勤予定時刻がある行に限る。打刻を両方削除した日は休憩の残りもクリアする
       （いずれもjinjerインポートの整合性チェック対策）

実績確定済みの扱い:
  実績確定済（180列=TRUE）は「本人が打刻申請を確定した状態」を示すだけで、
  勤怠が正しいことを保証しない。原則として請求勤怠の値が正であり、管理部の
  「人間判断=承認」があれば実績確定済でも上書きする。
  実績確定済の件数は参考表示としてサマリに出すが、上書きの可否判断には使わない。

dry-run（既定）: 承認件数・上書き予定件数を表示するだけ
--execute      : CSV を書き出す
"""

from __future__ import annotations

import argparse
import csv
import io
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, date, time, timedelta
from pathlib import Path
from typing import Any

import pandas as pd


# ===== jinjer CSV の上書き対象列（ヘッダー名で照合） =====
JINJER_COL_PUNCH_IN = "出勤1"
JINJER_COL_PUNCH_OUT = "退勤1"
JINJER_COL_BREAK_1_START = "休憩1"
JINJER_COL_BREAK_1_END = "復帰1"
JINJER_COL_BREAK_TOTAL = "休憩時間"
JINJER_COL_SCHED_IN = "出勤予定時刻"  # スケジュール（予定）開始時刻。出勤採用時に実打刻へ合わせる
JINJER_COL_SCHED_OUT = "退勤予定時刻"  # スケジュール未設定の行に出勤予定だけ書かないための判定用
JINJER_COL_EMP_ID = "*従業員ID"
JINJER_COL_DATE = "*年月日"
# 勤務状況（0:未打刻 1:欠勤）。打刻を書き込む日にこのフラグが残っていると
# jinjer インポートが「出勤1/退勤1/勤務状況が正しくありません」で弾くためクリアする。
JINJER_COL_WORK_STATUS_PREFIX = "勤務状況"

# 差異種別
DIFF_KIND_PUNCH_IN = "出勤"
DIFF_KIND_PUNCH_OUT = "退勤"
DIFF_KIND_BREAK = "休憩"
DIFF_KIND_TOTAL = "総労働時間"
# jinjer打刻なしの行は、汎用データの勤務状況により差異種別が「欠勤」「未打刻」と
# 表示される（2026-07-10）。書き戻し先（出勤1/退勤1）は quick_compare が生成する
# 警告理由の「jinjer出勤なし…／jinjer退勤なし…」から解決する。
# 警告理由で解決できない「欠勤」は日単位の転記行＝書き戻し対象外。
DIFF_KIND_ABSENCE = "欠勤"
DIFF_KIND_NO_PUNCH = "未打刻"
DIFF_KIND_SCHED_START = "スケジュール開始"  # 出勤予定時刻のみ書き戻す（quick_compare と一致）
# スケジュール開始合わせの参考シート名（quick_compare.write_excel と一致）。
# 差異一覧の行としては出さず（確認対象を増やさない・2026-08-13 谷津さん指定）、
# このシートを読んで出勤予定時刻を黙って合わせる。
SCHED_ALIGN_SHEET = "スケジュール開始合わせ"

# 人間判断の値。新ラベル: 「請求勤怠」=請求勤怠を正としてjinjerへ書き戻す（旧「承認」）/
# 「jinjer勤怠」=jinjerを正（書き戻さない・旧「却下」）/「保留」。
# すでに「承認/却下」で入力済みの差異一覧も読めるよう、旧ラベルも後方互換で受理する。
JUDGE_APPROVE = "請求勤怠"
JUDGE_REJECT = "jinjer勤怠"
JUDGE_HOLD = "保留"
APPROVE_LABELS = {JUDGE_APPROVE, "承認"}      # 請求勤怠を正 → 書き戻す
REJECT_LABELS = {JUDGE_REJECT, "却下"}        # jinjer勤怠を正 → 書き戻さない
HOLD_LABELS = {JUDGE_HOLD}
JUDGMENTS = APPROVE_LABELS | REJECT_LABELS | HOLD_LABELS

# 「人間判断」を入力すべきなのに、誤って判断（請求勤怠/jinjer勤怠/保留）が書き込まれやすい列。
# 本来これらの列には時刻・数値・メモが入るため、判断キーワードが入っていたら
# 「人間判断」列の入力ミスとみなして回収する（先頭ほど優先）。
# ※「自動修正提案値」はシステムが採用ラベル（請求勤怠/jinjer勤怠）を入れる列に変わったため、
#   ここから除外する（システムの提案を人間判断として誤って拾わないようにする）。
MISPLACED_JUDGE_COLS = [
    "打刻修正", "手入力修正値", "手入力休憩1", "手入力復帰1", "手入力休憩時間",
    "判断メモ",
]


# ----------------------------------------------------------------------
# データクラス
# ----------------------------------------------------------------------

@dataclass
class ApprovedRow:
    emp_id: str
    target_date_iso: str  # YYYY-MM-DD
    kind: str
    auto_fix_value: str
    manual_fix_value: str
    manual_break_start: str
    manual_break_end: str
    manual_break_total: str
    name: str  # 表示用
    warn_level: str
    source_diff_row_id: int


@dataclass
class ManualBreak:
    """手入力休憩の反映指示。(従業員ID, 日付) 単位で1件。

    休憩差異は「休憩」行として出ないことも多く（総労働時間差異にしか現れない等）、
    実務では出勤/退勤/総労働時間の行に手入力休憩が書かれる。そのため差異種別に
    関係なく、手入力休憩欄に値があれば反映指示として扱う（保留の行だけ除外）。
    """
    emp_id: str
    date_iso: str
    start: str = ""
    end: str = ""
    total: str = ""
    name: str = ""
    source_diff_row_id: int = 0


@dataclass
class Stats:
    total_diff_rows: int = 0
    approved: int = 0
    rejected: int = 0
    held: int = 0
    pending: int = 0
    # 差異種別ごとの承認内訳
    approved_punch_in: int = 0
    approved_punch_out: int = 0
    approved_break: int = 0
    approved_total: int = 0
    approved_sched_start: int = 0  # スケジュール開始合わせの採用行
    # 実際の上書き処理結果
    overwritten_punch_in: int = 0
    overwritten_punch_out: int = 0
    overwritten_sched_in: int = 0  # 出勤採用に合わせてスケジュール開始（出勤予定時刻）も更新した件数
    overwritten_sched_start: int = 0  # 「スケジュール開始」行の反映＝出勤予定時刻のみ更新（打刻は触らない）
    skipped_break: int = 0  # 承認されたが上書きしなかった
    skipped_total: int = 0
    overwritten_break_start: int = 0
    overwritten_break_end: int = 0
    overwritten_break_total: int = 0
    manual_break_days: int = 0  # 手入力休憩の反映指示 (従業員, 日付) 件数
    manual_break_conflicts: int = 0  # 同一 (従業員, 日付) で手入力休憩の値が食い違った件数
    held_rows_removed: int = 0  # 保留日のうち出力から外れた日数（jinjer手修正の保護）
    unchanged_rows_removed: int = 0  # 無変更のため出力しなかった行数
    manual_clear_days: int = 0  # 打刻削除を含むため出力から除外した日数（jinjer画面で手動対応）
    not_matched: int = 0  # jinjer 行が見つからない承認行
    overwritten_finalized: int = 0  # 上書きしたうち実績確定済みだった件数（参考表示）
    recovered_misplaced: int = 0  # 「人間判断」列以外に入っていた判断を回収した件数
    misplaced_by_col: dict[str, int] = field(default_factory=dict)  # 列名→誤入力件数
    warnings: list[str] = field(default_factory=list)


# ----------------------------------------------------------------------
# ユーティリティ
# ----------------------------------------------------------------------

def normalize_date_iso(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    s = str(value).strip()
    if not s or s.lower() in ("nan", "nat"):
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date().isoformat()
    except ValueError:
        pass
    try:
        return datetime.strptime(s, "%Y/%m/%d").date().isoformat()
    except ValueError:
        pass
    return None


def iso_to_jinjer_date(date_iso: str) -> str:
    """'2026-05-12' → '2026/5/12'（jinjer の 0パディングなし形式）"""
    d = datetime.strptime(date_iso, "%Y-%m-%d").date()
    return f"{d.year}/{d.month}/{d.day}"


_HHMM_RE = re.compile(r"^(\d{1,3}):(\d{2})")
# 秒付き時刻（10:00:00 / 33:30:00 等）。ExcelでCSVを開いて上書き保存すると、
# 24時超の時刻が [h]:mm:ss 形式になり秒が付く。jinjerは秒付きを受け付けないため
# H:MM に正規化する。日時文字列（"2026-06-01 10:00:00"）はセル全体が一致しないので対象外。
_TIME_WITH_SECONDS_RE = re.compile(r"^(\d{1,3}:\d{2}):\d{2}$")


def strip_time_seconds(value: str) -> str:
    """'10:00:00'/'33:30:00' → '10:00'/'33:30'。時刻でない値はそのまま返す。"""
    if not value:
        return value
    m = _TIME_WITH_SECONDS_RE.match(value.strip())
    return m.group(1) if m else value


def timedelta_to_hhmm(td: timedelta) -> str:
    """Excelの経過時間セルを jinjer の H:MM 表記（24時超そのまま）へ変換する。

    差異一覧xlsxの手入力セルに 31:00 のような24時超時刻を入力すると、Excelが
    [h]:mm:ss の経過時間として保存し、openpyxl が timedelta(days=1, seconds=25200)
    で返す。str() の '1 day, 7:00:00' はjinjerが受け付けないため、総分に換算して
    '31:00' 形式にする（加藤さん 2026-06 夜勤の復帰1で実際にインポートが弾かれた）。
    """
    total_sec = int(td.total_seconds())
    if total_sec < 0:  # Excelの手入力セルでは負の経過時間は作れない（念のため素通し）
        return str(td)
    total_min = (total_sec + 30) // 60  # 30秒以上は分へ繰り上げ
    return f"{total_min // 60}:{total_min % 60:02d}"


# 日付の正規化。jinjerのインポートはスラッシュ形式（2026/6/1・0パディングなし）しか
# 受け付けず、それ以外は「年月日が正しくありません」で全行弾かれる（2026-07-08実測）。
# 実際に遭遇した壊れ方:
# - ISO形式(2026-06-01): jinjer自身のエクスポートがこの形式で出す
# - 米国式(6/1/2026): ExcelでCSVを開いて上書き保存すると変換されることがある
_ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})$")
_US_DATE_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")


def iso_date_to_jinjer(value: str) -> str:
    """日付セルを jinjerインポート形式 'YYYY/M/D' へ正規化する。日付でない値はそのまま。

    '2026-06-01' → '2026/6/1'（ISO・jinjerエクスポート形式）
    '6/1/2026'   → '2026/6/1'（米国式・Excel保存で化けた形式。月>12なら日/月を入替）
    """
    if not value:
        return value
    s = value.strip()
    m = _ISO_DATE_RE.match(s)
    if m:
        return f"{int(m.group(1))}/{int(m.group(2))}/{int(m.group(3))}"
    m = _US_DATE_RE.match(s)
    if m:
        mm, dd, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if mm > 12 and dd <= 12:  # 実はD/M/YYYYだった場合の救済
            mm, dd = dd, mm
        return f"{yyyy}/{mm}/{dd}"
    return value


def _hhmm_to_minutes(value: Any) -> int | None:
    """'H:MM' / 'HH:MM'（24時超表記含む）を分に変換。不正値は None。"""
    if value is None:
        return None
    m = _HHMM_RE.match(str(value).strip())
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def to_overnight_punch_out(punch_in: Any, punch_out: str) -> str:
    """退勤が翌朝に跨ぐ場合、jinjer インポート用の 24時超表記へ変換する。

    出勤 > 退勤（夜勤で翌日にずれ込み）のときだけ退勤に 24h を足す（08:15→32:15）。
    すでに 24時超表記なら out >= in となり再変換しない（冪等）。出退勤の判定が
    できないときは元の退勤値をそのまま返す。
    """
    if not punch_out:
        return punch_out
    in_min = _hhmm_to_minutes(punch_in)
    out_min = _hhmm_to_minutes(punch_out)
    if in_min is None or out_min is None or out_min >= in_min:
        return punch_out
    adjusted = out_min + 24 * 60
    return f"{adjusted // 60:02d}:{adjusted % 60:02d}"


def _warn_if_not_hhmm(
    stats: Stats, row_id: int, col_label: str, value: str,
    emp_id: str, date_iso: str, name: str,
) -> None:
    """jinjerの時刻列へ書き込む値が H:MM 形式でなければ警告する（書き込みは行う）。

    実例: 手入力休憩時間に数値 1 を入力（'1時間' のつもり）→ '1' のまま出力される。
    数値の時刻への自動解釈は、'30'（30分のつもり）を30時間と取り違える危険があるため
    行わず、差異一覧xlsx側の入力を直してもらう。
    """
    if value and not _HHMM_RE.fullmatch(value):
        stats.warnings.append(
            f"⚠️行ID={row_id} {col_label}='{value}' は時刻形式(H:MM)ではないため、"
            f"jinjerインポートで弾かれる可能性があります。差異一覧の該当セルを "
            f"1:00 や 31:00 のような時刻で入力し直して再出力してください "
            f"(emp={emp_id} date={date_iso} {name})"
        )


def clean_excel_text(value: Any) -> str:
    """Excel由来の値を、jinjer汎用データCSVへ書く文字列に正規化する。

    時刻はjinjerエクスポートと同じ 0埋めなしの H:MM に揃える。Excelは同じ列でも
    入力のされ方でセル型が変わる（31:00→経過時間timedelta / 7:00→time /
    先頭'付き→文字列）ため、どの型で保存されていても同じ表記に落とす。
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, timedelta):
        # 24時超の時刻入力（31:00等）はExcelが経過時間として保存しtimedeltaで返る
        return timedelta_to_hhmm(value)
    if isinstance(value, datetime):
        return f"{value.hour}:{value.minute:02d}"
    if isinstance(value, time):
        return f"{value.hour}:{value.minute:02d}"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return ""
    # Excel編集で秒が付いた時刻（10:00:00等）はjinjerが弾くため H:MM へ正規化
    return strip_time_seconds(s)


# ----------------------------------------------------------------------
# 入力読み込み
# ----------------------------------------------------------------------

def load_approved_rows(
    diff_xlsx: Path, stats: Stats
) -> tuple[list[ApprovedRow], dict[tuple[str, str], ManualBreak], set[tuple[str, str]]]:
    """差異一覧xlsx の「差異一覧」シートから 人間判断=承認 の行と手入力休憩を抽出。

    Returns:
        (approved, manual_breaks, held_days)
        - approved      : 人間判断=請求勤怠(承認) の行
        - manual_breaks : 手入力休憩1/復帰1/休憩時間 が入力された行から集めた
                          (従業員ID, 日付ISO) → ManualBreak。差異種別・人間判断の
                          種類を問わず収集する（保留の行だけ除外）。
        - held_days     : 人間判断=保留 が1つでもある (従業員ID, 日付ISO) の集合。
                          アップロードCSVから行ごと除外する対象。
    """
    if not diff_xlsx.exists():
        raise FileNotFoundError(f"差異一覧xlsx が見つかりません: {diff_xlsx}")

    xl = pd.ExcelFile(diff_xlsx)
    if "差異一覧" not in xl.sheet_names:
        if "突合結果" in xl.sheet_names:
            raise ValueError(
                "選択されたファイルは手順1の勤怠突合結果xlsxです。"
                "手順3には、手順2で生成して人間判断を入力した差異一覧xlsxを指定してください。"
            )
        sheets = "、".join(xl.sheet_names)
        raise ValueError(
            "差異一覧シートが見つかりません。"
            "手順3には、手順2で生成して人間判断を入力した差異一覧xlsxを指定してください。"
            f" このファイルのシート: {sheets}"
        )

    df = pd.read_excel(xl, sheet_name="差異一覧", dtype=object)
    stats.total_diff_rows = len(df)

    approved: list[ApprovedRow] = []
    manual_breaks: dict[tuple[str, str], ManualBreak] = {}
    held_days: set[tuple[str, str]] = set()
    for _, row in df.iterrows():
        judge = str(row.get("人間判断") or "").strip()

        # 判断が「人間判断」列に無い場合、別列への入力ミスを回収する。
        # 承認/却下/保留 は時刻・メモには本来現れない閉じた語彙なので、
        # 他の列に厳密一致で入っていれば判断の置き場所間違いとみなして拾う。
        misplaced_col: str | None = None
        if judge not in JUDGMENTS:
            for col in MISPLACED_JUDGE_COLS:
                v = str(row.get(col) or "").strip()
                if v in JUDGMENTS:
                    judge = v
                    misplaced_col = col
                    stats.recovered_misplaced += 1
                    stats.misplaced_by_col[col] = stats.misplaced_by_col.get(col, 0) + 1
                    break

        # 判断を回収した列の値は「承認」等の文字列なので、時刻として書き込まないよう除外する。
        def _field(col: str) -> str:
            if col == misplaced_col:
                return ""
            v = clean_excel_text(row.get(col))
            # 念のため、どの入力欄にも判断キーワードが時刻として紛れ込まないよう除外
            return "" if v in JUDGMENTS else v

        emp_id = str(row.get("従業員ID") or "").strip()
        date_iso = normalize_date_iso(row.get("対象日付"))
        kind = str(row.get("差異種別") or "").strip()
        # 差異種別が欠勤/未打刻の打刻行は、警告理由から出勤/退勤どちらの書き戻しかを解決する。
        # どちらの語も無ければ日単位の欠勤転記行＝書き戻し対象外（種別のまま後段で無視される）。
        if kind in (DIFF_KIND_ABSENCE, DIFF_KIND_NO_PUNCH):
            reason = str(row.get("警告理由") or "")
            if "jinjer出勤なし" in reason:
                kind = DIFF_KIND_PUNCH_IN
            elif "jinjer退勤なし" in reason:
                kind = DIFF_KIND_PUNCH_OUT
        name = str(row.get("氏名") or "").strip()
        row_id_raw = row.get("行ID")
        try:
            row_id = int(row_id_raw) if row_id_raw is not None else 0
        except (ValueError, TypeError):
            row_id = 0

        # 手入力休憩は、差異種別（出勤/退勤/総労働時間/休憩）や人間判断のラベルに
        # 関係なく「入力があれば反映する」指示として全行から収集する。
        # 休憩差異は「休憩」行として出ず総労働時間差異にしか現れないことが多く、
        # 実務では出勤/退勤/総労働時間の行に手入力されるため（菅原さん 2026-06 で
        # 全行無視され休憩が反映されなかった実例あり）。保留の行だけは対象外。
        mb_start = _field("手入力休憩1")
        mb_end = _field("手入力復帰1")
        mb_total = _field("手入力休憩時間")
        if (mb_start or mb_end or mb_total) and emp_id and date_iso:
            if judge in HOLD_LABELS:
                stats.warnings.append(
                    f"行ID={row_id_raw} 手入力休憩がありますが人間判断=保留のため反映しません "
                    f"(emp={emp_id} date={date_iso} {name})"
                )
            else:
                key = (emp_id, date_iso)
                existing = manual_breaks.get(key)
                if existing is None:
                    manual_breaks[key] = ManualBreak(
                        emp_id=emp_id, date_iso=date_iso,
                        start=mb_start, end=mb_end, total=mb_total,
                        name=name, source_diff_row_id=row_id,
                    )
                else:
                    conflict = (
                        (mb_start and existing.start and mb_start != existing.start)
                        or (mb_end and existing.end and mb_end != existing.end)
                        or (mb_total and existing.total and mb_total != existing.total)
                    )
                    if conflict:
                        stats.manual_break_conflicts += 1
                        stats.warnings.append(
                            f"行ID={row_id_raw} 同じ日付の手入力休憩が食い違っています。"
                            f"先に読んだ行ID={existing.source_diff_row_id} の値を採用します "
                            f"(emp={emp_id} date={date_iso} {name}: "
                            f"採用={existing.start}/{existing.end}/{existing.total} "
                            f"無視={mb_start}/{mb_end}/{mb_total})"
                        )
                    else:
                        # 同値の重複入力（複数行に同じ休憩を記入）は空欄だけ補完する
                        existing.start = existing.start or mb_start
                        existing.end = existing.end or mb_end
                        existing.total = existing.total or mb_total

        if judge in APPROVE_LABELS:
            stats.approved += 1
        elif judge in REJECT_LABELS:
            stats.rejected += 1
            continue
        elif judge in HOLD_LABELS:
            stats.held += 1
            # 保留の日は行ごとアップロードCSVから除外する（jinjer手修正の保護）
            if emp_id and date_iso:
                held_days.add((emp_id, date_iso))
            continue
        else:
            stats.pending += 1
            continue
        # 「請求勤怠」を採用したときに書き戻す時刻。
        # 旧フォーマットは「自動修正提案値」列に時刻が入っていた。新フォーマットでは同列が
        # 採用ラベル（請求勤怠/jinjer勤怠/保留）に変わり _field で空に剥がれるため、表示している
        # 「請求勤怠値」（＝請求勤怠の打刻時刻）にフォールバックする。
        # ※これが無いと、退勤等を請求勤怠で承認したのに空で上書き＝打刻が消える不具合になる
        #   （夜勤の太田さん 5/7 等で実際に退勤1が空になった）。
        auto_fix = _field("自動修正提案値") or _field("請求勤怠値")
        # 列名は「打刻修正」（新）。旧フォーマットの「手入力修正値」も後方互換で読む。
        manual_fix = _field("打刻修正") or _field("手入力修正値")
        warn_level = str(row.get("警告レベル") or "").strip()

        if not emp_id or not date_iso or not kind:
            stats.warnings.append(
                f"行ID={row_id_raw} 承認だが必須フィールド欠落 (emp_id={emp_id}, date={date_iso}, kind={kind})"
            )
            continue

        approved.append(ApprovedRow(
            emp_id=emp_id, target_date_iso=date_iso, kind=kind,
            auto_fix_value=auto_fix,
            manual_fix_value=manual_fix,
            manual_break_start=mb_start,
            manual_break_end=mb_end,
            manual_break_total=mb_total,
            name=name, warn_level=warn_level,
            source_diff_row_id=row_id,
        ))

        # 種別別集計
        if kind == DIFF_KIND_SCHED_START:
            stats.approved_sched_start += 1
        elif kind == DIFF_KIND_PUNCH_IN:
            stats.approved_punch_in += 1
        elif kind == DIFF_KIND_PUNCH_OUT:
            stats.approved_punch_out += 1
        elif kind == DIFF_KIND_BREAK:
            stats.approved_break += 1
        elif kind == DIFF_KIND_TOTAL:
            stats.approved_total += 1

    # 判断が「人間判断」列以外に入っていたものを回収した場合、目立つ警告を残す。
    if stats.recovered_misplaced:
        detail = "、".join(
            f"『{col}』列 {cnt}件" for col, cnt in stats.misplaced_by_col.items()
        )
        stats.warnings.insert(
            0,
            f"⚠️ 判断（請求勤怠/jinjer勤怠/保留）が「人間判断」列ではなく {detail} に入力されていました。"
            f"合計 {stats.recovered_misplaced} 件を自動で回収して処理しました。"
            "次回は必ず「人間判断」列（プルダウンのある列）に入力してください。",
        )

    stats.manual_break_days = len(manual_breaks)
    return approved, manual_breaks, held_days


def _csv_input_files(path: Path) -> list[Path]:
    if path.is_file():
        return [path]
    return sorted([p for p in path.glob("*.csv") if not p.name.startswith("~$")])


def load_jinjer_csvs(jinjer_dir: Path) -> tuple[list[str], list[list[str]]]:
    """jinjer ダウンロードCSV を全件 CP932 で読む。
    複数ファイルがある場合は縦結合（ヘッダーは先頭ファイルのものを採用、列数一致を要求）。
    CSVファイル単体の指定にも対応する。

    Returns:
        (headers, rows) — rows は 1行 = list[str]
    """
    if not jinjer_dir.exists():
        raise FileNotFoundError(f"jinjer CSVまたはフォルダが見つかりません: {jinjer_dir}")

    csv_paths = _csv_input_files(jinjer_dir)
    if not csv_paths:
        raise RuntimeError(f"jinjer CSV が見つかりません: {jinjer_dir}")

    all_headers: list[str] | None = None
    all_rows: list[list[str]] = []
    for path in csv_paths:
        with open(path, encoding="cp932", newline="") as f:
            reader = csv.reader(f)
            headers = next(reader)
            if all_headers is None:
                all_headers = headers
            else:
                if len(headers) != len(all_headers):
                    raise RuntimeError(
                        f"jinjer CSV の列数が一致しません: {path.name} は {len(headers)} 列、"
                        f"先頭ファイルは {len(all_headers)} 列"
                    )
            for row in reader:
                # 列数を揃える（足りなければ空文字で埋める）
                if len(row) < len(all_headers):
                    row = row + [""] * (len(all_headers) - len(row))
                elif len(row) > len(all_headers):
                    row = row[: len(all_headers)]
                all_rows.append(row)
        print(f"[info] jinjer CSV 読込: {path.name} (累計 {len(all_rows)} 行)")

    assert all_headers is not None

    # jinjerインポートが受け付ける形式へ正規化する:
    # - 秒付き時刻(33:30:00等): ExcelでCSVを上書き保存すると24時超の時刻に秒が付く
    #   （Excelが経過時間 [h]:mm:ss として保存するため）。jinjerは秒付きを弾く。
    # - ISO日付(2026-06-01): jinjerのエクスポートはISO形式なのに、インポートは
    #   スラッシュ形式(2026/6/1)しか受け付けず「年月日が正しくありません」で弾く。
    time_fixed = 0
    date_fixed = 0
    for row in all_rows:
        for c, v in enumerate(row):
            if not v:
                continue
            if ":" in v:
                nv = strip_time_seconds(v)
                if nv != v:
                    row[c] = nv
                    time_fixed += 1
            elif "-" in v or "/" in v:
                nv = iso_date_to_jinjer(v)
                if nv != v:
                    row[c] = nv
                    date_fixed += 1
    if time_fixed:
        print(f"[info] 秒付き時刻 {time_fixed} セルを H:MM 形式へ正規化（Excel保存されたCSV対策）")
    if date_fixed:
        print(f"[info] ISO日付 {date_fixed} セルを jinjerインポート形式(YYYY/M/D)へ正規化")

    return all_headers, all_rows


def build_jinjer_row_index(
    headers: list[str], rows: list[list[str]]
) -> dict[tuple[str, str], int]:
    """(従業員ID, 日付ISO) → rows のインデックス"""
    try:
        emp_col_idx = headers.index(JINJER_COL_EMP_ID)
        date_col_idx = headers.index(JINJER_COL_DATE)
    except ValueError as e:
        raise RuntimeError(f"jinjer CSV に必須列が見つかりません: {e}")

    index: dict[tuple[str, str], int] = {}
    duplicates: list[tuple[str, str]] = []
    for i, row in enumerate(rows):
        emp_id = (row[emp_col_idx] or "").strip()
        date_iso = normalize_date_iso(row[date_col_idx])
        if not emp_id or not date_iso:
            continue
        key = (emp_id, date_iso)
        if key in index:
            duplicates.append(key)
        index[key] = i
    if duplicates:
        print(f"[warn] (従業員ID, 年月日) の重複が {len(duplicates)} 件ありました（最後の行を採用）")
    return index


# ----------------------------------------------------------------------
# CSV 上書き
# ----------------------------------------------------------------------

def load_sched_aligns(diff_xlsx: Path) -> list[dict]:
    """「スケジュール開始合わせ」シートを読む（無ければ空＝旧フォーマット互換）。

    戻り値: [{"emp", "date_iso", "new_start", "name"}, ...]
    """
    xl = pd.ExcelFile(diff_xlsx)
    if SCHED_ALIGN_SHEET not in xl.sheet_names:
        return []
    df = pd.read_excel(xl, sheet_name=SCHED_ALIGN_SHEET, dtype=object)
    out: list[dict] = []
    for _, row in df.iterrows():
        emp = clean_excel_text(row.get("従業員ID"))
        date_iso = normalize_date_iso(row.get("対象日付"))
        new_start = clean_excel_text(row.get("新しい開始(請求勤怠)"))
        if emp and date_iso and new_start:
            out.append({"emp": emp, "date_iso": date_iso, "new_start": new_start,
                        "name": clean_excel_text(row.get("氏名"))})
    return out


def apply_sched_aligns(
    headers: list[str],
    rows: list[list],
    row_index: dict,
    aligns: list[dict],
    held_days: set[tuple[str, str]],
    stats: "Stats",
    changed_days: set[tuple[str, str]],
) -> None:
    """スケジュール開始合わせを jinjer 行へ反映する（出勤予定時刻のみ・打刻は触らない）。

    方向ガード: 新しい開始が現在の予定より早いときだけ書く（遅刻方向は動かさない、
    という検知時のルールを適用時にも守る＝差異一覧が古い場合の保険）。
    保留の日は書かない（jinjer手修正の保護と同じ扱い）。
    """
    sched_in_col = headers.index(JINJER_COL_SCHED_IN) if JINJER_COL_SCHED_IN in headers else None
    sched_out_col = headers.index(JINJER_COL_SCHED_OUT) if JINJER_COL_SCHED_OUT in headers else None
    if sched_in_col is None or sched_out_col is None:
        if aligns:
            stats.warnings.append(
                f"スケジュール開始合わせ {len(aligns)} 件がありますが、jinjer CSV に "
                f"出勤予定時刻/退勤予定時刻の列が無いため反映できません")
        return
    for a in aligns:
        key = (a["emp"], a["date_iso"])
        if key in held_days:
            continue   # 保留日はツールで触らない（既存の保護と同じ）
        idx = row_index.get(key)
        if idx is None:
            stats.not_matched += 1
            stats.warnings.append(
                f"スケジュール開始合わせ: jinjer CSV に該当行なし "
                f"(emp={a['emp']} date={a['date_iso']} {a['name']})")
            continue
        cur = _hhmm_to_minutes(rows[idx][sched_in_col])
        new = _hhmm_to_minutes(a["new_start"])
        if new is None:
            stats.warnings.append(
                f"スケジュール開始合わせ: 時刻を解釈できません {a['new_start']!r} "
                f"(emp={a['emp']} date={a['date_iso']} {a['name']})")
            continue
        # 退勤予定が空の行に出勤予定だけ書くと jinjer が行ごと弾く
        if not (rows[idx][sched_out_col] or "").strip():
            continue
        if cur is None or new >= cur:
            continue   # 予定が既に同じ/より早い・遅刻方向 → 動かさない
        rows[idx][sched_in_col] = a["new_start"]
        stats.overwritten_sched_start += 1
        changed_days.add(key)


def apply_approved_rows(
    headers: list[str],
    rows: list[list[str]],
    row_index: dict[tuple[str, str], int],
    approved: list[ApprovedRow],
    stats: Stats,
    manual_breaks: dict[tuple[str, str], ManualBreak] | None = None,
) -> set[tuple[str, str]]:
    """承認行に応じて rows を in-place で上書きする。

    人間判断=承認 のみが上書きの条件。実績確定状況は上書き可否の判定材料にしない
    （実績確定済 = 本人が打刻申請を確定しただけで、勤怠が正しいことを保証しない。
    原則として請求勤怠が正しいため、管理部の判断＝承認なら実績確定済でも上書きする）。

    Returns:
        changed_days: このツールが値を書いた (従業員ID, 日付ISO) の集合。
        アップロードCSVにはこの日の行だけを出力する（無変更行を送り返すと、
        休暇残数や勤務状況の既存不整合でjinjerのエラー通知が汚れるだけで得がない）。
    """
    required_cols = [JINJER_COL_PUNCH_IN, JINJER_COL_PUNCH_OUT]
    missing = [col for col in required_cols if col not in headers]
    if missing:
        raise RuntimeError(f"jinjer CSV に上書き対象列がありません: {', '.join(missing)}")
    punch_in_col = headers.index(JINJER_COL_PUNCH_IN)
    punch_out_col = headers.index(JINJER_COL_PUNCH_OUT)
    break_start_col = headers.index(JINJER_COL_BREAK_1_START) if JINJER_COL_BREAK_1_START in headers else None
    break_end_col = headers.index(JINJER_COL_BREAK_1_END) if JINJER_COL_BREAK_1_END in headers else None
    break_total_col = headers.index(JINJER_COL_BREAK_TOTAL) if JINJER_COL_BREAK_TOTAL in headers else None
    sched_in_col = headers.index(JINJER_COL_SCHED_IN) if JINJER_COL_SCHED_IN in headers else None
    sched_out_col = headers.index(JINJER_COL_SCHED_OUT) if JINJER_COL_SCHED_OUT in headers else None
    # 勤務状況（0:未打刻1:欠勤）列。正式ヘッダーは注釈付きなので前方一致で解決する。
    work_status_col = next(
        (i for i, h in enumerate(headers) if str(h).startswith(JINJER_COL_WORK_STATUS_PREFIX)),
        None,
    )

    # 実績確定状況列があれば、上書きしたうち何件が実績確定済だったかを集計する（参考表示のみ）
    finalized_col = headers.index("実績確定状況") if "実績確定状況" in headers else None

    # このツールが値を書いた (従業員ID, 日付ISO)。アップロードCSVには変更行だけを出力する。
    changed_days: set[tuple[str, str]] = set()
    # 打刻を空で上書き（削除）した日。後段で休憩の残りをクリアし、片側残りを警告する。
    cleared_punch_days: set[tuple[str, str]] = set()

    # ---- 手入力休憩の反映（差異種別に依存しない・(従業員, 日付) 単位で1回） ----
    # 休憩は請求勤怠から自動推定せず、人間が差異一覧に入力した欄だけ反映する。
    manual_breaks = manual_breaks or {}
    for key, mb in manual_breaks.items():
        idx = row_index.get(key)
        if idx is None:
            stats.not_matched += 1
            stats.warnings.append(
                f"行ID={mb.source_diff_row_id} 手入力休憩の対象行が jinjer CSV にありません "
                f"(emp={mb.emp_id} date={mb.date_iso} {mb.name})"
            )
            continue
        # 時刻形式でない手入力値（数値の 1 など）は書き込む前に警告しておく
        _warn_if_not_hhmm(stats, mb.source_diff_row_id, "手入力休憩1", mb.start, mb.emp_id, mb.date_iso, mb.name)
        _warn_if_not_hhmm(stats, mb.source_diff_row_id, "手入力復帰1", mb.end, mb.emp_id, mb.date_iso, mb.name)
        _warn_if_not_hhmm(stats, mb.source_diff_row_id, "手入力休憩時間", mb.total, mb.emp_id, mb.date_iso, mb.name)
        if mb.start:
            if break_start_col is None:
                stats.warnings.append(f"行ID={mb.source_diff_row_id} 汎用データに '{JINJER_COL_BREAK_1_START}' 列がありません")
            else:
                rows[idx][break_start_col] = mb.start
                stats.overwritten_break_start += 1
                changed_days.add(key)
        if mb.end:
            if break_end_col is None:
                stats.warnings.append(f"行ID={mb.source_diff_row_id} 汎用データに '{JINJER_COL_BREAK_1_END}' 列がありません")
            else:
                rows[idx][break_end_col] = mb.end
                stats.overwritten_break_end += 1
                changed_days.add(key)
        if mb.total:
            if break_total_col is None:
                stats.warnings.append(f"行ID={mb.source_diff_row_id} 汎用データに '{JINJER_COL_BREAK_TOTAL}' 列がありません")
            else:
                rows[idx][break_total_col] = mb.total
                stats.overwritten_break_total += 1
                changed_days.add(key)

    for app in approved:
        if app.kind == DIFF_KIND_BREAK:
            # 上書き自体は manual_breaks で処理済み。手入力が無い承認だけ警告する。
            if (app.emp_id, app.target_date_iso) not in manual_breaks:
                stats.skipped_break += 1
                stats.warnings.append(
                    f"行ID={app.source_diff_row_id} 休憩差異が承認されましたが、手入力休憩欄が空のため反映しませんでした "
                    f"(emp={app.emp_id} date={app.target_date_iso} {app.name})"
                )
            continue
        if app.kind == DIFF_KIND_TOTAL:
            stats.skipped_total += 1
            # 同じ日に手入力休憩を反映済みなら、それが総労働時間差異への対処なので警告しない
            if (app.emp_id, app.target_date_iso) not in manual_breaks:
                stats.warnings.append(
                    f"行ID={app.source_diff_row_id} 総労働時間差異が承認されましたが、自動反映はスキップしました "
                    f"(emp={app.emp_id} date={app.target_date_iso} {app.name})"
                )
            continue

        if app.kind == DIFF_KIND_SCHED_START:
            # 実績（請求勤怠の出勤）がスケジュール開始より早い日の予定合わせ。
            # 打刻には一切触らず、出勤予定時刻だけを採用値へ書き換える（2026-08-13 谷津さん指定）。
            key = (app.emp_id, app.target_date_iso)
            idx = row_index.get(key)
            if idx is None:
                stats.not_matched += 1
                stats.warnings.append(
                    f"行ID={app.source_diff_row_id} jinjer CSV に該当行なし "
                    f"(emp={app.emp_id} date={app.target_date_iso} {app.name})"
                )
                continue
            new_sched = app.manual_fix_value or app.auto_fix_value
            _warn_if_not_hhmm(stats, app.source_diff_row_id, "出勤予定時刻への書込値", new_sched,
                              app.emp_id, app.target_date_iso, app.name)
            # 退勤予定が空の行に出勤予定だけ書くと jinjer が行ごと弾くため、揃っている行に限る
            # （quick_compare 側でも同条件で行を出すが、二重の防御）
            if (
                new_sched and sched_in_col is not None
                and sched_out_col is not None
                and (rows[idx][sched_out_col] or "").strip()
            ):
                rows[idx][sched_in_col] = new_sched
                stats.overwritten_sched_start += 1
                changed_days.add(key)
            else:
                stats.warnings.append(
                    f"行ID={app.source_diff_row_id} 出勤予定時刻を更新できません"
                    f"（退勤予定が空 か 予定列なし）"
                    f" (emp={app.emp_id} date={app.target_date_iso} {app.name})"
                )
            continue

        if app.kind in (DIFF_KIND_ABSENCE, DIFF_KIND_NO_PUNCH):
            # 警告理由から出勤/退勤を解決できなかった行（日単位の欠勤転記行）は書き戻し対象外
            stats.warnings.append(
                f"行ID={app.source_diff_row_id} 欠勤/未打刻の転記行は書き戻し対象外のため無視 "
                f"(emp={app.emp_id} date={app.target_date_iso} {app.name})"
            )
            continue
        if app.kind not in (DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT):
            stats.warnings.append(
                f"行ID={app.source_diff_row_id} 未知の差異種別 '{app.kind}' を無視 "
                f"(emp={app.emp_id} date={app.target_date_iso})"
            )
            continue

        # jinjer 行を引く
        key = (app.emp_id, app.target_date_iso)
        idx = row_index.get(key)
        if idx is None:
            stats.not_matched += 1
            stats.warnings.append(
                f"行ID={app.source_diff_row_id} jinjer CSV に該当行なし "
                f"(emp={app.emp_id} date={app.target_date_iso} {app.name})"
            )
            continue

        # 上書き（実績確定済かどうかに関わらず、承認されたものは上書きする）
        if app.kind == DIFF_KIND_PUNCH_IN:
            new_in = app.manual_fix_value or app.auto_fix_value
            _warn_if_not_hhmm(stats, app.source_diff_row_id, "出勤1への書込値", new_in,
                              app.emp_id, app.target_date_iso, app.name)
            rows[idx][punch_in_col] = new_in
            stats.overwritten_punch_in += 1
            changed_days.add(key)
            if not new_in:
                cleared_punch_days.add(key)
            # 人間判断で採用した就業開始時刻に、スケジュール開始（出勤予定時刻）も合わせる。
            # 採用値が空（jinjer打刻を消すケース）のときは予定を消さない。
            # 退勤予定時刻が空（スケジュール未設定）の行に出勤予定だけ書くと、jinjerが
            # 「退勤予定時刻が正しくありません」で行ごと弾くため、予定が揃っている行に限る。
            if (
                sched_in_col is not None and new_in
                and sched_out_col is not None
                and (rows[idx][sched_out_col] or "").strip()
            ):
                rows[idx][sched_in_col] = new_in
                stats.overwritten_sched_in += 1
        elif app.kind == DIFF_KIND_PUNCH_OUT:
            new_out = app.manual_fix_value or app.auto_fix_value
            # 夜勤で翌朝退勤の場合、jinjer は 24時超表記でないとインポートできないため
            # 現在の出勤1と突き合わせて 08:15→32:15 のように補正する（冪等・安全網）。
            new_out = to_overnight_punch_out(rows[idx][punch_in_col], new_out)
            _warn_if_not_hhmm(stats, app.source_diff_row_id, "退勤1への書込値", new_out,
                              app.emp_id, app.target_date_iso, app.name)
            rows[idx][punch_out_col] = new_out
            stats.overwritten_punch_out += 1
            changed_days.add(key)
            if not new_out:
                cleared_punch_days.add(key)

        # 打刻を書き込んだ行に「勤務状況（0:未打刻1:欠勤）」フラグが残っていると、
        # jinjerが「出勤1/退勤1/勤務状況が正しくありません」で弾くためクリアする。
        if (
            work_status_col is not None
            and (rows[idx][punch_in_col] or rows[idx][punch_out_col])
            and (rows[idx][work_status_col] or "").strip() in ("0", "1")
        ):
            rows[idx][work_status_col] = ""

        # 参考集計: 上書き対象が実績確定済だった件数
        if finalized_col is not None:
            finalized = (rows[idx][finalized_col] or "").strip().upper()
            if finalized == "TRUE":
                stats.overwritten_finalized += 1

    # ---- 打刻を削除（空で上書き）した日の扱い ----
    # jinjerインポートは打刻セルの「空」を削除とは解釈せず「出勤1が正しくありません」で
    # 行ごと弾く（2026-07-08実測: 上原さん6/14）。つまり打刻の削除はインポートでは
    # 反映できないため、削除を含む日はアップロードCSVから除外し、jinjer画面での
    # 手動対応を警告で案内する。
    for key in sorted(cleared_punch_days):
        idx = row_index.get(key)
        if idx is None:
            continue
        has_in = bool((rows[idx][punch_in_col] or "").strip())
        has_out = bool((rows[idx][punch_out_col] or "").strip())
        changed_days.discard(key)
        stats.manual_clear_days += 1
        if not has_in and not has_out:
            stats.warnings.append(
                f"⚠️要手動対応: 打刻の削除(空での上書き)はjinjerインポートでは反映できない"
                f"ため、この日はアップロードCSVから除外しました。jinjer画面で打刻を"
                f"手動削除してください (emp={key[0]} date={key[1]})"
            )
        else:
            remain = JINJER_COL_PUNCH_IN if has_in else JINJER_COL_PUNCH_OUT
            stats.warnings.append(
                f"⚠️要手動対応: 打刻の片側だけを削除する変更({remain}は残る)はjinjer"
                f"インポートでは反映できないため、この日はアップロードCSVから除外しました。"
                f"jinjer画面でこの日の打刻を手動修正してください (emp={key[0]} date={key[1]})"
            )

    return changed_days


def filter_output_rows(
    headers: list[str],
    rows: list[list[str]],
    changed_days: set[tuple[str, str]],
    held_days: set[tuple[str, str]],
    stats: Stats,
) -> list[list[str]]:
    """アップロードCSVに出力する行を「このツールが変更した (従業員, 日付)」だけに絞る。

    - 無変更の行を送り返しても意味がない上、jinjer側の既存不整合（休暇残数不足・
      勤務状況フラグの矛盾など、jinjer自身がエクスポートした行でも再インポートで
      弾かれるもの）でエラー通知が汚れるため、変更行だけを出力する。
    - 保留の日は変更されないので自然に出力から外れる（DL後にjinjer画面で手修正した
      日を保留にする運用の保護）。承認と保留が同じ日に混在する場合だけ行が残るため、
      警告して人に確認してもらう。
    """
    mixed = held_days & changed_days
    for emp_id, date_iso in sorted(mixed):
        stats.warnings.append(
            f"保留と承認(書き戻し)が同じ日に混在しています (emp={emp_id} date={date_iso})。"
            "行は出力されるため、保留側の項目はダウンロード時点の値がインポートされます。"
            "この日をjinjer画面で手修正している場合は、インポート前に値を確認してください。"
        )

    emp_col = headers.index(JINJER_COL_EMP_ID)
    date_col = headers.index(JINJER_COL_DATE)
    kept: list[list[str]] = []
    for row in rows:
        key = ((row[emp_col] or "").strip(), normalize_date_iso(row[date_col]))
        if key in changed_days:
            kept.append(row)
    stats.unchanged_rows_removed = len(rows) - len(kept)
    stats.held_rows_removed = len(held_days - changed_days)
    return kept


# ----------------------------------------------------------------------
# 出力
# ----------------------------------------------------------------------

def write_csv(output_path: Path, headers: list[str], rows: list[list[str]]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    # CRLF + CP932
    with open(output_path, "w", encoding="cp932", newline="") as f:
        writer = csv.writer(f, lineterminator="\r\n")
        writer.writerow(headers)
        writer.writerows(rows)


def print_summary(stats: Stats, output_path: Path, dry_run: bool, total_rows: int) -> None:
    mode = "[dry-run]" if dry_run else "[execute]"
    print()
    print(f"{mode} ===== 差異一覧の判断状況 =====")
    print(f"  総差異件数      : {stats.total_diff_rows}")
    print(f"  請求勤怠を正(書戻): {stats.approved}")
    print(f"  jinjer勤怠を正    : {stats.rejected}")
    print(f"  保留            : {stats.held}")
    print(f"  未判断          : {stats.pending}")
    if stats.recovered_misplaced:
        detail = "、".join(f"{col}={cnt}" for col, cnt in stats.misplaced_by_col.items())
        print(f"  ※ 判断の列ミスを回収: {stats.recovered_misplaced} 件 ({detail})")
    print()
    print(f"{mode} ===== 承認の種別内訳 =====")
    print(f"  出勤差異   : {stats.approved_punch_in} → 出勤1 列に上書き")
    print(f"  退勤差異   : {stats.approved_punch_out} → 退勤1 列に上書き")
    print(f"  休憩差異   : {stats.approved_break} → 手入力欄がある場合のみ休憩列へ上書き")
    print(f"  総労働時間 : {stats.approved_total} → 警告のみ、上書きしない")
    print(f"  手入力休憩 : {stats.manual_break_days} 日分（差異種別を問わず入力があれば反映）")
    if stats.manual_break_conflicts:
        print(f"  ⚠️ 手入力休憩の値の食い違い: {stats.manual_break_conflicts} 件（警告参照）")
    print()
    print(f"{mode} ===== 実際の上書き結果 =====")
    print(f"  出勤1 上書き         : {stats.overwritten_punch_in}")
    print(f"  └ 出勤予定時刻も更新 : {stats.overwritten_sched_in}")
    print(f"  スケジュール開始合わせ: {stats.overwritten_sched_start}（出勤予定時刻のみ・打刻は触らない）")
    print(f"  退勤1 上書き         : {stats.overwritten_punch_out}")
    print(f"  休憩1 上書き         : {stats.overwritten_break_start}")
    print(f"  復帰1 上書き         : {stats.overwritten_break_end}")
    print(f"  休憩時間 上書き      : {stats.overwritten_break_total}")
    print(f"  休憩スキップ         : {stats.skipped_break}")
    print(f"  総労働スキップ       : {stats.skipped_total}")
    print(f"  無変更行を除外       : {stats.unchanged_rows_removed}（変更した日の行だけを出力）")
    print(f"  └ うち保留日        : {stats.held_rows_removed} 日分（jinjer手修正の保護・インポートで触らない）")
    if stats.manual_clear_days:
        print(f"  ⚠️ 打刻削除の日を除外 : {stats.manual_clear_days} 日分（インポートでは削除不可。jinjer画面で手動対応→警告参照）")
    print(f"  jinjer行 未マッチ    : {stats.not_matched}")
    print(f"  (うち実績確定済を上書き: {stats.overwritten_finalized} 件 / 参考)")
    print()
    print(f"{mode} 出力CSV総行数: {total_rows + 1} 行 (ヘッダー含む)")
    if dry_run:
        print(f"{mode} 出力先（書き出されません）: {output_path}")
        print(f"{mode} 本番書き出しは --execute を指定してください")
    else:
        print(f"[done] 出力完了: {output_path}")

    if stats.warnings:
        print()
        print(f"{mode} ===== 警告 ({len(stats.warnings)} 件) =====")
        for w in stats.warnings[:30]:
            print(f"  - {w}")
        if len(stats.warnings) > 30:
            print(f"  ... ほか {len(stats.warnings) - 30} 件")


# ----------------------------------------------------------------------
# 共通実行関数（CLI / Flask 共用）
# ----------------------------------------------------------------------

@dataclass
class ExportResult:
    ok: bool
    output_path: Path
    dry_run: bool
    stats: Stats = field(default_factory=Stats)
    total_jinjer_rows: int = 0
    error: str = ""


def run_quick_export(
    diff_xlsx: Path,
    jinjer_dir: Path,
    output_path: Path,
    dry_run: bool,
    log_func=print,
) -> ExportResult:
    """差異一覧xlsx + jinjer CSV → アップロード用CSV を生成。CLI と Web UI から共用。"""
    stats = Stats()
    result = ExportResult(ok=False, output_path=output_path, dry_run=dry_run, stats=stats)

    log_func(f"[start] 差異一覧: {diff_xlsx}")
    log_func(f"[start] jinjer CSV/フォルダ: {jinjer_dir}")
    log_func(f"[start] 出力先: {output_path}")
    log_func(f"[start] モード: {'dry-run' if dry_run else 'execute'}")

    try:
        approved, manual_breaks, held_days = load_approved_rows(diff_xlsx, stats)
    except Exception as e:
        msg = f"差異一覧読み込み失敗: {e}"
        log_func(f"[error] {msg}")
        result.error = msg
        return result
    log_func(
        f"[info] 承認行 {len(approved)} 件 / 手入力休憩 {len(manual_breaks)} 日分 / "
        f"総差異 {stats.total_diff_rows} 件"
    )

    try:
        headers, rows = load_jinjer_csvs(jinjer_dir)
    except Exception as e:
        msg = f"jinjer CSV 読み込み失敗: {e}"
        log_func(f"[error] {msg}")
        result.error = msg
        return result
    log_func(f"[info] jinjer 行 {len(rows)} 件 / 列数 {len(headers)}")
    result.total_jinjer_rows = len(rows)

    try:
        row_index = build_jinjer_row_index(headers, rows)
    except Exception as e:
        msg = f"jinjer 行インデックス構築失敗: {e}"
        log_func(f"[error] {msg}")
        result.error = msg
        return result
    log_func(f"[info] (従業員ID, 年月日) インデックス {len(row_index)} 件")

    changed_days = apply_approved_rows(headers, rows, row_index, approved, stats, manual_breaks)

    # スケジュール開始合わせ（別シート）: 実績が予定より早い日の出勤予定時刻を黙って合わせる。
    # 差異一覧の行ではないので人間判断は無い＝保留日だけ除外して自動反映する。
    try:
        sched_aligns = load_sched_aligns(diff_xlsx)
    except Exception as e:  # noqa: BLE001 — 参考シートの読み損ねで全体を止めない
        sched_aligns = []
        stats.warnings.append(f"スケジュール開始合わせシートの読み込みに失敗: {e}")
    if sched_aligns:
        apply_sched_aligns(headers, rows, row_index, sched_aligns, held_days, stats, changed_days)
        log_func(f"[info] スケジュール開始合わせ: {stats.overwritten_sched_start} 件を反映"
                 f"（出勤予定時刻のみ・打刻は触らない）")

    # 変更した日の行だけを出力する（無変更行はjinjer側の既存不整合で弾かれて
    # エラー通知が汚れるだけ。保留の日も変更されないので自然に出力から外れる＝
    # ダウンロード後のjinjer手修正を巻き戻さない）。
    rows = filter_output_rows(headers, rows, changed_days, held_days, stats)
    result.total_jinjer_rows = len(rows)
    log_func(
        f"[info] 変更のあった {len(rows)} 行のみ出力"
        f"（無変更 {stats.unchanged_rows_removed} 行を除外・うち保留 {stats.held_rows_removed} 日分）"
    )

    if not dry_run:
        try:
            write_csv(output_path, headers, rows)
        except Exception as e:
            msg = f"CSV 書き出し失敗: {e}"
            log_func(f"[error] {msg}")
            result.error = msg
            return result

    print_summary(stats, output_path, dry_run, len(rows))
    result.ok = True
    return result


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="5月本番 MVP jinjer アップロード用 全件CSV生成")
    p.add_argument("--diff", required=True, help="差異一覧xlsx（人間判断埋め込み済み）")
    p.add_argument("--jinjer-dir", required=True, help="元 jinjer ダウンロードCSV のフォルダ")
    p.add_argument("--output", required=True, help="出力 CSV パス (CP932)")
    p.add_argument("--execute", action="store_true", help="実際に CSV を書き出す（既定は dry-run）")
    return p.parse_args()


def _unquote_path(s: str) -> str:
    """前後の空白と、前後が同じクォートで囲まれているときだけ1組を除去する。"""
    s = (s or "").strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ('"', "'"):
        s = s[1:-1].strip()
    return s


def main() -> int:
    args = parse_args()
    result = run_quick_export(
        diff_xlsx=Path(_unquote_path(args.diff)),
        jinjer_dir=Path(_unquote_path(args.jinjer_dir)),
        output_path=Path(_unquote_path(args.output)),
        dry_run=not args.execute,
    )
    return 0 if result.ok else 1


if __name__ == "__main__":
    sys.exit(main())
