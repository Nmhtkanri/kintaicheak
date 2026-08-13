"""差異行のトリアージ（要確認 / 自動採用 / 自動OK）分類。

差異一覧の判断件数を減らすため、各差異行を3区分に仕分けし、既定の「人間判断」を付ける。
人が読むのは「要確認」だけにし、自動採用/自動OK は既定値を入れて一括処理できるようにする。

純粋関数。quick_compare の DiffRow フィールドを引数で受ける（循環参照を避ける）。
kind は quick_compare の DIFF_KIND_PUNCH_IN/OUT（"出勤"/"退勤"）等と同じ文字列。
"""
from __future__ import annotations

# トリアージ区分
TRIAGE_NEEDS_CHECK = "要確認"
TRIAGE_AUTO_KINTAI = "自動(請求勤怠)"
TRIAGE_AUTO_OK = "自動OK(jinjer)"
TRIAGE_INFO_ONLY = "参考のみ"  # 手順3で書き戻さない・判断不要（INFOの総労働/休憩等）

# 既定の人間判断（quick_compare / quick_export と同じ語彙）
JUDGE_KINTAI = "請求勤怠"   # 請求勤怠を正 → jinjer へ書き戻す
JUDGE_JINJER = "jinjer勤怠"  # jinjer を正 → 書き戻さない
JUDGE_HOLD = "保留"          # 人が個別に確認する（書き戻さない）

# 警告レベル（quick_compare の LEVEL_* と一致させる）
LEVEL_DANGER = "DANGER"
LEVEL_WARN = "WARN"

# 出退勤の差異種別（quick_compare の DIFF_KIND_PUNCH_* と一致）
_PUNCH_KINDS = ("出勤", "退勤")
# スケジュール開始合わせ（quick_compare の DIFF_KIND_SCHED_START と一致）。
# 実績が予定より早い日の出勤予定時刻を請求勤怠へ寄せる行＝既定は自動採用。
_SCHED_START_KIND = "スケジュール開始"
# 半休（全日のみ自動OK対象。AM/PM半休は要確認に残す）
_HALF_DAY_TYPES = ("AM有休", "PM有休")
_FULL_DAY = "全日"


def is_zero_or_empty(value) -> bool:
    """請求勤怠の値が「0」または空か（全日休暇・休日の判定用）。"""
    s = str(value or "").strip()
    return s in ("", "0", "0:00", "00:00", "0:00:00", "00:00:00", "-")


# 旧名の別名（後方互換）
_is_zero_or_empty = is_zero_or_empty


def classify(
    *,
    kind: str,
    warn_level: str,
    punch_comment: str = "",
    stamp_comment: str = "",
    kintai_value: str = "",
    holiday_name1: str = "",
    holiday_name1_type: str = "",
) -> tuple[str, str]:
    """差異行を (トリアージ区分, 既定の人間判断) に分類する。

    既定の人間判断: 要確認="" / 自動採用="請求勤怠" / 自動OK="jinjer勤怠"。

    優先順位:
      1. DANGER（労基リスク）は必ず人が見る。
      2. 全日休暇＋打刻なし → jinjer が正（自動OK）。
      3. 打刻時/打刻修正コメントあり → 例外候補、人が読む。
      4. WARN（大差分・長時間・休憩過多）→ 人が見る。
      5. AM/PM 半休 → 人が見る。
      6. 出勤/退勤の小差分（INFO・コメント無し）→ 請求勤怠を自動採用。
      7. それ以外（INFOの休憩/総労働 等。手順3で書き戻せず判断不要）→ 参考のみ。
    """
    hol = str(holiday_name1 or "").strip()
    hol_type = str(holiday_name1_type or "").strip()
    has_comment = bool(str(punch_comment or "").strip() or str(stamp_comment or "").strip())

    if warn_level == LEVEL_DANGER:
        return (TRIAGE_NEEDS_CHECK, "")
    if hol and hol_type == _FULL_DAY and is_zero_or_empty(kintai_value):
        return (TRIAGE_AUTO_OK, JUDGE_JINJER)
    if has_comment:
        return (TRIAGE_NEEDS_CHECK, "")
    if warn_level == LEVEL_WARN:
        return (TRIAGE_NEEDS_CHECK, "")
    if hol_type in _HALF_DAY_TYPES:
        return (TRIAGE_NEEDS_CHECK, "")
    if kind in _PUNCH_KINDS or kind == _SCHED_START_KIND:
        return (TRIAGE_AUTO_KINTAI, JUDGE_KINTAI)
    # INFOの休憩/総労働 等は手順3で書き戻せず判断対象外 → 参考のみ（要確認に積まない）
    return (TRIAGE_INFO_ONLY, "")


# 表示・並べ替え用の優先度（小さいほど上）
TRIAGE_ORDER = {
    TRIAGE_NEEDS_CHECK: 0,
    TRIAGE_AUTO_OK: 1,
    TRIAGE_AUTO_KINTAI: 2,
    TRIAGE_INFO_ONLY: 3,
}
