import os
import sys
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from quick_compare import (  # noqa: E402
    DIFF_COLUMNS,
    DIFF_KIND_PUNCH_IN,
    DIFF_KIND_PUNCH_OUT,
    DIFF_KIND_TOTAL,
    JINJER_HEADERS,
    LogEntry,
    classify_total_work,
    clean_punch_comment,
    compute_diffs,
    format_stamp_comments,
    kintai_total_minutes,
    load_stamp_correction_reasons,
    normalize_kintai_result_columns,
    recommend_judge_label,
    resolve_jinjer_extra_columns,
    strip_punch_noise_words,
    to_jinjer_overnight_punch_out,
)


def _jinjer_row(**overrides):
    row = {
        JINJER_HEADERS["name"]: "上原 奏吾",
        JINJER_HEADERS["emp_id"]: "2018057",
        JINJER_HEADERS["date"]: "2026/4/1",
        JINJER_HEADERS["punch_in_1"]: "",
        JINJER_HEADERS["punch_out_1"]: "",
        JINJER_HEADERS["break_total"]: "",
        JINJER_HEADERS["total_work"]: "",
        JINJER_HEADERS["finalized"]: "",
    }
    row.update(overrides)
    return row


def test_compute_diffs_creates_punch_rows_when_jinjer_punches_are_blank():
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00",
        "jinjer_出勤": "",
        "出勤差分(分)": None,
        "勤務表_退勤": "18:00",
        "jinjer_退勤": "",
        "退勤差分(分)": None,
        "_source_file": "勤怠突合結果.xlsx",
    }])
    logs: list[LogEntry] = []

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): _jinjer_row()},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
    )

    punch_rows = [row for row in rows if row.kind in {DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT}]
    assert [row.kind for row in punch_rows] == [DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT]
    assert [row.auto_fix_value for row in punch_rows] == ["9:00", "18:00"]
    assert punch_rows[0].warn_reason == "jinjer出勤なし / 請求勤怠側に時刻あり"
    assert punch_rows[1].warn_reason == "jinjer退勤なし / 請求勤怠側に時刻あり"


def test_compute_diffs_creates_punch_rows_when_kintai_is_blank():
    """逆向きの片側欠落: jinjer に打刻あり / 請求勤怠側に時刻なし → 要確認行を生成。"""
    from services.triage import TRIAGE_NEEDS_CHECK

    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "",
        "jinjer_出勤": "9:00",
        "出勤差分(分)": None,
        "勤務表_退勤": "",
        "jinjer_退勤": "18:00",
        "退勤差分(分)": None,
        "_source_file": "勤怠突合結果.xlsx",
    }])
    logs: list[LogEntry] = []

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): _jinjer_row(
            **{JINJER_HEADERS["punch_in_1"]: "9:00", JINJER_HEADERS["punch_out_1"]: "18:00"})},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
    )

    punch_rows = [row for row in rows if row.kind in {DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT}]
    assert [row.kind for row in punch_rows] == [DIFF_KIND_PUNCH_IN, DIFF_KIND_PUNCH_OUT]
    # jinjer値が表示され、請求勤怠値は空。承認時は空で上書き（jinjer打刻を消す）。
    assert [row.jinjer_value for row in punch_rows] == ["9:00", "18:00"]
    assert [row.kintai_value for row in punch_rows] == ["", ""]
    assert [row.auto_fix_value for row in punch_rows] == ["", ""]
    assert punch_rows[0].warn_reason == "請求勤怠なし / jinjer側に時刻あり"
    assert punch_rows[1].warn_reason == "請求勤怠なし / jinjer側に時刻あり"
    # コメント・休暇なしなら要確認（人が見る）に分類される
    assert punch_rows[0].triage == TRIAGE_NEEDS_CHECK
    assert punch_rows[0].judge_default == ""


def test_compute_diffs_kintai_blank_fullday_holiday_is_auto_ok():
    """全日休暇日の jinjer 打刻（請求勤怠なし）は triage が自動OK(jinjer)へ回す。"""
    from services.triage import TRIAGE_AUTO_OK, JUDGE_JINJER

    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "", "jinjer_出勤": "9:00", "出勤差分(分)": None,
        "勤務表_退勤": "", "jinjer_退勤": "18:00", "退勤差分(分)": None,
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    jrow = _jinjer_row(**{
        JINJER_HEADERS["punch_in_1"]: "9:00", JINJER_HEADERS["punch_out_1"]: "18:00",
        "休日休暇名1": "有給休暇", "休日休暇名1：種別": "全日",
    })
    extra_cols = resolve_jinjer_extra_columns(list(jrow.keys()))

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): jrow},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
        extra_cols,
    )
    punch_rows = [r for r in rows if r.kind == DIFF_KIND_PUNCH_IN]
    assert punch_rows
    assert punch_rows[0].triage == TRIAGE_AUTO_OK
    assert punch_rows[0].judge_default == JUDGE_JINJER


def test_compute_diffs_surfaces_jinjer_unregistered_worker():
    """jinjerに氏名が無い(ID解決不可)が請求勤怠に勤務がある人は、捨てずに
    「jinjer未登録」要確認行(1人1行)として可視化する。"""
    from quick_compare import DIFF_KIND_UNMATCHED
    from services.triage import TRIAGE_NEEDS_CHECK

    kintai_df = pd.DataFrame([
        {"氏名": "幽霊 太郎", "日付": date(2026, 5, 1),
         "勤務表_出勤": "9:00", "jinjer_出勤": "", "出勤差分(分)": None,
         "勤務表_退勤": "18:00", "jinjer_退勤": "", "退勤差分(分)": None,
         "_source_file": "月給制1.xlsx"},
        {"氏名": "幽霊 太郎", "日付": date(2026, 5, 2),
         "勤務表_出勤": "9:00", "jinjer_出勤": "", "出勤差分(分)": None,
         "勤務表_退勤": "17:30", "jinjer_退勤": "", "退勤差分(分)": None,
         "_source_file": "月給制1.xlsx"},
        # 請求勤怠も空の日（休等）は未払い候補にしない＝行に数えない
        {"氏名": "幽霊 太郎", "日付": date(2026, 5, 3),
         "勤務表_出勤": "", "jinjer_出勤": "", "出勤差分(分)": None,
         "勤務表_退勤": "", "jinjer_退勤": "", "退勤差分(分)": None,
         "_source_file": "月給制1.xlsx"},
    ])
    logs: list[LogEntry] = []

    # name_map に「幽霊 太郎」は無い → 解決不可
    rows = compute_diffs(
        kintai_df,
        {},  # jinjer index 空
        {"別人": "2018999"},
        logs,
    )

    unmatched = [r for r in rows if r.kind == DIFF_KIND_UNMATCHED]
    assert len(unmatched) == 1, "未登録者は1人1行に集約"
    r = unmatched[0]
    assert r.name == "幽霊 太郎"
    assert r.emp_id == ""
    assert r.kintai_value == "勤務2日"   # 勤務がある日のみ計上（5/3は除外）
    assert r.triage == TRIAGE_NEEDS_CHECK
    assert r.judge_default == ""
    assert r.recommend_judge == ""       # 採用ラベルは付けない
    assert "jinjer未登録" in r.warn_reason


def test_compute_diffs_unresolved_with_no_work_emits_nothing():
    """ID解決不可でも請求勤怠に勤務が無ければ未登録行は出さない（休のみの人）。"""
    from quick_compare import DIFF_KIND_UNMATCHED
    kintai_df = pd.DataFrame([{
        "氏名": "休 のみ子", "日付": date(2026, 5, 1),
        "勤務表_出勤": "", "jinjer_出勤": "", "出勤差分(分)": None,
        "勤務表_退勤": "", "jinjer_退勤": "", "退勤差分(分)": None,
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    rows = compute_diffs(kintai_df, {}, {"別人": "2018999"}, logs)
    assert [r for r in rows if r.kind == DIFF_KIND_UNMATCHED] == []


def test_normalize_kintai_result_columns_accepts_seikyu_kintai_headers():
    df = pd.DataFrame([{
        "請求勤怠_出勤": "9:00",
        "請求勤怠_退勤": "18:00",
        "請求勤怠_総労働": "9:00",
    }])

    normalized = normalize_kintai_result_columns(df)

    assert normalized.iloc[0]["勤務表_出勤"] == "9:00"
    assert normalized.iloc[0]["勤務表_退勤"] == "18:00"
    assert normalized.iloc[0]["勤務表_総労働"] == "9:00"


def test_to_jinjer_overnight_punch_out():
    # 翌朝退勤（出勤 > 退勤）→ 24時超表記へ変換
    assert to_jinjer_overnight_punch_out("21:00", "08:15") == "32:15"
    assert to_jinjer_overnight_punch_out("17:00", "01:30") == "25:30"
    # 通常勤務（出勤 < 退勤）→ そのまま
    assert to_jinjer_overnight_punch_out("9:00", "18:00") == "18:00"
    # すでに24時超表記なら再変換しない（冪等）
    assert to_jinjer_overnight_punch_out("21:00", "32:15") == "32:15"
    # 出退勤のどちらかが空/不正ならそのまま
    assert to_jinjer_overnight_punch_out("", "08:15") == "08:15"
    assert to_jinjer_overnight_punch_out("21:00", "") == ""


def test_compute_diffs_converts_overnight_punch_out_to_jinjer_format():
    """夜勤（請求勤怠 出勤21:00→退勤翌08:15）の退勤提案値が 32:15 になる。"""
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "21:00",
        "jinjer_出勤": "21:00",
        "出勤差分(分)": 0,
        "勤務表_退勤": "08:15",
        "jinjer_退勤": "",
        "退勤差分(分)": None,
        "_source_file": "勤怠突合結果.xlsx",
    }])
    logs: list[LogEntry] = []

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): _jinjer_row(**{JINJER_HEADERS["punch_in_1"]: "21:00"})},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
    )

    out_rows = [r for r in rows if r.kind == DIFF_KIND_PUNCH_OUT]
    assert len(out_rows) == 1
    # 請求勤怠値（表示）は元の 08:15、自動修正提案値は jinjer 用の 32:15
    assert out_rows[0].kintai_value == "08:15"
    assert out_rows[0].auto_fix_value == "32:15"


def test_diff_columns_include_manual_review_fields():
    # 「手入力修正値」は「打刻修正」に改名済み
    assert "打刻修正" in DIFF_COLUMNS
    assert "手入力修正値" not in DIFF_COLUMNS
    assert "手入力休憩1" in DIFF_COLUMNS
    assert "手入力復帰1" in DIFF_COLUMNS
    assert "手入力休憩時間" in DIFF_COLUMNS


def test_diff_columns_include_schedule_and_leave_fields():
    for col in ("出勤予定", "退勤予定", "休憩予定", "復帰予定", "休日休暇名1", "休日休暇名1：種別"):
        assert col in DIFF_COLUMNS
    # 有休/AM有休/PM有休 は汎用データ194列に該当列が無く常に空のため出力しない
    for col in ("有休", "AM有休", "PM有休"):
        assert col not in DIFF_COLUMNS


def test_resolve_jinjer_extra_columns_exact_and_partial():
    cols = [
        "名前", "*従業員ID", "*年月日",
        "出勤予定時刻", "退勤予定時刻", "休憩予定時間",
        "有休", "AM有休", "PM有休",
    ]
    resolved = resolve_jinjer_extra_columns(cols)
    assert resolved["出勤予定"] == "出勤予定時刻"
    assert resolved["退勤予定"] == "退勤予定時刻"
    assert resolved["休憩予定"] == "休憩予定時間"
    assert resolved["有休"] == "有休"
    assert resolved["AM有休"] == "AM有休"
    assert resolved["PM有休"] == "PM有休"


def test_resolve_jinjer_extra_columns_leave_is_exact_only():
    # 「有休」列が無く AM有休/PM有休 だけある場合、「有休」を部分一致で誤ヒットさせない
    cols = ["名前", "AM有休", "PM有休"]
    resolved = resolve_jinjer_extra_columns(cols)
    assert "有休" not in resolved
    assert resolved.get("AM有休") == "AM有休"
    assert resolved.get("PM有休") == "PM有休"


def test_resolve_jinjer_extra_columns_missing_returns_no_key():
    resolved = resolve_jinjer_extra_columns(["名前", "*従業員ID", "*年月日"])
    assert resolved == {}


def test_compute_diffs_transcribes_schedule_and_leave():
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00",
        "jinjer_出勤": "",
        "出勤差分(分)": None,
        "勤務表_退勤": "18:00",
        "jinjer_退勤": "",
        "退勤差分(分)": None,
        "_source_file": "勤怠突合結果.xlsx",
    }])
    logs: list[LogEntry] = []
    jrow = _jinjer_row(**{
        "出勤予定時刻": "9:00",
        "退勤予定時刻": "18:00",
        "休憩予定時刻1": "12:00",
        "復帰予定時刻1": "13:00",
        "有休": "",
        "AM有休": "1",
        "PM有休": "",
    })
    extra_cols = resolve_jinjer_extra_columns(list(jrow.keys()))

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): jrow},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
        extra_cols,
    )

    assert rows, "差異行が生成されること"
    r = rows[0]
    assert r.sched_in == "9:00"
    assert r.sched_out == "18:00"
    assert r.sched_break == "12:00"     # 休憩予定時刻1（部分一致）
    assert r.sched_break_end == "13:00"  # 復帰予定時刻1（新規）
    assert r.am_yukyu == "1"
    assert r.yukyu == ""
    assert r.pm_yukyu == ""


def test_kintai_total_minutes_prefers_actual_work():
    """請求勤怠の正味労働(勤務表_実働時間)が拘束時間より優先される。"""
    row = pd.Series({
        "勤務表_実働時間": "7:00",     # 正味（休憩控除後）
        "勤務表_総労働時間": "9:00",   # 拘束（退勤−出勤）
        "勤務表_出勤": "9:00",
        "勤務表_退勤": "18:00",
    })
    minutes, hhmm = kintai_total_minutes(row)
    assert minutes == 420
    assert hhmm == "7:00"


def test_kintai_total_minutes_fallback_when_actual_blank():
    """実働列が空なら従来の拘束時間にフォールバックする（後方互換）。"""
    row = pd.Series({"勤務表_実働時間": "", "勤務表_総労働時間": "9:00"})
    minutes, _ = kintai_total_minutes(row)
    assert minutes == 540


def test_compute_diffs_total_compares_net_vs_net():
    """総労働差異は請求勤怠の正味(実働) vs jinjer 総労働 で突合される。

    拘束時間9:00ではなく実働8:00で比較されるため、jinjer総労働8:00とは差異なし。
    """
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00", "jinjer_出勤": "9:00", "出勤差分(分)": 0,
        "勤務表_退勤": "18:00", "jinjer_退勤": "18:00", "退勤差分(分)": 0,
        "勤務表_実働時間": "8:00",  # 正味（休憩1h控除後）
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    jrow = _jinjer_row(**{
        JINJER_HEADERS["total_work"]: "8:00",
        JINJER_HEADERS["break_total"]: "1:00",
        JINJER_HEADERS["punch_in_1"]: "9:00",
        JINJER_HEADERS["punch_out_1"]: "18:00",
    })
    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): jrow},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
    )
    total_rows = [r for r in rows if r.kind == DIFF_KIND_TOTAL]
    assert total_rows == []  # 実働8:00 = jinjer総労働8:00 → 差異なし


def test_format_stamp_comments():
    """打刻コメント list → 種別[方法] コメント / ... の整形。"""
    items = [
        {"type": "出勤", "method": "打刻修正申請", "comment": "KDX出社"},
        {"type": "退勤", "method": "PC", "comment": "私用のため早退"},
    ]
    assert format_stamp_comments(items) == "出勤[打刻修正申請] KDX出社 / 退勤[PC] 私用のため早退"
    assert format_stamp_comments([]) == ""
    assert format_stamp_comments([{"comment": "", "type": "出勤", "method": "PC"}]) == ""


def test_compute_diffs_attaches_stamp_comment():
    """打刻修正申請の従業員コメントが当日の差異行に併記される。"""
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00", "jinjer_出勤": "9:30", "出勤差分(分)": 30,
        "勤務表_退勤": "18:00", "jinjer_退勤": "18:00", "退勤差分(分)": 0,
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    stamp_comments = {
        ("2018057", "2026-04-01"): [
            {"type": "出勤", "method": "打刻修正申請", "comment": "KDX出社"},
        ],
    }
    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): _jinjer_row()},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
        None,
        stamp_comments,
    )
    punch_rows = [r for r in rows if r.kind == DIFF_KIND_PUNCH_IN]
    assert punch_rows
    assert punch_rows[0].jinjer_stamp_comment == "出勤[打刻修正申請] KDX出社"


def test_comment_columns_left_of_judgment_and_adjacent():
    """打刻時コメント／打刻修正時コメントは隣同士で、人間判断より左（判断材料）に置く。"""
    assert "打刻時コメント" in DIFF_COLUMNS
    assert "打刻修正時コメント" in DIFF_COLUMNS
    assert "jinjer打刻コメント" not in DIFF_COLUMNS
    # 隣同士（打刻時コメント → 打刻修正時コメント）
    assert DIFF_COLUMNS.index("打刻時コメント") + 1 == DIFF_COLUMNS.index("打刻修正時コメント")
    # どちらも人間判断より左
    assert DIFF_COLUMNS.index("打刻修正時コメント") < DIFF_COLUMNS.index("人間判断")


def test_diff_columns_layout_identity_first():
    """横スクロール対策: 識別4列(従業員ID/氏名/対象日付/差異種別)が先頭。手入力は右へ。"""
    assert DIFF_COLUMNS[:4] == ["従業員ID", "氏名", "対象日付", "差異種別"]
    # 従業員IDは先頭(A列)へ移動済み
    assert DIFF_COLUMNS.index("従業員ID") == 0
    # 打刻修正(手入力)は人間判断より右
    assert DIFF_COLUMNS.index("打刻修正") > DIFF_COLUMNS.index("人間判断")
    # 行ID・元突合結果ファイルは削除済み
    assert "行ID" not in DIFF_COLUMNS
    assert "元突合結果ファイル" not in DIFF_COLUMNS


def test_clean_punch_comment():
    """汎用データ#96『打刻時コメント』の整形（空ラベルのゴミは除去）。"""
    from quick_compare import clean_punch_comment
    assert clean_punch_comment("出勤: KDX出社 , 退勤:  , ") == "出勤: KDX出社"
    assert clean_punch_comment("出勤:  , 退勤:  , ") == ""
    assert clean_punch_comment("出勤:  , 退勤: テレワーク , ") == "退勤: テレワーク"
    assert clean_punch_comment("") == ""
    assert clean_punch_comment(None) == ""


def test_triage_column_present_and_placed():
    """確認区分（旧トリアージ区分）は差分とコメント列の間に置く。警告レベル列は廃止。"""
    assert "確認区分" in DIFF_COLUMNS
    assert "トリアージ区分" not in DIFF_COLUMNS
    assert "警告レベル" not in DIFF_COLUMNS
    assert DIFF_COLUMNS.index("差分(分)") < DIFF_COLUMNS.index("確認区分")
    assert DIFF_COLUMNS.index("確認区分") < DIFF_COLUMNS.index("打刻時コメント")


def test_schedule_leave_columns_left_of_judgment():
    """有休も判断材料なので、予定/休日休暇は人間判断より左に置く。"""
    for col in ("出勤予定", "退勤予定", "休憩予定", "復帰予定", "休日休暇名1", "休日休暇名1：種別"):
        assert DIFF_COLUMNS.index(col) < DIFF_COLUMNS.index("人間判断")
    # 復帰予定は 休憩予定 と 休日休暇名1 の間に挿入
    assert DIFF_COLUMNS.index("休憩予定") < DIFF_COLUMNS.index("復帰予定") < DIFF_COLUMNS.index("休日休暇名1")


def test_human_judgment_conditional_formatting(tmp_path):
    """人間判断の値に応じた条件付き書式（保留/jinjer勤怠）が差異一覧に入る。"""
    import html
    import zipfile
    from quick_compare import write_excel, DiffRow

    drow = DiffRow(
        row_id=1, emp_id="1001", name="山田 太郎", target_date="2026-05-01", kind="出勤",
        kintai_value="9:00", jinjer_value="9:30", diff_minutes="30",
        warn_level="INFO", warn_reason="", auto_fix_value="9:00",
        finalized="", source_file="x.xlsx",
    )
    out = tmp_path / "d.xlsx"
    write_excel(out, [drow], [], "2026-05")

    with zipfile.ZipFile(out) as z:
        sheets = " ".join(
            z.read(n).decode("utf-8") for n in z.namelist()
            if n.startswith("xl/worksheets/sheet")
        )
    assert "conditionalFormatting" in sheets
    # openpyxl は数式内の日本語を数値文字参照にエスケープするので unescape して判定
    decoded = html.unescape(sheets)
    assert '"保留"' in decoded
    assert '"jinjer勤怠"' in decoded
    # cellIs ルールが2件（保留 / jinjer勤怠）入っている
    assert decoded.count('type="cellIs"') >= 2


def test_recommend_judge_label_new_rule():
    """新ルール: コメントあり→jinjer勤怠 / 差分≧しきい値or片側欠落→保留 / それ以外→請求勤怠。"""
    # コメントなし & 差分がしきい値(10)未満 → 請求勤怠
    assert recommend_judge_label(DIFF_KIND_PUNCH_IN, "9", 10) == "請求勤怠"
    assert recommend_judge_label(DIFF_KIND_PUNCH_OUT, "-9", 10) == "請求勤怠"
    # コメントなし & 差分がしきい値以上 → 保留（10ちょうども「以上」で保留）
    assert recommend_judge_label(DIFF_KIND_PUNCH_IN, "10", 10) == "保留"
    assert recommend_judge_label(DIFF_KIND_PUNCH_OUT, "11", 10) == "保留"
    assert recommend_judge_label(DIFF_KIND_PUNCH_IN, "-30", 10) == "保留"
    # 片側欠落（差分なし）→ 保留（安全側）
    assert recommend_judge_label(DIFF_KIND_PUNCH_IN, "", 10) == "保留"
    # コメントあり → jinjer勤怠（差分の大小に関わらず）
    assert recommend_judge_label(DIFF_KIND_PUNCH_IN, "5", 10, has_comment=True) == "jinjer勤怠"
    assert recommend_judge_label(DIFF_KIND_PUNCH_OUT, "120", 10, has_comment=True) == "jinjer勤怠"
    # 総労働 → 対象外＝jinjer勤怠（コメント有無に関わらず変更なし）
    assert recommend_judge_label(DIFF_KIND_TOTAL, "5", 10) == "jinjer勤怠"
    assert recommend_judge_label(DIFF_KIND_TOTAL, "5", 10, has_comment=True) == "jinjer勤怠"


def test_recommend_judge_label_set_on_rows():
    """compute_diffs が各差異行に recommend_judge を設定する（新ルール）。"""
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:05", "jinjer_出勤": "9:00", "出勤差分(分)": 5,
        "勤務表_退勤": "20:00", "jinjer_退勤": "18:00", "退勤差分(分)": 120,
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): _jinjer_row(
            **{JINJER_HEADERS["punch_in_1"]: "9:00", JINJER_HEADERS["punch_out_1"]: "18:00"})},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
        threshold_minutes=10,
    )
    in_row = next(r for r in rows if r.kind == DIFF_KIND_PUNCH_IN)
    out_row = next(r for r in rows if r.kind == DIFF_KIND_PUNCH_OUT)
    assert in_row.recommend_judge == "請求勤怠"   # 差分5分 < 10 / コメントなし
    assert out_row.recommend_judge == "保留"      # 差分120分 ≥ 10 / コメントなし


def test_recommend_judge_label_comment_makes_jinjer():
    """打刻時コメントがある出退勤差異は、差分が大きくても jinjer勤怠 を提案する。"""
    jrow = _jinjer_row(**{
        JINJER_HEADERS["punch_in_1"]: "9:30", JINJER_HEADERS["punch_out_1"]: "18:00",
        "打刻時コメント": "出勤: 客先直行のため",
    })
    extra_cols = resolve_jinjer_extra_columns(list(jrow.keys()))
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾", "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00", "jinjer_出勤": "9:30", "出勤差分(分)": 30,
        "勤務表_退勤": "18:00", "jinjer_退勤": "18:00", "退勤差分(分)": 0,
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): jrow},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs, extra_cols, threshold_minutes=10,
    )
    in_row = next(r for r in rows if r.kind == DIFF_KIND_PUNCH_IN)
    assert in_row.punch_comment  # コメントが転記されている
    assert in_row.recommend_judge == "jinjer勤怠"  # コメントあり → jinjer勤怠


def test_long_work_over_10h_no_longer_warns():
    """>10h の長時間労働注意喚起は廃止。差分なしの総労働行は生成しない。集計不整合は残す。"""
    # 純粋な >10h（差分なし）→ 警告なし（None）
    assert classify_total_work(_jinjer_row(**{
        JINJER_HEADERS["total_work"]: "10:30",
        JINJER_HEADERS["punch_in_1"]: "9:00",
    })) is None
    # 集計不整合（出勤打刻あり / 総労働0:00）は引き続き DANGER
    res = classify_total_work(_jinjer_row(**{
        JINJER_HEADERS["total_work"]: "0:00",
        JINJER_HEADERS["punch_in_1"]: "9:00",
    }))
    assert res is not None and res[0] == "DANGER"

    # 請求=jinjer（差分なし）で >10h の行は総労働行を生成しない
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00", "jinjer_出勤": "9:00", "出勤差分(分)": 0,
        "勤務表_退勤": "20:30", "jinjer_退勤": "20:30", "退勤差分(分)": 0,
        "勤務表_実働時間": "10:30",
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): _jinjer_row(**{
            JINJER_HEADERS["total_work"]: "10:30",
            JINJER_HEADERS["punch_in_1"]: "9:00",
            JINJER_HEADERS["punch_out_1"]: "20:30",
        })},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
    )
    assert rows == []  # >10h 単独では行を作らない


def test_compute_diffs_sets_triage_default():
    """コメント無し・小差分の出勤差異 → 自動採用(請求勤怠) が既定セットされる。"""
    from services.triage import TRIAGE_AUTO_KINTAI

    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾", "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:05", "jinjer_出勤": "9:00", "出勤差分(分)": 5,
        "勤務表_退勤": "18:00", "jinjer_退勤": "18:00", "退勤差分(分)": 0,
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): _jinjer_row()},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
    )
    punch = [r for r in rows if r.kind == DIFF_KIND_PUNCH_IN]
    assert punch
    assert punch[0].triage == TRIAGE_AUTO_KINTAI
    assert punch[0].judge_default == "請求勤怠"


def test_load_stamp_correction_reasons(tmp_path):
    """申請データCSVの「理由」を (従業員ID,日付ISO)→コメント に読み込む。理由空は除外。"""
    import csv as _csv
    from pathlib import Path

    p = tmp_path / "app.csv"
    with open(p, "w", encoding="cp932", newline="") as f:
        w = _csv.writer(f)
        w.writerow([
            "No.", "従業員ID", "従業員名", "所属グループ", "打刻グループ名", "日付",
            "申請種別", "申請前", "申請内容", "理由", "ステータス", "承認者コメント",
            "申請日時", "対応日時",
        ])
        w.writerow(["1", "2018012", "加藤 昌", "本部", "T2課", "2026年05月07日",
                    "打刻修正申請", "なし", "出退勤 15:30~18:30", "午前休", "申請中", "", "", ""])
        w.writerow(["2", "2018012", "加藤 昌", "本部", "T2課", "2026年05月16日",
                    "打刻修正申請", "なし", "出退勤 09:00~17:30", "研修", "承認済", "", "", ""])
        w.writerow(["3", "2018012", "加藤 昌", "本部", "T2課", "2026年05月18日",
                    "打刻修正申請", "なし", "出退勤", "", "申請中", "", "", ""])  # 理由空→除外

    logs: list[LogEntry] = []
    d = load_stamp_correction_reasons(Path(p), logs)
    assert d[("2018012", "2026-05-07")][0]["comment"] == "午前休（申請中）"
    assert d[("2018012", "2026-05-16")][0]["comment"] == "研修（承認済）"
    assert ("2018012", "2026-05-18") not in d
    assert format_stamp_comments(d[("2018012", "2026-05-07")]) == "午前休（申請中）"


def test_holiday_columns_transcribed_from_jinjer():
    """汎用データの 休日休暇名1 / 休日休暇名1：種別 が差異行へ転記される。"""
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00", "jinjer_出勤": "9:30", "出勤差分(分)": 30,
        "勤務表_退勤": "18:00", "jinjer_退勤": "18:00", "退勤差分(分)": 0,
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    jrow = _jinjer_row(**{"休日休暇名1": "有給休暇", "休日休暇名1：種別": "1"})
    extra_cols = resolve_jinjer_extra_columns(list(jrow.keys()))
    assert extra_cols.get("休日休暇名1") == "休日休暇名1"
    assert extra_cols.get("休日休暇名1：種別") == "休日休暇名1：種別"

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): jrow},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
        extra_cols,
    )
    punch_rows = [r for r in rows if r.kind == DIFF_KIND_PUNCH_IN]
    assert punch_rows
    assert punch_rows[0].holiday_name1 == "有給休暇"
    assert punch_rows[0].holiday_name1_type == "1"


def test_punch_comment_transcribed_from_jinjer():
    """汎用データの『打刻時コメント』(#96) が整形のうえ差異行へ転記される。"""
    kintai_df = pd.DataFrame([{
        "氏名": "上原 奏吾",
        "日付": date(2026, 4, 1),
        "勤務表_出勤": "9:00", "jinjer_出勤": "9:30", "出勤差分(分)": 30,
        "勤務表_退勤": "18:00", "jinjer_退勤": "18:00", "退勤差分(分)": 0,
        "_source_file": "x.xlsx",
    }])
    logs: list[LogEntry] = []
    jrow = _jinjer_row(**{"打刻時コメント": "出勤: KDX出社のため、10:00時差出勤 , 退勤:  , "})
    extra_cols = resolve_jinjer_extra_columns(list(jrow.keys()))
    assert extra_cols.get("打刻時コメント") == "打刻時コメント"

    rows = compute_diffs(
        kintai_df,
        {("2018057", "2026-04-01"): jrow},
        {"上原 奏吾": "2018057", "上原奏吾": "2018057"},
        logs,
        extra_cols,
    )
    punch_rows = [r for r in rows if r.kind == DIFF_KIND_PUNCH_IN]
    assert punch_rows
    assert punch_rows[0].punch_comment == "出勤: KDX出社のため、10:00時差出勤"


# =============================================================================
# 打刻コメントの定型語除外（打刻忘れ／打刻漏れ／打刻修正）
# =============================================================================

def test_strip_punch_noise_words_removes_only_keywords():
    # 語そのものだけなら空に
    assert strip_punch_noise_words("打刻忘れ") == ""
    assert strip_punch_noise_words("打刻漏れ") == ""
    assert strip_punch_noise_words("打刻修正") == ""
    # 文中の語だけ除去し、残りは保持
    assert strip_punch_noise_words("打刻忘れ 客先で終日対応") == "客先で終日対応"
    assert strip_punch_noise_words("客先直行のため打刻修正") == "客先直行のため"
    # 空・NaN
    assert strip_punch_noise_words("") == ""
    assert strip_punch_noise_words(None) == ""
    assert strip_punch_noise_words("nan") == ""


def test_clean_punch_comment_strips_noise_per_side():
    # 出勤側が定型語のみ → 出勤ラベルは落ち、退勤側の本文は残る
    out = clean_punch_comment("出勤: 打刻忘れ , 退勤: 私用で早退")
    assert out == "退勤: 私用で早退"
    # 退勤側が定型語のみ → 退勤ラベルは落ち、出勤側は残る
    out2 = clean_punch_comment("出勤: 客先直行 , 退勤: 打刻修正")
    assert out2 == "出勤: 客先直行"
    # 両方本文あり（定型語なし）はそのまま
    out3 = clean_punch_comment("出勤: 直行 , 退勤: 直帰")
    assert out3 == "出勤: 直行 / 退勤: 直帰"


def test_format_stamp_comments_strips_noise():
    # 除外はコメント本文に対して行う（打刻方法ラベル [打刻修正申請] 等のメタ情報は対象外）。
    items = [
        {"type": "出勤", "method": "PC", "comment": "打刻忘れ KDX出社"},
        {"type": "退勤", "method": "PC", "comment": "打刻修正"},
    ]
    out = format_stamp_comments(items)
    # 1件目: 本文「KDX出社」が残り、定型語は消える
    assert "KDX出社" in out
    assert "打刻忘れ" not in out
    # 2件目: 定型語のみ → その項目ごと消える（退勤の項目が出ない）
    assert "退勤" not in out
