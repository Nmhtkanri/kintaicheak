import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
from datetime import date, time
from services.matcher import match, normalize_name


COLS = ["氏名", "日付", "出勤時刻", "退勤時刻", "コメント", "データソース"]

def make_df(records, source):
    rows = []
    for r in records:
        rows.append({
            "氏名": r[0],
            "日付": r[1],
            "出勤時刻": r[2],
            "退勤時刻": r[3],
            "コメント": r[4] if len(r) > 4 else None,
            "データソース": source,
        })
    if not rows:
        return pd.DataFrame(columns=COLS)
    return pd.DataFrame(rows, columns=COLS)


def test_ok():
    jinjer = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(18, 0))], "jinjer")
    sheet = make_df([("山田太郎", date(2024, 1, 15), time(9, 5), time(18, 0))], "勤務表")
    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)
    assert result.iloc[0]["判定"] == "OK"
    assert unsubmitted == []


def test_ng():
    jinjer = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(18, 0))], "jinjer")
    sheet = make_df([("山田太郎", date(2024, 1, 15), time(9, 20), time(18, 0))], "勤務表")
    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)
    assert result.iloc[0]["判定"] == "NG"
    assert unsubmitted == []


def test_caution_with_comment():
    jinjer = make_df([("佐藤花子", date(2024, 1, 16), time(10, 0), time(17, 0), "早退申請")], "jinjer")
    sheet = make_df([("佐藤花子", date(2024, 1, 16), time(10, 0), time(19, 0))], "勤務表")
    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)
    assert result.iloc[0]["判定"] == "要確認"
    assert unsubmitted == []


def test_missing_jinjer():
    """勤務表にはあるがjinjerにない → データ欠損"""
    jinjer = make_df([], "jinjer")
    sheet = make_df([("田中次郎", date(2024, 1, 15), time(9, 0), time(18, 0))], "勤務表")
    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)
    assert result.iloc[0]["判定"] == "データ欠損"
    assert unsubmitted == []


def test_unsubmitted():
    """jinjerにはあるが勤務表にない社員 → 未提出リストに入る"""
    jinjer = make_df([
        ("山田太郎", date(2024, 1, 15), time(9, 0), time(18, 0)),
        ("鈴木一郎", date(2024, 1, 15), time(9, 0), time(18, 0)),
        ("鈴木一郎", date(2024, 1, 16), time(9, 0), time(18, 0)),
    ], "jinjer")
    sheet = make_df([
        ("山田太郎", date(2024, 1, 15), time(9, 5), time(18, 0)),
    ], "勤務表")
    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)
    # 鈴木一郎は突合結果に含まれない
    assert "鈴木一郎" not in result["氏名"].values
    # 鈴木一郎は未提出リストに含まれる
    assert "鈴木一郎" in unsubmitted
    # 山田太郎は未提出リストに含まれない
    assert "山田太郎" not in unsubmitted
    # 突合結果は山田太郎の1件のみ
    assert len(result) == 1
    assert result.iloc[0]["判定"] == "OK"


def test_single_staff_code_sheet_name_matches_single_jinjer_employee():
    """画像解析で氏名がスタッフコードだけになった1人分の勤務表をjinjer氏名に寄せる"""
    jinjer = make_df([
        ("田村 栄和", date(2026, 4, 1), time(7, 0), time(17, 30)),
    ], "jinjer")
    sheet = make_df([
        ("TAM", date(2026, 4, 1), time(7, 0), time(17, 30)),
        ("TAM", date(2026, 4, 2), time(8, 0), time(17, 30)),
    ], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert unsubmitted == []
    assert len(result) == 2
    assert result["氏名"].tolist() == ["田村 栄和", "田村 栄和"]
    assert result.iloc[0]["判定"] == "OK"
    assert result.iloc[0]["jinjer_出勤時刻"] == time(7, 0)
    assert result.iloc[1]["判定"] == "データ欠損"


def test_staff_code_sheet_name_does_not_match_multiple_jinjer_employees():
    """複数人CSVではスタッフコードを勝手に1人へ結びつけない"""
    jinjer = make_df([
        ("田村 栄和", date(2026, 4, 1), time(7, 0), time(17, 30)),
        ("山田 太郎", date(2026, 4, 1), time(9, 0), time(18, 0)),
    ], "jinjer")
    sheet = make_df([
        ("TAM", date(2026, 4, 1), time(7, 0), time(17, 30)),
    ], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert set(unsubmitted) == {"田村 栄和", "山田 太郎"}
    assert result.iloc[0]["氏名"] == "TAM"
    assert result.iloc[0]["判定"] == "データ欠損"


def test_fuzzy_matches_single_kanji_ocr_misread():
    """OCRが氏名を1文字誤読しても、jinjer側に唯一の近似候補があれば突合する。

    実例: PDFの担当者名「矢野瑞穂」をOCRが「矢野瑞也」と誤読したケース。
    jinjer側の氏名（矢野 瑞穂）に寄せて突合・表示されること。
    """
    jinjer = make_df([
        ("矢野 瑞穂", date(2026, 5, 1), time(8, 50), time(19, 30)),
    ], "jinjer")
    sheet = make_df([
        ("矢野瑞也", date(2026, 5, 1), time(8, 50), time(19, 30)),
    ], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert unsubmitted == []
    assert len(result) == 1
    assert result.iloc[0]["氏名"] == "矢野 瑞穂"  # jinjer側の正しい氏名に寄る
    assert result.iloc[0]["判定"] == "OK"


def test_fuzzy_does_not_match_when_two_close_candidates():
    """近似候補がjinjer側に2人以上いるときは曖昧なので結びつけない。"""
    jinjer = make_df([
        ("山田太郎", date(2026, 5, 1), time(9, 0), time(18, 0)),
        ("山田次郎", date(2026, 5, 1), time(9, 0), time(18, 0)),
    ], "jinjer")
    sheet = make_df([
        ("山田三郎", date(2026, 5, 1), time(9, 0), time(18, 0)),  # 両者と編集距離1
    ], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert set(unsubmitted) == {"山田太郎", "山田次郎"}
    assert result.iloc[0]["氏名"] == "山田三郎"
    assert result.iloc[0]["判定"] == "データ欠損"


def test_fuzzy_does_not_match_short_surname():
    """4文字未満の短い姓は近似一致の対象外（田中/田口の誤マッチ防止）。"""
    jinjer = make_df([
        ("田口", date(2026, 5, 1), time(9, 0), time(18, 0)),
    ], "jinjer")
    sheet = make_df([
        ("田中", date(2026, 5, 1), time(9, 0), time(18, 0)),
    ], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert unsubmitted == ["田口"]
    assert result.iloc[0]["氏名"] == "田中"
    assert result.iloc[0]["判定"] == "データ欠損"


def test_unique_surname_sheet_name_matches_jinjer_employee():
    """PDFファイル名が姓だけでも、jinjer側で一意ならその社員に寄せる"""
    jinjer = make_df([
        ("奈良 隆宏", date(2026, 4, 1), time(9, 0), time(20, 30)),
        ("太田 琢也", date(2026, 4, 1), time(9, 0), time(21, 0)),
    ], "jinjer")
    sheet = make_df([
        ("奈良", date(2026, 4, 1), time(9, 0), time(20, 30)),
    ], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert "奈良 隆宏" not in unsubmitted
    assert "太田 琢也" in unsubmitted
    assert result.iloc[0]["氏名"] == "奈良 隆宏"
    assert result.iloc[0]["判定"] == "OK"


def test_normalize_name():
    assert normalize_name("山田　太郎") == "山田太郎"
    assert normalize_name(" 田中 次郎 ") == "田中次郎"
    # ローマ字は大文字小文字を畳む（ＡＢＣ→NFKC→ABC→casefold→abc）
    assert normalize_name("ＡＢＣ") == "abc"
    # 表記ゆれのあるローマ字氏名が同じキーに正規化される
    assert normalize_name("MAHARJAN RAMITA") == normalize_name("Maharjan, Ramita")


def test_match_romaji_name_is_case_insensitive():
    """ローマ字氏名は大文字小文字が違っても突合できる。"""
    jinjer = make_df([("MAHARJAN RAMITA", date(2026, 5, 1), time(9, 0), time(17, 30))], "jinjer")
    sheet = make_df([("Maharjan Ramita", date(2026, 5, 1), time(9, 0), time(17, 30))], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert unsubmitted == []
    assert len(result) == 1
    assert result.iloc[0]["判定"] == "OK"


def test_overnight_sap_split_rows_are_merged_to_shift_start_date():
    """SAPの日跨ぎ2行（深夜0時割り）を勤務開始日1行に結合してjinjerの夜勤1行と突合する。

    jinjerは夜勤を開始日の1行（退勤24時超表記）で記録するため、書き戻しが
    開始日の行へ行える形（1シフト=1行）で突合結果を出す。
    """
    jinjer = make_df([
        ("及川 航平", date(2026, 4, 3), time(16, 45), time(9, 30)),
    ], "jinjer")
    sheet = make_df([
        ("及川, 航平", date(2026, 4, 3), time(16, 45), time(0, 0)),
        ("及川, 航平", date(2026, 4, 4), time(0, 0), time(9, 30)),
    ], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert unsubmitted == []
    assert len(result) == 1
    row = result.iloc[0]
    assert row["判定"] == "OK"
    assert row["日付"] == date(2026, 4, 3)
    assert row["勤務表_出勤時刻"] == time(16, 45)
    assert row["勤務表_退勤時刻"] == time(9, 30)
    assert row["jinjer_退勤時刻"] == time(9, 30)


def test_overnight_merge_sums_net_work_minutes():
    """深夜0時割り2行の結合時、実働(総労働時間(分))は両行の合計になる。"""
    jinjer = make_df([
        ("及川 航平", date(2026, 4, 3), time(16, 45), time(9, 30)),
    ], "jinjer")
    sheet = make_df([
        ("及川, 航平", date(2026, 4, 3), time(16, 45), time(0, 0)),
        ("及川, 航平", date(2026, 4, 4), time(0, 0), time(9, 30)),
    ], "勤務表")
    sheet["総労働時間(分)"] = [390, 510]  # 6.5h + 8.5h

    result, _ = match(jinjer, sheet, threshold_minutes=10)

    assert len(result) == 1
    assert result.iloc[0]["勤務表_実働時間"] == "15:00"


def test_month_end_overnight_is_paired_with_tokki():
    """月末の夜勤（請求側が24:00打切り・翌月分未取得）は同日で突合し、特記を付ける。

    出勤は通常突合（差分が出る）。退勤・総労働は翌月確認のため差分を出さない。
    """
    jinjer = make_df([
        ("山口 太雅", date(2026, 6, 30), time(17, 15), time(9, 33)),
    ], "jinjer")
    sheet = make_df([
        ("山口, 太雅", date(2026, 6, 30), time(16, 45), time(0, 0)),
    ], "勤務表")

    result, _ = match(jinjer, sheet, threshold_minutes=10)

    assert len(result) == 1
    row = result.iloc[0]
    assert row["出勤差分(分)"] == 30
    assert pd.isna(row["退勤差分(分)"]) or row["退勤差分(分)"] is None
    assert pd.isna(row["総労働差分(分)"]) or row["総労働差分(分)"] is None
    assert "月末日跨ぎ" in str(row["特記"])
    assert "月末日跨ぎ" in str(row["詳細"])
    assert row["判定"] in ("NG", "要確認")  # 出勤30分差はそのまま検出される


def test_month_end_overnight_ok_start_becomes_needs_check():
    """月末夜勤で出勤が一致していても、退勤未確認のためOKにはせず要確認にする。"""
    jinjer = make_df([
        ("及川 航平", date(2026, 6, 30), time(16, 45), time(9, 33)),
    ], "jinjer")
    sheet = make_df([
        ("及川, 航平", date(2026, 6, 30), time(16, 45), time(0, 0)),
    ], "勤務表")

    result, _ = match(jinjer, sheet, threshold_minutes=10)

    assert len(result) == 1
    assert result.iloc[0]["判定"] == "要確認"
    assert "月末日跨ぎ" in str(result.iloc[0]["特記"])


def test_mid_month_24h_end_pairs_and_shows_real_diff():
    """月中で請求が24:00終わり・jinjerが翌日2:00退勤の場合は同日で突合し、実差120分を出す。"""
    jinjer = make_df([
        ("畑中 竜哉", date(2026, 6, 12), time(9, 0), time(2, 0)),
        ("畑中 竜哉", date(2026, 6, 13), time(9, 0), time(17, 30)),
    ], "jinjer")
    sheet = make_df([
        ("畑中, 竜哉", date(2026, 6, 12), time(9, 0), time(0, 0)),
        ("畑中, 竜哉", date(2026, 6, 13), time(9, 0), time(17, 30)),
    ], "勤務表")

    result, _ = match(jinjer, sheet, threshold_minutes=10)

    assert len(result) == 2
    row = result[result["日付"] == date(2026, 6, 12)].iloc[0]
    assert row["退勤差分(分)"] == 120
    assert str(row.get("特記") or "") == ""  # 月末ではないので特記なし＝実差として扱う
    assert row["判定"] == "NG"


def test_month_start_orphan_tail_matched_to_prev_month_end():
    """月初の0:00始まり孤児行（前月末勤務の後半）は、前月末日のjinjer行と退勤を突合する。

    対象日付が前月末日になるため、書き戻しは開始日の行へ24時超表記で行える。
    """
    jinjer = make_df([
        ("大堀 広智", date(2026, 5, 31), time(16, 45), time(9, 33)),
        ("大堀 広智", date(2026, 6, 1), None, None),
    ], "jinjer")
    sheet = make_df([
        ("大堀, 広智", date(2026, 6, 1), time(0, 0), time(9, 30)),
    ], "勤務表")

    result, _ = match(jinjer, sheet, threshold_minutes=10)

    paired = result[result["日付"] == date(2026, 5, 31)]
    assert len(paired) == 1
    row = paired.iloc[0]
    assert row["勤務表_退勤時刻"] == time(9, 30)
    assert row["退勤差分(分)"] == 3
    assert row["出勤差分(分)"] == 0  # 前半は前月分で確認済みのため差0
    assert "前月末夜勤の後半" in str(row["特記"])
    assert row["判定"] == "OK"  # 3分差は許容内。特記(突合済み)はOKのまま
    # 6/1 の勤務表のみ行は前月末日へ付け替えられて消えている
    sheet_only_61 = result[
        (result["日付"] == date(2026, 6, 1))
        & result["勤務表_出勤時刻"].notna()
    ]
    assert len(sheet_only_61) == 0


def test_month_start_orphan_tail_unmatched_gets_tokki():
    """前月末日のjinjer行が無い場合、月初の孤児行には「自動書戻し不可」の特記を付ける。"""
    jinjer = make_df([
        ("大堀 広智", date(2026, 6, 1), None, None),
    ], "jinjer")
    sheet = make_df([
        ("大堀, 広智", date(2026, 6, 1), time(0, 0), time(9, 30)),
    ], "勤務表")

    result, _ = match(jinjer, sheet, threshold_minutes=10)

    orphan = result[result["勤務表_出勤時刻"].notna()]
    assert len(orphan) == 1
    assert "自動書戻し不可" in str(orphan.iloc[0]["特記"])


def test_no_punch_night_worker_gets_single_merged_row():
    """jinjerに打刻が無い夜勤者でも、請求側は開始日1行（16:45〜翌9:30）に結合される。

    これにより書き戻しは開始日の行に 出勤16:45/退勤33:30（24時超表記）となり、
    翌日行へ 00:00〜09:30 が書かれる事故を防ぐ。
    """
    jinjer = make_df([
        ("大堀 広智", date(2026, 6, 2), None, None),
        ("大堀 広智", date(2026, 6, 3), None, None),
    ], "jinjer")
    sheet = make_df([
        ("大堀, 広智", date(2026, 6, 2), time(16, 45), time(0, 0)),
        ("大堀, 広智", date(2026, 6, 3), time(0, 0), time(9, 30)),
    ], "勤務表")

    result, _ = match(jinjer, sheet, threshold_minutes=10)

    merged_row = result[result["勤務表_出勤時刻"].notna()]
    assert len(merged_row) == 1
    row = merged_row.iloc[0]
    assert row["日付"] == date(2026, 6, 2)
    assert row["勤務表_出勤時刻"] == time(16, 45)
    assert row["勤務表_退勤時刻"] == time(9, 30)
    assert row["判定"] == "データ欠損"


def test_duplicate_sheet_rows_are_deduped():
    """請求勤怠側の完全重複行（Fieldglassレポートの複製行）は1行に畳まれる。"""
    jinjer = make_df([
        ("畑中 竜哉", date(2026, 6, 10), time(9, 0), time(17, 30)),
    ], "jinjer")
    sheet = make_df([
        ("畑中, 竜哉", date(2026, 6, 10), time(9, 0), time(17, 30)),
        ("畑中, 竜哉", date(2026, 6, 10), time(9, 0), time(17, 30)),
        ("畑中, 竜哉", date(2026, 6, 10), time(9, 0), time(17, 30)),
        ("畑中, 竜哉", date(2026, 6, 10), time(9, 0), time(17, 30)),
    ], "勤務表")

    result, _ = match(jinjer, sheet, threshold_minutes=10)

    assert len(result) == 1
    assert result.iloc[0]["判定"] == "OK"


def test_total_work_diff_is_included_in_judgment():
    jinjer = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(18, 0))], "jinjer")
    sheet = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(17, 30))], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert unsubmitted == []
    assert result.iloc[0]["勤務表_総労働時間"] == "8:30"
    assert result.iloc[0]["jinjer_総労働時間"] == "9:00"
    assert result.iloc[0]["総労働差分(分)"] == 30
    assert result.iloc[0]["判定"] == "NG"


def test_actual_work_minutes_from_timesheet():
    """請求勤怠ファイル記載の正味労働(分)が 勤務表_実働時間 として出力される。
    手順1の 勤務表_総労働時間 は従来どおり拘束時間(退勤−出勤)のまま据え置く。"""
    jinjer = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(18, 0))], "jinjer")
    sheet = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(18, 0))], "勤務表")
    sheet["総労働時間(分)"] = [420]  # 休憩1h控除後の正味 7:00

    result, _ = match(jinjer, sheet, threshold_minutes=10)
    assert result.iloc[0]["勤務表_実働時間"] == "7:00"
    assert result.iloc[0]["勤務表_総労働時間"] == "9:00"  # 拘束時間は不変


def test_actual_work_minutes_absent_fallback():
    """総労働時間(分) 列が無い勤務表でも例外なく動き、実働列は空になる（後方互換）。"""
    jinjer = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(18, 0))], "jinjer")
    sheet = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(18, 0))], "勤務表")

    result, _ = match(jinjer, sheet, threshold_minutes=10)
    assert "勤務表_実働時間" in result.columns
    assert result.iloc[0]["勤務表_実働時間"] == ""


if __name__ == "__main__":
    test_ok()
    test_ng()
    test_caution_with_comment()
    test_missing_jinjer()
    test_unsubmitted()
    test_normalize_name()
    test_actual_work_minutes_from_timesheet()
    test_actual_work_minutes_absent_fallback()
    print("全テスト通過")
