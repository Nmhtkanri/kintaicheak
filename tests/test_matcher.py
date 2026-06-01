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
    assert normalize_name("ＡＢＣ") == "ABC"


def test_overnight_jinjer_row_matches_split_sap_rows():
    """jinjerの夜勤1行を、SAPの日跨ぎ2行と突合する"""
    jinjer = make_df([
        ("及川 航平", date(2026, 4, 3), time(16, 45), time(9, 30)),
    ], "jinjer")
    sheet = make_df([
        ("及川, 航平", date(2026, 4, 3), time(16, 45), time(0, 0)),
        ("及川, 航平", date(2026, 4, 4), time(0, 0), time(9, 30)),
    ], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert unsubmitted == []
    assert len(result) == 2
    assert result["判定"].tolist() == ["OK", "OK"]
    assert result.iloc[0]["jinjer_退勤時刻"] == time(0, 0)
    assert result.iloc[1]["jinjer_出勤時刻"] == time(0, 0)
    assert result.iloc[1]["jinjer_退勤時刻"] == time(9, 30)


def test_total_work_diff_is_included_in_judgment():
    jinjer = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(18, 0))], "jinjer")
    sheet = make_df([("山田太郎", date(2024, 1, 15), time(9, 0), time(17, 30))], "勤務表")

    result, unsubmitted = match(jinjer, sheet, threshold_minutes=10)

    assert unsubmitted == []
    assert result.iloc[0]["勤務表_総労働時間"] == "8:30"
    assert result.iloc[0]["jinjer_総労働時間"] == "9:00"
    assert result.iloc[0]["総労働差分(分)"] == 30
    assert result.iloc[0]["判定"] == "NG"


if __name__ == "__main__":
    test_ok()
    test_ng()
    test_caution_with_comment()
    test_missing_jinjer()
    test_unsubmitted()
    test_normalize_name()
    print("全テスト通過")
