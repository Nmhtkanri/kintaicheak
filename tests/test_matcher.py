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


def test_normalize_name():
    assert normalize_name("山田　太郎") == "山田太郎"
    assert normalize_name(" 田中 次郎 ") == "田中次郎"
    assert normalize_name("ＡＢＣ") == "ABC"


if __name__ == "__main__":
    test_ok()
    test_ng()
    test_caution_with_comment()
    test_missing_jinjer()
    test_unsubmitted()
    test_normalize_name()
    print("全テスト通過")
