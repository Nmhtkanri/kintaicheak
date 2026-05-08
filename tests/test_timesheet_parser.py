import os
import sys
from datetime import date, time

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.timesheet_parser import parse_timesheet_smart


def test_parse_sap_timesheet_excel_direct(tmp_path):
    path = tmp_path / "sap_timesheet.xlsx"
    pd.DataFrame([
        {
            "スタッフ": "上原, 奏吾",
            "食事休憩 1 開始": "12:00:00",
            "食事休憩 1 終了": "13:00:00",
            "終了時刻": "17:30:00",
            "出勤時刻": "09:00:00",
            "タイムシートステータス": "請求済み",
            "時間エントリ日": "2026-04-01 00:00:00",
        },
        {
            "スタッフ": "二宮, 正年",
            "終了時刻": "18:00:00",
            "出勤時刻": "10:00:00",
            "タイムシートステータス": "請求済み",
            "時間エントリ日": "2026-04-02 00:00:00",
        },
    ]).to_excel(path, index=False)

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert len(df) == 2
    assert df["氏名"].tolist() == ["上原, 奏吾", "二宮, 正年"]
    assert df.iloc[0]["日付"] == date(2026, 4, 1)
    assert df.iloc[0]["出勤時刻"] == time(9, 0)
    assert df.iloc[0]["退勤時刻"] == time(17, 30)
    assert df["コメント"].isna().all()


def test_parse_sap_timesheet_csv_direct(tmp_path):
    path = tmp_path / "sap_timesheet.csv"
    pd.DataFrame([
        {
            "スタッフ": "上原, 奏吾",
            "終了時刻": "17:30",
            "出勤時刻": "9:00",
            "時間エントリ日": "2026/4/1",
        },
        {
            "スタッフ": "及川, 航平",
            "終了時刻": "09:30",
            "出勤時刻": "00:00",
            "時間エントリ日": "2026/4/4",
        },
    ]).to_csv(path, index=False, encoding="utf-8-sig")

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert len(df) == 2
    assert df.iloc[0]["氏名"] == "上原, 奏吾"
    assert df.iloc[0]["日付"] == date(2026, 4, 1)
    assert df.iloc[0]["出勤時刻"] == time(9, 0)
    assert df.iloc[1]["出勤時刻"] == time(0, 0)
    assert df.iloc[1]["退勤時刻"] == time(9, 30)


def test_parse_estaffing_timesheet_csv_direct(tmp_path):
    path = tmp_path / "estaffing.csv"
    pd.DataFrame([
        {
            "e-staffing契約No": "C100887997-040",
            "スタッフ氏名": "寺山 枝美",
            "就業年月日": "2026/4/1",
            "日々勤怠状況": "承認済",
            "区分": "通常",
            "開始時刻": "9:00",
            "終了時刻": "24:00:00",
            "休憩時間": "1:30",
            "備考コメント": "",
        },
        {
            "e-staffing契約No": "C100887997-040",
            "スタッフ氏名": "寺山 枝美",
            "就業年月日": "2026/4/2",
            "日々勤怠状況": "承認済",
            "区分": "通常",
            "開始時刻": "9:00",
            "終了時刻": "22:30",
            "休憩時間": "2:00",
            "備考コメント": "テレワーク",
        },
    ]).to_csv(path, index=False, encoding="cp932")

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert len(df) == 2
    assert df.iloc[0]["氏名"] == "寺山 枝美"
    assert df.iloc[0]["日付"] == date(2026, 4, 1)
    assert df.iloc[0]["出勤時刻"] == time(9, 0)
    assert df.iloc[0]["退勤時刻"] == time(0, 0)
    assert pd.isna(df.iloc[0]["コメント"])
    assert df.iloc[1]["退勤時刻"] == time(22, 30)
    assert df.iloc[1]["コメント"] == "テレワーク"


def test_parse_estaffing_timesheet_text_direct(tmp_path):
    path = tmp_path / "estaffing.txt"
    lines = [
        "H\tC102350238-017\t2026/01/01\t2026/03/31\tAHIG\t東 昭夫",
        "D\t2026/03/02\t1\t9:00\t16:30\t1:00\t\t*[テレワーク、場所：自宅]\t0\t6:30\t6:30",
        "D\t2026/03/03\t\t\t\t\t\t\t\t\t",
        "D\t2026/03/04\t1\t16:45\t33:30\t1:00\t\t夜勤\t0\t15:45\t15:45",
    ]
    path.write_text("\n".join(lines), encoding="cp932")

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert len(df) == 2
    assert df.iloc[0]["氏名"] == "東 昭夫"
    assert df.iloc[0]["日付"] == date(2026, 3, 2)
    assert df.iloc[0]["出勤時刻"] == time(9, 0)
    assert df.iloc[0]["退勤時刻"] == time(16, 30)
    assert df.iloc[0]["コメント"] == "*[テレワーク、場所：自宅]"
    assert df.iloc[1]["出勤時刻"] == time(16, 45)
    assert df.iloc[1]["退勤時刻"] == time(9, 30)
