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
