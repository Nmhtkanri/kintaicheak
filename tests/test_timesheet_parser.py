import os
import sys
from datetime import date, time

import pandas as pd
from openpyxl import Workbook
from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import services.timesheet_parser as timesheet_parser
from services.timesheet_parser import _load_file_for_claude, parse_timesheet_smart


def test_parse_itone_dispatch_timesheet_excel_direct(tmp_path):
    path = tmp_path / "中澤寿代さん(ITone)派遣労働者勤務報告書_202604.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "派遣労働者勤務報告書"
    ws["B12"] = "日付"
    ws["U6"] = "中澤　寿代"
    ws["B13"] = "2026/4/1"
    ws["X13"] = "09:45"
    ws["AC13"] = "19:33"
    ws["L13"] = "JANET更改業務：時間外会議"
    ws["B14"] = "2026/4/2"
    ws["X14"] = "09:45"
    ws["AC14"] = "18:21"
    wb.save(path)

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert len(df) == 2
    assert df["氏名"].tolist() == ["中澤　寿代", "中澤　寿代"]
    assert df.iloc[0]["日付"] == date(2026, 4, 1)
    assert df.iloc[0]["出勤時刻"] == time(9, 45)
    assert df.iloc[0]["退勤時刻"] == time(19, 33)
    assert df.iloc[0]["コメント"] == "JANET更改業務：時間外会議"


def test_parse_nmht_work_time_report_excel_direct(tmp_path):
    path = tmp_path / "勤務表（菅原孝）202509_202605.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "勤務時間報告書"
    ws["C2"] = "年"
    ws["E2"] = "月"
    ws["F2"] = "勤務時間報告書Ver3.3.3"
    ws["C3"] = 2026
    ws["E3"] = 5
    ws["C6"] = "氏名"
    ws["E6"] = "菅原　孝"
    ws["C11"] = "日"
    ws["E11"] = "勤務"
    ws["F11"] = "開始時間"
    ws["H11"] = "終了時間"
    ws["T11"] = "備考"
    ws["F12"] = "時"
    ws["G12"] = "分"
    ws["H12"] = "時"
    ws["I12"] = "分"
    ws["C13"] = 1
    ws["E13"] = "出勤"
    ws["F13"] = 9
    ws["G13"] = 0
    ws["H13"] = 19
    ws["I13"] = 45
    ws["T13"] = "テレワーク"
    ws["C14"] = 2
    ws["E14"] = "有休(全休)"
    ws["C15"] = 3
    ws["E15"] = "出勤"
    ws["F15"] = 6
    ws["G15"] = 30
    ws["H15"] = 24
    ws["I15"] = 0
    wb.save(path)

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert len(df) == 2
    assert df["氏名"].tolist() == ["菅原　孝", "菅原　孝"]
    assert df.iloc[0]["日付"] == date(2026, 5, 1)
    assert df.iloc[0]["出勤時刻"] == time(9, 0)
    assert df.iloc[0]["退勤時刻"] == time(19, 45)
    assert df.iloc[0]["コメント"] == "テレワーク"
    assert df.iloc[1]["日付"] == date(2026, 5, 3)
    assert df.iloc[1]["出勤時刻"] == time(6, 30)
    assert df.iloc[1]["退勤時刻"] == time(0, 0)


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


def test_estaffing_break_includes_night_break(tmp_path):
    """e-staffingの実休憩は 休憩時間(AD列) + 深夜休憩時間(AE列) の合算。

    夜勤者は深夜休憩が別列のため、AD列だけだと正味労働が過大になる
    （河端さんの90分・加藤英人さんの30分の偽差異の原因）。
    """
    path = tmp_path / "estaffing.csv"
    pd.DataFrame([
        {   # 夜勤: 20:00〜翌8:15 拘束735分 − (休憩60 + 深夜休憩30) = 645分
            "スタッフ氏名": "加藤 英人",
            "就業年月日": "2026/6/3",
            "開始時刻": "20:00",
            "終了時刻": "32:15:00",
            "休憩時間": "1:00",
            "深夜休憩時間": "0:30",
            "備考コメント": "",
        },
        {   # 深夜休憩のみ: 拘束735分 − 90分 = 645分
            "スタッフ氏名": "河端 桂大",
            "就業年月日": "2026/6/1",
            "開始時刻": "20:00",
            "終了時刻": "32:15:00",
            "休憩時間": "0:00",
            "深夜休憩時間": "1:30",
            "備考コメント": "",
        },
    ]).to_csv(path, index=False, encoding="cp932")

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert df.iloc[0]["総労働時間(分)"] == 645
    assert df.iloc[1]["総労働時間(分)"] == 645


def test_sap_placeholder_row_has_no_time_and_tokki(tmp_path):
    """Fieldglassの「00:00〜00:00・24h」プレースホルダ行は時刻なし＋特記付きで取り込む。

    実打刻ではないため、00:00打刻や総労働24:00として突合してはならない。
    """
    path = tmp_path / "sap_timesheet.csv"
    pd.DataFrame([
        {
            "スタッフ": "福家, 寛昭",
            "出勤時刻": "00:00",
            "終了時刻": "00:00",
            "エントリ日の労働時間 (ブレークダウンなし)": "24.000",
            "時間エントリ日": "2026-06-30",
        },
        {
            "スタッフ": "福家, 寛昭",
            "出勤時刻": "08:30",
            "終了時刻": "17:00",
            "エントリ日の労働時間 (ブレークダウンなし)": "7.500",
            "時間エントリ日": "2026-06-29",
        },
    ]).to_csv(path, index=False, encoding="utf-8-sig")

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert len(df) == 2
    placeholder = df[df["日付"] == date(2026, 6, 30)].iloc[0]
    assert placeholder["出勤時刻"] is None or pd.isna(placeholder["出勤時刻"])
    assert placeholder["退勤時刻"] is None or pd.isna(placeholder["退勤時刻"])
    assert placeholder["総労働時間(分)"] is None or pd.isna(placeholder["総労働時間(分)"])
    assert "Fieldglass時刻なし" in str(placeholder["特記"])
    normal = df[df["日付"] == date(2026, 6, 29)].iloc[0]
    assert normal["出勤時刻"] == time(8, 30)
    assert str(normal["特記"] or "") in ("", "nan")


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


def test_parse_fieldglass_pdf_direct(monkeypatch, tmp_path):
    path = tmp_path / "ts_uploadid_timesheet_ERCSTS01161606奈良.pdf"
    pdf_text = "\n".join([
        "Time Sheet",
        "ID ERCSTS01161606 Worker Nara, Takahiro(ERCSWK00144174)",
        "Period 2026-04-01 to 2026-04-30 Job Posting Support Engineer|JP|Job Stage",
        "Time in/time out",
        "Day 3-30 Mon 3-31 Tue 4-01 Wed 4-02 Thu 4-03 Fri 4-04 Sat 4-05 Sun Total",
        "Time In 09:00 09:00 09:00 00:00 00:00",
        "Meal Break 1 12:00 - 13:00 12:00 - 13:00 12:00 - 13:00",
        "Time Out 20:30 18:00 21:30 00:00 00:00",
        "Total 10h 30m 8h 0m 11h 30m 0h 0m 0h 0m 30h 0m",
    ])
    monkeypatch.setattr(
        timesheet_parser,
        "_pdf_to_text_or_bytes",
        lambda filepath: (pdf_text, None),
    )

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert df["氏名"].tolist() == ["奈良", "奈良", "奈良"]
    assert df["日付"].tolist() == [date(2026, 4, 1), date(2026, 4, 2), date(2026, 4, 3)]
    assert df.iloc[0]["出勤時刻"] == time(9, 0)
    assert df.iloc[0]["退勤時刻"] == time(20, 30)
    assert df.iloc[2]["退勤時刻"] == time(21, 30)


def test_parse_fieldglass_pdf_12h_ampm(monkeypatch, tmp_path):
    """Fieldglass PDF の 12時間制(AM/PM) を 24時間制に正しく変換する。

    実データ(MAHARJAN/奈良)の不具合の再現: "6:00 PM"(=18:00) の PM を落として
    6:00 と誤読し、出勤(9:00)より早い→夜勤跨ぎ+24h で 30:00 に化けていた。
    夜勤(太田)は Time In "8:45 PM"=20:45、Time Out 翌9:00 AM も確認する。
    """
    path = tmp_path / "timesheet_ERCSTS_ampm.pdf"
    pdf_text = "\n".join([
        "Time Sheet",
        "ID ERCSTS01182072 Worker Ramita, Maharjan(ERCSWK00144175)",
        "Period 2026-06-01 to 2026-06-30 Job Posting Support Engineer|JP|Job Stage",
        "Time in/time out",
        "Day 6-01 Mon 6-02 Tue 6-03 Wed 6-06 Sat 6-07 Sun Total",
        "Time In 9:00 AM 9:00 AM 9:00 AM 12:00 AM 12:00 AM",
        "Time Out 6:00 PM 8:00 PM 7:00 PM 12:00 AM 12:00 AM",
        "Total 8h 0m 10h 0m 9h 0m 0h 0m 0h 0m 27h 0m",
    ])
    monkeypatch.setattr(
        timesheet_parser,
        "_pdf_to_text_or_bytes",
        lambda filepath: (pdf_text, None),
    )

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert df.iloc[0]["出勤時刻"] == time(9, 0)
    assert df.iloc[0]["退勤時刻"] == time(18, 0)   # 6:00 PM → 18:00（30:00 に化けない）
    assert df.iloc[1]["退勤時刻"] == time(20, 0)   # 8:00 PM → 20:00
    assert df.iloc[2]["退勤時刻"] == time(19, 0)   # 7:00 PM → 19:00


def test_fieldglass_12h_to_24h_conversion():
    """AM/PM → 24時間制の境界値（12:00 AM=00:00, 12:00 PM=12:00）。"""
    f = timesheet_parser._fieldglass_12h_to_24h
    assert f(12, 0, "AM") == "00:00"   # 深夜0時
    assert f(12, 0, "PM") == "12:00"   # 正午
    assert f(6, 0, "PM") == "18:00"
    assert f(8, 45, "PM") == "20:45"
    assert f(9, 0, "AM") == "09:00"
    assert f(20, 30, None) == "20:30"  # AM/PM無し＝既に24時間制


def test_parse_fieldglass_pdf_katakana_filename_uses_romaji_worker_name(monkeypatch, tmp_path):
    """ファイル名がカタカナ通称（ラミタ）でも、PDF本文のローマ字氏名で突合できる。

    jinjer側はローマ字（MAHARJAN RAMITA）で登録されるため、カタカナ通称のままでは
    一致しない。PDF本文の "Worker Ramita, Maharjan" を採用して MAHARJAN RAMITA とする。
    """
    path = tmp_path / "timesheet_ERCSTS0ラミタさん.pdf"
    pdf_text = "\n".join([
        "Time Sheet",
        "ID ERCSTS01173544 Worker Ramita, Maharjan(ERCSWK00144175)",
        "Period 2026-05-01 to 2026-05-31 Job Posting Support Engineer|JP|Job Stage",
        "Time in/time out",
        "Day 5-01 Thu 5-02 Fri 5-03 Sat Total",
        "Time In 09:00 09:00 00:00",
        "Time Out 22:30 20:00 00:00",
        "Total 12h 30m 10h 0m 0h 0m 22h 30m",
    ])
    monkeypatch.setattr(
        timesheet_parser,
        "_pdf_to_text_or_bytes",
        lambda filepath: (pdf_text, None),
    )

    result = parse_timesheet_smart(str(path))
    df = result["df"]

    assert result["mode"] == "direct"
    assert df["氏名"].tolist() == ["MAHARJAN RAMITA", "MAHARJAN RAMITA"]
    assert df.iloc[0]["日付"] == date(2026, 5, 1)


def test_image_only_pdf_is_sent_as_image_for_ai_fallback(tmp_path):
    path = tmp_path / "scanned_timesheet.pdf"
    image = Image.new("RGB", (480, 640), "white")
    draw = ImageDraw.Draw(image)
    draw.text((40, 40), "タイムシート", fill="black")
    draw.text((40, 90), "氏名 小林 拓己", fill="black")
    image.save(path, "PDF")

    content, file_type, media_type = _load_file_for_claude(str(path))

    assert file_type == "image"
    assert media_type == "image/png"
    assert content.startswith(b"\x89PNG")


def test_parse_timesheet_smart_falls_back_when_legend_parser_returns_empty(monkeypatch, tmp_path):
    path = tmp_path / "timesheet.png"
    Image.new("RGB", (320, 240), "white").save(path)

    monkeypatch.setattr(
        timesheet_parser,
        "parse_with_legend_extraction",
        lambda file_content, file_type, media_type=None: {"mode": "direct", "data": []},
    )
    monkeypatch.setattr(
        timesheet_parser,
        "_parse_with_claude",
        lambda file_content, file_type, media_type=None: {
            "employee_name": "小林 拓己",
            "records": [
                {
                    "date": "2026-04-01",
                    "start_time": "09:00",
                    "end_time": "17:30",
                    "comment": None,
                }
            ],
        },
    )

    result = parse_timesheet_smart(str(path))

    assert result["mode"] == "direct"
    assert len(result["df"]) == 1
    assert result["df"].iloc[0]["氏名"] == "小林 拓己"
    assert result["df"].iloc[0]["出勤時刻"] == time(9, 0)
