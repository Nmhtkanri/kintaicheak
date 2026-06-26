# -*- coding: utf-8 -*-
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.expense_check import (
    summarize, build_telework_workbook, _md,
    read_commute_csv, add_commute_sheet, COMMUTE_OUTPUT_COLUMNS,
)


def test_summarize_dedups_and_filters():
    att = [
        {"date": "2026-05-01", "is_absent": False, "attended_at": "x",
         "stamp_classifications": [{"name": "テレワーク（4時間以上）"}]},
        # 同一日の重複レコード → 1日に集約
        {"date": "2026-05-01", "is_absent": False, "attended_at": "x",
         "stamp_classifications": [{"name": "テレワーク（4時間以上）"}]},
        # 出社（テレワーク区分なし）
        {"date": "2026-05-02", "is_absent": False, "attended_at": "x", "stamp_classifications": []},
        # 欠勤 → 除外
        {"date": "2026-05-03", "is_absent": True, "attended_at": None},
        # 打刻なし → 除外
        {"date": "2026-05-04", "is_absent": False, "attended_at": None},
    ]
    work_days, telework = summarize(att)
    assert work_days == 2                      # 5/1, 5/2 のみ
    assert telework == [("2026-05-01", "テレワーク（4時間以上）")]


def test_summarize_multiple_telework_kinds_joined():
    att = [
        {"date": "2026-05-10", "is_absent": False, "attended_at": "x",
         "stamp_classifications": [{"name": "テレワーク（午前）"}, {"name": "テレワーク（午後）"}]},
    ]
    work_days, telework = summarize(att)
    assert work_days == 1
    assert telework[0][0] == "2026-05-10"
    assert "テレワーク（午前）" in telework[0][1] and "テレワーク（午後）" in telework[0][1]


def test_summarize_empty():
    assert summarize([]) == (0, [])


def test_md_formats():
    assert _md("2026-05-01") == "5/1"
    assert _md("2026/5/9") == "5/9"


def test_build_telework_workbook(tmp_path):
    from openpyxl import load_workbook
    out = tmp_path / "tw.xlsx"
    build_telework_workbook("2026-05", [
        {"id": "2020001", "name": "田中 一郎", "work_days": 20,
         "telework_days": [("2026-05-01", "テレワーク"), ("2026-05-08", "テレワーク")]},
        {"id": "2020002", "name": "上原 奏吾", "work_days": 18, "telework_days": []},
    ], out)
    assert out.exists()
    wb = load_workbook(out)
    assert set(wb.sheetnames) == {"サマリ", "テレワーク明細"}
    assert wb.sheetnames[0] == "サマリ"  # サマリを先頭に表示
    ws = wb["サマリ"]
    assert [c.value for c in ws[1]] == ["社員番号", "氏名", "出勤日数", "テレワーク日数", "出社日数", "テレワーク実施日"]
    # 田中: テレワーク実施日は 5/1、5/8 表記
    assert ws.cell(row=2, column=6).value == "5/1、5/8"
    # テレワーク日数は COUNTIF 数式
    assert str(ws.cell(row=2, column=4).value).startswith("=COUNTIF")
    # 明細は田中の2行
    wd = wb["テレワーク明細"]
    assert wd.max_row == 3  # ヘッダー + 2行


def test_read_commute_csv_maps_fields(tmp_path):
    import csv
    p = tmp_path / "commute.csv"
    headers = ["社員番号", "職場氏名(氏)", "職場氏名(名)", "非課税通勤費(通勤1)", "課税通勤費(通勤1)",
               "支給金額(通勤1)", "利用開始日(通勤1)", "出発(通勤1)", "到着(通勤1)", "経由1(通勤1)",
               "経由2(通勤1)", "利用交通機関(通勤1)", "経路(通勤1)", "片道距離(km)(通勤1)"]
    with open(p, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerow(["2007002", "横山", "弘樹", "18350", "0", "18,350", "", "笹塚", "豊洲", "市ヶ谷",
                    "", "公共交通機関", "笹塚 → 京王線 → 豊洲", "0.00"])
        w.writerow(["2008002", "伊藤", "淳", "0", "0", "1056", "", "新八柱", "潮見", "",
                    "", "公共交通機関", "新八柱 → JR → 潮見", ""])
    rows = read_commute_csv(p)
    assert len(rows) == 2
    r = rows[0]
    assert r["社員番号"] == "2007002"
    assert r["氏名"] == "横山 弘樹"          # 氏＋名を結合
    assert r["出発"] == "笹塚" and r["到着"] == "豊洲"
    assert r["経由1"] == "市ヶ谷" and r["経由2"] == ""
    assert r["通勤経路"].startswith("笹塚")
    assert r["支給金額"] == 18350            # カンマ除去して int 化
    assert r["非課税通勤費"] == 18350


def test_add_commute_sheet(tmp_path):
    from openpyxl import Workbook, load_workbook
    wb = Workbook()
    add_commute_sheet(wb, [
        {"社員番号": "2007002", "氏名": "横山 弘樹", "出発": "笹塚", "到着": "豊洲", "経由1": "市ヶ谷",
         "経由2": "", "利用交通機関": "公共交通機関", "通勤経路": "笹塚 → 豊洲", "支給金額": 18350,
         "非課税通勤費": 18350, "課税通勤費": 0, "利用開始日": "", "片道距離(km)": ""},
        {"社員番号": "2008002", "氏名": "伊藤 淳", "出発": "新八柱", "到着": "潮見", "経由1": "",
         "経由2": "", "利用交通機関": "公共交通機関", "通勤経路": "新八柱 → 潮見", "支給金額": 1056,
         "非課税通勤費": 0, "課税通勤費": 0, "利用開始日": "", "片道距離(km)": ""},
    ])
    out = tmp_path / "c.xlsx"
    wb.save(out)
    wb2 = load_workbook(out)
    assert "通勤費" in wb2.sheetnames
    ws = wb2["通勤費"]
    assert [c.value for c in ws[1]] == COMMUTE_OUTPUT_COLUMNS
    assert ws.cell(row=2, column=1).value == "2007002"
    assert ws.cell(row=2, column=COMMUTE_OUTPUT_COLUMNS.index("出発") + 1).value == "笹塚"
    # 合計行（支給金額の SUM 数式）
    total_row = 2 + 2
    assert str(ws.cell(row=total_row, column=COMMUTE_OUTPUT_COLUMNS.index("支給金額") + 1).value).startswith("=SUM")
