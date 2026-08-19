import csv
import io
import json
from pathlib import Path

import pytest

from services import invoice_mode as im


STANDARD_TEXT = """請求書
アプリケーションアシスト株式会社 御中
請求日: 2026-07-31
請求書番号： INV-260731A&A
作業担当者：新井 結貴 1.00 × 650,000 650,000
小計 644,925
消費税 64,493
立替金（交通費等） 1,413
合計金額 ¥710,831
内 10％対象(税抜) ¥644,925
10％消費税 ¥64,493
入金期日： 2026-08-31
"""


ESTAFF_TEXT = """請求年月日 2026年07月31日 請求書コード BHI20260731IDE
エス・アンド・アイ株式会社 御中
御請求金額総計 ¥891,000 - お支払 期 日 2026年08月31日
請求小計 810,000 特別調整額１ 0 立替金額計 0
消費税額 0
交通費相当額小計 0
消費税額 81,000
請求合計 891,000
請求書明細
出澤 信晃
"""


COMMUTE_TEXT = """請求書
エス・アンド・アイ株式会社 御中
請求日: 2026-07-31
作業担当者：出澤 信晃
立替金（交通費等） 1.00 × 5,904 5,904
合計金額 ¥5,904
内 10％対象(税抜) ¥5,368
10％消費税 ¥536
入金期日： 2026-08-31
"""


def test_parse_standard_invoice_with_embedded_commute():
    got = im.parse_invoice_text(STANDARD_TEXT, "【A&A様】御請求書_202607_修正版.pdf")
    assert got["kind"] == "main"
    assert got["employee_name"] == "新井結貴"
    assert got["partner"] == "アプリケーションアシスト株式会社"
    assert got["issue_date"] == "2026-07-31"
    assert got["due_date"] == "2026-08-31"
    assert got["main_amount"] == 709418
    assert got["main_tax"] == 64493
    assert got["commute_amount"] == 1413
    assert got["commute_tax"] == 128


def test_parse_estaff_invoice_uses_positive_tax_line():
    got = im.parse_invoice_text(ESTAFF_TEXT, "【出澤信晃】御請求書_202607.pdf")
    assert got["employee_name"] == "出澤信晃"
    assert got["issue_date"] == "2026-07-31"
    assert got["due_date"] == "2026-08-31"
    assert got["main_amount"] == 891000
    assert got["main_tax"] == 81000
    assert got["commute_amount"] == 0


def test_parse_separate_commute_invoice():
    got = im.parse_invoice_text(COMMUTE_TEXT, "【出澤信晃】立替金御請求書_202607.pdf")
    assert got["kind"] == "commute"
    assert got["commute_amount"] == 5904
    assert got["commute_tax"] == 536


def test_find_invoice_files_prefers_revised_and_ignores_reports(tmp_path):
    month_dir = tmp_path / "会社（氏名）" / "01,提出データ" / "FY2026" / "2026年7月"
    month_dir.mkdir(parents=True)
    original = month_dir / "【氏名】御請求書_202607.pdf"
    revised = month_dir / "【氏名】御請求書_202607_修正版.pdf"
    report = month_dir / "【氏名】作業報告書_202607.pdf"
    original.write_bytes(b"old")
    revised.write_bytes(b"new")
    report.write_bytes(b"report")

    got = im.find_invoice_files("2026-07", [tmp_path / "会社（氏名）"])
    assert got["selected"] == [str(revised)]
    assert got["ignored"][0]["file"] == str(original)
    assert report.name not in " ".join(got["selected"])


def test_build_preview_groups_main_and_commute(monkeypatch, tmp_path):
    month_dir = tmp_path / "会社" / "01,提出データ" / "FY2026" / "2026年7月"
    month_dir.mkdir(parents=True)
    main_path = month_dir / "【出澤信晃】御請求書_202607.pdf"
    commute_path = month_dir / "【出澤信晃】立替金御請求書_202607.pdf"
    main_path.write_bytes(b"main")
    commute_path.write_bytes(b"commute")

    def fake_parse(path):
        text = COMMUTE_TEXT if "立替" in str(path) else ESTAFF_TEXT
        result = im.parse_invoice_text(text, Path(path).name)
        result["source_file"] = str(path)
        result["revised"] = False
        return result

    monkeypatch.setattr(im, "parse_invoice_pdf", fake_parse)
    monkeypatch.setattr(im, "load_employee_master", lambda *_: {
        im.normalize_name("出澤信晃"): {
            "employee_no": "2017012", "employee_name": "出澤信晃",
            "partner": "S＆I", "department": "SI：その他", "source": "test",
        }
    })
    got = im.build_preview("2026-07", roots=[tmp_path / "会社"])
    assert len(got["rows"]) == 2
    main, commute = got["rows"]
    assert main["管理番号"] == "2017012"
    # 売上高は外税＝金額は税抜、税額は別欄（2026-08 谷津さん指定）
    assert main["税計算区分"] == "外税"
    assert main["金額"] == 810000, "891,000(税込) − 81,000 = 810,000(税抜)"
    assert main["税額"] == 81000
    # 交通費だけ内税のまま
    assert commute["勘定科目"] == "売上高（交通費）"
    assert commute["税計算区分"] == "内税"
    assert commute["金額"] == 5904
    assert commute["税額"] == 536
    assert not main["_errors"]
    assert not commute["_errors"]


def _valid_rows():
    return [
        {
            "収支区分": "収入", "管理番号": "2017012", "発生日": "2026-07-31",
            "支払期日": "2026-08-31", "取引先": "エス・アンド・アイ株式会社",
            "勘定科目": "売上高", "税区分": "課税売上10%", "金額": 891000,
            "税計算区分": "内税", "税額": 81000, "備考": "総合計請求書：出澤信晃",
            "品目": "", "部門": "SI：その他", "メモタグ（複数指定可、カンマ区切り）": "",
            "従業員": "出澤信晃", "_row_type": "main",
        },
        {
            "収支区分": "", "管理番号": "", "発生日": "", "支払期日": "",
            "取引先": "", "勘定科目": "売上高（交通費）", "税区分": "課税売上10%",
            "金額": 5904, "税計算区分": "内税", "税額": 536,
            "備考": "総合計請求書：出澤信晃", "品目": "", "部門": "SI：その他",
            "メモタグ（複数指定可、カンマ区切り）": "", "従業員": "出澤信晃",
            "_row_type": "commute",
        },
    ]


def test_export_csv_matches_freee_columns_bom_crlf_and_writes_log(tmp_path):
    got = im.export_csv("2026-07", _valid_rows(), tmp_path,
                        log_context={"selected_files": ["invoice.pdf"]})
    raw = Path(got["csv_path"]).read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")
    assert b"\r\n" in raw
    rows = list(csv.DictReader(io.StringIO(raw.decode("utf-8-sig"))))
    assert list(rows[0]) == im.CSV_COLUMNS
    assert rows[0]["発生日"] == "2026/07/31"
    assert rows[1]["収支区分"] == ""
    assert rows[1]["取引先"] == ""
    log = json.loads(Path(got["log_path"]).read_text(encoding="utf-8"))
    assert log["row_count"] == 2
    assert log["selected_files"] == ["invoice.pdf"]


def test_export_stops_when_required_fields_are_missing(tmp_path):
    rows = _valid_rows()
    rows[0]["支払期日"] = ""
    with pytest.raises(im.InvoiceModeError, match="支払期日が未入力"):
        im.export_csv("2026-07", rows, tmp_path)
