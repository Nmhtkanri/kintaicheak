from services import invoice_mode as im


DTS_TEXT = """請求書
発注者 株式会社ＤＴＳ サプライヤ 株式会社エヌエム・ヒューマテック
スタッフ 岡崎, 修司 提出者 平良, 菜津子
提出日 2026-08-03
請求総額（税込み） 748,000
小計 680,000
消費税_10% 680,000 10.000 比率 68,000
小計 68,000
"""

IX_TEXT = """請求確定日時：2026/08/05 16:44
アイエックス・ナレッジ株式会社 御中
税抜合計額 1,110,000
消費税 111,000
総合計 1,221,000
"""

NTT_TEXT = """標準請求書
請求日 2026年8月4日 (火曜日)
支払金額 ¥1,601,875 JPY
明細の小計 ¥1,456,250 JPY
合計税額 ¥145,625 JPY
総計 ¥1,601,875 JPY
"""

IT_ONE_TEXT = """請求確定日時：2026/08/03 20:30
株式会社アイ・ティー・ワン 御中
請求金額 880,000円（税込）
10%対象 税抜金額 800,000円 消費税 80,000円
"""

FOCUS_TEXT = """請 求 書
発行日 : 2026/07/31
（請求先）株式会社フォーカスシステムズ 御中
10%対象 740,000
消費税(10%) 74,000
合計 814,000
"""

FIELDGLASS_TEXT = """Invoice
ID ERCSIN01294087 Buyer Ericsson
End Date 2026-07-31 Submitted By user
Worker Nara, Takahiro Submit Date 2026-07-31
Total Amount Due 975,100
Subtotal 975,100
"""


def test_parse_dts_fieldglass_japanese():
    got = im.parse_invoice_text(DTS_TEXT, "【DTS様】御請求書_202607.pdf")
    assert got["employee_name"] == "岡崎修司"
    assert got["partner"] == "株式会社DTS"
    assert got["issue_date"] == ""
    assert got["main_amount"] == 748000
    assert got["main_tax"] == 68000


def test_parse_combined_ix_invoice():
    got = im.parse_invoice_text(IX_TEXT, "【IXナレッジ様】御請求書_202607.pdf")
    assert got["issue_date"] == "2026-08-05"
    assert got["main_amount"] == 1221000
    assert got["main_tax"] == 111000


def test_parse_ntt_invoice():
    got = im.parse_invoice_text(NTT_TEXT, "【NTTデータ様】御請求書_202607.pdf")
    assert got["issue_date"] == "2026-08-04"
    assert got["main_amount"] == 1601875
    assert got["main_tax"] == 145625


def test_parse_it_one_invoice():
    got = im.parse_invoice_text(IT_ONE_TEXT, "【アイ・ティー・ワン様】御請求書_202607.pdf")
    assert got["issue_date"] == "2026-08-03"
    assert got["main_amount"] == 880000
    assert got["main_tax"] == 80000


def test_parse_focus_invoice():
    got = im.parse_invoice_text(FOCUS_TEXT, "【フォーカス様】御請求書_202607.pdf")
    assert got["partner"] == "株式会社フォーカスシステムズ"
    assert got["issue_date"] == "2026-07-31"
    assert got["main_amount"] == 814000
    assert got["main_tax"] == 74000


def test_parse_english_fieldglass_keeps_missing_tax_for_manual_review():
    got = im.parse_invoice_text(FIELDGLASS_TEXT, "invoice_ERCSIN01294087.pdf")
    assert got["employee_name"] == "NaraTakahiro"
    assert got["partner"] == "Ericsson"
    assert got["main_amount"] == 975100
    assert got["main_tax"] is None
    assert got["issue_date"] == ""


def test_unique_company_fallback_does_not_guess_when_multiple_people():
    document = {"employee_name": "", "text": "", "partner": "株式会社DTS",
                "source_file": r"Z:\NetMarks以外(常駐）\DTS（岡崎）\invoice.pdf"}
    unique = {
        "a": {"employee_no": "2016024", "employee_name": "岡崎修司",
              "partner": "DTS", "department": "", "source": "book"}
    }
    assert im._match_employee(document, unique)["employee_no"] == "2016024"

    duplicate = dict(unique)
    duplicate["b"] = {"employee_no": "2099999", "employee_name": "別人",
                       "partner": "DTS", "department": "", "source": "book"}
    assert im._match_employee(document, duplicate)["employee_no"] == ""
