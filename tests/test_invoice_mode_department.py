from openpyxl import Workbook

from services import invoice_mode as im


def test_sales_book_contract_type_maps_to_freee_department(tmp_path):
    book = Workbook()
    sheet = book.active
    sheet.title = "Other"
    sheet.append(["社員番号", "氏名", "No", "派遣/委任", "担当", "氏名", "部署", "作業名",
                  "契約先", "外注会社", "区分"])
    sheet.append(["2016024", "岡崎修司", 1, "委任", "", "岡崎修司", "", "",
                  "DTS", "社員", "総受注金額"])
    sheet.append(["2025015", "西村勇祐", 2, "派遣", "", "西村勇祐", "", "",
                  "NTTデータ", "社員", "総受注金額"])
    path = tmp_path / "sales.xlsx"
    book.save(path)

    master = im.load_employee_master(path)
    assert master[im.normalize_name("岡崎修司")]["department"] == "他社向け委任契約"
    assert master[im.normalize_name("西村勇祐")]["department"] == "他社向け派遣"
