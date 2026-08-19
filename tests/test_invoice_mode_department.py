import json

from openpyxl import Workbook

from services import invoice_mode as im
from services.keiri_api import PAYROLL_MENU_ID


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

def _custom_items_cache(tmp_path, rows):
    """jinjer custom-items のキャッシュを模した JSON を作る。

    customize_item[] の id が履歴行番号で、同じ id の要素群が1行になる
    （services/keiri_api.parse_payroll_custom_history と同じ形）。
    """
    data = {}
    for emp, history in rows.items():
        items = []
        for idx, (on_date, bumon) in enumerate(history, start=1):
            items.append({"id": str(idx), "title": "日付", "value": on_date})
            items.append({"id": str(idx), "title": "部門", "value": bumon})
        data[emp] = {"employee_id": emp,
                     "customize_menu": [{"id": PAYROLL_MENU_ID, "customize_item": items}]}
    path = tmp_path / "custom_items.json"
    path.write_text(json.dumps({"fetched_at": "x", "data": data}, ensure_ascii=False),
                    encoding="utf-8")
    return path


def test_department_comes_from_jinjer_custom_items(tmp_path):
    """部門は jinjer のカスタム項目（経理モードと同じ元データ）から引く。"""
    cache = _custom_items_cache(tmp_path, {
        "2024044": [("2020/4/1", "OT：その他")],
        "2007001": [("2020/4/1", "OT：UAL（地方）")],
    })
    got = im.load_departments(["2024044", "2007001"], "2026-07-31", cache_path=cache)
    assert got == {"2024044": "OT：その他", "2007001": "OT：UAL（地方）"}


def test_department_uses_the_value_effective_at_month_end(tmp_path):
    """月中異動は対象月末時点の部門になる（履歴の時点解決）。"""
    cache = _custom_items_cache(tmp_path, {
        "2024044": [("2020/4/1", "OT：UAL（首都圏）"), ("2026/7/16", "OT：その他")],
    })
    assert im.load_departments(["2024044"], "2026-07-31", cache_path=cache) ==         {"2024044": "OT：その他"}
    assert im.load_departments(["2024044"], "2026-06-30", cache_path=cache) ==         {"2024044": "OT：UAL（首都圏）"}


def test_department_lookup_is_not_fatal_when_jinjer_is_unavailable(tmp_path):
    """部門が引けなくても請求書CSVは作れる（画面で赤くなり人が気づく）。"""
    assert im.load_departments(["9999999"], "2026-07-31",
                               cache_path=tmp_path / "no_such_cache.json") == {}
    assert im.load_departments([], "2026-07-31") == {}
