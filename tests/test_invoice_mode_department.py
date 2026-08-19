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


def test_only_own_employees_are_in_the_master(tmp_path):
    """★回帰: 派遣の人がマスタに混ざると、氏名の入っていない請求書が拾えない。

    社員番号は 20YY 始まりが自社社員。IXナレッジ様の請求書には担当者名が
    書かれておらず契約先から補完するが、同じ契約先に派遣の八橋さん(5000002)が
    いたため「候補が複数」となり小島さん(2024044)が入らなかった。
    """
    book = Workbook()
    sheet = book.active
    sheet.title = "Other"
    sheet.append(["社員番号", "氏名", "No", "派遣/委任", "担当", "氏名", "部署", "作業名",
                  "契約先", "外注会社", "区分"])
    sheet.append(["2024044", "小島光晶", 1, "派遣", "", "小島光晶", "", "",
                  "IKI", "社員", "総受注金額"])
    sheet.append(["5000002", "八橋麻耶", 2, "派遣", "", "八橋麻耶", "", "",
                  "IKI", "派遣", "総受注金額"])
    path = tmp_path / "sales.xlsx"
    book.save(path)

    master = im.load_employee_master(path)
    assert im.normalize_name("小島光晶") in master
    assert im.normalize_name("八橋麻耶") not in master, "5始まりは自社社員ではない"

    document = {"employee_name": "", "text": "", "partner": "アイエックス・ナレッジ株式会社",
                "source_file": r"Z:\NetMarks以外(常駐）\IXナレッジ（小島・八橋）\請求書.pdf"}
    assert im._match_employee(document, master)["employee_name"] == "小島光晶"


def test_excluded_employees_are_dropped(tmp_path):
    """e-staffing / SAP Fieldglass 経由の人は別ルートでfreeeに入るので落とす。"""
    path = tmp_path / "excluded.csv"
    path.write_text("社員番号,氏名,理由\n2025029,奥山 昌苗,Estaffing\n"
                    "2018031,太田 裕一,SAP_Fieldglass\n", encoding="utf-8-sig")
    got = im.load_excluded_employees(path)
    assert got == {"2025029": "Estaffing", "2018031": "SAP_Fieldglass"}


def test_no_excluded_csv_means_no_exclusion(tmp_path):
    """リストが無いときに勝手に誰かを落とさない。"""
    assert im.load_excluded_employees(None) == {}
    assert im.load_excluded_employees(tmp_path / "none.csv") == {}
