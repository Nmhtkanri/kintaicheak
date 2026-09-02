# -*- coding: utf-8 -*-
"""健康診断申込: jinjer レスポンスの読み取り（社員一覧・選択肢マスタ・前年度の健診内容）。"""

from services.health_apply import jinjer_source as J


def emp(emp_id, last, first, email, enr_id="0", enr_name="在籍", retire=""):
    return {"id": emp_id, "company": {"last_name": last, "first_name": first, "email": email,
                                      "enrollment_classification": {"id": enr_id, "name": enr_name},
                                      "retirement_date": retire}}


def history_menu(*records):
    """records: (レコードID, {項目名: 値}) → menu 15 の customize_item（レコードIDごとに項目が並ぶ形）。"""
    items = []
    for rec_id, fields in records:
        for title, value in fields.items():
            items.append({"id": rec_id, "title": title, "value": value})
    return {"id": 15, "name": "健康診断履歴", "customize_item": items}


def current_menu(org="", date="", options=None):
    return {"id": 14, "name": "健康診断", "customize_item": [
        {"id": 1, "title": "検診機関", "value": org},
        {"id": 2, "title": "受診日", "value": date},
        {"id": 3, "title": "希望したオプション", "value": options},
    ]}


def person(*menus):
    return {"employee_id": "2099001", "customize_menu": list(menus)}


OPTION_NAMES = {"15": {"1": "基本健診", "3": "1日人間ドック・胃カメラ", "5": "婦人病検査"},
                "14": {"1": "基本健診", "5": "婦人病検査"}}


# --- 社員一覧 -----------------------------------------------------------------

def test_profiles_from_employees_reads_company_fields():
    profiles = J.profiles_from_employees([
        emp("2099001", "試験", "太郎", "t.shiken@nmht.co.jp"),
        emp("2099002", "退職", "花子", "", "1", "退職", "2026-03-31"),
        {"employee_id": "2099003", "company": {"last_name": "片方"}},
        "junk", {"company": {}},
    ])
    assert profiles["2099001"].as_dict() == {"employee_id": "2099001", "name": "試験 太郎", "email": "t.shiken@nmht.co.jp",
                                             "enrollment_id": "0", "enrollment_name": "在籍", "retirement_date": ""}
    assert profiles["2099002"].enrollment_id == "1" and profiles["2099002"].retirement_date == "2026-03-31"
    assert profiles["2099003"].name == "片方" and profiles["2099003"].email == ""
    assert set(profiles) == {"2099001", "2099002", "2099003"}


def test_fetch_profiles_asks_for_everyone_and_filters():
    class Client:
        def __init__(self):
            self.calls = []

        def get_employees(self, only_active=True):
            self.calls.append(only_active)
            return [emp("2099001", "a", "b", "a@nmht.co.jp"), emp("2099002", "c", "d", "c@nmht.co.jp")]

    client = Client()
    got = J.fetch_profiles(client, ["2099002", "2099009"])
    assert client.calls == [False]
    assert list(got) == ["2099002"]


# --- 選択肢マスタ ---------------------------------------------------------------

def test_option_names_from_menus():
    menus = [
        {"id": 15, "name": "健康診断履歴", "customize_item": [
            {"id": 4, "name": "健診機関"},
            {"id": 5, "name": "健診内容", "options": [{"id": 1, "name": "基本健診 "}, {"id": 5, "name": "婦人病検査"}]},
        ]},
        {"id": 14, "name": "健康診断", "customize_item": [{"id": 3, "name": "希望したオプション", "options": [{"id": 5, "name": "婦人病検査"}]}]},
        "junk",
    ]
    assert J.option_names_from_menus(menus) == {"15": {"1": "基本健診", "5": "婦人病検査"}, "14": {"5": "婦人病検査"}}


# --- 前年度 -------------------------------------------------------------------

def test_history_record_of_previous_year_is_preferred_over_current():
    p = person(
        history_menu(("101", {"年度": "2025", "健診機関": "旧院", "健診内容": ["1"], "受診日": "2025/7/1"}),
                     ("102", {"年度": "2026", "健診機関": "医療法人社団 同友会 春日クリニック", "健診内容": ["1", "5"], "受診日": "2026/7/3"})),
        current_menu("現在値の機関", "2026-08-01", ["3"]),
    )
    prev = J.previous_from_custom_items(p, 2026, OPTION_NAMES)
    assert prev.source == "履歴" and prev.year == "2026"
    assert prev.institution_text == "医療法人社団 同友会 春日クリニック"
    assert prev.content_labels == ["基本健診", "婦人病検査"]
    assert prev.exam_date == "2026-07-03"
    assert prev.notes == []


def test_multiple_records_in_previous_year_pick_latest_exam_date_with_note():
    p = person(history_menu(
        ("1", {"年度": "2026", "健診機関": "A", "健診内容": "1", "受診日": "2026/9/1"}),
        ("2", {"年度": "2026", "健診機関": "B", "健診内容": "3", "受診日": "2026/6/1"}),
    ))
    prev = J.previous_from_custom_items(p, 2026, OPTION_NAMES)
    assert prev.institution_text == "A"
    assert prev.content_labels == ["基本健診"]
    assert any("2 件" in n for n in prev.notes)


def test_falls_back_to_current_menu_when_history_lacks_previous_year():
    p = person(history_menu(("1", {"年度": "2024", "健診機関": "古い", "健診内容": "1"})),
               current_menu("医療法人社団 同友会 春日クリニック\n", "2026/7/3", "5"))
    prev = J.previous_from_custom_items(p, 2026, OPTION_NAMES)
    assert prev.source == "現在値"
    assert prev.institution_text == "医療法人社団 同友会 春日クリニック"
    assert prev.content_labels == ["婦人病検査"]
    assert prev.exam_date == "2026-07-03"
    assert any("現在値" in n for n in prev.notes)


def test_none_when_no_history_and_empty_current():
    assert J.previous_from_custom_items(person(current_menu()), 2026, OPTION_NAMES).source == "なし"
    assert J.previous_from_custom_items(person(), 2026, OPTION_NAMES).source == "なし"
    assert J.previous_from_custom_items(None, 2026, OPTION_NAMES).source == "なし"


def test_unknown_option_id_is_kept_with_note():
    p = person(history_menu(("1", {"年度": "2026", "健診機関": "A", "健診内容": ["1", "9"]})))
    prev = J.previous_from_custom_items(p, 2026, OPTION_NAMES)
    assert prev.content_labels == ["基本健診", "9"]
    assert any("選択肢ID 9" in n for n in prev.notes)


def test_value_ids_accepts_list_string_and_separators():
    assert J._value_ids(["1", 5, None]) == ["1", "5"]
    assert J._value_ids("1;5") == ["1", "5"]
    assert J._value_ids("1, 3 5") == ["1", "3", "5"]
    assert J._value_ids("") == [] and J._value_ids(None) == []


def test_menu_as_dict_is_tolerated():
    p = {"employee_id": "2099001", "customize_menu": history_menu(("1", {"年度": "2026", "健診機関": "A", "健診内容": "1"}))}
    assert J.previous_from_custom_items(p, 2026, OPTION_NAMES).source == "履歴"


def test_fetch_previous_calls_custom_items_once_for_all_ids():
    class Client:
        def __init__(self):
            self.calls = []

        def get_custom_items(self, ids):
            self.calls.append(list(ids))
            return {"2099001": person(history_menu(("1", {"年度": "2026", "健診機関": "A", "健診内容": "1"})))}

    client = Client()
    got = J.fetch_previous(client, ["2099001", "2099002"], 2026, OPTION_NAMES)
    assert client.calls == [["2099001", "2099002"]]
    assert got["2099001"].source == "履歴" and got["2099002"].source == "なし"
    assert J.fetch_previous(client, [], 2026, OPTION_NAMES) == {}
    assert client.calls == [["2099001", "2099002"]]


def test_keiri_client_get_custom_menus_uses_master_path(monkeypatch):
    from services.keiri_api import KeiriJinjerClient

    client = KeiriJinjerClient.__new__(KeiriJinjerClient)   # 認証情報なしで作る
    calls = []

    def fake_get(path, params):
        calls.append((path, params))
        return [{"id": 15}]

    monkeypatch.setattr(client, "_get", fake_get)
    assert client.get_custom_menus() == [{"id": 15}]
    assert calls == [("/v1/master/custom-menus", {})]
    monkeypatch.setattr(client, "_get", lambda path, params: {"unexpected": 1})
    assert client.get_custom_menus() == []
