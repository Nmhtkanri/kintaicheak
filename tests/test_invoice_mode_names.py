"""請求書モードの従業員名（jinjer の姓・名 → 「姓 名」）のテスト。

freee の従業員は「姓 名」（半角スペース区切り）で登録されている。氏名の元に
している年度営業実績も補正マスタも「出澤信晃」と姓名が続けて書かれていて
こちらでは割れないため、姓と名を別に持っている jinjer を正とする。
"""

import json

from services import invoice_mode as im


def _employees(people):
    return [{"id": emp, "company": {"last_name": last, "first_name": first}}
            for emp, (last, first) in people.items()]


def _roster_cache(tmp_path, people):
    """jinjer 従業員一覧のキャッシュ（経理モードが作る raw/roster.json）を模す。"""
    path = tmp_path / "roster.json"
    path.write_text(json.dumps({"fetched_at": "x", "employees": _employees(people)},
                               ensure_ascii=False), encoding="utf-8")
    return path


class _StubClient:
    """jinjer の従業員一覧APIの代わり。何回叩いたかを数える。"""

    def __init__(self, people=None, broken=False):
        self.people = people or {}
        self.broken = broken
        self.calls = 0

    def get_employees(self, only_active=True):
        self.calls += 1
        if self.broken:
            raise RuntimeError("jinjer に繋がらない")
        return _employees(self.people)


def test_name_comes_from_jinjer_sei_and_mei(tmp_path):
    cache = _roster_cache(tmp_path, {"2017012": ("出澤", "信晃"),
                                     "2022002": ("MAHARJAN", "RAMITA")})
    client = _StubClient()
    got = im.load_employee_names(["2017012", "2022002"], cache_path=cache, client=client)
    assert got == {"2017012": "出澤 信晃", "2022002": "MAHARJAN RAMITA"}
    assert client.calls == 0, "キャッシュで足りるなら jinjer は叩かない"


def test_missing_employee_makes_it_refetch_once(tmp_path):
    """入社したての人がキャッシュに載っていないときだけ取り直す。"""
    cache = _roster_cache(tmp_path, {"2017012": ("出澤", "信晃")})
    client = _StubClient({"2017012": ("出澤", "信晃"), "2026001": ("新井", "太郎")})
    got = im.load_employee_names(["2017012", "2026001"], cache_path=cache, client=client)
    assert got == {"2017012": "出澤 信晃", "2026001": "新井 太郎"}
    assert client.calls == 1
    assert "2026001" in cache.read_text(encoding="utf-8"), "取り直した分は残す"


def test_unknown_employee_is_not_guessed(tmp_path):
    """jinjer にいない社員番号はキーごと入れない（勝手に姓名を割らない）。"""
    cache = _roster_cache(tmp_path, {"2017012": ("出澤", "信晃")})
    client = _StubClient({"2017012": ("出澤", "信晃")})
    got = im.load_employee_names(["2017012", "9999999"], cache_path=cache, client=client)
    assert got == {"2017012": "出澤 信晃"}


def test_name_lookup_is_not_fatal_when_jinjer_is_unavailable(tmp_path):
    """氏名が引けなくても請求書CSVは作れる（画面に警告を出して人が直す）。"""
    cache = _roster_cache(tmp_path, {"2017012": ("出澤", "信晃")})
    broken = _StubClient(broken=True)
    assert im.load_employee_names(["2017012", "9999999"], cache_path=cache,
                                  client=broken) == {"2017012": "出澤 信晃"}
    assert im.load_employee_names(["9999999"], cache_path=tmp_path / "no_such.json",
                                  client=broken) == {}
    assert im.load_employee_names([], client=broken) == {}
