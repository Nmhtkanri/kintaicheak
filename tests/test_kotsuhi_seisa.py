"""承認前の交通費精査。

判定ロジック本体の細かいテストは Z:\\API連携\\tests\\test_kotsuhi_seisa.py にある。
ここでは「承認が進むたびに何度も回す」運用で効く部分だけを固める。
"""
import openpyxl
import pytest

from services.kotsuhi_seisa import (
    apply_diff,
    build_workdays,
    is_company_employee,
    read_previous_keys,
    row_key,
)


class _Sheet:
    def __init__(self, rows):
        self._rows = rows

    def iter_rows(self, min_row=1, values_only=False):
        return iter(self._rows)


class _Book:
    def __init__(self, summary, detail=None):
        self._s = summary
        self._d = detail
        self.sheetnames = ["サマリ"] + (["テレワーク明細"] if detail is not None else [])

    def __getitem__(self, name):
        return _Sheet(self._s if name == "サマリ" else self._d)


# ----------------------------------------------------------------------
# 繰り返し実行の差分
# ----------------------------------------------------------------------

def test_row_key_distinguishes_by_sheet_granularity():
    # 実費は日ごと、マスタ更新漏れは検知区分ごとに1件と数える
    assert row_key("実費突合", {"社員番号": "2024009", "利用日": "2026/7/1"}) \
        != row_key("実費突合", {"社員番号": "2024009", "利用日": "2026/7/2"})
    assert row_key("マスタ更新漏れ", {"社員番号": "2024009", "検知区分": "M2"}) \
        != row_key("マスタ更新漏れ", {"社員番号": "2024009", "検知区分": "M3"})
    assert row_key("定期代突合", {"社員番号": "2024009"}) == "定期代突合|2024009"


def test_apply_diff_marks_new_and_carried_over():
    prev = {"定期代突合": {"定期代突合|2025020"}}
    rows = [
        {"社員番号": "2025020", "区分": "要確認"},   # 前回もあった
        {"社員番号": "2026003", "区分": "要確認"},   # 今回から
    ]
    resolved = apply_diff("定期代突合", rows, prev)
    assert [r["前回比"] for r in rows] == ["継続", "新規"]
    assert resolved == 0


def test_apply_diff_counts_resolved():
    """申請者が直して要確認から外れたら「解消」として数える。"""
    prev = {"定期代突合": {"定期代突合|2025020", "定期代突合|2026003"}}
    rows = [
        {"社員番号": "2025020", "区分": "OK"},       # 直った
        {"社員番号": "2026003", "区分": "要確認"},   # まだ
    ]
    resolved = apply_diff("定期代突合", rows, prev)
    assert rows[0]["前回比"] == "解消"
    assert rows[1]["前回比"] == "継続"
    assert resolved == 1   # 行は残るが要確認から外れたので解消1件


def test_apply_diff_counts_rows_that_disappeared():
    """行ごと消えた（申請が取り下げられた等）ケースも解消として数える。"""
    prev = {"定期代突合": {"定期代突合|2025020"}}
    resolved = apply_diff("定期代突合", [], prev)
    assert resolved == 1


def test_apply_diff_is_noop_on_first_run():
    rows = [{"社員番号": "2025020", "区分": "要確認"}]
    assert apply_diff("定期代突合", rows, {}) == 0
    assert rows[0]["前回比"] == ""


def test_read_previous_keys_ignores_missing_file(tmp_path):
    assert read_previous_keys(tmp_path / "無い.xlsx") == {}


def test_read_previous_keys_picks_up_only_flagged_rows(tmp_path):
    p = tmp_path / "前回.xlsx"
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("定期代突合")
    ws.append(["社員番号", "氏名", "区分"])
    ws.append(["2025020", "戸松 乃天", "要確認"])
    ws.append(["2007001", "菅原 伸", "OK"])
    ws2 = wb.create_sheet("通勤費申請なし")
    # このシートの判定列は「確認要否」（「区分」は支給区分で別物）
    ws2.append(["社員番号", "氏名", "区分", "確認要否"])
    ws2.append(["2024009", "山口 太雅", "通勤費", "要確認"])
    ws2.append(["2007001", "菅原 伸", "通勤定期代", "OK"])
    wb.save(p)

    got = read_previous_keys(p)
    assert got["定期代突合"] == {"定期代突合|2025020"}
    assert got["通勤費申請なし"] == {"通勤費申請なし|2024009"}


def test_read_previous_keys_survives_a_broken_file(tmp_path):
    """前回ファイルが壊れていても実行は止めない（差分が出ないだけ）。"""
    p = tmp_path / "壊れ.xlsx"
    p.write_bytes(b"not an xlsx")
    assert read_previous_keys(p) == {}


# ----------------------------------------------------------------------
# サマリの数式（何度も回すので毎回ここを通る）
# ----------------------------------------------------------------------

def test_build_workdays_counts_from_detail_when_formula_has_no_cached_value():
    """テレワーク日数は COUNTIF の数式。書いたばかりのブックには結果が入っていない。"""
    wb = _Book(
        [("2008002", "伊藤 淳", 20, None, None, "7/1、7/2")],
        [("2008002", "伊藤 淳", "2026-07-%02d" % d, "テレワーク") for d in range(1, 17)],
    )
    got = build_workdays(wb)["2008002"]
    assert got["テレワーク日数"] == 16
    assert got["出社日数"] == 4       # 素直に読むと 20 になる箇所


def test_build_workdays_uses_cached_values_when_present():
    wb = _Book([("2008002", "伊藤 淳", 20, 16, 4, "7/1、7/2")])
    got = build_workdays(wb)["2008002"]
    assert (got["出勤日数"], got["テレワーク日数"], got["出社日数"]) == (20, 16, 4)


@pytest.mark.parametrize("emp,expected", [
    ("2024014", True),
    ("合計", False),      # 通勤費シート末尾の合計行
    ("5000001", False),   # 派遣
    ("3333008", False),   # 勤怠実績があっても給与計算対象外
])
def test_is_company_employee(emp, expected):
    assert is_company_employee(emp) is expected
