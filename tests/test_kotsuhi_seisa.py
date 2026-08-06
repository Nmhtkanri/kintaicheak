"""承認前の交通費精査。

判定ロジック本体の細かいテストは Z:\\API連携\\tests\\test_kotsuhi_seisa.py にある。
ここでは「承認が進むたびに何度も回す」運用で効く部分だけを固める。
"""
import openpyxl
import pytest

from services.kotsuhi_seisa import (
    apply_diff,
    build_limit_over_rows,
    build_workdays,
    is_company_employee,
    load_limit_exempt_members,
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


# ----------------------------------------------------------------------
# 通勤費の上限（月3万円）超過の検出
# ----------------------------------------------------------------------

LIMIT_IDX = {"交通機関": 0, "ステータス": 1, "社員番号": 2, "申請者": 3,
             "所属グループ": 4, "申請書No.": 5, "小計": 6}


def _row(kind, emp, name, amount, no="1", status="承認完了", group="UAL 平和島"):
    return [kind, status, emp, name, group, no, amount]


def test_limit_over_flags_only_people_missing_from_the_exempt_list():
    """許可者はOK、リストに無い人だけ要確認（＝上限の切り忘れ検知）。"""
    details = [
        _row("通勤定期代", "2014013", "柴田 和浩", "35090"),   # 許可なしで超過
        _row("通勤定期代", "2025029", "奥山 昌苗", "34350"),   # 許可あり
        _row("通勤定期代", "2026006", "大村 賢治", "30000"),   # 上限ちょうど＝対象外
    ]
    rows = build_limit_over_rows(details, LIMIT_IDX, {"2025029": "個別許可"}, limit=30000)

    by = {r["社員番号"]: r for r in rows}
    assert set(by) == {"2014013", "2025029"}        # 30,000ちょうどは挙がらない
    assert by["2014013"]["区分"] == "要確認"
    assert by["2014013"]["超過額"] == 5090
    assert by["2014013"]["上限免除"] == ""
    assert by["2025029"]["区分"] == "OK"
    assert by["2025029"]["上限免除"] == "○"
    assert "個別許可" in by["2025029"]["説明"]


def test_limit_over_treats_travel_expense_members_as_no_limit():
    """移動交通費（立替精算）対象者は通勤系で申請されていても上限が掛からない。

    2026-08 の山田大海さん（実費51,067円）がこれで要確認に挙がった。金額からは
    判別できないので、対象者リストを見に行かないと毎月ここで引っかかる。
    """
    details = [
        _row("通勤交通費（実費）", "2025033", "山田 大海", "51067"),
        _row("通勤交通費（実費）", "2018001", "有田 功太郎", "31832"),
    ]
    rows = build_limit_over_rows(details, LIMIT_IDX, {}, limit=30000,
                                 travel_members={"2025033"})

    by = {r["社員番号"]: r for r in rows}
    assert by["2025033"]["区分"] == "OK"
    assert by["2025033"]["上限免除"] == "○"
    assert "移動交通費" in by["2025033"]["説明"]
    assert by["2018001"]["区分"] == "要確認"       # リストに無い人はこれまでどおり


def test_limit_over_sums_split_applications_and_ignores_travel_expense():
    """定期代の分割申請は合算し、移動交通費（上限なし）は混ぜない。"""
    details = [
        _row("通勤定期代", "2021020", "稲場 直哉", "16000", no="1"),
        _row("通勤交通費（実費）", "2021020", "稲場 直哉", "16070", no="2"),
        # 移動交通費は何円あっても上限判定に入れない
        _row("交通費（電車・バス）", "2021020", "稲場 直哉", "99999", no="3"),
        _row("交通費（電車・バス）", "2019048", "阿部 涼平", "80000", no="4"),
    ]
    rows = build_limit_over_rows(details, LIMIT_IDX, {}, limit=30000)

    assert [r["社員番号"] for r in rows] == ["2021020"]
    got = rows[0]
    assert (got["通勤費合計"], got["うち定期代"], got["うち実費"]) == (32070, 16000, 16070)
    assert got["申請書No."] == "1, 2"


def test_limit_over_skips_withdrawn_applications_and_non_employees():
    details = [
        _row("通勤定期代", "2014013", "柴田 和浩", "35090", status="取下げ"),
        _row("通勤定期代", "5000001", "派遣 太郎", "40000"),
    ]
    assert build_limit_over_rows(details, LIMIT_IDX, {}, limit=30000) == []


def test_load_limit_exempt_members_reads_number_and_reason(tmp_path):
    p = tmp_path / "通勤費_上限免除者.csv"
    p.write_text("社員番号,氏名,理由\n2025017,杉原 司,個別許可\n\n", encoding="utf-8-sig")
    assert load_limit_exempt_members(p) == {"2025017": "個別許可"}


def test_load_limit_exempt_members_returns_empty_when_missing(tmp_path):
    """リストが無くても精査は止めない（超過者が全員 要確認 に出るので気づける）。"""
    assert load_limit_exempt_members(tmp_path / "ない.csv") == {}
