"""UAL勤務管理表（KDDI小山）パーサと、対象者リストによる絞り込みのテスト

背景（2026-07-31）:
  このブックは 1シート=1か月 で過去分を残していく運用のため、AI読み取りに回すと
  全シートを渡すことになり応答が返らず 25分以上ハングした。対象月のシートだけを
  構造化パースする。
  また **他社の方も同じ表に載っている**。姓だけの氏名なので、他社の「小島」が
  当社の小島さん(2024044)に名前一致してしまう。対象者リストで必ず絞り込む。
"""
from __future__ import annotations

import csv
import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

openpyxl = pytest.importorskip("openpyxl")

from services.employee_alias import (
    filter_employees_by_roster,
    load_roster,
    load_roster_for_source,
    normalize_name_key,
    roster_csv_path,
)
from services.ual_shift_parser import (
    UAL_SOURCE,
    build_ual_legend,
    is_ual_shift_xlsx,
    parse_ual_shift_xlsx,
    parse_ual_worksheet,
    sheet_name_for,
)


# =============================================================================
# テスト用ブックの生成（実物と同じレイアウト）
# =============================================================================

def _make_workbook(path, months=("202608",), *, with_rules=True, codes=None):
    """実物と同じレイアウトの UAL勤務管理表 を作る

    B2=月初日 / C2..=日 / B4..=氏名 / C4..=記号 / 集計行 / D22-D23=ルール
    """
    from datetime import datetime
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    default_codes = {
        "小島": ["B", "明", "A"],       # 他社
        "大堀": ["A", "B", "明"],       # 当社
        "木村": ["×", "年", "D"],       # 当社
    }
    people = codes if codes is not None else default_codes

    for sheet_name in months:
        year, month = int(sheet_name[:4]), int(sheet_name[4:])
        ws = wb.create_sheet(sheet_name)
        ws.cell(row=2, column=2, value=datetime(year, month, 1))
        import calendar
        days = calendar.monthrange(year, month)[1]
        for d in range(1, days + 1):
            ws.cell(row=2, column=2 + d, value=d)

        row = 4
        for name, day_codes in people.items():
            ws.cell(row=row, column=2, value=name)
            for i, c in enumerate(day_codes):
                if c:
                    ws.cell(row=row, column=3 + i, value=c)
            row += 1
        row += 1  # 空行
        for label in ("A", "B", "明", "小計"):
            ws.cell(row=row, column=2, value=label)
            row += 1

        if with_rules:
            ws.cell(row=22, column=4, value="・A勤：09:00～17:30 (休憩1H)")
            ws.cell(row=23, column=4, value="・B勤：16:45～33:30（09:30）(休憩1:45)")

    ws2 = wb.create_sheet("年休希望シート")
    ws2.cell(row=1, column=1, value="年休希望")
    wb.save(path)
    return path


# =============================================================================
# パーサ
# =============================================================================

def test_sheet_name_for():
    assert sheet_name_for(2026, 8) == "202608"
    assert sheet_name_for(2026, 12) == "202612"


def test_sniff_detects_ual_workbook(tmp_path):
    p = _make_workbook(str(tmp_path / "UAL勤務管理表.xlsx"))
    assert is_ual_shift_xlsx(p) is True


def test_sniff_rejects_workbook_without_rules(tmp_path):
    """YYYYMM シートがあってもA勤/B勤ルールが無ければ対象外"""
    p = _make_workbook(str(tmp_path / "別物.xlsx"), with_rules=False)
    assert is_ual_shift_xlsx(p) is False


def test_sniff_rejects_non_xlsx():
    assert is_ual_shift_xlsx("foo.pdf") is False
    assert is_ual_shift_xlsx("foo.csv") is False


def test_parse_reads_only_target_month(tmp_path):
    """過去分シートがあっても対象月のシートだけを読む（ハングの再発防止）"""
    p = _make_workbook(str(tmp_path / "UAL.xlsx"),
                       months=("202604", "202605", "202608", "202609"))
    result = parse_ual_shift_xlsx(p, 2026, 8)

    assert result["year"] == 2026 and result["month"] == 8
    assert result["source"] == UAL_SOURCE
    assert result["section_info"]["sheet"] == "202608"
    assert len(result["employees"]) == 3
    assert [e["name"] for e in result["employees"]] == ["小島", "大堀", "木村"]
    # 31日分そろっている
    assert all(len(e["shifts"]) == 31 for e in result["employees"])


def test_parse_missing_month_sheet_raises(tmp_path):
    p = _make_workbook(str(tmp_path / "UAL.xlsx"), months=("202608",))
    with pytest.raises(ValueError, match="202612"):
        parse_ual_shift_xlsx(p, 2026, 12)


def test_parse_requires_target_month(tmp_path):
    p = _make_workbook(str(tmp_path / "UAL.xlsx"))
    with pytest.raises(ValueError, match="対象年月"):
        parse_ual_shift_xlsx(p, None, None)


def test_parse_stops_at_summary_rows(tmp_path):
    """集計行（A / B / 明 / 小計）を従業員として拾わない"""
    p = _make_workbook(str(tmp_path / "UAL.xlsx"))
    names = [e["name"] for e in parse_ual_shift_xlsx(p, 2026, 8)["employees"]]
    for label in ("A", "B", "明", "小計"):
        assert label not in names


def test_parse_rules_from_workbook(tmp_path):
    """勤務時刻はブックの【ルール】欄から読む（表が変わったら追従する）"""
    p = _make_workbook(str(tmp_path / "UAL.xlsx"))
    legend = {e["code"]: e for e in parse_ual_shift_xlsx(p, 2026, 8)["legend"]}

    assert legend["A"]["start_time"] == "09:00"
    assert legend["A"]["end_time"] == "17:30"
    assert legend["A"]["break_minutes"] == 60
    assert legend["B"]["start_time"] == "16:45"
    assert legend["B"]["end_time"] == "33:30"      # 24時超表記のまま保持する
    assert legend["B"]["break_minutes"] == 105     # 1:45
    assert legend["A"]["is_off"] is False and legend["B"]["is_off"] is False


def test_legend_codes_and_labels(tmp_path):
    p = _make_workbook(str(tmp_path / "UAL.xlsx"))
    legend = {e["code"]: e for e in parse_ual_shift_xlsx(p, 2026, 8)["legend"]}

    # 「明」はラベルに「明」が入ることで exporter が「休み」にする
    assert "明" in legend["明"]["label"] and legend["明"]["is_off"] is True
    # 年休は exporter 側で「一般」雛形になるようラベルを完全一致させる
    assert legend["年"]["label"] == "年次有給休暇"
    assert legend["D"]["label"] == "特別休暇"
    assert legend["×"]["label"] == "公休"


def test_month_mismatch_in_sheet_raises(tmp_path):
    """シート名と中身の月初日が食い違ったら中止する（誤月投入の安全弁）"""
    from datetime import datetime
    p = str(tmp_path / "UAL.xlsx")
    _make_workbook(p, months=("202608",))
    wb = openpyxl.load_workbook(p)
    wb["202608"].cell(row=2, column=2, value=datetime(2026, 5, 1))
    wb.save(p)

    with pytest.raises(ValueError, match="一致しません"):
        parse_ual_shift_xlsx(p, 2026, 8)


def test_build_legend_only_includes_seen_codes():
    rules = {"A": ("9:00", "17:30", 60), "B": ("16:45", "33:30", 105)}
    codes = {e["code"] for e in build_ual_legend(rules, {"A", "B", "明"})}
    assert codes == {"A", "B", "明"}      # 年・D・× は出てこない


# =============================================================================
# 対象者リスト（他社の方を落とす）
# =============================================================================

def _write_roster(path, rows):
    with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)


def test_roster_filters_out_other_companies(tmp_path):
    """他社の「小島」を落とす。当社の小島さん(2024044)に化けるのを防ぐ。"""
    _write_roster(tmp_path / "スケジュール対象者_KDDI小山.csv",
                  [["シフト表氏名", "従業員ID"],
                   ["大堀", "2024042"], ["木村", ""]])

    names, ids, warning = load_roster_for_source(UAL_SOURCE, str(tmp_path))
    assert warning == ""
    assert names == {"大堀", "木村"}
    assert ids == {"大堀": "2024042"}      # ID空欄の木村は入らない

    employees = [{"name": "小島"}, {"name": "大堀"}, {"name": "木村"}]
    kept, excluded = filter_employees_by_roster(employees, names)
    assert [e["name"] for e in kept] == ["大堀", "木村"]
    assert excluded == ["小島"]


def test_roster_absent_source_does_not_filter(tmp_path):
    """対象者リストを持たない系統（KDX等）は絞り込まない"""
    names, ids, warning = load_roster_for_source("kdx", str(tmp_path))
    assert names is None and ids == {} and warning == ""

    employees = [{"name": "尾川"}, {"name": "椎津"}]
    kept, excluded = filter_employees_by_roster(employees, names)
    assert len(kept) == 2 and excluded == []


def test_roster_missing_file_warns(tmp_path):
    """リストが要る系統なのにファイルが無い＝他社を取り込む危険があるので警告する"""
    names, ids, warning = load_roster_for_source(UAL_SOURCE, str(tmp_path))
    assert names is None
    assert "対象者リストが見つかりません" in warning


def test_roster_empty_file_warns(tmp_path):
    _write_roster(tmp_path / "スケジュール対象者_KDDI小山.csv", [["シフト表氏名", "従業員ID"]])
    names, _ids, warning = load_roster_for_source(UAL_SOURCE, str(tmp_path))
    assert names is None and "空です" in warning


def test_roster_rejects_non_employee_id(tmp_path):
    """派遣・テスト番号は対象者としては残すが、IDとしては採用しない"""
    p = tmp_path / "roster.csv"
    _write_roster(p, [["シフト表氏名", "従業員ID"], ["派遣さん", "5000001"]])
    names, ids = load_roster(str(p))
    assert names == {"派遣さん"} and ids == {}


def test_roster_name_key_ignores_spaces():
    assert normalize_name_key(" 及川 航平 ") == "及川航平"
    assert normalize_name_key("及川　航平") == "及川航平"


def test_roster_csv_path_scoping(tmp_path):
    assert roster_csv_path(UAL_SOURCE, str(tmp_path)).endswith("スケジュール対象者_KDDI小山.csv")
    assert roster_csv_path("kdx", str(tmp_path)) == ""
    assert roster_csv_path(UAL_SOURCE, "") == ""


# =============================================================================
# 実データ（あれば）
# =============================================================================

_REAL = r"Z:\jinjer移行\カレンダー\KDDI小山\8月\UAL勤務管理表 (6) のコピー.xlsx"


@pytest.mark.skipif(not os.path.exists(_REAL), reason="実ファイルが無い環境")
def test_real_workbook_2026_08():
    result = parse_ual_shift_xlsx(_REAL, 2026, 8)
    assert len(result["employees"]) == 11
    assert [e["name"] for e in result["employees"]][:3] == ["小島", "角", "鈴木"]
    assert all(len(e["shifts"]) == 31 for e in result["employees"])
    assert result["unknown_codes"] == []

    legend = {e["code"]: e for e in result["legend"]}
    assert legend["B"]["end_time"] == "33:30"
    assert set(legend) == {"A", "B", "明", "×", "年", "D"}
