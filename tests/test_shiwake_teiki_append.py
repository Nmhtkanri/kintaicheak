# -*- coding: utf-8 -*-
"""仕訳データCSVを補う通勤定期代行の生成（イレギュラー救済処置）。

2026-08-07 の手作業（29名31行・562,570円）で起きた2つの事故を再発させないことを主眼に置く。
  ① Excel経由で仕訳No.の先頭ゼロが落ちた → 全工程で文字列のまま扱う
  ② 追記のついでに既存行が書き換わっていた → 既存ファイルには書き込みで触れない（standalone出力）

2026-08-12 拡張: 入力はフォルダでもよく（追加計上分が別CSVになるパターン）、
どれかのCSVに定期代が計上済みの人はスキップする＝二重計上が構造的に起きない。
"""
import csv
import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.kotsuhi_seisa import CommuteMaster  # noqa: E402
from services.shiwake_teiki_append import (  # noqa: E402
    SHIWAKE_COLS,
    ShiwakeError,
    booked_pass_members,
    build_teiki_append_rows,
    find_teiki_reference,
    load_shiwake_sources,
    month_bounds,
    render_shiwake_rows,
    write_standalone_shiwake_csv,
)


# ----------------------------------------------------------------------
# ヘルパー
# ----------------------------------------------------------------------

def _shiwake_row(emp="2026009", name="稲田 真勢", trans="通勤定期代", use="2026/7/7",
                 shiwake_no="00000239", company="株式会社エヌエム・ヒューマテック"):
    """実物（2026年08月06日仕訳データ.csv）の1行を模した33列。"""
    r = [""] * SHIWAKE_COLS
    r[0], r[1], r[2], r[3] = emp, name, use, "1182"
    r[5], r[6] = use, trans
    r[12], r[13], r[14] = "1182", "591", "往復"
    r[20], r[21] = "2026/7/31", "20260731"
    r[22], r[23] = "10", "0"
    r[24], r[25], r[26] = shiwake_no, "計上仕訳", company
    r[29] = "従業員立替"
    r[30], r[31] = "武蔵中原", "豊洲"
    return r


def _write_shiwake(path, rows, encoding="cp932"):
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
    w.writerow(["申請者社員番号", "申請者名"] + [f"c{i}" for i in range(SHIWAKE_COLS - 2)])
    w.writerows(rows)
    path.write_bytes(buf.getvalue().encode(encoding))
    return path


def _master(*legs):
    """(社員番号, 経路No, 出発, 到着, 交通機関, 支給間隔, 支給金額) から通勤費シートを作る。"""
    class _Sheet:
        def __init__(self, rows):
            self._rows = rows

        def iter_rows(self, min_row=1, values_only=False):
            return iter(self._rows)

    rows = [("社員番号",) + ("",) * 14]
    for emp, no, dep, arr, kind, interval, amount in legs:
        rows.append((emp, "氏名", no, dep, arr, "", "", "", kind, interval, "", amount, "", "", ""))
    return CommuteMaster(_Sheet(rows))


def _no_commute(emp="2008003", name="友納 英彦", kubun="通勤定期代",
                judge="マスタから支給", work=22):
    return {"社員番号": emp, "氏名": name, "区分": kubun, "判定": judge,
            "出勤日数": work, "支給金額": 18570}


# ----------------------------------------------------------------------
# 入力CSVの読み込み（ファイル／フォルダ・想定外の形は必ず止める）
# ----------------------------------------------------------------------

def test_load_sources_reads_a_single_file_and_keeps_zero_padded_shiwake_no(tmp_path):
    """仕訳No. の先頭ゼロを落とさない（手作業では 00000240 → 240 になっていた）。"""
    p = _write_shiwake(tmp_path / "仕訳.csv", [_shiwake_row(shiwake_no="00000240")])
    files = load_shiwake_sources(p)
    assert [f.name for f in files] == ["仕訳.csv"]
    assert len(files[0].header) == SHIWAKE_COLS
    assert files[0].body[0][24] == "00000240"


def test_load_sources_reads_every_csv_in_a_folder_sorted_by_name(tmp_path):
    """追加計上分が別CSVになるパターン（8/6本体＋8/7追加）に対応する。"""
    _write_shiwake(tmp_path / "2026年08月07日仕訳データ.csv", [_shiwake_row(emp="2020002")])
    _write_shiwake(tmp_path / "2026年08月06日仕訳データ.csv", [_shiwake_row(emp="2020001")])
    files = load_shiwake_sources(tmp_path)
    assert [f.name for f in files] == ["2026年08月06日仕訳データ.csv", "2026年08月07日仕訳データ.csv"]
    assert sum(len(f.body) for f in files) == 2


def test_load_sources_rejects_an_empty_folder(tmp_path):
    with pytest.raises(ShiwakeError, match="CSVファイルがありません"):
        load_shiwake_sources(tmp_path)


def test_load_sources_names_the_broken_file_in_the_folder(tmp_path):
    """フォルダ指定だとどのファイルが悪いのか分からないので、必ずファイル名を出す。"""
    _write_shiwake(tmp_path / "a_正常.csv", [_shiwake_row()])
    (tmp_path / "b_壊れ.csv").write_bytes("社員番号,氏名\r\n1,2\r\n".encode("cp932"))
    with pytest.raises(ShiwakeError, match="b_壊れ.csv"):
        load_shiwake_sources(tmp_path)


def test_load_sources_rejects_bom(tmp_path):
    p = tmp_path / "bom.csv"
    p.write_bytes(b"\xef\xbb\xbf" + b"a," * (SHIWAKE_COLS - 1) + b"b\r\n")
    with pytest.raises(ShiwakeError, match="BOM"):
        load_shiwake_sources(p)


def test_load_sources_rejects_utf8(tmp_path):
    p = _write_shiwake(tmp_path / "utf8.csv", [_shiwake_row(name="髙橋 淳")], encoding="utf-8")
    with pytest.raises(ShiwakeError, match="CP932"):
        load_shiwake_sources(p)


def test_load_sources_rejects_missing_path(tmp_path):
    with pytest.raises(ShiwakeError, match="見つかりません"):
        load_shiwake_sources(tmp_path / "ない.csv")


# ----------------------------------------------------------------------
# 計上済み検知（追加計上分・過去の追記出力との二重計上を防ぐ）
# ----------------------------------------------------------------------

def test_booked_members_are_collected_across_files_with_the_source_name(tmp_path):
    _write_shiwake(tmp_path / "a.csv", [_shiwake_row(emp="2020001")])
    _write_shiwake(tmp_path / "b_追加.csv", [_shiwake_row(emp="2020002", use="2026/7/1")])
    booked = booked_pass_members(load_shiwake_sources(tmp_path), "2026-07")
    assert booked == {"2020001": "a.csv", "2020002": "b_追加.csv"}


def test_booked_members_ignores_other_months_and_non_pass_rows(tmp_path):
    rows = [_shiwake_row(emp="2020001", use="2026/8/1"),                      # 別月
            _shiwake_row(emp="2020002", trans="交通費（電車・バス）")]        # 定期代でない
    files = load_shiwake_sources(_write_shiwake(tmp_path / "a.csv", rows))
    assert booked_pass_members(files, "2026-07") == {}


def test_build_rows_skips_people_already_booked_with_the_file_name():
    """計上済みの人を追記したら二重計上。スキップ理由にどのファイルかを出す。"""
    master = _master(("2008003", 1, "あざみ野", "銀座", "公共交通機関", "毎月", 18570))
    rows, skipped, _w = build_teiki_append_rows(
        [_no_commute()], master, "2026-07",
        booked={"2008003": "2026年08月07日仕訳データ.csv"})
    assert rows == []
    assert len(skipped) == 1
    assert "計上済み" in skipped[0]["理由"]
    assert "2026年08月07日仕訳データ.csv" in skipped[0]["理由"]


# ----------------------------------------------------------------------
# 雛形（仕訳No.・企業名）の引き継ぎ
# ----------------------------------------------------------------------

def test_find_teiki_reference_prefers_the_newest_file(tmp_path):
    """仕訳No.は出力ごとに変わる（8/6版=239 / 8/7版=240）。最新ファイルを優先する。"""
    _write_shiwake(tmp_path / "2026年08月06日仕訳データ.csv", [_shiwake_row(shiwake_no="00000239")])
    _write_shiwake(tmp_path / "2026年08月07日仕訳データ.csv", [_shiwake_row(shiwake_no="00000240")])
    no, company, source = find_teiki_reference(load_shiwake_sources(tmp_path))
    assert (no, source) == ("00000240", "2026年08月07日仕訳データ.csv")
    assert company == "株式会社エヌエム・ヒューマテック"


def test_find_teiki_reference_falls_back_to_older_files(tmp_path):
    """最新ファイルに定期代が無ければ前のファイルへ遡る（追加計上CSVは実費だけのことがある）。"""
    _write_shiwake(tmp_path / "a.csv", [_shiwake_row(shiwake_no="00000239")])
    _write_shiwake(tmp_path / "b_追加.csv", [_shiwake_row(trans="交通費（電車・バス）")])
    no, _c, source = find_teiki_reference(load_shiwake_sources(tmp_path))
    assert (no, source) == ("00000239", "a.csv")


def test_find_teiki_reference_stops_when_no_file_has_a_template_row(tmp_path):
    _write_shiwake(tmp_path / "a.csv", [_shiwake_row(trans="交通費（電車・バス）")])
    with pytest.raises(ShiwakeError, match="通勤定期代"):
        find_teiki_reference(load_shiwake_sources(tmp_path))


# ----------------------------------------------------------------------
# 対象の選定
# ----------------------------------------------------------------------

def test_build_rows_picks_only_confirmed_commuter_pass_people():
    master = _master(("2008003", 1, "あざみ野", "銀座", "公共交通機関", "毎月", 18570))
    rows, skipped, _w = build_teiki_append_rows([_no_commute()], master, "2026-07")
    assert [r["社員番号"] for r in rows] == ["2008003"]
    assert rows[0]["金額"] == 18570
    assert skipped == []


@pytest.mark.parametrize("kubun,judge,reason", [
    ("通勤費", "マスタから支給", "区分"),
    ("移動交通費", "マスタから支給", "区分"),
    ("通勤定期代", "支給漏れの疑い", "判定"),
    ("通勤定期代", "実費申請の日数不一致", "判定"),
    ("通勤定期代", "勤怠実績なし", "判定"),
])
def test_build_rows_skips_with_a_reason(kubun, judge, reason):
    """対象外は黙って消さず理由を出す（消えた人に気づけないと支給漏れになる）。"""
    master = _master(("2008003", 1, "あざみ野", "銀座", "公共交通機関", "毎月", 18570))
    rows, skipped, _w = build_teiki_append_rows(
        [_no_commute(kubun=kubun, judge=judge)], master, "2026-07")
    assert rows == []
    assert len(skipped) == 1
    assert reason in skipped[0]["理由"]


def test_build_rows_skips_people_without_a_monthly_leg():
    master = _master(("2008003", 1, "あざみ野", "銀座", "公共交通機関", "毎日", 590))
    rows, skipped, _w = build_teiki_append_rows([_no_commute()], master, "2026-07")
    assert rows == []
    assert "毎月支給の経路が無い" in skipped[0]["理由"]


def test_build_rows_includes_zero_attendance_people_with_a_flag():
    master = _master(("2007002", 1, "自宅", "本社", "公共交通機関", "毎月", 12000))
    rows, _s, _w = build_teiki_append_rows(
        [_no_commute(emp="2007002", name="代表 太郎", work=0)], master, "2026-07")
    assert len(rows) == 1
    assert "勤怠実績0日" in rows[0]["要確認"]


# ----------------------------------------------------------------------
# ※要確認フラグ
# ----------------------------------------------------------------------

def test_build_rows_flags_car_commuters():
    master = _master(("2007001", 1, "", "", "車", "毎月", 20000))
    rows, _s, _w = build_teiki_append_rows(
        [_no_commute(emp="2007001", name="菅原 伸")], master, "2026-07")
    assert rows[0]["要確認"] == "車通勤"
    assert rows[0]["乗車場所"] == "" and rows[0]["降車場所"] == ""
    assert "※要確認:車通勤" in rows[0]["備考(明細)"]


def test_build_rows_emits_one_row_per_route_with_position_note():
    master = _master(("2026010", 1, "本厚木", "新宿", "公共交通機関", "毎月", 14870),
                     ("2026010", 2, "新宿", "東京", "公共交通機関", "毎月", 14870))
    rows, _s, _w = build_teiki_append_rows(
        [_no_commute(emp="2026010", name="柳場 涼馬")], master, "2026-07")
    assert [r["金額"] for r in rows] == [14870, 14870]
    assert "経路2本のうち1本目" in rows[0]["要確認"]
    assert "経路2本のうち2本目" in rows[1]["要確認"]


def test_build_rows_flags_over_limit_but_never_cuts():
    master = _master(("2014013", 1, "A", "B", "公共交通機関", "毎月", 35090))
    rows, _s, _w = build_teiki_append_rows(
        [_no_commute(emp="2014013", name="柴田 和浩")], master, "2026-07")
    assert rows[0]["金額"] == 35090
    assert "上限30,000円超" in rows[0]["要確認"]


def test_build_rows_pays_exempt_members_in_full_without_a_flag():
    master = _master(("2016024", 1, "津田沼", "北府中", "公共交通機関", "毎月", 38500))
    rows, _s, _w = build_teiki_append_rows(
        [_no_commute(emp="2016024", name="岡崎 修司")], master, "2026-07",
        limit_exempt={"2016024": "個別許可"})
    assert rows[0]["金額"] == 38500
    assert rows[0]["要確認"] == ""


def test_build_rows_warns_when_daily_legs_also_exist():
    master = _master(("2008003", 1, "あざみ野", "銀座", "公共交通機関", "毎月", 18570),
                     ("2008003", 2, "銀座", "東京", "公共交通機関", "毎日", 590))
    rows, _s, warnings = build_teiki_append_rows([_no_commute()], master, "2026-07")
    assert [r["金額"] for r in rows] == [18570]
    assert any("毎日支給" in w for w in warnings)


# ----------------------------------------------------------------------
# 33列への整形（7月の手作業成果物と同じ形になること）
# ----------------------------------------------------------------------

def test_month_bounds_matches_jinjer_notation():
    assert month_bounds("2026-07") == ("2026/7/1", "2026/7/31", "20260731")
    assert month_bounds("2026-02") == ("2026/2/1", "2026/2/28", "20260228")
    assert month_bounds("2024-02") == ("2024/2/1", "2024/2/29", "20240229")


def test_month_bounds_rejects_bad_input():
    for bad in ("2026/07", "2026-13", "", "７月"):
        with pytest.raises(ShiwakeError):
            month_bounds(bad)


def test_render_matches_the_july_manual_output():
    """2026-08-07 の手作業成果物の1行と同じ内容になること。"""
    preview = [{"社員番号": "2008003", "氏名": "友納 英彦", "経路No": 1,
                "乗車場所": "あざみ野", "降車場所": "銀座", "利用交通機関": "公共交通機関",
                "金額": 18570, "要確認": "",
                "備考(明細)": "通勤費申請なし・マスタから支給（2026年7月）"}]
    row = render_shiwake_rows(preview, "2026-07", "00000240",
                              "株式会社エヌエム・ヒューマテック")[0]
    assert len(row) == SHIWAKE_COLS
    assert row[0] == "2008003" and row[1] == "友納 英彦"
    assert row[2] == ""                                   # 申請日は空
    assert row[3] == row[12] == row[13] == "18570"        # 合計＝小計＝金額(交通費)
    assert row[5] == "2026/7/1"                           # 利用日は月初
    assert row[6] == "通勤定期代"
    assert row[14] == "片道"                              # 定期代は往復＝片道
    assert row[19] == "通勤費申請なし・マスタから支給（2026年7月）"
    assert (row[20], row[21]) == ("2026/7/31", "20260731")
    assert (row[22], row[23]) == ("10", "0")
    assert row[24] == "00000240"                          # 先頭ゼロを保持
    assert row[25] == "計上仕訳"
    assert row[26] == "株式会社エヌエム・ヒューマテック"
    assert row[29] == "従業員立替"
    assert (row[30], row[31]) == ("あざみ野", "銀座")
    assert row[32] == ""                                  # 経路は空欄


def test_render_puts_the_confirm_note_after_a_full_width_space():
    preview = [{"社員番号": "2007001", "氏名": "菅原 伸", "経路No": 1,
                "乗車場所": "", "降車場所": "", "利用交通機関": "車", "金額": 20000,
                "要確認": "車通勤",
                "備考(明細)": "通勤費申請なし・マスタから支給（2026年7月）　※要確認:車通勤"}]
    row = render_shiwake_rows(preview, "2026-07", "240", "株式会社エヌエム・ヒューマテック")[0]
    assert row[19] == "通勤費申請なし・マスタから支給（2026年7月）　※要確認:車通勤"


# ----------------------------------------------------------------------
# standalone出力（既存ファイルには書き込みで触れない）
# ----------------------------------------------------------------------

def _new_rows():
    return render_shiwake_rows(
        [{"社員番号": "2008003", "氏名": "友納 英彦", "経路No": 1,
          "乗車場所": "あざみ野", "降車場所": "銀座", "利用交通機関": "公共交通機関",
          "金額": 18570, "要確認": "",
          "備考(明細)": "通勤費申請なし・マスタから支給（2026年7月）"}],
        "2026-07", "00000240", "株式会社エヌエム・ヒューマテック")


def _header():
    return ["申請者社員番号", "申請者名"] + [f"c{i}" for i in range(SHIWAKE_COLS - 2)]


def test_standalone_output_contains_only_header_and_new_rows(tmp_path):
    src = _write_shiwake(tmp_path / "元.csv", [_shiwake_row(), _shiwake_row(emp="2026010")])
    raw_before = src.read_bytes()
    out = tmp_path / "追記.csv"
    assert write_standalone_shiwake_csv(_header(), _new_rows(), out) == []

    written = out.read_bytes()
    assert written[:3] != b"\xef\xbb\xbf"        # BOMなし
    assert written.endswith(b"\r\n")
    rows = list(csv.reader(io.StringIO(written.decode("cp932"), newline="")))
    assert len(rows) == 2                        # ヘッダー + 追記1行だけ（元データは含まない）
    assert rows[0] == _header()
    assert rows[1][24] == "00000240"             # ゼロ埋め保持
    assert src.read_bytes() == raw_before        # 元ファイルは1バイトも変わらない


def test_standalone_output_stops_before_writing_when_cp932_cannot_encode(tmp_path):
    rows = _new_rows()
    rows[0][1] = "髙橋 淳µ"                 # µ は CP932 にできない
    out = tmp_path / "追記.csv"
    problems = write_standalone_shiwake_csv(_header(), rows, out)
    assert problems and "CP932" in problems[0]
    assert not out.exists()


def test_standalone_output_rejects_wrong_column_counts(tmp_path):
    with pytest.raises(ShiwakeError, match="列"):
        write_standalone_shiwake_csv(_header(), [["a", "b"]], tmp_path / "追記.csv")
    with pytest.raises(ShiwakeError, match="ヘッダー"):
        write_standalone_shiwake_csv(["a", "b"], _new_rows(), tmp_path / "追記.csv")


def test_rerun_with_previous_output_in_the_folder_is_idempotent(tmp_path):
    """出力CSVをフォルダに入れて再実行しても、同じ人が二重に追記されない。"""
    _write_shiwake(tmp_path / "a_元.csv", [_shiwake_row()])
    out = tmp_path / "通勤定期代追記_2026年7月.csv"
    write_standalone_shiwake_csv(_header(), _new_rows(), out)

    files = load_shiwake_sources(tmp_path)
    booked = booked_pass_members(files, "2026-07")
    assert booked["2008003"] == "通勤定期代追記_2026年7月.csv"

    master = _master(("2008003", 1, "あざみ野", "銀座", "公共交通機関", "毎月", 18570))
    rows, skipped, _w = build_teiki_append_rows([_no_commute()], master, "2026-07",
                                                booked=booked)
    assert rows == []
    assert "計上済み" in skipped[0]["理由"]
