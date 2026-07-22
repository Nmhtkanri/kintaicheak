# -*- coding: utf-8 -*-
"""kdx_shift_parser のテスト

合成 word リストでの純粋関数テストと、実PDF（共有フォルダにある場合のみ）の統合テスト。
"""
import os

import pytest

from services.kdx_shift_parser import (
    KDX_DAY_BREAK_MIN,
    KDX_DAY_END,
    KDX_DAY_START,
    KDX_NIGHT_BREAK_MIN,
    KDX_NIGHT_END,
    KDX_NIGHT_START,
    _extract_legend_breaks,
    _norm_code,
    build_kdx_legend,
    is_kdx_shift_pdf,
    parse_kdx_shift_pdf,
    parse_kdx_words,
)

REAL_PDF = r"Z:\jinjer移行\カレンダー\KDX\7月\(エヌエム・ヒューマテック様)2026年7月分勤務シフト表.pdf"


def _word(text, x0, top, width=12.0, height=8.0):
    return {"text": text, "x0": x0, "x1": x0 + width, "top": top, "bottom": top + height}


def _synth_words(title="2026年 7月度勤務スケジュール表"):
    """3名分（日勤2・夜勤1）×5日の最小シフト表 word リストを合成する"""
    words = []
    day_xs = {d: 100.0 + 20.0 * (d - 1) for d in range(1, 32)}
    # 日付ヘッダー行
    words.append(_word("日", 40, 100))
    words.append(_word("付", 55, 100))
    for d, x in day_xs.items():
        words.append(_word(str(d), x - 4, 100, width=8))
    # 曜日行（2026-07-01=水）
    weekday = ["水", "木", "金", "土", "日", "月", "火"]
    for d, x in day_xs.items():
        words.append(_word(weekday[(d - 1) % 7], x - 4, 110, width=8))
    # 日勤: 尾川
    words.append(_word("尾川", 40, 146))
    for d, code in [(1, "A5"), (2, "A1"), (3, "／"), (4, "／"), (5, "有")]:
        words.append(_word(code, day_xs[d] - 6, 150))
    # 日勤: 市川正（右端に「日勤」ラベル混入を再現）
    words.append(_word("市川正", 40, 166))
    for d, code in [(1, "A1"), (2, "A3"), (3, "A5"), (4, "／"), (5, "／")]:
        words.append(_word(code, day_xs[d] - 6, 170))
    words.append(_word("日勤", day_xs[31] + 40, 170))
    # 夜勤: 池田
    words.append(_word("池田", 40, 186))
    for d, code in [(1, "C4"), (2, "ー"), (3, "／"), (4, "C3"), (5, "ー")]:
        words.append(_word(code, day_xs[d] - 6, 190))
    return words


_LEGEND_TEXT = (
    "■休憩取得予定時間 凡例 ①日勤(計1h取得) ②夜勤(計2h取得)\n"
    "A1：11:30～12:30 C1：20:00～21:00／2:00～3:00\n"
    "A5：13:30～14:30 C4：23:00～24:00／5:00～6:00\n"
)


def _synth_text(title="2026年 7月度勤務スケジュール表"):
    return f"株式会社エヌエム・ヒューマテック様\n{title}\n{_LEGEND_TEXT}"


class TestNormCode:
    def test_shift_codes(self):
        assert _norm_code("A5") == "A5"
        assert _norm_code("C4") == "C4"
        assert _norm_code("Ａ５") == "A5"   # 全角も半角へ

    def test_off_and_ake_variants(self):
        assert _norm_code("／") == "／"
        assert _norm_code("/") == "／"
        assert _norm_code("ー") == "ー"
        assert _norm_code("-") == "ー"
        assert _norm_code("−") == "ー"

    def test_yukyu(self):
        assert _norm_code("有") == "有"

    def test_non_codes(self):
        assert _norm_code("日勤") is None
        assert _norm_code("尾川") is None
        assert _norm_code("A7") is None
        assert _norm_code("") is None


class TestLegendExtract:
    def test_day_and_night(self):
        got = _extract_legend_breaks(_LEGEND_TEXT)
        assert got["A1"] == "11:30～12:30"
        assert got["C4"] == "23:00～24:00/5:00～6:00"

    def test_build_kdx_legend_times(self):
        legend = build_kdx_legend(_extract_legend_breaks(_LEGEND_TEXT), {"A1", "C4", "有"})
        by_code = {e["code"]: e for e in legend}
        assert by_code["A1"]["start_time"] == KDX_DAY_START
        assert by_code["A1"]["end_time"] == KDX_DAY_END
        assert by_code["A1"]["break_minutes"] == KDX_DAY_BREAK_MIN
        assert by_code["C4"]["start_time"] == KDX_NIGHT_START
        assert by_code["C4"]["end_time"] == KDX_NIGHT_END      # "34:00" 24時超のまま
        assert by_code["C4"]["break_minutes"] == KDX_NIGHT_BREAK_MIN
        # ー は「明」を含む label（exporter の明け判定＝"休み" が効く）
        assert "明" in by_code["ー"]["label"]
        assert by_code["ー"]["is_off"] is True
        assert by_code["／"]["is_off"] is True
        # 有 は is_off=True（雛形時刻マッチ対象外。一般雛形化は exporter 優先順位0）
        assert by_code["有"]["is_off"] is True


class TestParseKdxWords:
    def test_basic(self):
        result = parse_kdx_words(
            _synth_words(), _synth_text(),
            filename="test.pdf", target_year=2026, target_month=7)
        assert result["year"] == 2026 and result["month"] == 7
        names = [e["name"] for e in result["employees"]]
        assert names == ["尾川", "市川正", "池田"]
        og = result["employees"][0]["shifts"]
        assert [s["code"] for s in og[:5]] == ["A5", "A1", "／", "／", "有"]
        assert og[0]["date"] == "2026-07-01"
        # 6日以降（記号なし）は空欄
        assert all(s["code"] == "" for s in og[5:])
        assert len(og) == 31

    def test_right_side_label_excluded(self):
        """右端の「日勤/シフト」ラベルが記号として混入しない"""
        result = parse_kdx_words(
            _synth_words(), _synth_text(),
            filename="test.pdf", target_year=2026, target_month=7)
        ichikawa = [e for e in result["employees"] if e["name"] == "市川正"][0]
        assert [s["code"] for s in ichikawa["shifts"][:5]] == ["A1", "A3", "A5", "／", "／"]

    def test_year_month_mismatch(self):
        with pytest.raises(ValueError, match="一致しません"):
            parse_kdx_words(
                _synth_words(), _synth_text(),
                filename="test.pdf", target_year=2026, target_month=8)

    def test_target_none_uses_title_year_month(self):
        """対象年月未入力（None）→ PDFタイトルの年月を採用して解析できる"""
        result = parse_kdx_words(
            _synth_words(), _synth_text(),
            filename="test.pdf", target_year=None, target_month=None)
        assert result["year"] == 2026 and result["month"] == 7
        assert len(result["employees"]) == 3

    def test_weekday_mismatch(self):
        """タイトルは合っているが曜日が別月のもの → 誤読み防止で中止"""
        words = _synth_words()
        bad_weekday = ["月", "火", "水", "木", "金", "土", "日"]
        for w in words:
            if w["top"] == 110:
                d = round((w["x0"] + 4 - 100.0) / 20.0) + 1
                w["text"] = bad_weekday[(d - 1) % 7]
        with pytest.raises(ValueError, match="曜日"):
            parse_kdx_words(
                words, _synth_text(),
                filename="test.pdf", target_year=2026, target_month=7)

    def test_no_date_header(self):
        words = [w for w in _synth_words() if not str(w["text"]).isdigit()]
        with pytest.raises(ValueError, match="日付ヘッダー"):
            parse_kdx_words(
                words, _synth_text(),
                filename="test.pdf", target_year=2026, target_month=7)

    def test_legend_and_off_markers(self):
        result = parse_kdx_words(
            _synth_words(), _synth_text(),
            filename="test.pdf", target_year=2026, target_month=7)
        assert set(result["off_markers"]) == {"／", "ー"}
        codes = {e["code"] for e in result["legend"]}
        # 凡例記載(A1/A5/C1/C4) + 表出現(A3) + 固定(ー/／) + 有
        assert {"A1", "A3", "A5", "C1", "C4", "ー", "／", "有"} <= codes


@pytest.mark.skipif(not os.path.exists(REAL_PDF), reason="実PDF（共有フォルダ）が無い環境")
class TestRealPdf:
    def test_sniff(self):
        assert is_kdx_shift_pdf(REAL_PDF) is True

    def test_parse_full(self):
        result = parse_kdx_shift_pdf(REAL_PDF, 2026, 7)
        assert result["year"] == 2026 and result["month"] == 7
        assert len(result["employees"]) == 11
        assert result["section_info"]["weekday_matched"] == 31
        by_name = {e["name"]: e for e in result["employees"]}
        # 日勤・夜勤の代表者をスポットチェック（PDF目視値）
        assert [s["code"] for s in by_name["尾川"]["shifts"][:3]] == ["A5", "A1", "A3"]
        assert [s["code"] for s in by_name["尾川"]["shifts"][-2:]] == ["有", "有"]
        assert [s["code"] for s in by_name["池田"]["shifts"][:3]] == ["C4", "ー", "／"]
        assert [s["code"] for s in by_name["亘（わたり）"]["shifts"][:3]] == ["ー", "／", "C1"]

    def test_year_mismatch_rejected(self):
        with pytest.raises(ValueError):
            parse_kdx_shift_pdf(REAL_PDF, 2026, 6)

    def test_parse_without_target_uses_title(self):
        """対象年月未入力でもタイトルの2026年7月で解析できる"""
        result = parse_kdx_shift_pdf(REAL_PDF)
        assert result["year"] == 2026 and result["month"] == 7
        assert len(result["employees"]) == 11

    def test_structured_files_mismatch_returns_warning(self):
        """年月不一致はAIへ黙って落とさず warning で見せる"""
        from services.multi_year_shift_parser import parse_structured_files
        result = parse_structured_files([REAL_PDF], 2026, 6)
        assert result is not None
        sheets, consumed, warnings = result
        assert sheets == [] and consumed == []
        assert any("フォールバック" in w for w in warnings)
