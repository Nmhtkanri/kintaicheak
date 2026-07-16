# -*- coding: utf-8 -*-
"""経費分類・集計エンジン（マクロ移植 P1b）のユニットテスト。

実データ検証: live 統合一覧表(2401行)で live 集計シートと 173名全員のC〜L完全一致・
集計ログ2457行一致を確認済み（2026-07-16）。ここでは判定順序・エッジケースを固定する。
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.keihi_classify import (  # noqa: E402
    DEFAULT_KEYWORDS, CAT_YAKAN, CAT_TRANS, CAT_NONTAX,
    normalize_str, normalize_id, parse_amount_vba, try_parse_date, hit_any,
    build_emp_key, parse_emp_no, classify_rows, aggregate_by_id,
    build_detail_totals, add_summary_sheet, add_log_sheet, load_keywords,
    B_YAKAN, B_RINK, B_TRANS, B_ETC, B_TW, B_BILL, B_NONTAX,
    COL_EMP, COL_NAME, COL_AMT, COL_MEMO_REQ, COL_TRANS, COL_UCH,
    COL_BILLTYPE, COL_EXPTYPE, COL_FARE, COL_MEMO_LINE, COL_BOOKED,
    COL_SOURCE, COL_ESTAFF,
)
from services.keihi_payroll_import import (  # noqa: E402
    build_import_rows, render_import_csv, IMPORT_HEADERS,
)


def _row(**kw):
    r = [""] * 34
    r[COL_EMP] = kw.get("emp", "2026013")
    r[COL_NAME] = kw.get("name", "川口 祐ノ輔")
    r[COL_AMT] = kw.get("amt", "")
    r[COL_UCH] = kw.get("uch", "")
    r[COL_TRANS] = kw.get("trans", "")
    r[COL_BILLTYPE] = kw.get("billtype_id", "")
    r[COL_EXPTYPE] = kw.get("exptype", "")
    r[COL_MEMO_REQ] = kw.get("memo_req", "")
    r[COL_MEMO_LINE] = kw.get("memo_line", "")
    r[COL_FARE] = kw.get("fare", "")
    r[COL_BOOKED] = kw.get("booked", "")
    r[COL_SOURCE] = kw.get("source", "")
    r[COL_ESTAFF] = kw.get("estaff", "")
    return r


def _bucket(rows):
    cls = classify_rows(rows)
    assert len(cls.agg) == 1
    return next(iter(cls.agg.values())), cls


# ----------------------------------------------------------------------
# VBA互換ヘルパー
# ----------------------------------------------------------------------

def test_normalize_str_vba():
    # CR/LF/Tab/全角空白→半角空白→Trim→LCase→全空白除去
    assert normalize_str(" A B\r\nC　D\t ") == "abcd"
    assert normalize_str(None) == ""


def test_parse_amount_vba_negative_parens():
    assert parse_amount_vba("1,375") == 1375
    assert parse_amount_vba("(123)") == -123
    assert parse_amount_vba("（4,560円）") == -4560
    assert parse_amount_vba("abc") == 0
    assert parse_amount_vba("") == 0


def test_try_parse_date():
    assert try_parse_date("2026/6/30") == date(2026, 6, 30)
    assert try_parse_date("2026-06-01") == date(2026, 6, 1)
    assert try_parse_date("") is None
    assert try_parse_date("30日") is None


def test_hit_any_order_first_wins():
    # 登録順の先勝ち: 交通費リストでは「通勤交通費（実費）」が「交通費」より先
    kw = DEFAULT_KEYWORDS[CAT_TRANS]
    assert hit_any(normalize_str("通勤交通費（実費）"), kw) == "通勤交通費（実費）".lower()
    # 部分一致: "jr西日本" は "JR"（LCase格納）にヒット
    assert hit_any(normalize_str("JR西日本"), kw) == "jr"
    assert hit_any("該当なしのテキスト", DEFAULT_KEYWORDS[CAT_YAKAN]) == ""


def test_emp_key_roundtrip():
    key = build_emp_key("2026013", "かわぐち")
    assert parse_emp_no(key) == "2026013"
    assert parse_emp_no(build_emp_key("", "かわぐち")) == ""


def test_normalize_id_digits_only():
    assert normalize_id("2026013") == "2026013"
    assert normalize_id("該当なし") == ""
    assert normalize_id(" 2026-013 ") == "2026013"


# ----------------------------------------------------------------------
# 分類の優先順位（Collect_From_Source 移植）
# ----------------------------------------------------------------------

def test_priority_0_honsha_to_nontax():
    b, _ = _bucket([_row(amt="500", uch="会議費", source="本社経費")])
    assert b[B_NONTAX] == 500 and b[B_ETC] == 0


def test_priority_1_yakan():
    b, _ = _bucket([_row(amt="2500", uch="夜間当番手当(旧_顧客対応当番手当)")])
    assert b[B_YAKAN] == 2500


def test_priority_2_telework_label_is_J():
    # 集計先はK(bucket4)だがログの判定結果はマクロのまま "J:テレワーク手当"
    b, cls = _bucket([_row(amt="3000", uch="テレワーク手当")])
    assert b[B_TW] == 3000
    assert any(e.result == "J:テレワーク手当" for e in cls.log)


def test_priority_3_rink():
    b, _ = _bucket([_row(amt="1000", uch="RINK日当（平日）1")])
    assert b[B_RINK] == 1000


def test_priority_4_nontax_uses_nontax_text_not_memo():
    # 癖1修正: 備考に旧分類名が残っていても非課税精算に引っ張られない
    b, _ = _bucket([_row(amt="356", trans="通勤交通費（実費）",
                         memo_line="交通費（電車・バス）⇒通勤交通費（実費）に変更")])
    # nonTaxText(内訳+交通機関+費用種別)に非課税KWなし → #5交通費でH
    assert b[B_TRANS] == 356 and b[B_NONTAX] == 0


def test_priority_4_nontax_by_transport():
    b, _ = _bucket([_row(amt="356", trans="交通費（電車・バス）")])
    assert b[B_NONTAX] == 356      # 交通機関が本当に非課税KWの行はI


def test_priority_5_trans_ng_goes_to_etc():
    # 内訳が交通費KWでも交通費除外KW（会議等）に当たればその他(J)
    b, cls = _bucket([_row(amt="800", uch="会議への移動")])
    assert b[B_ETC] == 800 and b[B_TRANS] == 0
    assert any("交通費NG" in e.result for e in cls.log)


def test_priority_6_etc_via_judgetext():
    # judgeText(備考含む)でその他KWにヒット
    b, _ = _bucket([_row(amt="1200", uch="雑品", memo_line="懇親会の飲み物")])
    assert b[B_ETC] == 1200


def test_priority_7_default_etc():
    b, cls = _bucket([_row(amt="999", uch="未知の内訳XYZ")])
    assert b[B_ETC] == 999
    assert any(e.matched_kw == "(該当キーワードなし)" for e in cls.log)


def test_judgetext_uses_billtype_id_column():
    # judgeTextの請求区分成分は請求区分ID列(0-based 8)。ID列に「会議」を置けば#6でヒットする
    b, _ = _bucket([_row(amt="700", uch="不明品目", billtype_id="会議")])
    assert b[B_ETC] == 700


# ----------------------------------------------------------------------
# G先行・estFilledガード・金額フォールバック
# ----------------------------------------------------------------------

def test_g_precheck_requires_emp_no():
    # 社員番号が無い行は顧客請求分(G)に加算されない
    cls = classify_rows([_row(emp="該当なし", name="稲場", estaff="1000")])
    assert all(v[B_BILL] == 0 for v in cls.agg.values()) or not cls.agg


def test_g_precheck_excluded_by_kokyaku_ng():
    # 顧客請求除外KW(夜間当番)に当たる → G加算なし（aggエントリ自体できない）・"G:除外" ログ
    cls = classify_rows([_row(amt="", uch="夜間当番手当", estaff="2500")])
    assert not cls.agg
    assert any(e.result == "G:除外" for e in cls.log)


def test_estfilled_guard_skips_amount_and_date():
    # 顧客請求分あり＆夜間当番でない → 金額分類スキップ（計上日もスキップ=GoTo NextR）
    rows = [_row(amt="1160", uch="客先請求分（交通費）", estaff="1160", booked="2026/6/30")]
    cls = classify_rows(rows)
    b = next(iter(cls.agg.values()))
    assert b[B_BILL] == 1160          # G先行は効く
    assert b[B_TRANS] == 0 and b[B_ETC] == 0   # 金額は分類されない
    assert not cls.max_date            # 計上日もスキップ
    assert any(e.result == "D:顧客請求費ありのため除外" for e in cls.log)


def test_estfilled_zero_string_still_guards():
    # AH="0" でも estFilled=True（文字列非空判定）
    rows = [_row(amt="500", uch="不明", estaff="0")]
    cls = classify_rows(rows)
    assert any(e.result == "D:顧客請求費ありのため除外" for e in cls.log)


def test_yakan_passes_guard_even_with_estaff():
    # estFilledでも内訳が夜間当番ならD計上（ガードを通過）
    b, _ = _bucket([_row(amt="2500", uch="顧客対応当番16", estaff="x")])
    assert b[B_YAKAN] == 2500


def test_amount_fallback_fare_positive_only():
    b, _ = _bucket([_row(amt="0", fare="178", trans="交通費（電車・バス）")])
    assert b[B_NONTAX] == 178          # 合計0→金額(交通費)>0を採用
    cls2 = classify_rows([_row(amt="0", fare="-50", uch="謎")])
    assert not any(e.result.startswith(("D:", "H:", "I:", "J:", "E:")) for e in cls2.log)


def test_booked_date_max():
    rows = [
        _row(amt="100", trans="交通費（電車・バス）", booked="2026/6/1"),
        _row(amt="100", trans="交通費（電車・バス）", booked="2026/6/30"),
        _row(amt="", booked="2026/7/15"),   # 金額0でも計上日は対象
    ]
    cls = classify_rows(rows)
    assert list(cls.max_date.values())[0] == date(2026, 7, 15)


# ----------------------------------------------------------------------
# 再集約・集計シート・新入社員包含
# ----------------------------------------------------------------------

def test_aggregate_merges_same_id_and_drops_name_only():
    rows = [
        _row(emp="2026013", name="川口 祐ノ輔", amt="100", trans="交通費（電車・バス）"),
        _row(emp="2026013", name="川口　祐ノ輔", amt="200", trans="交通費（電車・バス）"),  # 氏名表記ゆれ
        _row(emp="該当なし", name="不明 太郎", amt="300", uch="雑費"),                      # ID無し
    ]
    agg = aggregate_by_id(classify_rows(rows))
    assert agg.by_id["2026013"][B_NONTAX] == 300
    assert agg.unmatched_rows == 1 and agg.unmatched_amount == 300


def test_summary_sheet_includes_new_employee_and_excludes_569():
    from openpyxl import Workbook
    rows = [
        _row(emp="2026013", name="川口 祐ノ輔", amt="100", trans="交通費（電車・バス）", booked="2026/6/30"),
        _row(emp="5000001", name="派遣 太郎", amt="999", uch="雑費"),
    ]
    cls = classify_rows(rows)
    agg = aggregate_by_id(cls)
    wb = Workbook()
    stats = add_summary_sheet(wb, agg, cls.emp_names, build_detail_totals(rows))
    ws = wb["集計"]
    ids = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert "2026013" in ids            # 新入社員が自動で行になる（事故の根本対策）
    assert "5000001" not in ids        # 5始まりは給与計算対象外
    assert stats["excluded_out_of_scope"] == 1
    # C=合計、L=請求日、M=明細合計、O=判定
    i = ids.index("2026013") + 2
    assert ws.cell(row=i, column=3).value == 100
    assert ws.cell(row=i, column=12).value == date(2026, 6, 30)
    assert ws.cell(row=i, column=15).value == "OK"


def test_detail_totals_d_or_ah():
    rows = [
        _row(emp="2026013", amt="100"),           # D非空 → D
        _row(emp="2026013", amt="", estaff="50"), # D空欄 → AH
    ]
    assert build_detail_totals(rows)["2026013"] == 150


def test_log_sheet_headers():
    from openpyxl import Workbook
    cls = classify_rows([_row(amt="100", trans="交通費（電車・バス）")])
    wb = Workbook()
    add_log_sheet(wb, cls.log)
    ws = wb["集計ログ"]
    assert [c.value for c in ws[1]] == ["行番号", "社員番号", "氏名", "内訳", "金額", "判定結果", "マッチしたキーワード"]


def test_load_keywords_csv(tmp_path):
    import csv as _csv
    p = tmp_path / "kw.csv"
    with open(p, "w", encoding="cp932", newline="") as f:
        w = _csv.writer(f)
        w.writerow(["分類名", "キーワード"])
        w.writerow([CAT_YAKAN, "夜間当番"])
        w.writerow([CAT_YAKAN, ""])                 # 空KWはスキップ（R69対応）
        w.writerow([CAT_YAKAN, "顧客当番", "手当"])   # C列以降は無視（R6対応）
        w.writerow([CAT_NONTAX, "日当"])
    kw = load_keywords(p)
    assert kw[CAT_YAKAN] == ["夜間当番", "顧客当番"]
    assert kw[CAT_NONTAX] == ["日当"]


# ----------------------------------------------------------------------
# jinjer給与インポート行（P2）
# ----------------------------------------------------------------------

def test_build_import_rows_mapping():
    by_id = {
        # [夜間, RINK, 交通費, その他, テレワーク, 顧客請求, 非課税精算]
        "2026013": [2500.0, 1000.0, 5000.0, 300.0, 400.0, 8990.0, 1160.0],
        "5000001": [1.0, 0, 0, 0, 0, 0, 0],   # 対象外
    }
    rows, warnings = build_import_rows(by_id, {"2026013": "川口 祐ノ輔"})
    assert len(rows) == 1
    r = rows[0]
    assert r["夜間当番手当"] == 3500            # D+E（F列相当。live集計R/S実測に基づく）
    assert r["非課税通勤費"] == 5000            # H交通費
    assert r["立替金（顧客請求分）"] == 8990     # G
    assert r["立替金"] == 1160                  # I非課税精算
    assert r["その他"] == 300                   # J
    assert r["定常外業務対応手当"] == 0 and r["その他手当"] == 0 and r["現物支給"] == 0
    assert r["支給過不足調整"] == 0
    assert any("テレワーク手当" in w for w in warnings)   # K=400は未搬送の警告


def test_render_import_csv_sjis_and_quoting():
    rows, _ = build_import_rows({"2026013": [0, 0, 100.0, 0, 0, 0, 0]}, {"2026013": "川口 祐ノ輔"})
    data = render_import_csv(rows)
    text = data.decode("cp932")
    lines = text.strip().split("\r\n")
    assert lines[0] == ",".join(IMPORT_HEADERS)
    assert lines[1].startswith('"2026013","川口 祐ノ輔",0,0,0,100,')
