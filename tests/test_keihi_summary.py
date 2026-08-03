# -*- coding: utf-8 -*-
"""経費統合一覧表の生成（マクロ移植 P1a）のユニットテスト。

各ソース変換・SAP税抜補正・顧客請求分の寄せ・経路突合の判定・Excel出力を検証する。
実データ（Rev5）との全行一致は別途 谷津さん提供の突合セットで確認する。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.keihi_summary import (  # noqa: E402
    excel_coerce, norm_date_slash, norm_date_pad, normkey, parse_amount,
    _round_half_up, is_night_duty_text, in_company_scope,
    transform_jinjer, transform_estaffing, transform_sap, transform_freee,
    evaluate_route_check, add_integrated_sheet, run_keihi_integration,
    build_integrated_rows, add_roster_entry,
    INTEGRATED_HEADERS, NCOL,
    C_EMP, C_NAME, C_APPDATE, C_TOTAL, C_MEMO_REQ, C_USEDATE, C_TRANS, C_DETAIL,
    C_BILLTYPE, C_FARE, C_AMOUNT, C_SUBTOTAL, C_MEMO_LINE, C_ENTRY_TYPE,
    C_BOARD, C_ALIGHT, C_CUSTOMER_BILL,
)


# ----------------------------------------------------------------------
# ヘルパー関数
# ----------------------------------------------------------------------

def test_excel_coerce_strips_leading_zeros_and_dates():
    # 仕訳No. の先頭ゼロは Excel 取込で落ちる
    assert excel_coerce("00000236") == "236"
    # 日付はゼロ埋めなしで再表示
    assert excel_coerce("2026/06/01") == "2026/6/1"
    assert excel_coerce("2026/6/30") == "2026/6/30"
    # 社員番号・金額はそのまま
    assert excel_coerce("2026013") == "2026013"
    assert excel_coerce("356") == "356"
    # 8桁の計上日(yyyymmdd) は数値化されてもそのまま
    assert excel_coerce("20260630") == "20260630"
    # 文字列はそのまま
    assert excel_coerce("交通費（電車・バス）") == "交通費（電車・バス）"
    assert excel_coerce("") == ""


def test_norm_dates():
    assert norm_date_slash("2026/06/01") == "2026/6/1"
    assert norm_date_slash("2026-6-1") == "2026/6/1"
    assert norm_date_pad("2026/6/1") == "2026/06/01"
    assert norm_date_pad("2026/4/30 12:17") == "2026/04/30"   # 時刻切捨て


def test_normkey_and_night_duty():
    assert normkey(" 山田　太郎 ") == "山田太郎"
    assert is_night_duty_text("顧客対応当番16") is True
    assert is_night_duty_text("深夜出動手当") is True
    assert is_night_duty_text("交通費") is False


def test_excel_coerce_jp_date():
    from services.keihi_summary import excel_coerce_jp_date
    # Excel が OpenText で「6月5日」を日付化する挙動（年は year_hint から）
    assert excel_coerce_jp_date("6月5日", "2026/06/05") == "2026/06/05"
    assert excel_coerce_jp_date("6/17", "2026/06/17") == "2026/06/17"
    assert excel_coerce_jp_date("2026/6/4", "") == "2026/06/04"
    # 日付でない説明は原文そのまま（前後空白も保持）
    assert excel_coerce_jp_date("東京から潮見の往復", "2026/06/13") == "東京から潮見の往復"
    assert excel_coerce_jp_date("その他の費用(10%) 6/4 ", "2026/06/04") == "その他の費用(10%) 6/4 "
    assert excel_coerce_jp_date("", "2026/06/04") == ""


def test_parse_amount_and_round():
    assert parse_amount("1,375") == "1375"
    assert parse_amount("￥2,750円") == "2750"
    assert parse_amount("") == ""
    # SAP 税抜（÷1.1・四捨五入）
    assert _round_half_up(2750 / 1.1) == 2500
    assert _round_half_up(1375 / 1.1) == 1250
    assert _round_half_up(5500 / 1.1) == 5000


def test_in_company_scope():
    assert in_company_scope("2026013") is True
    assert in_company_scope("5000001") is False   # 派遣
    assert in_company_scope("9000001") is False
    assert in_company_scope("") is False


# ----------------------------------------------------------------------
# ソース別変換
# ----------------------------------------------------------------------

def _jinjer_row(**kw):
    r = [""] * 33
    r[C_EMP] = kw.get("emp", "2026013")
    r[C_NAME] = kw.get("name", "川口 祐ノ輔")
    r[C_TOTAL] = kw.get("total", "356")
    r[C_USEDATE] = kw.get("use", "2026/06/01")
    r[C_TRANS] = kw.get("trans", "交通費（電車・バス）")
    r[C_DETAIL] = kw.get("detail", "")
    r[C_BILLTYPE] = kw.get("billtype", "")
    return r


def test_transform_jinjer_identity_and_coerce():
    rows = transform_jinjer([], [_jinjer_row(use="2026/06/01", total="356")])
    assert len(rows) == 1
    row = rows[0]
    assert row[C_EMP] == "2026013"
    assert row[C_USEDATE] == "2026/6/1"       # ゼロ埋め落ち
    assert row[C_TOTAL] == "356"
    assert row[C_CUSTOMER_BILL] == ""          # 通常行は 34列目 空


def test_transform_jinjer_moves_customer_bill_to_col34():
    # 請求区分=顧客請求・夜間当番でない → 金額を 34列目へ、合計等はクリア
    r = _jinjer_row(billtype="顧客請求", total="8990", detail="客先請求分（交通費）")
    row = transform_jinjer([], [r])[0]
    assert row[C_CUSTOMER_BILL] == "8990"
    assert row[C_TOTAL] == ""
    assert row[C_SUBTOTAL] == ""
    assert row[C_FARE] == ""
    assert row[C_AMOUNT] == ""


def test_transform_jinjer_night_duty_customer_bill_stays():
    # 請求区分=顧客請求 でも夜間当番系なら 34列目へ寄せない（合計に残す）
    r = _jinjer_row(billtype="顧客請求", total="2500", detail="夜間当番手当")
    row = transform_jinjer([], [r])[0]
    assert row[C_CUSTOMER_BILL] == ""
    assert row[C_TOTAL] == "2500"


def test_transform_estaffing_all_to_customer_bill():
    # 立替金CSV 15列: [0]TC番号...[4]スタッフ名 [5]就業年月日 [8]出発地 [9]到着地 [10]交通手段 ...[14]金額
    r = [""] * 15
    r[4], r[5], r[8], r[9], r[10], r[14] = "衣笠 博陽", "2026/6/3", "大阪", "大津", "JR西日本", "960"
    roster = {}
    add_roster_entry(roster, "衣笠 博陽", "2021001")
    row = transform_estaffing([], [r], roster)[0]
    assert row[C_EMP] == "2021001"
    assert row[C_NAME] == "衣笠 博陽"
    assert row[C_TRANS] == "JR西日本"
    assert row[C_BOARD] == "大阪" and row[C_ALIGHT] == "大津"
    assert row[C_AMOUNT] == "960" and row[C_FARE] == "960"
    assert row[C_CUSTOMER_BILL] == "960"       # e-staffing は全額客先請求


def _sap_row(sei="太田", mei="裕一", amt="398", vendor="京葉線", desc="", cc="", sid=""):
    r = [""] * 13
    r[0], r[1], r[2], r[3], r[5], r[8], r[10] = sei, mei, amt, vendor, desc, cc, sid
    r[4] = "2026/6/13"   # 費用エントリ日
    r[6] = "2026/6/30"   # 承認日
    return r


def test_transform_sap_normal_to_customer_bill():
    row = transform_sap([], [_sap_row(vendor="京葉線", amt="398")], {})[0]
    assert row[C_DETAIL] == "京葉線"
    assert row[C_CUSTOMER_BILL] == "398"
    assert row[C_TOTAL] == ""                 # 通常行は D:合計 に入れない


def test_transform_sap_night_duty_tax_strip():
    # 業者名=顧客対応当番16・費用合計=2750（税込）→ ÷1.1=2500 を D:合計へ
    roster = {}
    add_roster_entry(roster, "大北 将司", "2024008")
    r = _sap_row(sei="将司", mei="大北", amt="2750", vendor="顧客対応当番16")
    row = transform_sap([], [r], roster)[0]
    assert row[C_TOTAL] == "2500"             # 税抜補正＋夜間当番→合計
    assert row[C_CUSTOMER_BILL] == ""
    assert row[C_EMP] == "2024008"            # 姓名逆でも照合


def test_transform_sap_8a_copies_desc_to_vendor():
    # 夜間当番KWが説明(F)のみ・業者名(D)空 → 8aでD←F、8bで税抜、夜間当番判定も効く
    r = _sap_row(sei="小池", mei="裕也", amt="2750", vendor="", desc="6/12 顧客当番16")
    row = transform_sap([], [r], {})[0]
    assert row[C_DETAIL] == "6/12 顧客当番16"   # 内訳=業者名(=説明の転記)
    assert row[C_TOTAL] == "2500"


def test_transform_sap_memo_coerces_date_desc():
    # 説明が「6月5日」の行 → Excel日付化を再現し 備考(明細)は費用エントリ日で始まる
    r = _sap_row(sei="将司", mei="大北", amt="2750", vendor="顧客対応当番16",
                 desc="6月5日", cc="ＵＡＬＨＷ保守全般", sid="36828ES00003651")
    row = transform_sap([], [r], {})[0]
    # 費用エントリ日=2026/6/13 → 年は2026、6月5日→2026/06/05
    assert row[C_MEMO_LINE] == "2026/06/05 / CC: ＵＡＬＨＷ保守全般 / ID: 36828ES00003651"
    # 説明が日付でない通常行は原文（前後空白保持）
    r2 = _sap_row(vendor="京葉線", amt="398", desc="東京から潮見の往復 ", cc="X", sid="Y")
    row2 = transform_sap([], [r2], {})[0]
    assert row2[C_MEMO_LINE] == "東京から潮見の往復  / CC: X / ID: Y"


def test_transform_freee_preserves_newline_in_memo():
    # freee 内容の埋め込み改行はそのまま保持する（マクロ CStr 相当）
    head = [""] * 32
    head[2], head[3], head[12], head[13] = "2026/7/3", "稲場　直哉", "6月分", "199"
    head[18], head[22], head[23], head[25], head[29] = "2026/6/22", "交通費", "入：新橋駅\n出：品川駅", "199", ""
    rows = transform_freee([], [head], roster={})
    assert rows[0][C_MEMO_LINE] == "入：新橋駅\n出：品川駅"
    assert rows[0][C_USEDATE] == "2026/06/22"


def _freee_rows():
    # 32列。前方フィル対象: 申請日(2) 申請者(3) 申請タイトル(12) 合計金額(13)
    head = [""] * 32
    head[2], head[3], head[12], head[13] = "2026/7/2", "kousei.shiokawa@nmht.co.jp", "6月_経費精算", "2958"
    head[18], head[22], head[23], head[25], head[29] = "2026/6/30", "会議接待費", "森田との面談", "1100", "承認願います"
    cont = [""] * 32
    cont[18], cont[22], cont[23], cont[25] = "2026/6/26", "交通費", "UAL定例会議", "448"
    return [head, cont]


def test_transform_freee_forward_fill_and_mapping():
    roster = {}
    add_roster_entry(roster, "塩川 浩生", "2019047")
    rows = transform_freee([], _freee_rows(), roster)
    assert len(rows) == 2
    r0, r1 = rows
    # 名称変換（メール→氏名）＋社員番号照合
    assert r0[C_NAME] == "塩川 浩生" and r0[C_EMP] == "2019047"
    # 内訳=経費科目、D=金額、仕訳区分=本社経費、利用日=日付（Append_本社経費 標準経路）
    assert r0[C_DETAIL] == "会議接待費" and r0[C_TOTAL] == "1100"
    assert r0[C_ENTRY_TYPE] == "本社経費"
    assert r0[C_USEDATE] == "2026/06/30"       # F:利用日 ← 日付(ゼロ埋め)
    assert r0[C_MEMO_REQ] == "6月_経費精算"
    assert "森田との面談" in r0[C_MEMO_LINE] and "承認願います" in r0[C_MEMO_LINE]
    # 2行目は申請者/申請日/タイトルが前方フィルされる
    assert r1[C_NAME] == "塩川 浩生"
    assert r1[C_DETAIL] == "交通費" and r1[C_TOTAL] == "448"


def test_freee_unmatched_name_marks_該当なし():
    rows = transform_freee([], _freee_rows(), roster={})
    assert rows[0][C_EMP] == "該当なし"


# ----------------------------------------------------------------------
# 経路突合チェック
# ----------------------------------------------------------------------

def _integrated_transit(emp, board, alight, kikan):
    row = [""] * NCOL
    row[C_EMP], row[C_NAME] = emp, "テスト 太郎"
    row[C_USEDATE] = "2026/6/1"
    row[C_TRANS] = kikan
    row[C_BOARD], row[C_ALIGHT] = board, alight
    row[C_TOTAL] = "300"
    return row


def test_route_check_flags_on_route_non_commute():
    commute = [{"社員番号": "2020001", "出発": "東京", "到着": "品川",
                "経由1": "", "経由2": "", "通勤経路": "", "利用交通機関": "電車"}]
    rows = [
        _integrated_transit("2020001", "東京", "品川", "交通費（電車・バス）"),  # ★経路内なのに通勤系以外
        _integrated_transit("2020001", "東京", "品川", "通勤定期代"),            # OK 通勤系
        _integrated_transit("2020001", "渋谷", "新宿", "通勤定期代"),            # △通勤系だが経路外
    ]
    res = evaluate_route_check(rows, commute)
    verdicts = [r["判定"] for r in res]
    assert verdicts[0].startswith("★")
    assert verdicts[1].startswith("OK")
    assert verdicts[2].startswith("△")


def test_route_check_skips_non_transit_rows():
    # 交通機関も乗降車も無い行（夜間当番手当等）はスキップ
    row = [""] * NCOL
    row[C_EMP], row[C_DETAIL], row[C_TOTAL] = "2020001", "夜間当番手当", "2500"
    assert evaluate_route_check([row], []) == []


def test_route_check_travel_members():
    """移動交通費（立替精算）対象者は経路の有無によらず立替精算が正。"""
    travel = {"2018017": "中村 淳一"}
    rows = [
        # 対象者が立替系で申請 → OK（通勤経路登録なしでも △ にしない）
        _integrated_transit("2018017", "東京", "品川", "交通費（電車・バス）"),
        # 対象者が通勤系を選択 → △ で知らせる
        _integrated_transit("2018017", "東京", "品川", "通勤定期代"),
        # 対象外の人は従来どおり（経路未登録＋通勤系 → △）
        _integrated_transit("2020001", "東京", "品川", "通勤定期代"),
    ]
    res = evaluate_route_check(rows, [], travel)
    assert res[0]["判定"] == "OK（移動交通費対象者＝立替精算）"
    assert res[1]["判定"] == "△逆要確認（移動交通費（立替精算）対象者が通勤系を選択）"
    assert res[2]["判定"] == "△逆要確認（通勤系なのに登録経路と不一致）"


def test_route_check_travel_member_overrides_on_route_flag():
    """対象者は経路内一致でも ★（経路内なのに通勤系以外）にしない。"""
    commute = [{"社員番号": "2018017", "出発": "東京", "到着": "品川",
                "経由1": "", "経由2": "", "通勤経路": "", "利用交通機関": "電車"}]
    rows = [_integrated_transit("2018017", "東京", "品川", "交通費（電車・バス）")]
    res = evaluate_route_check(rows, commute, {"2018017": "中村 淳一"})
    assert res[0]["判定"] == "OK（移動交通費対象者＝立替精算）"
    assert res[0]["一致"] == "経路内"   # 一致欄は事実をそのまま残す


# ----------------------------------------------------------------------
# Excel 出力・統合
# ----------------------------------------------------------------------

def test_add_integrated_sheet(tmp_path):
    from openpyxl import Workbook, load_workbook
    wb = Workbook()
    rows = [transform_jinjer([], [_jinjer_row()])[0]]
    add_integrated_sheet(wb, rows)
    out = tmp_path / "t.xlsx"
    wb.save(out)
    wb2 = load_workbook(out)
    ws = wb2["経費統合一覧表"]
    assert ws.max_column == NCOL
    assert [c.value for c in ws[1]] == INTEGRATED_HEADERS
    assert ws.cell(row=2, column=1).value == "2026013"
    wb2.close()


def test_build_integrated_rows_order_and_counts(tmp_path):
    # jinjer CSV を作って build_integrated_rows（jinjer→e-staffing→SAP→freee 順）
    import csv
    jcsv = tmp_path / "jinjer.csv"
    with open(jcsv, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f)
        header = ["申請者社員番号", "申請者名"] + [f"c{i}" for i in range(31)]
        w.writerow(header)
        w.writerow(_jinjer_row(emp="2026013", total="356"))
    rows, counts = build_integrated_rows(jinjer_csv=jcsv, log_func=lambda m: None)
    assert counts["jinjer"] == 1
    assert counts["estaffing"] is None      # 未取込
    assert rows[0][C_EMP] == "2026013"


def test_run_keihi_integration_jinjer_only(tmp_path):
    # route_check=False かつ jinjer のみ → API 不要でオフライン実行できる
    import csv
    from openpyxl import load_workbook
    jcsv = tmp_path / "jinjer.csv"
    with open(jcsv, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f)
        w.writerow(["申請者社員番号", "申請者名"] + [f"c{i}" for i in range(31)])
        w.writerow(_jinjer_row(emp="2026013", total="356"))
        w.writerow(_jinjer_row(emp="2026013", billtype="顧客請求", total="8990",
                               detail="客先請求分（交通費）"))
    out = tmp_path / "integrated.xlsx"
    res = run_keihi_integration(output_path=out, jinjer_csv=jcsv, route_check=False,
                                log_func=lambda m: None)
    assert res.ok is True
    assert res.integrated_rows == 2
    assert res.source_counts["jinjer"] == 2
    assert out.exists()
    wb = load_workbook(out)
    assert "経費統合一覧表" in wb.sheetnames
    wb.close()
