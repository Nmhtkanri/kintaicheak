"""支払期日ルール（PDFに入金期日が無い取引先を救う）。"""
from datetime import date

from services import invoice_mode as im


def _rules(tmp_path, body):
    path = tmp_path / "due.csv"
    path.write_text("取引先,支払期日\n" + body, encoding="utf-8-sig")
    return path


def test_rules_are_keyed_by_company_ignoring_legal_form(tmp_path):
    """「株式会社」の有無や全角半角が違っても同じ取引先として引ける。"""
    path = _rules(tmp_path, "アイエックス・ナレッジ株式会社,翌月末\n")
    rules = im.load_due_date_rules(path)
    assert rules[im._company_key("アイエックス・ナレッジ")] == "翌月末"


def test_resolve_month_end_rules():
    month_end = date(2026, 7, 31)
    assert im.resolve_due_date("当月末", month_end) == "2026-07-31"
    assert im.resolve_due_date("翌月末", month_end) == "2026-08-31"
    assert im.resolve_due_date("翌々月末", month_end) == "2026-09-30"


def test_resolve_fixed_day_rules():
    month_end = date(2026, 7, 31)
    assert im.resolve_due_date("翌月10日", month_end) == "2026-08-10"
    assert im.resolve_due_date("翌々月10日", month_end) == "2026-09-10"
    # 年をまたぐ
    assert im.resolve_due_date("翌月末", date(2026, 12, 31)) == "2027-01-31"
    # その月に無い日は月末に丸める
    assert im.resolve_due_date("翌月31日", date(2027, 1, 31)) == "2027-02-28"


def test_unknown_rule_is_not_guessed():
    """解釈できないルールで勝手な日付を作らない。"""
    month_end = date(2026, 7, 31)
    assert im.resolve_due_date("要相談", month_end) == ""
    assert im.resolve_due_date("", month_end) == ""


def test_no_csv_means_no_rules():
    assert im.load_due_date_rules(None) == {}
    assert im.load_due_date_rules("Z:/no/such/file.csv") == {}
