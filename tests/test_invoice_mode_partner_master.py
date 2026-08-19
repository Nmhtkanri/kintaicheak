"""取引先マスタ（freeeの登録名に寄せる／入金期日が無い取引先を救う）。"""
from datetime import date

from services import invoice_mode as im


def _master(tmp_path, body):
    path = tmp_path / "partner.csv"
    path.write_text("取引先,freee取引先名,支払期日,備考\n" + body, encoding="utf-8-sig")
    return path


def test_keyed_by_company_ignoring_legal_form(tmp_path):
    """「株式会社」の有無や全角半角が違っても同じ取引先として引ける。"""
    path = _master(tmp_path, "アイエックス・ナレッジ株式会社,,翌月末,\n")
    master = im.load_partner_master(path)
    assert master[im._company_key("アイエックス・ナレッジ")]["due"] == "翌月末"


def test_freee_official_name_is_kept(tmp_path):
    """freeeは取引先を名前で照合するので、登録名に寄せる必要がある。

    「アクシスＩＴパートナーズ株式会社（旧アクシス）」のように全角や
    （旧…）が付いていると、そのまま取り込むと別の取引先が作られてしまう。
    """
    path = _master(
        tmp_path,
        "アクシスITパートナーズ株式会社,アクシスＩＴパートナーズ株式会社（旧アクシス）,,\n")
    got = im.load_partner_master(path)[im._company_key("アクシスITパートナーズ株式会社")]
    assert got["freee_name"] == "アクシスＩＴパートナーズ株式会社（旧アクシス）"
    assert got["due"] == ""


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


def test_resolve_months_later_rule():
    """エリクソン様のように3か月先が期日になる取引先。

    freeeの実データで 2026-05-31→08-31、06-30→09-30、07-31→10-31 と確認。
    """
    assert im.resolve_due_date("3か月後末", date(2026, 7, 31)) == "2026-10-31"
    assert im.resolve_due_date("3ヶ月後末", date(2026, 6, 30)) == "2026-09-30"
    assert im.resolve_due_date("翌々々月末", date(2026, 5, 31)) == "2026-08-31"
    assert im.resolve_due_date("3か月後10日", date(2026, 7, 31)) == "2026-10-10"
    assert im.resolve_due_date("3か月後末", date(2026, 11, 30)) == "2027-02-28"


def test_unknown_rule_is_not_guessed():
    """解釈できないルールで勝手な日付を作らない。"""
    month_end = date(2026, 7, 31)
    assert im.resolve_due_date("要相談", month_end) == ""
    assert im.resolve_due_date("", month_end) == ""


def test_no_csv_means_no_settings():
    assert im.load_partner_master(None) == {}
    assert im.load_partner_master("Z:/no/such/file.csv") == {}
