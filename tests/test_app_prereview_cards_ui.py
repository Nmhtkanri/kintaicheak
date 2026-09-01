"""承認前精査「要確認」のカード表示の画面配線。

2026-09-01、画面に「要確認 13」と出ているのにカードが0枚という状態になった。
件数は6シート合計なのに、カードにしていたのは定期代突合・実費突合の2シートだけで、
その月の13件は通勤費申請なし(4)とマスタ更新漏れ(9)＝どちらもカード対象外だった。
「数えたものは必ず画面に出す」をここで固定する。
"""

import re

import app as app_module

from services.kotsuhi_seisa import FLAGGED_SHEETS


def _html():
    return app_module.app.test_client().get("/").get_data(as_text=True)


def _between(text, start, end):
    """start と end に挟まれた部分。正規表現より読みやすいので素で切る。"""
    assert start in text, start + " が画面に無い"
    i = text.index(start) + len(start)
    return text[i:text.index(end, i)]


def test_cards_cover_every_sheet_that_is_counted():
    # ここがずれると「件数には入っているのにカードが出ない人」が生まれる
    keys = re.findall("'([^']+)':", _between(_html(), "const SHEET_INFO = {", "};"))
    assert keys == list(FLAGGED_SHEETS)


def test_card_area_is_wired_to_flagged_rows():
    html = _html()
    for el in ("pr-flagged-area", "pr-flagged-table", "pr-flagged-filter"):
        assert 'id="' + el + '"' in html, el + " が無い"
    assert "prRenderFlagged(data.flagged_rows || [])" in html


def test_route_sheets_use_the_route_card_and_the_rest_use_the_plain_card():
    html = _html()
    assert "const routeCardHtml = (r) =>" in html      # 経路・差額つきのカード
    assert "const plainCardHtml = (r) =>" in html      # 列の顔ぶれが違うシート用
    route = re.findall("'([^']+)'", _between(html, "const ROUTE_SHEETS = [", "];"))
    assert route == ["定期代突合", "実費突合"]


def test_hint_no_longer_sends_the_user_to_the_xlsx_for_sheets_now_on_screen():
    # 4シートが画面に出るようになったので、旧案内文が残っていると嘘になる
    assert "マスタ更新漏れ・上限超過・申請なしの明細はxlsxで確認してください" not in _html()
