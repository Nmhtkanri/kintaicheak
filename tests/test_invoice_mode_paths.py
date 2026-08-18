from services import invoice_mode as im


def test_target_root_csv_preserves_fullwidth_parentheses(tmp_path):
    path = tmp_path / "targets.csv"
    expected = r"Z:\NetMarks以外(常駐）\A&A（新井）"
    path.write_text(f"対象,フォルダパス\n1,\"{expected}\"\n", encoding="utf-8-sig")
    assert im.load_target_roots(path) == [expected]
