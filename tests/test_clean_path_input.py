# -*- coding: utf-8 -*-
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import _clean_path_input


def test_strips_both_side_double_quotes():
    # Windows「パスのコピー」形式
    assert _clean_path_input('"Z:\\給与明細\\ERCSTS"') == "Z:\\給与明細\\ERCSTS"


def test_strips_one_side_double_quote():
    # 片側だけ残ったケース（従来は剥がれず失敗していた）
    assert _clean_path_input('"Z:\\a\\b.csv') == "Z:\\a\\b.csv"
    assert _clean_path_input('Z:\\a\\b.csv"') == "Z:\\a\\b.csv"


def test_strips_quotes_with_surrounding_spaces():
    assert _clean_path_input('   "Z:\\a\\b.csv"   ') == "Z:\\a\\b.csv"


def test_strips_multiple_double_quotes():
    assert _clean_path_input('""Z:\\a\\b.csv""') == "Z:\\a\\b.csv"


def test_matching_single_quotes_stripped():
    assert _clean_path_input("'Z:\\a\\b.csv'") == "Z:\\a\\b.csv"


def test_no_quotes_unchanged():
    assert _clean_path_input("Z:\\a\\b.csv") == "Z:\\a\\b.csv"


def test_apostrophe_in_path_preserved():
    # フォルダ名の途中のアポストロフィは壊さない
    assert _clean_path_input('"C:\\Users\\O\'Brien\\file.txt"') == "C:\\Users\\O'Brien\\file.txt"


def test_empty_and_none():
    assert _clean_path_input("") == ""
    assert _clean_path_input(None) == ""
