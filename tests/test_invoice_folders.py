# -*- coding: utf-8 -*-
"""請求書モードのフォルダ設定（一覧・保存・状態チェック）のテスト。

共有フォルダも Excel も使わず、tmp_path 上のダミーで確かめる。
"""

import csv
import datetime as dt
import io
from pathlib import Path

import pytest

from services import invoice_folders as folders
from services.invoice_pdf import load_settings


HEADER = ("対象,取引先,氏名,請求書Excel,シート名,勤怠フォルダ,勤怠ファイル,"
          "出力フォルダ,出力ファイル名")


def _write(path: Path, text: str, encoding: str = "utf-8-sig") -> Path:
    path.write_text(text, encoding=encoding, newline="")
    return path


def _rows(path: Path) -> list[dict]:
    text = path.read_text(encoding="utf-8-sig")
    return list(csv.DictReader(io.StringIO(text)))


# ----------------------------------------------------------------------
# 読み込み
# ----------------------------------------------------------------------

def test_対象列が無い古いCSVでも全員が対象になる(tmp_path):
    csv_path = _write(tmp_path / "s.csv",
                      "取引先,氏名,請求書Excel,シート名,勤怠フォルダ,勤怠ファイル,出力フォルダ,出力ファイル名\r\n"
                      "A社,細川昌広,a.xls,{YY}年{M}月,F,*.pdf,O,out.pdf\r\n")
    people = folders.load_people(csv_path)
    assert len(people) == 1
    assert people[0]["対象"] is True
    assert people[0]["氏名"] == "細川昌広"


def test_対象0の行は読めるが提出用PDF作成からは外れる(tmp_path):
    csv_path = _write(tmp_path / "s.csv", HEADER + "\r\n"
                      "1,A社,細川昌広,a.xls,S,F,*.pdf,O,out.pdf\r\n"
                      "0,A社,大村賢治,b.xls,S,F,*.pdf,O,out.pdf\r\n")
    # 画面には両方出す（チェックを外した状態で見せるため）
    people = folders.load_people(csv_path)
    assert [p["氏名"] for p in people] == ["細川昌広", "大村賢治"]
    assert [p["対象"] for p in people] == [True, False]
    # 実際の作成は対象の人だけ
    assert [r["氏名"] for r in load_settings(csv_path)] == ["細川昌広"]


def test_cp932のCSVも読める(tmp_path):
    csv_path = _write(tmp_path / "s.csv", HEADER + "\r\n"
                      "1,A社,細川昌広,a.xls,S,F,*.pdf,O,out.pdf\r\n", encoding="cp932")
    assert folders.load_people(csv_path)[0]["取引先"] == "A社"


def test_対象フォルダCSVの取引先はフォルダ名から補う(tmp_path):
    csv_path = _write(tmp_path / "r.csv",
                      "対象,フォルダパス\r\n"
                      '1,"Z:\\NetMarks以外(常駐）\\アクシスITパートナーズ（細川・渡会）"\r\n'
                      "0,Z:\\NetMarks以外(常駐）\\UBS（田中）\r\n")
    roots = folders.load_roots(csv_path)
    assert [r["取引先"] for r in roots] == ["アクシスITパートナーズ", "UBS"]
    assert [r["対象"] for r in roots] == [True, False]


def test_対象フォルダCSVが無ければ既定の28件を出す(tmp_path):
    roots = folders.load_roots(tmp_path / "ない.csv")
    assert len(roots) == 28
    assert all(r["対象"] for r in roots)


# ----------------------------------------------------------------------
# 保存
# ----------------------------------------------------------------------

def _people_row(**over):
    row = {"対象": True, "取引先": "A社", "氏名": "細川昌広", "請求書Excel": "a.xls",
           "シート名": "{YY}年{M}月", "勤怠フォルダ": "F", "勤怠ファイル": "*.pdf",
           "出力フォルダ": "O", "出力ファイル名": "out.pdf"}
    row.update(over)
    return row


def test_保存すると対象列つきで書き戻りバックアップが残る(tmp_path):
    settings = _write(tmp_path / "s.csv", HEADER + "\r\n"
                      "1,A社,細川昌広,a.xls,S,F,*.pdf,O,out.pdf\r\n")
    roots = _write(tmp_path / "r.csv", "対象,フォルダパス\r\n1,Z:\\A\r\n")

    result = folders.save_all(
        settings, roots,
        people=[_people_row(), _people_row(氏名="大村賢治", 対象=False)],
        roots=[{"対象": True, "取引先": "A社", "フォルダパス": "Z:\\A"}],
        signatures={"people": folders.signature(settings),
                    "roots": folders.signature(roots)})

    assert result["people_saved"] == 2
    saved = _rows(settings)
    assert [r["対象"] for r in saved] == ["1", "0"]
    assert saved[0]["氏名"] == "細川昌広"
    assert len(result["backups"]) == 2
    assert all(Path(b).exists() for b in result["backups"])


def test_知らない列を消さずに書き戻す(tmp_path):
    settings = _write(tmp_path / "s.csv", HEADER + ",メモ\r\n"
                      "1,A社,細川昌広,a.xls,S,F,*.pdf,O,out.pdf,来期で終了\r\n")
    roots = _write(tmp_path / "r.csv", "対象,フォルダパス\r\n1,Z:\\A\r\n")
    people = folders.load_people(settings)
    assert people[0]["_extra"] == {"メモ": "来期で終了"}

    folders.save_all(settings, roots, people=people,
                     roots=folders.load_roots(roots),
                     signatures={"people": folders.signature(settings),
                                 "roots": folders.signature(roots)})
    assert _rows(settings)[0]["メモ"] == "来期で終了"


def test_他の人が直していたら上書きせず止まる(tmp_path):
    settings = _write(tmp_path / "s.csv", HEADER + "\r\n"
                      "1,A社,細川昌広,a.xls,S,F,*.pdf,O,out.pdf\r\n")
    roots = _write(tmp_path / "r.csv", "対象,フォルダパス\r\n1,Z:\\A\r\n")
    stale = {"people": folders.signature(settings), "roots": folders.signature(roots)}

    _write(settings, HEADER + "\r\n1,A社,細川昌広,b.xls,S,F,*.pdf,O,out.pdf\r\n")

    with pytest.raises(folders.InvoiceFoldersConflict):
        folders.save_all(settings, roots, people=[_people_row()],
                         roots=[{"対象": True, "フォルダパス": "Z:\\A"}],
                         signatures=stale)
    # 中止したので中身は他の人が書いたまま
    assert _rows(settings)[0]["請求書Excel"] == "b.xls"


def test_空欄や重複は保存前に止める(tmp_path):
    settings = tmp_path / "s.csv"
    roots = tmp_path / "r.csv"
    with pytest.raises(folders.InvoiceFoldersError) as e:
        folders.save_all(settings, roots,
                         people=[_people_row(), _people_row()],
                         roots=[], signatures={})
    assert "同じ取引先・氏名" in str(e.value)

    with pytest.raises(folders.InvoiceFoldersError) as e:
        folders.save_all(settings, roots,
                         people=[_people_row(勤怠フォルダ="")], roots=[], signatures={})
    assert "勤怠フォルダが空です" in str(e.value)


def test_対象外の行は空欄でも保存できる(tmp_path):
    settings = tmp_path / "s.csv"
    roots = tmp_path / "r.csv"
    result = folders.save_all(settings, roots,
                              people=[_people_row(対象=False, 勤怠フォルダ="", 請求書Excel="")],
                              roots=[], signatures={})
    assert result["people_saved"] == 1


# ----------------------------------------------------------------------
# 今月の状態
# ----------------------------------------------------------------------

TARGET = dt.date(2026, 7, 1)


def test_勤怠PDFが1件なら通る(tmp_path):
    folder = tmp_path / "FY2026"
    folder.mkdir()
    (folder / "【26年7月業務実績表】細川様_v1.pdf").write_text("x", encoding="utf-8")
    state = folders.check_kintai(
        {"勤怠フォルダ": str(tmp_path / "FY{FY}"),
         "勤怠ファイル": "【{YY}年{M}月業務実績表】細川様*.pdf"}, TARGET)
    assert state["level"] == folders.LEVEL_OK
    assert state["text"] == "1件"


def test_勤怠PDFが複数なら候補として警告する(tmp_path):
    folder = tmp_path / "FY2026"
    folder.mkdir()
    for n in ("a", "b"):
        (folder / f"勤務実績表_大村賢治_202607_{n}.pdf").write_text("x", encoding="utf-8")
    state = folders.check_kintai(
        {"勤怠フォルダ": str(tmp_path / "FY{FY}"),
         "勤怠ファイル": "*勤務実績表_大村賢治_{YYYY}{MM}*.pdf"}, TARGET)
    assert state["level"] == folders.LEVEL_WARN
    assert state["text"] == "候補2件"


def test_勤怠フォルダが無ければ止まる(tmp_path):
    state = folders.check_kintai(
        {"勤怠フォルダ": str(tmp_path / "ない"), "勤怠ファイル": "*.pdf"}, TARGET)
    assert state["level"] == folders.LEVEL_STOP
    assert state["text"] == "フォルダなし"


def test_出力先に同名PDFがあれば警告する(tmp_path):
    out = tmp_path / "提出データ"
    out.mkdir()
    (out / "御請求書_202607.pdf").write_text("x", encoding="utf-8")
    row = {"出力フォルダ": str(out), "出力ファイル名": "御請求書_{YYYY}{MM}.pdf"}
    state = folders.check_output(row, TARGET)
    assert state["level"] == folders.LEVEL_WARN
    assert state["text"] == "同名あり"

    row["出力ファイル名"] = "御請求書_{YYYY}{MM}_2.pdf"
    assert folders.check_output(row, TARGET)["level"] == folders.LEVEL_OK


def test_会社フォルダは提出データの有無まで見る(tmp_path):
    company = tmp_path / "A社"
    (company / ",提出データ").mkdir(parents=True)
    assert folders.check_root({"フォルダパス": str(company)})["text"] == "提出データ1個"

    plain = tmp_path / "B社"
    plain.mkdir()
    state = folders.check_root({"フォルダパス": str(plain)})
    assert state["level"] == folders.LEVEL_INFO
    assert state["text"] == "提出データなし"

    assert folders.check_root(
        {"フォルダパス": str(tmp_path / "ない")})["level"] == folders.LEVEL_STOP


def test_対象外の行は状態を見に行かない(tmp_path):
    result = folders.check(
        "2026-07",
        people=[{"対象": False, "勤怠フォルダ": "Z:\\ないはず", "勤怠ファイル": "*.pdf",
                 "請求書Excel": "x.xls", "シート名": "S",
                 "出力フォルダ": "Z:\\ないはず", "出力ファイル名": "o.pdf"}],
        roots=[{"対象": False, "フォルダパス": "Z:\\ないはず"}],
        scope="all", read_sheets=False)
    assert result["people"][0]["kintai"]["level"] == folders.LEVEL_OFF
    assert result["roots"][0]["root"]["level"] == folders.LEVEL_OFF


def test_scopeで見る範囲を絞れる(tmp_path):
    people = [{"対象": True, "勤怠フォルダ": str(tmp_path), "勤怠ファイル": "*.pdf",
               "請求書Excel": "x.xls", "シート名": "S",
               "出力フォルダ": str(tmp_path), "出力ファイル名": "o.pdf"}]
    result = folders.check("2026-07", people=people, roots=[{"フォルダパス": str(tmp_path)}],
                           scope="kintai", read_sheets=False)
    assert "kintai" in result["people"][0]
    assert "excel" not in result["people"][0]
    assert result["roots"] == []


def test_対象月の形式が違えば断る():
    with pytest.raises(folders.InvoiceFoldersError):
        folders.check("2026/07", people=[], roots=[], read_sheets=False)
