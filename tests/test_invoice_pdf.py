"""請求書Excel＋勤怠PDF→提出用PDF（services/invoice_pdf）。

Excel COM を使う部分は実機依存なのでここでは触らず、
「何を綴じてどこへ出すか」を決めるところと、止める条件を確かめる。
"""
from datetime import date

import pytest

from services import invoice_pdf as ip


def test_fiscal_year_starts_in_april():
    assert ip.fiscal_year(date(2026, 4, 1)) == 2026
    assert ip.fiscal_year(date(2026, 7, 31)) == 2026
    assert ip.fiscal_year(date(2027, 3, 31)) == 2026, "1〜3月は前年度"
    assert ip.fiscal_year(date(2027, 4, 1)) == 2027


def test_expand_placeholders():
    target = date(2026, 7, 1)
    assert ip.expand("{YYYY}{MM}", target) == "202607"
    assert ip.expand("{YY}年{M}月", target) == "26年7月", "シート名は前ゼロなし"
    assert ip.expand("FY{FY}", target) == "FY2026"
    assert ip.expand("FY{FY}", date(2027, 2, 1)) == "FY2026"


def _settings(tmp_path, **over):
    row = {
        "取引先": "テスト社", "氏名": "山田太郎",
        "請求書Excel": str(tmp_path / "請求書_{FY}.xls"),
        "シート名": "{YY}年{M}月",
        "勤怠フォルダ": str(tmp_path / "kintai"),
        "勤怠ファイル": "勤務表_{YYYY}{MM}*.pdf",
        "出力フォルダ": str(tmp_path / "out"),
        "出力ファイル名": "請求書_{YYYY}{MM}.pdf",
    }
    row.update(over)
    return [row]


def _prepare(tmp_path, attendance=("勤務表_202607.pdf",)):
    (tmp_path / "請求書_2026.xls").write_text("x", encoding="utf-8")
    (tmp_path / "kintai").mkdir(exist_ok=True)
    for name in attendance:
        (tmp_path / "kintai" / name).write_text("x", encoding="utf-8")
    (tmp_path / "out").mkdir(exist_ok=True)


def test_plan_resolves_paths(tmp_path):
    _prepare(tmp_path)
    got = ip.plan("2026-07", _settings(tmp_path))[0]
    assert got.ok, got.errors
    assert got.sheet == "26年7月"
    assert got.attendance[0].name == "勤務表_202607.pdf"
    assert got.output.name == "請求書_202607.pdf"


def test_plan_stops_when_attendance_is_missing(tmp_path):
    _prepare(tmp_path, attendance=())
    got = ip.plan("2026-07", _settings(tmp_path))[0]
    assert not got.ok
    assert any("勤怠PDFが見つかりません" in e for e in got.errors)


def test_plan_stops_when_attendance_is_ambiguous(tmp_path):
    """★違う人の勤怠を綴じると事故になるので、候補が複数なら作らない。"""
    _prepare(tmp_path, attendance=("勤務表_202607.pdf", "勤務表_202607_捺印済.pdf"))
    got = ip.plan("2026-07", _settings(tmp_path))[0]
    assert not got.ok
    assert any("候補が複数" in e for e in got.errors)


def test_plan_never_overwrites_an_existing_pdf(tmp_path):
    """★提出済みのPDFを黙って置き換えない。"""
    _prepare(tmp_path)
    (tmp_path / "out" / "請求書_202607.pdf").write_text("既にある", encoding="utf-8")
    got = ip.plan("2026-07", _settings(tmp_path))[0]
    assert not got.ok
    assert any("上書きしません" in e for e in got.errors)


def test_plan_stops_when_workbook_is_missing(tmp_path):
    (tmp_path / "kintai").mkdir()
    (tmp_path / "kintai" / "勤務表_202607.pdf").write_text("x", encoding="utf-8")
    (tmp_path / "out").mkdir()
    got = ip.plan("2026-07", _settings(tmp_path))[0]
    assert not got.ok
    assert any("請求書Excelが見つかりません" in e for e in got.errors)


def test_bad_month_is_rejected(tmp_path):
    with pytest.raises(ip.InvoicePdfError):
        ip.plan("2026/07", _settings(tmp_path))


def _stub_excel(monkeypatch, tmp_path, notes=()):
    """Excel COM と PDF読み取りを外す（実機に依存させない）。"""
    def fake_export(workbook, sheet, out_pdf):
        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        out_pdf.write_text("dummy", encoding="utf-8")
        return out_pdf

    monkeypatch.setattr(ip, "export_sheet_to_pdf", fake_export)
    monkeypatch.setattr(ip, "check_dates", lambda *_: list(notes))
    monkeypatch.setattr(ip, "merge_pdfs",
                        lambda parts, out: (out.write_text("pdf", encoding="utf-8"), out)[1])


def test_dry_run_makes_nothing(tmp_path, monkeypatch):
    _prepare(tmp_path)
    _stub_excel(monkeypatch, tmp_path)
    res = ip.build("2026-07", _settings(tmp_path), work_dir=tmp_path / "work",
                   dry_run=True)
    assert res["made"] and not res["skipped"] and not res["needs_confirm"]
    assert not (tmp_path / "out" / "請求書_202607.pdf").exists()


def test_date_mismatch_asks_for_confirmation(tmp_path, monkeypatch):
    """★請求日が対象月とずれていたら、勝手に作らず人に確認してもらう。"""
    _prepare(tmp_path)
    _stub_excel(monkeypatch, tmp_path, notes=["請求日が 2026-06 になっています"])
    res = ip.build("2026-07", _settings(tmp_path), work_dir=tmp_path / "work",
                   dry_run=False)
    assert not res["made"]
    assert res["needs_confirm"][0]["氏名"] == "山田太郎"
    assert not (tmp_path / "out" / "請求書_202607.pdf").exists(), "確認前に作らない"


def test_force_builds_after_confirmation(tmp_path, monkeypatch):
    """人が確認したら、そのままの中身で作る。"""
    _prepare(tmp_path)
    _stub_excel(monkeypatch, tmp_path, notes=["請求日が 2026-06 になっています"])
    res = ip.build("2026-07", _settings(tmp_path), work_dir=tmp_path / "work",
                   dry_run=False, force=True)
    assert not res["needs_confirm"]
    assert res["made"][0]["確認事項"], "承知で作ったことを記録に残す"
    assert (tmp_path / "out" / "請求書_202607.pdf").exists()


def test_force_does_not_override_missing_material(tmp_path, monkeypatch):
    """★材料が揃わないものは force でも作らない（違う人の勤怠を綴じないため）。"""
    _prepare(tmp_path, attendance=())
    _stub_excel(monkeypatch, tmp_path)
    res = ip.build("2026-07", _settings(tmp_path), work_dir=tmp_path / "work",
                   dry_run=False, force=True)
    assert not res["made"] and res["skipped"]


def test_load_settings_ignores_blank_rows(tmp_path):
    path = tmp_path / "s.csv"
    path.write_text("取引先,氏名,請求書Excel\nテスト社,山田太郎,x.xls\n,,\n",
                    encoding="utf-8-sig")
    assert len(ip.load_settings(path)) == 1


def test_missing_settings_file_is_an_error(tmp_path):
    with pytest.raises(ip.InvoicePdfError):
        ip.load_settings(tmp_path / "no.csv")
