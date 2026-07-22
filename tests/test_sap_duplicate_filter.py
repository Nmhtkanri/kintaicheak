# -*- coding: utf-8 -*-
"""sap_duplicate_filter のテスト ＋ run_keihi_integration へのSAP重複除外統合テスト"""
import csv

import pytest
from openpyxl import load_workbook

from services.sap_duplicate_filter import (
    collect_past_files,
    parse_past_inputs,
    run_sap_dedup,
)
from services.keihi_summary import run_keihi_integration

SAP_HEADERS = ["姓", "名", "費用合計", "業者名", "費用エントリ日", "説明", "承認日",
               "事業単位", "コストセンター", "通貨", "費用シート ID", "勤務地", "ステータス"]


def _sap_row(sheet_id, sei="山田", mei="太郎", total="1000"):
    return [sei, mei, total, "業者A", "2026/06/01", "説明X", "2026/06/02",
            "BU", "CC1", "JPY", sheet_id, "東京", "承認済"]


def _write_sap_csv(path, rows, headers=None, encoding="cp932"):
    with open(path, "w", encoding=encoding, newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(headers if headers is not None else SAP_HEADERS)
        w.writerows(rows)
    return path


class _NoApiClient:
    """roster 構築の fetch_active_employees を失敗させてオフライン続行させるダミー"""


class TestRunSapDedup:
    def test_basic_dedup(self, tmp_path):
        past = _write_sap_csv(tmp_path / "past.csv", [_sap_row("EXP-1"), _sap_row("EXP-2")])
        cur = _write_sap_csv(tmp_path / "current.csv",
                             [_sap_row("EXP-2"), _sap_row("EXP-3"), _sap_row("EXP-4")])
        out = tmp_path / "out"
        r = run_sap_dedup([past], cur, out)
        assert r.removed_row_count == 1
        assert r.kept_row_count == 2
        assert r.remaining_match_count == 0
        assert r.removed_rows[0]["照合キー"] == "EXP-2"
        assert "past.csv" in r.removed_rows[0]["重複元ファイル"]
        # 除外済みCSVには EXP-3 / EXP-4 だけが残る
        with open(r.clean_path, encoding="cp932", newline="") as f:
            rows = list(csv.DictReader(f))
        assert [row["費用シート ID"] for row in rows] == ["EXP-3", "EXP-4"]
        # 除外一覧CSVには付加列がある
        with open(r.removed_path, encoding="cp932", newline="") as f:
            removed = list(csv.DictReader(f))
        assert removed[0]["除外理由"].startswith("過去SAP経費と")
        assert removed[0]["重複元件数"] == "1"

    def test_empty_id_never_removed(self, tmp_path):
        past = _write_sap_csv(tmp_path / "past.csv", [_sap_row(""), _sap_row("EXP-1")])
        cur = _write_sap_csv(tmp_path / "current.csv", [_sap_row(""), _sap_row("EXP-9")])
        r = run_sap_dedup([past], cur, tmp_path / "out")
        assert r.removed_row_count == 0
        assert r.kept_row_count == 2

    def test_multi_line_sheet_all_removed(self, tmp_path):
        """1つの費用シートに複数明細 → そのIDの当月行はすべて除外"""
        past = _write_sap_csv(tmp_path / "past.csv", [_sap_row("EXP-5")])
        cur = _write_sap_csv(tmp_path / "current.csv",
                             [_sap_row("EXP-5", total="100"), _sap_row("EXP-5", total="200"),
                              _sap_row("EXP-6")])
        r = run_sap_dedup([past], cur, tmp_path / "out")
        assert r.removed_row_count == 2
        assert r.kept_row_count == 1

    def test_id_whitespace_stripped(self, tmp_path):
        past = _write_sap_csv(tmp_path / "past.csv", [_sap_row(" EXP-7 ")])
        cur = _write_sap_csv(tmp_path / "current.csv", [_sap_row("EXP-7")])
        r = run_sap_dedup([past], cur, tmp_path / "out")
        assert r.removed_row_count == 1

    def test_folder_recursive_and_generated_skipped(self, tmp_path):
        """フォルダ指定でサブフォルダも走査し、生成物CSVは過去として読まない"""
        pastdir = tmp_path / "R8年"
        sub = pastdir / "6月" / "経費確認精査データ"
        sub.mkdir(parents=True)
        _write_sap_csv(sub / "SAP_経費月次出力 (3).csv", [_sap_row("EXP-A")])
        # 生成物（読み飛ばし対象）: もし読まれたら EXP-B も除外されテストが落ちる
        _write_sap_csv(sub / "SAP_経費月次出力 (3)_重複除外済.csv", [_sap_row("EXP-B")])
        _write_sap_csv(sub / "経費統合一覧表_SAP除外一覧.csv", [_sap_row("EXP-C")])
        cur = _write_sap_csv(tmp_path / "current.csv",
                             [_sap_row("EXP-A"), _sap_row("EXP-B"), _sap_row("EXP-C")])
        r = run_sap_dedup([pastdir], cur, tmp_path / "out")
        assert r.past_file_count == 1
        assert r.removed_row_count == 1
        assert {row["照合キー"] for row in r.removed_rows} == {"EXP-A"}

    def test_current_csv_not_matched_against_itself(self, tmp_path):
        """当月CSVが過去フォルダ内にあっても自分自身とは突合しない"""
        folder = tmp_path / "data"
        folder.mkdir()
        cur = _write_sap_csv(folder / "current.csv", [_sap_row("EXP-X")])
        with pytest.raises(ValueError, match="過去SAP CSVが1つも"):
            collect_past_files([folder], cur, recursive=True)

    def test_missing_id_column(self, tmp_path):
        headers = [h for h in SAP_HEADERS if h != "費用シート ID"]
        past = _write_sap_csv(tmp_path / "past.csv",
                              [[c for c in _sap_row("EXP-1")][:10] + ["東京", "承認済"]],
                              headers=headers)
        cur = _write_sap_csv(tmp_path / "current.csv", [_sap_row("EXP-1")])
        with pytest.raises(ValueError, match="必要な列がありません"):
            run_sap_dedup([past], cur, tmp_path / "out")

    def test_utf8_bom_csv(self, tmp_path):
        past = _write_sap_csv(tmp_path / "past.csv", [_sap_row("EXP-1")], encoding="utf-8-sig")
        cur = _write_sap_csv(tmp_path / "current.csv",
                             [_sap_row("EXP-1"), _sap_row("EXP-2")], encoding="utf-8-sig")
        r = run_sap_dedup([past], cur, tmp_path / "out")
        assert r.removed_row_count == 1

    def test_output_stem(self, tmp_path):
        past = _write_sap_csv(tmp_path / "past.csv", [_sap_row("EXP-1")])
        cur = _write_sap_csv(tmp_path / "current.csv", [_sap_row("EXP-2")])
        r = run_sap_dedup([past], cur, tmp_path / "out", output_stem="経費統合一覧表_2026年07月")
        assert r.clean_path.name == "経費統合一覧表_2026年07月_SAP重複除外済.csv"
        assert r.removed_path.name == "経費統合一覧表_2026年07月_SAP除外一覧.csv"


class TestParsePastInputs:
    def test_semicolon_and_newline(self):
        got = parse_past_inputs('Y:\\a.csv; "Y:\\b フォルダ"\nY:\\c.csv;')
        assert got == ["Y:\\a.csv", "Y:\\b フォルダ", "Y:\\c.csv"]

    def test_empty(self):
        assert parse_past_inputs("") == []
        assert parse_past_inputs("  \n ; ") == []


class TestKeihiIntegrationWithSapDedup:
    def _run(self, tmp_path, sap_past_inputs):
        past = _write_sap_csv(tmp_path / "past.csv", [_sap_row("EXP-1")])
        cur = _write_sap_csv(tmp_path / "current.csv",
                             [_sap_row("EXP-1"), _sap_row("EXP-2"), _sap_row("EXP-3")])
        out = tmp_path / "integrated.xlsx"
        res = run_keihi_integration(
            output_path=out, sap_csv=cur,
            route_check=False, classify=False,
            sap_past_inputs=[str(past)] if sap_past_inputs == "auto" else sap_past_inputs,
            client=_NoApiClient(),
            log_func=lambda m: None,
        )
        return res, out

    def test_dedup_applied_before_transform(self, tmp_path):
        res, out = self._run(tmp_path, "auto")
        assert res.ok is True
        assert res.sap_dedup["removed_rows"] == 1
        assert res.sap_dedup["kept_rows"] == 2
        # 統合一覧表にはEXP-1を除いた2行だけが載る
        assert res.source_counts["sap"] == 2
        # Excel に「SAP重複除外」シートがあり、除外行が載っている
        wb = load_workbook(out)
        assert "SAP重複除外" in wb.sheetnames
        ws = wb["SAP重複除外"]
        headers = [c.value for c in ws[1]]
        assert "費用シート ID" in headers and "除外理由" in headers
        keys = [ws.cell(row=r, column=headers.index("照合キー") + 1).value
                for r in range(2, ws.max_row + 1)]
        assert keys == ["EXP-1"]
        # 除外済みCSV・除外一覧CSVが出力フォルダに残る
        assert (tmp_path / "integrated_SAP重複除外済.csv").exists()
        assert (tmp_path / "integrated_SAP除外一覧.csv").exists()

    def test_without_past_behaves_as_before(self, tmp_path):
        res, out = self._run(tmp_path, None)
        assert res.ok is True
        assert res.sap_dedup == {}
        assert res.source_counts["sap"] == 3
        wb = load_workbook(out)
        assert "SAP重複除外" not in wb.sheetnames

    def test_dedup_failure_aborts(self, tmp_path):
        cur = _write_sap_csv(tmp_path / "current.csv", [_sap_row("EXP-1")])
        res = run_keihi_integration(
            output_path=tmp_path / "integrated.xlsx", sap_csv=cur,
            route_check=False, classify=False,
            sap_past_inputs=[str(tmp_path / "存在しないフォルダ")],
            client=_NoApiClient(),
            log_func=lambda m: None,
        )
        assert res.ok is False
        assert "SAP重複除外に失敗" in res.error
