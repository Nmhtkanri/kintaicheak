# -*- coding: utf-8 -*-
"""jinjer給与へ投入した経費インポートの台帳と、追加投入のマージ。

投入後にイレギュラー経費が出たとき、統合一覧表を作り直さずに
「前回分＋追加分」で再投入するための土台。台帳が狂うと投入内容が狂うので、
同時書き込みの守り（SAP台帳と同じSHA-256比較）を重点的に固める。
"""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.keihi_import_ledger import (  # noqa: E402
    LEDGER_COLUMNS,
    ImportLedger,
    LedgerConflictError,
    load_ledger,
    merge_addon,
    month_rows,
    months_summary,
    parse_import_csv,
    record_submission,
    replace_month,
    resolve_names,
    save_ledger,
)
from services.keihi_payroll_import import IMPORT_HEADERS, render_import_csv  # noqa: E402


def _rows(*specs):
    """(社員番号, 氏名, 非課税通勤費, 社保調整) から投入行を作る。"""
    return [{"社員番号": e, "氏名": n, "非課税通勤費": t, "社保調整": s}
            for e, n, t, s in specs]


def _saved(tmp_path, month="2026-07", rows=None, user="谷津晴香"):
    led = ImportLedger(path=tmp_path / "投入台帳.csv", rows=[])
    replace_month(led, month, rows or _rows(("2020001", "山田 太郎", 18570, 0)),
                  "経費統合一覧表_jinjerインポート.csv", "44450", "1", user=user)
    save_ledger(led)
    return led


# ----------------------------------------------------------------------
# 保存と読み込み
# ----------------------------------------------------------------------

class TestSaveLoad:
    def test_round_trip_keeps_columns_and_amounts(self, tmp_path):
        led = _saved(tmp_path, rows=_rows(("2020001", "山田 太郎", 18570, -506)))
        again = load_ledger(led.path)
        assert [r["社員番号"] for r in again.rows] == ["2020001"]
        row = again.rows[0]
        assert row["対象月"] == "2026-07"
        assert row["非課税通勤費"] == "18570"
        assert row["社保調整"] == "-506"          # 控除のマイナスも保つ
        assert row["投入結果"] == "成功"
        assert row["投入者"] == "谷津晴香"
        assert row["元CSV"] == "経費統合一覧表_jinjerインポート.csv"

    def test_written_file_is_bom_utf8_with_crlf(self, tmp_path):
        led = _saved(tmp_path)
        raw = led.path.read_bytes()
        assert raw.startswith(b"\xef\xbb\xbf")     # Excelで開いて文字化けしない
        assert b"\r\n" in raw
        header = raw.decode("utf-8-sig").split("\r\n")[0].split(",")
        assert header == LEDGER_COLUMNS

    def test_missing_file_gives_an_empty_ledger(self, tmp_path):
        led = load_ledger(tmp_path / "まだ無い.csv")
        assert led.rows == [] and led.months() == []

    def test_broken_file_raises_instead_of_looking_empty(self, tmp_path):
        """空の台帳として黙って受け入れると、追加投入で前回分が消える。"""
        p = tmp_path / "別物.csv"
        p.write_text("名前,金額\nテスト,100\n", encoding="utf-8-sig")
        with pytest.raises(ValueError, match="対象月"):
            load_ledger(p)

    def test_backup_is_taken_before_overwriting(self, tmp_path):
        led = _saved(tmp_path)
        replace_month(led, "2026-08", _rows(("2020002", "佐藤 花子", 1000, 0)),
                      "x.csv", "44450", "1")
        bak = save_ledger(led)
        assert bak is not None and bak.exists()
        assert bak.parent.name == "_台帳バックアップ"

    def test_months_summary_is_newest_first(self, tmp_path):
        led = _saved(tmp_path, month="2026-07")
        replace_month(led, "2026-08", _rows(("2020002", "佐藤 花子", 1000, 0)),
                      "x.csv", "44450", "0")
        save_ledger(led)
        got = months_summary(load_ledger(led.path))
        assert [m["month"] for m in got] == ["2026-08", "2026-07"]
        assert got[0]["last_result"] == "タイムアウト（jinjer画面で要確認）"
        assert got[1]["rows"] == 1


# ----------------------------------------------------------------------
# 同時書き込み（谷津さんと平良さんの操作が重なっても記録を消さない）
# ----------------------------------------------------------------------

class TestConcurrentWrite:
    def test_overwrite_after_someone_else_saved_is_blocked(self, tmp_path):
        a = _saved(tmp_path)                       # 谷津さんが読み込んだ状態
        b = load_ledger(a.path)                    # 平良さんが同じ台帳を読み込む

        replace_month(b, "2026-08", _rows(("2020002", "佐藤 花子", 1000, 0)),
                      "b.csv", "44450", "1", user="平良菜津子")
        save_ledger(b)                             # 平良さんが先に保存

        replace_month(a, "2026-09", _rows(("2020003", "鈴木 次郎", 2000, 0)),
                      "a.csv", "44450", "1", user="谷津晴香")
        with pytest.raises(LedgerConflictError, match="他の人によって更新"):
            save_ledger(a)

        again = load_ledger(a.path)
        assert again.months() == ["2026-08", "2026-07"]   # 2026-09 は書かれていない
        assert month_rows(again, "2026-08")[0]["投入者"] == "平良菜津子"

    def test_force_overwrites(self, tmp_path):
        a = _saved(tmp_path)
        b = load_ledger(a.path)
        replace_month(b, "2026-08", _rows(("2020002", "佐藤 花子", 1000, 0)),
                      "b.csv", "44450", "1")
        save_ledger(b)
        save_ledger(a, force=True)
        assert load_ledger(a.path).months() == ["2026-07"]

    def test_sequential_saves_by_the_same_holder(self, tmp_path):
        """保存後にシグネチャを更新するので、同じ画面から続けて保存できる。"""
        led = _saved(tmp_path)
        replace_month(led, "2026-08", _rows(("2020002", "佐藤 花子", 1000, 0)),
                      "b.csv", "44450", "1")
        save_ledger(led)
        assert load_ledger(led.path).months() == ["2026-08", "2026-07"]


# ----------------------------------------------------------------------
# 投入済みCSVの読み取り（列の並びはテンプレ次第で変わる）
# ----------------------------------------------------------------------

class TestParseImportCsv:
    def test_reads_the_canonical_layout(self):
        rows, warns = parse_import_csv(render_import_csv(
            [{"社員番号": "2020001", "氏名": "山田 太郎", "非課税通勤費": 18570,
              "社保調整": -506}]))
        assert warns == []
        assert rows[0]["社員番号"] == "2020001"
        assert rows[0]["非課税通勤費"] == 18570
        assert rows[0]["社保調整"] == -506
        assert rows[0]["立替金"] == 0            # 未指定の項目は0で埋める

    def test_follows_a_reordered_template_header(self):
        """テンプレの並びが違っても、見出しから正規のキーへ引き直せる。"""
        header = ["従業員番号", "名前", "", "社保調整", "非課税通勤費", "未知の列"]
        rows, warns = parse_import_csv(render_import_csv(
            [{"社員番号": "2020001", "氏名": "山田 太郎", "非課税通勤費": 18570,
              "社保調整": -506}], template_header=header))
        assert warns == []
        assert rows[0]["社員番号"] == "2020001" and rows[0]["氏名"] == "山田 太郎"
        assert (rows[0]["非課税通勤費"], rows[0]["社保調整"]) == (18570, -506)

    def test_captures_telework_allowance_column(self):
        header = list(IMPORT_HEADERS) + ["テレワーク手当"]
        rows, _w = parse_import_csv(render_import_csv(
            [{"社員番号": "2020001", "氏名": "山田 太郎", "テレワーク手当": 3000}],
            template_header=header))
        assert rows[0]["テレワーク手当"] == 3000

    def test_handles_quoted_names_with_commas(self):
        rows, _w = parse_import_csv(render_import_csv(
            [{"社員番号": "2020001", "氏名": "山田, 太郎", "非課税通勤費": 100}]))
        assert rows[0]["氏名"] == "山田, 太郎"

    def test_reports_a_csv_without_employee_column(self):
        rows, warns = parse_import_csv("金額\r\n100\r\n".encode("cp932"))
        assert rows == []
        assert warns and "社員番号" in warns[0]


# ----------------------------------------------------------------------
# 月の入れ替えと記録
# ----------------------------------------------------------------------

class TestReplaceAndRecord:
    def test_replace_month_touches_only_that_month(self, tmp_path):
        led = _saved(tmp_path, month="2026-07")
        replace_month(led, "2026-08", _rows(("2020002", "佐藤 花子", 1000, 0)),
                      "b.csv", "44450", "1")
        replace_month(led, "2026-07", _rows(("2020009", "新 太郎", 500, 0)),
                      "c.csv", "44450", "1")
        assert [r["社員番号"] for r in month_rows(led, "2026-07")] == ["2020009"]
        assert [r["社員番号"] for r in month_rows(led, "2026-08")] == ["2020002"]

    @pytest.mark.parametrize("status,expected", [
        ("1", "成功"), ("0", "タイムアウト（jinjer画面で要確認）")])
    def test_status_maps_to_a_readable_result(self, tmp_path, status, expected):
        led = ImportLedger(path=tmp_path / "台帳.csv", rows=[])
        replace_month(led, "2026-07", _rows(("2020001", "山田 太郎", 1, 0)),
                      "a.csv", "44450", status)
        assert led.rows[0]["投入結果"] == expected

    def test_record_submission_stores_the_submitted_bytes(self, tmp_path):
        p = tmp_path / "台帳.csv"
        csv_bytes = render_import_csv(
            [{"社員番号": "2020001", "氏名": "山田 太郎", "非課税通勤費": 18570}])
        ok, warns = record_submission(p, "2026-07", csv_bytes, "投入.csv", "44450", "1")
        assert ok is True and warns == []
        assert month_rows(load_ledger(p), "2026-07")[0]["非課税通勤費"] == "18570"

    def test_record_submission_retries_once_after_a_conflict(self, tmp_path, monkeypatch):
        """保存の瞬間に他の人が書いていたら、読み直して1回だけやり直す。"""
        import services.keihi_import_ledger as mod
        p = tmp_path / "台帳.csv"
        record_submission(p, "2026-06", render_import_csv(
            [{"社員番号": "2020009", "氏名": "前月 太郎", "非課税通勤費": 1}]),
            "a.csv", "44450", "1")

        real_save, calls = mod.save_ledger, {"n": 0}

        def flaky(ledger, force=False):
            calls["n"] += 1
            if calls["n"] == 1:
                raise LedgerConflictError("他の人によって更新されています")
            return real_save(ledger, force)

        monkeypatch.setattr(mod, "save_ledger", flaky)
        ok, warns = record_submission(p, "2026-07", render_import_csv(
            [{"社員番号": "2020001", "氏名": "山田 太郎", "非課税通勤費": 18570}]),
            "b.csv", "44450", "1")
        assert ok is True and warns == [] and calls["n"] == 2
        assert load_ledger(p).months() == ["2026-07", "2026-06"]   # 前月分も残る

    def test_record_submission_warns_instead_of_raising_on_repeated_conflict(
            self, tmp_path, monkeypatch):
        """投入自体は終わっているので、記録の失敗で例外を投げない（警告で伝える）。"""
        import services.keihi_import_ledger as mod
        monkeypatch.setattr(mod, "save_ledger", lambda *a, **k: (_ for _ in ()).throw(
            LedgerConflictError("他の人によって更新されています")))
        ok, warns = record_submission(
            tmp_path / "台帳.csv", "2026-07",
            render_import_csv([{"社員番号": "2020001", "氏名": "山田 太郎"}]),
            "b.csv", "44450", "1")
        assert ok is False
        assert any("追加投入" in w for w in warns)

    def test_record_submission_warns_when_the_path_is_unwritable(self, tmp_path):
        ok, warns = record_submission(
            tmp_path, "2026-07",           # ディレクトリを指定＝書けない
            render_import_csv([{"社員番号": "2020001", "氏名": "山田 太郎"}]),
            "b.csv", "44450", "1")
        assert ok is False and warns


# ----------------------------------------------------------------------
# 追加投入のマージ
# ----------------------------------------------------------------------

class TestMergeAddon:
    def _base(self):
        return [{"社員番号": "2020001", "氏名": "山田 太郎", "非課税通勤費": "18570"},
                {"社員番号": "2020002", "氏名": "佐藤 花子", "立替金": "3000"}]

    def test_adds_to_an_existing_person(self):
        rows, preview, warns = merge_addon(
            self._base(), {"その他手当": {"2020001": 5000}})
        assert warns == []
        by = {r["社員番号"]: r for r in rows}
        assert by["2020001"]["非課税通勤費"] == 18570     # 前回分は残る
        assert by["2020001"]["その他手当"] == 5000
        pv = {p["社員番号"]: p for p in preview}
        assert pv["2020001"]["区分"] == "追加あり"
        assert "その他手当 +5,000円" in pv["2020001"]["追加内容"]
        assert pv["2020002"]["区分"] == "前回のみ"

    def test_adds_a_person_who_was_not_in_the_previous_batch(self):
        rows, preview, _w = merge_addon(self._base(), {"現物支給": {"2020009": 3000}})
        assert [r["社員番号"] for r in rows][-1] == "2020009"
        assert preview[-1]["区分"] == "新規"
        assert preview[-1]["現物支給"] == 3000

    def test_negative_amount_cancels(self):
        """取り消しはマイナスで入れる（jinjerは上書きなので0でも送る意味がある）。"""
        rows, preview, _w = merge_addon(
            [{"社員番号": "2020001", "氏名": "山田 太郎", "その他手当": "5000"}],
            {"その他手当": {"2020001": -5000}})
        assert rows[0]["その他手当"] == 0
        assert preview[0]["区分"] == "追加あり"

    def test_sums_the_same_person_appearing_twice_in_the_base(self):
        rows, _p, _w = merge_addon(
            [{"社員番号": "2020001", "氏名": "山田 太郎", "非課税通勤費": "100"},
             {"社員番号": "2020001", "氏名": "山田 太郎", "非課税通勤費": "200"}], {})
        assert len(rows) == 1 and rows[0]["非課税通勤費"] == 300

    def test_rejects_an_item_that_cannot_be_submitted(self):
        rows, _p, warns = merge_addon(self._base(), {"謎の手当": {"2020001": 100}})
        assert any("投入できる項目ではない" in w for w in warns)
        assert all("謎の手当" not in r for r in rows)


# ----------------------------------------------------------------------
# 氏名の解決（jinjer APIのレート制限があるので必要なときだけ取る）
# ----------------------------------------------------------------------

class TestResolveNames:
    def _ledger(self):
        return ImportLedger(path=None, rows=[
            {"対象月": "2026-06", "社員番号": "2020001", "氏名": "山田 太郎"},
            {"対象月": "2026-07", "社員番号": "2020002", "氏名": ""},
        ])

    def test_finds_names_across_months_without_calling_the_api(self):
        called = []
        names, warns = resolve_names(self._ledger(), ["2020001"],
                                     roster_fetch=lambda: called.append(1) or {})
        assert names["2020001"] == "山田 太郎"
        assert called == [] and warns == []

    def test_calls_the_api_only_once_for_unknown_ids(self):
        called = []

        def fetch():
            called.append(1)
            return {"2020009": "新 太郎", "2020002": "佐藤 花子"}

        names, warns = resolve_names(self._ledger(), ["2020001", "2020002", "2020009"],
                                     roster_fetch=fetch)
        assert len(called) == 1
        assert names["2020009"] == "新 太郎" and names["2020002"] == "佐藤 花子"
        assert warns == []

    def test_degrades_to_a_warning_when_the_api_fails(self):
        def boom():
            raise RuntimeError("429 Too Many Requests")

        names, warns = resolve_names(self._ledger(), ["2020009"], roster_fetch=boom)
        assert names.get("2020009") is None
        assert any("429" in w for w in warns)
        assert any("台帳に無い社員番号" in w for w in warns)

    def test_warns_about_unknown_ids_without_a_fetcher(self):
        names, warns = resolve_names(self._ledger(), ["2020009"])
        assert names.get("2020009") is None
        assert any("2020009" in w for w in warns)
