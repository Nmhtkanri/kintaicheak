# -*- coding: utf-8 -*-
"""sap_import_ledger のテスト（SAP経費 取込済み費用シート台帳）

実データ由来のシナリオを重視している:
  - 2026-08: 7月に取り込んだ費用シートが8月のSAP CSVに丸ごと再掲（93行/285,642円）→ 除外
  - 2026-07: 同じ費用シートIDのまま金額が訂正されて再発行（27件）→ 要確認
  - 山田さんの例: 同じID・同じ日・同じ金額でも業者名と説明だけ違う明細が5行並ぶ
"""
import csv

from services.sap_import_ledger import (
    ID_COLUMN,
    LEDGER_COLUMNS,
    STATUS_CONFIRMED,
    STATUS_PROVISIONAL,
    Ledger,
    append_provisional,
    can_write,
    classify_rows,
    confirm_month,
    default_import_month,
    load_ledger,
    row_key,
    save_ledger,
    unconfirm_month,
)


def _sap(sheet_id, sei="上原", mei="奏吾", date="2026/6/1", total="2096",
         vendor="京成、都営浅草、京急", desc="東松戸駅から平和島駅の往復"):
    return {ID_COLUMN: sheet_id, "姓": sei, "名": mei, "費用エントリ日": date,
            "費用合計": total, "業者名": vendor, "説明": desc}


def _ledger_row(row, month="2026-07", status=STATUS_CONFIRMED):
    out = {c: "" for c in LEDGER_COLUMNS}
    out.update(row)
    out["状態"] = status
    out["取込年月"] = month
    return out


def _make_ledger(tmp_path, rows):
    return Ledger(path=tmp_path / "台帳.csv", rows=list(rows))


class TestClassifyRows:
    def test_three_buckets(self, tmp_path):
        led = _make_ledger(tmp_path, [_ledger_row(_sap("EXP-1"))])
        cur = [
            _sap("EXP-1"),                          # 明細まで一致 → 除外
            _sap("EXP-1", total="2200"),            # IDだけ一致（金額訂正） → 要確認
            _sap("EXP-9"),                          # 台帳に無い → 取込
        ]
        plan = classify_rows(cur, led)
        assert len(plan.excluded_rows) == 1
        assert len(plan.review_rows) == 1
        assert len(plan.kept_rows) == 1
        assert plan.excluded_rows[0]["判定"] == "除外（支給済み）"
        assert plan.review_rows[0]["台帳の取込年月"] == "2026-07"
        assert plan.kept_rows[0][ID_COLUMN] == "EXP-9"
        assert plan.current_row_count == 3

    def test_only_confirmed_rows_are_used(self, tmp_path):
        """暫定行は判定に使わない（取り込む前に弾いてしまうと未払いになる）"""
        led = _make_ledger(tmp_path, [_ledger_row(_sap("EXP-1"), status=STATUS_PROVISIONAL)])
        plan = classify_rows([_sap("EXP-1")], led)
        assert len(plan.kept_rows) == 1
        assert len(plan.excluded_rows) == 0
        assert len(plan.review_rows) == 0

    def test_duplicate_detail_counted_not_setwise(self, tmp_path):
        """台帳に1行しかない明細が当月に2行あれば、除外は1行だけ（残り1行は取込）"""
        led = _make_ledger(tmp_path, [_ledger_row(_sap("EXP-1"))])
        plan = classify_rows([_sap("EXP-1"), _sap("EXP-1")], led)
        assert len(plan.excluded_rows) == 1
        assert len(plan.kept_rows) == 0
        # 2行目はIDが台帳にあるので「要確認」へ回る（黙って取り込まない）
        assert len(plan.review_rows) == 1

    def test_same_id_same_amount_different_vendor_are_distinct(self, tmp_path):
        """山田さんの実例: 同じID・同じ日・同じ金額でも業者名/説明が違えば別明細"""
        a = _sap("EXP-Y", sei="山田", mei="友輔", date="2026/5/1", total="5000",
                 vendor="市バス", desc="踊場駅→領家中学校前　片道")
        b = _sap("EXP-Y", sei="山田", mei="友輔", date="2026/5/1", total="5000",
                 vendor="ブルーライン線", desc="湘南台駅→踊場駅　片道")
        assert row_key(a) != row_key(b)
        led = _make_ledger(tmp_path, [_ledger_row(a)])
        plan = classify_rows([a, b], led)
        assert len(plan.excluded_rows) == 1
        assert len(plan.review_rows) == 1   # b は同じIDなので要確認へ

    def test_empty_sheet_id_always_kept(self, tmp_path):
        led = _make_ledger(tmp_path, [_ledger_row(_sap(""))])
        plan = classify_rows([_sap(""), _sap("  ")], led)
        assert len(plan.kept_rows) == 2
        assert len(plan.excluded_rows) == 0

    def test_date_and_amount_normalized(self, tmp_path):
        """2026/06/01 と 2026/6/1、"2,096" と "2096" は同じ明細"""
        led = _make_ledger(tmp_path, [_ledger_row(_sap("EXP-1", date="2026/06/01", total="2,096"))])
        plan = classify_rows([_sap("EXP-1", date="2026/6/1", total="2096")], led)
        assert len(plan.excluded_rows) == 1

    def test_empty_ledger_keeps_everything(self, tmp_path):
        plan = classify_rows([_sap("EXP-1"), _sap("EXP-2")], _make_ledger(tmp_path, []))
        assert len(plan.kept_rows) == 2
        assert plan.ledger_row_count == 0

    def test_tax_reissue_goes_to_review(self, tmp_path):
        """2026-07 実例: 当番手当が税抜2,500円 → 税込2,750円で再発行された27件"""
        led = _make_ledger(tmp_path, [
            _ledger_row(_sap("EXP-B", sei="砂子", mei="領吾", date="2026/5/3",
                             total="2500", vendor="顧客対応当番　16", desc="5月3日"), month="2026-06"),
        ])
        cur = [_sap("EXP-B", sei="砂子", mei="領吾", date="2026/5/3",
                    total="2750", vendor="顧客対応当番　16", desc="5月3日")]
        plan = classify_rows(cur, led)
        assert len(plan.review_rows) == 1
        assert len(plan.excluded_rows) == 0
        assert "再発行" in plan.review_rows[0]["理由"]
        assert plan.review_rows[0]["台帳の取込年月"] == "2026-06"


class TestAppendAndConfirm:
    def test_append_then_confirm(self, tmp_path):
        led = _make_ledger(tmp_path, [])
        n = append_provisional(led, [_sap("EXP-1"), _sap("EXP-2")], "2026-08",
                               "経費月次出力 (10).csv", user="谷津晴香")
        assert n == 2
        assert led.provisional_months == ["2026-08"]
        assert led.confirmed_rows == []
        assert led.rows[0]["元CSV"] == "経費月次出力 (10).csv"
        assert led.rows[0]["記録者"] == "谷津晴香"

        assert confirm_month(led, "2026-08", user="谷津晴香") == 2
        assert led.provisional_months == []
        assert len(led.confirmed_rows) == 2
        assert led.rows[0]["確定者"] == "谷津晴香"
        assert led.rows[0]["確定日時"]

    def test_rerun_same_month_replaces_provisional(self, tmp_path):
        """同じ月をやり直しても暫定行が二重に積まれない"""
        led = _make_ledger(tmp_path, [])
        append_provisional(led, [_sap("EXP-1"), _sap("EXP-2")], "2026-08", "a.csv")
        append_provisional(led, [_sap("EXP-3")], "2026-08", "b.csv")
        assert len(led.rows) == 1
        assert led.rows[0][ID_COLUMN] == "EXP-3"

    def test_rerun_does_not_touch_confirmed(self, tmp_path):
        led = _make_ledger(tmp_path, [])
        append_provisional(led, [_sap("EXP-1")], "2026-07", "a.csv")
        confirm_month(led, "2026-07")
        append_provisional(led, [_sap("EXP-2")], "2026-08", "b.csv")
        assert len(led.rows) == 2
        assert len(led.confirmed_rows) == 1

    def test_rows_without_sheet_id_are_not_recorded(self, tmp_path):
        led = _make_ledger(tmp_path, [])
        assert append_provisional(led, [_sap(""), _sap("EXP-1")], "2026-08", "a.csv") == 1

    def test_unconfirm(self, tmp_path):
        led = _make_ledger(tmp_path, [])
        append_provisional(led, [_sap("EXP-1")], "2026-08", "a.csv")
        confirm_month(led, "2026-08")
        assert unconfirm_month(led, "2026-08") == 1
        assert led.provisional_months == ["2026-08"]
        assert led.rows[0]["確定日時"] == ""

    def test_default_import_month(self):
        import datetime as dt
        assert default_import_month(dt.date(2026, 8, 6)) == "2026-08"
        assert default_import_month(dt.date(2026, 12, 31)) == "2026-12"


class TestSaveLoad:
    def test_round_trip_bom_utf8(self, tmp_path):
        led = _make_ledger(tmp_path, [])
        append_provisional(led, [_sap("EXP-1")], "2026-08", "a.csv", user="谷津晴香")
        confirm_month(led, "2026-08", user="谷津晴香")
        save_ledger(led)

        raw = led.path.read_bytes()
        assert raw.startswith(b"\xef\xbb\xbf")      # Excel で文字化けしないBOM付きUTF-8

        again = load_ledger(led.path)
        assert len(again.confirmed_rows) == 1
        assert again.rows[0]["姓"] == "上原"
        assert again.rows[0][ID_COLUMN] == "EXP-1"
        # 読み直した台帳でも判定が効く
        assert len(classify_rows([_sap("EXP-1")], again).excluded_rows) == 1

    def test_load_missing_file_is_empty(self, tmp_path):
        led = load_ledger(tmp_path / "まだ無い.csv")
        assert led.rows == []
        assert led.confirmed_rows == []

    def test_broken_file_raises_not_silently_empty(self, tmp_path):
        """壊れた台帳を「空の台帳」として通すと除外が効かず二重支給になる → 必ず落とす"""
        import pytest
        broken = tmp_path / "壊れた.csv"
        broken.write_bytes(b"\xff\xfe\x00\x00not a csv")
        with pytest.raises(ValueError, match="台帳CSVの形式が不正"):
            load_ledger(broken)

    def test_wrong_csv_raises(self, tmp_path):
        """列違いのCSV（SAP生CSVを間違って指定した等）も弾く"""
        import pytest
        wrong = tmp_path / "別物.csv"
        with wrong.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, lineterminator="\r\n")
            w.writerow(["姓", "名", "費用合計"])
            w.writerow(["上原", "奏吾", "2096"])
        with pytest.raises(ValueError, match="台帳CSVの形式が不正"):
            load_ledger(wrong)

    def test_backup_created_on_overwrite(self, tmp_path):
        led = _make_ledger(tmp_path, [])
        append_provisional(led, [_sap("EXP-1")], "2026-08", "a.csv")
        assert save_ledger(led) is None          # 初回はバックアップ無し
        append_provisional(led, [_sap("EXP-2")], "2026-09", "b.csv")
        bak = save_ledger(led)
        assert bak is not None and bak.exists()
        assert bak.parent.name == "_台帳バックアップ"

    def test_header_order_is_fixed(self, tmp_path):
        led = _make_ledger(tmp_path, [])
        append_provisional(led, [_sap("EXP-1")], "2026-08", "a.csv")
        save_ledger(led)
        with led.path.open(encoding="utf-8-sig", newline="") as f:
            assert next(csv.reader(f)) == LEDGER_COLUMNS


class TestCanWrite:
    def _writers(self, tmp_path, rows):
        p = tmp_path / "writers.csv"
        with p.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=["ユーザー名", "表示名", "備考"],
                               lineterminator="\r\n")
            w.writeheader()
            w.writerows(rows)
        return p

    def test_allowed(self, tmp_path):
        p = self._writers(tmp_path, [{"ユーザー名": "谷津晴香", "表示名": "谷津さん", "備考": ""}])
        ok, msg = can_write(p, user="谷津晴香")
        assert ok is True and "書き込み可" in msg

    def test_case_insensitive(self, tmp_path):
        p = self._writers(tmp_path, [{"ユーザー名": "Taira", "表示名": "平良さん", "備考": ""}])
        assert can_write(p, user="taira")[0] is True

    def test_denied_shows_allowed_names(self, tmp_path):
        p = self._writers(tmp_path, [
            {"ユーザー名": "谷津晴香", "表示名": "谷津さん", "備考": ""},
            {"ユーザー名": "taira", "表示名": "平良さん", "備考": ""},
        ])
        ok, msg = can_write(p, user="someone")
        assert ok is False
        assert "谷津さん" in msg and "平良さん" in msg and "someone" in msg

    def test_missing_list_denies(self, tmp_path):
        """許可リストが読めないときは全員書き込み不可に倒す"""
        ok, msg = can_write(tmp_path / "無い.csv", user="谷津晴香")
        assert ok is False
        assert "読めません" in msg
