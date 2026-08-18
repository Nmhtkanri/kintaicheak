# -*- coding: utf-8 -*-
"""標報投入の突合・計画・実行。

jinjer は一切叩かない（クライアントは差し替える）。見張るのは主に4つ:

  1. PDF・jinjer登録値・当方計算値の3点でステータスが正しく分かれること
  2. **標準報酬が動いていない人には書かない**こと（有効な値との比較）
  3. 投入が冪等であること（途中で落ちても同じPDFで再開できる）
  4. dry-run では1回もAPIを呼ばないこと
"""

from __future__ import annotations

import csv
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services import shaho_writer  # noqa: E402
from services.shaho_pdf import PdfPerson, PdfStatement  # noqa: E402
from services.shaho_writer import (CalcResult, build_plan, execute_plan,  # noqa: E402
                                   ledger_entries, pick_records, plan_hash,
                                   summarize, verify_after)

TARGET = "2026-07"


def person(emp="2099001", name="試験 太郎", kenpo=300000, konen=300000,
           reason="月額変更", **kw):
    return PdfPerson(emp=emp, name=name, kenpo_smr=kenpo, konen_smr=konen,
                     reason_kenpo=reason, **kw)


def statement(*persons, target_ym=TARGET):
    return PdfStatement(target_ym=target_ym, pay_ym="2026-08", office_code="263",
                        persons=list(persons))


def rec(ym=TARGET, kenpo=300000, konen=300000, updater="管理者登録"):
    year, month = ym.split("-")
    return {"year": year, "month": month, "collection_month": "2026-08",
            "health_insurance": {"fee": str(kenpo)},
            "employee_pension": {"fee": str(konen)},
            "last_update": {"classification": {"id": "1", "name": updater},
                            "updater": {"name": "谷津"}}}


def roster(*emps, name="試験 太郎", enrollment="在籍", retired_on=""):
    return {e: {"name": name, "enrollment": enrollment, "joined_on": "2020-04-01",
                "retired_on": retired_on} for e in emps}


_DEFAULT = object()


def plan_for(persons, current=None, roster_map=_DEFAULT, calc=None, pdf_issues=None):
    """expected_smr を差し替えて build_plan だけを見る。

    roster_map に空 dict を渡す＝「jinjer に居ない」テストなので、
    既定値の判定は None ではなくセンチネルで行う（`or` だと空dictが潰れる）。
    """
    if roster_map is _DEFAULT:
        roster_map = roster(*[p.emp for p in persons])
    stmt = statement(*persons)
    with mock.patch.object(shaho_writer, "expected_smr",
                           side_effect=lambda p, ym, ctx: calc or CalcResult(note="計算なし")):
        return build_plan(stmt, current or {}, roster_map, object(), pdf_issues)


class PickRecordsTests(unittest.TestCase):
    def test_effective_is_the_newest_at_or_before_target(self):
        recs = [rec("2026-04", 260000, 260000), rec("2026-06", 300000, 300000),
                rec("2026-09", 320000, 320000)]
        at_target, effective = pick_records(recs, TARGET)
        self.assertIsNone(at_target)                       # 7月ちょうどのレコードは無い
        self.assertEqual(effective["month"], "06")         # 7月時点で効いているのは6月分

    def test_record_at_target_is_found(self):
        at_target, effective = pick_records([rec("2026-06"), rec(TARGET)], TARGET)
        self.assertIsNotNone(at_target)
        self.assertEqual(effective["month"], "07")

    def test_empty(self):
        self.assertEqual(pick_records([], TARGET), (None, None))


class StatusTests(unittest.TestCase):
    def test_auto_ok_when_all_three_agree(self):
        row = plan_for([person(kenpo=300000, konen=300000)],
                       current={"2099001": [rec(kenpo=280000, konen=280000)]},
                       calc=CalcResult(kenpo=300000, konen=300000, source="随時改定"))[0]
        self.assertEqual(row.status, "AUTO_OK")
        self.assertTrue(row.default_selected)
        self.assertEqual(row.operation, "PATCH")           # 7月のレコードがある → 更新

    def test_calc_mismatch_is_not_selected_by_default(self):
        row = plan_for([person(kenpo=300000)],
                       current={"2099001": [rec(kenpo=280000)]},
                       calc=CalcResult(kenpo=320000, konen=320000, source="随時改定"))[0]
        self.assertEqual(row.status, "CALC_MISMATCH")
        self.assertFalse(row.default_selected)
        self.assertTrue(row.needs_force)                   # 承知のうえ投入なら通せる
        self.assertTrue(row.selectable)

    def test_no_calc_is_selectable_without_force(self):
        row = plan_for([person(reason="取得時決定")],
                       current={}, calc=CalcResult(note="見込み額なので計算できない"))[0]
        self.assertEqual(row.status, "NO_CALC")
        self.assertFalse(row.needs_force)
        self.assertTrue(row.selectable)
        self.assertEqual(row.operation, "POST")            # レコードが無い → 新規

    def test_no_change_when_effective_value_matches(self):
        """料率変更の月に、標準報酬が動いていない人へ履歴を足さない。"""
        row = plan_for([person(reason="料率変更")],
                       current={"2099001": [rec("2026-04", 300000, 300000)]})[0]
        self.assertEqual(row.status, "NO_CHANGE")
        self.assertFalse(row.selectable)

    def test_old_record_still_counts_as_the_current_value(self):
        """★回帰: 対象月にレコードが無くても、過去のレコードが効いていれば書かない。

        2026-08-17 に本番exeで踏んだ。報酬月額APIを year/month で絞って取ると
        「対象月のレコードが無い人＝登録なし」に見えてしまい、標準報酬が動いて
        いない人（料率変更のみ）に要らない履歴を足すところだった。
        **プレビューと実行直前の取得は月で絞らないこと。**
        """
        rows = plan_for([person(reason="料率変更")],
                        current={"2099001": [rec("2025-09", 300000, 300000)]})
        self.assertEqual(rows[0].status, "NO_CHANGE")
        self.assertEqual(rows[0].cur_kenpo, 300000)
        self.assertEqual(rows[0].cur_ym, "2025-09")

    def test_future_record_is_not_treated_as_current(self):
        """先付けのレコード（9月の定時決定など）は対象月の現在値にしない。"""
        rows = plan_for([person(reason="料率変更")],
                        current={"2099001": [rec("2025-09", 300000, 300000),
                                             rec("2026-09", 500000, 500000)]})
        self.assertEqual(rows[0].cur_kenpo, 300000)
        self.assertEqual(rows[0].status, "NO_CHANGE")

    def test_not_in_jinjer(self):
        row = plan_for([person()], roster_map={})[0]
        self.assertEqual(row.status, "NOT_IN_JINJER")
        self.assertFalse(row.selectable)

    def test_excluded_employee_number(self):
        row = plan_for([person(emp="9999999")], roster_map=roster("9999999"))[0]
        self.assertEqual(row.status, "EXCLUDED")
        self.assertFalse(row.selectable)

    def test_retired_needs_force(self):
        row = plan_for([person()], roster_map=roster("2099001", enrollment="退職",
                                                    retired_on="2026-06-30"))[0]
        self.assertEqual(row.status, "RETIRED")
        self.assertTrue(row.needs_force)

    def test_pdf_inconsistent_cannot_be_forced(self):
        row = plan_for([person()], pdf_issues={"2099001": ["健康保険計が内訳と合いません"]})[0]
        self.assertEqual(row.status, "PDF_INCONSISTENT")
        self.assertFalse(row.selectable)     # 読み取りを信用できない人は投入させない

    def test_row_issue_also_blocks(self):
        p = person()
        p.issues.append("2行目の数値が想定と違う")
        self.assertEqual(plan_for([p])[0].status, "PDF_INCONSISTENT")

    def test_name_difference_is_noted(self):
        row = plan_for([person(name="別人 太郎")], roster_map=roster("2099001"))[0]
        self.assertTrue(any("氏名がjinjerと違う" in n for n in row.notes))

    def test_summary_counts(self):
        rows = plan_for([person(emp="2099001"), person(emp="2099002")],
                        roster_map=roster("2099001", "2099002"))
        summary = summarize(rows)
        self.assertEqual(sum(s["count"] for s in summary), 2)


class PlanHashTests(unittest.TestCase):
    def test_hash_is_stable_and_sensitive(self):
        rows = plan_for([person()])
        first = plan_hash(rows, TARGET)
        self.assertEqual(first, plan_hash(rows, TARGET))
        rows[0].pdf_kenpo = 999999
        self.assertNotEqual(first, plan_hash(rows, TARGET))

    def test_hash_depends_on_month(self):
        rows = plan_for([person()])
        self.assertNotEqual(plan_hash(rows, TARGET), plan_hash(rows, "2026-08"))


class FakeClient:
    """書き込みクライアントの代役。呼ばれた内容を全部覚える。"""

    def __init__(self, fail_on=(), state=None):
        self.calls = []
        self.fail_on = set(fail_on)
        self.state = dict(state or {})

    def _write(self, op, emp, year, month, kenpo, konen):
        self.calls.append((op, emp, year, month, kenpo, konen))
        if emp in self.fail_on:
            raise RuntimeError("jinjer が 400 を返しました")
        self.state[emp] = (kenpo, konen)
        return {}

    def post_monthly_remuneration(self, emp, year, month, kenpo, konen):
        return self._write("POST", emp, year, month, kenpo, konen)

    def patch_monthly_remuneration(self, emp, year, month, kenpo, konen):
        return self._write("PATCH", emp, year, month, kenpo, konen)

    def get_monthly_remunerations(self, emps, year=None, month=None):
        return {e: [rec(TARGET, *self.state[e])] for e in emps if e in self.state}


class ExecuteTests(unittest.TestCase):
    def test_dry_run_never_calls_the_api(self):
        rows = plan_for([person()], current={"2099001": [rec(kenpo=280000)]},
                        calc=CalcResult(kenpo=300000, konen=300000))
        client = FakeClient()
        results = execute_plan(rows, client, TARGET, dry_run=True)
        self.assertEqual(client.calls, [])
        self.assertEqual(results[0]["result"], "dry-run")

    def test_patch_and_post_are_chosen_per_row(self):
        rows = plan_for([person(emp="2099001"), person(emp="2099002")],
                        current={"2099001": [rec(kenpo=280000)]},
                        roster_map=roster("2099001", "2099002"),
                        calc=CalcResult(kenpo=300000, konen=300000))
        client = FakeClient()
        execute_plan(rows, client, TARGET, dry_run=False)
        self.assertEqual([c[0] for c in client.calls], ["PATCH", "POST"])
        self.assertEqual(client.calls[0][3:], ("07", 300000, 300000))

    def test_already_correct_row_is_skipped(self):
        """再実行したときに、入っている人をもう一度書かない（冪等）。"""
        rows = plan_for([person()], current={"2099001": [rec(kenpo=280000)]},
                        calc=CalcResult(kenpo=300000, konen=300000))
        fresh = {"2099001": [rec(TARGET, 300000, 300000)]}
        results = execute_plan(rows, FakeClient(), TARGET, dry_run=False, fresh=fresh)
        self.assertEqual(results[0]["result"], "スキップ")

    def test_value_changed_since_preview_aborts_that_row(self):
        rows = plan_for([person()], current={"2099001": [rec(kenpo=280000)]},
                        calc=CalcResult(kenpo=300000, konen=300000))
        fresh = {"2099001": [rec(TARGET, 260000, 260000)]}   # 誰かが別の値に変えた
        client = FakeClient()
        results = execute_plan(rows, client, TARGET, dry_run=False, fresh=fresh)
        self.assertEqual(results[0]["result"], "中止")
        self.assertEqual(client.calls, [])

    def test_one_failure_does_not_stop_the_rest(self):
        rows = plan_for([person(emp="2099001"), person(emp="2099002")],
                        roster_map=roster("2099001", "2099002"),
                        calc=CalcResult(kenpo=300000, konen=300000))
        results = execute_plan(rows, FakeClient(fail_on={"2099001"}), TARGET, dry_run=False)
        self.assertEqual([r["result"] for r in results], ["失敗", "OK"])

    def test_progress_is_reported_per_row(self):
        rows = plan_for([person(emp="2099001"), person(emp="2099002")],
                        roster_map=roster("2099001", "2099002"),
                        calc=CalcResult(kenpo=300000, konen=300000))
        seen = []
        execute_plan(rows, FakeClient(), TARGET, dry_run=False,
                     progress=lambda done, total, entry: seen.append((done, total)))
        self.assertEqual(seen, [(1, 2), (2, 2)])


class VerifyTests(unittest.TestCase):
    def test_verify_ok_and_ng(self):
        rows = plan_for([person(emp="2099001"), person(emp="2099002")],
                        roster_map=roster("2099001", "2099002"),
                        calc=CalcResult(kenpo=300000, konen=300000))
        client = FakeClient(state={"2099001": (300000, 300000),
                                   "2099002": (280000, 280000)})
        result = verify_after(client, rows, TARGET)
        self.assertEqual(result["2099001"], "OK")
        self.assertTrue(result["2099002"].startswith("NG"))

    def test_wrong_collection_month_is_flagged(self):
        """★金額が合っていても、控除される月がずれていたらOKと呼ばない。

        2026-08-18 に実際に起きた: API で作ったレコードは徴収年月が
        「基準年月と同じ」になり、画面入力（基準+1か月）とずれていた。
        """
        rows = plan_for([person()], calc=CalcResult(kenpo=300000, konen=300000))

        class Client(FakeClient):
            def get_monthly_remunerations(self, emps, year=None, month=None):
                rec_ = rec(TARGET, 300000, 300000)
                rec_["collection_month"] = TARGET          # 本来は 2026-08
                return {e: [rec_] for e in emps}

        result = verify_after(Client(), rows, TARGET)
        self.assertIn("徴収年月", result["2099001"])
        self.assertNotEqual(result["2099001"], "OK")

    def test_correct_collection_month_is_ok(self):
        rows = plan_for([person()], calc=CalcResult(kenpo=300000, konen=300000))
        client = FakeClient(state={"2099001": (300000, 300000)})
        self.assertEqual(verify_after(client, rows, TARGET)["2099001"], "OK")

    def test_missing_record_is_ng(self):
        rows = plan_for([person()], calc=CalcResult(kenpo=300000, konen=300000))
        self.assertTrue(verify_after(FakeClient(), rows, TARGET)["2099001"].startswith("NG"))


class IdempotencyTests(unittest.TestCase):
    def test_replanning_after_a_write_shows_no_change(self):
        """途中で落ちても「同じPDFでもう一度プレビュー」で残りだけになる。"""
        people = [person(emp="2099001"), person(emp="2099002")]
        calc = CalcResult(kenpo=300000, konen=300000)
        rows = plan_for(people, roster_map=roster("2099001", "2099002"), calc=calc)
        client = FakeClient()
        execute_plan(rows[:1], client, TARGET, dry_run=False)     # 1人だけ投入して中断

        again = plan_for(people, current=client.get_monthly_remunerations(
            ["2099001", "2099002"]), roster_map=roster("2099001", "2099002"), calc=calc)
        by_emp = {r.emp: r for r in again}
        self.assertEqual(by_emp["2099001"].status, "NO_CHANGE")   # 投入済み
        self.assertEqual(by_emp["2099002"].status, "AUTO_OK")     # 残り


class LedgerTests(unittest.TestCase):
    def test_header_is_written_once_and_rows_appended(self):
        rows = plan_for([person()], calc=CalcResult(kenpo=300000, konen=300000))
        results = execute_plan(rows, FakeClient(), TARGET, dry_run=False)
        entries = ledger_entries(results, {"2099001": "OK"}, target_ym=TARGET,
                                 pdf_name="7月保険料一覧表.pdf", backup="backup.json",
                                 forced={"2099001"})
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "台帳.csv")
            shaho_writer.append_ledger(entries, path=path)
            shaho_writer.append_ledger(entries, path=path)
            with open(path, encoding="utf-8-sig", newline="") as f:
                data = list(csv.DictReader(f))
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0]["社員番号"], "2099001")
        self.assertEqual(data[0]["健保後"], "300000")
        self.assertEqual(data[0]["検証"], "OK")
        self.assertEqual(data[0]["承知投入"], "○")


class PermissionTests(unittest.TestCase):
    """許可は2段構え: 投入できる人と、要確認を承知のうえ投入できる人。"""

    def _allowlist(self, tmp, rows, header=("ユーザー名", "表示名", "承知投入", "備考")):
        path = os.path.join(tmp, "許可.csv")
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(list(header))
            writer.writerows(rows)
        return path

    def test_missing_allowlist_means_nobody_can_write(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch.object(shaho_writer.Config, "SHAHO_IMPORT_ALLOWED_USERS_CSV",
                                   os.path.join(tmp, "ない.csv")):
                allowed, why = shaho_writer.can_write("谷津晴香")
                self.assertFalse(shaho_writer.can_force("谷津晴香")[0])
        self.assertFalse(allowed)
        self.assertIn("許可リスト", why)

    def test_listed_user_can_write_others_cannot(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._allowlist(tmp, [["谷津晴香", "谷津さん", "○", "管理部"]])
            with mock.patch.object(shaho_writer.Config,
                                   "SHAHO_IMPORT_ALLOWED_USERS_CSV", path):
                self.assertTrue(shaho_writer.can_write("谷津晴香")[0])
                self.assertFalse(shaho_writer.can_write("別の人")[0])

    def test_force_needs_its_own_mark(self):
        """投入はできるが、要確認の承知投入は印のある人だけ。"""
        with tempfile.TemporaryDirectory() as tmp:
            path = self._allowlist(tmp, [["谷津晴香", "谷津さん", "○", ""],
                                         ["平良菜津子", "平良さん", "", ""]])
            with mock.patch.object(shaho_writer.Config,
                                   "SHAHO_IMPORT_ALLOWED_USERS_CSV", path):
                self.assertTrue(shaho_writer.can_write("平良菜津子")[0])
                allowed, why = shaho_writer.can_force("平良菜津子")
                self.assertFalse(allowed)
                self.assertIn("谷津さん", why)
                self.assertTrue(shaho_writer.can_force("谷津晴香")[0])

    def test_missing_force_column_denies_everyone(self):
        """列を足し忘れた古いCSVで、全員が承知投入できるようになってはいけない。"""
        with tempfile.TemporaryDirectory() as tmp:
            path = self._allowlist(tmp, [["谷津晴香", "谷津さん", "管理部"]],
                                   header=("ユーザー名", "表示名", "備考"))
            with mock.patch.object(shaho_writer.Config,
                                   "SHAHO_IMPORT_ALLOWED_USERS_CSV", path):
                self.assertTrue(shaho_writer.can_write("谷津晴香")[0])
                self.assertFalse(shaho_writer.can_force("谷津晴香")[0])

    def test_unlisted_user_cannot_force_either(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._allowlist(tmp, [["谷津晴香", "谷津さん", "○", ""]])
            with mock.patch.object(shaho_writer.Config,
                                   "SHAHO_IMPORT_ALLOWED_USERS_CSV", path):
                self.assertFalse(shaho_writer.can_force("別の人")[0])


class LockTests(unittest.TestCase):
    def test_second_run_is_refused_then_allowed_after_release(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "実行中.lock")
            shaho_writer.acquire_lock(TARGET, 3, path=path)
            with self.assertRaises(shaho_writer.ShahoWriteError):
                shaho_writer.acquire_lock(TARGET, 3, path=path)
            shaho_writer.release_lock(path)
            shaho_writer.acquire_lock(TARGET, 3, path=path)      # 解放後は取れる

    def test_stale_lock_is_ignored(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "実行中.lock")
            shaho_writer.acquire_lock(TARGET, 3, path=path)
            shaho_writer.acquire_lock(TARGET, 3, path=path, max_age_hours=0)


class WindowTests(unittest.TestCase):
    def test_only_the_windows_actually_needed_are_loaded(self):
        """月額変更だけの紙で、定時決定の窓（前年4〜6月）まで探しにいかない。"""
        revision_only = statement(person(reason="月額変更"), person(reason="取得時決定"))
        self.assertEqual(shaho_writer.needed_windows(TARGET, revision_only),
                         ["2026-04", "2026-05", "2026-06"])
        teiji = statement(person(reason="定時決定"), target_ym="2026-09")
        self.assertEqual(shaho_writer.needed_windows("2026-09", teiji),
                         ["2026-04", "2026-05", "2026-06"])
        both = statement(person(reason="月額変更"), person(reason="定時決定"))
        self.assertEqual(len(shaho_writer.needed_windows(TARGET, both)), 6)
        nothing = statement(person(reason="料率変更"))
        self.assertEqual(shaho_writer.needed_windows(TARGET, nothing), [])

    def test_revision_window_is_the_three_months_before(self):
        self.assertEqual(shaho_writer.revision_window("2026-07"),
                         ["2026-04", "2026-05", "2026-06"])
        self.assertEqual(shaho_writer.revision_window("2026-01"),
                         ["2025-10", "2025-11", "2025-12"])

    def test_teiji_window_follows_september_application(self):
        self.assertEqual(shaho_writer.teiji_window("2026-09"),
                         ["2026-04", "2026-05", "2026-06"])
        self.assertEqual(shaho_writer.teiji_window("2026-07"),
                         ["2025-04", "2025-05", "2025-06"])

    def test_fiscal_year_starts_in_april(self):
        self.assertEqual(shaho_writer._fiscal_year("2026-07"), 2026)
        self.assertEqual(shaho_writer._fiscal_year("2026-03"), 2025)


class ExpectedSmrTests(unittest.TestCase):
    """理由ごとの分岐（実データを使わない範囲）。"""

    def test_no_master_means_no_calc(self):
        result = shaho_writer.expected_smr(person(), TARGET,
                                           shaho_writer.CalcContext(error="等級表が読めない"))
        self.assertIsNone(result.kenpo)
        self.assertIn("等級表", result.note)

    def test_shutoku_is_not_computable(self):
        ctx = shaho_writer.CalcContext(master=object())
        result = shaho_writer.expected_smr(person(reason="取得時決定"), TARGET, ctx)
        self.assertIsNone(result.kenpo)
        self.assertIn("見込み額", result.note)

    def test_rate_change_is_not_computable(self):
        ctx = shaho_writer.CalcContext(master=object())
        result = shaho_writer.expected_smr(person(reason="料率変更"), TARGET, ctx)
        self.assertIsNone(result.kenpo)
        self.assertIn("料率変更", result.note)

    def test_unknown_reason_is_reported(self):
        ctx = shaho_writer.CalcContext(master=object())
        result = shaho_writer.expected_smr(person(reason=""), TARGET, ctx)
        self.assertIsNone(result.kenpo)
        self.assertIn("判定できない", result.note)

    def test_missing_month_cache_is_explained(self):
        ctx = shaho_writer.CalcContext(master=object(), months={})
        result = shaho_writer.expected_smr(person(reason="月額変更"), TARGET, ctx)
        self.assertIsNone(result.kenpo)
        self.assertIn("給与明細がない", result.note)


class RealDataTests(unittest.TestCase):
    """実PDF＋実キャッシュでの回帰。社労士の値と当方計算が一致することを見る。"""

    @classmethod
    def setUpClass(cls):
        from config import Config
        from services.shaho_pdf import read_pdf
        pdf = (r"Z:\NMHT総務関係\社労士提出\提出書類\2026年度\2026.07\送付0730"
               r"\7月保険料一覧表（8月給与控除分）\7月保険料一覧表（8月給与控除分）.pdf")
        roster_json = os.path.join(Config.KEIRI_OUTPUT_DIR, "raw", "roster.json")
        if not (os.path.exists(pdf) and os.path.exists(roster_json)):
            raise unittest.SkipTest("実PDFまたは従業員一覧キャッシュがありません")
        cls.stmt = read_pdf(pdf, expected_office="263")
        cls.ctx = shaho_writer.load_calc_context(cls.stmt.target_ym)
        if cls.ctx.master is None:
            raise unittest.SkipTest(f"等級表が読めません: {cls.ctx.error}")

        import json

        from services.keiri_api import roster_index
        with open(roster_json, encoding="utf-8") as f:
            cls.roster = roster_index(json.load(f)["employees"])

    def test_calculation_agrees_with_the_shiroushi_for_every_revision(self):
        """月額変更33名すべてで、当方の随時改定計算が社労士の値と一致する。"""
        rows = build_plan(self.stmt, {}, self.roster, self.ctx)
        mismatch = [r for r in rows if r.status == "CALC_MISMATCH"]
        self.assertEqual(mismatch, [], "計算が社労士と食い違う人がいます")
        self.assertEqual(sum(1 for r in rows if r.status == "AUTO_OK"), 33)

    def test_only_uncomputable_reasons_fall_back(self):
        rows = build_plan(self.stmt, {}, self.roster, self.ctx)
        fallback = {r.emp: r.reason for r in rows if r.status == "NO_CALC"}
        self.assertEqual(sorted(fallback.values()), ["取得時決定", "取得時決定", "料率変更"])


if __name__ == "__main__":
    unittest.main()
