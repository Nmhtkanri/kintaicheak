# -*- coding: utf-8 -*-
"""メール下書きモードの純ロジックテスト（実データ・Outlook・台帳ファイルは使わない）。"""
import os
import sys
import tempfile
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.mail_draft import (
    STATUS_NG,
    STATUS_OK,
    build_mail_plans,
    create_drafts,
    extract_placeholders,
    invalid_addresses,
    load_templates,
    render_template_text,
    save_template,
    select_recipients,
    split_addresses,
    value_to_text,
)


def entry(employee_id="2024001", name="山田太郎", company=(), client=(), personal=()):
    return {"employee_id": employee_id, "name": name,
            "company": tuple(company), "client": tuple(client), "personal": tuple(personal)}


TEMPLATE = {"subject": "【連絡】{{氏名}}さんへ", "body": "{{氏名}}さん\nあと{{不足日数}}日です。",
            "cc": "", "importance": "normal"}
HEADERS = ["社員番号", "氏名", "不足日数"]


def plans_for(rows, book, template=None):
    return build_mail_plans(rows, "社員番号", "氏名", HEADERS, book, template or dict(TEMPLATE))


class ValueToTextTest(unittest.TestCase):
    def test_formats(self):
        self.assertEqual(value_to_text(3.0), "3")
        self.assertEqual(value_to_text(2.5), "2.5")
        self.assertEqual(value_to_text(date(2027, 3, 31)), "2027年3月31日")
        self.assertEqual(value_to_text(None), "")
        self.assertEqual(value_to_text(" 佐藤 "), "佐藤")


class RenderTest(unittest.TestCase):
    def test_substitution_and_spaces(self):
        row = {"氏名": "山田太郎", "不足日数": 2.0}
        rendered, missing = render_template_text("{{ 氏名 }}さん あと{{不足日数}}日", row, set(HEADERS))
        self.assertEqual(rendered, "山田太郎さん あと2日")
        self.assertEqual(missing, [])

    def test_missing_column_and_empty_value(self):
        row = {"氏名": "", "不足日数": 1}
        _, missing = render_template_text("{{氏名}} {{期限}}", row, set(HEADERS))
        self.assertIn("{{氏名}}（空欄）", missing)
        self.assertIn("{{期限}}（列がありません）", missing)

    def test_extract_placeholders(self):
        self.assertEqual(extract_placeholders("{{氏名}}", "{{氏名}} {{期限}}"), ["氏名", "期限"])


class AddressTest(unittest.TestCase):
    def test_split_and_validate(self):
        self.assertEqual(split_addresses("a@x.co.jp; b@x.co.jp、a@x.co.jp"),
                         ("a@x.co.jp", "b@x.co.jp"))
        self.assertEqual(invalid_addresses(["a@x.co.jp", "壊れたアドレス"]), ["壊れたアドレス"])

    def test_select_recipients_company_first(self):
        to, bcc, breakdown = select_recipients(
            entry(company=("c@x.jp",), client=("s@y.jp",), personal=("p@z.jp",)))
        self.assertEqual(to, ("c@x.jp",))
        self.assertEqual(bcc, ("s@y.jp", "p@z.jp"))
        self.assertEqual(breakdown, "To:社用 / BCC:就業先・個人")

    def test_select_recipients_fallback(self):
        to, _, breakdown = select_recipients(entry(client=("s@y.jp",)))
        self.assertEqual(to, ("s@y.jp",))
        self.assertEqual(breakdown, "To:就業先")


class BuildPlansTest(unittest.TestCase):
    def setUp(self):
        self.row = {"社員番号": "2024001", "氏名": "山田太郎", "不足日数": 2.0}
        self.book = {"2024001": [entry(company=("c@x.jp",), personal=("p@z.jp",))]}

    def test_ok_plan_default_is_to_only(self):
        plan = plans_for([self.row], self.book)[0]
        self.assertEqual(plan["status"], STATUS_OK)
        self.assertEqual(plan["to"], ["c@x.jp"])
        self.assertEqual(plan["bcc"], [])  # 既定はToのみ（2026-07-29変更）
        self.assertEqual(plan["breakdown"], "To:社用")
        self.assertEqual(plan["subject"], "【連絡】山田太郎さんへ")
        self.assertIn("あと2日です", plan["body"])

    def test_bcc_mode_adds_personal_bcc(self):
        template = dict(TEMPLATE, bcc_mode="bcc")
        plan = plans_for([self.row], self.book, template)[0]
        self.assertEqual(plan["bcc"], ["p@z.jp"])
        self.assertEqual(plan["breakdown"], "To:社用 / BCC:個人")

    def test_missing_ledger_entry(self):
        plan = plans_for([{"社員番号": "9999999", "氏名": "誰か", "不足日数": 1}], self.book)[0]
        self.assertEqual(plan["status"], STATUS_NG)
        self.assertTrue(any("メール台帳行がありません" in issue for issue in plan["issues"]))

    def test_name_mismatch_blocks(self):
        book = {"2024001": [entry(name="別人花子", company=("c@x.jp",))]}
        plan = plans_for([self.row], book)[0]
        self.assertEqual(plan["status"], STATUS_NG)
        self.assertTrue(any("氏名相違" in issue for issue in plan["issues"]))

    def test_duplicate_same_addresses_is_note_only(self):
        same = entry(company=("c@x.jp",))
        plan = plans_for([self.row], {"2024001": [same, dict(same)]})[0]
        self.assertEqual(plan["status"], STATUS_OK)
        self.assertTrue(any("重複行" in issue for issue in plan["issues"]))

    def test_merge_failure_blocks(self):
        row = {"社員番号": "2024001", "氏名": "山田太郎", "不足日数": None}
        plan = plans_for([row], self.book)[0]
        self.assertEqual(plan["status"], STATUS_NG)
        self.assertTrue(any("差し込みできません" in issue for issue in plan["issues"]))

    def test_invalid_cc_raises(self):
        template = dict(TEMPLATE, cc="これはアドレスではない")
        with self.assertRaises(ValueError):
            plans_for([self.row], self.book, template)

    def test_to_only_mode_drops_bcc(self):
        template = dict(TEMPLATE, bcc_mode="to_only")
        plan = plans_for([self.row], self.book, template)[0]
        self.assertEqual(plan["status"], STATUS_OK)
        self.assertEqual(plan["to"], ["c@x.jp"])
        self.assertEqual(plan["bcc"], [])
        self.assertEqual(plan["breakdown"], "To:社用")


class DummyMailer:
    def __init__(self, fail_times=0):
        self.calls = []
        self.fail_times = fail_times

    def create_draft(self, **kwargs):
        if self.fail_times > 0:
            self.fail_times -= 1
            raise RuntimeError("ダミー失敗")
        self.calls.append(kwargs)


class CreateDraftsTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        book = {"2024001": [entry(company=("c@x.jp",))]}
        rows = [{"社員番号": "2024001", "氏名": "山田太郎", "不足日数": 2},
                {"社員番号": "9999999", "氏名": "台帳なし", "不足日数": 1}]
        self.plans = plans_for(rows, book)

    def test_creates_only_selected_ok(self):
        mailer = DummyMailer()
        result = create_drafts(self.plans, only_ids=["2024001", "9999999"],
                               log_dir=self.tmp.name, mailer=mailer)
        self.assertEqual(result["processed"], 1)
        self.assertEqual(result["skipped"], 1)  # 要確認は選択されていても作らない
        self.assertEqual(len(mailer.calls), 1)
        self.assertEqual(mailer.calls[0]["to"], "c@x.jp")
        self.assertTrue(os.path.exists(result["log_path"]))

    def test_unselected_not_created(self):
        mailer = DummyMailer()
        result = create_drafts(self.plans, only_ids=[], log_dir=self.tmp.name, mailer=mailer)
        self.assertEqual(result["processed"], 0)
        self.assertEqual(mailer.calls, [])


class TemplateStoreTest(unittest.TestCase):
    def test_save_load_delete_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "テンプレ.json")
            save_template(path, {"name": "案内", "subject": "S", "body": "B",
                                 "cc": "", "bcc_mode": "to_only", "importance": "high"})
            templates = load_templates(path)
            self.assertEqual(len(templates), 1)
            self.assertEqual(templates[0]["importance"], "high")
            self.assertEqual(templates[0]["bcc_mode"], "to_only")
            save_template(path, {"name": "案内"}, delete=True)
            self.assertEqual(load_templates(path), [])

    def test_missing_file_returns_empty(self):
        self.assertEqual(load_templates(r"C:\存在しない\テンプレ.json"), [])


if __name__ == "__main__":
    unittest.main()
