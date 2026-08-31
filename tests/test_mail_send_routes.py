"""一斉送信のルート（バッチ一覧・送信前確認・送信）を test_client で直接叩く。

app.py には似たルートが並んでいて、直したつもりが隣に入る事故が起きやすいので、
サービス層のテストとは別に、ルートを通した状態でも確かめる。
Outlook には触れない（OutlookMailer を fake に差し替える）。
"""

import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app as app_module  # noqa: E402
from config import Config  # noqa: E402
from services import mail_draft  # noqa: E402


class FakeMailer:
    """送った EntryID を覚えるだけの偽メーラー。"""

    instances = []

    def __init__(self):
        self.sent = []
        FakeMailer.instances.append(self)

    def send_saved(self, entry_id, store_id=""):
        if entry_id == "MISSING":
            raise RuntimeError("下書きが見つかりません")
        self.sent.append((entry_id, store_id))


@pytest.fixture
def client():
    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as c:
        yield c


@pytest.fixture
def mail_dir(tmp_path, monkeypatch):
    """送信バッチの置き場を tmp に向け、Outlook を fake にする。"""
    monkeypatch.setattr(Config, "MAIL_OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(mail_draft, "OutlookMailer", FakeMailer)
    FakeMailer.instances = []
    return tmp_path


def make_batch(mail_dir, *, user=None, items=None):
    batch = {
        "batch_id": "20260831_170000",
        "created_at": "2026-08-31 17:00:00",
        "user": user or mail_draft.current_user(),
        "template_name": "有休取得のお願い",
        "items": items or [
            {"employee_id": "2024001", "name": "山田太郎", "to": "a@x.jp",
             "cc": "kanri@nmht.co.jp", "bcc": "", "subject": "件名A",
             "entry_id": "EID001", "store_id": "STORE", "state": "draft"},
            {"employee_id": "2024002", "name": "鈴木花子", "to": "b@x.jp",
             "cc": "kanri@nmht.co.jp", "bcc": "", "subject": "件名B",
             "entry_id": "EID002", "store_id": "STORE", "state": "draft"},
        ],
    }
    return mail_draft.save_send_batch(batch, log_dir=mail_dir)


def test_batches_lists_only_my_own(client, mail_dir):
    make_batch(mail_dir)
    make_batch(mail_dir, user="ほかの人")   # 同じ batch_id・別ユーザー
    data = client.post("/mail_send_batches").get_json()
    assert data["success"] is True
    assert len(data["batches"]) == 1
    assert data["batches"][0]["sendable"] == 2
    assert data["batches"][0]["template_name"] == "有休取得のお願い"


def test_preview_does_not_touch_outlook(client, mail_dir):
    make_batch(mail_dir)
    data = client.post("/mail_send_preview",
                       data={"batch_id": "20260831_170000"}).get_json()
    assert data["success"] is True
    assert data["sendable"] == 2
    assert data["confirm_text"] == "SEND 2"
    assert [i["employee_id"] for i in data["items"]] == ["2024001", "2024002"]
    assert all(i["sendable"] for i in data["items"])
    assert FakeMailer.instances == []


def test_preview_marks_sent_and_missing_entry_id(client, mail_dir):
    make_batch(mail_dir, items=[
        {"employee_id": "2024001", "name": "山田太郎", "to": "a@x.jp", "cc": "", "bcc": "",
         "subject": "件名A", "entry_id": "EID001", "store_id": "", "state": "sent",
         "sent_at": "2026-08-31 17:05:00"},
        {"employee_id": "2024002", "name": "鈴木花子", "to": "b@x.jp", "cc": "", "bcc": "",
         "subject": "件名B", "entry_id": "", "store_id": "", "state": "draft"},
    ])
    data = client.post("/mail_send_preview",
                       data={"batch_id": "20260831_170000"}).get_json()
    assert data["sendable"] == 0
    assert data["confirm_text"] == "SEND 0"
    reasons = {i["employee_id"]: i["reason"] for i in data["items"]}
    assert "送信済" in reasons["2024001"]
    assert "EntryID" in reasons["2024002"]


def test_unknown_batch_is_404(client, mail_dir):
    make_batch(mail_dir)
    res = client.post("/mail_send_preview", data={"batch_id": "存在しないID"})
    assert res.status_code == 404


def test_other_users_batch_cannot_be_sent(client, mail_dir):
    make_batch(mail_dir, user="ほかの人")
    res = client.post("/mail_send_execute", data={
        "batch_id": "20260831_170000",
        "selected_ids": json.dumps(["2024001"]),
        "confirm_text": "SEND 1"})
    assert res.status_code == 404
    assert FakeMailer.instances == []


def test_wrong_confirm_text_sends_nothing(client, mail_dir):
    path = make_batch(mail_dir)
    res = client.post("/mail_send_execute", data={
        "batch_id": "20260831_170000",
        "selected_ids": json.dumps(["2024001", "2024002"]),
        "confirm_text": "SEND 1"})
    assert res.status_code == 400
    assert "SEND 2" in res.get_json()["errors"][0]
    assert FakeMailer.instances == []
    batch = mail_draft.load_send_batch(path)
    assert all(item["state"] == "draft" for item in batch["items"])


def test_execute_sends_the_saved_drafts(client, mail_dir):
    path = make_batch(mail_dir)
    data = client.post("/mail_send_execute", data={
        "batch_id": "20260831_170000",
        "selected_ids": json.dumps(["2024001", "2024002"]),
        "confirm_text": "SEND 2"}).get_json()
    assert data["success"] is True
    assert data["processed"] == 2
    assert data["failed"] == 0
    assert FakeMailer.instances[0].sent == [("EID001", "STORE"), ("EID002", "STORE")]
    batch = mail_draft.load_send_batch(path)
    assert all(item["state"] == "sent" for item in batch["items"])
    # 一覧では未送信0件になり、二度目は送るものが無い
    listed = client.post("/mail_send_batches").get_json()["batches"][0]
    assert listed["sendable"] == 0 and listed["sent"] == 2


def test_execute_only_the_selected_person(client, mail_dir):
    path = make_batch(mail_dir)
    data = client.post("/mail_send_execute", data={
        "batch_id": "20260831_170000",
        "selected_ids": json.dumps(["2024002"]),
        "confirm_text": "SEND 1"}).get_json()
    assert data["processed"] == 1
    assert FakeMailer.instances[0].sent == [("EID002", "STORE")]
    states = {i["employee_id"]: i["state"]
              for i in mail_draft.load_send_batch(path)["items"]}
    assert states == {"2024001": "draft", "2024002": "sent"}


def test_missing_draft_is_reported_but_others_are_sent(client, mail_dir):
    path = make_batch(mail_dir, items=[
        {"employee_id": "2024001", "name": "山田太郎", "to": "a@x.jp", "cc": "", "bcc": "",
         "subject": "件名A", "entry_id": "MISSING", "store_id": "", "state": "draft"},
        {"employee_id": "2024002", "name": "鈴木花子", "to": "b@x.jp", "cc": "", "bcc": "",
         "subject": "件名B", "entry_id": "EID002", "store_id": "", "state": "draft"},
    ])
    data = client.post("/mail_send_execute", data={
        "batch_id": "20260831_170000",
        "selected_ids": json.dumps(["2024001", "2024002"]),
        "confirm_text": "SEND 2"}).get_json()
    assert data["processed"] == 1
    assert data["failed"] == 1
    states = {i["employee_id"]: i["state"]
              for i in mail_draft.load_send_batch(path)["items"]}
    assert states["2024001"] == "draft"     # 送れていないので未送信のまま
    assert states["2024002"] == "sent"


def test_nothing_selected_is_rejected(client, mail_dir):
    make_batch(mail_dir)
    res = client.post("/mail_send_execute", data={
        "batch_id": "20260831_170000",
        "selected_ids": json.dumps([]),
        "confirm_text": "SEND 0"})
    assert res.status_code == 400
    assert FakeMailer.instances == []
