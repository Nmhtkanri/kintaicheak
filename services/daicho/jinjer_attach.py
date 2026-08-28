# -*- coding: utf-8 -*-
"""生成済みPDFを jinjer カスタム項目「派遣元管理台帳」（menu_id=16・項目追加（横））へ添付する。

項目構成（/v1/master/custom-menus 実測 2026-08-29）:
  item 1 = 添付日付（カレンダー） / item 2 = 派遣元管理台帳PDF（ファイル type=3）

1PDFの流れ:
  1. GET /v1/employees/addible-custom-items?customize-menu-id=16 で既存レコードを取得
  2. 同じ添付日付で既にファイルが付いていればスキップ（再実行安全）
  3. 同じ日付でファイル未添付のレコードがあればそれに、無ければ POST で日付レコードを
     作成→再GETでデータIDを特定
  4. PATCH /v1/async/files（type.id=8=カスタム項目ファイル）で record_code=データID に
     Base64 の PDF を添付。※非同期＝200は受付。反映は最大15分後、--verify で確認する

書き込みレートが厳しい（429）ため write_interval 秒空ける。429時はクライアントが待って再試行。
添付日付はファイル名の期間ラベル先頭 = 「2025年7-9月分」→ 2025/7/1（谷津さん指定 2026-08-29）。
"""
from __future__ import annotations

import base64
import datetime as dt
import re
from pathlib import Path

from .config import DATA_ROOT, LOG_DIR

PDF_ROOT = DATA_ROOT / "PDF"
LOG_PATH = LOG_DIR / "jinjer添付ログ.txt"
MENU_ID = "16"
ITEM_DATE = "1"
ITEM_FILE = "2"
FILE_KIND_CUSTOM = "8"      # async/files type.id: カスタム項目＞詳細項目（入力形式=3:ファイル）

_DATE_IN_NAME = re.compile(r"_(\d{4})年(\d{1,2})")


def _client(write_interval: float):
    from .hub_client import DaichoAttachClient
    client = DaichoAttachClient(write_interval=write_interval)
    client.authenticate()
    return client


def attach_date_of(pdf_name: str) -> str | None:
    """'2007002_横山弘樹_2025年7-9月分.pdf' → '2025/7/1'（期間ラベル先頭の年月の1日）。"""
    m = _DATE_IN_NAME.search(pdf_name)
    if not m:
        return None
    return f"{int(m.group(1))}/{int(m.group(2))}/1"


def _norm_date(s: str) -> str:
    """'2025/7/1'・'2025-07-01' → '2025-07-01'。読めなければ原文。"""
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$", str(s or "").strip())
    return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else str(s or "")


def scan_pdfs(pdf_root: Path = PDF_ROOT, employees: set[str] | None = None):
    """[(社員番号, フォルダ名, PDFパス, 添付日付)] を返す。未確定_*（社員番号なし）は除外。"""
    jobs, skipped = [], []
    for folder in sorted(pdf_root.iterdir()):
        if not folder.is_dir():
            continue
        emp = folder.name.split("_", 1)[0]
        if not emp.isdigit():
            skipped.append(f"{folder.name}: 社員番号が無いフォルダ → 添付対象外")
            continue
        if employees and emp not in employees:
            continue
        for pdf in sorted(folder.glob("*.pdf")):
            date = attach_date_of(pdf.name)
            if date is None:
                skipped.append(f"{pdf.name}: ファイル名から期間が読めない")
                continue
            jobs.append((emp, folder.name, pdf, date))
    return jobs, skipped


def fetch_records(client, employee_ids: list[str]) -> dict[str, list[dict]]:
    """{社員番号: [{id, date, attached}]}。

    ファイル項目の value/label は添付済みでも常に空で返る（2026-08-29 実測・画面では見える）。
    添付済みかどうかは「ファイル項目の updated_at が created_at から進んでいるか」で判定する
    （添付の非同期ワーカーが処理時に updated_at を進める。横山さん4件で実証）。
    """
    out: dict[str, list[dict]] = {emp: [] for emp in employee_ids}
    for i in range(0, len(employee_ids), 50):
        chunk = employee_ids[i:i + 50]
        page = 1
        while True:
            resp = client._request(
                "GET", "/v1/employees/addible-custom-items",
                params={"customize-menu-id": MENU_ID, "employee-ids": ",".join(chunk), "page": page},
            )
            # _request は data 部分（リスト）を直接返す実装。dictで来ても拾えるように両対応
            data = resp.get("data", []) if isinstance(resp, dict) else (resp or [])
            for rec in data:
                emp = str(rec.get("employee_id", ""))
                menus = rec.get("customize_menu") or []
                if isinstance(menus, dict):   # customize-menu-id で絞ると配列でなく単一dictで返る（2026-08-29実測）
                    menus = [menus]
                for menu in menus:
                    if str(menu.get("id")) != MENU_ID:
                        continue
                    for row in menu.get("customize_data", []):
                        item = {"id": str(row.get("id")), "date": "", "attached": False}
                        for it in row.get("customize_item", []):
                            val = it.get("value")
                            val = "" if val is None else (";".join(map(str, val)) if isinstance(val, list) else str(val))
                            if str(it.get("id")) == ITEM_DATE:
                                item["date"] = _norm_date(val)
                            elif str(it.get("id")) == ITEM_FILE:
                                upd, cre = it.get("updated_at"), it.get("created_at")
                                item["attached"] = bool(val) or bool(upd and upd != cre)
                        out.setdefault(emp, []).append(item)
            if len(data) < 100:
                break
            page += 1
    return out


def post_date_record(client, emp: str, date: str):
    # カレンダー項目の POST は yyyy-MM-dd 固定（GET は YYYY/M/D で返すのに別形式。2026-08-29 実測）
    body = {"employee_id": emp,
            "customize_menu": {"id": MENU_ID, "customize_item": [{"id": ITEM_DATE, "value": _norm_date(date)}]}}
    return client._request("POST", "/v1/employees/addible-custom-items", json_body=body)


def attach_file(client, emp: str, record_id: str, pdf: Path):
    body = {
        "type": {"id": FILE_KIND_CUSTOM},
        "employee_id": emp,
        "customize_menu": {"id": MENU_ID, "customize_item": {"id": ITEM_FILE}},
        "record_code": str(record_id),
        "file": {"name": pdf.name, "encoded_string": base64.b64encode(pdf.read_bytes()).decode("ascii")},
    }
    return client._request("PATCH", "/v1/async/files", json_body=body)


def _log(line: str) -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as fp:
        fp.write(f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {line}\n")


def _say(msg: str) -> None:
    """printの安全版。リダイレクト先（NAS等）のハンドルが死んでもバッチを止めない。
    2026-08-28 実障害: 60件目の進捗printが [Errno 22] → 例外処理内のprintも失敗 → プロセス死。
    exe（コンソールがcp932）では ⚠ 等が UnicodeEncodeError（ValueError系）になるのも同日実測。"""
    try:
        print(msg, flush=True)
    except (OSError, ValueError):
        pass


def _write_with_retry(client, fn, what: str, max_tries: int = 10, cooldown: float = 180.0):
    """書き込みを 429・トークン失効(403) に負けずに実行する。

    - 429: jinjer の書き込み制限は「間隔」でなく時間あたりの枠らしく（2026-08-28 実測:
      朝に速いペースで叩くと以後の書き込みが数分間すべて 429）、クライアントの3回
      リトライでは足りないことがある。cooldown 秒待って自前で粘る。
    - 403 "Access token verification failed": トークンは4時間で失効し、クライアントの
      自動再認証は 401 しか拾わない（2026-08-28 実障害: 4時間経過後の残り全件が403でNG）。
      再認証して再試行する。"""
    import time
    for attempt in range(1, max_tries + 1):
        try:
            return fn()
        except Exception as exc:
            msg = str(exc)
            if attempt >= max_tries:
                raise
            if "403" in msg and "token" in msg.lower():
                _say(f"    … トークン失効: 再認証して {what} を再試行 ({attempt}/{max_tries})")
                client.authenticate()
                continue
            if "429" not in msg:
                raise
            _say(f"    … 429（書き込み枠待ち）: {what} を{int(cooldown)}秒後に再試行 ({attempt}/{max_tries})")
            time.sleep(cooldown)


def run(employees: list[str] | None = None, dry_run: bool = True,
        write_interval: float = 12.0, limit: int | None = None,
        on_progress=None, should_stop=None) -> tuple[int, int, int]:
    """PDFを添付する。戻り値 (添付数, 添付済みスキップ数, 対象数)。

    on_progress(done, skip, total, current) と should_stop() はハブのジョブ用
    （省略時は従来のCLI動作のまま）。should_stop が True を返すと次の1件から止まる。
    """
    emp_filter = set(employees) if employees else None
    jobs, skipped = scan_pdfs(employees=emp_filter)
    for s in skipped:
        _say(f"  ⚠ {s}")
    if limit:
        jobs = jobs[:limit]
    emp_ids = sorted({emp for emp, *_ in jobs})
    _say(f"対象: {len(jobs)}ファイル / {len(emp_ids)}人（menu {MENU_ID}・添付日付=期間開始日）")
    if not jobs:
        return 0, 0, 0

    client = _client(write_interval)
    records = fetch_records(client, emp_ids)

    # (社員×添付日付) でグループ化して消し込む＝再実行安全。
    # 同月開始の契約が2本ある人は同じ日付のレコードが2件になるが、
    # 「ファイル済みレコード数ぶんは添付済みとみなす」ので二重添付にならない。
    from collections import defaultdict
    groups: dict[tuple[str, str], list] = defaultdict(list)
    for emp, folder, pdf, date in jobs:
        groups[(emp, _norm_date(date))].append((emp, folder, pdf, date))

    done = skip = 0
    stopped = False

    def _notify(current: str = "") -> None:
        if on_progress is not None:
            try:
                on_progress(done, skip, len(jobs), current)
            except Exception:  # noqa: BLE001 — 進捗表示の失敗で添付を止めない
                pass

    for (emp, date_n), job_list in groups.items():
        if stopped:
            break
        rows = records.get(emp, [])
        same_date = [r for r in rows if r["date"] == date_n]
        attached_n = sum(1 for r in same_date if r["attached"])
        empty = [r for r in same_date if not r["attached"]]
        # 添付済みレコードの数だけ「済み」として消し込む（ファイル名はAPIから見えないため数で対応）
        skip += min(attached_n, len(job_list))
        _notify()
        remaining = job_list[min(attached_n, len(job_list)):]
        for emp_, folder, pdf, date in remaining:
            if should_stop is not None and should_stop():
                stopped = True
                _say("  キャンセル指示を受けて停止（再実行すれば添付済みスキップで続きから）")
                break
            plan = f"既存レコード{empty[0]['id']}に添付" if empty else "レコード作成→添付"
            if dry_run:
                _say(f"  [DRY] {emp_} {pdf.name} 添付日付={date} → {plan}")
                continue
            try:
                if empty:
                    rec_id = empty.pop(0)["id"]
                else:
                    _write_with_retry(client, lambda: post_date_record(client, emp_, date),
                                      f"{emp_} 日付レコード作成")
                    fresh = fetch_records(client, [emp_]).get(emp_, [])
                    known = {r["id"] for r in rows}
                    cands = [r for r in fresh if r["date"] == date_n and not r["attached"] and r["id"] not in known]
                    if not cands:
                        _log(f"NG {emp_} {pdf.name} 作成したレコードが見つからない")
                        _say(f"  ✗ {emp_} {pdf.name}: 作成したレコードをGETで特定できず → スキップ")
                        continue
                    rec_id = cands[0]["id"]
                    rows.append({"id": rec_id, "date": date_n, "attached": True})
                _write_with_retry(client, lambda: attach_file(client, emp_, rec_id, pdf),
                                  f"{emp_} {pdf.name} 添付")
                done += 1
                _log(f"OK {emp_} {pdf.name} 日付={date} record={rec_id}")
                _notify(pdf.name)
                if done % 10 == 0:
                    _say(f"  … {done}件添付済み（残り約{len(jobs) - done - skip}件）")
            except Exception as exc:                  # 1件の失敗で全体を止めない
                _log(f"NG {emp_} {pdf.name} {exc}")
                _say(f"  ✗ {emp_} {pdf.name}: {exc}")
    _notify()
    mode = "DRY-RUN" if dry_run else "実行"
    _say(f"[{mode}] 添付 {done} / スキップ（添付済み） {skip} / 対象 {len(jobs)}")
    if not dry_run:
        _say("※ async/files は非同期のため、15分後に --verify で反映を確認すること")
    return done, skip, len(jobs)


def verify(employees: list[str] | None = None) -> None:
    """PDFフォルダと jinjer 側レコードを突き合わせ、未反映を一覧する。"""
    emp_filter = set(employees) if employees else None
    jobs, _ = scan_pdfs(employees=emp_filter)
    emp_ids = sorted({emp for emp, *_ in jobs})
    client = _client(25.0)
    records = fetch_records(client, emp_ids)
    from collections import defaultdict
    groups: dict[tuple[str, str], list] = defaultdict(list)
    for emp, folder, pdf, date in jobs:
        groups[(emp, _norm_date(date))].append((emp, pdf, date))
    ok = missing = 0
    for (emp, date_n), job_list in groups.items():
        attached_n = sum(1 for r in records.get(emp, [])
                         if r["date"] == date_n and r["attached"])
        ok += min(attached_n, len(job_list))
        for emp_, pdf, date in job_list[min(attached_n, len(job_list)):]:
            missing += 1
            _say(f"  未反映: {emp_} {pdf.name}（添付日付={date}）")
    _say(f"[verify] 反映済み {ok} / 未反映 {missing} / 対象 {len(jobs)}")
    if missing:
        _say("※ 添付は非同期処理（実測で最大35分遅れ）。実行直後の未反映は時間を置いて再確認")
