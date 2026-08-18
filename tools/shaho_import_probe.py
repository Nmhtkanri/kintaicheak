# -*- coding: utf-8 -*-
r"""標報投入を始める前の下調べ（**読み取りだけ**。1件も書きません）。

Phase 0 の確認用。これを通してから画面で投入を有効にする。

    python tools/shaho_import_probe.py --emp 2024027 2011001
    python tools/shaho_import_probe.py --pdf "Z:\...\7月保険料一覧表（8月給与控除分）.pdf"
    python tools/shaho_import_probe.py --templates

確かめたいこと:

1. **基準年月の意味**。PDFは「令和08年07月分（08月給与控除分）」と書いてある。
   これを jinjer の報酬月額に入れるとき year=2026 / month=07 でよいか
   （＝既存レコードの `collection_month` が 2026-08 になっているか）を実データで見る。
2. **最終更新種別**（0:自動登録 / 1:管理者登録 / 2:随時改定 / 3:定時改定）の実値。
   手入力された値かどうかがここで分かる。
3. **一括登録テンプレート**に社会保険系があるか（`--templates`）。
   あれば126名を1リクエストで送れる経路（POST /v1/jinji-imports）に切り替えられる。
4. `--pdf` を付けると、PDFに載っている人の現在値をまとめて並べる（突合の下見）。
"""

from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv  # noqa: E402

load_dotenv()

from config import Config  # noqa: E402
from services.jinjer_api_client import JinjerAPIError  # noqa: E402


def show_remunerations(client, emps, year="", month=""):
    data = client.get_monthly_remunerations(emps, year=year, month=month)
    if not data:
        print("  （報酬月額のレコードがありません）")
        return data
    for emp in sorted(data):
        print(f"\n  ── {emp} ──")
        for rec in sorted(data[emp], key=lambda r: (r.get("year", ""), r.get("month", ""))):
            last = rec.get("last_update") or {}
            cls = (last.get("classification") or {})
            updater = (last.get("updater") or {})
            print("   基準 {y}-{m} / 徴収 {c} / 健保 {k} / 厚年 {n} / {cn}({ci}) {u}".format(
                y=rec.get("year"), m=rec.get("month"),
                c=rec.get("collection_month") or "—",
                k=(rec.get("health_insurance") or {}).get("fee") or "—",
                n=(rec.get("employee_pension") or {}).get("fee") or "—",
                cn=cls.get("name") or "—", ci=cls.get("id") or "—",
                u=updater.get("name") or updater.get("employee_id") or ""))
    return data


def check_collection_month(data, expect_pay_ym: str, target_ym: str) -> None:
    """「基準年月＝保険料の対象月」で合っているかを徴収年月から裏取りする。"""
    print("\n[基準年月の意味]")
    hits = []
    for recs in data.values():
        for rec in recs:
            try:
                ym = f"{int(rec['year']):04d}-{int(rec['month']):02d}"
            except (KeyError, TypeError, ValueError):
                continue
            if ym == target_ym and rec.get("collection_month"):
                hits.append(str(rec["collection_month"]))
    if not hits:
        print(f"  {target_ym} のレコードが無く、確認できませんでした")
        return
    uniq = sorted(set(hits))
    if uniq == [expect_pay_ym]:
        print(f"  ✅ 基準 {target_ym} → 徴収 {expect_pay_ym}。"
              "PDFの「○月分（翌月給与控除分）」と一致します。この指定で投入して問題ありません")
    else:
        print(f"  ⚠ 基準 {target_ym} のレコードの徴収年月は {'、'.join(uniq)} で、"
              f"PDFの {expect_pay_ym} と違います。投入前に谷津さんへ確認してください")


def list_templates(client) -> None:
    """一括登録テンプレートの一覧（社会保険系があるかを見る）。"""
    import requests

    print("\n[一括登録テンプレート]")
    url = f"{client.base_url}/v1/master/jinji-import-templates"
    try:
        res = requests.get(url, headers=client._auth_headers(), timeout=30)
    except requests.RequestException as e:
        print(f"  取得できませんでした: {e}")
        return
    if res.status_code != 200:
        print(f"  取得できませんでした (status={res.status_code}): {res.text[:200]}")
        return
    items = res.json().get("data", []) or []
    print(f"  {len(items)} 件")
    keywords = ("社会保険", "標準報酬", "報酬月額", "社保", "保険")
    for item in items:
        name = str(item.get("name") or "")
        menu = ((item.get("menu") or {}).get("name")
                if isinstance(item.get("menu"), dict) else item.get("menu"))
        mark = " ★" if any(k in name for k in keywords) else ""
        print(f"   id={item.get('id')} / {menu} / {name}{mark}")
    print("  ★ が付いたものがあれば、CSV一括投入（POST /v1/jinji-imports）で"
          "1リクエストにまとめられます（126名でも25秒×126が不要になる）")


def main() -> int:
    parser = argparse.ArgumentParser(description="標報投入の下調べ（読み取り専用）")
    parser.add_argument("--emp", nargs="*", default=[], help="社員番号（複数可）")
    parser.add_argument("--pdf", default="", help="保険料一覧表PDF（載っている人を対象にする）")
    parser.add_argument("--year", default="", help="基準年 YYYY（省略時は全件）")
    parser.add_argument("--month", default="", help="基準月 MM（year 必須）")
    parser.add_argument("--templates", action="store_true",
                        help="一括登録テンプレートの一覧も出す")
    parser.add_argument("--json", default="", help="取得結果をJSONで保存するパス")
    args = parser.parse_args()

    from services.keiri_api import get_client

    emps = list(args.emp)
    target_ym = pay_ym = ""
    if args.pdf:
        from services.shaho_pdf import read_pdf, verify_totals
        stmt = read_pdf(args.pdf, expected_office=Config.SHAHO_IMPORT_EXPECTED_OFFICE)
        checked = verify_totals(stmt)
        target_ym, pay_ym = stmt.target_ym, stmt.pay_ym
        print(f"[PDF] {target_ym} 分（{pay_ym} 給与控除）／{len(stmt.persons)}名／"
              f"事業所 {stmt.office_code}:{stmt.office_name}")
        print(f"  チェックサム照合 OK: {'、'.join(checked)}")
        emps += [p.emp for p in stmt.persons]

    if not emps:
        parser.error("--emp か --pdf のどちらかを指定してください")

    try:
        client = get_client()
        print(f"\n[jinjer] 認証OK（{client.base_url}）")
        if args.templates:
            list_templates(client)
        print(f"\n[報酬月額] {len(set(emps))}名ぶんを取得します（GETのみ）")
        data = show_remunerations(client, sorted(set(emps)), args.year, args.month)
        if target_ym:
            check_collection_month(data, pay_ym, target_ym)
        if args.json:
            os.makedirs(os.path.dirname(args.json) or ".", exist_ok=True)
            with open(args.json, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"\n保存しました: {args.json}")
    except JinjerAPIError as e:
        print(f"\njinjer API エラー: {e}", file=sys.stderr)
        return 1
    print("\n※ このツールは1件も書き込んでいません。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
