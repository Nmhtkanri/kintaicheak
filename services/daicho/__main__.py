# -*- coding: utf-8 -*-
"""コマンドライン。

  python -X utf8 -m services.daicho make-template
  python -X utf8 -m services.daicho build --quarter 2025Q3
      [--tc TCnmht*.csv] [--cpi CPInmht*.csv] [--roster 従業員一覧*.xlsx] [--template ...] [--out-dir ...]

入力を省略すると Z:\\派遣元管理台帳\\input の最新ファイルを使う。
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

from dotenv import load_dotenv

# .env（JINJER_API_KEY 等・DAICHO_DATA_ROOT）を config の DATA_ROOT 確定前に読む
load_dotenv(Path(__file__).resolve().parents[2] / ".env")

from . import config  # noqa: E402
from .estaffing import load_contracts  # noqa: E402
from .roster import load_roster  # noqa: E402
from .template import build_template  # noqa: E402


def _newest(folder: Path, pattern: str) -> Path:
    from .inputs import newest
    try:
        return newest(folder, pattern)
    except FileNotFoundError as e:
        raise SystemExit(str(e))


def cmd_make_template(args) -> int:
    src = Path(args.source) if args.source else config.SOURCE_FORM_XLSM
    out = Path(args.out) if args.out else config.TEMPLATE_XLSX
    if not src.exists():
        raise SystemExit(f"旧フォームが見つかりません: {src}")
    build_template(src, out)
    print(f"テンプレートを作成: {out}")
    return 0


def cmd_build(args) -> int:
    from .build import build_quarter
    try:
        result = build_quarter(
            args.quarter, tc=args.tc, cpi=args.cpi, roster=args.roster,
            template=args.template, out_dir=args.out_dir, fg=args.fg,
            fg_details=args.fg_details, no_fg=args.no_fg, jinjer_api=args.jinjer_api)
    except FileNotFoundError as e:
        raise SystemExit(str(e))
    for note in result["notes"]:
        print(note)
    print(result["summary"])
    return 0


def cmd_fg_check(args) -> int:
    """Fieldglass Work Order の取り込み検証レポート（台帳はまだ作らない）。"""
    import csv as _csv

    from .fieldglass import load_workorders, name_candidates, workorders_in_quarter
    from .roster import normalize_name

    q = args.quarter.upper()
    q_start, q_end = config.quarter_range(q)
    wo_path = Path(args.wo) if args.wo else _newest(config.INPUT_DIR, "*WorkOrder*.csv")
    tc = Path(args.tc) if args.tc else _newest(config.INPUT_DIR, "TCnmht*.csv")
    cpi = Path(args.cpi) if args.cpi else _newest(config.INPUT_DIR, "CPInmht*.csv")
    roster_path = Path(args.roster) if args.roster else _newest(config.INPUT_DIR, "従業員一覧*.xlsx")

    wos = workorders_in_quarter(load_workorders(wo_path), q_start, q_end)
    roster = load_roster(roster_path)
    contracts, _ = load_contracts(tc, cpi)
    by_name: dict[str, list] = {}
    for c in contracts:
        sei, mei = c.worker_sei_mei
        by_name.setdefault(normalize_name(sei, mei), []).append(c)
    for v in by_name.values():
        v.sort(key=lambda c: (c.end or dt.date.min), reverse=True)

    out_dir = Path(args.out_dir) if args.out_dir else config.OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"FG検証_{q}.csv"
    counts = {"OK": 0, "引継ぎ元なし": 0, "jinjer不一致": 0}
    with out.open("w", encoding="utf-8-sig", newline="") as fp:
        w = _csv.writer(fp)
        w.writerow(["判定", "氏名(WO表記)", "社員番号", "名寄せ", "WO_ID", "改訂", "WO期間",
                    "勤務地", "事業単位", "Supervisor", "引継ぎ元契約No", "引継ぎ元期間", "引継ぎ元派遣先", "備考"])
        for wo_ in wos:
            cands = name_candidates(wo_.staff_raw)
            person = None
            how = ""
            for i, k in enumerate(cands):
                hits = roster.by_name.get(k, [])
                if len(hits) == 1:
                    person, how = hits[0], ("そのまま" if i == 0 else "姓名逆で一致")
                    break
            src = next((by_name[k][0] for k in cands if k in by_name), None)
            if person is None:
                verdict = "jinjer不一致"
            elif src is None:
                verdict = "引継ぎ元なし"
            else:
                verdict = "OK"
            counts[verdict] += 1
            note = "" if src else "e-staffing履歴なし（4月以降の新規配属？契約書面の確認が必要）"
            w.writerow([verdict, wo_.staff_display, person.emp_id if person else "", how,
                        wo_.wo_id, wo_.revision,
                        f"{wo_.start or ''}～{wo_.end or '(継続中)'}", wo_.site, wo_.business_unit,
                        "；".join(wo_.supervisors),
                        src.contract_no if src else "", f"{src.start}～{src.end}" if src else "",
                        src.client_name if src else "", note])
    print(f"[{q}] Fieldglass WO {len(wos)}件 → {dict(counts)}\n  検証レポート: {out}")
    return 0


def cmd_extract_direct(args) -> int:
    """契約書フォルダから直接契約マスタ（自動）を再生成する。"""
    from .direct import AUTO_CSV
    from .extract_direct import CONTRACT_BASE, extract_all

    n, warns = extract_all(base=args.base or CONTRACT_BASE, out_csv=args.out or AUTO_CSV,
                           debug=args.debug)
    print(f"直接契約マスタ_自動: {n}行 → {args.out or AUTO_CSV}")
    for w in warns:
        print("  ⚠", w)
    return 0


def cmd_export_pdf(args) -> int:
    """四半期台帳ブック → 人別フォルダへ契約ごとのPDF。"""
    from .export_pdf import PDF_ROOT, export_quarter

    quarters = [q.upper() for q in (args.quarter or [])]
    if args.all:
        quarters = ["2025Q3", "2025Q4", "2026Q1", "2026Q2"]
    if not quarters:
        raise SystemExit("--quarter を指定（複数可）するか --all を付けてください")
    total = 0
    for q in quarters:
        n, warns = export_quarter(q, xlsx_path=args.xlsx, pdf_root=args.out_dir or PDF_ROOT)
        total += n
        print(f"[{q}] PDF {n}枚 → {args.out_dir or PDF_ROOT}")
        for w in warns:
            print("  ⚠", w)
    print(f"合計 {total}枚")
    return 0


def cmd_attach_jinjer(args) -> int:
    """PDFフォルダ → jinjer カスタム項目「派遣元管理台帳」(menu 16) へ添付。"""
    from .jinjer_attach import run, verify

    if args.verify:
        verify(employees=args.employee or None)
        return 0
    run(employees=args.employee or None, dry_run=not args.execute,
        write_interval=args.interval, limit=args.limit)
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="daicho", description="派遣元管理台帳ジェネレーター")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("make-template", help="旧フォームからテンプレート .xlsx を作る")
    p1.add_argument("--source")
    p1.add_argument("--out")
    p1.set_defaults(func=cmd_make_template)
    p2 = sub.add_parser("build", help="四半期分の台帳を生成する")
    p2.add_argument("--quarter", required=True, help="例: 2025Q3（7-9月期）")
    p2.add_argument("--tc")
    p2.add_argument("--cpi")
    p2.add_argument("--roster")
    p2.add_argument("--template")
    p2.add_argument("--out-dir")
    p2.add_argument("--fg", help="Fieldglass WorkOrder CSV（省略時は input の最新 *WorkOrder*.csv）")
    p2.add_argument("--fg-details", help="SAP詳細JSON（省略時は input の最新 *fieldglass_details*.json）")
    p2.add_argument("--no-fg", action="store_true", help="Fieldglass 分を含めない")
    p2.add_argument("--jinjer-api", action="store_true",
                    help="jinjer API から退職者込みの人マスタを再取得してキャッシュ（無指定ならキャッシュを使う）")
    p2.set_defaults(func=cmd_build)
    p3 = sub.add_parser("fg-check", help="Fieldglass Work Order の取り込み検証レポートを出す")
    p3.add_argument("--quarter", required=True)
    p3.add_argument("--wo")
    p3.add_argument("--tc")
    p3.add_argument("--cpi")
    p3.add_argument("--roster")
    p3.add_argument("--out-dir")
    p3.set_defaults(func=cmd_fg_check)
    p4 = sub.add_parser("extract-direct", help="契約書フォルダから直接契約マスタ_自動.csv を再生成")
    p4.add_argument("--base")
    p4.add_argument("--out")
    p4.add_argument("--debug", action="store_true")
    p4.set_defaults(func=cmd_extract_direct)
    p6 = sub.add_parser("attach-jinjer", help="PDFを jinjer カスタム項目「派遣元管理台帳」へ添付する")
    p6.add_argument("--execute", action="store_true", help="実書き込み（既定はdry-run）")
    p6.add_argument("--employee", action="append", help="社員番号で絞る（複数可）")
    p6.add_argument("--interval", type=float, default=12.0, help="書き込み間隔秒（429が出たら25へ）")
    p6.add_argument("--limit", type=int, help="先頭N件だけ処理")
    p6.add_argument("--verify", action="store_true", help="添付の反映状況をGETで突き合わせる")
    p6.set_defaults(func=cmd_attach_jinjer)
    p5 = sub.add_parser("export-pdf", help="台帳ブックから人別フォルダへ契約ごとのPDFを書き出す")
    p5.add_argument("--quarter", action="append", help="例: 2025Q3（複数指定可）")
    p5.add_argument("--all", action="store_true", help="4四半期すべて")
    p5.add_argument("--xlsx", help="ブックを直接指定（通常は不要）")
    p5.add_argument("--out-dir", help="出力先（既定: Z:\\派遣元管理台帳\\PDF）")
    p5.set_defaults(func=cmd_export_pdf)
    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
