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
from .estaffing import contracts_in_quarter, load_contracts
from .records import build_record
from .roster import load_roster
from .template import build_template
from .writer import write_quarter


def _newest(folder: Path, pattern: str) -> Path:
    files = sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise SystemExit(f"入力が見つかりません: {folder}\\{pattern}")
    return files[0]


def cmd_make_template(args) -> int:
    src = Path(args.source) if args.source else config.SOURCE_FORM_XLSM
    out = Path(args.out) if args.out else config.TEMPLATE_XLSX
    if not src.exists():
        raise SystemExit(f"旧フォームが見つかりません: {src}")
    build_template(src, out)
    print(f"テンプレートを作成: {out}")
    return 0


def cmd_build(args) -> int:
    q = args.quarter.upper()
    q_start, q_end = config.quarter_range(q)
    tc = Path(args.tc) if args.tc else _newest(config.INPUT_DIR, "TCnmht*.csv")
    cpi = Path(args.cpi) if args.cpi else _newest(config.INPUT_DIR, "CPInmht*.csv")
    roster_path = Path(args.roster) if args.roster else _newest(config.INPUT_DIR, "従業員一覧*.xlsx")
    template = Path(args.template) if args.template else config.TEMPLATE_XLSX
    out_dir = Path(args.out_dir) if args.out_dir else config.OUTPUT_DIR
    if not template.exists():
        print(f"テンプレートが無いので作成します: {template}")
        build_template(config.SOURCE_FORM_XLSM, template)

    started = dt.datetime.now()
    contracts, load_warnings = load_contracts(tc, cpi)
    roster = load_roster(roster_path)
    # jinjer API の人マスタ（退職者込み）を合流: --jinjer-api で再取得、無指定ならキャッシュがあれば使う
    from .jinjer_api import CACHE_PATH, load_cache, merge_into_roster, refresh_cache
    api_note = ""
    if getattr(args, "jinjer_api", False):
        people = refresh_cache()
        api_note = merge_into_roster(roster, people) + f"（APIから再取得→{CACHE_PATH.name}）"
    else:
        people, fetched_at = load_cache()
        if people:
            api_note = merge_into_roster(roster, people) + f"（キャッシュ {fetched_at}。更新は --jinjer-api）"
        else:
            api_note = "jinjer API 人マスタ未使用（--jinjer-api で退職者込みに取得できる）"
    hit = contracts_in_quarter(contracts, q_start, q_end)
    records = []
    match = {"ok": 0, "none": 0, "ambiguous": 0}
    for c in hit:
        sei, mei = c.worker_sei_mei
        person, state = roster.find(sei, mei)
        match[state] += 1
        records.append(build_record(c, person, state, q_start, q_end, generated_at=started))

    # --- SAP Fieldglass（2026年4月以降のユニアデックス分）---
    fg_records: list = []
    fg_note = ""
    fg_path = None
    if not getattr(args, "no_fg", False):
        if getattr(args, "fg", None):
            fg_path = Path(args.fg)
        else:
            fg_files = sorted(config.INPUT_DIR.glob("*WorkOrder*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
            fg_path = fg_files[0] if fg_files else None
    if fg_path is not None:
        from .fieldglass import (FG_CARRIED_FIELDS, apply_defaults, contract_from_workorder,
                                 derive_defaults, detail_has_schedule, fill_from_person,
                                 load_details, load_workorders, name_candidates, workorders_in_quarter)
        from .roster import normalize_name
        det_by_staff: dict = {}
        det_by_wo: dict = {}
        det_path = None
        if getattr(args, "fg_details", None):
            det_path = Path(args.fg_details)
        else:
            det_files = sorted(config.INPUT_DIR.glob("*fieldglass_details*.json"),
                               key=lambda p: p.stat().st_mtime, reverse=True)
            det_path = det_files[0] if det_files else None
        if det_path is not None:
            det_by_staff, det_by_wo = load_details(det_path)
        wos = workorders_in_quarter(load_workorders(fg_path), q_start, q_end)
        if wos:
            latest_by_name: dict = {}
            for c in contracts:
                k = normalize_name(*c.worker_sei_mei)
                cur = latest_by_name.get(k)
                if cur is None or (c.end or dt.date.min) > (cur.end or dt.date.min):
                    latest_by_name[k] = c
            es_hit_names = {normalize_name(*c.worker_sei_mei) for c in hit}
            carried = 0
            n_detail = 0
            defaults = None
            for wo in wos:
                cands = name_candidates(wo.staff_raw)
                person, state = None, "none"
                for k in cands:
                    person, state = roster.find_by_key(k)
                    if person is not None:
                        break
                match[state if state in match else "none"] += 1
                src = next((latest_by_name[k] for k in cands if k in latest_by_name), None)
                detail = det_by_staff.get(wo.staff_fg_id) or det_by_wo.get(wo.wo_id)
                vc = contract_from_workorder(wo, src, q_end, person, detail)
                default_filled: list = []
                if src is None:
                    if defaults is None:
                        defaults = derive_defaults(contracts, q_start, q_end)
                    default_filled = apply_defaults(vc, defaults)
                    fill_from_person(vc, person, wo.start)
                rec = build_record(vc, person, state, q_start, q_end, generated_at=started)
                sap_fields = ""
                if detail:
                    n_detail += 1
                    sap_fields = "派遣先責任者・苦情申出先・責任の程度・業務内容" + (
                        "・就業時間・休憩・休日" if detail_has_schedule(detail) else "")
                    rec.warnings.append(f"FG: {sap_fields} は SAP Fieldglass 詳細（Job Posting {detail.get('jobPostingId', '')}）の現行値")
                    if src is not None:
                        for label, old, new in (
                                ("派遣先責任者", src.cpi.get("派遣先責任者 氏名", ""),
                                 (detail.get("clientResponsible") or {}).get("name", "")),
                                ("苦情申出先", src.cpi.get("派遣先苦情申出先 氏名", ""),
                                 (detail.get("complaintRecipient") or {}).get("name", ""))):
                            if old and new and normalize_name(old) != normalize_name(new):
                                rec.warnings.append(f"FG: {label}が3月契約から変更されている: {old} → {new}（台帳はSAP側を採用）")
                if src is not None:
                    carried += 1
                    carried_fields = "36協定・保険・抵触日" + (
                        "" if detail_has_schedule(detail) else "・就業時間・休憩・休日") + (
                        "" if detail else f"・{FG_CARRIED_FIELDS}")
                    src_note = (f"直近e-staffing契約 {src.contract_no}（{src.start}～{src.end}）から"
                                f"{carried_fields}を引き継ぎ" + (f"／{sap_fields}はSAP詳細の現行値" if detail else ""))
                    if not detail:
                        rec.warnings.append(f"FG: {FG_CARRIED_FIELDS} は3月以前の契約からの引き継ぎ値（SAP詳細なし＝終了済みWO等）")
                else:
                    src_note = ("引き継ぎ元の e-staffing 契約なし。36協定・保険・抵触日・派遣元責任者等は"
                                "e-staffing全体の標準値・jinjerで補完"
                                + (f"／{sap_fields}はSAP詳細の現行値" if detail else ""))
                    rec.warnings.append(
                        f"FG: 引き継ぎ元なし（4月以降の新規配属）→ {len(default_filled)}項目を"
                        "e-staffing標準値・jinjer属性で補完＝契約書面での確認推奨")
                if len(wo.supervisors) > 1:
                    rec.warnings.append(f"FG: スーパーバイザが{len(wo.supervisors)}名登録 → 指揮命令者の特定要確認")
                if any(k in es_hit_names for k in cands):
                    rec.warnings.append("FG: 同じ四半期に e-staffing 契約もある（二重か、2派遣先か要確認）")
                rec.fields["備考"] = (f"SAP Fieldglass WO {wo.wo_id} 改訂{wo.revision}（求人情報 {wo.job_posting_id}）"
                                     f"／{src_note}／派遣許可番号 {rec.fields.get('派遣許可番号', '')}"
                                     f"／作成 {started:%Y/%m/%d %H:%M}")
                fg_records.append(rec)
            records += fg_records
            fg_note = (f"Fieldglass: {fg_path.name} → WO {len(wos)}件（SAP詳細あり {n_detail} / 引き継ぎ元あり {carried} / 元なし {len(wos) - carried}）。"
                       f"責任者・業務内容等はSAP詳細の現行値、36協定・保険・抵触日（と詳細の無い分）は直近e-staffing契約からの引き継ぎ"
                       + (f"／詳細: {det_path.name}" if det_path is not None else "／SAP詳細JSONなし"))

    # --- 直接契約（紙/Excel契約 → マスタCSV）---
    direct_records: list = []
    from .direct import contract_from_row, load_master, rows_in_quarter
    d_master = load_master()
    d_hit = rows_in_quarter(d_master, q_start, q_end)
    if d_hit:
        from .fieldglass import apply_defaults as _apply_d, derive_defaults as _derive_d, fill_from_person as _fill_d
        from .roster import normalize_name as _nn
        covered_names = {_nn(*r.contract.worker_sei_mei) for r in records if r.contract}
        base_defaults = _derive_d(contracts, q_start, q_end)
        # 会社共通のものだけ補完に使う（事業所抵触日・安全衛生・法定休日は派遣先固有なので触らない）
        d_defaults = {"cpi": dict(base_defaults["cpi"]), "tc": {}}
        for row in d_hit:
            vc = contract_from_row(row)
            sei, mei = vc.worker_sei_mei
            person, state = roster.find(sei, mei)
            if person is None:   # 外国名など姓名が逆に書かれた契約書に対応
                person, state = roster.find_by_key(_nn(mei, sei))
            match[state if state in match else "none"] += 1
            filled = _apply_d(vc, d_defaults)
            _fill_d(vc, person, vc.start)
            rec = build_record(vc, person, state, q_start, q_end, generated_at=started)
            src_name = Path(row.get("出所ファイル") or "").name or "手入力"
            rec.fields["備考"] = (f"直接契約（{row.get('様式') or '?'}）／契約書: {src_name}"
                                 + (f"／限定の別: {row['限定の別']}" if row.get("限定の別") else "")
                                 + (f"／{row['備考']}" if row.get("備考") else "")
                                 + (f"／{len(filled)}項目は標準値補完（36協定・保険等）" if filled else "")
                                 + f"／作成 {started:%Y/%m/%d %H:%M}")
            rec.warnings.append(f"直接契約: 契約書（{row.get('様式') or '?'}）から抽出"
                                + (f"。{len(filled)}項目を標準値・jinjerで補完" if filled else ""))
            if _nn(sei, mei) in covered_names:
                rec.warnings.append("直接契約: 同じ四半期に e-staffing/Fieldglass の契約もある（重複か2契約か要確認）")
            direct_records.append(rec)
        records += direct_records

    global_warnings = [f"入力: 契約データ={tc.name} / 契約書・通知書データ={cpi.name} / 従業員一覧={roster_path.name}"
                       + (f" / Fieldglass={fg_path.name}" if fg_records else "")
                       + (" / 直接契約マスタ" if direct_records else "")]
    if direct_records:
        global_warnings.append(f"直接契約: マスタ {len(d_master)}行のうち {len(d_hit)}行がこの四半期に該当")
    if api_note:
        global_warnings.append(api_note)
    if fg_note:
        global_warnings.append(fg_note)
    global_warnings += load_warnings + roster.warnings
    if match["none"] or match["ambiguous"]:
        global_warnings.append(
            f"jinjer 従業員一覧と氏名で突合: 一致 {match['ok']} / 不一致 {match['none']} / 同姓同名 {match['ambiguous']}"
            "（不一致は退職者の可能性。jinjer API で退職者込みに取ると埋まる）")
    paths = write_quarter(records, q, template, out_dir, global_warnings)

    people = {r.name for r in records}
    n_warn = sum(len(r.warnings) for r in records)
    summary = (f"[{q} {config.quarter_label(q)}] 台帳 {len(records)}枚（e-staffing {len(hit)} / Fieldglass {len(fg_records)} / 直接 {len(direct_records)}） / {len(people)}人 / "
               f"jinjer一致 {match['ok']} 不一致 {match['none']} 同姓同名 {match['ambiguous']} / 台帳警告 {n_warn}件\n"
               f"  台帳: {paths['xlsx']}\n  一覧: {paths['csv']}\n  警告: {paths['warnings']}")
    print(summary)
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    with (config.LOG_DIR / "実行ログ.txt").open("a", encoding="utf-8") as fp:
        fp.write(f"{started:%Y-%m-%d %H:%M:%S} {summary}\n")
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
