# -*- coding: utf-8 -*-
"""四半期分の台帳を生成する本体（__main__.cmd_build から抽出。2026-08-28 ハブ移設）。

CLI とハブ画面の両方から呼ぶ。処理内容・警告文言は抽出前の cmd_build と同一に保つこと
（一覧CSV・警告CSVのバイト一致を移設検証の基準にしているため）。
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from . import config
from .estaffing import contracts_in_quarter, load_contracts
from .inputs import (PATTERN_CPI, PATTERN_FG_DETAILS, PATTERN_FG_ERICSSON,
                     PATTERN_FG_UAL_REPORT, PATTERN_FG_WO, PATTERN_ROSTER,
                     PATTERN_TC, newest, newest_or_none)
from .records import build_record
from .roster import load_roster
from .template import build_template
from .writer import write_quarter


def build_quarter(quarter: str, *, tc=None, cpi=None, roster=None, template=None,
                  out_dir=None, fg=None, fg_details=None, no_fg: bool = False,
                  jinjer_api: bool = False, fg_mode: str = "auto") -> dict:
    """台帳xlsx・一覧CSV・警告CSVを out_dir へ生成し、結果サマリの dict を返す。

    入力が見つからないときは FileNotFoundError（CLI 側で SystemExit に変換する）。
    fg_mode: auto=新レポート（fieldglass_report）が input にあればそれを使う /
             report=新レポート必須 / legacy=旧 WO CSV＋details JSON（移設前と同一出力）。
    エリクソンの直接契約上書きも legacy では行わない（比較検証のため）。
    戻り値: {quarter, label, counts, match, n_warn, paths, inputs, global_warnings,
             notes, summary}
    """
    q = quarter.upper()
    q_start, q_end = config.quarter_range(q)
    tc = Path(tc) if tc else newest(config.INPUT_DIR, PATTERN_TC)
    cpi = Path(cpi) if cpi else newest(config.INPUT_DIR, PATTERN_CPI)
    roster_path = Path(roster) if roster else newest(config.INPUT_DIR, PATTERN_ROSTER)
    template = Path(template) if template else config.TEMPLATE_XLSX
    out_dir = Path(out_dir) if out_dir else config.OUTPUT_DIR
    notes: list[str] = []
    if not template.exists():
        notes.append(f"テンプレートが無いので作成します: {template}")
        build_template(config.SOURCE_FORM_XLSM, template)

    started = dt.datetime.now()
    contracts, load_warnings = load_contracts(tc, cpi)
    roster = load_roster(roster_path)
    # jinjer API の人マスタ（退職者込み）を合流: jinjer_api=True で再取得、無指定ならキャッシュがあれば使う
    from .jinjer_api import CACHE_PATH, load_cache, merge_into_roster, refresh_cache
    api_note = ""
    if jinjer_api:
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
    if fg_mode not in ("auto", "legacy", "report"):
        raise ValueError(f"fg_mode は auto / legacy / report のいずれか: {fg_mode!r}")
    fg_records: list = []
    fg_note = ""
    fg_path = None
    det_path = None
    report_path = None
    use_report = False
    fg_src_name = ""
    if not no_fg:
        if fg_mode != "legacy" and not fg:
            report_path = newest_or_none(config.INPUT_DIR, PATTERN_FG_UAL_REPORT)
        use_report = fg_mode == "report" or (fg_mode == "auto" and report_path is not None)
        if fg_mode == "report" and report_path is None:
            raise FileNotFoundError(
                f"入力が見つかりません: {config.INPUT_DIR}\\{PATTERN_FG_UAL_REPORT}")
        if not use_report:
            report_path = None
            if fg:
                fg_path = Path(fg)
            else:
                fg_path = newest_or_none(config.INPUT_DIR, PATTERN_FG_WO)
    if use_report or fg_path is not None:
        from .fieldglass import (FG_CARRIED_FIELDS, apply_defaults, contract_from_workorder,
                                 derive_defaults, detail_has_schedule, fill_from_person,
                                 load_details, load_workorders, name_candidates, workorders_in_quarter)
        from .roster import normalize_name
        det_by_staff: dict = {}
        det_by_wo: dict = {}
        if use_report:
            from .fieldglass_report import load_ual_report
            wos_all, det_by_staff, det_by_wo = load_ual_report(report_path)
            fg_src_name = report_path.name
        else:
            if fg_details:
                det_path = Path(fg_details)
            else:
                det_path = newest_or_none(config.INPUT_DIR, PATTERN_FG_DETAILS)
            if det_path is not None:
                det_by_staff, det_by_wo = load_details(det_path)
            wos_all = load_workorders(fg_path)
            fg_src_name = fg_path.name
        wos = workorders_in_quarter(wos_all, q_start, q_end)
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
                if use_report and wo.end is None:
                    # 新レポートは「最新の終了日」が空のWOがまれにある（実測: 即日終了の1件）。
                    # 旧CSVには実終了日が入っていたため、継続中扱いで期末まで延ばした事実を人に見せる
                    rec.warnings.append("FG: レポートの終了日が空（継続中として期末まで延長）"
                                        "→ 中止・即日終了のWOでないか実際の終了日を要確認")
                if any(k in es_hit_names for k in cands):
                    rec.warnings.append("FG: 同じ四半期に e-staffing 契約もある（二重か、2派遣先か要確認）")
                rec.fields["備考"] = (f"SAP Fieldglass WO {wo.wo_id} 改訂{wo.revision}（求人情報 {wo.job_posting_id}）"
                                     f"／{src_note}／派遣許可番号 {rec.fields.get('派遣許可番号', '')}"
                                     f"／作成 {started:%Y/%m/%d %H:%M}")
                if detail and (detail.get("monthlyStdLower") or detail.get("monthlyStdUpper")):
                    rec.fields["備考"] += (f"／月間標準時間 {detail.get('monthlyStdLower', '')}"
                                          f"〜{detail.get('monthlyStdUpper', '')}時間")
                fg_records.append(rec)
            records += fg_records
            fg_note = (f"Fieldglass: {fg_src_name} → WO {len(wos)}件（SAP詳細あり {n_detail} / 引き継ぎ元あり {carried} / 元なし {len(wos) - carried}）。"
                       f"責任者・業務内容等はSAP詳細の現行値、36協定・保険・抵触日（と詳細の無い分）は直近e-staffing契約からの引き継ぎ"
                       + (f"／詳細: {det_path.name}" if det_path is not None
                          else ("／新レポート1本読み（詳細JSONなし運用）" if use_report else "／SAP詳細JSONなし")))

    # --- 直接契約（紙/Excel契約 → マスタCSV）---
    direct_records: list = []
    from .direct import contract_from_row, load_master, rows_in_quarter
    d_master = load_master()
    d_hit = rows_in_quarter(d_master, q_start, q_end)
    # エリクソンの新レポートがあれば、直接契約マスタの該当行へ現行値を上書きする
    # （レポートに期間・WOが無いため行は増やさない。legacy モードでは行わない＝比較検証用）
    eric_note = ""
    if d_hit and fg_mode != "legacy":
        eric_path = newest_or_none(config.INPUT_DIR, PATTERN_FG_ERICSSON)
        if eric_path is not None:
            from .fieldglass_report import apply_ericsson_report, load_ericsson_report
            n_upd, unmatched = apply_ericsson_report(
                load_ericsson_report(eric_path), d_hit, source_name=eric_path.name)
            eric_note = (f"エリクソン: {eric_path.name} の現行値で直接契約 {n_upd}行を更新"
                         + (f"／突合できず: {'・'.join(n for n in unmatched if n)}" if unmatched else ""))
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
            if row.get("_エリクソン注記"):
                rec.warnings.append(f"直接契約: {row['_エリクソン注記']}")
            if _nn(sei, mei) in covered_names:
                rec.warnings.append("直接契約: 同じ四半期に e-staffing/Fieldglass の契約もある（重複か2契約か要確認）")
            direct_records.append(rec)
        records += direct_records

    global_warnings = [f"入力: 契約データ={tc.name} / 契約書・通知書データ={cpi.name} / 従業員一覧={roster_path.name}"
                       + (f" / Fieldglass={fg_src_name}" if fg_records else "")
                       + (" / 直接契約マスタ" if direct_records else "")]
    if direct_records:
        global_warnings.append(f"直接契約: マスタ {len(d_master)}行のうち {len(d_hit)}行がこの四半期に該当")
    if api_note:
        global_warnings.append(api_note)
    if fg_note:
        global_warnings.append(fg_note)
    if eric_note:
        global_warnings.append(eric_note)
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
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    with (config.LOG_DIR / "実行ログ.txt").open("a", encoding="utf-8") as fp:
        fp.write(f"{started:%Y-%m-%d %H:%M:%S} {summary}\n")
    return {
        "quarter": q,
        "label": config.quarter_label(q),
        "counts": {"total": len(records), "estaffing": len(hit),
                   "fieldglass": len(fg_records), "direct": len(direct_records),
                   "people": len(people)},
        "match": dict(match),
        "n_warn": n_warn,
        "paths": {k: str(v) for k, v in paths.items()},
        "inputs": {"tc": tc.name, "cpi": cpi.name, "roster": roster_path.name,
                   "fg": fg_src_name, "fg_mode": ("report" if use_report else "legacy"),
                   "fg_details": det_path.name if det_path is not None else ""},
        "global_warnings": global_warnings,
        "notes": notes,
        "summary": summary,
    }
