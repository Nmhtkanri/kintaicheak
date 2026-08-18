r"""社労士の保険料一覧表PDF × jinjer登録値 × 当方計算値 の3点突合と、報酬月額の投入。

## トリプルチェック（谷津さんの要件・2026-08-17）

| # | 誰が | 何を見るか | このモジュールでの扱い |
|---|---|---|---|
| A | 社労士（前田事務所） | PDFの標準報酬月額 | 投入する値。**法的な正はここ** |
| B | jinjer | いま登録されている値 | 差分を出す。同じなら書かない |
| C | オペレーションハブ | 4〜6月等の給与明細から**独立に計算**した値 | A と食い違ったら投入から外す |

そのうえで **人間（谷津さん）が画面で確認して実行**する。A・B・C が揃って初めて
自動で投入対象になり、C が出せない人（資格取得時決定の新入社員など）は
「差分を確認して投入」に落ちる。**C が A と食い違う人は既定で投入しない**
（承知のうえ投入は個別に選べる）。

2026-08-17 に発覚した遠田恭実さんの手入力ミス（正 280,000 を 300,000 と登録）は、
このうち B が A・C とずれる形で現れる。

## C（当方計算）の求め方は改訂理由で変わる

標準報酬モード（`shaho_engine`）の部品をそのまま使って計算する。

| PDFの改訂理由 | 算定に使う3か月 | 採用条件 |
|---|---|---|
| 月額変更（随時改定） | 対象月の**直前3か月**（例: 7月分 → 4・5・6月支給） | 3か月すべてが支払基礎日数を満たすこと |
| 定時決定 | その年の**4〜6月支給** | 1か月でも満たせば採用（原則どおり） |
| 取得時決定 | — | 見込み額なので**計算しない** |
| 料率変更 | — | 標準報酬は動いていないので計算しない |

## 書き込みの階層とガード

標準報酬月額は給与計算の土台で、階層は「書く（マスタ級）」。本来は共有exeに
入れない決まりだが、2026-08-17 に谷津さんの判断でオペレーションハブに載せる
**例外**とした。その代わりに以下を必須にしている:

- 実行者の許可リスト（`Config.SHAHO_IMPORT_ALLOWED_USERS_CSV`）。読めなければ書けない
- dry-run 既定＋計画ハッシュ照合（見た内容と違うものを書かせない）
- 書き込み前に現在値をJSONへバックアップ
- 書き込み後に**取り直して照合**（VERIFY_MISMATCH）
- 実行台帳CSVへ新旧値つきで追記
- 同時実行ロック（テナント単位のレート制限があるので他の投入と並行させない）
"""

from __future__ import annotations

import csv
import datetime
import hashlib
import json
import os
from dataclasses import dataclass, field

from config import Config
from services.keiri_api import classify_employee
from services.keiri_engine import ym_add
from services.sap_import_ledger import can_write as _can_write_csv
from services.sap_import_ledger import current_user

# 判定ステータス（画面の並び順＝気にすべき順）。
# ⚠ shaho_check.STATUS_PRIORITY には足さないこと。あちらは REVIEW_STATUSES を
#    スライス（[:6]）で切り出しているので、要素を挿すと要確認の範囲が黙って変わる。
STATUS_ORDER = ["UNRESOLVED", "PDF_INCONSISTENT", "NOT_IN_JINJER", "CALC_MISMATCH",
                "RETIRED", "NO_CALC", "AUTO_OK", "EXCLUDED", "NO_CHANGE"]
STATUS_JA = {
    "UNRESOLVED": "要確認（本人を特定できない）",
    "PDF_INCONSISTENT": "要確認（PDFの検算NG）",
    "NOT_IN_JINJER": "要確認（jinjerに居ない）",
    "CALC_MISMATCH": "要確認（当方の計算と不一致）",
    "RETIRED": "要確認（退職済み）",
    "NO_CALC": "投入できる（計算値なし・差分確認）",
    "AUTO_OK": "投入できる（3点一致）",
    "EXCLUDED": "対象外（自社の社員番号でない）",
    "NO_CHANGE": "書込不要（登録済みと同じ）",
}
# 既定でチェックが入る＝3点そろって一致した人だけ
DEFAULT_SELECTED = frozenset({"AUTO_OK"})
# 「承知のうえ投入」を明示しないと選べない
FORCEABLE = frozenset({"CALC_MISMATCH", "RETIRED"})
# 画面で選べる（＝投入しうる）
SELECTABLE = frozenset({"AUTO_OK", "NO_CALC"}) | FORCEABLE

LEDGER_COLUMNS = ["実行日時", "実行者", "対象年月", "PDFファイル名", "社員番号", "氏名",
                  "健保前", "健保後", "厚年前", "厚年後", "操作", "結果", "検証",
                  "承知投入", "バックアップ"]


class ShahoWriteError(Exception):
    """投入を続けてはいけない状態（許可なし・計画の食い違い・ロック中など）。"""


# ---------------------------------------------------------------------------
# 計画の行
# ---------------------------------------------------------------------------

@dataclass
class PlanRow:
    emp: str
    name: str = ""                      # PDFの氏名
    jinjer_name: str = ""               # jinjer登録の氏名（別人検知用）
    pdf_kenpo: int = 0                  # A: 社労士PDF（円）
    pdf_konen: int = 0
    cur_kenpo: int | None = None        # B: 対象年月時点で有効なjinjerの値
    cur_konen: int | None = None
    cur_ym: str = ""                    # その値の基準年月
    cur_updated_by: str = ""            # 最終更新種別（管理者登録／定時改定 など）
    record_at_target: bool = False      # 対象年月ちょうどのレコードがあるか
    calc_kenpo: int | None = None       # C: 当方の計算値
    calc_konen: int | None = None
    calc_source: str = ""               # 随時改定／定時決定
    reason: str = ""                    # PDFの改訂理由
    status: str = "NO_CHANGE"
    notes: list = field(default_factory=list)

    @property
    def operation(self) -> str:
        """jinjer への操作。対象年月のレコードがあれば更新、無ければ新規登録。"""
        return "PATCH" if self.record_at_target else "POST"

    @property
    def selectable(self) -> bool:
        return self.status in SELECTABLE

    @property
    def needs_force(self) -> bool:
        return self.status in FORCEABLE

    @property
    def default_selected(self) -> bool:
        return self.status in DEFAULT_SELECTED

    def to_dict(self) -> dict:
        return {
            "emp": self.emp, "name": self.name, "jinjer_name": self.jinjer_name,
            "pdf_kenpo": self.pdf_kenpo, "pdf_konen": self.pdf_konen,
            "cur_kenpo": self.cur_kenpo, "cur_konen": self.cur_konen,
            "cur_ym": self.cur_ym, "cur_updated_by": self.cur_updated_by,
            "calc_kenpo": self.calc_kenpo, "calc_konen": self.calc_konen,
            "calc_source": self.calc_source, "reason": self.reason,
            "operation": self.operation if self.selectable else "",
            "status": self.status, "status_ja": STATUS_JA.get(self.status, self.status),
            "notes": list(self.notes), "selectable": self.selectable,
            "needs_force": self.needs_force, "default_selected": self.default_selected,
        }


# ---------------------------------------------------------------------------
# jinjer の現在値
# ---------------------------------------------------------------------------

def _rec_ym(rec) -> str:
    try:
        return f"{int((rec or {}).get('year')):04d}-{int((rec or {}).get('month')):02d}"
    except (TypeError, ValueError):
        return ""


def _rec_fee(rec, key: str) -> int | None:
    """レコードから標準報酬月額（円）を取り出す。GETは fee、書き込みは standard_fee。"""
    block = (rec or {}).get(key) or {}
    value = block.get("fee")
    if value in (None, ""):
        value = block.get("standard_fee")
    if value in (None, ""):
        return None
    try:
        return int(float(str(value).replace(",", "")))
    except (TypeError, ValueError):
        return None


def pick_records(records: list, target_ym: str) -> tuple:
    """(対象年月ちょうどのレコード, 対象年月時点で有効なレコード) を返す。

    報酬月額は履歴テーブルで、基準年月から先に効く。**対象年月以前で最も新しい
    レコードが「いま効いている値」**。ここを対象年月ちょうどのレコードだけで見ると、
    標準報酬が動いていない月（料率変更のみ）にも新しい履歴を足してしまう。
    """
    dated = []
    for rec in records or []:
        ym = _rec_ym(rec)
        if ym:
            dated.append((ym, rec))
    dated.sort(key=lambda t: t[0])
    at_target = next((r for ym, r in dated if ym == target_ym), None)
    effective = None
    for ym, rec in dated:
        if ym <= target_ym:
            effective = rec
    return at_target, effective


def record_ym(rec) -> str:
    """レコードの基準年月 "YYYY-MM"（読めなければ空文字）。"""
    return _rec_ym(rec)


def _updater_label(rec) -> str:
    cls = ((rec or {}).get("last_update") or {}).get("classification") or {}
    name = str(cls.get("name") or "").strip()
    updater = ((rec or {}).get("last_update") or {}).get("updater") or {}
    who = str(updater.get("name") or updater.get("employee_id") or "").strip()
    return f"{name}／{who}" if name and who else (name or who)


# ---------------------------------------------------------------------------
# C: 当方の計算値
# ---------------------------------------------------------------------------

@dataclass
class CalcContext:
    master: object = None
    class_master: dict = field(default_factory=dict)
    months: dict = field(default_factory=dict)     # {ym: {社員番号: rec}}
    threshold: int = 17
    missing_months: list = field(default_factory=list)
    error: str = ""                                # マスタが読めない等（Cなしで続行）


@dataclass
class CalcResult:
    kenpo: int | None = None
    konen: int | None = None
    source: str = ""
    note: str = ""


def _fiscal_year(ym: str) -> int:
    """保険料の年度（4月始まり）。料率表の年度チェックに使う。"""
    return int(ym[:4]) if int(ym[5:7]) >= 4 else int(ym[:4]) - 1


def _teiji_year(ym: str) -> int:
    """その月に効いている定時決定の算定年（9月適用）。"""
    return int(ym[:4]) if int(ym[5:7]) >= 9 else int(ym[:4]) - 1


def revision_window(target_ym: str) -> list:
    """随時改定の算定3か月＝対象月の直前3か月。"""
    return [ym_add(target_ym, -3), ym_add(target_ym, -2), ym_add(target_ym, -1)]


def teiji_window(target_ym: str) -> list:
    y = _teiji_year(target_ym)
    return [f"{y}-04", f"{y}-05", f"{y}-06"]


def needed_windows(target_ym: str, stmt=None) -> list:
    """その紙を突き合わせるのに要る算定月だけを返す。

    stmt を渡すと**改訂理由から必要な窓だけ**に絞る。月額変更だけの月に
    定時決定の窓（前年4〜6月）まで読もうとすると、必要でもないキャッシュの
    欠落を警告してしまう。
    """
    reasons = [f"{p.reason_kenpo}{p.reason_konen}" for p in getattr(stmt, "persons", [])]
    if stmt is None:
        return list(dict.fromkeys(revision_window(target_ym) + teiji_window(target_ym)))
    months = []
    if any("月額変更" in r or "随時" in r for r in reasons):
        months += revision_window(target_ym)
    if any("定時" in r for r in reasons):
        months += teiji_window(target_ym)
    return list(dict.fromkeys(months))


def load_calc_context(target_ym: str, *, stmt=None, cache_dir: str = None,
                      grade_xlsx: str = None, class_csv: str = None,
                      insurer: str = None, threshold: int = None) -> CalcContext:
    """C（当方計算）に必要な等級表・分類マスタ・給与明細キャッシュを読む。

    読めないものがあっても例外にせず、`error` / `missing_months` に残して返す。
    C が出せない人は NO_CALC（差分確認で投入）に落ちるだけで、
    A（PDF）と B（jinjer）の突合は続けられるようにするため。
    """
    from services.shaho_engine import load_statements_full
    from services.shaho_master import ShahoMasterError, load_class_master, load_grade_table

    cache_dir = cache_dir or Config.KEIRI_OUTPUT_DIR
    ctx = CalcContext(threshold=threshold if threshold is not None
                      else Config.SHAHO_BASE_DAYS_THRESHOLD)
    try:
        ctx.master = load_grade_table(grade_xlsx or Config.SHAHO_GRADE_TABLE_XLSX,
                                      insurer or Config.SHAHO_INSURER,
                                      _fiscal_year(target_ym))
        ctx.class_master = load_class_master(class_csv or Config.SHAHO_CLASS_MASTER_CSV)
    except ShahoMasterError as e:
        ctx.error = str(e)
        return ctx
    except FileNotFoundError as e:
        ctx.error = f"マスタが見つかりません: {e}"
        return ctx

    for ym in needed_windows(target_ym, stmt):
        path = os.path.join(cache_dir, "raw", f"salary_statements_{ym}.json")
        if not os.path.exists(path):
            ctx.missing_months.append(ym)
            continue
        try:
            ctx.months[ym] = load_statements_full(cache_dir, ym)
        except Exception:  # キャッシュが壊れていてもCだけ諦めて先へ進む
            ctx.missing_months.append(ym)
    return ctx


def _compute_window(emp: str, window: list, ctx: CalcContext, *, need_all: bool):
    """算定3か月から標準報酬を計算する。(TeijiKettei, 説明) を返す。"""
    from services.shaho_engine import assess_month, calc_teiji_kettei, salary_system_name

    recs = [(m, (ctx.months.get(m) or {}).get(emp)) for m in window]
    missing = [m for m, r in recs if r is None]
    if missing:
        return None, "、".join(missing) + " の給与明細がない"
    if any((r.get("n_nonzero") or 0) > 1 for _m, r in recs):
        return None, "同じ月に給与明細が複数ある（2社勤務など）"

    assessments = [assess_month(m, r["payroll_info"], salary_system_name(r["basic_info"]),
                                ctx.class_master, ctx.threshold)
                   for m, r in recs]
    tk = calc_teiji_kettei(assessments, ctx.master)
    if tk.gate_ng:
        return None, "検算ゲート不一致の月がある（報酬計≠雇用保険対象額）"
    if tk.unclassified:
        return None, "分類できない支給項目に金額がある"
    if tk.adopted_n == 0 or (need_all and tk.adopted_n < len(window)):
        ng = "／".join(a.reason for a in tk.months if a.reason)
        return None, f"算定に使える月が足りない（{ng}）"
    return tk, ""


def expected_smr(person, target_ym: str, ctx: CalcContext) -> CalcResult:
    """PDF 1名分について、当方の計算値（C）を求める。出せないときは note に理由。"""
    if ctx is None or ctx.master is None:
        return CalcResult(note=(ctx.error if ctx else "") or "計算用のマスタが読めていない")

    reason = f"{person.reason_kenpo}{person.reason_konen}"
    if "取得" in reason:
        return CalcResult(note="資格取得時決定（見込み額）は給与実績からは計算できない")
    if "定時" in reason:
        window, source, need_all = teiji_window(target_ym), "定時決定", False
    elif "月額変更" in reason or "随時" in reason:
        window, source, need_all = revision_window(target_ym), "随時改定", True
    elif "料率" in reason:
        return CalcResult(note="料率変更のみ（標準報酬は改定されていない）")
    else:
        return CalcResult(note=f"改訂理由が判定できない（{person.reason or '空欄'}）")

    tk, why = _compute_window(person.emp, window, ctx, need_all=need_all)
    if tk is None:
        return CalcResult(source=source, note=f"{source}の計算ができない: {why}")
    return CalcResult(kenpo=tk.kenpo_smr, konen=tk.konen_smr, source=source,
                      note=f"{source}／{'・'.join(window)}の平均 {tk.average:,}円")


# ---------------------------------------------------------------------------
# 投入計画
# ---------------------------------------------------------------------------

def build_plan(stmt, current: dict, roster: dict, ctx: CalcContext = None,
               pdf_issues: dict = None) -> list:
    """PDF・jinjer現在値・当方計算値を突き合わせて投入計画を作る（純関数）。

    Args:
        stmt: `shaho_pdf.PdfStatement`
        current: `{社員番号: [報酬月額レコード, ...]}`（jinjer GET の結果）
        roster: `{社員番号: {"name","enrollment","retired_on",...}}`
        ctx: C（当方計算）の材料。None なら C なしで組む
        pdf_issues: `shaho_pdf.verify_person_premiums` の結果
    """
    pdf_issues = pdf_issues or {}
    rows = []
    for person in stmt.persons:
        row = PlanRow(emp=person.emp, name=person.name, reason=person.reason,
                      pdf_kenpo=person.kenpo_smr, pdf_konen=person.konen_smr)
        info = roster.get(person.emp) or {}
        row.jinjer_name = str(info.get("name") or "")

        at_target, effective = pick_records(current.get(person.emp) or [], stmt.target_ym)
        row.record_at_target = at_target is not None
        if effective is not None:
            row.cur_kenpo = _rec_fee(effective, "health_insurance")
            row.cur_konen = _rec_fee(effective, "employee_pension")
            row.cur_ym = _rec_ym(effective)
            row.cur_updated_by = _updater_label(effective)

        same = (row.cur_kenpo == row.pdf_kenpo and row.cur_konen == row.pdf_konen)

        if not person.emp:
            # 関東ITSのCSVのように、突合で社員番号が決まらないことがある。
            # 誰に書くか分からないものは絶対に書かない（承知のうえ投入も不可）。
            row.status = "UNRESOLVED"
            row.notes.extend(person.issues)
        elif classify_employee(person.emp) != "target":
            row.status = "EXCLUDED"
            row.notes.append("社員番号が 20YY 始まりではない（派遣・テスト番号）")
        elif not info:
            row.status = "NOT_IN_JINJER"
            row.notes.append("この社員番号が jinjer の従業員一覧にない")
        elif person.issues or pdf_issues.get(person.emp):
            row.status = "PDF_INCONSISTENT"
            row.notes.extend(person.issues)
            row.notes.extend(pdf_issues.get(person.emp, []))
        elif same:
            row.status = "NO_CHANGE"
            row.notes.append(f"{row.cur_ym} 登録の値と同じ")
        elif str(info.get("enrollment") or "") == "退職":
            row.status = "RETIRED"
            row.notes.append(f"jinjer上は退職（{info.get('retired_on') or '日付不明'}）")
        else:
            calc = expected_smr(person, stmt.target_ym, ctx)
            row.calc_kenpo, row.calc_konen = calc.kenpo, calc.konen
            row.calc_source = calc.source
            if calc.kenpo is None:
                row.status = "NO_CALC"
                row.notes.append(calc.note)
            elif calc.kenpo == row.pdf_kenpo and calc.konen == row.pdf_konen:
                row.status = "AUTO_OK"
                row.notes.append(calc.note)
            else:
                row.status = "CALC_MISMATCH"
                row.notes.append(
                    f"当方の計算は健保 {calc.kenpo:,}／厚年 {calc.konen:,} 円"
                    f"（{calc.note}）。社労士の値と違うので既定では投入しません")

        row.notes.extend(getattr(person, "warnings", []) or [])
        if (row.status != "UNRESOLVED" and row.jinjer_name and row.name
                and _name_key(row.jinjer_name) != _name_key(row.name)
                and not any("氏名" in n for n in row.notes)):
            row.notes.append(f"氏名がjinjerと違う（jinjer: {row.jinjer_name}）")
        rows.append(row)

    rows.sort(key=lambda r: (STATUS_ORDER.index(r.status)
                             if r.status in STATUS_ORDER else 99, r.emp))
    return rows


def _name_key(name: str) -> str:
    return str(name or "").replace(" ", "").replace("　", "").strip()


def plan_hash(rows: list, target_ym: str = "") -> str:
    """計画の同一性トークン。dry-run で見たものと違う計画を書かせないために使う。"""
    payload = [target_ym] + [
        f"{r.emp}|{r.pdf_kenpo}|{r.pdf_konen}|{r.cur_kenpo}|{r.cur_konen}"
        f"|{r.status}|{r.operation}" for r in rows]
    return hashlib.sha256("\n".join(payload).encode("utf-8")).hexdigest()[:16]


def summarize(rows: list) -> list:
    """ステータスごとの件数（画面のサマリカード用）。"""
    counts: dict[str, int] = {}
    for row in rows:
        counts[row.status] = counts.get(row.status, 0) + 1
    return [{"status": s, "label": STATUS_JA.get(s, s), "count": counts[s],
             "selectable": s in SELECTABLE, "review": s in FORCEABLE or s in
             ("UNRESOLVED", "PDF_INCONSISTENT", "NOT_IN_JINJER")}
            for s in STATUS_ORDER if counts.get(s)]


# ---------------------------------------------------------------------------
# 実行者の許可と同時実行ロック
# ---------------------------------------------------------------------------

# 許可リストの列。「承知投入」に印がある人だけが要確認の人を通せる。
FORCE_COLUMN = "承知投入"


def can_write(user: str = None) -> tuple:
    """(書き込み可否, 理由)。許可リストが読めなければ**書けない**側に倒れる。"""
    return _can_write_csv(Config.SHAHO_IMPORT_ALLOWED_USERS_CSV, user)


def can_force(user: str = None) -> tuple:
    """(要確認の人を「承知のうえ投入」できるか, 理由)。

    許可は2段構え（2026-08-17 谷津さん決定）:

    - **投入そのもの**は許可リストに載っていればできる（管理部6名）。
      3点（社労士PDF・jinjer登録値・当方計算）が一致した人を入れるだけなので、
      判断の余地がない。
    - **要確認の人（当方の計算と食い違う・退職済み）を承知のうえで投入する**のは
      判断が要るので、許可リストの「承知投入」列に印がある人だけに限る。

    列が無い・空欄の人は不可に倒れる（増やすときは明示的に印を付ける）。
    """
    from services.sap_import_ledger import _norm, current_user, load_writers

    who = user or current_user()
    allowed, why = can_write(who)
    if not allowed:
        return False, why

    rows = load_writers(Config.SHAHO_IMPORT_ALLOWED_USERS_CSV)
    for row in rows:
        if _norm(row.get("ユーザー名", "")).lower() == _norm(who).lower():
            if _norm(row.get(FORCE_COLUMN, "")):
                return True, f"要確認の人も承知のうえ投入できます（ユーザー: {who}）"
            break
    names = "・".join(_norm(r.get("表示名") or r.get("ユーザー名"))
                     for r in rows if _norm(r.get(FORCE_COLUMN, ""))) or "（誰もいません）"
    return False, (f"要確認の人の「承知のうえ投入」は {names} のみです。"
                   f"3点が一致した人の投入はできます（ユーザー: {who}）")


def acquire_lock(target_ym: str, count: int, *, path: str = None,
                 max_age_hours: int = 2) -> str:
    """同時実行ロックを取る。既に有効なロックがあれば ShahoWriteError。

    jinjer のレート制限は**テナント単位**なので、別PC・別モードの投入と
    並走させると両方が429で詰まる。取れなければ実行しない（フェイルクローズ）。
    """
    path = path or Config.SHAHO_IMPORT_LOCK_FILE
    now = datetime.datetime.now()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                held = json.load(f)
            started = datetime.datetime.fromisoformat(held.get("started_at"))
        except (OSError, ValueError, TypeError):
            held, started = {}, None
        if started and (now - started).total_seconds() < max_age_hours * 3600:
            raise ShahoWriteError(
                f"いま別の投入が動いています（{held.get('user', '不明')} が "
                f"{started:%H:%M} に開始・{held.get('count', '?')}件）。"
                "終わってから実行してください")
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"user": current_user(), "started_at": now.isoformat(),
                   "target_ym": target_ym, "count": count}, f, ensure_ascii=False)
    return path


def release_lock(path: str = None) -> None:
    path = path or Config.SHAHO_IMPORT_LOCK_FILE
    try:
        os.remove(path)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# バックアップ・実行・検証・台帳
# ---------------------------------------------------------------------------

def output_dir(target_ym: str, base: str = None) -> str:
    path = os.path.join(base or Config.SHAHO_IMPORT_OUTPUT_DIR, target_ym.replace("-", ""))
    os.makedirs(path, exist_ok=True)
    return path


def write_backup(target_ym: str, rows: list, current: dict, meta: dict,
                 base: str = None) -> str:
    """**1件も書く前に**現在値を丸ごと残す。戻せるのはこれがあるときだけ。

    新規登録（POST）だった人は元のレコードが無いので、PATCH では戻せない
    （jinjer の画面から履歴を消すことになる）。その旨をファイル内に明記する。
    """
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(output_dir(target_ym, base), f"投入前バックアップ_{stamp}.json")
    payload = {
        "説明": ("投入前の jinjer 報酬月額。操作=POST の人は対象年月のレコードが"
               "存在しなかったので、戻すときは PATCH ではなく jinjer 画面での"
               "履歴削除になります。"),
        "対象年月": target_ym,
        "取得日時": datetime.datetime.now().isoformat(timespec="seconds"),
        "メタ": meta,
        "対象": [{"社員番号": r.emp, "氏名": r.name, "操作": r.operation,
                "対象年月のレコード有無": r.record_at_target,
                "投入前_健保": r.cur_kenpo, "投入前_厚年": r.cur_konen,
                "投入前_基準年月": r.cur_ym,
                "投入予定_健保": r.pdf_kenpo, "投入予定_厚年": r.pdf_konen}
               for r in rows],
        "生レスポンス": {r.emp: current.get(r.emp, []) for r in rows},
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return path


def execute_plan(rows: list, client, target_ym: str, *, dry_run: bool = True,
                 progress=None, fresh: dict = None) -> list:
    """選択済みの行を1件ずつ投入する。

    Args:
        rows: 投入対象に絞った PlanRow のリスト
        client: `JinjerWriteClient`（dry_run なら None でよい）
        target_ym: 基準年月 "YYYY-MM"
        dry_run: True なら API を一切叩かず、やることだけ返す
        progress: `progress(done, total, entry)` で1件ごとに呼ばれる
        fresh: 直前に取り直した現在値。dry-run で見た時から動いた人を弾く

    Returns:
        1行1件の結果 dict のリスト。
    """
    year, month = target_ym.split("-")
    total = len(rows)
    results = []
    for i, row in enumerate(rows, start=1):
        entry = {"emp": row.emp, "name": row.name,
                 "before_kenpo": row.cur_kenpo, "before_konen": row.cur_konen,
                 "after_kenpo": row.pdf_kenpo, "after_konen": row.pdf_konen,
                 "operation": row.operation, "status": row.status,
                 "result": "", "message": ""}

        operation = row.operation
        if fresh is not None:
            at_target, effective = pick_records(fresh.get(row.emp) or [], target_ym)
            now_kenpo = _rec_fee(effective, "health_insurance") if effective else None
            now_konen = _rec_fee(effective, "employee_pension") if effective else None
            if (now_kenpo, now_konen) == (row.pdf_kenpo, row.pdf_konen):
                entry.update(result="スキップ", operation="スキップ",
                             message="すでに同じ値が入っています（投入済みとみなします）")
                results.append(entry)
                if progress:
                    progress(i, total, entry)
                continue
            if (now_kenpo, now_konen) != (row.cur_kenpo, row.cur_konen):
                entry.update(result="中止", operation="スキップ",
                             message=(f"確認したときと登録値が変わっています"
                                      f"（今: 健保 {now_kenpo}／厚年 {now_konen}）。"
                                      "もう一度プレビューし直してください"))
                results.append(entry)
                if progress:
                    progress(i, total, entry)
                continue
            operation = "PATCH" if at_target is not None else "POST"
            entry["operation"] = operation

        if dry_run:
            entry.update(result="dry-run", message="実行するとこの内容で書き込みます")
            results.append(entry)
            if progress:
                progress(i, total, entry)
            continue

        try:
            if operation == "PATCH":
                client.patch_monthly_remuneration(row.emp, year, month,
                                                  row.pdf_kenpo, row.pdf_konen)
            else:
                client.post_monthly_remuneration(row.emp, year, month,
                                                 row.pdf_kenpo, row.pdf_konen)
            entry["result"] = "OK"
        except Exception as e:                       # 1名の失敗で全体は止めない
            entry.update(result="失敗", message=str(e))
        results.append(entry)
        if progress:
            progress(i, total, entry)
    return results


def verify_after(client, rows: list, target_ym: str) -> dict:
    """書き込んだあと**取り直して**PDFの値と一致するか確かめる。

    書き込みAPIが200を返しても入っているとは限らない（jinjerには予約が
    サイレント破棄される前例がある）。独立に読み直すまで成功とは呼ばない。

    ⚠ ここだけは**対象年月で絞って**取る。絞らないと、書き込みが失敗して
    レコードができていない人でも「過去のレコードの値がたまたま同じ」だと
    OK に見えてしまう（標準報酬が動いていない人で実際に起こる）。
    """
    emps = [r.emp for r in rows]
    if not emps:
        return {}
    year, month = target_ym.split("-")
    fetched = client.get_monthly_remunerations(emps, year=year, month=month)
    out = {}
    for row in rows:
        _at, effective = pick_records(fetched.get(row.emp) or [], target_ym)
        got_kenpo = _rec_fee(effective, "health_insurance") if effective else None
        got_konen = _rec_fee(effective, "employee_pension") if effective else None
        if (got_kenpo, got_konen) == (row.pdf_kenpo, row.pdf_konen):
            out[row.emp] = "OK"
        else:
            out[row.emp] = (f"NG（健保 {got_kenpo}／厚年 {got_konen} が入っています。"
                            f"期待は {row.pdf_kenpo}／{row.pdf_konen}）")
    return out


def append_ledger(entries: list, *, path: str = None) -> str:
    """実行台帳へ追記する（新旧値つき）。ヘッダが無ければ作る。

    台帳に書けなくても投入の成否は変えない（書き込みは終わっている）。
    呼び出し側で例外を受けて退避先に書く。
    """
    path = path or Config.SHAHO_IMPORT_LEDGER_CSV
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    exists = os.path.exists(path) and os.path.getsize(path) > 0
    with open(path, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=LEDGER_COLUMNS, extrasaction="ignore")
        if not exists:
            writer.writeheader()
        for row in entries:
            writer.writerow({c: row.get(c, "") for c in LEDGER_COLUMNS})
    return path


def ledger_entries(results: list, verified: dict, *, target_ym: str, pdf_name: str,
                   backup: str, forced: set = None) -> list:
    """execute_plan の結果を台帳の行に変換する。"""
    forced = forced or set()
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    user = current_user()
    rows = []
    for r in results:
        rows.append({
            "実行日時": now, "実行者": user, "対象年月": target_ym,
            "PDFファイル名": pdf_name, "社員番号": r["emp"], "氏名": r["name"],
            "健保前": r.get("before_kenpo"), "健保後": r.get("after_kenpo"),
            "厚年前": r.get("before_konen"), "厚年後": r.get("after_konen"),
            "操作": r.get("operation", ""), "結果": r.get("result", ""),
            "検証": verified.get(r["emp"], ""),
            "承知投入": "○" if r["emp"] in forced else "",
            "バックアップ": os.path.basename(backup or ""),
        })
    return rows
