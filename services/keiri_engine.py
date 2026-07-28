r"""経理モード C-2/C-4: 給与明細 API → freee 取引インポート 4CSV の生成エンジン。

Z:\API連携\scripts\keiri_freee_engine.py から移植（2026-07-27）。本体はこちらが正で、
Z:\API連携 側には CLI ラッパと差分突合などの開発ツールだけを残す。

マッピングマスタ（既定 Z:\API連携\docs\経理モード_品目マッピングマスタ_draftC.csv）駆動。
本番は status=確定 の行のみ使用（min_status で開発時は 推定/要確認 まで緩和可能）。
レビュー結果はマスタ CSV の差し替えだけで反映される（exe の再ビルドは不要）。

入力 : jinjer の給与明細 API（月別 JSON をキャッシュ。当月と前月の2か月ぶん要る）
       カスタム項目（部門・人件費区分）も API 取得してキャッシュ
       Y:\給与明細\R8年\{M}月\経費利用履歴 RevN.xlsm（C-4 の経費転記の分解材料。無くても動く）
出力 : {出力先}\{YYYYMM}\freee_data_給与（…）.csv ほか4CSV（cp932）
       ＋ 検算_{YYYYMM}.md ／ 要確認_{YYYYMM}.md（utf-8）    ※すべて個人情報・非コミット

未対応（＝要確認_*.md の TODO 節に毎回出力）:
  未収入金・未払金（B8）／ミロク情報サービス取引／賞与（スコープ外）／
  慶弔見舞金・研修交通費の振替（経理判断）／住民税の会社立替行（税額通知書由来）／
  時点差（API は再計算後の値を返す）／過去月の育成区分・経費ブックが無い月の経費転記

2026-07 実績: 住民税・健保・厚年は完全一致、給与の残差は2件（jinjer 手入力の資格取得費）。
"""

from __future__ import annotations

import csv
import glob
import io
import json
import os
from collections import Counter, defaultdict
from decimal import ROUND_HALF_UP, Decimal

from config import Config
from services.keiri_api import (
    JINKENHI_KUBUN_DEFAULT,
    JINKENHI_KUBUN_NAMES,
    classify_employee,
    get_client,
    load_or_fetch_roster,
    midmonth_records,
    normalize_label,
    now_iso,
    parse_payroll_custom_history,
    resolve_custom_value,
    roster_index,
    statement_flag,
    to_number,
)
from services.keiri_keihi_tenki import decompose as keihi_decompose
from services.keiri_keihi_tenki import find_book as keihi_find_book
from services.keiri_keihi_tenki import load_details as keihi_load_details
from services.keiri_keihi_tenki import load_mapping as keihi_load_mapping

MASTER_CSV = Config.KEIRI_MASTER_CSV
KEIHI_MAPPING_CSV = Config.KEIRI_KEIHI_MAPPING_CSV
FINAL_CSV_DIR = Config.KEIRI_FINAL_CSV_DIR
OUT_BASE = Config.KEIRI_OUTPUT_DIR

JINKENHI = "人件費（{人件費区分}）"
STATUS_ORDER = {"確定": 0, "推定": 1, "要確認": 2}
FREEE_COLS = ["収支区分", "管理番号", "発生日", "支払期日", "取引先", "勘定科目", "税区分",
              "金額", "税計算区分", "税額", "備考", "品目", "部門", "従業員", "支払日",
              "支払口座", "支払金額"]

# 部門マスタ（Phase B 監査済みの 8 種。これ以外の値は要確認へ）
KNOWN_BUMON = {"本社", "OT：UAL（首都圏）", "OT：UAL（地方）", "OT：その他",
               "FS：UAL（受託）", "FS：UAL（常駐）", "SI：UAL（受託）", "SI：UAL（常駐）"}

# 会社側でまとめて計上する行の従業員名（実データの表記）
SONOTA_HONSHA = "その他本社経費"

# --- 特別ルール（2026-07-24 谷津さん回答＋実データ検証で確定） -------------------
# B3: 現物支給は購入時に仮払金計上 → 給与支給時に給料手当へ振替。
#     支給側は実績差分の人件費行へ合算（対象者が少数なら独立行）、
#     人件費区分=本社の人は常に独立行。仮払金のマイナス行は必ず1人1行。
GENBUTSU_KEY = "salary_items:allowance53"
GENBUTSU_BULK_THRESHOLD = 30   # これ以上なら支給側を人件費行へ合算
GENBUTSU_BIKO = "現物支給"      # 実際の文言（例「カタログギフト現物支給」）は都度手入力

# B2: 通勤手当は人件費区分=本社/育成の人だけ品目が人件費（…）になる。
#     ただし通勤費の品目は育成でも「人件費（本社）」に寄せる（2026-07-24 谷津さん確認。
#     稲田2026009・柳場2026010 の通勤費 16,940+29,740=46,680 が最終CSVで人件費（本社））。
TSUKIN_KEYS = ("salary_items:allowance35", "salary_items:allowance34")

# 立替金（経費の立替）は人件費区分=本社の人は計上しない（2026-07-24 谷津さん確認）。
#   ビジネス企画部・管理監督者は通勤費以外の経費を freee へ直接申請しているため重複になる
#   （ルール5「交通費（本社）を削除」と同じ理由）。
TATEKAE_KEYS = ("salary_items:allowance50", "salary_items:allowance51")

# 育成期間中は部門を「本社」に寄せる（2026-07-24 谷津さん確認）
IKUSEI_ITEM = "人件費（育成）"
# 入社から何か月以内なら「育成だったのでは？」と知らせるか
#   （実績: 2026-03入社の3名は4月まで育成・2026-01入社の池村2025022 も4月まで育成）
IKUSEI_WATCH_MONTHS = 6
HONSHA_ITEM = "人件費（本社）"
HONSHA_BUMON = "本社"

# B6/①: 三谷一志さんの役員貸付金は元利均等返済。元本＋利息の2行へ分解する。
#     実績7か月（2026-01〜07）に完全フィットするパラメータ。将来1年分も3種の丸めで一致を確認済。
#     jinjer 側の入力先は月によって揺れる（2026-02/06/07 は「貸付金返済」deduction3、
#     2026-05 は「家賃控除」deduction2、他の月は未入力）ので**両方を見る**。
#     ここに挙げた項目は master 側の行を使わず、この特別ルールで生成する（二重計上の防止）。
YAKUIN_LOAN = {
    "employee_id": "2023004",
    "source_keys": ["salary_deduction_items:deduction3",
                    "salary_deduction_items:deduction2"],
    "payment": 341659,          # 毎月の返済額（元本＋利息）
    "principal": 20_005_000,    # 当初借入
    "monthly_rate": 0.00081223,  # 月利（年利約0.975%）
    "anchor_ym": "2026-01", "anchor_count": 7, "total_count": 60,
}

# B7: 子ども・子育て拠出金は全員分を1行に一括計上（厚年ファイル）
KYOSHUTSUKIN_KEY = "salary_other_items:other19"
KYOSHUTSUKIN_BIKO = "子ども・子育て拠出金"

# B9/実測: 納付ファイルには従業員から預かった分を取り崩す行が品目ごとに1行入る。
#   金額は労使折半のため会社負担と同額＝**主取引に計上した人の会社負担額を品目別に集計**する
#   （2026-07 実測で主取引の合計と預り金の合計が完全一致 4,473,085）。
KENPO_AZUKARI = [("salary_other_items:other15", "健康保険料（預り分）"),
                 ("salary_other_items:other16", "介護保険料（預り分）"),
                 ("salary_deduction_items:child_support", "子ども・子育て支援金（預り分）")]
KONEN_AZUKARI = [("salary_other_items:other17", "厚生年金保険料（預り分）")]

# C3: 休職者かつ社保が発生している人は暫定支給とは別取引（管理番号・支払期日が空欄）
SHAHO_KEYS = ("salary_deduction_items:deduction29", "salary_deduction_items:deduction31")

# 社保調整(deduction4) は月変対応漏れ等の遡及調整。jinjer は1項目にまとめているが、
# 最終CSVでは 健保／介護／厚年／子ども子育て支援金 の品目別に分かれる。
# 本人の当月控除額で按分すると1円まで一致する（同じ標準報酬に率をかけた金額なので比が揃う）。
#   2026-07 実測: 小池2023019 の 16,680 = 厚年10,980 ＋ 健保5,562 ＋ 子育て138
SHAHO_CHOSEI_KEY = "salary_deduction_items:deduction4"
SHAHO_CHOSEI_BASIS = ("salary_deduction_items:deduction29",    # 健康保険料
                      "salary_deduction_items:deduction30",    # 介護保険料
                      "salary_deduction_items:deduction31",    # 厚生年金保険料
                      "salary_deduction_items:child_support")  # 子ども・子育て支援金
SHAHO_CHOSEI_BIKO = "社保調整"   # 実際の文言（例「社保調整分（4月、5月月変対応漏れのため）」）は手入力

# jinjer の社員番号が 20YY 始まりでないが計上対象にする人（classify_employee の例外）。
#   7777777 今井 保 … 監査役。jinjer に給与明細があり freee でも毎月 役員報酬 で計上されている。
SPECIAL_TARGET_EMPLOYEES = {"7777777"}

# jinjer 名簿と freee の従業員名の表記ゆれ（freee 側の表記に合わせる）
NAME_ALIASES = {"2014009": "ブアー マーティン"}   # jinjer は "Ver Martin"

# カスタム項目（部門・人件費区分）が jinjer に無い人の固定値
EMP_OVERRIDES = {"7777777": {"bumon": HONSHA_BUMON, "jinkenhi": HONSHA_ITEM}}

# ③ 給与ファイルの2取引は「当月(executed_on=支給月)の給与明細を2つに割ったもの」
#    （2026-07-24 実データで確定。奈良隆宏さん2022013で1円まで一致を確認）:
#      暫定取引（発生日=当月15日）   = 基本給＋みなし時間外＋みなし深夜＋調整＋リーダー＋役職 ＋ 全控除
#      実績差分取引（発生日=前月末） = 上記以外の支給項目（夜間当番・テレワーク・差額調整・
#                                    支給過不足調整・その他手当…）＋ 通勤費
#    → **前月の statements は使わない**（社保ファイルだけが前月分を使う）。
#    ⚠ allowance2 は体系で意味が変わるため id ではなく 体系別名(salary_system_label) で判定する:
#      月給制/140-180時間制 → 「当月みなし時間外手当」「みなし手当」「当月みなし深夜手当」= 暫定側
#      時給制               → 「前月超過勤務」= どちらにも計上しない（freee 実データでも未計上）
# id が体系によらず同じ意味の項目は id で判定する
#   （allowance15=調整手当 は月給制だと体系別名まで汎用名「給与支給項目15」になるため、
#     体系別名だけで判定すると暫定から漏れる。2026-07 実測: 小山2011012・岡野2024011）
ZANTEI_FIXED_KEYS = {
    "salary_items:allowance1",    # 基本給
    "salary_items:allowance10",   # 役職手当
    "salary_items:allowance11",   # リーダー手当
    "salary_items:allowance15",   # 調整手当
}
# allowance2 だけは体系で意味が変わるので体系別名で判定する
ZANTEI_LABELS = {"当月みなし時間外手当", "みなし手当", "当月みなし深夜手当"}
# 支給項目だが freee には計上しない（体系別名で判定）
#   前月超過勤務・基礎時給 … 時給制の情報項目（freee 実データでも未計上）
SKIP_LABELS = {"前月超過勤務", "前月基礎時給", "基礎時給"}

# 差額調整(allowance24) は前月精算分(allowance3/4/5/6)の再掲。同額のときは重複計上しない。
#   2026-06/07 実測の同額ペア: 前月実績分44件・前月控除精算分27件・超過精算分28件・休日労働精算分1件
#   （同額でない月・人は正当な別項目なので、リストではなく「同額なら除外」の動的判定にする）
#   allowance3 も構成要素に入る（2026-05 実測: 内田2025028 の差額調整134,064 =
#   allowance3 739 ＋ allowance5 124,983 ＋ allowance6 8,342。
#   体系別名が「給与支給項目3」の汎用名のままなので 前月〜 のスキップ判定に掛からない）
SAGAKU_KEY = "salary_items:allowance24"
SEISAN_KEYS = ("salary_items:allowance3", "salary_items:allowance4",
               "salary_items:allowance5", "salary_items:allowance6")

# allowance52「その他」= 経費統合一覧表からの追加明細の合計。人件費には混ぜず、
#   消耗品費/社員育成費/会議接待費/通信費/福利厚生費/雑費 に分けて計上する（C-4）。
#   例: 岡野2024011 の 62,580 = 会議接待費13,080 + 社員育成費49,500。
#   分解は keiri_keihi_tenki.py が担当し、合計が合わない人は行を作らず要確認へ回す。
KEIHI_TENKI_KEY = "salary_items:allowance52"

# 対象外にしたが値が出たら知らせてほしい項目（マスタの判断が正しいかを毎月見張る）
WATCH_KEYS = {
    "salary_items:allowance8":
        "過不足調整。2026-03/04 は暫定支給額の再掲で最終CSVは13名全員未計上のため対象外にした",
}

# 役員（谷津さん確定）: 給与の勘定科目が 給料手当 ではなく 役員報酬 になる。
# 賞与・現物支給時は 役員賞与（3月カタログギフトの実績で確認）。
YAKUIN_EMPLOYEES = {"2008003", "2023004", "7777777"}
YAKUIN_ACCOUNT = "役員報酬"
YAKUIN_BONUS_ACCOUNT = "役員賞与"

SKELETON_TODOS = [
    "未収入金・未払金（B8: 急な退職で社保を徴収しきれない等のイレギュラー）: "
    "システム化対象外。下の検知リストに出すので手で追加すること",
    "社保調整(deduction4)の備考: 都度内容が違うため手入力（品目の分割は実装済み）",
    "現物支給の備考: 実際の文言（例「カタログギフト現物支給」）は手入力",
    "ミロク情報サービス取引: 未実装（出どころ要確認）",
    "賞与ファイル: スコープ外（ルール7）",
    "慶弔見舞金・研修交通費の振替: 経理の判断でしか決まらない"
    "（実績: 立替金→福利厚生費、非課税通勤費→社員育成費）。jinjer 側からは判別できない",
    "住民税の会社立替行: 特別徴収の開始月に会社が立て替える分は税額通知書由来で"
    "jinjer に無いため生成できない。当月の相殺は前月の最終CSVから読む",
    "時点差: API は現在の再計算値を返すため、支払時点で凍結された最終CSVとは差が出る"
    "（C1回答: 現時点の値で進める）。2026-04 は家賃控除・役員報酬が jinjer 側にまだ無い",
    "過去月の再現限界: 人件費区分（育成）は 2026-07 の一括投入以降しか jinjer に無いため、"
    "それ以前の月は復元できない。経費一覧表マクロのブックが無い月は経費転記も分解できない",
]


# ---------------------------------------------------------------------------
# 日付ユーティリティ
# ---------------------------------------------------------------------------
def ym_add(ym, delta):
    y, m = int(ym[:4]), int(ym[5:7])
    total = y * 12 + (m - 1) + delta
    return f"{total // 12:04d}-{total % 12 + 1:02d}"


def month_last_day(ym):
    import calendar
    y, m = int(ym[:4]), int(ym[5:7])
    return f"{y:04d}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}"


def jp_date(iso):
    """'2026-07-15' → '2026/7/15'（freee CSV の 0 パディングなし表記）。"""
    if not iso:
        return ""
    y, m, d = iso.split("-")
    return f"{int(y)}/{int(m)}/{int(d)}"


def ym_compact(ym):
    return ym.replace("-", "")


# ---------------------------------------------------------------------------
# マスタ読み込み
# ---------------------------------------------------------------------------
def load_master(path, min_status):
    """マッピングマスタを読み、行を用途別にグループして返す。

    戻り値 dict:
      by_key      : {source_key: [全行]}（対象外含む全行。新規使用アラート用）
      z_jinkenhi  : 暫定側・人件費集約の source_key リスト（給与）
      z_deduction : 暫定側・控除等の行リスト（給与、sign=-1）
      j_items     : 実績差分側の行リスト（給与）
      juminzei / kenpo / konen : 各納付ファイルの行リスト
      skipped     : min_status で除外された行数（status別）
      skipped_rows: 同・除外された行そのもの（検算シートで項目名まで出すため）
      total_n     : マッピング表の全行数
      offscope    : 仕訳に使わない行数（target_file別＝対象外・賞与スコープ外・要確認）
    """
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    limit = STATUS_ORDER[min_status]
    by_key = defaultdict(list)
    active = []
    skipped = Counter()
    skipped_rows = []
    offscope = Counter()
    for r in rows:
        by_key[r["source_key"]].append(r)
        if r["target_file"] in ("対象外", "賞与スコープ外", "要確認"):
            offscope[r["target_file"]] += 1
            continue
        if STATUS_ORDER.get(r["status"], 9) <= limit:
            active.append(r)
        else:
            skipped[r["status"]] += 1
            skipped_rows.append(r)
    g = {
        "by_key": by_key,
        "z_jinkenhi": [], "z_deduction": [], "j_items": [],
        "juminzei": [], "kenpo": [], "konen": [],
        "skipped": skipped, "active_n": len(active),
        "skipped_rows": skipped_rows, "total_n": len(rows),
        "offscope": offscope, "min_status": min_status, "path": path,
    }
    for r in active:
        if r["target_file"] == "給与":
            if r["transaction_side"] == "暫定":
                if r["freee_item"] == JINKENHI:
                    g["z_jinkenhi"].append(r)
                else:
                    g["z_deduction"].append(r)
            else:
                g["j_items"].append(r)
        elif r["target_file"] == "住民税":
            g["juminzei"].append(r)
        elif r["target_file"] == "健康保険":
            g["kenpo"].append(r)
        elif r["target_file"] == "厚生年金":
            g["konen"].append(r)
    g["z_by_key"] = {r["source_key"]: r for r in g["z_deduction"]}
    return g


# ---------------------------------------------------------------------------
# statements キャッシュ
# ---------------------------------------------------------------------------
def fetch_statements(cache_dir, ym, client=None, refresh=False):
    """指定月の給与明細を API 取得して JSON キャッシュに保存し、生 data[] を返す。

    1か月ぶんで 200 人分・数十ページになるので、既にキャッシュがあれば使い回す
    （締め後の月は値が動かない。再取得したいときだけ refresh=True）。
    """
    path = os.path.join(cache_dir, "raw", f"salary_statements_{ym}.json")
    if not refresh and os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)["data"]
    data = (client or get_client()).get_salary_statements(ym)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"fetched_at": now_iso(), "executed_on": ym, "data": data},
                  f, ensure_ascii=False)
    return data


def load_statements(cache_dir, ym, client=None, refresh=False):
    """{社員番号(20YYのみ): payroll_info} を返す（複数statementsは基本給非ゼロを採用）。"""
    data = fetch_statements(cache_dir, ym, client=client, refresh=refresh)
    out = {}
    for person in data:
        emp = str(person.get("employee_id", "")).strip()
        if classify_employee(emp) != "target" and emp not in SPECIAL_TARGET_EMPLOYEES:
            continue
        statements = person.get("statements") or []
        if not statements:
            continue
        best = statements[0]
        for st in statements:
            pi = st.get("payroll_info", {}) or {}
            for it in pi.get("salary_items", []) or []:
                if str(it.get("id")) == "allowance1" and to_number(it.get("value")):
                    best = st
                    break
        out[emp] = best.get("payroll_info", {}) or {}
    return out


def pi_item(pi, source_key):
    """payroll_info から source_key（'salary_items:allowance1'）の項目 dict を取る。"""
    atype, item_id = source_key.split(":", 1)
    for it in pi.get(atype, []) or []:
        if str(it.get("id")) == item_id:
            return it
    return None


def pi_value(pi, source_key):
    """payroll_info から source_key の数値を取る。"""
    it = pi_item(pi, source_key)
    if it is None:
        return 0.0
    n = to_number(it.get("value"))
    return n if n is not None else 0.0


def item_label(item):
    return (str(item.get("salary_system_label") or "").strip()
            or str(item.get("label") or "").strip())


def is_zantei_item(item, source_key):
    """暫定支給額（基本給＋みなし時間外＋みなし深夜＋調整＋リーダー＋役職）に含める項目か。"""
    if source_key in ZANTEI_FIXED_KEYS:
        return True
    return item_label(item) in ZANTEI_LABELS


def is_skip_item(item):
    """freee に計上しない支給項目。

    体系別名が「前月〜」の項目（前月超過勤務・前月深夜残業分・前月休日出勤時間分・
    前月実績分・前月基礎時給…）は差額調整に含まれており、最終CSVには計上されない
    （2026-07 実測: 畑中2011004 の差 12,299 = 前月深夜残業分10,339＋前月休日出勤1,960、
      ほか時給制8名の残差がすべてこれで説明できた）。
    """
    label = item_label(item)
    return label in SKIP_LABELS or label.startswith("前月")


def sagaku_dupe_keys(pi):
    """差額調整の再掲になっている精算項目のキー集合を返す。

    実測パターン: 単独一致（allowance5 == allowance24）と
    合計一致（allowance5 + allowance6 == allowance24。河端2024045 で 53,907+22,275=76,182）。
    """
    sagaku = pi_value(pi, SAGAKU_KEY)
    if not sagaku:
        return set()
    nonzero = {k: pi_value(pi, k) for k in SEISAN_KEYS if pi_value(pi, k)}
    if not nonzero:
        return set()
    if abs(sum(nonzero.values()) - sagaku) < 0.5:
        return set(nonzero)                       # 精算分の合計＝差額調整
    return {k for k, v in nonzero.items() if abs(v - sagaku) < 0.5}


def calc_zantei(pi, master, detail=None):
    """暫定支給額（基本給＋みなし時間外＋みなし深夜＋調整＋リーダー＋役職）を返す。"""
    total = 0.0
    for r in master["z_jinkenhi"]:
        it = pi_item(pi, r["source_key"])
        if it is None or not is_zantei_item(it, r["source_key"]):
            continue
        v = to_number(it.get("value")) or 0.0
        total += v
        if detail is not None and v:
            detail.append((r["source_key"], str(it.get("salary_system_label") or it.get("label")), v))
    return total


def split_shaho_chosei(amount, pi):
    """社保調整を本人の当月控除額（健保・介護・厚年・子育て支援金）で按分する。

    戻り値は [(source_key, 符号つき金額), ...]。分割できないときは空リスト＝マスタどおり1行。

    **端数が1円でも出るときは分割しない。** 社保調整の中身は「率×標準報酬の差額」なので、
    本人の控除額（同じ率×標準報酬）で按分すると必ず割り切れる。割り切れない調整は
    月変差額ではない（実測: 2026-05 村山2013020 の 25円は最終CSVで雇用保険料だった）ため、
    機械的に分けずマスタどおり1行にして要確認へ回す。
      2026-07 小池2023019: 16,680 → 厚年10,980／健保5,562／子育て138 がいずれも割り切れる
    """
    basis = [(k, v) for k, v in ((k, pi_value(pi, k)) for k in SHAHO_CHOSEI_BASIS) if v]
    total_basis = sum(v for _k, v in basis)
    if not total_basis:
        return []
    sign = -1 if amount < 0 else 1
    amt = int(round(abs(amount)))
    parts = []
    for k, v in basis:
        exact = amt * v / total_basis
        if abs(exact - round(exact)) > 1e-6:
            return []               # 端数が出る＝月変差額ではないので分割しない
        parts.append((k, sign * int(round(exact))))
    return [(k, v) for k, v in parts if v]


def split_halves(amount):
    """ルール9(d): ÷2して端数は前半（旧部門）側に寄せる。両側の合計＝元額。"""
    sign = -1 if amount < 0 else 1
    total = abs(amount)
    new_half = int(total // 2)
    return sign * (total - new_half), sign * new_half     # (旧側, 新側)


def midmonth_date(histories, emp, ym):
    """ym 内に月中（2日〜末日）の履歴があればその日付を返す（ルール9の分割スイッチ）。"""
    recs = midmonth_records(histories.get(emp, []), ym)
    return recs[-1]["date"] if recs else None


def jinkenhi_account(emp, is_bonus=False):
    """役員は 役員報酬（賞与・現物支給なら 役員賞与）、それ以外は 給料手当。"""
    if emp in YAKUIN_EMPLOYEES:
        return YAKUIN_BONUS_ACCOUNT if is_bonus else YAKUIN_ACCOUNT
    return "給料手当"


# ---------------------------------------------------------------------------
# カスタム項目（部門・人件費区分）
# ---------------------------------------------------------------------------
def load_custom_histories(out_base, roster_ids, refresh=False):
    cache = os.path.join(out_base, "raw", "custom_items.json")
    if not refresh and os.path.exists(cache):
        with open(cache, "r", encoding="utf-8") as f:
            raw = json.load(f)["data"]
    else:
        client = get_client()
        raw = client.get_custom_items(sorted(roster_ids))
        os.makedirs(os.path.dirname(cache), exist_ok=True)
        with open(cache, "w", encoding="utf-8") as f:
            json.dump({"fetched_at": now_iso(), "data": raw}, f, ensure_ascii=False)
    return {emp: parse_payroll_custom_history(person) for emp, person in raw.items()}


class Resolver:
    """部門・人件費区分の時点解決（ルール2/3）＋未解決アラート収集。"""

    def __init__(self, histories, alerts, ridx=None):
        self.histories = histories
        self.alerts = alerts
        self.ridx = ridx

    def bumon(self, emp, on_date, context):
        if emp in EMP_OVERRIDES:
            return EMP_OVERRIDES[emp]["bumon"]
        # 育成期間中は部門を「本社」に寄せる（谷津さん確認 2026-07-24）
        if self.jinkenhi_item(emp, on_date) == IKUSEI_ITEM:
            return HONSHA_BUMON
        hist = self.histories.get(emp, [])
        v = resolve_custom_value(hist, "部門", on_date, fallback_earliest=False)
        if not v:
            # 対象日より前に履歴が無い（入社直後・後から登録した等）→ 最古の値で代用し記録する
            v = resolve_custom_value(hist, "部門", on_date, fallback_earliest=True)
            self.alerts["bumon_missing"].add((emp, context, v or "(解決不能)"))
        if v and v not in KNOWN_BUMON:
            self.alerts["bumon_unknown"].add((emp, v))
        return v

    def jinkenhi_item(self, emp, on_date):
        if emp in EMP_OVERRIDES:
            return EMP_OVERRIDES[emp]["jinkenhi"]
        # 人件費区分は 2026-07-01 付けで一括投入したため、それ以前の月には
        # 最古レコードを遡及適用する（fallback_earliest=True）
        v = resolve_custom_value(self.histories.get(emp, []), "人件費区分", on_date)
        name = JINKENHI_KUBUN_NAMES.get(v, "")
        if not name:
            # jinjer は 本社/育成 の例外だけを持ち、空欄＝本社以外（既定）。
            # ただし育成レコードは 2026-07 の一括投入以降しか無いので、
            # **入社まもない人が既定に落ちた場合だけ**「本当は育成では？」と知らせる。
            name = JINKENHI_KUBUN_DEFAULT
            joined = str((self.ridx or {}).get(emp, {}).get("joined_on") or "")
            if joined and ym_add(on_date[:7], -IKUSEI_WATCH_MONTHS) <= joined[:7] <= on_date[:7]:
                self.alerts["ikusei_maybe"].add((emp, joined, on_date[:7]))
        return f"人件費（{name}）"


# ---------------------------------------------------------------------------
# 行・取引の構築
# ---------------------------------------------------------------------------
def detail_row(account, tax, amount, item, bumon, name, biko=""):
    return {"勘定科目": account, "税区分": tax, "金額": int(round(amount)),
            "税計算区分": "内税", "税額": "", "備考": biko, "品目": item, "部門": bumon,
            "従業員": name, "支払日": "", "支払口座": "", "支払金額": ""}  # ルール8: 右4列は常に空欄


def resolve_item_name(row, resolver, emp, on_date):
    item = row["freee_item"]
    if item == JINKENHI:
        return resolver.jinkenhi_item(emp, on_date)
    return item


def loan_split(month):
    """役員貸付金の当月分を (回数, 元本, 利息) で返す（元利均等・実績にフィット済）。"""
    cfg = YAKUIN_LOAN
    n = cfg["anchor_count"] + (
        (int(month[:4]) * 12 + int(month[5:7])) -
        (int(cfg["anchor_ym"][:4]) * 12 + int(cfg["anchor_ym"][5:7])))
    bal = cfg["principal"]
    principal = interest = 0
    for _i in range(n):
        interest = int(Decimal(str(bal * cfg["monthly_rate"])).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP))
        principal = cfg["payment"] - interest
        bal -= principal
    return n, principal, interest


def is_shaho_menjo(pi):
    """社会保険料が免除されている人か（育児休業・産前産後休業）。

    jinjer の basic_info.parental_leave_classification は当テナントでは全員「未選択」で
    使えないため、**本人負担がゼロなのに会社負担だけ計上されている**形で判定する
    （労使折半なので本人ゼロ＝会社もゼロが正しい。2026-07-24 谷津さん確認:
      井料2017037・有田2018001 は育休中で社保は発生しない）。
    """
    honnin = sum(pi_value(pi, k) for k in SHAHO_KEYS)
    kaisha = sum(pi_value(pi, k) for k in ("salary_other_items:other15",
                                           "salary_other_items:other17"))
    return kaisha > 0 and honnin == 0


def is_kyushoku(pi):
    """休職者かつ社保発生（C3: 支払期日が空欄の別取引になる人）の判定。"""
    if pi is None:
        return False
    net = pi_value(pi, "salary_payment_items:payment1")
    shaho = any(pi_value(pi, k) for k in SHAHO_KEYS)
    return net <= 0 and shaho


def build_kyuyo(month, prev, st_m, st_prev, ridx, resolver, master, paid_on, alerts,
                split_midmonth=True, keihi_details=None, keihi_mapping=None):
    """給与ファイル: 当月の給与明細を「実績差分（発生日=前月末）」と「暫定支給（当月15日）」に割る。

    ③で確定した構造。st_prev は使わない（社保ファイルのみ前月分を使う）。
    """
    m_first = f"{month}-01"
    month_end_m = month_last_day(month)
    prev_end = month_last_day(prev)
    zantei_rows, jisseki_rows, karibarai_rows = [], [], []
    kyushoku = {}
    keihi_details = keihi_details or {}

    # 現物支給: 対象者が閾値以上なら支給側を人件費行へ合算（B3）
    genbutsu_emps = {e for e, pi in st_m.items() if pi_value(pi, GENBUTSU_KEY)}
    genbutsu_bulk = len(genbutsu_emps) >= GENBUTSU_BULK_THRESHOLD
    if genbutsu_emps:
        alerts["genbutsu"].add((month, len(genbutsu_emps), genbutsu_bulk))

    for emp in sorted(st_m):
        name = ridx.get(emp, {}).get("name", emp)
        pi = st_m[emp]
        dupes = sagaku_dupe_keys(pi)
        # --- 暫定支給（当月15日）: 基本給＋みなし系＋手当1 と 全控除 ---
        # ルール2の「当月1日以降の部門」＝当月内の最新レコードを採用する
        # （月中異動者は異動後の部門になる。稲田2026009・柳場2026010 の 7/16 異動で確認）
        bumon_m = resolver.bumon(emp, month_end_m, f"{month}暫定")
        rows = []
        base = calc_zantei(pi, master)
        if base:
            rows.append(detail_row(jinkenhi_account(emp), "対象外", base,
                                   resolver.jinkenhi_item(emp, month_end_m), bumon_m, name))
        # ①三谷さんの役員貸付金は返済スケジュールから毎月生成する。
        #   jinjer 側は入力先（貸付金返済/家賃控除）も入力の有無も月によって揺れるため
        #   API 非依存にし、値が入っている月だけ返済額と突き合わせて見張る。
        if emp == YAKUIN_LOAN["employee_id"]:
            n, principal, interest = loan_split(month)
            if 1 <= n <= YAKUIN_LOAN["total_count"]:
                api_v = sum(pi_value(pi, k) for k in YAKUIN_LOAN["source_keys"])
                if api_v and abs(principal + interest - api_v) >= 1:
                    alerts["loan_mismatch"].add((month, api_v, principal + interest))
                rows.append(detail_row("役員貸付金", "対象外", -principal, "役員貸付金",
                                       bumon_m, name, f"役員貸付金（{n}/{YAKUIN_LOAN['total_count']}回）"))
                rows.append(detail_row("受取利息", "非課売上", -interest, "雑収入",
                                       bumon_m, name, "役員貸付金利息"))
        for r in master["z_deduction"]:
            v = pi_value(pi, r["source_key"])
            if not v:
                continue
            if (emp == YAKUIN_LOAN["employee_id"]
                    and r["source_key"] in YAKUIN_LOAN["source_keys"]):
                continue   # 上で生成済み
            if r["source_key"] == SHAHO_CHOSEI_KEY:
                # 社保調整は品目別（健保・介護・厚年・子育て支援金）に分けて計上する
                alerts["biko_needed"].add((emp, name, "社保調整", int(v)))
                parts = split_shaho_chosei(int(r["amount_sign"]) * v, pi)
                if parts:
                    for key, amt in parts:
                        mr = master["z_by_key"].get(key, r)
                        rows.append(detail_row(
                            mr["freee_account"], mr["freee_tax"], amt,
                            resolve_item_name(mr, resolver, emp, m_first), bumon_m, name,
                            SHAHO_CHOSEI_BIKO))
                    alerts["shaho_chosei_split"].add(
                        (emp, name, int(v), tuple((master["z_by_key"].get(k, r)["freee_item"], a)
                                                  for k, a in parts)))
                    continue
            rows.append(detail_row(
                r["freee_account"], r["freee_tax"], int(r["amount_sign"]) * v,
                resolve_item_name(r, resolver, emp, m_first), bumon_m, name))
        # 差引支給額がマイナス＝給与から引ききれない分は仮払金として計上する
        #   （2026-07 実測: 村岡2023017 の -117,300 が最終CSVで仮払金 117,300）
        net = pi_value(pi, "salary_payment_items:payment1")
        if net < 0:
            rows.append(detail_row("仮払金", "対象外", -net, "仮払金", bumon_m, name))
            alerts["karibarai"].add((emp, name, int(-net)))
        if is_kyushoku(pi):
            kyushoku[emp] = rows          # C3: 休職者は1名=1取引（管理番号・支払期日は空欄）
            alerts["kyushoku"].add((emp, name))
        else:
            zantei_rows.extend(rows)

        # --- 実績差分（発生日=前月末）: 暫定に入らない支給項目（部門は前月時点＝ルール2） ---
        bumon_p = resolver.bumon(emp, prev_end, f"{prev}実績差分")
        kubun_item = resolver.jinkenhi_item(emp, prev_end)
        is_honsha = kubun_item != "人件費（本社以外）"
        jinkenhi_diff = 0.0
        others = []
        for r in master["j_items"]:
            it = pi_item(pi, r["source_key"])
            if it is None:
                continue
            v = to_number(it.get("value")) or 0.0
            if not v or is_zantei_item(it, r["source_key"]) or is_skip_item(it):
                continue
            key = r["source_key"]
            if key in dupes:
                continue                       # 差額調整の再掲になっている精算分は飛ばす
            if key == KEIHI_TENKI_KEY:
                # C-4: 経費一覧表マクロの明細（集計ログの J:その他）から品目別に分解する。
                # 明細の合計が allowance52 と合わないときは行を作らず要確認へ回す
                # （前月分の合算・マクロ側の重複・経費申請でない手入力などがあるため）
                det = keihi_details.get(emp, [])
                det_total = sum(a for _u, a, _m in det)
                if det and abs(det_total - v) < 0.5:
                    by_item, biko, reasons, _t = keihi_decompose(det, keihi_mapping)
                    for (item, account, tax), amt in sorted(by_item.items()):
                        others.append(detail_row(account, tax, amt, item, bumon_p, name,
                                                 biko.get((item, account, tax), "")))
                    for d in reasons:
                        # 同額・同内訳の明細が複数あるので set にしない（合計が合わなくなる）
                        alerts["keihi_bunkai"].append((emp, name, d["内訳"], int(d["金額"]),
                                                       d["品目"], d["根拠"], d["status"]))
                else:
                    alerts["keihi_tenki"].add((emp, name, int(v), bumon_p,
                                               int(det_total) if det else None))
                continue
            # B3 現物支給: 仮払金の取り崩し行は必ず1人1行（部門=本社・その他本社経費）
            if key == GENBUTSU_KEY:
                karibarai_rows.append(detail_row("仮払金", "対象外", -v, "仮払金",
                                                 "本社", SONOTA_HONSHA, GENBUTSU_BIKO))
                if genbutsu_bulk and not is_honsha:
                    jinkenhi_diff += v       # 支給側は人件費行へ合算
                else:
                    others.append(detail_row(jinkenhi_account(emp, is_bonus=True), "対象外",
                                             v, kubun_item, bumon_p, name, GENBUTSU_BIKO))
                continue
            # B2 通勤手当: 本社・育成の人は品目が人件費（…）になる（育成も本社に寄せる）
            if key in TSUKIN_KEYS and is_honsha:
                others.append(detail_row(r["freee_account"], r["freee_tax"],
                                         int(r["amount_sign"]) * v, HONSHA_ITEM, bumon_p, name))
                continue
            # 立替金: 本社（ビジネス企画部・管理監督者）は freee へ直接申請済みのため計上しない
            if key in TATEKAE_KEYS and kubun_item == HONSHA_ITEM:
                alerts["tatekae_skip"].add((emp, name, int(v)))
                continue
            if r["freee_item"] == JINKENHI:
                jinkenhi_diff += int(r["amount_sign"]) * v
            else:
                others.append(detail_row(
                    r["freee_account"], r["freee_tax"], int(r["amount_sign"]) * v,
                    resolve_item_name(r, resolver, emp, prev_end), bumon_p, name))
        if round(jinkenhi_diff):
            # ルール9: 前月に月中異動があれば給料手当を÷2して旧部門・新部門の2行にする。
            # 当月払いのため、7月の異動は8月CSVの実績差分（発生日7/31＝7月分の費用）で分割する
            # （谷津さん確認 2026-07-24。7月CSVの暫定が分割されていないのはこのため）。
            idou = midmonth_date(resolver.histories, emp, prev) if split_midmonth else None
            if idou:
                old_amt, new_amt = split_halves(jinkenhi_diff)
                first = f"{prev}-01"
                jisseki_rows.append(detail_row(
                    jinkenhi_account(emp), "対象外", old_amt,
                    resolver.jinkenhi_item(emp, first), resolver.bumon(emp, first, f"{prev}異動前"),
                    name, f"{idou} 異動（異動前分）"))
                jisseki_rows.append(detail_row(
                    jinkenhi_account(emp), "対象外", new_amt, kubun_item, bumon_p,
                    name, f"{idou} 異動（異動後分）"))
                alerts["split_done"].add((emp, name, idou, int(old_amt), int(new_amt)))
            else:
                jisseki_rows.append(detail_row(jinkenhi_account(emp), "対象外", jinkenhi_diff,
                                               kubun_item, bumon_p, name))
        jisseki_rows.extend(others)  # 最終CSVの慣例: 人件費行が従業員の先頭

    kanri = min(st_m or st_prev)  # 最終CSV慣例: 管理番号=先頭従業員コード（取引単位）
    transactions = []
    # 現物支給の仮払金取り崩しは支給側（実績差分）と同じ取引に入れる
    if jisseki_rows or karibarai_rows:
        transactions.append({"管理番号": kanri, "発生日": jp_date(prev_end),
                             "支払期日": jp_date(paid_on), "取引先": "従業員",
                             "rows": jisseki_rows + karibarai_rows})
    if zantei_rows:
        transactions.append({"管理番号": kanri, "発生日": jp_date(f"{month}-15"),
                             "支払期日": jp_date(paid_on), "取引先": "従業員",
                             "rows": zantei_rows})
    for emp, rows in kyushoku.items():   # C3: 管理番号・支払期日とも空欄
        transactions.append({"管理番号": "", "発生日": jp_date(f"{month}-15"),
                             "支払期日": "", "取引先": "従業員", "rows": rows})
    return transactions


def load_prev_juminzei(prev, out_base, final_dir=None):
    """前月の住民税CSVから従業員別の金額を読む（会社が立て替えて納付した分）。

    異動届が前月給与に間に合わなかった人は、前月に会社が立て替えて計上し、
    当月に2か月分を従業員から徴収して相殺する（谷津さん確認 2026-07-24）。
    当月のCSVには当月分だけを載せるため、前月の計上額を差し引く。

    立替額は特別徴収税額決定通知書由来で jinjer には無いため、**経理担当が手で足した行**が
    唯一の情報源になる。よって前月の最終CSV（実際に freee へ取り込んだもの）を優先して読み、
    それが無いときだけ自前の生成結果で代用する。
    """
    cands = glob.glob(os.path.join(final_dir or FINAL_CSV_DIR, f"{int(prev[5:7])}月",
                                   "freee", "*住民税*.csv"))
    if not cands:
        cands = glob.glob(os.path.join(out_base, ym_compact(prev), "*住民税*.csv"))
    if not cands:
        return {}
    raw = open(cands[0], "rb").read()
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            rows = list(csv.reader(io.StringIO(raw.decode(enc))))
            break
        except UnicodeDecodeError:
            continue
    else:
        return {}
    idx = {c: i for i, c in enumerate(rows[0]) if c}
    out = {}
    for r in rows[1:]:
        def v(col, r=r):
            i = idx.get(col)
            return r[i].strip() if i is not None and i < len(r) else ""
        amt = to_number(v("金額"))
        if amt and v("従業員"):
            out[normalize_label(v("従業員"))] = out.get(normalize_label(v("従業員")), 0.0) + amt
    return out


def build_juminzei(month, st_m, ridx, resolver, master, prev_paid=None, alerts=None):
    rows = []
    m_first = f"{month}-01"
    prev_paid = prev_paid or {}
    for emp in sorted(st_m):
        pi = st_m[emp]
        name = ridx.get(emp, {}).get("name", emp)
        for r in master["juminzei"]:
            v = pi_value(pi, r["source_key"])
            if not v:
                continue
            # 前月に会社が立て替えた分は前月CSVで計上済みなので当月から差し引く
            tatekae = prev_paid.get(normalize_label(name), 0.0)
            if tatekae and alerts is not None and (emp, name) in {
                    (e, n) for e, n, _a in alerts.get("juminzei_shokai", set())}:
                v -= tatekae
                alerts["juminzei_soosai"].add((emp, name, int(tatekae), int(v)))
            if v:
                rows.append(detail_row(r["freee_account"], r["freee_tax"],
                                       int(r["amount_sign"]) * v, r["freee_item"],
                                       resolver.bumon(emp, m_first, f"{month}住民税"), name))
    if not rows:
        return []
    return [{"管理番号": min(st_m), "発生日": jp_date(month_last_day(month)),
             "支払期日": jp_date(f"{ym_add(month, 1)}-10"), "取引先": "住民税", "rows": rows}]


def build_shaho(month, prev, st_prev, st_m, ridx, resolver, master_rows, kanri, torihikisaki,
                azukari_specs, alerts, kyoshutsukin=False, split_midmonth=True):
    """健保/厚年ファイルの4取引を作る（2026-07-24 実データで構造確定）:

      ① 主取引     発生日=前月末・期日=当月末 … 在籍者の会社負担（法定福利費）
      ② 預り金     発生日=当月末・期日=当月末 … ①に対応する従業員預り分の取り崩し
      ③ 退職者分   発生日=当月末・期日=翌月末 … 当月退職者の翌月分（社保2倍回収の会社負担）
      ④ 退職者預り 発生日=翌月末・期日=翌月末 … ③に対応する預り分

    B7: 健保=健保＋介護＋子ども子育て支援金の合算／厚年=厚年のみ。
        子ども・子育て拠出金は全員分を1行に一括計上（品目=人件費（本社）・厚年ファイル）。
    """
    prev_end = month_last_day(prev)
    month_end = month_last_day(month)
    next_end = month_last_day(ym_add(month, 1))

    def shaho_total(pi):
        return sum(int(r["amount_sign"]) * pi_value(pi, r["source_key"]) for r in master_rows)

    def is_retiree(emp):
        """当月末で退職した人だけが社保2倍回収の対象。

        社会保険の資格喪失は退職日の翌日なので、月中退職（例: 池村2025022 の 7/15）は
        当月分の保険料が発生せず、前月分だけを通常どおり計上する。
        """
        return str(ridx.get(emp, {}).get("retired_on") or "") == month_end

    def is_new_hire(emp):
        """入社月が対象月（＝前月）の人。

        jinjer は入社月の明細に社保を載せない（控除開始が翌月給与のため）が、
        社保の発生自体は入社月から。最終CSVは翌月＝当月明細の値で入社月分を計上している
        （2026-07 実測: 橘2026012 の 25,175＝7月明細の健保24,566＋子育て609、
          川口2026013 の 13,300＝12,978＋322。厚年 48,495／25,620 も当月明細と一致）。
        """
        return str(ridx.get(emp, {}).get("joined_on") or "")[:7] == prev

    def source_pi(emp):
        """当月退職者は社保を2倍徴収されているので、当月値の半分を1か月分として使う。
        入社月の人は前月明細が空なので当月明細を使う。それ以外は前月の値をそのまま使う。"""
        if is_retiree(emp) and emp in st_m:
            return st_m[emp], 0.5
        if (is_new_hire(emp) and emp in st_m
                and not shaho_total(st_prev[emp]) and shaho_total(st_m[emp])):
            alerts["shaho_new_hire"].add(
                (emp, ridx.get(emp, {}).get("name", emp),
                 str(ridx.get(emp, {}).get("joined_on") or ""), int(shaho_total(st_m[emp]))))
            return st_m[emp], 1.0
        return st_prev[emp], 1.0

    # ②預り金・拠出金の集計対象（当月給与から控除・負担が発生する人）
    azukari_emps = []
    for emp in sorted(st_m):
        retired = str(ridx.get(emp, {}).get("retired_on") or "")
        if retired and retired <= prev_end:
            continue                      # 前月末までに退職＝前月CSVで処理済み
        if is_shaho_menjo(st_m[emp]):
            continue                      # 育休等の社保免除
        azukari_emps.append((emp, 0.5 if is_retiree(emp) else 1.0))

    # --- ① 主取引（当月も在籍している人。前月末で退職した人は前月のCSVで処理済み）---
    main_rows = []
    azukari_base = []            # ②の集計対象（①に載せた人）
    for emp in sorted(st_prev):
        if emp not in st_m:
            continue
        retired = str(ridx.get(emp, {}).get("retired_on") or "")
        if retired and retired <= prev_end:
            continue          # 前月末までに退職＝前月CSVの③④で処理済み
        if is_shaho_menjo(st_prev[emp]):
            alerts["shaho_menjo"].add((emp, ridx.get(emp, {}).get("name", emp), prev))
            continue          # 育休等で社保免除＝計上しない
        pi_src, ratio = source_pi(emp)
        total = shaho_total(pi_src) * ratio
        if not total:
            continue
        azukari_base.append((emp, pi_src, ratio))
        name = ridx.get(emp, {}).get("name", emp)
        idou = midmonth_date(resolver.histories, emp, prev) if split_midmonth else None
        if idou:   # ルール9(b): 法定福利費も旧部門／新部門で÷2
            old_amt, new_amt = split_halves(total)
            first = f"{prev}-01"
            main_rows.append(detail_row(
                "法定福利費", "対象外", old_amt, resolver.jinkenhi_item(emp, first),
                resolver.bumon(emp, first, f"{prev}社保・異動前"), name, f"{idou} 異動（異動前分）"))
            main_rows.append(detail_row(
                "法定福利費", "対象外", new_amt, resolver.jinkenhi_item(emp, prev_end),
                resolver.bumon(emp, prev_end, f"{prev}社保"), name, f"{idou} 異動（異動後分）"))
        else:
            main_rows.append(detail_row(
                "法定福利費", "対象外", total, resolver.jinkenhi_item(emp, prev_end),
                resolver.bumon(emp, prev_end, f"{prev}社保"), name))
    if kyoshutsukin:
        # 拠出金も預り金と同じく当月の値（2026-07 実測: 7月の other19 合計 295,056 =
        # 主取引 293,832 ＋ 退職者取引 1,224 で完全一致）
        total = sum(pi_value(st_m[e], KYOSHUTSUKIN_KEY) * ratio for e, ratio in azukari_emps)
        if total:
            main_rows.append(detail_row("法定福利費", "対象外", total, "人件費（本社）",
                                        "本社", SONOTA_HONSHA, KYOSHUTSUKIN_BIKO))

    # --- ② 預り金の取り崩し ---
    # ①（前月分の費用）と違い、**当月給与から控除した分**なので当月(st_m)の値を使う
    #   （2026-07 実測: 介護 451,260 が7月の other16 と完全一致。6月の値では合わない）。
    #   当月退職者は2倍徴収されているため、半分は④へ回す。
    azukari_rows = []
    for key, item in azukari_specs:
        total = sum(pi_value(st_m[e], key) * ratio for e, ratio in azukari_emps)
        if total:
            azukari_rows.append(detail_row("預り金", "対象外", total, item, "本社", SONOTA_HONSHA))

    # --- ③④ 当月退職者の翌月分（社保2倍回収の残り半分）---
    retire_rows, retire_azukari, retire_base = [], [], []
    for emp in sorted(st_m):
        if not is_retiree(emp):
            continue
        total = shaho_total(st_m[emp]) * 0.5
        if not total:
            continue
        name = ridx.get(emp, {}).get("name", emp)
        retire_base.append(emp)
        retire_rows.append(detail_row(
            "法定福利費", "対象外", total, resolver.jinkenhi_item(emp, month_end),
            resolver.bumon(emp, month_end, f"{month}退職者社保"), name))
        alerts["retiree"].add((emp, name, str(ridx.get(emp, {}).get("retired_on")), int(total)))
    if retire_rows:
        if kyoshutsukin:
            total = sum(pi_value(st_m[e], KYOSHUTSUKIN_KEY) * 0.5 for e in retire_base)
            if total:
                retire_rows.append(detail_row("法定福利費", "対象外", total, "人件費（本社）",
                                              "本社", SONOTA_HONSHA, KYOSHUTSUKIN_BIKO))
        for key, item in azukari_specs:
            total = sum(pi_value(st_m[e], key) * 0.5 for e in retire_base)
            if total:
                retire_azukari.append(detail_row("預り金", "対象外", total, item,
                                                 "本社", SONOTA_HONSHA))

    transactions = []
    head = {"管理番号": kanri, "取引先": torihikisaki}
    if main_rows:
        transactions.append(dict(head, 発生日=jp_date(prev_end), 支払期日=jp_date(month_end),
                                 rows=main_rows))
    if azukari_rows:
        transactions.append(dict(head, 発生日=jp_date(month_end), 支払期日=jp_date(month_end),
                                 rows=azukari_rows))
    if retire_rows:
        transactions.append(dict(head, 発生日=jp_date(month_end), 支払期日=jp_date(next_end),
                                 rows=retire_rows))
    if retire_azukari:
        transactions.append(dict(head, 発生日=jp_date(next_end), 支払期日=jp_date(next_end),
                                 rows=retire_azukari))
    return transactions


def write_freee_csv(path, transactions):
    with open(path, "w", encoding="cp932", errors="replace", newline="") as f:
        w = csv.writer(f)
        w.writerow(FREEE_COLS)
        for t in transactions:
            head = {"収支区分": "支出", "管理番号": t["管理番号"], "発生日": t["発生日"],
                    "支払期日": t["支払期日"], "取引先": t["取引先"]}
            for i, row in enumerate(t["rows"]):
                merged = dict(row)
                merged.update(head if i == 0 else {k: "" for k in head})
                w.writerow([merged.get(c, "") for c in FREEE_COLS])


# ---------------------------------------------------------------------------
# アラート・検算
# ---------------------------------------------------------------------------
def check_new_usage(master, st_maps, alerts):
    """マスタで『対象外(全期間ゼロ)』の項目に値が出ていないか（新規使用の検知網）。"""
    for ym, st in st_maps.items():
        for emp, pi in st.items():
            for atype in ("salary_items", "salary_deduction_items", "salary_other_items"):
                for it in pi.get(atype, []) or []:
                    n = to_number(it.get("value"))
                    if not n:
                        continue
                    key = f"{atype}:{it.get('id')}"
                    rows = master["by_key"].get(key, [])
                    if rows and all(r["target_file"] == "対象外" and "全期間ゼロ" in r.get("evidence", "")
                                    for r in rows):
                        alerts["new_usage"].add((key, str(it.get("label")), ym))
                    if key in WATCH_KEYS:
                        alerts["watch_used"].add((key, ym, emp, int(n)))


def detect_midmonth(histories, month, prev, alerts):
    for emp, hist in histories.items():
        for ym in (prev, month):
            for h in midmonth_records(hist, ym):
                alerts["midmonth"].add((emp, h["date"]))


def totals_by(transactions, key):
    c = Counter()
    for t in transactions:
        for r in t["rows"]:
            c[r[key]] += r["金額"]
    return c


def special_rule_owner(source_key, emp):
    """master 行の代わりに特別ルールで生成している組み合わせなら、その理由を返す。"""
    if emp == YAKUIN_LOAN["employee_id"] and source_key in YAKUIN_LOAN["source_keys"]:
        return "役員貸付金の返済スケジュールから生成（元本＋受取利息の2行）"
    return ""


def master_skipped_report(master, st_m, ridx=None):
    """status で採用しなかったマスタ行に、当月そもそも金額が出ているかを見に行く。

    「採用マスタ行: 38行（status除外: 推定=1）」だけでは何を見る欄か分からない、という
    指摘（谷津さん 2026-07-28）への対応。除外した項目名と、当月の実データで金額が
    出ているかまで出し、出ていなければ「影響なし」と明記する。
    特別ルールで別途生成している項目（役員貸付金など）は「計上済み」と区別して出す
    ——ここで誤警報を出すと、本物の漏れを見落とす欄になってしまうため。
    """
    ridx = ridx or {}
    out = []
    for r in master["skipped_rows"]:
        missing, handled, total = [], [], 0.0
        for emp, pi in st_m.items():
            v = pi_value(pi, r["source_key"])
            if not v:
                continue
            total += v
            who = f"{emp} {ridx.get(emp, {}).get('name', '')}".strip()
            why = special_rule_owner(r["source_key"], emp)
            (handled if why else missing).append((who, v, why))
        if missing:
            impact = (f"⚠️ 当月に金額あり（{len(missing)}名 {sum(v for _w, v, _y in missing):,.0f}円）"
                      "＝生成CSVから漏れています: "
                      + "、".join(f"{w} {v:,.0f}円" for w, v, _y in missing))
        elif handled:
            impact = ("影響なし（特別ルールで計上済み: "
                      + "、".join(f"{w} {v:,.0f}円 … {y}" for w, v, y in handled) + "）")
        else:
            impact = "影響なし（当月は該当者なし）"
        out.append({"row": r, "people": len(missing) + len(handled), "total": total,
                    "impact": impact})
    return out


def build_kensan(month, prev, files, st_m, st_prev, master, alerts=None, ridx=None):
    lines = [f"# 検算シート {ym_compact(month)}（C-2骨格）", "",
             f"生成: 勤怠チェッカー 経理モード（{now_iso()}）", "",
             "| ファイル | 取引数 | 行数 | 金額合計 |", "|---|---|---|---|"]
    for name, trans in files.items():
        n_rows = sum(len(t["rows"]) for t in trans)
        total = sum(r["金額"] for t in trans for r in t["rows"])
        lines.append(f"| {name} | {len(trans)} | {n_rows} | {total:,} |")
    lines += ["", "## 勘定科目別合計（給与ファイル）", "", "| 勘定科目 | 合計 |", "|---|---|"]
    for k, v in sorted(totals_by(files["給与"], "勘定科目").items(), key=lambda kv: -abs(kv[1])):
        lines.append(f"| {k} | {v:,} |")
    lines += ["", "## 品目別合計（給与ファイル）", "", "| 品目 | 合計 |", "|---|---|"]
    for k, v in sorted(totals_by(files["給与"], "品目").items(), key=lambda kv: -abs(kv[1])):
        lines.append(f"| {k} | {v:,} |")
    # jinjer 側との突合（構成上一致するはずの代表値＝パイプライン健全性チェック）
    soosai_rows = sorted((alerts or {}).get("juminzei_soosai", set()))
    soosai = sum(t for _e, _n, t, _a in soosai_rows)
    juminzei_api = sum(pi_value(pi, "salary_deduction_items:deduction41") for pi in st_m.values())
    juminzei_csv = sum(r["金額"] for t in files["住民税"] for r in t["rows"])
    # 前月に会社が立て替えた分は前月CSVで計上済み＝当月は差し引いて出すのが正しい。
    # 素の合計と比べると毎月「不一致」で赤くなり本物の異常を見落とすため、相殺後で比べる。
    zansa = juminzei_api - soosai - juminzei_csv
    if abs(zansa) < 1:
        juminzei_verdict = "一致" + (f"（うち前月立替の相殺 {soosai:,}円・{len(soosai_rows)}名）"
                                   if soosai else "")
    else:
        juminzei_verdict = f"**不一致（残差 {zansa:,.0f}円）**"
    juminzei_line = f"- 住民税: API合計 {juminzei_api:,.0f}"
    if soosai:
        juminzei_line += f" − 前月立替の相殺 {soosai:,} = {juminzei_api - soosai:,.0f}"
    juminzei_line += f" vs 住民税ファイル {juminzei_csv:,} → {juminzei_verdict}"
    lines += ["", "## パイプライン健全性（構成上一致するはずの値）", "", juminzei_line]
    for emp, name, tatekae, after in soosai_rows:
        lines.append(f"    - 相殺の内訳: {emp} {name} 立替 {tatekae:,}円 → 当月計上 {after:,}円")
    lines += [f"- 対象従業員: 当月={len(st_m)}名 / 前月={len(st_prev)}名"]

    # マッピング表のうち今回の仕訳に使った行数（何を見る欄か分かるように内訳を書く）
    skipped_report = master_skipped_report(master, st_m, ridx)
    off_n = sum(master["offscope"].values())
    off_detail = "・".join(f"{k} {v}行" for k, v in master["offscope"].items()) or "なし"
    lines += ["", "## 採用マスタ行（マッピング表のうち今回の仕訳に使った行）", "",
              "マッピング表＝ジンジャーの支給・控除項目を freee の品目へ対応づけた表"
              f"（`{os.path.basename(master['path'])}`）。この欄はその表のうち何行を今回使ったかを示す。", "",
              f"- マッピング表 全{master['total_n']}行",
              f"- うち仕訳に使わない {off_n}行（{off_detail}）＝ freee に計上しない項目",
              f"- 残り {master['total_n'] - off_n}行のうち status「{master['min_status']}」までを採用 → "
              f"**{master['active_n']}行を使用**"]
    if skipped_report:
        lines += ["", f"### status で採用しなかった {len(skipped_report)}行", "",
                  "| 項目 | source_key | status | 当月の金額 | 判定 |", "|---|---|---|---|---|"]
        for s in skipped_report:
            r = s["row"]
            amt = f"{s['total']:,.0f}円（{s['people']}名）" if s["people"] else "0円（該当者なし）"
            lines.append(f"| {r['label']} | {r['source_key']} | {r['status']} | {amt} | {s['impact']} |")
    else:
        lines += ["", f"- status で除外した行はなし（{master['min_status']}以外の行が無い）"]
    return lines


def detect_mishunyukin(st_m, st_prev, ridx, alerts):
    """B8: 社保を給与から徴収しきれていない人＝未収入金・未払金の手動計上候補を検知する。

    検知条件: 会社負担（other15/17）が発生しているのに従業員控除（deduction29/31）が
    ゼロ、または差引支給額がマイナス（回収不能）。
    """
    for ym, st in (("当月", st_m), ("前月", st_prev)):
        for emp, pi in st.items():
            name = ridx.get(emp, {}).get("name", emp)
            kaisha = pi_value(pi, "salary_other_items:other15") + pi_value(pi, "salary_other_items:other17")
            honnin = pi_value(pi, "salary_deduction_items:deduction29") + pi_value(pi, "salary_deduction_items:deduction31")
            net = pi_value(pi, "salary_payment_items:payment1")
            if kaisha and not honnin:
                pass   # 育休等の社保免除 → 社保ファイル側で除外・別リストへ
            elif net < 0:
                alerts["mishunyukin"].add((ym, emp, name, "差引支給額がマイナス", int(net)))


def detect_juminzei_shokai(st_m, st_prev, ridx, alerts):
    """住民税の特別徴収が当月から始まった人を検知する。

    新入社員は前年所得がないため特別徴収の開始が遅れ、開始月に複数月分が
    まとめて控除されることがある（2026-07 実測: 柳場2026010 が 21,400 に対し
    最終CSVは 10,300）。前月分は前月CSVで計上済みのことがあるため要確認に出す。
    """
    key = "salary_deduction_items:deduction41"
    for emp, pi in st_m.items():
        cur = pi_value(pi, key)
        prev_v = pi_value(st_prev[emp], key) if emp in st_prev else 0.0
        if cur and not prev_v:
            alerts["juminzei_shokai"].add((emp, ridx.get(emp, {}).get("name", emp), int(cur)))


def build_yokakunin(month, alerts, master):
    lines = [f"# 要確認リスト {ym_compact(month)}（C-2）", ""]
    lines += ["## 骨格の未実装（TODO）", ""] + [f"- {t}" for t in SKELETON_TODOS]
    lines += ["", "## ⚠️ 未収入金・未払金の候補（B8: 手で追加が必要）", "",
              "急な退職等で社保を給与から徴収しきれない場合、経理が未収入金として手計上している。"
              "システム化対象外のため、下記に該当者がいれば手で行を追加すること。", ""]
    if alerts["mishunyukin"]:
        lines += ["| 月 | 社員番号 | 氏名 | 検知理由 | 金額 |", "|---|---|---|---|---|"]
        for ym, emp, name, why, amt in sorted(alerts["mishunyukin"]):
            lines.append(f"| {ym} | {emp} | {name} | {why} | {amt:,} |")
    else:
        lines.append("- なし")
    lines += ["", "## 経費転記の品目分解（C-4・実施済み）", "",
              "jinjer給与の「その他」(allowance52) を経費一覧表マクロの明細（集計ログの J:その他）から"
              "消耗品費／社員育成費／会議接待費／通信費／福利厚生費／雑費 に分解した。"
              "**品目の判定根拠を全件出すので目視で確認すること。**"
              "備考はマクロ側の備考(明細)をそのまま入れているので、必要なら書き直すこと。", ""]
    if alerts["keihi_bunkai"]:
        lines += ["| 社員番号 | 氏名 | 内訳 | 金額 | 判定した品目 | 根拠 | 確度 |",
                  "|---|---|---|---|---|---|---|"]
        for emp, name, uchiwake, amt, item, why, status in sorted(alerts["keihi_bunkai"]):
            mark = "" if status == "確定" else f" **{status}**"
            lines.append(f"| {emp} | {name} | {uchiwake} | {amt:,} | {item} | {why} |{mark or ' 確定'} |")
        total = sum(a for _e, _n, _u, a, _i, _w, _s in alerts["keihi_bunkai"])
        lines.append(f"| **合計** | | | **{total:,}** | | | |")
    else:
        lines.append("- なし")
    lines += ["", "## ⚠️ 経費転記で分解できず計上しなかった人（手で追加が必要）", "",
              "明細の合計が jinjer の「その他」と一致しないため、行を作っていない。"
              "前月分をまとめて計上した／マクロ側に重複行がある／経費申請ではなく手入力した、"
              "などの理由がある。金額と品目を確認して手で追加すること。", ""]
    if alerts["keihi_tenki"]:
        lines += ["| 社員番号 | 氏名 | jinjerの「その他」 | 明細の合計 | 部門 |", "|---|---|---|---|---|"]
        for emp, name, amt, bumon, det_total in sorted(alerts["keihi_tenki"],
                                                       key=lambda x: (x[0], x[1])):
            shown = f"{det_total:,}" if det_total is not None else "明細なし"
            lines.append(f"| {emp} | {name} | {amt:,} | {shown} | {bumon} |")
        total = sum(a for _e, _n, a, _b, _d in alerts["keihi_tenki"])
        lines.append(f"| **合計** | | **{total:,}** | | |")
    else:
        lines.append("- なし")
    if alerts["keihi_book_missing"]:
        lines += ["", "### ⚠️ 経費一覧表マクロのブックが見つからない", "",
                  "経費転記の分解ができないため、上の表に全員が並ぶ。ブックの場所を確認すること。", ""]
        for p in sorted(alerts["keihi_book_missing"]):
            lines.append(f"- `{p}`")
    lines += ["", "## 差引支給額がマイナスのため仮払金を計上した人", "",
              "給与から控除しきれなかった分を仮払金として計上している（後日回収）。", ""]
    if alerts["karibarai"]:
        for emp, name, amt in sorted(alerts["karibarai"]):
            lines.append(f"- {emp} {name}: {amt:,}円")
    else:
        lines.append("- なし")
    lines += ["", "## 立替金を計上しなかった人（本社＝freee へ直接申請済み）", ""]
    if alerts["tatekae_skip"]:
        for emp, name, amt in sorted(alerts["tatekae_skip"]):
            lines.append(f"- {emp} {name}: {amt:,}円（ビジネス企画部・管理監督者は通勤費以外を"
                         "freee へ直接申請しているため計上しない）")
    else:
        lines.append("- なし")
    lines += ["", "## 備考の手入力が必要な行", ""]
    if alerts["biko_needed"] or alerts["genbutsu"]:
        split_emps = {e for e, _n, _a, _p in alerts["shaho_chosei_split"]}
        for emp, name, kind, amt in sorted(alerts["biko_needed"]):
            lines.append(f"- {emp} {name}: {kind} {amt:,}円 → 備考に理由を記入"
                         "（例「社保調整分（4月、5月月変対応漏れのため）」）"
                         + ("" if emp in split_emps else
                            "。**按分で割り切れないため品目を分けていない**"
                            "＝品目が正しいか確認すること（実績: 端数調整は雇用保険料（預り分）"
                            "だった月がある）"))
        for ym, n, bulk in sorted(alerts["genbutsu"]):
            lines.append(f"- {ym} 現物支給 {n}名（支給側は{'人件費行へ合算' if bulk else '独立行'}）"
                         f"→ 備考「{GENBUTSU_BIKO}」を実際の文言（例「カタログギフト現物支給」）に修正")
    else:
        lines.append("- なし")
    if alerts["loan_mismatch"]:
        lines += ["", "## ⚠️ 役員貸付金の分解が合わない", ""]
        for ym, api_v, calc in sorted(alerts["loan_mismatch"]):
            lines.append(f"- {ym}: API家賃控除 {api_v:,.0f} vs 返済額 {calc:,}（元利均等パラメータ要更新）")
    lines += ["", "## 月中異動の÷2分割（ルール9・実施済み）", "",
              "当月払いのため、異動月の翌月CSVの実績差分（＝異動月分の費用）で分割する。", ""]
    if alerts["split_done"]:
        lines += ["| 社員番号 | 氏名 | 異動日 | 異動前分 | 異動後分 |", "|---|---|---|---|---|"]
        for emp, name, date, old, new in sorted(alerts["split_done"]):
            lines.append(f"| {emp} | {name} | {date} | {old:,} | {new:,} |")
    else:
        lines.append("- 対象なし（前月内に月中日付の履歴レコードなし）")
    lines += ["", "## 月中異動の検知（当月・翌月のCSVで分割対象になる）", ""]
    if alerts["midmonth"]:
        for emp, date in sorted(alerts["midmonth"]):
            lines.append(f"- {emp}: 履歴日付 {date}")
    else:
        lines.append("- なし")
    lines += ["", "## 住民税の特別徴収が当月から始まった人（金額の確認が必要）", "",
              "新入社員などで特別徴収の開始月には複数月分がまとめて控除されることがある。"
              "前月分を前月CSVで計上済みの場合は当月分だけに調整すること。", ""]
    if alerts["juminzei_soosai"]:
        for emp, name, tatekae, after in sorted(alerts["juminzei_soosai"]):
            lines.append(f"- {emp} {name}: 前月の立替 {tatekae:,}円を差し引いて {after:,}円を計上")
    for emp, name, amt in sorted(alerts["juminzei_shokai"]):
        if not any(e == emp for e, _n, _t, _a in alerts["juminzei_soosai"]):
            lines.append(f"- {emp} {name}: {amt:,}円（前月の立替が見つからずそのまま計上）")
    if not alerts["juminzei_shokai"]:
        lines.append("- なし")
    lines += ["", "## 社保免除で社保ファイルに載せなかった人（育休・産休）", "",
              "本人負担がゼロなのに会社負担だけ jinjer に計上されている人＝育児休業等で"
              "社会保険料が免除されている人（谷津さん確認 2026-07-24）。労使折半の原則から"
              "会社負担も計上しない。**jinjer 側の会社負担額が残っている点は要確認**。", ""]
    if alerts["shaho_menjo"]:
        for emp, name, ym in sorted(alerts["shaho_menjo"]):
            lines.append(f"- {emp} {name}（{ym}分）")
    else:
        lines.append("- なし")
    lines += ["", "## 入社月の社保を当月明細から拾った人", "",
              "jinjer は入社月の明細に社保を載せない（控除開始が翌月給与のため）が、"
              "社保の発生自体は入社月から。前月明細がゼロで当月明細に値がある入社月の人は、"
              "当月明細の値で前月分（健保＝健保＋介護＋子ども子育て支援金／厚年＝厚年）を計上する。", ""]
    if alerts["shaho_new_hire"]:
        lines += ["| 社員番号 | 氏名 | 入社日 | 会社負担 |", "|---|---|---|---|"]
        for emp, name, joined, amt in sorted(alerts["shaho_new_hire"]):
            lines.append(f"| {emp} | {name} | {joined} | {amt:,} |")
    else:
        lines.append("- なし")
    lines += ["", "## 社保調整の品目別分割（実施済み）", "",
              "jinjer の社保調整(deduction4)は1項目だが、freee では品目別に分かれる。"
              "本人の当月控除額（健保・介護・厚年・子育て支援金）で按分している。", ""]
    if alerts["shaho_chosei_split"]:
        for emp, name, amt, parts in sorted(alerts["shaho_chosei_split"]):
            uchiwake = "／".join(f"{item} {a:,}" for item, a in parts)
            lines.append(f"- {emp} {name}: {amt:,}円 → {uchiwake}")
    else:
        lines.append("- なし")
    lines += ["", "## 当月退職者の社保2倍回収（翌月分を当月計上）", ""]
    if alerts["retiree"]:
        lines += ["| 社員番号 | 氏名 | 退職日 | 翌月分の会社負担 |", "|---|---|---|---|"]
        for emp, name, retired, amt in sorted(alerts["retiree"]):
            lines.append(f"| {emp} | {name} | {retired} | {amt:,} |")
    else:
        lines.append("- なし")
    lines += ["", "## 育成期間だったかもしれない人（人件費区分が既定に落ちた新入社員）", "",
              f"jinjer の人件費区分は 本社／育成 の例外だけを持ち、空欄は既定の"
              f"{JINKENHI_KUBUN_DEFAULT}になる。育成のレコードは 2026-07 の一括投入以降しか"
              "無いため、それ以前の月は育成を復元できない。入社から"
              f"{IKUSEI_WATCH_MONTHS}か月以内の人が既定に落ちたら、"
              "本当は 人件費（育成）ではないか確認すること。", ""]
    if alerts["ikusei_maybe"]:
        lines += ["| 社員番号 | 入社日 | 対象月 |", "|---|---|---|"]
        for emp, joined, ym in sorted(alerts["ikusei_maybe"]):
            lines.append(f"| {emp} | {joined} | {ym} |")
    else:
        lines.append("- なし")
    lines += ["", "## 部門が対象日時点で解決できず、最古の履歴で代用した従業員", ""]
    if alerts["bumon_missing"]:
        for emp, ctx, used in sorted(alerts["bumon_missing"]):
            lines.append(f"- {emp}（{ctx}）: 対象日以前の履歴なし → 「{used}」を使用")
    else:
        lines.append("- なし")
    lines += ["", "## 部門マスタ（8種）に無い部門値", ""]
    if alerts["bumon_unknown"]:
        for emp, v in sorted(alerts["bumon_unknown"]):
            lines.append(f"- {emp}: 「{v}」")
    else:
        lines.append("- なし")
    lines += ["", "## 『対象外(全期間ゼロ)』項目の新規使用検知", ""]
    if alerts["new_usage"]:
        for key, label, ym in sorted(alerts["new_usage"]):
            lines.append(f"- {ym}: {label}（{key}）に値 → マスタ再マッピングが必要")
    else:
        lines.append("- なし")
    lines += ["", "## 見張り項目に値が出た（対象外の判断が正しいか確認）", ""]
    if alerts["watch_used"]:
        for key, reason in sorted(WATCH_KEYS.items()):
            hits = [(ym, emp, n) for k, ym, emp, n in alerts["watch_used"] if k == key]
            if not hits:
                continue
            total = sum(n for _ym, _e, n in hits)
            lines.append(f"- **{key}**: {len(hits)}人月・計 {total:,}円 → {reason}")
            for ym, emp, n in sorted(hits):
                lines.append(f"  - {ym} {emp}: {n:,}")
    else:
        lines.append("- なし")
    return lines


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------
def generate(month, out_base=None, master_csv=None, keihi_mapping_csv=None,
             keihi_book=None, final_csv_dir=None, min_status="確定",
             refresh_statements=False, refresh_custom=False, client=None):
    """支給月ぶんの4CSV＋検算＋要確認を作る。

    Args:
        month: 支給月 'yyyy-MM'
        out_base: 出力の親フォルダ（既定 Config.KEIRI_OUTPUT_DIR）。{YYYYMM} を掘る
        master_csv / keihi_mapping_csv: マッピング表（空欄なら Config の共有フォルダ）
        keihi_book: 経費利用履歴 RevN.xlsm。空欄なら {M}月フォルダから自動検出
        final_csv_dir: 経理の最終CSVの親フォルダ（住民税の前月立替を読むのに使う）
        min_status: マスタの採用範囲（本番は '確定' のみ）
        refresh_statements / refresh_custom: API を取り直すか（既定はキャッシュ優先）

    Returns:
        dict — out_dir / files（種別→{path,transactions,rows}）/ kensan_path /
               yokakunin_path / alerts の要約
    """
    out_base = out_base or OUT_BASE
    prev = ym_add(month, -1)
    master = load_master(master_csv or MASTER_CSV, min_status)

    cache_dir = out_base
    st_m = load_statements(cache_dir, month, client=client, refresh=refresh_statements)
    st_prev = load_statements(cache_dir, prev, client=client, refresh=refresh_statements)
    if not st_m:
        raise ValueError(f"{month} の給与明細が jinjer から取得できませんでした。"
                         "給与計算が済んでいる月か確認してください。")
    roster = load_or_fetch_roster(client, os.path.join(cache_dir, "raw", "roster.json"),
                                  refresh=refresh_statements)
    ridx = roster_index(roster)
    for emp, alias in NAME_ALIASES.items():        # freee 側の表記に合わせる
        if emp in ridx:
            ridx[emp]["name"] = alias

    target_ids = {e for e in ridx if classify_employee(e) == "target"}
    histories = load_custom_histories(out_base, target_ids, refresh=refresh_custom)
    alerts = {"bumon_missing": set(), "bumon_unknown": set(), "midmonth": set(),
              "new_usage": set(), "mishunyukin": set(), "biko_needed": set(),
              "genbutsu": set(), "kyushoku": set(), "loan_mismatch": set(),
              "keihi_tenki": set(), "keihi_bunkai": [], "keihi_book_missing": set(),
              "tatekae_skip": set(), "split_done": set(),
              "retiree": set(), "shaho_menjo": set(), "juminzei_shokai": set(),
              "juminzei_soosai": set(), "karibarai": set(),
              "shaho_chosei_split": set(), "shaho_new_hire": set(), "watch_used": set(),
              "ikusei_maybe": set()}
    resolver = Resolver(histories, alerts, ridx)

    paid_counter = Counter()
    for pi in st_m.values():
        p = statement_flag(None, {"payroll_info": pi}, "paid_on")
        if p:
            paid_counter[str(p)] += 1
    paid_on = paid_counter.most_common(1)[0][0] if paid_counter else ""

    # C-4: 経費転記の分解材料（経費一覧表マクロのブック）。無ければ検知だけする
    keihi_mapping = keihi_load_mapping(keihi_mapping_csv or KEIHI_MAPPING_CSV)
    book = keihi_book or keihi_find_book(month)
    keihi_details = {}
    if book and os.path.exists(book):
        keihi_details = keihi_load_details(book)
    else:
        alerts["keihi_book_missing"].add(book or f"{int(month[5:7])}月フォルダに見つからず")

    detect_juminzei_shokai(st_m, st_prev, ridx, alerts)
    prev_juminzei = load_prev_juminzei(prev, out_base, final_dir=final_csv_dir)
    files = {
        "給与": build_kyuyo(month, prev, st_m, st_prev, ridx, resolver, master, paid_on, alerts,
                          keihi_details=keihi_details, keihi_mapping=keihi_mapping),
        "住民税": build_juminzei(month, st_m, ridx, resolver, master, prev_juminzei, alerts),
        "健康保険": build_shaho(month, prev, st_prev, st_m, ridx, resolver, master["kenpo"],
                            "KEMPO", "関東ＩＴソフトウェア健康保険組合", KENPO_AZUKARI, alerts),
        # 拠出金は従業員別に配分せず一括1行にするため、従業員別合算の対象からは外す
        "厚生年金": build_shaho(month, prev, st_prev, st_m, ridx, resolver,
                            [r for r in master["konen"] if r["source_key"] != KYOSHUTSUKIN_KEY],
                            "KONEN", "厚生労働省", KONEN_AZUKARI, alerts, kyoshutsukin=True),
    }
    detect_midmonth(histories, month, prev, alerts)
    check_new_usage(master, {month: st_m, prev: st_prev}, alerts)
    detect_mishunyukin(st_m, st_prev, ridx, alerts)
    detect_juminzei_shokai(st_m, st_prev, ridx, alerts)   # 住民税の相殺判定に先立って実行する

    mc, pc = ym_compact(month), ym_compact(prev)
    out_dir = os.path.join(out_base, mc)
    os.makedirs(out_dir, exist_ok=True)
    names = {
        "給与": f"freee_data_給与（{pc}実績差分+{mc}暫定支給分）.csv",
        "住民税": f"freee_data_住民税（{mc}）.csv",
        "健康保険": f"freee_data_健康保険（{pc}+{mc}退職者分）.csv",
        "厚生年金": f"freee_data_厚生年金（{pc}+{mc}退職者分）.csv",
    }
    result_files = {}
    for key, trans in files.items():
        path = os.path.join(out_dir, names[key])
        write_freee_csv(path, trans)
        result_files[key] = {"path": path, "filename": names[key],
                             "transactions": len(trans),
                             "rows": sum(len(t["rows"]) for t in trans)}

    kensan_path = os.path.join(out_dir, f"検算_{mc}.md")
    with open(kensan_path, "w", encoding="utf-8") as f:
        f.write("\n".join(build_kensan(month, prev, files, st_m, st_prev, master, alerts, ridx)))
    yokakunin_path = os.path.join(out_dir, f"要確認_{mc}.md")
    with open(yokakunin_path, "w", encoding="utf-8") as f:
        f.write("\n".join(build_yokakunin(month, alerts, master)))

    return {
        "month": month, "prev": prev, "out_dir": out_dir, "files": result_files,
        "kensan_path": kensan_path, "yokakunin_path": yokakunin_path,
        "paid_on": paid_on, "employees": len(st_m),
        "keihi_book": book if book and os.path.exists(book) else "",
        "master_csv": master_csv or MASTER_CSV,
        "alerts": {
            "月中異動検知": len(alerts["midmonth"]),
            "部門未解決": len(alerts["bumon_missing"]),
            "部門未知値": len(alerts["bumon_unknown"]),
            "対象外項目の新規使用": len(alerts["new_usage"]),
            "経費転記の分解": len(alerts["keihi_bunkai"]),
            "経費転記で保留": len(alerts["keihi_tenki"]),
            "未収入金の候補": len(alerts["mishunyukin"]),
            "備考の手入力が必要": len(alerts["biko_needed"]),
            "育成かもしれない人": len(alerts["ikusei_maybe"]),
        },
    }
