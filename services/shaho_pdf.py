r"""社労士（前田事務所）から届く「保険料一覧表」PDFを読む（確定的パース）。

健診PDF（`health_hpm_pdf.py`）と違い **AI を使わない**。この帳票にはテキスト層があり、
pdfplumber の `extract_text()` だけで全項目が取れることを実測で確認した
（2026-08-17・36名/36名。見本 = 7月保険料一覧表（8月給与控除分））。
確定的に読めるものに課金と非決定性を持ち込まない、という方針は
`kdx_shift_parser.py`（シフト表）と同じ。

## 帳票の構造（実測）

1名が **3行** で出る。列見出し自体も3行に折り返されている。

    1073 2011001 齋藤 公伸 月額変更 第２号
    73 昭55年12月06日 45 1 620 620 35,030 35,030 28,737 28,737 5,580 5,580 56,730 56,730 0 0 0 0 月額変更
    ( 0) 713 713

| 行 | 中身 |
|---|---|
| 1行目 | 健保№／**個人コード（=社員番号）**／氏名／改訂理由(健保)／介護区分 |
| 2行目 | 厚年№／生年月日／年齢／種別／**標準報酬 健保・厚年（千円単位）**／保険料6組／改訂理由(厚年) |
| 3行目 | (内特定保険)／子ども・子育て支援金（事業主・個人分）／２社勤務 |

保険料6組は左から (事業主, 個人分) の順で
**健康保険計・(健康保険)本体・(介護保険)・厚生年金・(厚生年金)基金・(厚生年金基金)**。

### ⚠ 標準報酬は千円単位。健保と厚年で値が違うことがある

`620 620` は 620,000円。厚年は上限650（65万円）で頭打ちになるため、
`1150 650`（西村稔さん）のように**健保と厚年で別の値**になる人がいる。
どちらか片方だけを見て両方に入れると上限超えの人が壊れる。

### ⚠ 改訂理由は2つある。列挙で受けない

1行目の理由＝健保、2行目の理由＝厚年。**片方だけのことがある**
（例: 健保だけ「月額変更」で厚年は上限のまま動かない → 2行目は空）。
また「料率変更」は標報が動かない月にも出る。9月には「定時決定」が来る見込みで、
理由の語彙は今後も増える。そのため**列挙せず「変更・決定・改定…で終わる語」**として拾う。

## 検算（このモジュールが持つ2段の網）

1. **PDF内部の整合**（マスタ不要・常に実行）
   健康保険計 = 本体 + 介護 + 子ども・子育て支援金。事業主・個人分の両方で成立する。
2. **料率検算**（等級表マスタを渡したとき）
   (事業主 + 個人分) = 標準報酬 × 全体料率。**足し合わせれば端数処理が消えるので厳密に一致する**。
   個人分単体は丸め方向が項目で違う（健保本体は50銭切捨てだが子育て支援金は逆に出る）ため、
   個人分の比較は許容差つきで行う。

さらに末尾の【総合計】行（人数・保険料合計）を**チェックサム**として使う。
1行でも読み落とせば人数か合計が合わなくなるので、静かな取りこぼしが起きない。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# 令和1年 = 2019年
REIWA_BASE = 2018
TITLE_KEYWORD = "保険料一覧表"

# 介護区分（1行目の最右列）。ここに載っている語だけを区分として切り出し、
# それ以外は氏名の一部として扱う（知らない語で氏名を削らない）。
KAIGO_TOKENS = frozenset({"第２号", "第2号", "第１号", "第1号", "非対象", "対象外",
                          "該当", "非該当"})
# 改訂理由は月・年度で増える（月額変更／取得時決定／料率変更／9月は定時決定）。
# 列挙せずに語尾で判定する。氏名がこれらで終わることは実務上ない。
REASON_SUFFIXES = ("変更", "決定", "改定", "訂正", "取得", "喪失")

_NUM_RE = re.compile(r"^-?[\d,]+$")
_BIRTH_RE = re.compile(r"^[明大昭平令]\d{2}年\d{2}月\d{2}日$")
_PERIOD_RE = re.compile(r"令和(\d{1,2})年(\d{1,2})月分\s*[（(]\s*令和(\d{1,2})年(\d{1,2})月給与分")
_OFFICE_RE = re.compile(r"事業所\s+(\d+)\s*[:：]\s*(.+?)\s+令和")
_TOTAL_RE = re.compile(r"^【総\s*合\s*計】\s*人数\s+(\d+)\s+(.*)$")
_KODOMO_TOTAL_RE = re.compile(r"^子ども・子育て拠出金\s+(.*)$")
_TOKUTEI_RE = re.compile(r"^[（(]\s*([\d,]+)\s*[）)]\s*(.*)$")

# 2行目の保険料6組（事業主, 個人分）の並び。標準報酬2つの後ろに続く。
PREMIUM_PAIRS = ("kenpo_total", "kenpo", "kaigo", "konen", "kikin", "kikin_fund")
# 標準報酬2 + 6組×2 = 14個の数値が2行目に並ぶ
EXPECTED_NUMBERS = 2 + len(PREMIUM_PAIRS) * 2


class ShahoPdfError(ValueError):
    """PDF全体として扱えない（別の帳票・別の事業所・チェックサム不一致など）。

    1名分の読み取り異常は例外にせず ``PdfPerson.issues`` に残して先へ進む。
    止めるべきは「この紙をこの月の投入に使ってよいか」が崩れたときだけ。
    """


@dataclass
class PdfPerson:
    """保険料一覧表の1名分。金額はすべて円。"""

    emp: str                                   # 個人コード（= 社員番号）
    name: str = ""
    kenpo_no: str = ""
    konen_no: str = ""
    birth: str = ""                            # 和暦のまま（照合には使わない）
    age: int = 0
    sex_code: str = ""                         # 種別（1/2）
    kenpo_smr: int = 0                         # 健保 標準報酬月額（円）
    konen_smr: int = 0                         # 厚年 標準報酬月額（円）
    premiums: dict = field(default_factory=dict)
    kaigo_kubun: str = ""                      # 第２号 / 非対象
    reason_kenpo: str = ""                     # 1行目の改訂理由
    reason_konen: str = ""                     # 2行目の改訂理由
    two_company: bool = False                  # ２社勤務
    page: int = 0
    issues: list = field(default_factory=list)
    # issues は「値を信用できない」＝投入させない。warnings は「投入はできるが
    # 人に伝えたいこと」（氏名の表記ゆれ・厚年が上限で頭打ち など）。
    warnings: list = field(default_factory=list)

    @property
    def reason(self) -> str:
        """画面表示用の改訂理由（健保・厚年が違うときは両方出す）。"""
        if self.reason_kenpo and self.reason_konen and self.reason_kenpo != self.reason_konen:
            return f"健保:{self.reason_kenpo} / 厚年:{self.reason_konen}"
        return self.reason_kenpo or self.reason_konen

    @property
    def ok(self) -> bool:
        return not self.issues


@dataclass
class PdfStatement:
    """保険料一覧表1枚ぶん。"""

    target_ym: str = ""                        # 保険料の対象月 "2026-07"
    pay_ym: str = ""                           # 控除される給与月 "2026-08"
    office_code: str = ""
    office_name: str = ""
    persons: list = field(default_factory=list)
    totals: dict = field(default_factory=dict)  # 【総合計】行の値
    total_count: int | None = None              # 【総合計】の人数
    warnings: list = field(default_factory=list)

    @property
    def by_emp(self) -> dict:
        return {p.emp: p for p in self.persons}


# ---------------------------------------------------------------------------
# 小道具
# ---------------------------------------------------------------------------

def _to_int(token: str) -> int:
    return int(str(token).replace(",", "").replace("　", "").strip())


def _is_num(token: str) -> bool:
    return bool(_NUM_RE.match(str(token).strip()))


def _reiwa_to_iso(year: int, month: int) -> str:
    return f"{REIWA_BASE + int(year):04d}-{int(month):02d}"


def _numbers(text: str) -> list:
    """空白区切りのトークンから数値だけを順に取り出す。"""
    return [_to_int(t) for t in str(text).split() if _is_num(t)]


# ---------------------------------------------------------------------------
# 行のパース
# ---------------------------------------------------------------------------

def _parse_line1(line: str):
    """1行目 → (健保№, 個人コード, 氏名, 改訂理由, 介護区分)。合わなければ None。"""
    tokens = line.split()
    if len(tokens) < 3:
        return None
    if not re.fullmatch(r"\d{2,5}", tokens[0]) or not re.fullmatch(r"\d{6,8}", tokens[1]):
        return None

    rest = tokens[2:]
    kaigo = ""
    if rest and rest[-1] in KAIGO_TOKENS:
        kaigo = rest.pop()
    reason = ""
    # 氏名まで理由として食べないよう、氏名が1トークンも残らなくなる場合は取らない
    if len(rest) > 1 and rest[-1].endswith(REASON_SUFFIXES) and not _is_num(rest[-1]):
        reason = rest.pop()
    name = " ".join(rest).strip()
    if not name:
        return None
    return tokens[0], tokens[1], name, reason, kaigo


def _parse_line2(line: str):
    """2行目 → (厚年№, 生年月日, 年齢, 種別, 数値リスト, 改訂理由)。合わなければ None。"""
    tokens = line.split()
    if len(tokens) < 6 or not re.fullmatch(r"\d{1,5}", tokens[0]):
        return None
    if not _BIRTH_RE.match(tokens[1]):
        return None

    reason = ""
    if tokens and tokens[-1].endswith(REASON_SUFFIXES) and not _is_num(tokens[-1]):
        reason = tokens[-1]
        tokens = tokens[:-1]

    tail = tokens[2:]
    if len(tail) < 2 or not (_is_num(tail[0]) and _is_num(tail[1])):
        return None
    age, sex = _to_int(tail[0]), tail[1]
    values = [_to_int(t) for t in tail[2:] if _is_num(t)]
    return tokens[0], tokens[1], age, sex, values, reason


def _parse_line3(line: str):
    """3行目 → (内特定保険, 子育て事業主, 子育て個人分, ２社勤務)。合わなければ None。"""
    m = _TOKUTEI_RE.match(line.strip())
    if not m:
        return None
    tokutei = _to_int(m.group(1))
    rest = m.group(2).split()
    nums = [_to_int(t) for t in rest if _is_num(t)]
    two = any("２社" in t or "2社" in t for t in rest)
    er = nums[0] if len(nums) > 0 else 0
    ee = nums[1] if len(nums) > 1 else 0
    return tokutei, er, ee, two


def _build_person(l1, l2, l3, page: int) -> PdfPerson:
    kenpo_no, emp, name, reason_k, kaigo = l1
    konen_no, birth, age, sex, values, reason_n = l2

    person = PdfPerson(emp=emp, name=name, kenpo_no=kenpo_no, konen_no=konen_no,
                       birth=birth, age=age, sex_code=sex, kaigo_kubun=kaigo,
                       reason_kenpo=reason_k, reason_konen=reason_n, page=page)

    if len(values) != EXPECTED_NUMBERS:
        person.issues.append(
            f"2行目の数値が{len(values)}個（想定{EXPECTED_NUMBERS}個）。列が想定と違う可能性")
        return person

    # 標準報酬は千円単位。円に直してから先へ渡す（以降どこでも円で扱う）
    person.kenpo_smr = values[0] * 1000
    person.konen_smr = values[1] * 1000
    for i, key in enumerate(PREMIUM_PAIRS):
        person.premiums[f"{key}_er"] = values[2 + i * 2]
        person.premiums[f"{key}_ee"] = values[3 + i * 2]

    if l3 is None:
        person.issues.append("3行目（子ども・子育て支援金）が読めない")
        person.premiums.setdefault("kodomo_er", 0)
        person.premiums.setdefault("kodomo_ee", 0)
    else:
        tokutei, kodomo_er, kodomo_ee, two = l3
        person.premiums["tokutei"] = tokutei
        person.premiums["kodomo_er"] = kodomo_er
        person.premiums["kodomo_ee"] = kodomo_ee
        person.two_company = two
        if two:
            person.issues.append("２社勤務（標準報酬が按分されている可能性）")
    return person


# ---------------------------------------------------------------------------
# 本体
# ---------------------------------------------------------------------------

def parse_text(text: str, *, expected_office: str = "") -> PdfStatement:
    """PDFから抜いたテキスト全体を PdfStatement にする（純関数・I/Oなし）。

    Args:
        text: 全ページを連結したテキスト
        expected_office: 事業所コード（指定時、一致しなければ ShahoPdfError）
    """
    lines = [ln.strip() for ln in str(text).splitlines()]
    if not any(TITLE_KEYWORD in ln for ln in lines):
        raise ShahoPdfError(
            f"「{TITLE_KEYWORD}」の見出しがありません。"
            "社労士から届く保険料一覧表のPDFを選んでください")

    stmt = PdfStatement()

    # ---- 見出し（事業所・対象月）。全ページに出るので食い違いを検知する ----
    periods, offices = set(), set()
    for ln in lines:
        m = _PERIOD_RE.search(ln)
        if m:
            periods.add((_reiwa_to_iso(int(m.group(1)), int(m.group(2))),
                         _reiwa_to_iso(int(m.group(3)), int(m.group(4)))))
        m = _OFFICE_RE.search(ln)
        if m:
            offices.add((m.group(1), m.group(2).strip()))
    if not periods:
        raise ShahoPdfError("対象月（令和○年○月分）が読み取れません")
    if len(periods) > 1:
        raise ShahoPdfError(
            "1つのPDFに複数の対象月が混じっています: "
            + "、".join(sorted(f"{t}分（{p}給与）" for t, p in periods)))
    stmt.target_ym, stmt.pay_ym = periods.pop()
    if offices:
        stmt.office_code, stmt.office_name = sorted(offices)[0]
        if len(offices) > 1:
            stmt.warnings.append("事業所名がページごとに違います")
    if expected_office and stmt.office_code and stmt.office_code != str(expected_office):
        raise ShahoPdfError(
            f"別の事業所のPDFです（{stmt.office_code}:{stmt.office_name}）。"
            f"想定は事業所コード {expected_office}")

    # ---- 従業員行 ----
    page = 1
    i = 0
    while i < len(lines):
        line = lines[i]
        if "頁" in line and "事業所" in line:
            m = re.search(r"(\d+)頁", line)
            if m:
                page = int(m.group(1))
        l1 = _parse_line1(line)
        if l1 and i + 1 < len(lines):
            l2 = _parse_line2(lines[i + 1])
            if l2:
                l3 = _parse_line3(lines[i + 2]) if i + 2 < len(lines) else None
                stmt.persons.append(_build_person(l1, l2, l3, page))
                i += 3 if l3 else 2
                continue
        m = _TOTAL_RE.match(line)
        if m:
            _parse_totals(stmt, m, lines, i)
        i += 1

    if not stmt.persons:
        raise ShahoPdfError("従業員の行が1件も読み取れませんでした（帳票の様式が変わった可能性）")

    seen = {}
    for p in stmt.persons:
        if p.emp in seen:
            p.issues.append(f"同じ個人コードが複数行あります（{p.emp}）")
            seen[p.emp].issues.append(f"同じ個人コードが複数行あります（{p.emp}）")
        seen[p.emp] = p
    return stmt


def _parse_totals(stmt: PdfStatement, m, lines: list, idx: int) -> None:
    """【総合計】ブロックを読む。読めない部分は静かに諦める（人数だけは必ず取る）。

    実測の並び:
        【総 合 計】人数 36 759,410 690,619 51,660 1,282,830 0 0     ← 事業主
        子ども・子育て拠出金 50,472 14,020 759,410 690,611 …        ← 末尾6個が個人分
        ( 0) 17,131                                                  ← 子育て支援金 事業主
        17,139                                                       ← 子育て支援金 個人分
    """
    stmt.total_count = int(m.group(1))
    er = _numbers(m.group(2))
    keys = ("kenpo_total", "kenpo", "kaigo", "konen")
    for k, v in zip(keys, er):
        stmt.totals[f"{k}_er"] = v

    for j in range(idx + 1, min(idx + 5, len(lines))):
        km = _KODOMO_TOTAL_RE.match(lines[j])
        if not km:
            continue
        ee = _numbers(km.group(1))
        if len(ee) >= len(er):
            for k, v in zip(keys, ee[-len(er):]):
                stmt.totals[f"{k}_ee"] = v
        # 続く2行が子ども・子育て支援金の事業主・個人分
        tm = _TOKUTEI_RE.match(lines[j + 1]) if j + 1 < len(lines) else None
        if tm:
            nums = _numbers(tm.group(2))
            if nums:
                stmt.totals["kodomo_er"] = nums[0]
            nxt = lines[j + 2].strip() if j + 2 < len(lines) else ""
            if _is_num(nxt):
                stmt.totals["kodomo_ee"] = _to_int(nxt)
        break


def verify_totals(stmt: PdfStatement) -> list:
    """【総合計】と読み取った明細を突き合わせる。合わなければ ShahoPdfError。

    ここを通れば「1行も落としていない・金額を読み違えていない」と言える。
    合わない紙をそのまま投入計画に流すと、静かに数名が欠けたまま処理が進む。
    """
    problems = []
    if stmt.total_count is not None and stmt.total_count != len(stmt.persons):
        problems.append(
            f"人数が合いません（【総合計】{stmt.total_count}名 / 読み取り{len(stmt.persons)}名）")

    for key, total in sorted(stmt.totals.items()):
        got = sum(p.premiums.get(key, 0) for p in stmt.persons)
        if got != total:
            problems.append(f"{key} の合計が合いません（帳票 {total:,} / 読み取り {got:,}）")

    if problems:
        raise ShahoPdfError(
            "PDFのチェックサムが合いません。読み取りを信用できないので中止します:\n- "
            + "\n- ".join(problems))
    checked = ["人数"] if stmt.total_count is not None else []
    checked += sorted(stmt.totals)
    return checked


def verify_person_premiums(stmt: PdfStatement, master=None, *, rounding: str = "50sen",
                           tolerance: int = 1) -> dict:
    """1名ごとの保険料を検算する。{社員番号: [問題文, ...]}（問題なしの人は入らない）。

    2段構え:
      1. **PDF内部の整合**（master 不要）: 健保計 = 本体 + 介護 + 子育て支援金
      2. **料率検算**（master 指定時）: 事業主+個人分 = 標準報酬 × 全体料率（端数が消えるので厳密）

    値は直さない。原票が正で、ここは**読み違い・料率変更を見つける網**として置く
    （健診PDFの BMI 検算と同じ位置づけ）。
    """
    from services.shaho_check import round_premium

    out: dict[str, list] = {}
    for p in stmt.persons:
        problems = []
        pr = p.premiums
        if pr:
            for side, label in (("er", "事業主"), ("ee", "個人分")):
                parts = (pr.get(f"kenpo_{side}", 0) + pr.get(f"kaigo_{side}", 0)
                         + pr.get(f"kodomo_{side}", 0))
                total = pr.get(f"kenpo_total_{side}", 0)
                if total and parts != total:
                    problems.append(
                        f"健康保険計({label})が内訳と合いません"
                        f"（計 {total:,} ≠ 本体+介護+子育て {parts:,}）")

        if master is not None and p.kenpo_smr:
            kaigo_on = pr.get("kaigo_er", 0) or pr.get("kaigo_ee", 0)
            expect = [
                ("kenpo", p.kenpo_smr, master.rates["kenpo"], "健康保険"),
                ("kodomo", p.kenpo_smr, master.rates["kodomo"], "子ども・子育て支援金"),
            ]
            if kaigo_on:
                expect.append(("kaigo", p.kenpo_smr, master.rates["kaigo"], "介護保険"))
            if p.konen_smr:
                expect.append(("konen", p.konen_smr, master.rates["konen"], "厚生年金"))
            for key, smr, rate, label in expect:
                both = pr.get(f"{key}_er", 0) + pr.get(f"{key}_ee", 0)
                want = round_premium(smr * rate.total, rounding)
                if abs(both - want) > tolerance:
                    problems.append(
                        f"{label}が標準報酬{smr:,}円×{rate.total:.4%}と合いません"
                        f"（帳票 {both:,} / 計算 {want:,}）")
                ee = pr.get(f"{key}_ee", 0)
                want_ee = round_premium(smr * rate.employee, rounding)
                if abs(ee - want_ee) > tolerance:
                    problems.append(
                        f"{label}の個人分が計算と合いません（帳票 {ee:,} / 計算 {want_ee:,}）")

        if problems:
            out[p.emp] = problems
    return out


def read_pdf(path: str, *, expected_office: str = "") -> PdfStatement:
    """PDFファイルを読んで PdfStatement を返す（このモジュール唯一のI/O）。"""
    import pdfplumber

    try:
        with pdfplumber.open(path) as pdf:
            pages = [page.extract_text() or "" for page in pdf.pages]
    except ShahoPdfError:
        raise
    except Exception as e:
        raise ShahoPdfError(f"PDFを開けませんでした: {e}") from e

    text = "\n".join(pages)
    if not text.strip():
        raise ShahoPdfError(
            "PDFにテキストがありません（画像として取り込まれたPDFの可能性）。"
            "社労士に、印刷ではなくPDF出力で送ってもらってください")
    return parse_text(text, expected_office=expected_office)
