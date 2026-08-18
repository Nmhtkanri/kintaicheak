r"""関東ITS（健康保険組合）の「被保険者標準報酬決定通知書一覧」を読む。

社労士の保険料一覧表（`shaho_pdf.py`）と同じ投入計画に載せるため、
出力は `shaho_pdf.PdfStatement` / `PdfPerson` に揃えてある。

## 入力（2026-09 定時決定分で実測）

PATPOST で PDF から起こした **cp932 の8列CSV**。1名1行。

| 列 | 中身 | 使い方 |
|---|---|---|
| A | 連番（ページごとに振り直す） | 使わない |
| B | **被保険者証番号** | ★突合の主キー |
| C | 氏名 | ★突合の照合用 |
| D | 生年月日 | 使わない（OCRの欠けが多い） |
| E | 性別 | 使わない |
| F | 従前の標準報酬月額（**千円単位**） | 参考表示 |
| G | 平均額 | 参考表示 |
| H | **決定 標準報酬月額（円単位）** | ★投入する値 |

**F は千円・H は円**と単位が違う。H には PATPOST の変換で
`440 000` `1,330 000` のように空白やカンマが混ざるので、数字以外を落として読む。

## ★行落ち検知（A列のページ内連番）

この通知には社労士PDFの【総合計】のようなチェックサム行が無い。代わりに
**A列（ページ内連番）が 1,2,3…と連続しているか**で PATPOST の行落ちを検知する
（ページが変わると1に戻る）。連番が飛んだら・1から始まらなかったら・重複したら
**中止**する。あわせて「A列は数字なのに証番号や決定額が読めない行」も黙って捨てず
中止する（静かな取りこぼしを許さない）。

**限界: 各ページの最終行が丸ごと落ちた場合だけは検知できない**
（…18,19 → 次ページ1 と …18 → 次ページ1 を区別する材料が無い）。
そのため読み取り後に画面へ出る人数を、原本の最終ページの連番と目で照らすこと。

## ★これは健康保険の通知（厚生年金は載っていない）

関東ITSは健康保険組合なので、H列は**健康保険の**標準報酬月額。
実測で 179名中7名が厚年の上限65万円を超えている（最高139万円）。
そのため**厚生年金は等級表から導出する**（`master.find_grade(H).konen_smr`）。
上限にかかる人は自動的に650,000になる。179名すべてで
`find_grade(H).kenpo_smr == H` が成立することを確認済み＝H は正規の等級値。

## ★本人の特定は「証番号」と「氏名」の両方で行う（2026-08-17 実測）

どちらか片方だけでは事故る。実測（179名）:

| 判定 | 人数 | 扱い |
|---|---|---|
| 証番号・氏名とも一致 | 156 | 自動で投入対象 |
| **証番号と氏名が別人を指す** | 2 | **本人を特定できない → 投入させない** |
| 証番号は一致・氏名が表記ゆれ | 3 | 投入可。警告を出す |
| 氏名のみ一致（jinjerに証番号が未登録） | 17 | 投入可。警告を出す |
| どちらでも決まらない | 1 | 投入させない |

証番号だけで突合していたら、**証1151/1152 で柴田さんと岡村さんの標準報酬を
入れ替えて書き込むところだった**（CSV側の取り違えと判明。jinjerの登録が正）。
氏名だけなら Ver Martin（CSVはカタカナ）・宮嵜/宮寄・井料 香凜/香凛 を取り逃す。

jinjer側の証番号は `GET /v1/employees/social-and-labor-insurances` の
`social_insurance.health_insurance.number`。
"""

from __future__ import annotations

import csv
import re

from services.keiri_api import classify_employee
from services.shaho_pdf import PdfPerson, PdfStatement, ShahoPdfError

# この通知が持つのは定時決定だけ（随時改定は社労士の保険料一覧表から来る）
REASON = "定時決定"
COLUMNS = 8
COL_NO, COL_NAME, COL_PREV, COL_AVG, COL_DECIDED = 1, 2, 5, 6, 7


def _num(value):
    """`440 000` `1,330 000` `620 千円` → 数字だけ取り出す。"""
    digits = re.sub(r"[^\d]", "", str(value or ""))
    return int(digits) if digits else None


def name_key(name) -> str:
    """氏名の突合キー（空白の入り方だけを吸収する。異体字は吸収しない）。"""
    return re.sub(r"[\s　]", "", str(name or ""))


def verify_sequence(numbers: list) -> list:
    """A列（ページ内連番）の連続性を検証する。問題文のリストを返す（正常なら空）。

    正しい並びは「1,2,3,…（ページが変わると1に戻る）」。
    - 先頭が1でない → 最初の行が落ちている
    - 前の値+1 でも 1 でもない → 間の行が落ちた・ずれた・重複した
    """
    problems = []
    prev = None
    for n in numbers:
        if prev is None:
            if n != 1:
                problems.append(
                    f"最初の行のページ内連番が {n} です"
                    "（1から始まっていない＝先頭の行が落ちている可能性）")
        elif n != prev + 1 and n != 1:
            problems.append(
                f"ページ内連番が {prev} の次に {n} になっています"
                "（間の行が落ちているか、行が重複・ずれています）")
        prev = n
    return problems


def read_rows(path: str) -> list:
    """CSVから明細行だけを取り出す（A列が数字の行）。cp932→UTF-8 の順で試す。"""
    with open(path, "rb") as f:
        data = f.read()
    for encoding in ("cp932", "utf-8-sig", "utf-8"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ShahoPdfError("CSVの文字コードを判別できません（cp932/UTF-8 のどちらでもない）")

    out, seq, problems = [], [], []
    for row in csv.reader(text.splitlines()):
        if len(row) < COLUMNS or not row[0].strip().isdigit():
            continue
        seq.append(int(row[0].strip()))
        no = row[COL_NO].strip()
        decided = _num(row[COL_DECIDED])
        if not no or decided is None:
            # A列は明細行なのに中身が読めない＝変換の崩れ。黙って捨てると
            # その人の定時決定だけが静かに欠けるので、必ず止める。
            problems.append(
                f"ページ内連番 {row[0].strip()} の行"
                f"（{row[COL_NAME].strip() or '氏名不明'}）から"
                "証番号または決定標準報酬月額を読み取れません")
            continue
        out.append({
            "no": no,
            "name": row[COL_NAME].strip(),
            "prev": (_num(row[COL_PREV]) or 0) * 1000,   # F列は千円単位
            "avg": _num(row[COL_AVG]) or 0,
            "decided": decided,                          # H列は円単位
        })
    problems.extend(verify_sequence(seq))
    if problems:
        raise ShahoPdfError(
            "PATPOST変換に取りこぼしの疑いがあります。"
            "読み取りを信用できないので中止します:\n- " + "\n- ".join(problems)
            + "\n元PDFと突き合わせて、PATPOSTで変換し直すかCSVを直してください")
    if not out:
        raise ShahoPdfError(
            "決定標準報酬月額の行が1件も読み取れませんでした。"
            "関東ITSの「被保険者標準報酬決定通知書一覧」をPATPOSTでCSVにしたものを選んでください")
    return out


def build_statement(rows: list, target_ym: str, *, roster: dict, number_map: dict,
                    master) -> PdfStatement:
    """明細行 → 投入計画に載せられる PdfStatement（純関数）。

    Args:
        rows: `read_rows` の戻り値
        target_ym: 適用年月 "YYYY-MM"（CSVに書かれていないので画面から受け取る）
        roster: `{社員番号: {"name", ...}}`
        number_map: `{健保証番号: 社員番号}`（`load_number_map` で作る）
        master: 等級表（厚生年金の標準報酬を導出するのに使う）
    """
    from services.keiri_engine import ym_add

    by_name: dict = {}
    for emp, info in roster.items():
        by_name.setdefault(name_key(info.get("name")), []).append(emp)

    def pick_by_name(cands: list):
        """氏名の候補から1人に絞る。(社員番号 or None, 注記)。

        jinjerには同じ人が2つの社員番号で登録されていることがある
        （例: 谷津晴香さん = 2026007 と 3333003。旧登録が閉じられていない）。
        **投入できるのは 20YY 始まりの人だけ**なので、候補のうち 20YY が
        ちょうど1人ならその人に決めてよい。
        """
        if len(cands) == 1:
            return cands[0], ""
        targets = [c for c in cands if classify_employee(c) == "target"]
        if len(targets) == 1:
            others = "・".join(c for c in cands if c != targets[0])
            return targets[0], (f"jinjerに同じ氏名の登録が{len(cands)}件あり"
                                f"（{others}）、社員番号が 20YY の {targets[0]} を採りました")
        return None, ""

    stmt = PdfStatement(target_ym=target_ym, pay_ym=ym_add(target_ym, 1),
                        office_name="関東ITソフトウェア健康保険組合",
                        total_count=len(rows))
    seen: dict = {}
    for row in rows:
        by_no = number_map.get(row["no"])
        cands = by_name.get(name_key(row["name"]), [])
        by_nm, dup_note = pick_by_name(cands)

        person = PdfPerson(emp="", name=row["name"], kenpo_no=row["no"],
                           reason_kenpo=REASON, reason_konen=REASON)
        person.premiums = {"prev_kenpo": row["prev"], "average": row["avg"]}

        if by_no and by_nm and by_no != by_nm:
            person.issues.append(
                f"証番号は {by_no}（{roster[by_no]['name']}）・氏名は {by_nm}"
                f"（{roster[by_nm]['name']}）を指しており、本人を特定できません")
        elif by_no:
            person.emp = by_no
            jname = str((roster.get(by_no) or {}).get("name") or "")
            if name_key(jname) != name_key(row["name"]):
                person.warnings.append(
                    f"氏名が jinjer と違います（jinjer:{jname}）。証番号で突合しました")
        elif by_nm:
            person.emp = by_nm
            person.warnings.append(
                f"健保証番号 {row['no']} が jinjer に未登録のため、氏名で突合しました")
            if dup_note:
                person.warnings.append(dup_note)
        elif cands:
            person.issues.append(
                f"同じ氏名が {len(cands)}名（{'・'.join(cands)}）いて、"
                f"証番号 {row['no']} も jinjer に未登録のため特定できません")
        else:
            person.issues.append(
                f"証番号 {row['no']}・氏名『{row['name']}』のどちらでも "
                "jinjer の社員を特定できません")

        # 健保は通知の値そのまま。厚年は等級表から引く（上限65万で頭打ちになる）
        person.kenpo_smr = row["decided"]
        grade = master.find_grade(row["decided"]) if master is not None else None
        if grade is None:
            person.issues.append("等級表が読めないので厚生年金の標準報酬を出せません")
        elif grade.kenpo_smr != row["decided"]:
            person.issues.append(
                f"決定額 {row['decided']:,} 円が等級表の健保標準報酬と一致しません"
                f"（最も近い等級は {grade.kenpo_smr:,} 円）")
        else:
            person.konen_smr = grade.konen_smr
            if grade.konen_smr < row["decided"]:
                person.warnings.append(
                    f"厚生年金は上限のため {grade.konen_smr:,} 円"
                    f"（健保は {row['decided']:,} 円）")

        if person.emp:
            if person.emp in seen:
                person.issues.append(f"同じ社員番号が複数行にあります（{person.emp}）")
            seen[person.emp] = person.emp
        stmt.persons.append(person)
    return stmt


def load_number_map(client, employee_ids: list, batch: int = 50) -> dict:
    """jinjer の社会保険情報から `{健保証番号: 社員番号}` を作る。

    `GET /v1/employees/social-and-labor-insurances` の
    `social_insurance.health_insurance.number`。同じ番号が2人に付いていたら
    **どちらも採らない**（取り違えたまま書くより、特定できないとして人に返す）。
    """
    import requests

    from services.jinjer_api_client import JinjerAPIError

    out: dict = {}
    dupes = set()
    ids = [str(e).strip() for e in employee_ids if str(e or "").strip()]
    for i in range(0, len(ids), batch):
        response = requests.get(
            f"{client.base_url}/v1/employees/social-and-labor-insurances",
            headers=client._auth_headers(),
            params={"employee-ids": ",".join(ids[i:i + batch])}, timeout=60)
        if response.status_code != 200:
            raise JinjerAPIError(
                f"社会保険情報の取得に失敗 (status={response.status_code})")
        for item in response.json().get("data", []) or []:
            number = str((((item.get("social_insurance") or {})
                           .get("health_insurance") or {}).get("number") or "")).strip()
            emp = str(item.get("employee_id") or "").strip()
            if not number or not emp:
                continue
            if number in out and out[number] != emp:
                dupes.add(number)
            out[number] = emp
    for number in dupes:
        out.pop(number, None)
    return out
