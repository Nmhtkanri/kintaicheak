"""請求書Excelの対象月シートと勤怠PDFを綴じて、提出用PDFを作るサービス。

いまは人が手でやっている次の作業を置き換える。

  ① 請求書Excel（年度ファイル）の対象月シートに金額を入れる  ← 人のまま
  ② そのシートをPDF化し、勤怠PDF（業務実績表・タイムシート等）と結合する
  ③ 提出データフォルダへ決まった名前で置く

②③をここで行う。常駐先ごとにフォルダ構成もファイル名もバラバラなので、
推測はせず外部CSV（1人1行）に書いてもらった場所だけを見る。違う人の勤怠を
綴じると事故になるため、少しでも決まらなければ作らずに止める。

外部サービスへは書き込まない。共有フォルダに新しいPDFを作るだけで、
提出は人が行う。既にあるファイルは上書きしない。
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence


class InvoicePdfError(ValueError):
    """設定の不備や、綴じる材料が決まらないときに投げる。"""


@dataclass
class BuildPlan:
    """1人分の「何を綴じてどこへ出すか」。"""

    partner: str
    person: str
    workbook: Path
    sheet: str
    attendance: list[Path] = field(default_factory=list)
    output: Path | None = None
    errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors


def fiscal_year(target: date) -> int:
    """年度。4月始まりなので1〜3月は前年になる。"""
    return target.year if target.month >= 4 else target.year - 1


def expand(pattern: str, target: date) -> str:
    """{YYYY}/{YY}/{MM}/{M}/{FY} を対象月で埋める。

    {MM} は2桁ゼロ埋め、{M} は前ゼロなし（「26年7月」のようなシート名用）。
    """
    return (str(pattern)
            .replace("{YYYY}", f"{target.year:04d}")
            .replace("{YY}", f"{target.year % 100:02d}")
            .replace("{MM}", f"{target.month:02d}")
            .replace("{M}", str(target.month))
            .replace("{FY}", f"{fiscal_year(target):04d}"))


def _match_files(folder: Path, pattern: str) -> list[Path]:
    """フォルダ直下から glob で探す。見つかった順は名前順に固定する。"""
    if not folder.exists():
        return []
    return sorted(p for p in folder.glob(pattern) if p.is_file())


def load_settings(csv_path: os.PathLike[str] | str) -> list[dict[str, str]]:
    """1人1行の設定を読む。

    列:
        取引先, 氏名, 請求書Excel, シート名, 勤怠フォルダ, 勤怠ファイル, 出力フォルダ, 出力ファイル名
    値には {YYYY} {YY} {MM} {M} {FY} が使える。勤怠ファイルは glob（* が使える）。
    """
    import csv
    import io

    path = Path(csv_path)
    if not path.exists():
        raise InvoicePdfError(f"設定CSVがありません: {path}")
    for encoding in ("utf-8-sig", "cp932"):
        try:
            rows = list(csv.DictReader(io.StringIO(path.read_text(encoding=encoding))))
            break
        except UnicodeDecodeError:
            continue
    else:
        raise InvoicePdfError(f"設定CSVの文字コードが読めません: {path}")
    return [r for r in rows if (r.get("氏名") or "").strip()]


def plan(month: str, settings: Sequence[dict[str, str]]) -> list[BuildPlan]:
    """対象月について、誰の何を綴じるかを決める。ファイルはまだ作らない。"""
    matched = re.fullmatch(r"(\d{4})-(\d{2})", str(month).strip())
    if not matched:
        raise InvoicePdfError("対象月は YYYY-MM 形式で指定してください")
    target = date(int(matched.group(1)), int(matched.group(2)), 1)

    plans: list[BuildPlan] = []
    for row in settings:
        workbook = Path(expand(row.get("請求書Excel") or "", target))
        item = BuildPlan(
            partner=(row.get("取引先") or "").strip(),
            person=(row.get("氏名") or "").strip(),
            workbook=workbook,
            sheet=expand(row.get("シート名") or "", target),
        )
        if not workbook.name or not workbook.exists():
            item.errors.append(f"請求書Excelが見つかりません: {workbook}")

        folder = Path(expand(row.get("勤怠フォルダ") or "", target))
        pattern = expand(row.get("勤怠ファイル") or "", target)
        found = _match_files(folder, pattern) if pattern else []
        if not found:
            item.errors.append(f"勤怠PDFが見つかりません: {folder}\\{pattern}")
        elif len(found) > 1:
            # どれを綴じるか決まらない。人が消すか設定を絞るまで作らない。
            item.errors.append(
                "勤怠PDFの候補が複数あります（どれを綴じるか決められません）: "
                + " / ".join(p.name for p in found))
        else:
            item.attendance = found

        out_dir = Path(expand(row.get("出力フォルダ") or "", target))
        out_name = expand(row.get("出力ファイル名") or "", target)
        if not out_name:
            item.errors.append("出力ファイル名が未設定です")
        else:
            item.output = out_dir / out_name
            if not out_dir.exists():
                item.errors.append(f"出力フォルダがありません: {out_dir}")
            elif item.output.exists():
                item.errors.append(f"すでに同じ名前のPDFがあります（上書きしません）: {item.output.name}")
        plans.append(item)
    return plans


def export_sheet_to_pdf(workbook: Path, sheet: str, out_pdf: Path) -> Path:
    """Excelの1シートだけをPDFにする（Excel COM）。

    .xls も .xlsx も扱える。ブックは読み取り専用で開き、保存はしない。
    """
    import win32com.client as win32

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = None
    try:
        book = excel.Workbooks.Open(str(workbook), False, True)
        names = [s.Name for s in book.Worksheets]
        # シート名の末尾に空白が入っているブックがある（'26年7月 '）ので詰めて比べる
        hit = next((n for n in names if n.strip() == sheet.strip()), None)
        if hit is None:
            raise InvoicePdfError(
                f"シート「{sheet}」がありません: {workbook.name}（あるのは {', '.join(names)}）")
        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        # 対象シートだけを新しいブックへ複写してから出す。
        # ブックが複数シートを選択した状態で保存されていると、シートに対して
        # そのまま ExportAsFixedFormat すると選択中の他の月まで出てしまう
        # （大村さんのブックで6月分が2枚目に混ざった）。
        book.Worksheets(hit).Copy()
        single = excel.ActiveWorkbook
        try:
            # 印刷範囲が設定されていないブックがあり、そのままだと用紙幅に
            # 収まらず横にページが割れる（大村さんの請求書が2枚になった）。
            # 横は必ず1ページに収める。縦は内容なりに流す。
            setup = single.Worksheets(1).PageSetup
            setup.Zoom = False
            setup.FitToPagesWide = 1
            setup.FitToPagesTall = False
            single.ExportAsFixedFormat(0, str(out_pdf))
        finally:
            single.Close(False)
        return out_pdf
    finally:
        if book is not None:
            book.Close(False)
        excel.Quit()


def check_dates(pdf_path: Path, target: date) -> list[str]:
    """請求書PDFの請求日・入金期日が対象月とかみ合っているかを見る。

    前月シートをコピーして作る運用なので、金額だけ直して請求日・請求書番号・
    入金期日を戻し忘れることがある（2026-07 の大村さんのシートが実際に
    請求日 2026-06-30 / INV-260630KOM のまま残っていた）。綴じて提出した
    あとでは気づけないので、ここで拾って人に確認してもらう。

    日付の記載が無いテンプレートもあるので、見つからないものは何も言わない。

    Returns:
        気になる点の一覧。空なら問題なし。
    """
    import calendar

    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        text = " ".join((page.extract_text() or "") for page in pdf.pages)
    flat = re.sub(r"\s+", "", text)
    notes: list[str] = []

    issued = re.search(r"(?:請求日|請求年月日)[:：]?(\d{4})[-/年](\d{1,2})", flat)
    if issued:
        year, month = int(issued.group(1)), int(issued.group(2))
        if (year, month) != (target.year, target.month):
            notes.append(
                f"請求日が {year}-{month:02d} になっています"
                f"（対象は {target.year}-{target.month:02d}）")

    due = re.search(r"(?:入金期日|お?支払期日|支払期限)[:：]?(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})",
                    flat)
    if due:
        year, month, day = (int(due.group(1)), int(due.group(2)), int(due.group(3)))
        try:
            due_date = date(year, month, min(day, calendar.monthrange(year, month)[1]))
        except ValueError:
            due_date = None
        month_end = date(target.year, target.month,
                         calendar.monthrange(target.year, target.month)[1])
        if due_date and due_date <= month_end:
            notes.append(
                f"入金期日が {due_date.isoformat()} で、対象月（{month_end.isoformat()}）"
                "より後になっていません")

    if notes:
        notes.append("前月シートから請求日・請求書番号・入金期日を直し忘れていませんか")
    return notes


def merge_pdfs(parts: Iterable[os.PathLike[str] | str], out_pdf: Path) -> Path:
    """PDFを順番どおりに綴じる。"""
    from pypdf import PdfWriter

    writer = PdfWriter()
    used = 0
    for part in parts:
        writer.append(str(part))
        used += 1
    if not used:
        raise InvoicePdfError("綴じる材料がありません")
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    with out_pdf.open("wb") as f:
        writer.write(f)
    writer.close()
    return out_pdf


def build(month: str, settings: Sequence[dict[str, str]], *,
          work_dir: os.PathLike[str] | str | None = None,
          dry_run: bool = True, force: bool = False) -> dict[str, Any]:
    """対象月の提出用PDFを作る。

    dry_run でも請求書シートのPDF化までは行い、請求日・入金期日を検査する。
    そうしないと画面で確認する材料が出ないため。最終PDFの書き出しだけを止める。

    結果は3つに分ける。
      made          … 作った（dry_run のときは「作れる」）
      needs_confirm … 日付が対象月とかみ合わない。人が見て問題なければ force で作る
      skipped       … 材料が揃わない等。force でも作らない
    """
    import tempfile

    matched = re.fullmatch(r"(\d{4})-(\d{2})", str(month).strip())
    if not matched:
        raise InvoicePdfError("対象月は YYYY-MM 形式で指定してください")
    target = date(int(matched.group(1)), int(matched.group(2)), 1)

    plans = plan(month, settings)
    work = Path(work_dir) if work_dir else Path(tempfile.mkdtemp(prefix="invoice_pdf_"))
    made: list[dict[str, Any]] = []
    needs_confirm: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []

    for item in plans:
        base = {"取引先": item.partner, "氏名": item.person}
        if not item.ok:
            skipped.append({**base, "理由": " / ".join(item.errors)})
            continue
        detail = {
            **base,
            "請求書": f"{item.workbook.name}［{item.sheet}］",
            "勤怠": " / ".join(p.name for p in item.attendance),
            "出力先": str(item.output),
        }
        sheet_pdf = export_sheet_to_pdf(
            item.workbook, item.sheet, work / f"{item.person}_請求書.pdf")
        notes = check_dates(sheet_pdf, target)
        if notes and not force:
            needs_confirm.append({**detail, "確認事項": notes})
            continue
        detail["確認事項"] = notes          # force で通したときも記録に残す
        if dry_run:
            made.append(detail)
            continue
        merge_pdfs([sheet_pdf, *item.attendance], item.output)
        detail["ページ数"] = 1 + len(item.attendance)
        made.append(detail)

    return {"month": month, "dry_run": dry_run, "force": force,
            "made": made, "needs_confirm": needs_confirm, "skipped": skipped,
            "work_dir": str(work)}
