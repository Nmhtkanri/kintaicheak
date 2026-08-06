"""SAP経費CSV の重複除外フィルタ【2026-08-06 廃止・現在どこからも使っていない】

⚠️ このモジュールは使わないでください。後継は services/sap_import_ledger.py です。

廃止した理由: 「前月の生CSVを過去データに指定して突合する」方式そのものが、実際に取り込んだ
ファイルがどれだったかを人間が思い出す前提になっており、2026-08-06 に破綻したため。
7月に取り込んだのは (4)_重複除外済.csv だったが、下記 GENERATED_MARKERS の仕組みが
名前に「重複除外済」を含むCSVを過去データから自動スキップするため選べず、SAP元データ
フォルダにあった (5).csv（7月には使わなかった別ダウンロード・22ID欠落）を指定した結果、
除外0行となり 285,642円（13名/93行）が二重支給になりかけた。

現在は「取り込んだ費用シートを台帳に記録し、翌月は台帳と突合する」方式に移行済み。
残してあるのは経緯の記録と、万一の切り戻し用。新しいコードから import しないこと。

--- 以下、旧仕様の説明 ---


SAP Fieldglass の月次経費CSVには前々月分の費用シートが再掲されてくることがあり、
そのまま統合一覧表へ取り込むと二重計上になる。当月CSVを過去CSVと「費用シート ID」
完全一致で突合し、一致行を取込前に除外する。

ロジックは単体ツール `Y:\\SAP経費重複の除外システム\\sap_expense_duplicate_excluder.py`
（2026-07 谷津作・運用実績あり）からの移植。判定仕様は同一:
  - 「費用シート ID」の前後空白を除いた完全一致のみで判定
  - ID が空欄の行は判定に使わず、除外もしない
  - 1つの費用シートに複数明細がある場合、そのIDの当月行はすべて除外
  - 過去フォルダ指定時、生成物（重複除外済/除外一覧 CSV）は自動で読み飛ばす

出力は元CSVを変更せず、除外済みCSV・除外一覧CSVを output_dir へ書き出す。
"""

from __future__ import annotations

import csv
import datetime as dt
import logging
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

logger = logging.getLogger(__name__)

ID_COLUMN = "費用シート ID"
# 過去フォルダ走査時に読み飛ばす生成物マーカー（単体ツール出力＋当モジュール出力の両方）
GENERATED_MARKERS = ("重複除外済", "除外一覧", "SAP重複除外済", "SAP除外一覧")
REMOVED_EXTRA_COLUMNS = ["除外理由", "照合キー", "重複元ファイル", "重複元CSV行", "重複元件数"]


@dataclass(frozen=True)
class _CsvData:
    path: Path
    encoding: str
    fieldnames: list[str]
    rows: list[dict[str, str]]


@dataclass(frozen=True)
class _MatchSource:
    path: Path
    row_number: int


@dataclass
class SapDedupResult:
    current_path: Path
    clean_path: Path
    removed_path: Path
    past_file_count: int
    past_row_count: int
    current_row_count: int
    removed_row_count: int
    kept_row_count: int
    remaining_match_count: int          # 検算: 除外済みCSVに残った一致行数（0 が正常）
    fieldnames: list[str] = field(default_factory=list)
    removed_fieldnames: list[str] = field(default_factory=list)
    removed_rows: list[dict[str, str]] = field(default_factory=list)   # Excel シート化用

    def summary_dict(self) -> dict:
        """app.py の stats 用に JSON 化しやすい形へ落とす。"""
        return {
            "past_files": self.past_file_count,
            "past_rows": self.past_row_count,
            "current_rows": self.current_row_count,
            "removed_rows": self.removed_row_count,
            "kept_rows": self.kept_row_count,
            "remaining_match": self.remaining_match_count,
            "clean_csv_name": self.clean_path.name,
            "removed_csv_name": self.removed_path.name,
        }


def _detect_encoding(path: Path) -> str:
    data = path.read_bytes()
    if data.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    try:
        data.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError:
        return "cp932"


def _read_csv(path: Path) -> _CsvData:
    encoding = _detect_encoding(path)
    with path.open("r", encoding=encoding, newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError(f"CSVヘッダーがありません: {path}")
        fieldnames = [name.lstrip("\ufeff") for name in reader.fieldnames]
        rows: list[dict[str, str]] = []
        for row in reader:
            normalized = {}
            for raw_name, value in row.items():
                if raw_name is None:
                    continue
                normalized[raw_name.lstrip("\ufeff")] = value if value is not None else ""
            rows.append(normalized)
    return _CsvData(path=path, encoding=encoding, fieldnames=fieldnames, rows=rows)


def _require_columns(data: _CsvData, columns: Iterable[str]) -> None:
    missing = [column for column in columns if column not in data.fieldnames]
    if missing:
        raise ValueError(f"必要な列がありません: {', '.join(missing)} / {data.path}")


def _make_key(row: dict[str, str], id_column: str) -> str:
    """費用シート ID を正規化した照合キー。空IDは "" を返し、判定に使わない。"""
    return str(row.get(id_column, "")).strip()


def collect_past_files(
    inputs: Iterable[str | Path],
    current_path: Path,
    recursive: bool,
) -> list[Path]:
    """過去CSV（ファイル/フォルダ混在可）を集める。生成物CSVと当月CSV自身は除外する。"""
    current_resolved = current_path.resolve()
    files: list[Path] = []

    for raw in inputs:
        path = Path(raw).expanduser()
        if not path.exists():
            raise FileNotFoundError(f"過去SAPデータが見つかりません: {path}")
        if path.is_dir():
            pattern = "**/*.csv" if recursive else "*.csv"
            candidates = path.glob(pattern)
        else:
            candidates = [path]

        for candidate in candidates:
            if not candidate.is_file():
                continue
            if candidate.suffix.lower() != ".csv":
                continue
            if any(marker in candidate.stem for marker in GENERATED_MARKERS):
                continue
            try:
                if candidate.resolve() == current_resolved:
                    continue
            except OSError:
                pass
            files.append(candidate)

    seen: set[str] = set()
    unique_files: list[Path] = []
    for file_path in files:
        key = os.path.normcase(str(file_path.resolve()))
        if key in seen:
            continue
        seen.add(key)
        unique_files.append(file_path)

    if not unique_files:
        raise ValueError("過去SAP CSVが1つも見つかりませんでした（フォルダ指定の場合は中のCSVをご確認ください）。")
    return unique_files


def _unique_output_paths(stem: str, output_dir: Path) -> tuple[Path, Path]:
    clean = output_dir / f"{stem}_SAP重複除外済.csv"
    removed = output_dir / f"{stem}_SAP除外一覧.csv"
    if not clean.exists() and not removed.exists():
        return clean, removed
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return (
        output_dir / f"{stem}_SAP重複除外済_{stamp}.csv",
        output_dir / f"{stem}_SAP除外一覧_{stamp}.csv",
    )


def _write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]], encoding: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding=encoding, newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore", lineterminator="\r\n")
        writer.writeheader()
        writer.writerows(rows)


def run_sap_dedup(
    past_inputs: Iterable[str | Path],
    current_csv: str | Path,
    output_dir: str | Path,
    output_stem: str | None = None,
    recursive: bool = True,
    id_column: str = ID_COLUMN,
) -> SapDedupResult:
    """当月SAP CSVから過去と費用シートIDが一致する行を除外し、除外済みCSVを出力する。

    Args:
        past_inputs: 過去CSVのファイル/フォルダのリスト（フォルダは recursive に従い走査）
        current_csv: 当月の SAP 生CSV
        output_dir: 除外済みCSV・除外一覧CSVの出力先
        output_stem: 出力ファイル名の頭（未指定なら当月CSVの stem）

    Returns:
        SapDedupResult（clean_path を後段の取込に使う。removed_rows はシート化用）
    """
    current_path = Path(current_csv).expanduser()
    if not current_path.exists():
        raise FileNotFoundError(f"当月SAP CSVが見つかりません: {current_path}")

    output_dir = Path(output_dir).expanduser()
    past_files = collect_past_files(past_inputs, current_path, recursive)

    current_data = _read_csv(current_path)
    _require_columns(current_data, (id_column,))

    key_sources: dict[str, list[_MatchSource]] = {}
    past_row_count = 0
    for past_file in past_files:
        past_data = _read_csv(past_file)
        _require_columns(past_data, (id_column,))
        for index, row in enumerate(past_data.rows, start=2):
            key = _make_key(row, id_column)
            if not key:
                continue
            key_sources.setdefault(key, []).append(_MatchSource(past_file, index))
        past_row_count += len(past_data.rows)

    kept_rows: list[dict[str, str]] = []
    removed_rows: list[dict[str, str]] = []
    for row in current_data.rows:
        key = _make_key(row, id_column)
        sources = key_sources.get(key, []) if key else []
        if sources:
            first = sources[0]
            removed_row = dict(row)
            removed_row["除外理由"] = f"過去SAP経費と{id_column}が一致"
            removed_row["照合キー"] = key
            removed_row["重複元ファイル"] = str(first.path)
            removed_row["重複元CSV行"] = str(first.row_number)
            removed_row["重複元件数"] = str(len(sources))
            removed_rows.append(removed_row)
        else:
            kept_rows.append(row)

    stem = output_stem or current_path.stem
    clean_path, removed_path = _unique_output_paths(stem, output_dir)
    removed_fieldnames = current_data.fieldnames + REMOVED_EXTRA_COLUMNS

    _write_csv(clean_path, current_data.fieldnames, kept_rows, current_data.encoding)
    _write_csv(removed_path, removed_fieldnames, removed_rows, current_data.encoding)

    # 検算: 除外済みCSVを読み直して一致が残っていないことを確認（単体ツールと同じ）
    verification = _read_csv(clean_path)
    remaining_match_count = sum(
        1 for row in verification.rows if _make_key(row, id_column) in key_sources
    )
    if remaining_match_count:
        raise ValueError(
            f"SAP重複除外の検算に失敗しました（除外済みCSVに一致が {remaining_match_count} 行残存）: {clean_path}")

    logger.info(
        "SAP重複除外: 過去 %d ファイル/%d 行と突合 → 当月 %d 行から %d 行除外（残り %d 行）",
        len(past_files), past_row_count,
        len(current_data.rows), len(removed_rows), len(kept_rows),
    )
    return SapDedupResult(
        current_path=current_path,
        clean_path=clean_path,
        removed_path=removed_path,
        past_file_count=len(past_files),
        past_row_count=past_row_count,
        current_row_count=len(current_data.rows),
        removed_row_count=len(removed_rows),
        kept_row_count=len(kept_rows),
        remaining_match_count=remaining_match_count,
        fieldnames=list(current_data.fieldnames),
        removed_fieldnames=removed_fieldnames,
        removed_rows=removed_rows,
    )


def parse_past_inputs(raw: str) -> list[str]:
    """画面のテキスト入力（; または改行区切りの複数パス）→ パスリスト。"""
    if not raw:
        return []
    parts: list[str] = []
    for chunk in str(raw).replace("\r", "\n").split("\n"):
        for p in chunk.split(";"):
            p = p.strip().strip('"')
            if p:
                parts.append(p)
    return parts
