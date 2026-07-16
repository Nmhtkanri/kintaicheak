"""jinjer給与へのインポート行生成とAPI投入（経費マクロ移植 P2）

集計結果（keihi_classify の bucket）から jinjer給与インポートCSV（11列・谷津さん確定ヘッダー）を作り、
`POST /v1/jinji-imports`（メニュー種別=給与計算、テンプレート「経費インポート用」id=37047）で投入する。

列マッピング（live 集計シート実測 2026-07-16。マクロは dRY列振り分け→Module2 の2段変換）:
  - 夜間当番手当       ← F列相当 = 夜間当番(D) + RINK(E)   ※dRYがF値をR/Sペアに入れるため合算値になる
  - 定常外業務対応手当  ← 仕訳データ振り分け由来（本移植の対象外）→ 当面 0
  - 支給過不足調整     ← 固定 0（旧「過不足調整」を名称変更）
  - 非課税通勤費       ← 交通費(H)
  - 立替金（顧客請求分）← 顧客請求分(G)
  - 立替金            ← 非課税精算(立替金)(I)
  - その他            ← その他(J)
  - その他手当         ← R/S・T/Uペア由来の別支給項目 → 当面 0
  - 現物支給          ← 固定 0
  ※テレワーク手当(K) は現行マクロでもインポートCSVに乗らない（dRYでK廃止）。金額があれば警告表示する。

投入は人間の最終チェック（プレビューHTML＋CSVダウンロード）を経てから行う（確認後実行）。
"""

from __future__ import annotations

import time as _time
from dataclasses import dataclass, field
from pathlib import Path

from services.keihi_classify import B_YAKAN, B_RINK, B_TRANS, B_ETC, B_TW, B_BILL, B_NONTAX

# 谷津さん確定の正式ヘッダー（2026-07-16。jinjerテンプレート「経費インポート用」と一致させること）
IMPORT_HEADERS = [
    "社員番号", "氏名", "夜間当番手当", "定常外業務対応手当", "支給過不足調整",
    "非課税通勤費", "立替金（顧客請求分）", "立替金", "その他", "その他手当", "現物支給",
]
# jinjer側テンプレート（GET /v1/master/jinji-import-templates で実測）
DEFAULT_TEMPLATE_ID = "37047"   # 「経費インポート用」menu=payroll_calculation


def build_import_rows(by_id: dict, emp_names: dict, roster_names: "dict | None" = None) -> tuple[list[dict], list[str]]:
    """集計（数字ID→7バケット）からインポート行と警告リストを返す。

    行は 20YY 始まりの社員のみ・社員番号昇順（5/6/9始まりは給与計算対象外）。
    """
    from services.keihi_summary import in_company_scope
    roster_names = roster_names or {}
    rows: list[dict] = []
    warnings: list[str] = []
    for emp_id in sorted(by_id):
        if not in_company_scope(emp_id):
            continue
        v = by_id[emp_id]
        name = emp_names.get(emp_id) or roster_names.get(emp_id, "")
        tw = int(v[B_TW])
        if tw:
            warnings.append(
                f"{emp_id} {name}: テレワーク手当 {tw:,}円 はインポートCSVに含まれません（現行マクロ仕様）")
        rows.append({
            "社員番号": emp_id,
            "氏名": name,
            "夜間当番手当": int(v[B_YAKAN] + v[B_RINK]),   # F列相当（夜間＋RINK）
            "定常外業務対応手当": 0,
            "支給過不足調整": 0,
            "非課税通勤費": int(v[B_TRANS]),
            "立替金（顧客請求分）": int(v[B_BILL]),
            "立替金": int(v[B_NONTAX]),
            "その他": int(v[B_ETC]),
            "その他手当": 0,
            "現物支給": 0,
        })
    return rows, warnings


def _escape_csv(s: str) -> str:
    """Module2.EscapeCSV: 文字列は必ずダブルクォート囲み・内部の " は "" に。"""
    return '"' + str(s).replace('"', '""') + '"'


def render_import_csv(rows: list[dict]) -> bytes:
    """インポートCSVを Shift-JIS バイト列で返す（Module2 と同形式: 文字列のみクォート）。"""
    lines = [",".join(IMPORT_HEADERS)]
    for r in rows:
        cells = [_escape_csv(r["社員番号"]), _escape_csv(r["氏名"])]
        for h in IMPORT_HEADERS[2:]:
            cells.append(str(int(r[h])))
        lines.append(",".join(cells))
    return ("\r\n".join(lines) + "\r\n").encode("cp932")


def write_import_csv(rows: list[dict], out_path: "str | Path") -> Path:
    """インポートCSVをファイルに書き出す（人間チェック用ダウンロード）。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(render_import_csv(rows))
    return out_path


# ----------------------------------------------------------------------
# jinjer API 投入（確認後実行）
# ----------------------------------------------------------------------

@dataclass
class PayrollImportResult:
    ok: bool
    month: str = ""
    file_name: str = ""
    status: str = ""      # "0"=予約 / "1"=成功 / "2"=失敗 / "" = 不明
    error: str = ""
    logs: list = field(default_factory=list)


def post_payroll_import(
    client,
    month: str,
    csv_bytes: bytes,
    file_name: str,
    template_id: str = DEFAULT_TEMPLATE_ID,
    apply_formulas_off: bool = True,
    poll_seconds: int = 180,
    log_func=print,
) -> PayrollImportResult:
    """給与計算インポートを `POST /v1/jinji-imports` で投入し、完了をポーリングする。

    Args:
        client: JinjerClient（認証済み）
        month: 処理月 "YYYY-MM"（salary_setting.executed_on）
        csv_bytes: Shift-JIS のインポートCSV
        file_name: インポートファイル名
        template_id: 入力テンプレートID（既定 37047「経費インポート用」）
        apply_formulas_off: 「編集した項目は計算式を適用しない」を有効にするか
        poll_seconds: 完了待ちの最大秒数（バッチ処理のため時間がかかる）
    """
    result = PayrollImportResult(ok=False, month=month, file_name=file_name)

    def log(msg):
        result.logs.append(msg)
        log_func(msg)

    try:
        data = client.post_jinji_import(
            csv_bytes=csv_bytes, file_name=file_name, template_id=template_id,
            executed_on=month, apply_formulas_off=apply_formulas_off,
        )
        log(f"[info] jinji-imports 投入予約 OK: {file_name} (処理月 {month} / テンプレ {template_id})")
    except Exception as e:  # noqa: BLE001
        result.error = f"投入予約に失敗しました: {e}"
        log(f"[error] {result.error}")
        return result

    # ポーリング（status: 0=予約 / 1=成功 / 2=失敗）
    waited = 0
    interval = 10
    while waited < poll_seconds:
        _time.sleep(interval)
        waited += interval
        try:
            item = client.find_jinji_import(file_name)
        except Exception as e:  # noqa: BLE001
            log(f"[warn] 状況確認に失敗（リトライ）: {e}")
            continue
        status = str((item or {}).get("status") or "")
        if status == "1":
            result.status = "1"
            result.ok = True
            log(f"[done] インポート成功（{waited}秒）。jinjer給与画面で金額をご確認ください。")
            return result
        if status == "2":
            result.status = "2"
            result.error = "jinjer側でインポートが失敗しました。通知メールとjinjer画面のインポート履歴をご確認ください。"
            log(f"[error] {result.error}")
            return result
        log(f"[info] 処理待ち... ({waited}秒 / status={status or '未検出'})")

    result.status = result.status or "0"
    result.ok = True  # 予約自体は成功（完了待ちタイムアウトは失敗扱いにしない）
    result.error = ""
    log("[warn] 完了確認がタイムアウトしました（投入予約は成功）。後ほどjinjer画面のインポート履歴か通知メールをご確認ください。")
    return result
