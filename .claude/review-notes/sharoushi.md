# 社労士対応モード（前田事務所向け CSV）— 2026-09-10 転記・同日コード検証済

## 入口
- `/sharoushi_run` `/sharoushi_biko_save` `/sharoushi_download/<ym>/<path:filename>`
- services: `sharoushi_export`

## 入力
- **jinjer の給与明細 API 1 本だけ**（`fetch_statements`。キャッシュは経理モードと共用）。勤怠時間の列も勤怠 API ではなく給与明細内の `salary_attendance_items:kintai*` から取る。
- 外部マスタ（`config.py`）：列マッピング `SHAROUSHI_COLUMN_MAPPING_CSV`（直せば exe 再ビルドなしで次回実行から効く）、追加支給台帳 `SHAROUSHI_EXTRA_LEDGER_CSV`。
- 備考台帳 `Z:\API連携\docs\社労士モード_備考台帳.csv`（画面から保存）。

## 出力と不可逆な処理
- 社労士事務所へ渡す CSV（cp932、49 列 `LAYOUT_V2`）。2026-08-28 に 60 列→49 列（欠勤控除→社保調整）。旧 60 列は `LAYOUT_V1` として控えが残り、旧ファイルは `_旧60列` 接尾辞。
- 備考の保存（`/sharoushi_biko_save`）は共有側を書き換える。

## 業務ルール
- 列は番号でなく列 id で指す。旧スキーマ（列番号）のマッピング CSV は読まずエラー停止する設計。列番号の直書きがあれば軽微以上。
- 貸付金返済が社宅家賃の枠（deduction2）に入る月がある。コード上は社員番号 2023004 限定の特例として実装（`sharoushi_export.py` の決定事項 8 付近）。対象者が増えたら特例の書き方ごと見直す。

## 宿題
- 列定義の変更履歴が増えるので、旧列との対応表の保守。
- 2023004 限定の特例が他の人にも必要になったときの扱い。
