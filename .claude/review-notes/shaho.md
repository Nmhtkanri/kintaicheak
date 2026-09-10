# 標準報酬月額チェック・標報投入 — 2026-09-10 転記・同日コード検証済

## 入口
- `/shaho_run` `/shaho_download/<year>/<path:filename>` `/shaho_import_preview` `/shaho_import_execute` `/shaho_import_status` `/shaho_import_download/<ym>/<path:filename>`
- services: `shaho_engine` `shaho_check` `shaho_master` `shaho_pdf` `shaho_its` `shaho_report` `shaho_writer`

## 入力
- 社労士の PDF、関東 ITS の CSV、jinjer の報酬月額。

## 出力と不可逆な処理
- **jinjer の報酬月額（マスタ級）への投入。** ハブがマスタを書く唯一の例外。
  - 証番号と氏名の二重突合（`shaho_its.py`）。片方だけだと証番号 1151/1152 で別人が入れ替わる事故が実測にある。突合できない行は `UNRESOLVED` で投入させない。重複証番号は両方捨てる。この二重突合を外す差分は重大。
  - 投入のガード（`config.py`）：許可ユーザー CSV、実行台帳、ロックファイル、事業所コード固定 `SHAHO_IMPORT_EXPECTED_OFFICE=263`、書込間隔 25 秒。どれかを外す差分は重大。
- 検算ゲート（`shaho_engine.py` の `gate_diff`、±0.5 まで。NG なら計算値を出さない）。ゲートを緩める差分は重大。

## 業務ルール
- 検算ゲートの実績：2026-04〜06 の実データ全 713 人月で一致（`docs/PLAN_標準報酬月額チェック.md`）。
- 2026-08 の登録は 33 名中 32 名一致。残 1 名は jinjer 手入力ミスの検出例。人数は当時の実測で、コードに直書きは無い。

## 宿題
- なし（2026-08-17 本番稼働）。
