# 請求書モード（常駐先 PDF → freee 売上 CSV・提出用 PDF）— 2026-09-10 転記・同日コード検証済

## 入口
- `/invoice_preview` `/invoice_export` `/invoice_folders` `/invoice_folders_check` `/invoice_folders_save` `/invoice_pdf_companies` `/invoice_pdf_run` `/invoice_download/<ym>/<path:filename>`
- services: `invoice_mode` `invoice_folders` `invoice_pdf`

## 入力
- 常駐先ごとの PDF。見るフォルダは設定 CSV（画面 3 タブから編集、書き込みは許可ユーザー CSV の 3 名限定、判定はフェイルクローズ）。
- 外部マスタ一式（`config.py`）：対象フォルダ（対象=1 が 28 件）、年度営業実績ブック、freee 補正マスタ、取引先マスタ、対象外リスト、内訳マスタ、PDF 作成設定、書込許可ユーザー。

## 出力と不可逆な処理
- freee 売上 CSV。**freee は H 列を税込として読む。** 外税＋税抜で出すと二重に引かれる（2026-08 に税抜 810,000 が 729,000 で計上された実害あり。2026-09-02 に税込へ戻した）。
- 提出用 PDF（請求書 Excel＋勤怠 PDF。2026-08-20 実装。exe 再ビルドに pypdf が必要）。**対象者は設定 CSV `請求書モード_PDF作成設定.csv` が正で、コードに人数の直書きは無い。** 2026-09-10 時点の CSV は 2 社 3 名（アクシス＝細川さん／NIC ソフト＝熊崎さん・岡崎さん）。実装時は 2 社 4 名だったので、アクシスの大村さんの行が意図的に外れたのかは要確認。
- 見るフォルダ設定の保存（`/invoice_folders_save`）。常駐先フォルダ改名時はパス CSV 2 本（`請求書モード_PDF作成設定.csv` と `請求書モード_対象フォルダ.csv`）とも直す。

## 業務ルール
- 氏名は jinjer の姓・名から「姓 名」（半角スペース）。

## 宿題
- PDF 作成設定 CSV の大村さんの行の有無（意図的か欠落か）。
- 対象者を増やすときの PDF 構成。
