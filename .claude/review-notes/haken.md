# 派遣台帳モード（四半期作成〜jinjer 添付）— 2026-09-10 転記・同日コード検証済

## 入口
- `/haken_build` `/haken_verify` `/haken_download` `/haken_export_pdf` `/haken_pdf_status/<job_id>` `/haken_freshness` `/haken_quarter_status` `/haken_attach_preview` `/haken_attach_execute` `/haken_attach_cancel` `/haken_attach_status`
- services: `services/daicho`（台帳の正はここ。20 モジュールのパッケージ）

## 入力
- **`Z:\派遣元管理台帳\input` 配下のファイル群**（`services/daicho/config.py`、`inputs.py`）：e-staffing 契約 CSV `TCnmht*.csv` / `CPInmht*.csv`、従業員一覧 `従業員一覧*.xlsx`、SAP Fieldglass 3 種（`*WorkOrder*.csv` / `*fieldglass_details*.json` / `*ユニアデックス*業務内容*.xlsx`）、エリクソン `派遣元管理台帳作成用*.csv`、直接契約マスタ 2 本（`direct.py`）。
- jinjer API は「退職者込みの人マスタをキャッシュして roster に合流する」用途だけ。契約情報は jinjer 由来ではない。
- **ファイル選択はワイルドカード＋`newest()`（最新 1 本）。** 四半期をまたいで古い WorkOrder や業務内容 xlsx を掴む余地があるので、重点確認 4 はここを見る。

## 出力と不可逆な処理
- build の成果物は 3 ファイル（台帳 xlsx／一覧 CSV／警告 CSV）を上書き生成。ダウンロードもこの 3 種限定。警告 CSV は前回比の突合元でもある。
- **jinjer への添付（`/haken_attach_execute`）** は不可逆（`services/daicho/jinjer_attach.py` の POST）。ガードは三重：許可 CSV フェイルクローズ、確認欄「添付」、実行中・予約中・他 PC ロックは 409。添付はデタッチした子プロセス（exe は `launcher.py --daicho-attach`、python は `daicho_attach_run.py`）で走り、キャンセルは `attach_job.request_cancel`。スレッド実行へのフォールバックは `HAKEN_ATTACH_FALLBACK_THREAD=1` のときだけ。

## 業務ルール
- `fg_mode` は auto / legacy / report の 3 値（`build.py` で検証）。ただし CLI 専用で、画面は `no_fg` だけを送り `fg_mode` は auto 固定。

## 既知の落とし穴
- `fg_mode` の 3 分岐を直接検証するテストが無い（2026-09-10 時点）。

## 宿題
- 2026-08-28 全フェーズ実装完了。運用実績の確認。特に `newest()` が四半期をまたぐ古いファイルを掴んでいないかの実データ確認。
