# スケジュール取込（シフト表 → jinjer 予定）— 2026-09-10 転記・同日コード検証済

## 入口
- シフト表の受け口は `/upload`（`mode=csv_export`。KDX 専用欄は `kdx_files`）→ 凡例レビュー → `/export_jinjer_csv`（jinjer インポート用 CSV 生成）。
- API 投入は `/schedule_api_import` `/schedule_start_edit` `/resolve_schedule_names`。
- services: `schedule_import_runner` `schedule_start_edit` `shift_resolver` `shift_legend_parser` `kdx_shift_parser` `bbs_shift_parser` `higashi_shift_parser` `ual_shift_parser` `multi_year_shift_parser` `jinjer_schedule_csv_exporter`

## 入力
- 常駐先ごとのシフト表（KDX・BBS・東・UAL など）。形式が常駐先ごとに違い、月で変わる。パーサは常駐先ごとに 1 本ずつ並列に存在する。

## 出力と不可逆な処理
- jinjer への予定投入（API）。**雛形 ID は使わず、出勤予定／退勤予定／休憩予定の時刻を直書きする方式**（`schedule_import_runner.py`、`schedule_start_edit.py`）。「スケジュール雛形ID」列は値 0（休み、`REST_TEMPLATE_ID`）だけ使う。勤怠側の「雛形 ID を空欄化する」対応（kintai.md）とは別物で、こちらで雛形 ID=0 を落とす修正は誤り。
- CSV インポート用グリッド（`jinjer_schedule_csv_exporter`）は ID 必須で、API 投入とはまた別。

## 業務ルール
- KDX：2026-09 から文字なし PDF。専用アップロード欄に入れると固定凡例（A→9:00-17:30 / C→16:30-34:00）を強制する（`kdx_shift_parser.py`）。
- BBS：2026-08-31 に構造化パーサ化。AI 読みは月初の空欄で 1 日ズレる。9 月は凡例欄なし。
- **夜勤明けは「休み」にする（意図した仕様）。** 退勤 30:00 以降のシフトは翌日を必ず明け休にする（`jinjer_schedule_csv_exporter.py`、回帰テスト `tests/test_auto_ake_and_alias.py`）。連日夜勤はシフト表優先、月末の夜勤は要確認扱い。
- 24:00 跨ぎは `shift_resolver.merge_overnight` で統合。

## 既知の落とし穴
- 月初の空欄で 1 日ズレる（BBS の AI 読み）。
- 月末夜勤の翌月 1 日の扱い。

## 宿題
- 常駐先が増えるたびにパーサが増える構造。統合時の共通化余地。
