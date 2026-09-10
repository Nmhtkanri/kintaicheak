# 勤怠チェック（差異一覧・書き戻し・API 投入）— 2026-09-10 転記・同日コード検証済

## 入口
- `/upload` `/quick_compare` `/quick_export` `/batch_compare` `/resolve_and_match` `/api_import` `/api_import_status/<job_id>` `/download/<filename>`
  - `/quick_export` は差異一覧 xlsx＋jinjer CSV からアップロード用 CSV を作る書き戻し経路。dry-run と本実行（`execute=1`）がある。
  - `/export_jinjer_csv` はシフト表→jinjer CSV 変換モードのルートで、勤怠突合ではない（schedule.md 参照）。
- 実装：**差異一覧はリポジトリ直下の `quick_compare.py`、書き戻しは直下の `quick_export.py`。** services 側は `timesheet_parser` `jinjer_parser` `matcher` `triage` `excel_exporter`（手順 1 の 17 列出力）`kintai_import_runner` `jinjer_api_client` `jinjer_template_matcher` `employee_alias` `batch_runner`。

## 入力
- jinjer 勤怠 CSV、**jinjer 申請データ（打刻修正申請）CSV（`/quick_compare` では必須。申請理由を差異一覧に載せるため）**、各部署・常駐先の勤務表（Excel/PDF/画像）。
- Fieldglass（ERCSTS）は「氏名_社員番号.pdf」に付け替える運用。本文はローマ字しか無く、名前で突合すると未提出者に落ちる（`timesheet_parser.py`）。

## 出力と不可逆な処理
- 差異一覧 Excel（2026-08-20 に 25 列→23 列。休憩時間は数式。シート保護のため並べ替え不可、`autoFilter=False`。回帰テストは `tests/test_quick_compare.py`）。
- 書き戻し CSV（`/quick_export`）と jinjer API 投入（`/api_import`、`kintai_import_runner` に投入前の門番チェックあり）。投入後の取り消しは手作業。

## 業務ルール（コードに実装あり）
- 休日・休暇の打刻抑制は「jinjer にも打刻が無い日」だけ（`quick_compare.py`、2026-09-02 追加。抑制を広げて休日出勤が丸ごと消えた事故が発端。回帰テストあり）。

## 業務ルール（jinjer 側の設定運用。このリポジトリに実装は無い）
- 140-180 時間制：基礎時給 =（基本給＋調整手当）÷160 切上。140 割れ控除 =（140 − 総労働）× 基礎時給を差額調整に翌月計上。ハブ側は経理モードで「基礎時給」を freee 非計上ラベルとして扱うだけ。
- 退職・休職者の支給項目ゼロ化：手当 1 は調整・リーダー・役職の 3 つ。みなし深夜は月給制 3 のみ。
- jinjer 計算式は文字列比較不可（" は禁止文字）。在籍区分は 0=在籍/1=退職/2=休職 の数値で比較（`expense_check.py` にも同じ前提）。

## 既知の落とし穴
- **API 投入の雛形 ID（2026-09-09 修正）。** 7〜9 月の検証 NG 165 件中 162 件が雛形時刻に負けて出勤予定が 9:00 のままになった。`quick_export.py` で雛形 ID を空欄化し、休暇日は打刻だけ送る（`kintai_import_runner.strip_schedule_for_kyuka`）。この扱いは勤怠側だけの話で、スケジュール取込は別方式（schedule.md）。

## 宿題
- 退職者での支給ゼロ化の実地確認（jinjer 側）。
