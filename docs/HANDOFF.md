# 月次勤怠ワークフロー自動化 — 新セッション開始ハンドオフ

**作成日**: 2026-05-26（最終更新: 2026-05-26 MVP復活反映）
**目的**: 新規 Claude セッションで本プロジェクトを継続するための最初の 1 枚

## 🚩 まずこれだけ理解してください（30秒）

1. **目的**: 230名規模の月次給与勤怠処理（請求勤怠 ↔ jinjer 突合 → 修正 → jinjer 投入）を自動化
2. **2 トラック構成**:
   - **5月分（5/29 金 本番）**: MVP（2 スクリプト）で乗り切る ← **明日 5/27 に実装**
   - **6月以降**: 設計書 P0〜P6 を本格実装 ← **5/30 以降に着手**
3. **コード配置先**: `Z:\勤怠チェックシステム\kintai-checker\` 配下
   - MVP: 直下に `quick_merge.py` / `quick_export.py`
   - 本格版: `services/` に新規モジュール追加
4. **既存資産が豊富**: `Z:\API連携\` の jinjer API クライアント・dry-run フレームワーク等を流用する前提

## 📖 必読ドキュメント（順番に）

1. **本ハンドオフ** ← いま読んでいるこれ
2. **5月対応**: `docs/PLAN_5月本番_3営業日MVP.md` — MVP 実装プラン（明日着手）
3. **6月対応**: `docs/DESIGN_月次マスター_P0_P3.md` — フル設計書（改訂5、約1,000行）
4. `Z:\API連携\README.md` — 既存資産の総目録、特に「月次勤怠ワークフロー自動化プロジェクトとの関係」セクション
5. Claude Code メモリ:
   - `memory/MEMORY.md` — 索引
   - `memory/project_kintai_2026_05_mvp.md` — プロジェクト方針サマリ
   - `memory/jinjer_api_cheatsheet.md` — jinjer API ハマりどころ

## 📌 確定済みの方針（再議論不要）

| 項目 | 決定 |
|---|---|
| 5月本番（5/29 金）| MVP（quick_merge + quick_export）で対応。Web 画面手動アップロード |
| MVP の位置づけ | 設計書 P1 / P6 の **劣化版**。triage なし、全行を人間判断 |
| 6月以降 | 設計書 P0〜P6 を本格実装、`services/` 配下に追加 |
| 修正キューの構造（本格版）| 月次マスター_<YYYY-MM>.xlsx 内のシート。承認/却下/保留/未判断のプルダウン、却下通知ステータス列あり |
| トリアージロジック（本格版）| match_confidence / parse_failed / data_missing / day_crossing / submission_missing / fuzzy_match / monthly_only / comment_present / unknown_reason の各 reason コードを付与 |
| インポート方式（本格版）| dry-run → 本番、未判断1件でも既定で中断（--force-incomplete で続行可）|
| MVP のレビューゲート思想 | **承認行のみが CSV に乗る**ルールは MVP でも堅持 |

## ⏭️ 明日 5/27 水の作業フロー

### ユーザー側（並行）
1. AM: 既存 kintai-checker で 5 月分突合を走らせる
2. AM: jinjer から「汎用データ（まるめ適用後）」CSV をダウンロード
3. PM: 差異一覧シートで人間判断列を埋め始める

### Claude 側（AM）
1. ユーザーから「実装スタート」の合図
2. 突合結果xlsx 1 ファイルの列構成を確認
3. **`quick_merge.py` を実装**（2〜3 時間）
   - 配置: `Z:\勤怠チェックシステム\kintai-checker\quick_merge.py`
   - 入力: `Y:\給与明細\R8年\5月\突合結果\*.xlsx`
   - 出力: `Y:\給与明細\R8年\5月\差異一覧_2026-05.xlsx`
   - 詳細仕様: `docs/PLAN_5月本番_3営業日MVP.md` §2.1
   - 参考実装: `Z:\API連携\create_template.py`（openpyxl + データ検証プルダウン）
4. **`quick_export.py` を実装**（半日）
   - 配置: `Z:\勤怠チェックシステム\kintai-checker\quick_export.py`
   - 入力: 差異一覧xlsx + jinjer ダウンロードCSV
   - 出力: jinjer アップロード用CSV（CP932）
   - 詳細仕様: `docs/PLAN_5月本番_3営業日MVP.md` §2.2
   - 参考実装: `Z:\API連携\update_salary_unit_prices.py`（dry-run → execute パターン）

## 🛠️ 既存資産マップ（流用前提）

### Z:\API連携\

| ファイル | 流用ポイント |
|---|---|
| `jinjer_client.py` | **6月 P0** で `get_employees()` 呼び出し。認証＋全ページ取得済み |
| `config.py` | `.env` 読み込みパターン踏襲 |
| `update_salary_unit_prices.py` | **MVP quick_export** および **6月 P6** の dry-run → 本番フロー参考 |
| `change_attendance_group.py` / `add_affiliation_bulk.py` | jinjer 書き込み系 POST/PATCH 実装パターン |
| `create_template.py` / `create_attendance_template.py` | **MVP quick_merge** および **6月 P1** の openpyxl 入力フォーム生成参考 |
| `crossstaff-builder/services/` | サービス分割設計の構造参考 |
| `docs/jinjerAPI資料/` | jinjer 汎用データインポート CSV 仕様参考 |

### Z:\勤怠チェックシステム\kintai-checker\services\（既存）

| ファイル | 状態 |
|---|---|
| `jinjer_api_client.py` | 既存。`get_employees_with_affiliation()` 等の拡張が 6月 P0 で必要 |
| `jinjer_parser.py` | 既存。jinjer CSV パース |
| `timesheet_parser.py` | 既存。請求勤怠多形式パース、Vision OCR 含む |
| `shift_legend_parser.py` | 既存。シフト記号解決 |
| `matcher.py` | 既存。突合ロジック、`normalize_name_keys` 等 |
| `excel_exporter.py` | 既存。Excel 出力 |
| `multi_year_shift_parser.py` | 既存 |
| `jinjer_template_matcher.py` | 既存 |
| `jinjer_schedule_csv_exporter.py` | 既存 |
| `shift_resolver.py` | 既存 |

### 新規追加するもの

**MVP（5月、明日 5/27 実装）**:
```
quick_merge.py        ★明日 AM 実装
quick_export.py       ★明日 AM〜PM 実装
```

**本格実装（6月以降、5/30〜着手予定）** — 設計書 §2:
```
services/
├── employee_resolver.py          ★P0 新規
├── master_aggregator.py          ★P1 新規
├── triage.py                     ★P1 新規
├── comment_carryover.py          ★P2 新規
├── batch_runner.py               ★P3 新規
├── review_loader.py              ★P6 新規
└── kintai_import_builder.py      ★P6 新規

batch_run.py                       ★P3 新規 CLI エントリ
import_run.py                      ★P6 新規 CLI エントリ
```

## 📐 jinjer 汎用データ CSV（重要、type.id=5）

**ファイル**: `Z:\勤怠チェックシステム\kintai-checker\汎用データテンプレート\汎用データ(まるめ適用後)ダウンロード_9637_20260522150030.csv`

- 文字コード: **CP932（Shift-JIS）**
- 列数: **194 列**
- 必須列: `名前 / *従業員ID / *年月日 / *打刻グループID`
- 修正で実際に書き換える列:
  - 「勤務時間」差異 → 22列目（出勤1）/ 23列目（退勤1）
  - 「休憩時間」差異 → 42列目（休憩1）/ 43列目（復帰1）
  - 「実働時間」差異 → 機械判定不能、手動対応
- ダウンロード CSV をそのまま上書きアップロードで動作する（ユーザー実機確認済み）

ヘッダー全列一覧: `汎用データテンプレート\_headers.txt`

## 💻 環境

- Python 3.9 以上（既存）
- 既にインストール済み依存（requirements.txt）: `flask, pandas, openpyxl, pdfplumber, Pillow, pypdfium2, anthropic, python-dotenv, requests`
- 6月本格実装で追加推奨: `pytest, pytest-mock, tenacity, rapidfuzz, rich`
- `.env` 設定済み: `ANTHROPIC_API_KEY, JINJER_API_KEY, JINJER_SECRET_KEY, JINJER_BASE_URL`

## ⚠️ 罠と過去の認識ミス（重要）

- **「既存資産を見落として工数を倍に見積もる」のは私が一度やった失敗**。新規実装に着手する前に `Z:\API連携\` を最低限眺めること。
- **MVP の位置づけを誤解しない**：MVP は「設計書のフル機能を 3 日で詰め込む」ではなく、「**P1/P6 の最小機能だけを 3 日で動かす**」もの。triage / comment_carryover / API投入 等は 6月以降。
- **5月本番までに P0〜P6 のフル実装を試みない**：物理的に 9〜13 日必要、3 日には収まらない。試みると誤投入リスク大。
- **jinjer API の50件/リクエスト制限**は実用上の壁。`salaries` / `custom-items` などの複数指定系で要注意（`memory/jinjer_api_cheatsheet.md` 参照）。
- **修正キューのトリアージは「保守的」が正解**（本格版）。fuzzy_match / monthly_only / parse_failed / data_missing / day_crossing / submission_missing / unknown_reason は **コメントなしでも needs_check に降格**（設計書 §4.7）。
- **Excel/CSV 文字コード**: jinjer 関連 CSV は基本 CP932。openpyxl の xlsx はそのまま。

## 📞 ユーザーへの問い合わせ運用

ユーザーは管理部所属、Windows 11 + PowerShell 環境。バイブコーディング前提・1人運用。
複雑な選択は `AskUserQuestion` で 2〜4 択提示すると速い。
今日 2026-05-26 時点の確認済み事項は本ハンドオフに集約済み。

---

## 変更履歴

| 日付 | 内容 |
|---|---|
| 2026-05-26 初版 | MVP 廃止前提で作成 |
| 2026-05-26 更新 | MVP 復活反映。5月=MVP / 6月=本格 P0-P6 の 2 トラック構成に書き直し |
