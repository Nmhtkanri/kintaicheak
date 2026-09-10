# 経費チェック（SAP 経費・通勤費・定期代）— 2026-09-10 転記・同日コード検証済

## 入口
- `/expense_prereview` `/expense_integration` `/expense_integration_preview` `/expense_import_addon_preview` `/expense_payroll_import` `/expense_telework` `/expense_month_status` `/sap_ledger_confirm` `/sap_ledger_status` `/keihi_import_ledger_status` `/travel_expense_members`（GET/POST）`/no_commute_extract` `/teiki_shiwake_preview` `/teiki_shiwake_generate`
- services: `expense_check` `keihi_classify` `keihi_summary` `keihi_payroll_import` `keihi_import_ledger` `sap_duplicate_filter` `sap_import_ledger` `kotsuhi_seisa` `no_commute_extract` `shiwake_teiki_append`

## 入力
- SAP 経費データ、jinjer の通勤費申請、テレワーク日数。
- 外部 CSV（`config.py`）：通勤費上限・免除者 `Z:\API連携\docs\通勤費_上限免除者.csv`（人数は CSV が正。転記時点で 7 名）、投入台帳。

## 出力と不可逆な処理
- 給与への追加投入（`/expense_payroll_import`）。jinjer 側は同月同社員を置換する仕様で、これが二重計上を防いでいる。
- 台帳は 2 種類で役割が違う：
  - `keihi_import_ledger` は「追加投入の土台として投入内容を保持する」台帳。status "2" は記録しない、記録に失敗しても投入は成功扱いで警告のみ、という仕様（`app.py`）。
  - 二重支給のガードは SAP 側の `sap_import_ledger`（費用シート ID／明細の 3 段階判定）。
- 仕訳 CSV への定期代追記（`shiwake_teiki_append`）。

## 業務ルール
- 通勤費の月額上限は 3 万円。**上限は定期代と実費の月累計の合算に掛かる（2026-08-06 確定、`kotsuhi_seisa.py`）。** 免除者と移動交通費対象者は OK 扱い。
- 通勤費申請なしリスト：2026-08-12 に「実費＋テレワーク＝出勤日数なら通過」へ拡張（`kotsuhi_seisa.py`、③単独出力の `no_commute_extract` も同ロジック共用）。
- ▲片側一致はレビュー対象として残す（`keihi_summary.py`、2026-08-12 決定、テストあり）。

## 既知の落とし穴
- 通勤費上限の関数は `limit=COMMUTE_MONTHLY_LIMIT` のデフォルト引数（monkeypatch が効かない形）。

## 宿題
- なし（実費への上限適用は確定済み）。
