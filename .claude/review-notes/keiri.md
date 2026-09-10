# 経理モード（jinjer 給与明細 → freee 仕訳）— 2026-09-10 転記・同日コード検証済

## 入口
- `/keiri_run` `/keiri_sonota_save` `/keiri_download/<ym>/<path:filename>`
- services: `keiri_engine` `keiri_api` `keiri_diff` `keiri_keihi_tenki`

## 入力
- jinjer API の給与明細（キャッシュは社労士モードと共用）。
- **経費チェックモードが出力する `経費統合一覧表.xlsx` は読まない。** ただし `経費利用履歴 RevN.xlsm` の中の「経費統合一覧表」シートは読む（`services/keiri_keihi_tenki.py` で allowance52 の分解に使う「備考(明細)」を取る。列は定数直書き `I_MEMO_LINE = 19`、行の引き当ては「集計ログの行番号＝シートの 1-based 行番号」という前提）。このシートの列追加・並べ替え・行ズレは備考の科目判定を壊す。
- 外部マスタ（`config.py`）：品目マッピングマスタ `KEIRI_MASTER_CSV`、経費転記マッピング `KEIRI_KEIHI_MAPPING_CSV`、経費利用履歴ブックの置き場 `KEIRI_KEIHI_BOOK_DIR`。マスタ 1 行の変更で freee CSV の中身が変わる。
- その他台帳 CSV（手入力分の科目決め）。追加投入 5 項目（定常外業務対応手当・その他手当・現物支給・支給過不足調整・社保調整）は自動計上。

## 出力と不可逆な処理
- freee 取込用仕訳 CSV。その他台帳 CSV の保存（`/keiri_sonota_save`）は共有側 `Z:\API連携\docs\経理モード_その他手入力.csv` を書き換える。

## 業務ルール
- 社保の 1 か月ズレ：会社負担＝当月分／本人控除＝前月分（`keiri_engine.py` の `build_shaho` 周辺）。子ども・子育て支援金は API に会社負担側の項目が無く（jinjer 画面の「事業主保険料 > 子ども・子育て支援金」は API の salary_other_items に出てこない。2026-04〜08 で確認）、同額の本人控除 child_support で代用している。①主取引では会社負担を前月明細、支援金を当月明細から読んで月を揃える。
- その他手当は allowance20/21 の 2 つある。調整手当は allowance15→12 へ移設済み（2026-08）。暫定固定項目には 12 と 15 の両方を残している。
- 対象外項目の新規使用検知：画面赤枠・人数金額・要確認 md（先頭見出しと「骨格の未実装（TODO）」節の次）に出す。`KEIRI_NEW_USAGE_BLOCK` で停止切替（既定オフ、テストで assert 済）。要確認 md の新規使用検知を毎月見る。

## 既知の落とし穴
- jinjer API 429 はテナント単位。他の大量取得と並行すると経理モードが落ちる（コード側のレート制御の有無は未確認）。
- 有給買取・手数料が jinjer 手入力で「その他」に混ざる。台帳で科目を決める。

## 宿題
- なし。子育て支援金の 1 か月ズレは 2026-09-10 に解消（①主取引の支援金だけ当月明細の本人控除から読む。`build_shaho.shaho_total(pi, pi_deduction)`）。実データ 2026-07／08 で修正前後を比較し、差分は月変・復職の 34 人分の支援金だけであることを確認済み（`tests/test_keiri_shaho_child_support.py` の docstring）。
