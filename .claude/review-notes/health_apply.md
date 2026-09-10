# 健診申込モード（Google スプレッドシート連携）— 2026-09-10 転記・同日コード検証済

## 入口
- `/health_apply_active_employees` `/health_apply_targets_preview` `/health_apply_targets_commit` `/health_apply_responses` `/health_apply_status`
- services: `services/health_apply`（11 モジュール）

## 入力
- jinjer の在籍者（GET のみ・逐次）。
- Google シート。**鍵 JSON は管理者 PC のローカル固定で、共有フォルダ・リポジトリ・.env には置かない**（`config.py`、`services/health_apply/settings.py`）。年度設定 JSON と閲覧許可ユーザー CSV は共有 Z: に置く。
- プレビューは個人情報を含むためローカルに 2 時間だけ保持。

## 出力と不可逆な処理
- **シートへの対象者登録（`/health_apply_targets_commit`）** が唯一の外部書き込み。ガードは、許可 CSV、シート構成の検証、確認語、プレビュー指紋の一致、読み直して変わっていれば 409。追記と監査ログを残す。

## 既知の落とし穴
- 2026-09-02 実装・main マージ済だが、実接続テストは未実施。exe 再ビルドも実接続テスト後に 1 回の予定。鍵 JSON とシートは谷津さん待ち。鍵はローカル固定なので、どの PC で誰が実接続テストをするかを決める必要がある。

## 宿題
- 実接続テスト。鍵の無い PC での挙動。
