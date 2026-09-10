# 共通の前提（全モード共通）— 2026-09-10 転記・同日コード検証済

## リポジトリの構造

- Flask 1 本（`app.py`、`@app.route` 78 本。Blueprint や `add_url_rule` は無い）＋ `services/` 配下のモジュール。画面は `templates/base.html` `templates/index.html` と `static/script.js` `static/style.css` だけ。
- **`services/` の外にも本体がある。** 勤怠モードの差異一覧（23 列・数式・シート保護）はリポジトリ直下の `quick_compare.py`、jinjer への書き戻し CSV は直下の `quick_export.py`。`services/excel_exporter.py` は手順 1 の突合結果（17 列）で別物。`services/` だけ読んで合格にしない。
- 製品名は「オペレーションハブ」。フォルダ名・exe 名（KintaiChecker.exe）は旧名のまま（意図的）。
- 既存の運用ルールは `AGENTS.md`、開発フローは `CLAUDE.md`。

## 既知の落とし穴（コードの書き方）

- **同型コードが複数ルートにある。** 例：`/resolve_and_match` と `/export_jinjer_csv` に「未マッチ雛形→新規雛形 CSV」の別実装が並んでいる。片方だけ直す事故が起きた。変更した関数の複製が無いか Grep で確認し、テストは test_client で該当ルートを直接叩いて確かめる。
- **インライン JS の罠。** `script.js` は body 末尾で読み込まれる（`base.html`）。インラインスクリプトの先頭で `script.js` の関数を呼ぶと画面の全ボタンが死ぬ。機械的な検査は `tests/test_app_*_ui.py` と id 重複検査 `tests/test_app_index_unique_ids.py` にあり、画面変更ではこれらを通したうえで「コンソールにエラーなし＋既存ボタンを含む実クリック」まで確認する。実クリックは Claude 側では出来ない（UNC パス上でブラウザペインが使えない）ので、人間確認事項になる。
- **デフォルト引数 × monkeypatch。** 定数をデフォルト引数にした関数は monkeypatch でパスを差し替えられず、テストが実データ（共有フォルダ）を読んだニアミスがある。この形は現存する：`keiri_keihi_tenki.load_mapping(path=MAPPING_CSV)`、`kotsuhi_seisa`（`limit=COMMUTE_MONTHLY_LIMIT`）、`services/daicho` の `jinjer_api` `jinjer_attach` `direct`。
- **共有 CSV/JSON の外部マスタが多い。** 投入台帳、経理その他手入力、請求書フォルダ書込許可者、社労士備考台帳、通勤費上限・免除者など（`config.py`）。パスの直書きと、ファイルが無いときの挙動を見る。

## 既知の落とし穴（環境）

- **NAS の怪現象。** 削除保留で Access denied になる。書き込みが他プロセスに数分見えず、Jinja が旧画面を配信し続けることがある（ビルド前は python の目で内容を確認してから touch）。
- **exe 再ビルドは worktree でなくルートで。** worktree は未コミット分を持たず機能欠落 exe になる（2026-08-21 実例）。最新判定は mtime でなく中身で。
- **5000 番ポート占有で配布が止まる。** `起動.bat` のポート判定が robocopy 同期より前にあり、already_running へ飛ぶ。
- **並行セッション。** 同じ作業ツリーで別の Claude セッションが動くことがある。着手前に mtime と `git status` を見る。
- **jinjer API のレート制限（429）はテナント単位。** 他の大量取得と並行すると落ちる。10〜15 分空けて再実行（`AGENTS.md`）。

## 不可逆な処理の一覧（横断）

- jinjer への投入：勤怠 `/api_import`、勤怠の書き戻し CSV `/quick_export`（`execute=1`）、スケジュール `/schedule_api_import`、標報 `/shaho_import_execute`、派遣台帳の添付 `/haken_attach_execute`
- freee への取込用 CSV：`/keiri_run`、`/invoice_export`（取り込むのは人だが、金額の誤りは後から直しにくい）
- メール送信 `/mail_send_execute`（Outlook COM）
- Google シートへの登録 `/health_apply_targets_commit`
- 共有フォルダ上の設定 CSV の書き換え：`/invoice_folders_save`、`/keiri_sonota_save`、`/sharoushi_biko_save`、`/mail_ledger_apply`、`/mail_templates_save`
