import glob
import os
import tempfile


DEFAULT_JINJER_TEMPLATE_DIRS = (
    r"Z:\jinjer移行\カレンダー",
    r"\\Nmht\NMHumaTach\jinjer移行\カレンダー",
)


def _resolve_jinjer_template_csv_path() -> str:
    """jinjer 雛形 CSV の最新版を自動検出する

    優先順位:
      1. 環境変数 JINJER_TEMPLATE_CSV_PATH が指すファイルが存在すればそれ
      2. 雛形フォルダ内の `スケジュール雛形一覧_YYYY-MM-DD.csv` で
         ファイル名の日付が最も新しいもの
      3. どちらも無ければ環境変数の値（無ければ空文字）をそのまま返す
         → 後段の load_jinjer_templates が「ファイル無し」警告を出す
    """
    env_path = os.environ.get("JINJER_TEMPLATE_CSV_PATH", "")
    if env_path and os.path.exists(env_path):
        return env_path

    candidates = []
    for default_dir in DEFAULT_JINJER_TEMPLATE_DIRS:
        if os.path.isdir(default_dir):
            candidates.extend(
                glob.glob(os.path.join(default_dir, "スケジュール雛形一覧_*.csv"))
            )
    if candidates:
        # Z: が見えない実行環境では同じ共有先の UNC パスを使う。
        return max(candidates, key=lambda path: os.path.basename(path))

    return env_path  # 見つからなくてもそのまま返す（警告は load 側で出る）


class Config:
    # Flask
    SECRET_KEY = os.environ.get("SECRET_KEY", "dev-secret-key")
    UPLOAD_FOLDER = "uploads"
    OUTPUT_FOLDER = "outputs"
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB

    # uploads に残った入力ファイルの写しを何日で消すか。
    # 処理の途中で例外が出た実行では消し漏れるため、起動時にまとめて掃除する。
    # uploads にあるのは「処理のために保存した入力のコピー」で、成果物は outputs 側。
    # 氏名や勤怠が入ったものが共有フォルダに残り続けないようにするのが目的。
    # 0 以下にすると掃除しない。
    UPLOAD_RETENTION_DAYS = int(os.environ.get("UPLOAD_RETENTION_DAYS", "7"))

    # 許容差分（デフォルト値、画面から変更可能）
    DEFAULT_THRESHOLD_MINUTES = 10

    # Claude API
    # モデルは .env の ANTHROPIC_MODEL で上書き可能（未設定ならデフォルト値）。
    # → 将来モデルが提供終了になっても .env を書き換えるだけでよく、exe 再ビルド不要。
    # 履歴: claude-sonnet-4-20250514 (Sonnet 4) は 2026-06-15 に提供終了 → claude-sonnet-4-6 へ。
    ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")
    # 全員分のシフト表(例: 21名×30日)を1回で返すと 4096 では出力が途中で切れて壊れる。
    # 余裕を持たせる。.env の ANTHROPIC_MAX_TOKENS で上書き可能。
    ANTHROPIC_MAX_TOKENS = int(os.environ.get("ANTHROPIC_MAX_TOKENS", "16000"))
    # AI 読み取り1回あたりの上限時間（秒）。**必ず設定すること**。
    # 未設定だと、応答が返らないときに画面が「処理中...」のまま無限に固まる
    # （2026-07-31 実例: 複数シートのブックで25分以上ハング）。
    ANTHROPIC_TIMEOUT_SECONDS = float(os.environ.get("ANTHROPIC_TIMEOUT_SECONDS", "300"))
    # AI に送るテキストの上限文字数。これを超えるファイルは送信前にエラーで止める
    # （送っても max_tokens に収まる回答が返らず、時間だけ溶けるため）。
    AI_MAX_INPUT_CHARS = int(os.environ.get("AI_MAX_INPUT_CHARS", "400000"))

    # シフト記号モード（jinjer 雛形マッチング用 CSV のパス）
    # 設定されている場合のみ、未マッチ記号から「新規雛形 CSV」を自動生成する
    # フォルダ内で最新の「スケジュール雛形一覧_YYYY-MM-DD.csv」を自動検出する
    JINJER_TEMPLATE_CSV_PATH = _resolve_jinjer_template_csv_path()

    @staticmethod
    def get_jinjer_template_csv_path() -> str:
        """Return the current latest jinjer schedule template CSV path."""
        return _resolve_jinjer_template_csv_path()

    # シフトデータの一時保存ディレクトリ（凡例レビュー → resolve のセッション用）
    SHIFT_SESSION_FOLDER = "uploads/sessions"

    # スケジュールアップロード: シフト表氏名 → 従業員ID の読み替え表（エイリアス）
    # 同姓が複数いる姓（吉田・佐藤 等）はシフト表が姓だけだと自動確定できないため、
    # **シフト表の系統ごと**に読み替えを持つ（services/employee_alias.py 参照）。
    # マッピング表方式＝谷津さんが直したら exe 再ビルドなしで次回実行から効く。
    SCHEDULE_NAME_ALIAS_DIR = os.environ.get(
        "SCHEDULE_NAME_ALIAS_DIR", r"Z:\API連携\docs")

    # jinjer API（CSV変換モードで従業員ID取得に使用）
    JINJER_API_KEY = os.environ.get("JINJER_API_KEY")
    JINJER_SECRET_KEY = os.environ.get("JINJER_SECRET_KEY")
    JINJER_BASE_URL = os.environ.get("JINJER_BASE_URL", "https://api.jinjer.biz")

    # ------------------------------------------------------------------
    # 経費統合: SAP重複除外（取込済み費用シート台帳方式。2026-08-06 谷津さん決定）
    # ------------------------------------------------------------------
    # SAPの月次CSVは前々月分が再掲されて落ちてくるため、取込前に除外する必要がある。
    # 旧方式（前月の生CSVを画面で選んで突合）は 2026-08-06 に破綻した:
    #   実際に7月へ取り込んだのは (4)_重複除外済.csv だったが、旧ツールは名前に
    #   「重複除外済」を含むCSVを過去データから自動スキップする。SAP元データ フォルダに
    #   あった (5).csv は7月に使わなかった別DLで22IDが欠落しており、これを過去に指定すると
    #   除外0行になって 285,642円（13名/93行）が二重支給になりかけた。
    # → 取り込んだ費用シートを台帳に記録し、翌月はファイルではなく台帳と突合する方式へ移行。
    #   年度をまたぐので R8年 フォルダの外に置く。
    KEIHI_SAP_LEDGER_CSV = os.environ.get(
        "KEIHI_SAP_LEDGER_CSV", r"Y:\給与明細\_SAP取込済み費用シート台帳.csv")
    # 台帳へ書き込めるユーザー（列: ユーザー名, 表示名, 備考）。台帳が狂うと給与が狂うため、
    # 共有exeを使う全員が書ける状態にしない。谷津さんが直すファイルなので共有フォルダを読む
    # （再ビルドなしで平良さんのユーザー名を追加できる）。読めないときは全員書き込み不可。
    KEIHI_SAP_LEDGER_WRITERS_CSV = os.environ.get(
        "KEIHI_SAP_LEDGER_WRITERS_CSV", r"Z:\API連携\docs\SAP台帳_書き込み許可ユーザー.csv")

    # 経費チェック: jinjer給与へ投入したインポート内容の台帳（2026-08-12 実装）。
    # 投入後にイレギュラー経費が出たとき、統合一覧表を4ソースから作り直さずに
    # 「前回分＋追加分」を再投入するための土台。月をまたぐので年度フォルダの外に置く。
    KEIHI_IMPORT_LEDGER_CSV = os.environ.get(
        "KEIHI_IMPORT_LEDGER_CSV", r"Y:\給与明細\_経費インポート投入台帳.csv")

    # 経費チェック: 移動交通費（立替精算）対象者リスト（2026-08-03 谷津さん指定）。
    # 現場直行などで通勤経路の登録が無くて正しい人たち＝交通費は通勤費でなく
    # 移動交通費（立替精算）で計上する。谷津さんが直すファイルなので exe に同梱せず
    # 共有フォルダを読む（再ビルド不要で反映）。列: A=社員番号, B=氏名
    KEIHI_TRAVEL_EXPENSE_MEMBERS_CSV = os.environ.get(
        "KEIHI_TRAVEL_EXPENSE_MEMBERS_CSV", r"Z:\API連携\docs\経費チェック_移動交通費対象者.csv")

    # ------------------------------------------------------------------
    # 経理モード（freee 給与仕訳の生成）
    # ------------------------------------------------------------------
    # マッピング表は谷津さんがレビューで直すファイルなので **exe に同梱せず共有フォルダを読む**。
    # 直したら再ビルドなしで次回実行から効く。画面のパス欄で 1回限りの上書きもできる。
    KEIRI_MASTER_CSV = os.environ.get(
        "KEIRI_MASTER_CSV", r"Z:\API連携\docs\経理モード_品目マッピングマスタ_draftC.csv")
    # 給与の「その他」(allowance52) のうち経費一覧表マクロに明細が無いもの（jinjer へ手入力
    # された有給買取・事務手数料など）を、画面で入力して仕訳に載せるための台帳。
    # 列: 支給月, 社員番号, 氏名, 金額, 勘定科目, 品目, 税区分, 備考
    # 画面から追記するが、共有フォルダに置くので Excel で直接直しても次回実行から効く。
    KEIRI_SONOTA_MANUAL_CSV = os.environ.get(
        "KEIRI_SONOTA_MANUAL_CSV", r"Z:\API連携\docs\経理モード_その他手入力.csv")
    KEIRI_KEIHI_MAPPING_CSV = os.environ.get(
        "KEIRI_KEIHI_MAPPING_CSV", r"Z:\API連携\docs\経理モード_経費転記マッピング_draftD.csv")
    # C-4 の分解材料（経費利用履歴 RevN.xlsm）と、突合先の経理最終CSV。どちらも {M}月 フォルダ配下
    KEIRI_KEIHI_BOOK_DIR = os.environ.get("KEIRI_KEIHI_BOOK_DIR", r"Y:\給与明細\R8年")
    KEIRI_FINAL_CSV_DIR = os.environ.get("KEIRI_FINAL_CSV_DIR", r"Y:\給与明細\R8年")
    # 生成物と API キャッシュの置き場（個人情報を含むのでローカル）
    KEIRI_OUTPUT_DIR = os.environ.get("KEIRI_OUTPUT_DIR", "outputs/keiri")

    # ------------------------------------------------------------------
    # 社労士モード（前田事務所へ渡す給与CSVの生成）
    # ------------------------------------------------------------------
    # 列の対応は給与体系ごとに変わる。谷津さんが直せるよう共有フォルダに置き、
    # 直したら再ビルドなしで次回実行から効く。無ければコード内の既定表で動く。
    # 列: 列番号, 列名, 給与体系, source_key, 体系別名条件, 備考
    SHAROUSHI_COLUMN_MAPPING_CSV = os.environ.get(
        "SHAROUSHI_COLUMN_MAPPING_CSV", r"Z:\API連携\docs\社労士モード_列マッピング.csv")
    # 給与計算のあとに発生した支給（jinjer に入っていない分）を足すための台帳。
    # 台帳の額は差引支給額と立替金列に加算し、口座1振込額と総支給額は据え置く。
    # 列: 支給月, 社員番号, 氏名, 項目, 金額, メモ
    SHAROUSHI_EXTRA_LEDGER_CSV = os.environ.get(
        "SHAROUSHI_EXTRA_LEDGER_CSV", r"Z:\API連携\docs\社労士モード_追加支給台帳.csv")
    # イレギュラー発生5項目（定常外業務対応手当・その他手当・現物支給・支給過不足調整・
    # 社保調整）の発生理由。金額は jinjer が正で、ここには理由だけを持たせる。
    # 画面から入力するが、共有フォルダなので Excel で直接直しても次回実行から効く。
    # 列: 支給月, 社員番号, 氏名, 項目, 理由
    SHAROUSHI_BIKO_LEDGER_CSV = os.environ.get(
        "SHAROUSHI_BIKO_LEDGER_CSV", r"Z:\API連携\docs\社労士モード_備考台帳.csv")
    # 生成物の置き場（氏名・給与額を含むのでローカル）
    SHAROUSHI_OUTPUT_DIR = os.environ.get("SHAROUSHI_OUTPUT_DIR", "outputs/sharoushi")

    # 手順③ API直接投入（kintai-imports）
    # executor = jinjer標準の完了通知メール宛先。
    # ⚠️ 明示指定できるのは「勤怠管理者権限ロール」を持つ社員番号のみ。
    #    ロールが無い番号を指定すると POST は 200 でも予約がサイレント破棄される
    #    （2026-07-09 実測: 9999999・2026007 とも現状ロール無しで破棄。
    #      未指定ならマスタ扱いで通り、通知は 2018013 宛に届く）。
    # 管理部メール（9999999）宛にしたい場合は、jinjer画面で 9999999 に
    # 勤怠管理者ロールを付与してから .env に JINJER_IMPORT_EXECUTOR_ID=9999999 を設定する。
    JINJER_IMPORT_EXECUTOR_ID = os.environ.get("JINJER_IMPORT_EXECUTOR_ID", "")

    # 通勤費の精査対象外者リスト（社員番号,氏名,理由）。共有フォルダを読む（再ビルド不要）。
    # 勤怠データからは判別できない事情で通勤費が発生しない人を落とす。
    KOTSUHI_EXCLUDED_MEMBERS_CSV = os.environ.get(
        "KOTSUHI_EXCLUDED_MEMBERS_CSV", r"Z:\API連携\docs\通勤費_精査対象外者.csv")

    # 通勤費の月額上限と、上限を適用しない免除者リスト（社員番号,氏名,理由）。
    # 当社の通勤費は月3万円が上限で、超えた分は基本的に切る。ただし個別に実費全額を
    # 認めている人がいるため明示リストで持つ（2026-08-06 谷津さん指定）。
    # 移動交通費（立替精算）には上限が無いので、この判定には含めない。
    # 共有フォルダを読むので、対象者が増えても exe 再ビルドなしで反映される。
    KOTSUHI_MONTHLY_LIMIT = int(os.environ.get("KOTSUHI_MONTHLY_LIMIT", "30000"))
    KOTSUHI_LIMIT_EXEMPT_MEMBERS_CSV = os.environ.get(
        "KOTSUHI_LIMIT_EXEMPT_MEMBERS_CSV", r"Z:\API連携\docs\通勤費_上限免除者.csv")

    # ------------------------------------------------------------------
    # メール下書きモード（Outlook 下書きの一括作成。送信機能は作らない）
    # ------------------------------------------------------------------
    # 台帳・テンプレートは共有フォルダを読む（経理モードのマッピング表方式。再ビルド不要）。
    MAIL_ADDRESS_BOOK = os.environ.get(
        "MAIL_ADDRESS_BOOK", r"Z:\NMHT社員勤務表\便利マクロ\メール一斉送信マクロ.xlsm")
    MAIL_TEMPLATES_JSON = os.environ.get(
        "MAIL_TEMPLATES_JSON", r"Z:\API連携\docs\メールテンプレート.json")
    # CC が空欄のときに自動で入れる既定CC（管理部の控え。谷津さん指定 2026-07-29）
    MAIL_DEFAULT_CC = os.environ.get("MAIL_DEFAULT_CC", "kanri@nmht.co.jp")
    # 下書き作成ログ（宛先を含むのでローカル）
    MAIL_OUTPUT_DIR = os.environ.get("MAIL_OUTPUT_DIR", "outputs/mail")

    # ------------------------------------------------------------------
    # 健康診断HPMモード（整形済ExcelからHPM取込用CSVを作る。取込は手動のまま）
    # ------------------------------------------------------------------
    # 機関コード・健診コース・302列ヘッダー・項目マッピングは共有フォルダの
    # 変換マスタで持つ。HPM側で列や機関が増えても再ビルドなしで反映できる。
    HEALTH_HPM_MASTER_XLSX = os.environ.get(
        "HEALTH_HPM_MASTER_XLSX", r"Z:\NMHT総務関係\健康診断\健康診断HPM変換マスタ.xlsx")
    # 出力先のベース。実際は {年度}\{年度}年度健康診断受診者結果\CSV格納 へ書く。
    HEALTH_HPM_OUTPUT_BASE = os.environ.get(
        "HEALTH_HPM_OUTPUT_BASE", r"Z:\NMHT総務関係\健康診断")
    # 健診結果は要配慮個人情報。launcher.py が作業フォルダを共有NASへ chdir するため、
    # 既定の uploads/sessions に置くと6人が読める場所にpklが残る。各PCのローカルに隔離する。
    HEALTH_HPM_SESSION_DIR = os.environ.get(
        "HEALTH_HPM_SESSION_DIR",
        os.path.join(
            os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(),
            "KintaiChecker", "health_sessions",
        ),
    )
    # 上記フォルダに残った一時ファイルを消すまでの時間（プレビューのたびに掃除する）
    HEALTH_HPM_SESSION_MAX_AGE_HOURS = int(
        os.environ.get("HEALTH_HPM_SESSION_MAX_AGE_HOURS", "8"))
    # PDFを直接読むときのページ数上限。1ページ＝Claude API 1回（10〜20秒）なので、
    # 間違って分厚いPDFを入れたときに時間と課金が伸び続けるのを止める。
    # 実績: 同友会の1回分は3〜7ページ。増やすなら環境変数で。
    HEALTH_HPM_PDF_MAX_PAGES = int(os.environ.get("HEALTH_HPM_PDF_MAX_PAGES", "20"))

    # jinjer CSV カラムマッピング候補
    # ※先頭に置くほど優先度が高い（完全一致を試みたあと部分一致）
    JINJER_COLUMN_MAPPING = {
        "氏名": ["名前", "氏名", "社員名", "従業員名", "スタッフ名"],
        "日付": ["*年月日", "年月日", "日付", "勤務日", "出勤日", "対象日"],
        # 出勤1/退勤1 は完全一致で先にヒットさせ、出勤10等への誤マッチを防ぐ
        "出勤時刻": ["出勤1", "出勤", "出勤時刻", "始業", "始業時刻", "出勤時間"],
        "退勤時刻": ["退勤1", "退勤", "退勤時刻", "終業", "終業時刻", "退勤時間"],
        # 「実労働時間」は別の意味を持つため代用せず、総労働時間の完全一致だけを使う
        "総労働時間": ["総労働時間"],
        # 第1コメント列（打刻時）と第2コメント列（管理者備考）を別キーで管理
        "コメント": ["打刻時コメント", "備考", "コメント", "メモ", "申請理由", "理由", "申請事由"],
        "コメント2": ["管理者備考"],
    }

    # 対応ファイル形式
    ALLOWED_EXTENSIONS = {
        "jinjer": {"csv"},
        "timesheet": {"xlsx", "xls", "xlsb", "csv", "txt", "pdf", "png", "jpg", "jpeg"},
        "health_excel": {"xlsx"},
    }
