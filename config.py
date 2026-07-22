import glob
import os


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

    # jinjer API（CSV変換モードで従業員ID取得に使用）
    JINJER_API_KEY = os.environ.get("JINJER_API_KEY")
    JINJER_SECRET_KEY = os.environ.get("JINJER_SECRET_KEY")
    JINJER_BASE_URL = os.environ.get("JINJER_BASE_URL", "https://api.jinjer.biz")

    # 経費統合: SAP重複除外の過去CSV既定フォルダ（画面の初期値。サブフォルダも走査）。
    # SAPの月次CSVは前々月分が混ざって落ちてくるため、過去CSVと費用シートIDで突合して
    # 取込前に除外する。画面で空欄にすれば従来通り除外なし。.env の KEIHI_SAP_PAST_DIR で上書き可
    KEIHI_SAP_PAST_DIR = os.environ.get("KEIHI_SAP_PAST_DIR", r"Y:\給与明細\R8年")

    # 手順③ API直接投入（kintai-imports）
    # executor = jinjer標準の完了通知メール宛先。
    # ⚠️ 明示指定できるのは「勤怠管理者権限ロール」を持つ社員番号のみ。
    #    ロールが無い番号を指定すると POST は 200 でも予約がサイレント破棄される
    #    （2026-07-09 実測: 9999999・2026007 とも現状ロール無しで破棄。
    #      未指定ならマスタ扱いで通り、通知は 2018013 宛に届く）。
    # 管理部メール（9999999）宛にしたい場合は、jinjer画面で 9999999 に
    # 勤怠管理者ロールを付与してから .env に JINJER_IMPORT_EXECUTOR_ID=9999999 を設定する。
    JINJER_IMPORT_EXECUTOR_ID = os.environ.get("JINJER_IMPORT_EXECUTOR_ID", "")

    # jinjer CSV カラムマッピング候補
    # ※先頭に置くほど優先度が高い（完全一致を試みたあと部分一致）
    JINJER_COLUMN_MAPPING = {
        "氏名": ["名前", "氏名", "社員名", "従業員名", "スタッフ名"],
        "日付": ["*年月日", "年月日", "日付", "勤務日", "出勤日", "対象日"],
        # 出勤1/退勤1 は完全一致で先にヒットさせ、出勤10等への誤マッチを防ぐ
        "出勤時刻": ["出勤1", "出勤", "出勤時刻", "始業", "始業時刻", "出勤時間"],
        "退勤時刻": ["退勤1", "退勤", "退勤時刻", "終業", "終業時刻", "退勤時間"],
        # 第1コメント列（打刻時）と第2コメント列（管理者備考）を別キーで管理
        "コメント": ["打刻時コメント", "備考", "コメント", "メモ", "申請理由", "理由", "申請事由"],
        "コメント2": ["管理者備考"],
    }

    # 対応ファイル形式
    ALLOWED_EXTENSIONS = {
        "jinjer": {"csv"},
        "timesheet": {"xlsx", "xls", "xlsb", "csv", "txt", "pdf", "png", "jpg", "jpeg"},
    }
