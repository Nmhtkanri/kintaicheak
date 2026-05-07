import glob
import os


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

    # デフォルトの雛形フォルダ
    default_dir = r"Z:\jinjer移行\カレンダー"
    if os.path.isdir(default_dir):
        candidates = glob.glob(os.path.join(default_dir, "スケジュール雛形一覧_*.csv"))
        if candidates:
            # ファイル名にYYYY-MM-DDが入っているので辞書順ソートで最新が末尾
            candidates.sort()
            return candidates[-1]

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
    ANTHROPIC_MODEL = "claude-sonnet-4-20250514"
    ANTHROPIC_MAX_TOKENS = 4096

    # シフト記号モード（jinjer 雛形マッチング用 CSV のパス）
    # 設定されている場合のみ、未マッチ記号から「新規雛形 CSV」を自動生成する
    # フォルダ内で最新の「スケジュール雛形一覧_YYYY-MM-DD.csv」を自動検出する
    JINJER_TEMPLATE_CSV_PATH = _resolve_jinjer_template_csv_path()

    # シフトデータの一時保存ディレクトリ（凡例レビュー → resolve のセッション用）
    SHIFT_SESSION_FOLDER = "uploads/sessions"

    # jinjer API（CSV変換モードで従業員ID取得に使用）
    JINJER_API_KEY = os.environ.get("JINJER_API_KEY")
    JINJER_SECRET_KEY = os.environ.get("JINJER_SECRET_KEY")
    JINJER_BASE_URL = os.environ.get("JINJER_BASE_URL", "https://api.jinjer.biz")

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
        "timesheet": {"xlsx", "xls", "csv", "txt", "pdf", "png", "jpg", "jpeg"},
    }
