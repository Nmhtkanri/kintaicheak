import os


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
        "timesheet": {"xlsx", "xls", "pdf", "png", "jpg", "jpeg"},
    }
