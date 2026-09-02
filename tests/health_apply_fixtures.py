# -*- coding: utf-8 -*-
"""健康診断申込テストの共有フィクスチャ: 偽の Sheets ゲートウェイとサンプルブック。

個人情報は一切含めない（社員番号は 20 始まり7桁の架空値、氏名も架空）。
"""

from __future__ import annotations

from services.health_apply import schema as S
from services.health_apply.sheets_gateway import GatewayError


class FakeSheetsGateway:
    """{シート名: [[セル...], ...]}（1行目=ヘッダー）を持つメモリ上のブック。

    Sheets API と同じく、read_values は行末の空セルを切り落として返す。
    append_rows は本物と同じ assert_writable を通す。
    """

    def __init__(self, sheets: dict[str, list[list[str]]], title: str = "テスト 健康診断申込"):
        self.sheets = {k: [list(map(str, r)) for r in v] for k, v in sheets.items()}
        self.title = title
        self.appended: list[tuple[str, list[list[str]]]] = []
        self.calls: list[tuple] = []
        self.fail_append_on: str | None = None   # このシートへの append で例外を出す
        self.fail_read: bool = False

    def get_metadata(self) -> dict:
        self.calls.append(("get_metadata",))
        return {"title": self.title,
                "sheets": [{"title": k, "sheetId": i, "rowCount": len(v), "columnCount": max((len(r) for r in v), default=0)}
                           for i, (k, v) in enumerate(self.sheets.items())]}

    def read_values(self, sheet: str, a1_range: str = "") -> list[list[str]]:
        self.calls.append(("read_values", sheet, a1_range))
        if self.fail_read:
            raise GatewayError("読み取りに失敗しました（テスト）")
        if sheet not in self.sheets:
            raise GatewayError(f"シートがありません: {sheet}")
        rows = self.sheets[sheet]
        if a1_range == "1:1":
            rows = rows[:1]
        out = []
        for r in rows:
            cells = list(r)
            while cells and cells[-1] == "":
                cells.pop()
            out.append(cells)
        return out

    def append_rows(self, sheet: str, rows: list[list[str]]) -> int:
        self.calls.append(("append_rows", sheet, len(rows)))
        try:
            S.assert_writable(sheet)
        except S.SchemaError as e:
            raise GatewayError(str(e)) from e
        if self.fail_append_on == sheet:
            raise GatewayError(f"{sheet} への書き込みに失敗しました（テスト）")
        if sheet not in self.sheets:
            raise GatewayError(f"シートがありません: {sheet}")
        rows = [list(map(str, r)) for r in rows]
        self.sheets[sheet].extend(rows)
        self.appended.append((sheet, rows))
        return len(rows)


# --- 行ビルダー -----------------------------------------------------------

def settings_rows(**overrides) -> list[list[str]]:
    kv = {
        "スキーマ版": S.SCHEMA_VERSION, "年度": "2027", "前年度": "2026",
        "受付開始": "2027-02-01", "受付終了": "2027-02-28",
        "受診期間開始": "2027-04-01", "受診期間終了": "2028-03-31", "回答受付": "1",
        "WebアプリURL": "https://script.google.com/macros/s/test/exec",
        "案内メール件名": "【健康診断】2027年度の申込", "案内メール本文": "{氏名} さん {URL}",
        "担当者連絡先": "kanri@example.invalid",
    }
    kv.update(overrides)
    return [list(S.SETTINGS_HEADERS)] + [[k, v, ""] for k, v in kv.items()]


def options_rows() -> list[list[str]]:
    rows = [list(S.OPTION_HEADERS)]
    rows += [
        ["機関", "1310528885", "医療法人社団 同友会 春日クリニック", "1", "10", "医療法人社団同友会 春日クリニック;同友会", ""],
        ["機関", "0301619", "医療法人徳洲会 生駒市立病院", "1", "20", "", ""],
        ["機関", "13X5035440", "関東ITソフトウェア健康組合(大久保健診センター)", "1", "30", "", ""],
        ["機関", "1310136358", "MYメディカルクリニック 大手町", "1", "40", "", ""],
        ["機関", "130192", "東京品川病院 総合健診センター", "0", "50", "", "無効化した例"],
        ["機関", "OTHER", "その他", "1", "999", "", ""],
        ["種別", "10", "定期健康診断", "1", "1", "基本健診", ""],
        ["種別", "11", "人間ドックA", "1", "2", "", ""],
        ["種別", "12", "人間ドックB", "1", "3", "1日人間ドック・バリウム", ""],
        ["種別", "13", "人間ドックC", "1", "4", "1日人間ドック・胃カメラ", ""],
        ["種別", "14", "雇用時の健康診断", "1", "5", "", ""],
        ["種別", "15", "人間ドック1日コース", "1", "6", "1日人間ドック", ""],
        ["追加検査", "GYN", "婦人科検診", "1", "1", "婦人病検査;婦人科健診", ""],
        ["続柄", "妻", "妻", "1", "1", "", ""],
        ["続柄", "夫", "夫", "1", "2", "", ""],
    ]
    return rows


def target_row(**kw) -> list[str]:
    base = {name: "" for name in S.TARGET_HEADERS}
    base.update({
        "年度": "2027", "社員番号": "2099001", "氏名": "試験 太郎", "社用メール": "t.shiken@nmht.co.jp",
        "在籍区分": "0", "前年度情報元": S.SOURCE_HISTORY,
        "前年度健診機関コード": "1310528885", "前年度健診機関名": "医療法人社団 同友会 春日クリニック",
        "前年度健診種別コード": "10", "前年度健診種別名": "定期健康診断", "前年度追加検査": "",
        "前年度健診機関(原文)": "医療法人社団 同友会 春日クリニック",
        "登録日時": "2027-01-10T09:00:00", "登録者": "yatsu",
        "申込状態": S.STATUS_UNSENT, "送信回数": "0", "回答版": "0",
    })
    base.update(kw)
    return [base[name] for name in S.TARGET_HEADERS]


def response_row(**kw) -> list[str]:
    base = {name: "" for name in S.RESPONSE_HEADERS}
    base.update({
        "回答日時": "2027-02-05T10:00:00", "受付番号": "HC-2027-2099001-01", "年度": "2027",
        "社員番号": "2099001", "回答版": "1", "氏名": "試験 太郎", "社用メール": "t.shiken@nmht.co.jp",
        "申込区分": S.KIND_CHANGE, "健診機関コード": "1310136358", "健診機関名": "MYメディカルクリニック 大手町",
        "その他医療機関名": "", "健診種別コード": "13", "健診種別名": "人間ドックC", "追加検査": "",
        "その他健診予定日": "", "被扶養者申込": "0", "続柄": "", "被扶養者氏名": "", "備考": "",
        "トークンハッシュ": "ab" * 32, "回答元": "Web",
    })
    base.update(kw)
    return [base[name] for name in S.RESPONSE_HEADERS]


def workbook(targets=(), responses=(), audit=(), settings=None, options=None) -> dict[str, list[list[str]]]:
    return {
        S.SHEET_SETTINGS: settings if settings is not None else settings_rows(),
        S.SHEET_OPTIONS: options if options is not None else options_rows(),
        S.SHEET_TARGETS: [list(S.TARGET_HEADERS)] + [list(r) for r in targets],
        S.SHEET_RESPONSES: [list(S.RESPONSE_HEADERS)] + [list(r) for r in responses],
        S.SHEET_AUDIT: [list(S.AUDIT_HEADERS)] + [list(r) for r in audit],
    }
