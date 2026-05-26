# 月次給与勤怠ワークフロー自動化 — P0〜P3 詳細設計（+ P5/P6 拡張）

**作成日**: 2026-05-19（最終更新 2026-05-26）
**対象**: 230名規模の月次給与勤怠処理を、突合〜マスター集約〜人間レビュー〜jinjer インポートまで自動化する
**前提となる調査結果**:
- jinjer API `POST /v1/kintai-imports` (type.id=5) で汎用データCSVインポート可能（公式確定）
- 実績確定申請の取消・再申請 API は提供されていない → **「承認保留」運用**で迂回
- 給与計算式は jinjer 側で完結（kintai勤怠 → kintai給与）
- 自動修正は危険を伴うため、**P6 のインポート用CSV生成前に人間レビューゲートを必須**で挟む（2026-05-22 追加）

---

## 1. 目的・背景

### 1.1 解こうとしている問題

| 現状の課題 | 原因 |
|---|---|
| 230人 × 形式バラバラの請求勤怠を突合するのに膨大な時間 | 月次バッチが組まれていない、画面アップロードで1ファイルずつ |
| 進捗管理が属人化、管理部で詰まりが見えない | 突合結果が12個に分散、誰がどの段階かが見えない |
| 修正→jinjerインポート作業に時間がかかる | 手動アップロード、実績確定申請の取消待ちが頻発 |
| 同じ差異の説明を毎月繰り返し書く | コメントの引き継ぎ機構がない |

### 1.2 ゴール

「請求勤怠フォルダにファイルが揃った状態でコマンド1発、月次マスター.xlsx が出来上がっており、管理部はそれを見て進捗管理できる」

ただし **jinjer への修正反映は完全自動化しない**。月次マスター内の「修正キュー」シートを介し、**人間が承認した行だけが jinjer インポート用CSVに乗る**。本設計のゴールは「自動化」ではなく「**自動準備 + 人間判断ゲート + 自動投入**」のパイプラインを成立させること。

#### 自動修正候補とするための運用方針（合意）

| 状態 | 扱い | インポート対象 |
|---|---|---|
| 差異なし | 修正不要（修正キューに乗せない）| - |
| 差異あり + コメントあり | **要チェック**（自動修正しない、人間判断必須）| 承認時のみ |
| 差異あり + コメントなし + 低リスク | **自動修正候補**（請求勤怠を正とする提案を生成）| 承認時のみ |
| 差異あり + コメントなし + **高リスク** | **要チェック**に降格（社員紐付け曖昧／データ欠損／パース失敗／日跨ぎ等）| 承認時のみ |

「自動修正候補」も含め、**人間の承認なしで jinjer に投入される行は1件も存在しない**。

---

## 2. 全体アーキテクチャ

```
[Z:\勤怠チェックシステム\kintai-checker\] (既存資産)
│
├── services/
│   ├── jinjer_api_client.py          ★拡張：在籍者・所属取得＋kintai-imports POST を追加
│   ├── jinjer_parser.py              既存：jinjer CSV パース
│   ├── timesheet_parser.py           既存：請求勤怠多形式パース
│   ├── shift_legend_parser.py        既存：シフト記号解決
│   ├── matcher.py                    既存：突合ロジック
│   ├── excel_exporter.py             既存：Excel出力
│   │
│   ├── employee_resolver.py          ★P0 新規：在籍者取得＋ファイル→社員紐付け
│   ├── master_aggregator.py          ★P1 新規：月次マスター生成＋修正キュー初期化
│   ├── triage.py                     ★P1 新規：差異行を auto_fix_candidate / needs_check に分類
│   ├── comment_carryover.py          ★P2 新規：前月コメント引き継ぎ＋再トリアージ
│   ├── batch_runner.py               ★P3 新規：バッチ実行オーケストレーション
│   ├── review_loader.py              ★P6 新規：修正キューシート読込＋承認状態バリデーション
│   └── kintai_import_builder.py      ★P6 新規：承認行から汎用データCSV生成＋POST 投入
│
├── batch_run.py                       ★P3 新規：CLI エントリ
├── import_run.py                      ★P6 新規：レビュー済キュー → jinjer インポート CLI
├── tests/
│   ├── test_employee_resolver.py     ★P0 新規
│   ├── test_master_aggregator.py     ★P1 新規
│   ├── test_triage.py                ★P1 新規
│   ├── test_comment_carryover.py     ★P2 新規
│   ├── test_batch_runner.py          ★P3 新規
│   ├── test_review_loader.py         ★P6 新規
│   └── test_kintai_import_builder.py ★P6 新規
└── docs/
    └── DESIGN_月次マスター_P0_P3.md   ★この文書（P5/P6 拡張を含む）
```

### データの流れ

```
[1] jinjer API
      ├─ 在籍者一覧 (GET /v1/employees)
      └─ 所属履歴   (GET /v1/employees/affiliations)
                ↓
           P0: employee_resolver
                ↓
       在籍者マスタ（社員番号・氏名・所属・除外フラグ）
                ↓
[2] 請求勤怠フォルダ (Y:\給与明細\R8年\<月>\請求勤怠データ\)
                ↓
           P3: batch_runner
              ├─ ファイル名/内容から氏名抽出
              ├─ 氏名 → 社員番号 マッチング（曖昧マッチ含む）
              ├─ ソース戦略の自動判定（ファイル名パターン/内容 sniffing）
              └─ ソース戦略別にグループ化して kintai-checker 突合を並列実行
                ↓
       突合結果xlsx (Y:\給与明細\R8年\<月>\突合結果\*.xlsx)
                ↓
[3] P1: master_aggregator + triage
       在籍者マスタ ⊃ ファイルから抽出された社員 を判定:
         ├─ ファイル抽出済 → match（突合済）
         ├─ ファイル無し & 除外グループ → skip（本社常駐等）
         └─ ファイル無し & 除外条件にも該当せず → 提出忘れ候補
       さらに差異あり行をトリアージ:
         ├─ コメントあり / 高リスク → needs_check（要チェック）
         └─ コメントなし & 低リスク → auto_fix_candidate（自動修正候補）
                ↓
       月次マスター_<YYYY-MM>.xlsx（仮、修正キューシート含む）
                ↓
[4] P2: comment_carryover
       前月マスターのコメントを引き継ぎ後、差異あり行を再トリアージ
       （引き継ぎコメントが付いた行は needs_check へ降格）
                ↓
       月次マスター_<YYYY-MM>.xlsx（完成版・全差異が未判断状態）
                ↓
[5] 人間レビュー（手作業 / P5 運用）
       Excel 上で 修正キューシートの「人間判断」列を埋める
         承認 / 却下 / 保留 / （空欄=未判断）
                ↓
[6] P6: kintai_import_builder
       修正キューを読み込み、人間判断=承認 の行のみを抽出
       → 汎用データCSVを生成（dry-run / 本実行）
       → POST /v1/kintai-imports
```

---

## 3. P0: 在籍者マスタ取得とファイル→社員紐付け（employee_resolver）

### 3.1 目的
**jinjer タグに依存せず**、以下を実現する:
1. jinjer から在籍者全員の社員番号・氏名・所属グループを取得
2. 請求勤怠フォルダ内のファイル群から、各ファイル/レコードに対応する社員番号を特定（氏名ベース）
3. 在籍者と「ファイルから抽出された社員」の差分を取り、突合区分を確定

突合区分は P0 単独では確定せず、**P1 で集約時に決まる** ことに注意:

| 状況 | 区分 |
|---|---|
| ファイルから抽出された | **match** |
| ファイルが無く、所属が除外グループ（本社常駐・F付き等）| **skip** |
| ファイルが無く、除外グループでもない | **submission_missing**（提出忘れ候補・要確認）|

### 3.2 入力API

| API | 用途 |
|---|---|
| `GET /v1/employees?enrollment-classification-id=0` | 在籍者の社員番号・氏名・カナ |
| `GET /v1/employees/affiliations` | 各人の最新所属グループ名（バッチ50件）|

タグ系API（`/v1/employees/kintai-tags`、`/v1/employees/tags`）は**使用しない**。

### 3.3 除外グループ判定

```python
SKIP_KEYWORDS = ['管理部', 'ビジネス企画本部', 'ディスパッチグループ']

def is_excluded_group(group_name: str) -> bool:
    """本社常駐など、請求勤怠が発生しないグループか"""
    if any(kw in group_name for kw in SKIP_KEYWORDS):
        return True
    if 'F' in group_name:  # 例: "FXXX", "XXXF" など F付き
        return True
    return False
```

### 3.4 氏名抽出ロジック（ファイル → 氏名候補）

```python
def extract_name_candidates(file_path: Path) -> list[str]:
    """ファイル名・ファイル内容から氏名候補を抽出"""
    candidates = []

    # 1. ファイル名から漢字氏名を抽出
    #    例: "勤務表2026年04月_大場.pdf" → "大場"
    #    例: "2026年04月勤務表_住吉健一.pdf" → "住吉健一"
    #    例: "中澤寿代さん(ITone)派遣労働者勤務報告書..." → "中澤寿代"
    name_from_filename = _extract_jp_name_from_filename(file_path.name)
    if name_from_filename:
        candidates.append(name_from_filename)

    # 2. アグリゲートファイル（SAP一括Excel等）はファイル内のシート/行から抽出
    #    既存 timesheet_parser.py のロジックを利用
    if _is_aggregate_file(file_path):
        candidates.extend(_extract_names_from_content(file_path))

    # 3. アルファベット名（外国籍）対応
    #    例: "timesheet_..._MaharjanRamita.pdf" → "MaharjanRamita"
    alpha_name = _extract_alpha_name_from_filename(file_path.name)
    if alpha_name:
        candidates.append(alpha_name)

    return candidates
```

### 3.5 氏名 → 社員番号マッチング

```python
def match_name_to_employee(name: str, employees: list[Employee]) -> MatchResult:
    """既存 normalize_name_keys を活用した名前マッチング"""
    normalized = normalize_name_keys(name)  # 既存実装（matcher.py）

    # 完全一致 → 1人 → 確定
    exact = [e for e in employees if any(k in e.name_keys for k in normalized)]
    if len(exact) == 1:
        return MatchResult(kind='matched', employee=exact[0])
    if len(exact) > 1:
        return MatchResult(kind='ambiguous', candidates=exact)

    # 部分一致（姓のみ等）
    partial = _fuzzy_match(name, employees)
    if len(partial) == 1:
        return MatchResult(kind='matched_fuzzy', employee=partial[0])
    if len(partial) > 1:
        return MatchResult(kind='ambiguous', candidates=partial)

    return MatchResult(kind='not_found')
```

**マッチ結果の扱い:**

| 種別 | match_confidence | 動作 | P1 トリアージへの影響 |
|---|---|---|---|
| `matched` | 1.0 | 確定。突合パイプラインに流す | 通常判定 |
| `matched_fuzzy` | 0.5〜0.9（具体値はマッチ手段による）| 確定だが「あいまいマッチ」フラグ付与、月次マスター上で目印表示 | **自動的に needs_check に降格**（社員紐付けが曖昧なため）|
| `ambiguous` | 0.0 | **アラート**。複数候補を「取込ログ」シートに表示、人手で確定 | 突合自体保留、修正キューに乗らない |
| `not_found` | -- | **アラート**。「未紐付けファイル」として「取込ログ」に記録 | 突合自体保留、修正キューに乗らない |

`MatchResult` には `kind` だけでなく `confidence: float`（0.0〜1.0）を持たせる。トリアージ側は `confidence < FUZZY_THRESHOLD (=1.0)` で降格判定するため、将来「カナで絞り込んで 1 名確定したが信頼度 0.8」のような中間ケースが出てきても閾値を調整するだけで対応できる。

`matched_fuzzy` の社員に紐付いた差異は、たとえコメントなし・低リスク条件を満たしても自動修正候補にせず、必ず人間判断を経由させる（P1 §4.7 トリアージロジック参照）。

### 3.6 出力

`Y:\給与明細\R8年\<月>\employees_<YYYYMM>.xlsx`（中間ファイル）

| 社員番号 | 氏名 | カナ | 所属グループ | 除外グループ判定 | 名前突合用キー |
|---|---|---|---|---|---|
| 2011013 | 大貫 崇 | オオヌキ タカシ | NMHT〇〇 | False | `大貫崇`, `おおぬきたかし` |
| 2017022 | 渡辺 聖 | ワタナベ ヒジリ | ホンダ芳賀 | False | `渡辺聖`, `わたなべひじり` |
| 2099001 | 田中 一郎 | タナカ イチロウ | 管理部 | True | - |

「除外グループ判定」は P1 で submission_missing 判定に使う。「名前突合用キー」は P3 のマッチング高速化のため事前計算しておく。

### 3.7 実装ファイル

- `services/employee_resolver.py` — 在籍者取得＋氏名抽出＋マッチング
- `services/jinjer_api_client.py` — `get_employees_with_affiliation()` を追加
- `tests/test_employee_resolver.py` — マッチングロジックの単体テスト（同姓同名、ゆれ、外国籍名）

### 3.8 工数

**既存資産（流用前提で再見積り）:**
- `services/jinjer_api_client.py` に在籍者取得・所属履歴・氏名→IDマップ・打刻グループ取得が **既に実装済み**
- `services/matcher.py` の `normalize_name_keys` も活用可

**新規に書くもの:**
- 除外グループ判定（`SKIP_KEYWORDS` + F付き判定）— 数十行
- ファイル名/内容からの氏名候補抽出（漢字氏名・アルファベット名の正規表現、アグリゲートファイル対応）
- マッチング結果の `exact / matched_fuzzy / ambiguous / not_found` 分岐 + `MatchResult` データクラス
- 出力 Excel `employees_<YYYYMM>.xlsx` の生成

実装: **1〜1.5日**（旧見積 2〜3日 から短縮）

---

## 4. P1: 月次マスター集約

### 4.1 目的
kintai-checker の突合結果xlsx（複数）と P0 の社員区分を統合し、**全員1冊の月次マスター.xlsx** を生成。

### 4.2 入力

- `Y:\給与明細\R8年\<月>\突合結果\*.xlsx` — kintai-checker 出力（複数）
- `employees_<YYYYMM>.xlsx` — P0 出力（在籍者マスタ）
- `不明ファイル_<YYYYMM>.xlsx` — P0/P3 で社員特定できなかったファイルのログ

※ jinjer の当月勤怠サマリは取得しない（Phase C で確定。skip 者を含む全員分の取得は API 負荷大きすぎ＆運用上不要）。突合結果に含まれる「jinjer側の勤務時間」等を流用する

### 4.3 シート構成

| シート名 | 内容 | 想定行数 |
|---|---|---|
| **サマリ** | ステージ別人数、突合区分内訳、submission_missing 件数、ambiguous 件数、修正キューの判断状況（未判断/承認/却下/保留 件数）、**却下未通知件数**、**保留3ヶ月以上滞留件数** | 数十行 |
| **詳細** | 1社員=1行、全 230 人 | 230 |
| **日次突合結果** | 1社員×1日=1行、勤務/休憩/実働の日次差異 | 230 × 30 ≒ 6,900 |
| **修正キュー** ★再設計 | 差異あり行を **auto_fix_candidate / needs_check** に分類し、人間判断（承認/却下/保留）と判断メモを記録する **レビューゲート用シート**。月次・日次の両差異が混在 | 動的 |
| **コメント履歴** | 前月から引き継いだコメント（P2 で書き込み）| 動的 |
| **取込ログ** | ファイル→社員紐付け結果、ambiguous/not_found ファイル、突合粒度（月次のみ/日次あり）| 動的 |
| **提出忘れ候補** | 「除外グループでない」かつ「ファイル未発見」の社員リスト | 動的 |
| **インポート履歴** ★新規 | P6 で kintai-imports を実行した日時・件数・誰が承認したか・jinjer 側応答の監査ログ | 動的 |

### 4.4 「詳細」シートの列設計

| 列 | 内容 | データ源 |
|---|---|---|
| 社員番号 | キー | P0 |
| 氏名 | 表示用 | P0 |
| 所属グループ | 表示用 | P0 |
| **区分** | match / skip / submission_missing / ambiguous | P1 集約時に確定 |
| **ソース推定** | ファイル名/内容から推定したソース（例: SAP_Fieldglass）| kintai-checker |
| **突合粒度** | `日次+月次` / `月次のみ` / `突合不可` | kintai-checker |
| **ステージ** | 0〜7 | 計算 |
| **判定種別** | OK / NG / 要確認 / データ欠損（P2 引き継ぎキーに使用）| kintai-checker (matcher.judge) |
| 月次差異件数 | 月合計の差異 | kintai-checker |
| 日次差異件数 | 日次突合で出た差異の合計件数 | P1 集計 |
| 休憩差異件数 | うち休憩時間に起因する差異件数 | P1 集計 |
| 最大差分(分) | 日次差異のうち最大の絶対値 | P1 集計 |
| 日次突合不可理由 | 日次データが取れなかったソースの場合に理由を記載 | P1 |
| **マッチ精度** | exact / fuzzy / ambiguous | P0 マッチング結果 |
| **修正キュー件数(auto)** ★新規 | 当該社員の auto_fix_candidate 行数 | P1 トリアージ |
| **修正キュー件数(check)** ★新規 | 当該社員の needs_check 行数 | P1 トリアージ |
| **未判断件数** ★新規 | 修正キュー行のうち人間判断が空欄の件数（P6 ゲートで警告対象）| P1/P6 |
| 差異内訳 | 詳細メッセージ（参考表示用、引き継ぎキーには使用しない）| kintai-checker |
| コメント | 人が記入 / P2で引き継ぎ | 手動 + P2 |
| 修正方針 | 「請求勤怠に合わせて修正」など | 手動 |
| 元突合結果ファイル | リンク | P1 |
| 最終更新 | タイムスタンプ | 自動 |

### 4.5 「日次突合結果」シートの列設計

1行 = (社員番号, 日付)。同一社員の日数分の行が連続する。月次のみ突合の場合はこのシートに行が生成されない。

| 列 | 内容 |
|---|---|
| 社員番号 | キー |
| 氏名 | 表示用 |
| 日付 | キー（YYYY-MM-DD）|
| 請求勤怠_勤務時間 | 始業〜終業の総拘束時間 |
| jinjer_勤務時間 | 同上、jinjer 側 |
| 請求勤怠_休憩時間 | 休憩合計 |
| jinjer_休憩時間 | 同上、jinjer 側 |
| 請求勤怠_実働時間 | 勤務 - 休憩 |
| jinjer_実働時間 | 同上、jinjer 側 |
| **差分(分)** | jinjer_実働 - 請求勤怠_実働（マイナス=jinjer側少ない）|
| **差異理由候補** | 「休憩打刻ミス」「日跨ぎ」「未打刻」など、ヒューリスティック判定 |
| **トリアージ区分** ★再定義 | `no_diff` / `auto_fix_candidate` / `needs_check`（P1 §4.7 で確定）|
| **要チェック理由** ★新規 | `needs_check` の場合の理由コード（カンマ区切り複数可）。値域: `comment_present` / `fuzzy_match` / `monthly_only` / `parse_failed` / `data_missing` / `day_crossing` / `submission_missing` / `unknown_reason` |
| 修正キューID | 修正キューシートへの参照（後方リンク）|
| コメント | 人が記入 |

#### 差異理由候補の自動推定ロジック

| 条件 | 推定理由 | 既定トリアージ（理由コード）|
|---|---|---|
| 休憩時間に差があり、勤務時間は一致 | **休憩打刻ミス** | auto_fix_candidate |
| 勤務時間が大幅に短い（jinjer側）| **未打刻** または **早退/遅刻** | auto_fix_candidate |
| 勤務時間が深夜帯に偏る | **日跨ぎシフトの集計位置ズレ** | needs_check (`day_crossing`) |
| 請求勤怠側が空・jinjer側に値あり | **請求勤怠未提出** | needs_check (`submission_missing`) |
| パース時に必須フィールドが None / 空 | **データ欠損** | needs_check (`data_missing`) |
| パーサーが例外を投げた | **パース失敗** | needs_check (`parse_failed`) |
| 月次のみソース由来の差異 | **粒度不足** | needs_check (`monthly_only`) |
| 社員紐付けが fuzzy で確定 | **紐付け曖昧** | needs_check (`fuzzy_match`) |
| 差分が ±5分以内 | **誤差許容範囲** | no_diff（修正キューに乗せない）|
| 上記いずれにも該当しない | **不明** | needs_check (`unknown_reason`) |
| （ヒューリスティック判定の上に）当月コメントあり / 前月引き継ぎコメントあり | **既往説明あり** | needs_check (`comment_present`) を **重ねて付与**（複数 reason 可）|

reason コード一覧（needs_check 時の `要チェック理由` 列）: `fuzzy_match / monthly_only / parse_failed / data_missing / day_crossing / submission_missing / unknown_reason / comment_present`

### 4.6 「修正キュー」シートの列設計 ★新規

修正キューシートは**人間レビューゲート**の実体。P1 で生成され、人間が Excel 上で「人間判断」列を埋めるまで P6（CSV 生成）に進めない。

| 列 | 内容 | 値域 / 例 |
|---|---|---|
| 修正キューID | 一意キー `<社員番号>_<YYYY-MM-DD or YYYY-MM>_<差異種別>` | `2017022_2026-05-12_break` |
| 社員番号 | キー | `2017022` |
| 氏名 | 表示用 | `渡辺 聖` |
| マッチ精度 | exact / fuzzy（fuzzy は強制 needs_check）| `exact` |
| 粒度 | 月次 / 日次 | `日次` |
| 対象日付 | 日次は YYYY-MM-DD、月次は YYYY-MM | `2026-05-12` |
| 差異種別 | `勤務時間` / `休憩時間` / `実働時間` / `月合計` | `休憩時間` |
| 請求勤怠値 | パース値（分） | `60` |
| jinjer値 | jinjer 側値（分） | `45` |
| 差分(分) | jinjer - 請求勤怠 | `-15` |
| 差異理由候補 | §4.5 のヒューリスティック判定結果 | `休憩打刻ミス` |
| **トリアージ区分** | `auto_fix_candidate` / `needs_check` | `auto_fix_candidate` |
| **要チェック理由** | needs_check の場合のコード（複数可カンマ区切り）| `comment_present,fuzzy_match` |
| **自動修正提案値** | 請求勤怠値（=請求勤怠を正とする提案。承認時にこの値で CSV 生成）| `60` |
| 既存コメント | 当月手入力 + P2 引き継ぎ | `毎月昼休憩45分の運用` |
| **人間判断** ★新規 | 空欄 / `承認` / `却下` / `保留` （Excel データ検証でプルダウン化）| 空欄 |
| **判断メモ** ★新規 | 人手記入欄 | `本人確認済み` |
| 判断者 | 任意。Excel ユーザー名で自動入力可 | `谷津` |
| 判断日時 | 「人間判断」列が変わった最後のタイムスタンプ（任意）| `2026-05-22 14:33` |
| ステータス | `pending` / `reviewed` / `exported` / `imported` | `pending` |
| **却下通知ステータス** ★新規 | 却下行の請求勤怠側修正依頼の通知状況。`不要`（承認/保留/未判断）/ `未通知`（却下直後の初期値）/ `通知済` | `未通知` |
| **通知先** ★新規 | 却下時に請求勤怠側を修正依頼する相手（派遣会社名・担当者名など）| `ITone 山田` |
| **通知日時** ★新規 | 通知済に切替えたタイムスタンプ | `2026-05-23 11:00` |
| 元突合行リンク | 「日次突合結果」シートの行番号へのリンク | `=HYPERLINK(...)` |
| 最終更新 | 自動 | `2026-05-22 09:00` |

#### 却下通知ステータスの遷移ルール

| 「人間判断」 | 「却下通知ステータス」初期値 | 説明 |
|---|---|---|
| 空欄 / 承認 / 保留 | `不要` | 通知対象外。Excel データ検証で `却下` 以外の場合は灰色表示推奨 |
| 却下 | `未通知` | 担当者が請求勤怠側（派遣会社・本人）への連絡を行ったら手動で `通知済` に切替 |

P6 のレビューゲートで「却下 かつ 通知ステータス = 未通知」が残っている場合は警告のみ（インポート自体は阻害しない、§10.4 参照）。「却下行の宙浮き」を翌月レビュー時に管理部が拾えるよう、サマリシートに「却下未通知件数」を表示する。

#### 「人間判断」値の意味（運用合意済）

| 値 | 意味 | P6 CSV への扱い | 翌月への持ち越し |
|---|---|---|---|
| 空欄（未判断）| まだ誰も見ていない | **除外**（CSV 未生成）。P6 ゲートで警告対象 | 翌月マスターに未解決として再表示 |
| **承認** | 自動修正提案値（請求勤怠値）で jinjer を修正してよい | **CSV 出力対象** | 完了。履歴に残る |
| **却下** | 請求勤怠側が誤り。jinjer 側の値を正とする（修正不要）| **除外** | 完了。翌月引き継がない |
| **保留** | 判断保留。今月は触らず来月以降に再判断 | **除外** | **翌月マスターに未解決として再表示**し、コメントを引き継ぎ |

#### auto_fix_candidate の承認 UX（運用補助）

Excel 上で大量の auto_fix_candidate を 1 件ずつクリックするのを助ける補助を用意するが、**「フィルタ全可視行を一括承認」ボタンは原則提供しない**（誤承認による誤投入のリスクを優先）:

- 修正キューシート先頭にフィルタ + 「人間判断」列にプルダウン（承認 / 却下 / 保留）
- 行高を広めにとり、目視で 1 件あたり 5〜10 秒で「自動修正提案値」と「jinjer 値」の差を確認できるレイアウト
- フィルタ条件例: `トリアージ区分 = auto_fix_candidate` AND `マッチ精度 = exact`
- ワンクリック承認マクロは将来オプション扱い（§9.2 末尾参照）

### 4.7 トリアージロジック ★新規

差異あり行を `auto_fix_candidate` / `needs_check` のいずれかに分類するルールセット。**順番に評価し、最初にヒットした条件で確定**する（早期 needs_check への降格を優先）。

```python
FUZZY_THRESHOLD = 1.0  # match_confidence がこの値未満なら fuzzy 扱い
ALLOW_DIFF_MINUTES = 5

def triage(row, match_confidence: float) -> TriageResult:
    """差異あり行をトリアージする。

    Args:
        row: 突合済み 1 行（差異情報・コメント・粒度・パース状態を持つ）
        match_confidence: P0 employee_resolver が返した社員紐付けの信頼度 (0.0〜1.0)

    Returns:
        TriageResult(kind, reasons)
        kind: 'no_diff' | 'auto_fix_candidate' | 'needs_check'
        reasons: needs_check の場合の理由コード list
    """
    # ① 差異なし（誤差許容範囲含む）はそもそも修正キューに乗らない
    if not row.has_diff or abs(row.diff_minutes) <= ALLOW_DIFF_MINUTES:
        return TriageResult(kind='no_diff', reasons=[])

    reasons = []

    # ② 社員紐付けが曖昧 → 必ず needs_check
    if match_confidence < FUZZY_THRESHOLD:
        reasons.append('fuzzy_match')

    # ③ 月次のみ突合（日次データなし）→ 必ず needs_check
    if row.granularity == 'monthly_only':
        reasons.append('monthly_only')

    # ④ パース失敗 → 必ず needs_check
    if row.parse_failed:
        reasons.append('parse_failed')

    # ⑤ データ欠損（請求勤怠 or jinjer 側のフィールドが None / 空）→ 必ず needs_check
    if row.has_missing_data:
        reasons.append('data_missing')

    # ⑥ 日跨ぎシフト・深夜帯偏り → 必ず needs_check
    if row.diff_reason == '日跨ぎシフトの集計位置ズレ':
        reasons.append('day_crossing')

    # ⑦ 請求勤怠側が空 → 提出忘れの可能性、必ず needs_check
    if row.diff_reason == '請求勤怠未提出':
        reasons.append('submission_missing')

    # ⑧ 差異理由が「不明」→ ヒューリスティック判定不能、念のため needs_check
    if row.diff_reason == '不明':
        reasons.append('unknown_reason')

    # ⑨ コメントあり（手入力 or 前月引き継ぎ）→ needs_check
    if row.comment.strip():
        reasons.append('comment_present')

    if reasons:
        return TriageResult(kind='needs_check', reasons=reasons)

    # ⑩ 上記いずれにも該当しない → auto_fix_candidate
    return TriageResult(kind='auto_fix_candidate', reasons=[])
```

**設計意図:**
- 「コメントあり」だけでなく、**リスクが高いケースはコメントなしでも needs_check に降格**（fuzzy_match / monthly_only / parse_failed / **data_missing** / day_crossing / submission_missing / unknown_reason）
- 「パース失敗」と「データ欠損」は別物として理由コードを分離（parse_failed = ファイル形式・OCRの失敗 / data_missing = パースは成功したが必須フィールドが空）
- 社員紐付け信頼度は float（match_confidence）で扱い、`FUZZY_THRESHOLD` の閾値調整で「将来カナで絞り込めたが信頼度 0.8 の中間ケース」にも対応可能
- 自動修正候補は「請求勤怠側のデータが信頼でき、jinjer 側の差分理由もヒューリスティックで説明できる」ケースに限定
- 設定 `ALLOW_DIFF_MINUTES`（既定 5 分）で誤差許容範囲を調整可能

### 4.8 ステージ定義

```
ステージ 0: 突合不要 (skip 区分の本社常駐者)
ステージ 1: jinjer 勤怠取得済（match 区分のみ）
ステージ 2: 突合実行済
ステージ 2.5: トリアージ実行済（auto_fix_candidate / needs_check に分類済）
ステージ 3a: 全差異が修正キューに登録済、未判断あり（人間レビュー待ち）
ステージ 3b: 修正キュー全行が判断済（承認/却下/保留が全て埋まっている）
ステージ 4: 承認行から修正データCSV生成済（P6 dry-run 通過）
ステージ 5: jinjer インポート済（POST /v1/kintai-imports 成功）
ステージ 6: 給与計算待ち
ステージ 7: 完了（実績確定申請の承認まで）
```

- **skip 区分は ステージ 0 のまま**：jinjer サマリも取得しない。詳細シートには「区分=skip / ステージ=0 / 突合不要」とだけ表示
- **ステージ 2 → 2.5** は P1 master_aggregator の triage モジュールで自動進行
- **ステージ 2.5 → 3a → 3b** は人間レビュー（P5 運用）で手動進行。3a は判断空欄が残っている状態
- **ステージ 3b → 4 → 5** は P6 import_run.py で進行。3b 未達でも `--force-incomplete` で部分実行可能だが警告
- 今回 P0〜P3 では **match 区分のステージ 1〜2.5 の自動化** が範囲。3a/3b/4/5 は P5/P6 の責務

### 4.9 出力

`Y:\給与明細\R8年\<月>\月次マスター_<YYYY-MM>.xlsx`

### 4.10 実装ファイル

- `services/master_aggregator.py` — シート組み立て
- `services/triage.py` — §4.7 のトリアージロジック
- `tests/test_master_aggregator.py`
- `tests/test_triage.py` — ルール ②〜⑨ の各分岐 + 優先順位テスト

### 4.11 工数

P1 はシート構成・トリアージ判定・修正キュー生成が本体で、既存資産に直接該当するものは少ない。

- 実装: **4〜5日**（修正キューシート + triage モジュール + openpyxl `DataValidation` プルダウン + サマリ集計）
- 内訳目安: master_aggregator 1.5日 / triage 1日 / シート整形(openpyxl DataValidation/auto_filter/色塗り) 1.5日 / テスト 1日

---

## 5. P2: 前月コメント引き継ぎ

### 5.1 目的
前月マスターに書かれたコメントを、今月マスターの該当行（同じ差異理由）に自動転記。

### 5.2 引き継ぎキー

**Phase C 確定:** 粗い粒度（判定種別のみ）でハッシュ化する。

```python
def carryover_key(row):
    """社員番号 + 判定種別"""
    return (row['社員番号'], row['判定種別'])
```

判定種別は matcher.py の `judge()` 関数の出力に揃える:
- `OK` / `NG` / `要確認` / `データ欠損` の4種

引き継ぎ動作:
- 同じ社員で前月と判定種別が同じ → 前月コメントを今月にコピー
- 例: 「この人は毎月『送遅』気味でNG扱い」のコメントが毎月引き継がれる
- 引き継がれたコメントには `[前月引き継ぎ]` マーク付与（手動更新できるよう保護はしない）

**意図:** 細かい粒度（日付・時刻まで含める）だと「同じパターン」と認識されず引き継がれない問題があるため、まず粗い粒度で「引っ掛けやすさ」を優先。粗すぎる場合は運用で粒度を上げる。

### 5.3 引き継ぎポリシー

```python
def carryover(prev_master, curr_master):
    prev = read_comments_indexed_by(prev_master, carryover_key)
    for row in curr_master.detail_rows:
        key = carryover_key(row)
        if key in prev and not row['コメント']:
            row['コメント'] = prev[key]['comment']
            row['コメント'] += f'\n[前月 {prev_master.month} から引き継ぎ]'
            log_to_history_sheet(curr_master, row, prev[key])
```

**ルール:**
- 今月コメントが空 かつ 前月に同キーのコメントあり → 転記
- 今月コメントが手動記入済み → 上書きしない
- 引き継いだコメントには `[前月引き継ぎ]` マーク付与
- 「コメント履歴」シートに転記履歴を残す

### 5.4 修正キューへの保留行引き継ぎ ★新規

P2 では、前月マスターの **修正キューシート** から「人間判断 = 保留」だった行を抽出し、当月マスターの修正キューに再投入する。

```python
def carryover_held_review_items(prev_master, curr_master):
    prev_held = [
        row for row in prev_master.review_queue
        if row['人間判断'] == '保留'
    ]
    for row in prev_held:
        if not still_relevant_in_current_month(row, curr_master):
            continue  # 該当社員が当月退職等で対象外なら捨てる
        curr_master.review_queue.append({
            **row,
            '人間判断': '',  # リセット
            '判断メモ': f"{row['判断メモ']}\n[前月 {prev_master.month} 保留分]",
            'ステータス': 'pending',
        })
```

### 5.5 引き継ぎ後の再トリアージ ★新規

コメント引き継ぎを行うと、**前月コメントが付いた行は当月コメントが空でも実質コメント付き状態**になる。
P2 の末尾で全差異行を再トリアージし、新たに `comment_present` 理由が付いた行は `auto_fix_candidate → needs_check` に降格させる。

```python
def re_triage_after_carryover(curr_master):
    for row in curr_master.review_queue:
        if row['人間判断']:
            continue  # 既に人間が判断済みの行は触らない（保留含む）
        new_result = triage(row, employee_match_kind=row['マッチ精度'])
        if new_result.kind != row['トリアージ区分']:
            row['トリアージ区分'] = new_result.kind
            row['要チェック理由'] = ','.join(new_result.reasons)
```

**狙い:** コメント引き継ぎは「過去に説明済みの差異」を示すため、自動修正せず必ず人間に確認させる。

### 5.6 実装ファイル

- `services/comment_carryover.py` — §5.3 + §5.4
- `services/triage.py` の `re_triage_after_carryover()` を呼び出し
- `tests/test_comment_carryover.py` — 保留引き継ぎ + 再トリアージ降格のテスト

### 5.7 工数

- 実装: 2〜3日（保留引き継ぎ + 再トリアージ分で +1日）

---

## 6. P3: バッチ実行モード

### 6.1 目的
請求勤怠フォルダを指定して **コマンド1発** で P0→突合→P1→P2 まで全自動。

### 6.2 CLI 仕様

```powershell
cd Z:\勤怠チェックシステム\kintai-checker

python batch_run.py `
  --month        2026-05 `
  --billing-dir  "Y:\給与明細\R8年\5月\請求勤怠データ" `
  --output-dir   "Y:\給与明細\R8年\5月\突合結果" `
  --master-out   "Y:\給与明細\R8年\5月\月次マスター_2026-05.xlsx" `
  --prev-master  "Y:\給与明細\R8年\4月\月次マスター_2026-04.xlsx" `
  --workers      4
```

### 6.3 処理フロー

```
[Step 1] P0 実行
  ├─ jinjer 在籍者・所属取得 → employees_2026-05.xlsx
  └─ 除外グループ判定済みの在籍者マスタ生成

[Step 2] 請求勤怠フォルダをスキャン
  └─ 全ファイルリストを作成

[Step 3] ファイル → 氏名候補抽出
  ├─ ファイル名から氏名抽出
  ├─ アグリゲートファイルは内容から抽出（既存 timesheet_parser 利用）
  └─ 抽出失敗ファイルは「取込ログ」に not_found として記録

[Step 4] 氏名 → 社員番号マッチング
  ├─ exact / fuzzy / ambiguous / not_found を判定
  ├─ ambiguous は「取込ログ」に複数候補表示
  └─ matched/matched_fuzzy のみ突合パイプラインに流す

[Step 5] ソース戦略の判定
  ├─ ファイル名パターン（"SAP" "TCD" "e-staffing" 等）
  ├─ ファイル内容 sniffing（ヘッダー文字列、列構成）
  └─ 既存 timesheet_parser.parse_timesheet_smart のモード自動判定を活用

[Step 6] ソース戦略別グルーピング & 並列突合
  └─ ProcessPoolExecutor(max_workers=N)
      ├─ 戦略A（例: SAP月次）: kintai-checker 突合 → 月次のみ突合結果
      ├─ 戦略B（例: Estaffing日次）: kintai-checker 突合 → 月次＋日次
      └─ ...

[Step 7] P1: 月次マスター集約
  ├─ 在籍者マスタ ⊃ 突合済み社員 を判定して区分確定
  │   ├─ 突合済み → match
  │   ├─ 未突合 & 除外グループ → skip
  │   └─ 未突合 & 除外でない → submission_missing（提出忘れ候補）
  ├─ 「詳細」シート生成
  ├─ 「日次突合結果」シート生成
  ├─ 「提出忘れ候補」シート生成
  └─ 月次マスター_2026-05.xlsx 生成（修正キューはまだ空）

[Step 7.5] P1: トリアージ ★新規
  ├─ services/triage.py で全差異行を auto_fix_candidate / needs_check に分類
  ├─ 「修正キュー」シートに登録（人間判断列は空欄）
  ├─ 「日次突合結果」シートにトリアージ区分 + 要チェック理由を反映
  └─ 「詳細」シートの修正キュー件数(auto/check)・未判断件数 を集計

[Step 8] P2: 前月コメント引き継ぎ + 再トリアージ
  ├─ 前月コメントを「詳細」シートに転記
  ├─ 前月修正キューの「保留」行を当月修正キューに再投入
  └─ 引き継ぎコメントが付いた auto_fix_candidate 行を needs_check へ降格

[Step 9] 完了通知
  ├─ コンソールにサマリ出力（match/skip/submission_missing/ambiguous 人数）
  ├─ 修正キュー件数（auto_fix_candidate / needs_check / 保留引き継ぎ）を表示
  ├─ 「未判断件数」と「人間レビュー必須」を強調表示
  └─ 月次マスターの保存先パスを表示
```

### 6.4 ソース戦略の判定（フォルダ名一次キー + 中身ガード）★改訂7

**運用前提（2026-05-26 確定）**: ユーザーが請求勤怠を **ソース別サブフォルダに事前格納する**。これにより `SOURCE_STRATEGIES` のファイル名/中身 sniffing による一次判定は不要となり、フォルダ名 → ソース戦略の単純マップで決定する。中身の列名チェックは「誤格納の検知」のためのガードとして残す（ファイルは弾く、エラーログに記録）。

**運用フォルダ構成（推奨）**:

```
Y:\給与明細\R8年\<月>\請求勤怠データ\
├── SAP_Fieldglass\      ← 「スタッフ・時間エントリ日…」14列の Excel
├── e-staffing\          ← TCD で始まる Excel、または e-staffing 契約No 列の Excel
├── Excel\               ← 個別 Excel（DELL、ホンダ等の独自フォーマット）
└── その他\              ← PDF・画像個別、ERCSTS、判別困難なもの（既存 parse_timesheet_smart にフォールバック）
```

```python
SOURCE_STRATEGIES = {
    'SAP_Fieldglass': {
        # 「スタッフ・時間エントリ日・出勤時刻・終了時刻・食事休憩・タイムシートID」の14列構造
        'parser': 'sap_fieldglass_daily',
        'granularity': 'daily_and_monthly',  # 1人複数行、日次あり
        'name_format': 'lastname_comma_firstname',  # 例: "上原, 奏吾"
        'expected_headers': ['スタッフ', '時間エントリ日', 'タイムシート ID'],  # ガード用
    },
    'e-staffing': {
        # ヘッダー: e-staffing契約No, スタッフ氏名, 就業年月日, 区分, 開始時刻, 終了時刻, 休憩時間, ...
        'parser': 'estaffing_daily',
        'granularity': 'daily_and_monthly',
        'name_format': 'lastname_space_firstname',  # 例: "寺山 枝美"
        'expected_headers': ['e-staffing契約No', '就業年月日', 'スタッフ氏名'],
    },
    'Excel': {
        # 個別 Excel フォーマット。既存 parse_timesheet_smart のモード判定を流用
        'parser': 'generic_excel',
        'granularity': 'daily_and_monthly',
        'expected_headers': None,  # ガードなし（多形式すぎるため）
    },
    'その他': {
        # PDF / 画像 / ERCSTS など。拡張子で個別ディスパッチ
        'parser': 'auto_dispatch',  # .pdf → pdf_ocr / .jpg|.png → image_ocr / .xlsx → generic_excel
        'granularity': 'daily_and_monthly',
        'expected_headers': None,
    },
}
```

#### 判定アルゴリズム

```python
def resolve_source_strategy(file_path: Path, base_dir: Path) -> SourceStrategy | None:
    """ファイルのフォルダ名からソース戦略を引く。中身ガードに失敗した場合は None を返し
    呼び出し側で「取込ログ」に警告を出す。"""
    # 1. フォルダ名で一次判定
    rel = file_path.relative_to(base_dir)
    folder_name = rel.parts[0] if len(rel.parts) > 1 else None
    if folder_name not in SOURCE_STRATEGIES:
        log_warning(f"未知のソースフォルダ: {folder_name} / file: {file_path}")
        return None
    strategy = SOURCE_STRATEGIES[folder_name]

    # 2. 中身ガード（expected_headers が指定されているもののみ）
    if strategy['expected_headers']:
        actual_headers = _read_top_headers(file_path)  # Excel/CSV の1〜数行目を読む
        missing = [h for h in strategy['expected_headers'] if h not in actual_headers]
        if missing:
            log_warning(
                f"誤格納の疑い: {file_path.name} は {folder_name} フォルダに格納されているが、"
                f"期待されるヘッダー {missing} が見つかりません。「取込ログ」に記録します。"
            )
            return None  # 処理スキップ → 取込ログへ

    return strategy
```

**狙い:**
- ユーザーが事前にフォルダ分けすることで、`detect_by_filename` の正規表現メンテと `detect_by_content_header` の sniffing ロジックが不要に
- 中身ガード（`expected_headers`）で「SAP_Fieldglass フォルダに e-staffing ファイルを入れた」等の誤格納を検知
- 「その他」フォルダは拡張子で個別ディスパッチし、既存 `parse_timesheet_smart` のフォールバックモード判定にも委譲

**月次のみソースに該当する社員は、月次マスターの:**
- 「詳細」シート → `突合粒度=月次のみ` + `日次突合不可理由` 列に理由を記載
- 「日次突合結果」シート → 該当社員の行は**生成しない**
- 「サマリ」シート → 月次のみ突合人数を可視化

### 6.5 並列化

- ソース戦略単位で並列実行
- ボトルネックは Claude Vision OCR（PDF/画像処理）→ 並列化効果大
- `--workers` で同時実行数指定（デフォルト 4）

### 6.6 エラーハンドリング

| エラー種別 | 動作 |
|---|---|
| ファイル → 氏名抽出失敗 | 「取込ログ」シートに `not_found` で記録、処理続行 |
| 氏名 → 社員番号 マッチで `ambiguous` | 「取込ログ」に複数候補表示、ステージ「人手介在要」、処理続行 |
| ソース戦略の自動判定失敗 | 既存 `parse_timesheet_smart` のフォールバックモード判定へ、最終的に失敗なら未処理ログ |
| パーサー例外 | 該当社員をステージ「失敗」に、処理続行。差異行は **needs_check (parse_failed)** で修正キュー登録 |
| jinjer API タイムアウト | 3回リトライ、それでも失敗ならステージ「失敗」 |
| 月次マスター書き込み失敗 | 致命的、中断 |
| トリアージ実行中の例外 | 該当行を needs_check (unknown_reason) に強制降格、処理続行 |

### 6.7 実装ファイル

- `batch_run.py` — CLI エントリ（argparse）
- `services/batch_runner.py` — オーケストレーション本体
- `tests/test_batch_runner.py` — 統合テスト

### 6.8 工数

**既存資産:**
- 既存 `parse_timesheet_smart` のフォールバックモード判定・既存 kintai-checker の突合パイプラインを呼び出すだけ
- 並列化は `concurrent.futures.ProcessPoolExecutor` の薄いラッパー

**運用前提による簡略化（改訂7）:**
- §6.4 のソース戦略を「フォルダ名一次キー + 中身ガード」に変更したことで、`detect_by_filename` / `detect_by_content_header` の自動判定ロジックとそのテストが不要に
- 「その他」フォルダのみ拡張子ディスパッチ + 既存 `parse_timesheet_smart` のフォールバックに委譲

実装: **1.5〜2日**（旧見積 2〜3日 から更に短縮）— フォルダ→戦略マップ + 中身ガード + P0/P1/P2 直列呼び出し + 並列化ラッパー。

---

## 7. 工数まとめ（2026-05-26 再精査・既存資産流用前提）

### 7.1 既存資産棚卸し

| 資産 | 場所 | カバー範囲 |
|---|---|---|
| `JinjerClient`（認証・在籍者・所属履歴・打刻グループ・氏名→IDマップ）| `kintai-checker/services/jinjer_api_client.py` | P0 の 70% |
| `parse_timesheet_smart` / 各種パーサ | `kintai-checker/services/` | P3 の 60% |
| `request_with_retry`（429リトライ・`Retry-After`解釈）/ dry-run→本実行フロー | `Z:/API連携/update_salary_unit_prices.py`, `change_attendance_group.py` | P6 の通信基盤 50% |

### 7.2 各 Phase 再見積り

| Phase | 旧見積 | 新見積（流用後）| 主な圧縮要因 |
|---|---|---|---|
| P0 employee_resolver | 2〜3日 | **1〜1.5日** | jinjer API クライアント既存・氏名→IDマップ既存 |
| P1 master_aggregator + triage | 4〜5日 | **4〜5日**（据置）| シート構成・triage が本体、流用先なし |
| P2 comment_carryover + 再トリアージ | 2〜3日 | **2〜3日**（据置）| 6月最小構成からは省略可（初月のため）|
| P3 batch_runner | 3〜5日 | **1.5〜2日** | 既存 parse_timesheet_smart の薄いラッパー **+ ユーザーがソース別フォルダで事前格納する運用を採用したため、ソース自動判定ロジック不要（改訂7）** |
| P5 運用整備（**openpyxl `DataValidation` で実装、VBAマクロではない**）| 1日 | **0.5日** | P1 のシート生成内に組み込み、運用ガイド文書化のみ |
| P6 kintai_import_builder | 4〜6日 | **3〜4日** | 429リトライ・dry-run フレームワーク流用。kintai-imports POST のフォーマット検証・ジョブポーリング・監査ログは新規 |

### 7.3 累計

| 構成 | 工数 |
|---|---|
| **フル（P0+P1+P2+P3+P5+P6）** | **12〜16 日**（旧 16〜23 日 → 改訂6 12.5〜17日 → 改訂7 で更に圧縮）|
| **本格最小構成（P2 スキップ）** | **10〜13 日**（旧 13〜18 日 → 改訂6 10.5〜14日 → 改訂7 で更に圧縮）|

> Codex の指摘「既存資産流用なら 6〜13日」の **6日下限** は P3 もほぼゼロにした極端な楽観値であり、本設計書では P1（修正キュー + triage + openpyxl DataValidation）と P6（kintai-imports POST の動作検証 + 監査ログ + 二重投入防止）が本番品質では譲れないため、**8〜10 日を下限、12〜13 日を中央値** として扱う。改訂7（フォルダ事前分類運用）で Codex の上限値 13日に一致した。

### 7.4 5月分との切り分け（2026-05-26 確定）

- **5月分（5/29 締切、残り 3 営業日）**：[`PLAN_5月本番_3営業日MVP.md`](./PLAN_5月本番_3営業日MVP.md) の `quick_merge.py` + `quick_export.py` で乗り切る（P1/P6 の劣化版を半日強で実装）
- **6月分**：本設計書の P0〜P6（本格最小構成は P2 スキップ）を **5/30 着手で 3〜4 週間** をターゲット
- **3 営業日で P0〜P6 を完成させるのは依然として不可能** — P1 の本番品質シート生成 + P6 の API 投入安全装置を 3 日に圧縮することはできない

---

## 8. リスク・未解決事項

### 8.1 リスク

| リスク | 影響 | 対策 |
|---|---|---|
| 将来、同姓同名が発生 | ambiguous 多発、マニュアル介在 | 月次マスター上で目立たせる。発生時にはカナ・所属で絞り込む副ロジック追加 |
| 漢字違い・カナ/全角半角ゆれ | 名前マッチ失敗 | `normalize_name_keys` のテストケースを請求勤怠の実例で増強 |
| 外国籍名（アルファベット）| マッチ失敗 | アルファベット名抽出ルールを明示的に追加（"Maharjan Ramita" 等）|
| アグリゲートファイル内の名前抽出失敗 | 一部社員が「未紐付け」化 | パーサーごとのテストで担保、新パターン出現時に追加 |
| ソース戦略の自動判定ミス | パース失敗 | フォールバックで `parse_timesheet_smart` のモード判定に委譲、未判定ファイルはログへ |
| 除外グループ名の表記ゆれ（"F付き" の定義揺れ）| 提出忘れ誤検知 | 除外条件を運用ルールとして文書化、起動オプションで上書き可能に |
| 差異種別ハッシュの粒度設計 | 関係ないコメントが引き継がれる | 初回は保守的に（細かい粒度）、運用で調整 |
| 日次データを持たないソースが多い場合の検知力低下 | 休憩打刻ミスなど月次では拾えない | 「日次突合不可」を可視化、運用で日次取得できる手段を継続的に追加 |
| jinjer API レート制限 | 取得に時間 | キャッシュ・バッチサイズ調整 |
| **修正キューの誤承認**（auto_fix_candidate を内容確認せず一括承認）| 誤データが jinjer に投入される | needs_check の判定条件を保守的に保つ。fuzzy_match / monthly_only / day_crossing / submission_missing / unknown_reason は **常に needs_check に降格**。auto_fix_candidate のみの一括承認 UI を限定提供 |
| **未判断行の見落とし** | レビュー未完で投入される | P6 ゲートが既定で未判断ゼロを要求。サマリシートに未判断件数を強調表示。--force-incomplete は明示指定が必要 |
| **却下/保留行の意味取り違え** | 「却下」した行を翌月以降も誤って繰り返し処理 | §10 表で運用合意を明文化、判断メモを必須運用化 |
| **保留行の塩漬け** | 「保留」のまま放置されて月跨ぎで増殖 | 翌月マスターで再表示 + コメント引き継ぎ。サマリシートに「保留累計」を表示し、3ヶ月以上保留を強調 |
| **同一行への二重投入**（CSV を別経路から再実行）| jinjer 側で打刻重複 | 修正キュー行に `import_id` + ステータス `imported` を残し、再実行時はスキップ |

### 8.2 未解決事項（**Phase C 棚卸し済 + レビューゲート設計確定**）

2026-05-19 + 2026-05-22 に全て確定済み。

1. ✅ **本社常駐者の jinjer 勤怠取得方法** — **取得しない**。詳細シートにステージ 0「突合不要」とだけ表示。230人全体の進捗を見える化しつつ API 負荷を最小化
2. ✅ **「差異種別ハッシュ」の具体的なハッシュ関数** — `(社員番号, 判定種別)` の粗い粒度で。判定種別は matcher.py 出力の `OK / NG / 要確認 / データ欠損` の4種
3. ✅ **各ソースの日次データ取得可否** — **全ソース日次取得可能**（実物棚卸し済、§6.4 参照）。SAP も e-staffing も 1人複数行で日次データを持つ。`monthly_only` 型は将来の保険として残す
4. ✅ **除外グループ名のリスト確定** — 「管理部」「ビジネス企画本部」「ディスパッチグループ」「F付き」の4種で確定。F は「グループ名のどこかに F が含まれる」判定
5. ✅ **アグリゲートファイル抽出ロジック** — 既存 `timesheet_parser.py` でカバー済。Vision プロンプトに「複数人分は配列で返す」と明示済、`for sheet in wb.worksheets` で複数シート対応済。**テストケース追加のみで OK**:
   - 「上原, 奏吾」（SAP の「姓, 名」カンマ区切り）→ `normalize_name_keys` で対応済
   - 「寺山 枝美」（e-staffing の「姓 名」スペース区切り）→ 同上
6. ✅ **修正キューの「人間判断」値の意味（2026-05-22 確定）** — 承認=請求勤怠で修正、**却下=jinjer 側を正に確定（請求勤怠を直してもらう運用通知は別途）**、**保留=判断保留で翌月マスターに未解決として再表示+コメント引き継ぎ**、空欄=未判断
7. ✅ **auto_fix_candidate の初期人間判断値（2026-05-22 確定）** — **空欄=未判断、明示承認が必要**。auto_fix_candidate 一括承認の運用補助を Excel 上で提供
8. ✅ **P6 ゲートの厳格度（2026-05-22 確定）** — **未判断が1件でもあれば既定で WARNING を出して中断**（exit 1）。`--force-incomplete` フラグで承認分のみ反映可（未判断・保留は次回繰越）
9. ✅ **修正キューの実装場所（2026-05-22 確定）** — **月次マスター_<YYYY-MM>.xlsx 内の「修正キュー」シート**。別ファイル分離・SQLite 化は不採用（二重管理リスク・実装コスト増のため）

---

## 9. P5: 人間レビュー運用 ★詳細化

### 9.1 目的

P1/P2 で生成された修正キューシートに対し、人間が「人間判断」列を埋めることで、P6 のインポート対象を確定する。**自動修正の暴走を防ぐためのゲート**。

### 9.2 運用フロー（管理部の月次オペレーション）

```
1. batch_run.py 実行（P3）→ 月次マスター_<YYYY-MM>.xlsx 生成
2. 月次マスターを開き、「サマリ」シートで未判断件数・却下未通知件数を確認
3. 「修正キュー」シートに移動
4. auto_fix_candidate（exact match のみ）をレビュー
   - フィルタで絞り込み、行を上から目視
   - 「自動修正提案値（=請求勤怠値）」と「jinjer 値」の差を確認し、納得できれば 承認
   - 全件機械的に承認するのは禁止。最低でも「差分が極端でないか」「対象日が実労働日か」をスポット確認
   - 1人あたり数行〜十数行程度なので、慣れれば 1 件 5〜10 秒で判断できる想定
5. needs_check 行を 1 件ずつレビュー
   - 要チェック理由列を確認
   - 必要に応じて元データ（請求勤怠 / jinjer）を確認
   - 「人間判断」を 承認 / 却下 / 保留 のいずれかに記入
   - 「判断メモ」に判断根拠を簡潔に記載
   - 却下を選んだ場合は「通知先」列を埋める（請求勤怠提供元の連絡先）
6. 全行の「人間判断」列が埋まったら保存（ステージ 3b 到達）
7. import_run.py を実行（P6）
8. 翌日以降、却下行の「通知済」へのステータス更新を運用で実施
```

#### 一括承認に関する運用ルール

- Excel データ検証で auto_fix_candidate の「人間判断」プルダウンを **空欄スタートに固定**
- マクロや「フィルタ可視範囲を一括承認」ボタンは **提供しない**（誤承認リスクが大きいため）
- 代わりに、auto_fix_candidate を絞り込んだ後で「判断者」列に自分の名前を入力し、行ごとにプルダウンで承認する運用にする
- どうしても運用負荷が問題になった場合のみ、「マッチ精度 = exact」かつ「差分絶対値 ≤ 30 分」かつ「ヒューリスティック理由が休憩打刻ミス」に限定したワンクリック承認マクロを後日追加検討

### 9.3 承認/却下/保留の運用合意（再掲）

| 値 | 業務的意味 | jinjer への動作 | 通知運用 |
|---|---|---|---|
| **承認** | 請求勤怠の値で修正してよい | 自動修正提案値（請求勤怠値）で打刻データを上書き | 不要 |
| **却下** | 請求勤怠側が誤り。jinjer 側が正しい | 何もしない（jinjer を変更しない）| **「却下通知ステータス」初期値=未通知**。担当者が請求勤怠提供元（派遣会社・本人）に修正依頼を出したら手動で「通知済」に更新 |
| **保留** | 判断保留 | 何もしない。翌月マスター生成時に未解決として再表示 | 不要 |

### 9.4 承認権限のルール（運用ガイドライン）

- 修正キューの「人間判断」列を埋めるのは **管理部の指定担当者** のみ
- 1人 == 1か月分を1日でレビューできる量を想定（230人 × 平均1〜2行 = 約400行/月）
- 判断者と判断日時を残し、監査可能にする（Excel ユーザー名 + 入力日時の自動記録）

### 9.5 実装ファイル

- なし（コード上の独立ファイルは不要）
- 「人間判断」「却下通知ステータス」のプルダウンは **openpyxl の `DataValidation`** で月次マスター生成時に直接セルに設定する（P1 master_aggregator のシート組み立て内に組み込む）
- VBA マクロは作らない。`tools/setup_review_validation.bas` のようなヘルパーマクロは **不要**（過去版で記載していたが撤回）

```python
# P1 master_aggregator から呼ぶ想定の最小例
from openpyxl.worksheet.datavalidation import DataValidation

def attach_review_dropdowns(ws, header_row: int, last_row: int):
    judge_col_letter = 'O'  # 「人間判断」列の列文字（実装時に確定）
    notify_col_letter = 'R'  # 「却下通知ステータス」列

    judge_dv = DataValidation(
        type='list',
        formula1='"承認,却下,保留"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle='不正な値',
        error='承認 / 却下 / 保留 のいずれかを選択してください',
    )
    ws.add_data_validation(judge_dv)
    judge_dv.add(f'{judge_col_letter}{header_row + 1}:{judge_col_letter}{last_row}')

    notify_dv = DataValidation(
        type='list',
        formula1='"不要,未通知,通知済"',
        allow_blank=True,
    )
    ws.add_data_validation(notify_dv)
    notify_dv.add(f'{notify_col_letter}{header_row + 1}:{notify_col_letter}{last_row}')
```

### 9.6 工数

- openpyxl DataValidation の組み込み（P1 内）: **数時間**（追加実装としては実質ゼロ）
- 運用ガイド（管理部向け Markdown）整備: **0.5日**

合計: **0.5日**（旧見積 1日 から短縮）

---

## 10. P6: jinjer kintai-imports 自動投入 ★詳細化

### 10.1 目的

修正キューの **承認済み** 行から汎用データCSVを生成し、`POST /v1/kintai-imports` (type.id=5) で jinjer へ自動投入する。**人間レビューゲートを経由しないルートは存在しない**。

### 10.2 CLI 仕様

```powershell
cd Z:\勤怠チェックシステム\kintai-checker

# dry-run（CSV を生成して保存するだけ、jinjer へは投げない）
python import_run.py `
  --month               2026-05 `
  --master              "Y:\給与明細\R8年\5月\月次マスター_2026-05.xlsx" `
  --output-csv          "Y:\給与明細\R8年\5月\jinjer_import_2026-05.csv" `
  --dry-run

# 本実行（jinjer へ POST）
python import_run.py `
  --month               2026-05 `
  --master              "Y:\給与明細\R8年\5月\月次マスター_2026-05.xlsx" `
  --output-csv          "Y:\給与明細\R8年\5月\jinjer_import_2026-05.csv" `
  --execute

# 未判断が残っていても強制実行（承認分のみ反映、その他は次回持ち越し）
python import_run.py --execute --force-incomplete ...
```

### 10.3 処理フロー

```
[Step 1] 月次マスターを開く（読み取りモード推奨）
  └─ services/review_loader.load_review_queue()

[Step 2] レビュー状態バリデーション ★レビューゲート
  ├─ 修正キュー全行を走査し、人間判断の値を集計
  ├─ 「承認」: CSV 対象
  ├─ 「却下」「保留」: スキップ（理由をログ出力）
  ├─ 「空欄(未判断)」が 1 件でもあれば:
  │   ├─ 既定: WARNING を出して中断（exit code 1）
  │   └─ --force-incomplete 指定時: 警告のみ出して継続、未判断行は次回繰越
  ├─ 「却下 かつ 通知ステータス=未通知」を集計
  │   └─ 中断はしないが、コンソールに件数を表示（運用フォローアップを促す）
  ├─ ステータスが exported/imported の行はスキップ（再投入防止）
  └─ 異常検出（同一キーで複数行・データ欠損など）は中断

[Step 3] 承認行から自動修正提案値を抽出
  └─ services/kintai_import_builder.build_rows(approved_rows)
      ├─ 社員番号・日付・項目コード（勤務開始/終了/休憩等）に変換
      └─ 汎用データCSV形式（type.id=5）に整形

[Step 4] CSV 出力（dry-run / 本実行 共通）
  └─ output_csv 指定先に保存

[Step 5] dry-run プレビュー
  ├─ 標準出力に「承認 N件 / 却下 M件（うち未通知 X件）/ 保留 K件 / 未判断 L件」表示
  ├─ CSV の先頭 20 行をコンソールに表示
  └─ ここで終了（--dry-run の場合）

[Step 6] 本実行（--execute 指定時）
  ├─ 直前に「これから N 件を jinjer に送信します。続行しますか？ (y/N)」プロンプト
  ├─ POST /v1/kintai-imports に CSV をアップロード
  ├─ レスポンスの import_id を取得
  └─ ジョブ完了までステータスポーリング（GET /v1/kintai-imports/{id}）

[Step 7] 結果反映
  ├─ 「修正キュー」シートの承認行を ステータス = exported / imported に更新
  ├─ 該当行に「import_id」「投入日時」「実行者」を記録
  ├─ 「インポート履歴」シートに監査ログを追記（承認・却下・保留・未判断・却下未通知の件数を含む）
  └─ 該当社員の「ステージ」を 4→5 に更新
```

### 10.4 レビューゲートの実装方針

```python
def validate_review_state(rows, force_incomplete: bool = False) -> GateResult:
    approved = [r for r in rows if r['人間判断'] == '承認']
    rejected = [r for r in rows if r['人間判断'] == '却下']
    held = [r for r in rows if r['人間判断'] == '保留']
    pending = [r for r in rows if not r['人間判断'].strip()]
    rejected_unnotified = [
        r for r in rejected if r.get('却下通知ステータス') == '未通知'
    ]

    if pending and not force_incomplete:
        raise ReviewIncompleteError(
            f"未判断が {len(pending)} 件あります。"
            f"修正キューシートを確認するか、--force-incomplete を指定してください。"
        )
    if rejected_unnotified:
        # 警告のみ。インポートは継続。
        logger.warning(
            f"却下行のうち {len(rejected_unnotified)} 件が「未通知」のままです。"
            f"請求勤怠提供元への連絡が完了したら、修正キューの『却下通知ステータス』を"
            f"『通知済』に更新してください。"
        )
    return GateResult(
        approved=approved, rejected=rejected, held=held, pending=pending,
        rejected_unnotified=rejected_unnotified,
        will_export=approved,  # 承認のみが CSV 出力対象
    )
```

### 10.5 汎用データCSV（type.id=5）のマッピング

| 修正キュー列 | CSV 列 | 備考 |
|---|---|---|
| 社員番号 | `staff_code` | jinjer 側の社員番号と完全一致を要求 |
| 対象日付 | `target_date` | YYYY-MM-DD |
| 差異種別 + 自動修正提案値 | `item_code` + `value` | 「休憩時間」60分 → 該当のkintai項目コードと分換算値 |
| (省略) | ... | 詳細は jinjer 仕様書に従う |

実際の項目コードは [jinjer API チートシート](Z:/API連携/memory/jinjer_api_cheatsheet.md) を参照。

### 10.6 安全装置

- **二重投入防止**: 修正キュー行のステータスが `imported` の場合、再実行時にスキップ
- **import_id ロギング**: jinjer 側で承認待ち中の jobs を `GET /v1/kintai-imports` で確認できる状態を維持
- **ロールバック**: jinjer 側の取消 API は提供されないため、誤投入時は **承認保留運用** を介して打刻データを手修正（P5 周知済み運用）
- **アクセス制御**: import_run.py 実行は管理部の本番担当者のみ。`JINJER_API_KEY` は本番用キーを `.env` 経由で
- **却下行の宙浮き防止**: 「却下 かつ 通知ステータス=未通知」をサマリシートに件数表示、import 実行時に警告。翌月マスター生成時にも「却下未通知件数」を引き継ぎ集計

### 10.7 実装ファイル

- `import_run.py` — CLI エントリ（argparse）
- `services/review_loader.py` — 修正キューシート読込 + バリデーション
- `services/kintai_import_builder.py` — CSV 生成 + POST 投入
- `services/jinjer_api_client.py` — `post_kintai_imports()` メソッド追加
- `tests/test_review_loader.py` — ゲート判定（未判断あり/なし、保留・却下の扱い、二重投入防止）
- `tests/test_kintai_import_builder.py` — CSV フォーマット、項目コードマッピング

### 10.8 工数

**既存資産（流用前提）:**
- `Z:/API連携/update_salary_unit_prices.py` の `request_with_retry` / `calc_retry_wait_seconds` / `parse_retry_after_seconds`（429リトライ・`Retry-After`解釈）をそのまま移植
- 同ファイルの dry-run → 本実行フロー、子プロセス呼び出し、ログ出力フォーマットを踏襲
- 認証は既存 `JinjerClient.authenticate()` をそのまま使用

**新規に書くもの:**
- `services/review_loader.py`（修正キュー読込 + バリデーション）
- `services/kintai_import_builder.py`（汎用データCSV 生成 + POST 投入 + ジョブポーリング）
- `JinjerClient.post_kintai_imports()` / `get_kintai_import_status()` メソッド追加
- 監査ログ（インポート履歴シート書き戻し）
- 二重投入防止（ステータス `imported` の行スキップ + `import_id` ロギング）

実装: **3〜4日**（旧見積 4〜6日 から短縮）— ただし POST `/v1/kintai-imports` (type.id=5) のフォーマット動作検証が初回は実機で必須のため、見積の下限を割るのは難しい。

---

## 11. その後のフェーズ

| Phase | 内容 | 着手予定 |
|---|---|---|
| P7 | 「承認保留」運用ルール周知（紙ベース）| P0〜P6 安定稼働後 |
| P8 | Web ダッシュボード化（Flask 拡張、修正キューの Web レビュー UI を含む）| P6 完了後 |

P0〜P3 が稼働すれば、**管理部の進捗管理は劇的に改善される**。
P6 まで進めば修正反映も自動化されるが、**人間レビューゲートにより安全性は担保**される。
P8 で Web UI 化することで Excel ファイルの取り回し負荷を解消し、完全な属人化解消となる。

---

## 12. レビューしてほしいポイント

「✅」は確定済。残りは実装着手前に判断したい項目。

1. **アーキテクチャ全体**: services/ に新規追加する分割粒度（7ファイル: employee_resolver, master_aggregator, triage, comment_carryover, batch_runner, review_loader, kintai_import_builder）でOKか
2. ✅ **P0 の方針**: タグ API を使わず、在籍者マスタ + ファイル名/内容から氏名抽出する方針（B案で合意済）
3. **区分の判定（P1 で確定）**: match / skip / **submission_missing** / **ambiguous** の4区分でOKか
4. **氏名抽出ロジック**: ファイル名から漢字氏名・アルファベット名を抽出する正規表現の精度、アグリゲートファイルの扱い
5. **氏名マッチング結果の扱い**: exact / matched_fuzzy / ambiguous / not_found の4分類で過不足は。matched_fuzzy を常に needs_check に降格する方針でOKか
6. **P1 のシート構成**: サマリ・詳細・日次突合結果・修正キュー・コメント履歴・取込ログ・提出忘れ候補・**インポート履歴** の8シートでOKか
7. **「日次突合結果」シートの列設計**: 勤務/休憩/実働＋差分＋差異理由候補＋トリアージ区分＋要チェック理由 で過不足は
8. **「詳細」シートの新規列**: 修正キュー件数(auto/check)・未判断件数 を追加した8列強で過不足は
9. ✅ **差異理由候補の自動推定ロジック**: 「休憩打刻ミス」「未打刻」「日跨ぎ」「請求勤怠未提出」「誤差許容範囲」「不明」の6分類
10. **トリアージロジック ★更新（§4.7）**: ②〜⑩ の優先順位、`match_confidence` での閾値判定、理由コード（comment_present / fuzzy_match / monthly_only / parse_failed / **data_missing** / day_crossing / submission_missing / unknown_reason）の網羅性
11. **「修正キュー」シート設計 ★更新（§4.6）**: 列構成（人間判断・判断メモ・自動修正提案値 + **却下通知ステータス・通知先・通知日時**）と auto_fix_candidate の目視承認 UX（一括承認マクロは原則提供しない）
12. ✅ **却下/保留の運用合意**: 却下=請求勤怠が誤り（jinjer を正に確定）/ 保留=判断保留で翌月持ち越し
13. **ステージ定義（0〜7、2.5/3a/3b 追加）**: 粒度感はOKか
14. ✅ **P2 引き継ぎキーの粒度**: `(社員番号, 判定種別)` の粗い粒度（OK/NG/要確認/データ欠損）で確定済
15. **P2 再トリアージ ★新規（§5.5）**: 引き継ぎコメントが付いた行を auto_fix_candidate → needs_check に降格する方針でOKか
16. **P3 の SOURCE_STRATEGIES**: SAP_Fieldglass / e-staffing / ERCSTS / PDF個別 / 画像個別 / 個別Excel のフォールバックチェーンで網羅できるか
17. **P3 のCLI設計**: 引数の過不足、Windows での運用イメージ
18. **P6 のレビューゲート ★更新（§10.4）**: 未判断あり → 既定で中断、`--force-incomplete` で続行可能。**却下未通知は警告のみで継続**の方針でOKか
19. **P6 の CSV マッピング ★新規（§10.5）**: 修正キュー → 汎用データCSV（type.id=5）の項目コード変換、別途仕様確定が必要
20. **P6 の安全装置 ★更新（§10.6）**: 二重投入防止・import_id ロギング・取消手段・**却下未通知の宙浮き防止**の運用合意
21. **却下行の通知運用 ★新規（§4.6 / §10.4）**: 修正キューに通知ステータス列を持つ案で OK か。別シート「却下→修正依頼」に分離するか

レビュー後、修正点を反映してから実装に着手します。

---

## 変更履歴

| 日付 | 変更内容 |
|---|---|
| 2026-05-19 (初版) | P0〜P3 初期設計 |
| 2026-05-19 (改訂1) | 日次突合シート追加・管理タグ判定ロジック修正・月次のみ突合フォールバック追加 |
| 2026-05-19 (改訂2) | **タグ運用を廃止、ファイル名/内容からの氏名抽出ベースに全面書き換え**。employee_resolver 新設、ambiguous/submission_missing 区分追加、提出忘れ候補シート追加 |
| 2026-05-19 (改訂3) | **Phase C 棚卸し反映**。全ソースが日次取得可と確定（SAP_Fieldglass を `monthly_only` から `daily_and_monthly` に修正）、skip 者は jinjer サマリ取得せずステージ0表示、差異種別ハッシュは判定種別のみの粗い粒度、除外グループ4種で確定 |
| 2026-05-22 (改訂4) | **人間レビューゲートを必須化**。修正キューシートを再設計（人間判断・判断メモ列追加）、トリアージロジック新設（§4.7）、リスク高ケース（fuzzy_match / monthly_only / parse_failed / day_crossing / submission_missing / unknown_reason）はコメントなしでも needs_check に降格。P2 に再トリアージ追加、P3 にトリアージステップ追加、P5/P6 詳細フェーズ設計を書き下ろし（CSV 生成前に承認状態バリデーション、未判断あれば既定中断・--force-incomplete で続行可）。ステージ定義に 2.5/3a/3b を追加 |
| 2026-05-26 (改訂5) | **運用方針の明文化と詰めの補強**。①§1.2 ゴールにレビューゲート必須を明記＋運用方針表追加。②triage の理由コードを `parse_failed`（パース失敗）と `data_missing`（必須フィールド空）に分離。③社員紐付け信頼度を `match_confidence: float` に変更し閾値判定 (`FUZZY_THRESHOLD`) で扱う。④§9.2 P5 運用フローを「フィルタ一括承認」から「目視承認」に厳格化（auto_fix_candidate でも 1 件ずつプルダウン操作）。⑤却下行の宙浮きを防ぐため、修正キューに「却下通知ステータス／通知先／通知日時」列を追加。サマリシートに「却下未通知件数」表示、P6 ゲートで警告。レビューポイント 21 を追加 |
| 2026-05-26 (改訂6) | **既存資産流用前提で工数再精査**。Codex 指摘を反映：①P0 employee_resolver は `services/jinjer_api_client.py`（認証・在籍者・所属履歴・氏名→IDマップ・打刻グループ）が既に揃っているため **2〜3日 → 1〜1.5日**。②P3 batch_runner は既存 `parse_timesheet_smart` の薄いラッパーなので **3〜5日 → 2〜3日**。③P6 kintai_import_builder は `Z:/API連携/update_salary_unit_prices.py` の 429リトライ・dry-run フレームワークを移植して **4〜6日 → 3〜4日**。④**P5 はデータ検証プルダウンを openpyxl `DataValidation` で実装（VBA マクロは不要）**。`tools/setup_review_validation.bas` の記載を撤回し、§9.5 に最小実装例を追加、**1日 → 0.5日**。⑤§7 を「7.1 既存資産棚卸し / 7.2 各 Phase 再見積り / 7.3 累計 / 7.4 5月分との切り分け」に再構成。フル **16〜23日 → 12.5〜17日**、本格最小構成（P2 スキップ）**13〜18日 → 10.5〜14日**。**3 営業日 P0〜P6 完成は依然として不可能**、5月は MVP 維持・6月以降に本格版移行の方針は不変 |
| 2026-05-26 (改訂7) | **請求勤怠のソース判定をフォルダ名一次キー + 中身ガード方式に変更**。ユーザーが請求勤怠を `SAP_Fieldglass / e-staffing / Excel / その他` のサブフォルダに事前格納する運用前提に変更。§6.4 の `SOURCE_STRATEGIES` を `detect_by_filename` / `detect_by_content_header` の自動判定リストから「フォルダ名→戦略」の dict に簡素化、中身の `expected_headers` チェックは誤格納検知のためのガードとして残す。P3 batch_runner **2〜3日 → 1.5〜2日** に短縮。§7.3 累計：フル **12.5〜17日 → 12〜16日**、本格最小構成 **10.5〜14日 → 10〜13日**（Codex の上限値 13日と一致） |
