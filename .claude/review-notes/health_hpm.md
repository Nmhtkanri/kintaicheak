# 健康診断 HPM モード（健診結果 → HPM 取込 CSV）— 2026-09-10 転記・同日コード検証済

## 入口
- `/health_hpm_preview` `/health_hpm_generate` `/health_hpm_pdf_preview` `/health_hpm_page_image/<session_id>/<int:page>`
- services: `health_hpm_csv` `health_hpm_excel` `health_hpm_master` `health_hpm_match` `health_hpm_options` `health_hpm_pdf`

## 入力
- 健診機関の結果（Excel/PDF）。
- 選択肢シート（Google）。鍵の無い PC は共有フォルダの写し JSON を読む。呼び出し順は「Google 成功→写しを置き直す／失敗→写し→fallback」（`app.py`、`health_hpm_options.py`。写しは一時ファイル→`os.replace`）。
- 変換マスタ `Z:\NMHT総務関係\健康診断\健康診断HPM変換マスタ.xlsx`（`config.py`）。

## 出力と不可逆な処理
- HPM 取込 CSV（302 列固定）。要配慮個人情報なので一時ファイルはローカルへ隔離し 8 時間で掃除、PDF ページ上限 20（`config.py`）。
- 医師名と所見は 302 列に列が無いので CSV には出さず、画面と監査 Excel のみ。

## 業務ルール（絶対ルール）
- **血圧の平均を作らない。** 数値を補完・加工して新しい値を作らない（`health_hpm_csv.py` の書き出し直前の検算、`health_hpm_master.py` の列固定の二重突合、PDF 側は測定回不明なら値ごと捨てる）。
- PDF 直接読み取り：ページ全体 1 枚だけ送ると検査値が入れ替わる（%肺活量と 1 秒率の実測記録）。必ず上下バンドも送る（3 枚。`BAND_TOP/BAND_BOTTOM`）。
- 婦人科検診の既定は女性 ON。CSV には出さない旨の警告あり。
- 所見（聴力〜腹部超音波）と医師名の読み取りは 2026-09-09 追加。

## 既知の落とし穴
- **13X5035440 の機関名がリポジトリ内で未追随。** 運用は「関東ITソフトウェア健康組合(大久保健診センター)」で確定しマスタ Excel を改名したが、`tools/build_health_hpm_master.py` の初期データは旧名「桜十字グランフロント大阪クリニック」のまま。`--force` で再生成すると旧名に戻り、`merged_institutions` の重複除外で選択肢シート側の新名が消える。

## 宿題
- `tools/build_health_hpm_master.py` の初期データを新名に直すか、再生成を禁止するかの判断。
- 機関ごとの帳票差。新しい機関が来たときの未対応検知。
