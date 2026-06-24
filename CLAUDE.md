> 作業を始める前にCLAUDE.mdを読んで前回の状態を把握して。
> 作業が終わったらCLAUDE.mdの「最新の作業状態」を更新してからセッションを終えて。
> 適宜良きタイミングでブランチを切ること、GitHubに保存すること
> APIと.venvは.gitignoreに入れること

# 著作権調査支援ツール（CyosakukenJIdouka_app）

## 🔧 プロジェクト概要

NUENDO から書き出した Cue CSV と、PC 内の WAV/MP3 音源ファイルを照合し、
番組で使用した楽曲を整理するアプリ。

音響効果・選曲業務における Cue Sheet 作成補助ツール。
最終的には JASRAC/J-WID・NexTone 検索補助と Excel 出力ができるツールにする。

### 技術スタック

- **Python + Streamlit**（ローカル Web アプリ）
- pandas：CSV/DataFrame 処理
- openpyxl：Excel 出力
- chardet：文字コード自動判定

### フォルダ構成

```
CyosakukenJIdouka_app/
├── app.py                  # メインアプリ（Streamlit）
├── run.bat                 # Windows 起動バッチ
├── requirements.txt
├── .gitignore              # .venv / API キーを除外済み
├── modules/
│   ├── csv_reader.py       # CSV 読み込み・列バリデーション
│   ├── number_parser.py    # ライブラリ管理番号・Audiostock 番号分解
│   ├── normalizer.py       # 照合用文字列正規化・タイトル検出
│   ├── matcher.py          # Cue × WAV 照合ロジック（優先度6段階）
│   ├── excel_exporter.py   # Excel 6シート出力（色分け・フィルター）
│   └── search_helper.py    # J-WID / NexTone / Google 検索語生成
├── samples/
│   ├── sample_cue.csv
│   ├── sample_wav.csv
│   └── sample_mp3.csv
└── scripts/
    ├── Get-WavList.ps1     # WAV 一覧取得 PowerShell スクリプト
    └── Get-Mp3List.ps1     # MP3 一覧取得 PowerShell スクリプト
```

### 管理番号ルール

- **ライブラリ管理番号**: `数字1桁 + 英字2文字 - 3桁 - 2桁`
  - 有効系列: 1ST〜7ST / 1AN〜5AN / 1VO〜2VO / 1VJ〜2VJ
  - 例: `6ST-653-09 GO! GO!`
- **Audiostock**: `audiostock_数字`
  - 例: `audiostock_856447_残念なシーンのジングル`

### 照合優先順位（matcher.py）

1. イベント名 ↔ WAV ファイル名の完全一致
2. 正規化一致（拡張子除去・記号正規化・全半角統一）
3. 管理番号一致
4. 曲名一致（管理番号除去後）
5. NUENDO ファイル名一致
6. 部分一致

WAV で照合できなかった場合のみ MP3 を補助として使う。

## 📋 作業ルール

### セッション終了時は必ず以下を更新すること
作業を終える前に、このファイルの「最新の作業状態」セクションを更新する。
更新内容：
- 今日やったこと（箇条書き）
- 次にやること（箇条書き）
- 未解決の問題や懸念点
- 重要な決定事項とその理由

## 📅 最新の作業状態

### 最終更新：2026-06-25
**やったこと**
- MVP 実装（前回）
- サンプル CSV による動作確認 → 全 6 曲「正規化一致」・イベント 7→楽曲 6 件集約を確認
- J-WID / NexTone 自動スクレイピング機能を追加
  - `modules/scraper.py`：requests + BeautifulSoup4 による HTML スクレイピング
    - J-WID: EUC-JP 対応、テーブルパース、デバッグ HTML 表示
    - NexTone: JSON API → HTML の順でフォールバック、カード/テーブル両対応
    - レート制限（2 秒間隔）・タイムアウト設定済み
  - `app.py` 検索補助タブ：自動調査ボタン、結果表示、「適用」ボタンで楽曲まとめに反映
  - `requirements.txt`：requests / beautifulsoup4 / lxml を追加

**次にやること**
- 実際の J-WID / NexTone でスクレイパーをテスト（HTML 構造確認・パーサー調整）
- 実際の NUENDO Cue CSV でテスト・列名ゆれの確認
- 既存 Excel 再読み込み機能（手入力情報の引き継ぎ）
- 類似一致の精度向上（表記ゆれ対応の拡充）

**未解決の問題**
- J-WID / NexTone のスクレイパーは実サイトで HTML 構造を確認するまで動作保証なし
  → パース失敗時は「デバッグ: 取得 HTML」エキスパンダーで生 HTML を確認して調整
- J-WID は EUC-JP エンコーディング前提（サイト変更で文字化けの可能性）
- NexTone は Next.js 製で内部 API パスが変わると JSON 取得に失敗する可能性
- Duration の形式が環境により異なる可能性（PowerShell の Shell API 依存）

**重要な決定事項**
- 技術スタック: Python + Streamlit（複数人共有・初心者保守性・ローカル処理を優先）
- MP3 は補助のみ（主軸は WAV）・MP3 タグ情報は前提にしない
- J-WID / NexTone 自動スクレイピングを実装（個人業務用ツールとして使用）
- 既存 Excel 読み込みは将来実装（MVP 外）
