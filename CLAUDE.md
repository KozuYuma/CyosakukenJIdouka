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

### 最終更新：2026-06-22
**やったこと**
- 仕様書をもとに MVP を実装
  - `app.py`：Streamlit 5 タブ構成（読み込み / 楽曲まとめ / イベント一覧 / 検索補助 / Excel 出力）
  - `modules/csv_reader.py`：UTF-8 / Shift_JIS / UTF-8 BOM 自動判定、列名ゆれ吸収
  - `modules/number_parser.py`：ライブラリ管理番号・Audiostock 番号の分解
  - `modules/normalizer.py`：照合用文字列正規化、ファイル名からのタイトル検出
  - `modules/matcher.py`：優先度6段階の照合ロジック、楽曲まとめ・イベント一覧生成
  - `modules/excel_exporter.py`：6 シート Excel 出力（色分け・フィルター・列幅自動調整）
  - `modules/search_helper.py`：検索語生成（J-WID / NexTone / Google URL）
  - `scripts/Get-WavList.ps1`・`Get-Mp3List.ps1`：WAV/MP3 一覧取得スクリプト
  - `samples/`：サンプル CSV 3 種
  - `README.md`：日本語セットアップガイド
  - `run.bat`：ワンクリック起動バッチ
  - `.gitignore`：.venv / API キー除外設定

**次にやること**
- 実際の NUENDO Cue CSV でテスト・列名ゆれの確認と吸収
- 既存 Excel 再読み込み機能（手入力情報の引き継ぎ）
- Git 初期化・GitHub リポジトリ作成
- 類似一致の精度向上（表記ゆれ対応の拡充）
- MP3 自動収集補助 PowerShell スクリプトの追加

**未解決の問題**
- 実際の NUENDO Cue CSV の列名が想定と異なる可能性あり（normalize_cue_columns で吸収予定）
- J-WID は直接 URL パラメータを渡せない（トップページへ誘導 → 手動検索）
- Duration の形式が環境により異なる可能性（PowerShell の Shell API 依存）

**重要な決定事項**
- 技術スタック: Python + Streamlit（複数人共有・初心者保守性・ローカル処理を優先）
- MP3 は補助のみ（主軸は WAV）・MP3 タグ情報は前提にしない
- J-WID / NexTone の自動スクレイピングは行わない
- 既存 Excel 読み込みは将来実装（MVP 外）
