# 著作権調査支援ツール

NUENDO Cue CSV × WAV 一覧照合 → 権利情報管理 → Excel 出力

音響効果・選曲業務における Cue Sheet 作成補助ツールです。

---

## 機能概要（MVP）

1. NUENDO から書き出した Cue CSV を読み込む
2. プロジェクト Audio フォルダ内の WAV 一覧 CSV を読み込む
3. WAV ファイル名から曲タイトル候補を検出する
4. WAV フル尺を取得・表示する
5. 管理番号を分解する（ライブラリ管理番号 / Audiostock）
6. イベント名優先で Cue と WAV を照合する
7. WAV で照合できなかったものを「MP3 補助確認」ステータスにする
8. MP3 一覧 CSV を補助として読み込める
9. 楽曲まとめ表を画面に表示する
10. 手入力で作家情報・作品コード・確認ステータスを編集できる
11. Excel に出力できる（6 シート）

---

## セットアップ

### 必要環境

- Windows 10 / 11
- Python 3.11 以上（[python.org](https://www.python.org/) からダウンロード）

### 初回セットアップ・起動

`run.bat` をダブルクリックするだけで自動セットアップと起動を行います。

手動で実行する場合：

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

ブラウザで `http://localhost:8501` にアクセスしてください。

---

## 使い方

### ① WAV 一覧 CSV の準備

PowerShell スクリプトで Audio フォルダの WAV 一覧を取得します。

```powershell
# PowerShell を開いて実行
.\scripts\Get-WavList.ps1 -AudioFolder "H:\プロジェクト名\Audio" -OutputCsv "wav_list.csv"
```

MP3 一覧が必要な場合（補助）：

```powershell
.\scripts\Get-Mp3List.ps1 -Mp3Folder "H:\MP3ライブラリ" -OutputCsv "mp3_list.csv"
```

### ② アプリを起動して CSV を読み込む

1. `run.bat` をダブルクリックしてアプリを起動
2. 「ファイル読み込み」タブを開く
3. Cue CSV と WAV 一覧 CSV をアップロード
4. 「照合実行」ボタンを押す

### ③ 楽曲まとめで確認・編集

「楽曲まとめ」タブで：
- 照合結果を確認
- 作曲者・作詞者・JASRAC作品コードなどを手入力
- 確認ステータスを更新

### ④ 検索補助で著作権情報を調査

「検索補助」タブで：
- 楽曲を選択すると検索語を自動生成
- J-WID・NexTone・Google の検索リンクを表示

### ⑤ Excel に出力

「Excel 出力」タブで「Excel ファイルを生成」ボタンを押してダウンロード。

---

## 管理番号について

### ライブラリ管理番号

形式: `数字1桁 + 英字2文字 - 3桁 - 2桁`

| 系列 | 使用可能な先頭数字 |
|------|----------------|
| ST   | 1〜7           |
| AN   | 1〜5           |
| VO   | 1〜2           |
| VJ   | 1〜2           |

例: `6ST-653-09 GO! GO!` → 管理番号: `6ST-653-09` / 盤番号: `6ST-653` / トラック: `09`

### Audiostock 管理番号

形式: `audiostock_数字`

例: `audiostock_856447_残念なシーンのジングル` → 番号: `856447`

---

## フォルダ構成

```
CyosakukenJIdouka_app/
├── app.py              # メインアプリ
├── run.bat             # Windows 起動スクリプト
├── requirements.txt    # Python ライブラリ一覧
├── modules/
│   ├── csv_reader.py       # CSV 読み込み（文字コード自動判定）
│   ├── number_parser.py    # 管理番号分解
│   ├── normalizer.py       # 文字列正規化
│   ├── matcher.py          # Cue × WAV 照合ロジック
│   ├── excel_exporter.py   # Excel 出力
│   └── search_helper.py    # 検索語生成
├── samples/
│   ├── sample_cue.csv      # Cue CSV サンプル
│   ├── sample_wav.csv      # WAV 一覧 CSV サンプル
│   └── sample_mp3.csv      # MP3 一覧 CSV サンプル
└── scripts/
    ├── Get-WavList.ps1     # WAV 一覧取得スクリプト
    └── Get-Mp3List.ps1     # MP3 一覧取得スクリプト
```

---

## 確認ステータス一覧

| ステータス | 意味 |
|-----------|------|
| 未調査 | まだ調査していない |
| 候補あり | 候補は特定できているが確認中 |
| 確定 | 権利情報確定 |
| 要確認 | 何らかの問題あり |
| J-WID要確認 | J-WID での確認が必要 |
| NexTone要確認 | NexTone での確認が必要 |
| ライブラリ元確認 | ライブラリ元への確認が必要 |
| Audiostock確認 | Audiostock での確認が必要 |
| MP3補助確認 | WAV で照合できず MP3 での補助確認が必要 |

---

## 注意事項

- J-WID / NexTone の自動スクレイピングは行いません（検索補助と手確認のみ）
- NUENDO の使用尺は編集後の使用尺であり、曲のフル尺ではありません
- 曲特定には Audio フォルダ内の WAV フル尺を使います
- MP3 タグ情報は前提にしません（タグなしファイルにも対応）
