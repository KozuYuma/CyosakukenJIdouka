# 著作権調査支援ツール

NUENDO Cue CSV × WAV 一覧照合 → J-WID / NexTone 自動調査 → Excel 出力

音響効果・選曲業務における Cue Sheet 作成補助ツールです。

---

## 機能一覧

| 機能 | 内容 |
|------|------|
| CSV 読み込み | NUENDO Cue CSV・WAV 一覧・MP3 一覧を読み込む。UTF-8 / Shift_JIS / BOM 自動判定 |
| WAV 照合 | イベント名優先の6段階照合でプロジェクト Audio フォルダの WAV と突き合わせ |
| 管理番号分解 | ライブラリ管理番号（6ST-653-09 等）・Audiostock 番号を自動分解 |
| タイトル検出 | WAV ファイル名から管理番号を除去して曲タイトル候補を取得 |
| J-WID 自動調査 | J-WID（JASRAC）を自動検索。作品コード・作曲者・作詞者を取得して楽曲まとめに反映 |
| NexTone 自動調査 | NexTone を自動検索。管理番号・作曲者・アーティストを取得して楽曲まとめに反映 |
| 手入力・編集 | 作曲者・作詞者・JASRAC作品コード・確認ステータスなどを画面で直接編集 |
| Excel 出力 | 楽曲まとめ / イベント一覧 / WAV一覧 / MP3一覧 / 検索語一覧 / 確認メモ の6シートを出力 |

---

## セットアップ

### 必要環境

- Windows 10 / 11
- Python 3.11 以上（[python.org](https://www.python.org/) からダウンロード）

### 起動方法

**`run.bat` をダブルクリックするだけで自動セットアップ＆起動します。**

初回は仮想環境の作成とライブラリのインストールが自動で行われます。  
ブラウザで `http://localhost:8501` が自動的に開きます。

手動で実行する場合：

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

---

## 使い方

### ステップ 1：WAV 一覧 CSV を準備する

PowerShell スクリプトでプロジェクトの Audio フォルダを一覧化します。

```powershell
.\scripts\Get-WavList.ps1 -AudioFolder "H:\プロジェクト名\Audio" -OutputCsv "wav_list.csv"
```

MP3 一覧が必要な場合（WAV で照合できなかった場合の補助）：

```powershell
.\scripts\Get-Mp3List.ps1 -Mp3Folder "H:\MP3ライブラリ" -OutputCsv "mp3_list.csv"
```

### ステップ 2：CSV を読み込んで照合する

1. `run.bat` をダブルクリックしてアプリを起動
2. 「ファイル読み込み」タブを開く
3. Cue CSV・WAV 一覧 CSV をアップロード（MP3 は任意）
4. 「照合実行」ボタンを押す

### ステップ 3：楽曲まとめを確認・手入力する

「楽曲まとめ」タブで：
- 照合結果（照合ステータス・WAVフル尺など）を確認
- 作曲者・作詞者・JASRAC作品コードなどを直接入力
- 確認ステータスをプルダウンで更新

### ステップ 4：J-WID / NexTone で自動調査する

「検索補助」タブで：

1. 調査したい楽曲を選択
2. 検索語を選ぶ（WAV検出タイトル・管理番号・盤番号などから選択）
3. **「自動検索」ボタン**を押す
4. J-WID と NexTone の検索結果が並んで表示される
5. 正しい候補の **「✅ このデータを楽曲まとめに適用」** を押す  
   → 作曲者・作詞者・JASRAC作品コード・NexTone管理番号が自動入力される

> **パース失敗時の対処**  
> サイトの HTML 構造が変わるとパースに失敗する場合があります。  
> 各結果タブの「デバッグ: 取得 HTML」を開いて生 HTML を確認し、  
> `modules/scraper.py` の `_parse_jwid_table` / `_parse_nextone_html` を修正してください。

### ステップ 5：Excel に出力する

「Excel 出力」タブで「Excel ファイルを生成」ボタンを押してダウンロード。

---

## 管理番号について

### ライブラリ管理番号

形式：`数字1桁 + 英字2文字 - 3桁 - 2桁`

| 系列 | 使用可能な先頭数字 |
|------|----------------|
| ST   | 1〜7           |
| AN   | 1〜5           |
| VO   | 1〜2           |
| VJ   | 1〜2           |

例：`6ST-653-09 GO! GO!` → 管理番号: `6ST-653-09` ／ 盤番号: `6ST-653` ／ トラック: `09`

### Audiostock 管理番号

形式：`audiostock_数字`

例：`audiostock_856447_残念なシーンのジングル` → 番号: `856447`

---

## 照合の優先順位

WAV との照合は以下の順で試みます。上位でマッチするほど信頼度が高いです。

1. イベント名と WAV ファイル名の完全一致
2. 正規化一致（拡張子除去・全角半角統一・記号正規化）
3. 管理番号一致
4. 曲名一致（管理番号除去後）
5. NUENDO ファイル名一致
6. 部分一致

WAV で照合できなかった場合のみ MP3 を補助として試みます。

---

## フォルダ構成

```
CyosakukenJIdouka_app/
├── app.py                  # メインアプリ（Streamlit 5 タブ）
├── run.bat                 # Windows ワンクリック起動
├── requirements.txt        # Python ライブラリ一覧
├── modules/
│   ├── csv_reader.py       # CSV 読み込み（文字コード自動判定・列名ゆれ吸収）
│   ├── number_parser.py    # ライブラリ管理番号・Audiostock 番号の分解
│   ├── normalizer.py       # 照合用文字列正規化・ファイル名からのタイトル検出
│   ├── matcher.py          # Cue × WAV 照合ロジック（優先度6段階）
│   ├── scraper.py          # J-WID / NexTone 自動スクレイピング
│   ├── excel_exporter.py   # Excel 6 シート出力（色分け・フィルター・列幅）
│   └── search_helper.py    # 検索語生成
├── samples/
│   ├── sample_cue.csv      # Cue CSV サンプル
│   ├── sample_wav.csv      # WAV 一覧 CSV サンプル
│   └── sample_mp3.csv      # MP3 一覧 CSV サンプル
└── scripts/
    ├── Get-WavList.ps1     # WAV 一覧取得 PowerShell スクリプト
    └── Get-Mp3List.ps1     # MP3 一覧取得 PowerShell スクリプト
```

---

## 確認ステータス一覧

| ステータス | 意味 |
|-----------|------|
| 未調査 | まだ調査していない |
| 候補あり | 候補は特定できているが確認中 |
| 確定 | 権利情報確定 |
| 要確認 | 何らかの問題あり |
| J-WID要確認 | J-WID での最終確認が必要 |
| NexTone要確認 | NexTone での最終確認が必要 |
| ライブラリ元確認 | ライブラリ元への確認が必要 |
| Audiostock確認 | Audiostock での確認が必要 |
| MP3補助確認 | WAV で照合できず MP3 での補助確認が必要 |

---

## 注意事項

- NUENDO の使用尺は編集後の使用尺であり、曲のフル尺ではありません
- 曲特定には Audio フォルダ内の WAV フル尺を使います
- MP3 タグ情報は前提にしません（タグなしファイルにも対応）
- J-WID / NexTone の自動調査は個人業務用途での使用を想定しています
- サイトの仕様変更によりスクレイピングが動作しなくなる場合があります
