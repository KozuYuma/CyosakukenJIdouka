> 作業を始める前にCLAUDE.mdを読んで前回の状態を把握して。
> 作業が終わったらCLAUDE.mdの「最新の作業状態」セクションを更新してからセッションを終えて。
> 適宜良きタイミングでブランチを切ること、GitHubに保存すること
> APIと.venvは.gitignoreに入れること

# 著作権調査支援ツール（CyosakukenJIdouka_app）

## プロジェクト概要

NUENDO から書き出した Cue CSV と PC 内の WAV/MP3 音源を照合し、番組使用楽曲を整理するアプリ。
音響効果・選曲業務における Cue Sheet 作成補助ツール。
**起動: `run.bat` をダブルクリック → http://localhost:8501/**

---

## 技術スタック

- Python 3.x + Streamlit（ローカル Web アプリ、Windows）
- pandas / openpyxl / chardet / requests / BeautifulSoup4
- Playwright（MINC ログイン専用）
- 外部 API: Claude (Anthropic) / Spotify / MusicBrainz（すべてオプション）

---

## フォルダ構成

```
CyosakukenJIdouka_app/
├── app.py                  # メインアプリ（2311行）
├── run.bat
├── .gitignore              # .venv / .env を除外済み
├── modules/
│   ├── csv_reader.py       # CSV 読み込み・列バリデーション
│   ├── number_parser.py    # ライブラリ管理番号・Audiostock 番号分解
│   ├── normalizer.py       # 照合用文字列正規化・タイトル検出
│   ├── matcher.py          # Cue × WAV 照合ロジック（優先度6段階）
│   ├── database.py         # 永続化。DATABASE_URL 未設定=SQLite / 設定時=Supabase
│   ├── excel_exporter.py   # Excel 出力（申告フォーマット+5シート）
│   ├── search_helper.py    # J-WID/NexTone/Google 検索語生成、JWID_BASE 定数
│   ├── pipeline.py         # 全自動調査パイプライン（run_pipeline 関数）
│   ├── musicbrainz.py      # MusicBrainz API クライアント
│   ├── musicforest.py      # MINC (minc.or.jp) スクレイパー + Cookie 管理
│   ├── scraper.py          # J-WID / NexTone スクレイパー
│   ├── claude_lookup.py    # Claude API 楽曲情報ルックアップ（オプション）
│   └── spotify.py          # Spotify API 楽曲検索（オプション）
└── H:\PROGRAM\search_music\
    ├── src\login_browser.py  # Playwright MINC ログイン（reCAPTCHA対策）
    ├── auth\state.json       # MINC Cookie（デフォルト、セッション約3時間）
    ├── auth\state_{mail}.json # ユーザー別 Cookie
    ├── auth\chrome-profile\  # Chrome 永続プロファイル（reCAPTCHA回避）
    └── .env                  # MINC_MAIL_ADDRESS / MINC_PASSWORD
```

---

## app.py 画面構成（tabs）

```
tabs[0] 申告フォーム作成
  ① NUENDO Cue CSV（必須）
  ② WAV ファイル一覧（任意）
  ③ MusicBrainz 自動調査
  ④ MP3 ファイル一覧（mp3finder 出力も自動認識）
  ⑤ マスターDB CSV（任意）
  ── 照合実行 ──
  📋 申告フォーマット
    ├─ MINCログイン設定 expander（per-user メール/パスワード）← 表の上
    ├─ MINC セッション状態 + 確認ボタン
    ├─ 確認ステータス / 管理番号種別 フィルター
    ├─ 🔍 一括検索 expander（MINC / J-WID / NexTone）
    ├─ 縮小表示トグル（260px / 560px）
    ├─ data_editor（全楽曲、編集可能）
    └─ 📋 申告フォーマット プレビュー expander（イベント行単位・提出用）
  補完検索（J-WID / MINC / NexTone）
    ├─ 楽曲選択（デフォルト「すべて」表示）
    ├─ 検索語候補（WAV検出タイトル優先）
    ├─ 手動検索リンク（J-WID作品検索 / eJwid直リンク）
    ├─ J-WID 手動検索コード入力 → 反映（expander）
    ├─ 全自動パイプライン（MusicBrainz → MINC / J-WID / NexTone）
    │    ├─ チェックボックス: MINC も使う / Claude API / Spotify API
    │    ├─ 検索語ラジオボタン（term_candidates から選択）
    │    └─ 結果タブ: J-WID / NexTone / 🌲 MINC
    ├─ MINC 楽曲検索（個別）← パイプライン失敗時の保険
    │    ├─ 検索語セレクトボックス（ラベル付き候補）
    │    ├─ 一致方式: デフォルト「2: 前方一致」（match=1は不安定）
    │    └─ MINC詳細 / J-WID作家情報 / 委任者確認 / 申告フォーマット反映
    ├─ 全検索語一覧
    ├─ Excel 出力
    └─ J-WID 管理状況 CSV（JASRAC作品コードが入っている曲を一括取得）

tabs[1] イベント一覧
```

---

## 管理番号ルール

- **ライブラリ**: `数字1桁+英字2文字-3桁-2桁`（1ST〜7ST / 1AN〜5AN / 1VO〜2VO / 1VJ〜2VJ）
  - 例: `6ST-653-09 GO! GO!`
- **Audiostock**: `audiostock_数字`
  - 例: `audiostock_856447_残念なシーンのジングル`

---

## 照合優先順位（matcher.py）

1. イベント名 ↔ WAV ファイル名 完全一致
2. 正規化一致（拡張子除去・記号正規化・全半角統一）
3. 管理番号一致
4. 曲名一致（管理番号除去後）
5. NUENDO ファイル名一致
6. 部分一致

WAV で照合できなかった場合のみ MP3 を補助として使う。

---

## 重要な技術決定事項

### MINC (musicforest.py)

- `SEARCH_URL = "https://www.minc.or.jp/music/list"` — **末尾スラッシュなし**
  - スラッシュありだと Laravel がリダイレクト時にクエリパラメータを消失させる
- `_TIMEOUT = 30`（旧15秒→タイムアウト多発のため延長）
- **`match=1` は MINC 側でキーワード検索として動作**し、全く別の曲が返ることがある
  - デフォルトは `match=2`（前方一致）が最も安定
  - UIの選択肢順: `["2: 前方一致", "3: キーワード", "1: 完全一致"]`
- `_parse_search_results`: 品番が空のとき `collapseDetail` アンカーテキストから正規表現で補完（2枚組対応）
- `_parse_detail`: 作詞作曲同一人の場合、**`elif` ではなく `if`** で両フィールドに入れる
- 区切り文字は **半角 `/`**（全角 `／` は使わない）

### MINC ログイン（login_browser.py）

- reCAPTCHA のため自動ログイン不可 → 実物 Chrome + 永続プロファイルで検知回避
- セッション有効期間: **約3時間**（サーバー側で切れる）
- Cookie 保存先: `H:\PROGRAM\search_music\auth\state.json`
- **ユーザー別対応**: `MINC_STATE_PATH` / `MINC_PROFILE_DIR` 環境変数で切り替え可能
  - app.py でメール入力 → `state_{mail}.json` / `chrome-profile_{mail}` を自動生成
- `check_session()` はセッション経過時間を返す（2.5時間超で警告表示）

### J-WID / NexTone（scraper.py）

- `search_jwid()` / `search_nextone()`: セッション共有・レートリミット付き
- `fetch_jwid_detail(detail_url)`: 詳細ページから作曲者・作詞者を取得
- `fetch_jwid_rights_by_code(jasrac_code)`: JASRACコード直引き（管理状況 CSV・J-WID反映ボタンで使用）
- J-WIDスクレイパー: `elif` ではなく `if` で作詞作曲を両フィールドへ（訳詞のみ `elif`）

### パイプライン（pipeline.py）

- `run_pipeline(event_name, ..., song_title="")`: `song_title` が最優先の検索語
- 優先順: `song_title` > `wav_detected_title` > イベント名から抽出
- `song_title` に `term_candidates` から選んだ語を渡す（app.py でラジオボタン選択）

### 申告フォーマット反映ルール

- **MINC詳細 / J-WID の作曲者・作詞者を優先**（MP3タグより上位）
- 作詞作曲同一人 → 作曲者・作詞者の**両方**に記入
- 複数人 → **半角 `/` 区切り**
- 反映ボタン押下時、作曲者・作詞者未取得なら MINC 詳細を自動フェッチ
- アーティスト名を作曲者フィールドにフォールバックする処理は**削除済み**

### 著作者名フィールド（MINC個別検索）

- `value=""` (空) + `placeholder=作曲者列の値` — デフォルトで曲名を入れると 0件になるため

---

## セッション状態の主要キー（st.session_state）

| キー | 内容 |
|------|------|
| `songs_df` | 楽曲まとめ DataFrame |
| `events_df` | イベント一覧 DataFrame |
| `mf_auth_state` | `(ok: bool, msg: str)` MINC 認証状態キャッシュ |
| `mf_results_{no}` | MINC 検索結果 |
| `mf_detail_{no}_{i}` | MINC 楽曲詳細 |
| `mf_delg_{no}_{i}` | MINC 委任者確認結果 |
| `jwid_minc_{no}_{i}` | J-WID作家情報（MINCから取得） |
| `jwid_manual_{no}` | J-WID手動検索結果 |
| `pipeline_result_{no}` | 全自動パイプライン結果 |
| `jwid_rights_batch` | J-WID 管理状況 CSV 一括取得結果 |

---

## 環境変数（.env — Git 除外済み）

```
MINC_MAIL_ADDRESS=xxx@example.com
MINC_PASSWORD=xxxxxxxx
MINC_STATE_PATH=（省略時はデフォルトパス）
MINC_PROFILE_DIR=（省略時はデフォルトパス）
ANTHROPIC_API_KEY=sk-ant-xxxxx
SPOTIFY_CLIENT_ID=xxxxx
SPOTIFY_CLIENT_SECRET=xxxxx
```

---

## 残っている TODO

### 優先度: 高

- [ ] **MINC ログインセッション UI**: メール入力欄が空のとき `.env` の値を使うことの確認・ドキュメント化
- [ ] **Git ブランチ & GitHub push**: 今セッションの変更をコミット・プッシュ

### 優先度: 中

- [ ] **既存 Excel 再読み込み**: 手入力済みの作曲者・JASRAC コード等を引き継ぐ機能（MVP 外だが要望あり）
- [ ] **NexTone スクレイパー安定性**: Next.js 製のため内部 API パスが変わると失敗する可能性

### 優先度: 低（将来）

- [ ] **J-WID eJwid 検索 URL 自動補完**: `trxID=F00100` で検索語を URL に含める方法の調査（現状 POST のためリンクで完結しない可能性）
- [ ] **MINC 一括検索**: 結果がない曲のリトライロジック
- [ ] **マルチユーザー対応**: 複数人が同時アクセスする場合の session_state 分離

---

## 作業ルール

- セッション終了時に「最新の作業状態」を更新すること
- `.env` / `.venv` は絶対に Git に含めない
- 適宜ブランチを切って GitHub に保存する

---

## 最新の作業状態

### 最終更新: 2026-08-27（session 6）

**やったこと（session 6）** — ブランチ `feat/ui-foundation`

進め方: 「土台を先、構造を後」。この順で1つずつ進めている。
**0. UI土台 ✅ / 1. ログイン＋所有者分け ✅ / 2. 共有楽曲データ song_master ⏳ /
3. その管理タブ / 4. TSP CD 取り込み / 5. Render デプロイ / 6. UI構造（st.navigation）**

- **ステップ0: UI土台**（`3cc2c77`）
  - `.streamlit/config.toml` … 明暗テーマ一式。アクセントは藍 `#2F4B8F`。
    **緑・黄・橙・赤は使わない**。確認ステータスの意味に予約してあるので、
    飾りで使うと状態が読めなくなる
  - `modules/ui.py` … `status_tone` / `inject_css` / `status_bar` /
    `style_status` / `count_done`。
    `st.dataframe` は Styler を受けるが **`st.data_editor` は受けない**ので、
    色付けは読み取り専用の表だけ
  - ヘッダーの帯に ログイン中 / 案件 / 保存先 / MINC と進捗ゲージを出す。
    MINC の状態は `state.json` の更新時刻から作る。`check_session()` は
    毎回 HTTP を叩くので帯には使わない
- **⚙️ タブの並べ替え**（`7e8bf9c`, `87cfef9`）
  - 🗄️ プロジェクト管理（旧⑤マスターDBを内包・畳んだまま）→ ①Cue CSV ／
    ②MP3 → ③WAV → 照合実行
  - 旧③ MusicBrainz 一括補完は**消さずに**照合後の 🔍 一括検索の隣へ移した。
    一括検索が見るのは MINC / J-WID / NexTone だけで、MusicBrainz を消すと
    一括で補完する手段が無くなるため
  - MP3 は「📄 CSV をアップロード」を既定（左）に
- **データが消える穴を塞いだ**（`37d7671`, `3506e7b`）
  - DB への保存は `if st.session_state.project_id:` の中だけ。案件を選んで
    いないと**一切保存されない**のが原因だった
  - 🔄 照合実行の時点で案件が無ければ Cue CSV のファイル名で自動作成する。
    データが生まれるのがこの瞬間なので、ここで作らないと以降の自動保存が
    全部素通りする
  - 保存する場所: 照合実行 / 一括検索の完了 / 一括補完の完了 /
    申告フォーマットCSVのダウンロード / Excel の**生成時**
  - Excel は「生成」ボタンの中にダウンロードボタンがあり、押した瞬間の
    再実行で消える。だから `on_click` ではなく生成時に保存している
  - `on_click` の中の `st.success` は直後の再実行で消えるので、伝言を
    `_autosave_msg` に置いて本体側で `st.toast` に出している
  - **残っている穴**: ✏️ 直接編集の手入力は 💾 ボタンか次のダウンロードまで
    保存されない。1文字ごとに書くと Supabase への書き込みが増えすぎるため
- **ステップ1: ログイン＋案件の所有者分け**（`24fb802`）
  - `modules/auth.py` 新規。環境変数 `APP_USERS=名前:合言葉,名前2:合言葉2`。
    **未設定ならログイン画面を出さない**ので手元の使い方は変わらない。
    ユーザー表も登録画面も作っていない（人の増減は env 1行の書き換えで済む）
  - 合言葉の照合は `hmac.compare_digest`。`==` は違った時点で止まるので
    返答の速さの差から1文字ずつ言い当てられる
  - `projects` に `owner` 列。PostgreSQL は `ADD COLUMN IF NOT EXISTS`、
    **SQLite にはそれが無い**ので重複エラーを握りつぶす（他の DDL を
    巻き込まないよう接続を分ける）
  - `list_projects(owner)` は「自分のもの＋所有者が空のもの」。空のものを
    隠すと分ける前の案件が消えたように見えるため。読み込むとその人のものに
    なる（`set_project_owner`）ので、**移行スクリプトは要らない**。
    一覧では `🔓 所有者なし` と表示される
  - `require_login()` は Chrome 拡張の Cookie 同期処理より**後**に置くこと。
    拡張は裏で新しいセッションを開くので、先に止めると同期が届かない
  - `_load_local_env()` が `DATABASE_URL` があると早期に戻っていたのを修正。
    `.env` には APP_USERS など DB 以外の設定も入るようになったため。
    `setdefault` なので Render 側の環境変数が `.env` に負けることはない

**やったこと（session 5）**

- **NUENDO MP3 Finder（GUI exe）の仕上げ** — ブランチ `feat/mp3-finder-gui-exe`
  - 中止ボタンを追加。スレッドは外から止められないので `threading.Event` を
    要所で見る協調方式（`Cancelled` 例外）。CSV 書き出し直前にも見て、
    中止時に中途半端な CSV が残らないようにしてある
  - 進捗バーを実測に変更。`_PHASE` で各段階に％の帯を割り当てる。
    総数が分からない MP3 スキャン中だけ流し表示（indeterminate）
  - ログ欄のフォントを `TkDefaultFont` の複製に。`tk.Text` の既定は等幅の
    `TkFixedFont` で、日本語が痩せて見えていた
  - exe: `dist\NUENDO_MP3_Finder.exe`
- **Chrome 拡張 MINC Session Sync の画面切り替えを解消** — ブランチ `feat/minc-sync-silent`
  - `chrome.tabs.create({active:false})` で裏タブを開き、同期後に自動で閉じる
  - 既存タブを使い回さないのは、リロードすると Streamlit の
    `st.session_state`（読み込み済み楽曲データ）が消えるため
  - アプリ未起動時に「同期した」と誤表示する問題も修正（事前に fetch で疎通確認）
  - **要確認**: `chrome://extensions` で拡張を再読み込みして実機テスト
- **Supabase 対応（DB を外に出す準備）** — ブランチ `feat/supabase-db`（詳細は下）

**DB まわりの変更点（重要）**

- 接続先は環境変数 `DATABASE_URL` で切り替え。未設定なら従来どおり
  `data/cyosakuken.db`（SQLite）。設定すれば Supabase（PostgreSQL）
- `app.py` は無改修。`modules/database.py` の 8 関数のシグネチャを変えていない
- スキーマ変更: 旧 `songs` / `nuendo_events`（列がそのままテーブル列で
  `ALTER TABLE ADD COLUMN` していた）→ 新 `song_rows` / `event_rows`
  （1行 = 1 JSON）。**旧テーブルは残してある**ので戻せる
- **Supabase 接続・移行・アプリ動作まで確認済み（2026-08-27）**。
  リージョンは `ap-northeast-2`（ソウル）、**Session pooler** 接続。
  Data API は無効のまま（Python から直接 PostgreSQL に繋ぐので不要）
- ローカルの実データは移行済み（楽曲 136 行 / 9 プロジェクト）。
  移行前のバックアップ: `data/cyosakuken.db.bak_before_migrate`
- **移行中に見つけて直した SQLite との差** ―― 今後 PostgreSQL 特有の
  問題が出たらまずここを疑う:
  - 全行が空の列を pandas が float と見なし、`NaN`（JSON に無い語）を
    書き出していた → 値ごとに None へ直す `_jsonable()`。
    `json.dumps(..., allow_nan=False)` で再発時は即例外にする
  - 日時が SQLite=文字列 / PostgreSQL=datetime で、`app.py:1110` の
    `p["updated_at"][:10]` が落ちた → `list_projects()` 側で文字列に揃える
- `scripts/check_db.py` … 接続確認、`scripts/migrate_db.py` … 移行（`--dry-run` あり）
- 手順書: `docs/Supabase設定手順.md`、雛形: `.env.example`

**前セッション（session 3）やったこと（記録）**

- **一括検索 詳細フィールド補完**:
  - J-WID 自動適用時に `fetch_jwid_detail(r["_detail_url"])` を呼んで作曲者・作詞者・編曲者・訳詞者を取得
    - 従来: J-WID 検索結果に `著作者名`（合体）しかなく `作曲者` は常に None → 未入力
    - 修正: 詳細ページを追加フェッチして各フィールドを分割取得
  - MINC マッチ時に `fetch_product_detail(_album_id, _track_id)` で委任者（集中管理 ○/×）を取得
  - I/V区分 判定を改善:
    - 作詞者あり（新規 or 既存）→「ヴォーカル」
    - J-WID 詳細取得済みで作詞者なし → 「インスト」（詳細ページで明示的に空が確認できた場合のみ）
    - None/nan/"none" も「未設定」扱いに修正
  - 推定所要時間: MINC使用時 8秒/曲、MINC未使用時 5秒/曲 に更新
  - 説明文を更新（編曲者・訳詞者・I/V区分 の取得も明示）
- **邦洋区分 自動判定**: JASRACコードの2文字目が数字→邦楽、英字→洋楽
  - 例: `188-2861-2`（2文字目=8）→ 邦楽、`0E6-6061-1`（2文字目=E）→ 洋楽
  - 一括検索でJASRACコードが判明した場合に自動設定（既存コードからも判定）

**前々セッション（session 2）やったこと（記録）**

- レイアウト変更: MINCログイン設定・セッション状態を表の上へ、プレビューを下expander、縮小トグル追加
- パイプラインにMINC追加: チェックボックス、実行後自動検索、「🌲 MINC 結果」タブ
- ステータス「該当なし」追加: 一括検索でヒットなし→「該当なし」
- 補完検索フィルターのデフォルトを「すべて」に変更
- 一括検索後に表が空になるバグ修正（stale session state の検出 + empty filter fallback）

**前前前セッション（session 1）やったこと（記録）**

- MINC `match=1` バグ修正（デフォルト `2: 前方一致`）、タイムアウト延長（30秒）、SEARCH_URL 末尾スラッシュ除去
- MINCログイン per-user 対応、セッション経過時間表示
- 検索語候補タブ、J-WID 手動検索 → 反映 expander
- 作詞作曲同一人対応、区切り文字 `/` 統一

**次にやること**

1. ~~Supabase 対応~~ **完了（2026-08-27）**
2. **Render デプロイ** ← 次はここ
   - 既知の制約: WAV/MP3 のフォルダ走査は不可（CSV アップロード経路を使う）、
     Playwright の MINC ログインも不可（Chrome 拡張の Cookie 同期に寄せ、
     `popup.js` の `APP_URL` を公開URLに書き換える）、無料枠は15分でスリープ
3. **① TSP CD データ取り込み**（管理番号 → 楽曲データ）
   - サンプル CSV の列名と、キーが `6ST-653-09` 形式かの確認待ち
   - 案: `tsp_cd` テーブル＋WEB結果キャッシュで「ローカル → キャッシュ → WEB」の3段
4. **② UI の刷新**（管理アプリ風）
   - Step1 テーマ+上部ステータスバー+ダッシュボード / Step2 `st.navigation` で
     ページ分割 / Step3 マスター詳細（`st.dataframe` の行選択）
5. 実機で一括検索（作曲者/作詞者/編曲者/訳詞者/I/V区分/委任者）を動作確認

**未解決の問題・懸念点**

- NexTone のみマッチでJASRACコードなし → 邦洋区分が判定できない場合がある
- MINC `match=1` の正確な挙動は MINC 側の仕様でありドキュメント化されていない
- NexTone は Next.js 製のため内部 API パス変更で scraper が壊れる可能性
- J-WID eJwid は検索語を URL パラメータで渡せない可能性（POST 送信のため）
