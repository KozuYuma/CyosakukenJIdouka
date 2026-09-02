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

- [ ] **ステップ6: UI構造**（`st.navigation` / `st.Page`）← 次にやること
- [ ] **本番 Supabase に TSP 台帳を投入**: `python scripts/import_tsp.py <フォルダ>`。
      10〜30分・約130MB増。入れるまでサーバー側では台帳補完が効かない
- [ ] **`?sync_minc=` に合言葉**: 受け取り口がログインの前にあり、URL を知って
      いれば誰でも Cookie を送り込める（共有の Cookie を壊せる）
- [ ] **migration 管理の方針決め**: 今は `init_db()` の冪等 DDL だけ。版番号も
      戻す手段もドリフト検知も無い。`schema_version` 表か Alembic か
- [ ] **MINC ログインセッション UI**: メール入力欄が空のとき `.env` の値を使うことの確認・ドキュメント化

### 優先度: 中

- [ ] **`song_master.make_keys` の `endswith` 判定**: `cd_master.keys_of` で直した
      のと同じ穴が残っている。直すと既存のキーが変わるので要判断

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

### 最終更新: 2026-09-03（session 8）

**やったこと（session 8）** — ブランチ `feat/ui-foundation`（`master` にも反映済み＝本番に配られている）

テーマは **「検索結果を人が絞り込めるようにする」「間違ったものを勝手に入れない」**。
利用者から挙がった9件の要望を①から順に片づけた。

- **① 管理番号を手で足して台帳から埋められるように**（`c7539bc`）
  - 申告フォーマットの表に「管理番号」欄（元管理番号）を出して、その場で
    書けるようにした。書いたあと「🗃️ データベースで空欄を埋める」を押すと、
    共有楽曲データと TSP 台帳を引き直す
  - 番号の正規化は既存の `norm_id` をそのまま使う。盤番号まで（`1AN-001`）で
    書いた場合はトラック番号を足した形でも引く。**埋めるのは空欄だけ**

- **② 2枚組の2枚目の収録曲からも逆引き**（`41dba3e`）
  - MINC の CD商品詳細は2枚組だと収録曲の表が盤ごとに分かれている。
    `select_one` で最初の表しか読んでいなかったのが原因。全部読むようにした
  - 品番も盤ごとにあるので出てきた順に控え、曲ごとに「ディスク」と
    「その盤の品番」を持たせた。表に反映すると、**その曲が入っている方の盤の
    CD番号**が入る（一括のCD情報反映も同じ）

- **③ MINC検索を JASRAC作品コードでも絞れるように**（`eb8bfb9`）
  - ハイフン・空白を無視した部分一致。前の方だけ書いても効く
  - **初期値は空**。行の番号を勝手に入れると、番号を付け直したくて検索した
    ときにその番号の候補しか出なくなる

- **④ MINC検索の絞り込み項目を増やした**（`0a95440`）
  - 先頭20件しか描いていないのに断りが無く「400件出ているのに見られない」
    状態だった。「🔎 さらに絞り込む」を追加（取得済みの候補をその場で絞る
    だけなので通信しない）
  - 曲名／CD名・品番／作家名／レコード会社／どの表から出た候補か／
    作品コードの有無／出す件数（20・50・100・すべて）／解除

- **⑤ 管理番号が無くても、トラック番号＋曲名で台帳から埋める**（`00d10e1`）
  - 「1曲目・オープニング」のような重なりが1万3千種類あるので、勝手に別の盤の
    曲を入れないよう決まりを置いた。行のCD番号・CD名・アーティストでまず絞る
    （空になる条件は表記違いとみなして使わない）→ それでも複数残るときは
    **全部の当たりが同じことを言っている欄だけ**埋める → 60件超のキーは諦める
  - 確認ステータス「台帳一致（曲名）」を新設（管理番号の「台帳一致」より弱い
    当たりなので分けてある）。`track_key` に索引を追加

- **⑥ MINC検索で前の検索結果が灰色の欄に残るのを直した**（`96b3c8d`）
  - **原因**: 候補ごとの欄の `key` が候補番号だけで作られていた。Streamlit は
    `key` の付いた欄の中身を覚えていて、次に描くときは渡した `value` より
    **覚えている値を優先する**。だから検索し直しても前の候補1の値が残る
  - 検索するたびに上がる版（`mf_ver_◯`）を `key` に混ぜ、古い版の覚え書きは
    捨てる（`_mf_reset_result_state`）。CD商品リストの「収録曲を表示」の
    覚え書きも同じ形にそろえた
  - 捨てる側の正規表現の末尾 `_\d+` は**外さないこと**。外すと
    `selected_no=1` のときに `mf_results_12`（別の行の結果）まで消える

- **⑦ チェックした行が表の真ん中に来るように**（`afcae63`）
  - チェックを付け替えるたびに表の `key` を変えて作り直している（表側が持って
    いる編集中の印を消すため）。**作り直すと表の中のスクロールが先頭に戻る**
    ので、下の方の行にチェックを入れるとその行を見失っていた
  - 選んだ行を控え、描き直したあとでその行が中ほどに来る位置までスクロールを
    戻す（`_TABLE_SCROLL_JS`）。表の中身は canvas なので行が要素になって
    いない → 行の高さは「中身の高さ ÷ 行数」から割り出す。描き上がりが遅れる
    ので 80/250/600/1200 ミリ秒で数回試す
  - タブの裏にも表があるので、`<a id="sec-shinkok-table">` を目印にして
    「この目印より後ろの表」を狙っている

- **⑧ NexTone で同じ題名の別の曲を出さない**（`ad27bfa`）
  - NexTone は曲名（`freeWord`）でしか引けず、結果の表に JASRAC コードの欄も
    無い。そこで**作家名で見分ける**。J-WID の検索結果の「著作者名」と
    NexTone の「作曲者」がそろう行だけを同じ曲とみなす
  - 残るのは「JASRAC にも NexTone にもある曲」と「NexTone にしかない曲
    （J-WID が0件で比べる相手がいない）」の2通り
  - `modules/scraper.py` に `split_nextone_same_work()` を追加、作家名の照合は
    `composer_matches` として公開。**作家名が空の行は判断材料が無いので捨てない**
  - 一斉検索・一括は外した件数を伝言に出す。個別のNexToneタブは既定で隠し、
    「別の曲らしい候補も出す」の印で出せる（並びが変わるので `key` に印を混ぜる）

- **⑨ CD情報検索を CD名・CD番号でも絞れるように**（`25ef484`）
  - 1曲が何十枚ものCDに入っていることがあり、アーティストだけでは減らせなかった。
    CD名（CD商品タイトル）と CD番号（品番）の欄を足して事前に効かせる
  - 品番はハイフン・全角の有無を気にせず比べる（`_norm_hinban`）。
    **どの条件も、一致が0件になるときは全件表示に戻す**（打ち間違いで表ごと
    消えないようにするため）
  - JASRACコードも曲名も分からない曲のために「🔢 品番のCDを直接開く」を追加。
    `search_cds_by_hinban` をそのまま使う。作品名・作品コードが無い一覧になるので
    見出しは「品番◯◯で引いたCD商品」

**同じ期間に入れた、上の①〜⑨より前の直し**

- `d1c4c70` CD情報検索の3つの不具合 / `f3e9ff3` コード無しの曲を拾えるようにし、
  逆引きを両方に置いた
- `2d7e515` Cue CSV: タブ区切りの NUENDO 書き出しで見出し行が曲として並ぶのを直した /
  `f351f54` 時間順に並べ、申告フォーマットからトラック列を外した
- `91d4adc` **再実行のたびに DB へ16回問い合わせていたのをやめた**
  （`init_db` を `cache_resource` で process ごと1回に、`list_projects` と
  `master_search` は使い回して直した・消したときに捨てる）。1往復175ミリ秒
- `0a7cfd3` 一括検索を module 直下の `_run_bulk_search()` に切り出し、
  **画面を描き終えたスクリプト末尾で回す**ようにした（描きかけのまま何分も
  置かれて画面が二重に見えていた）
- `eef82d0` 申告フォーマットの「選択」を常に1つか0個にする（⑦の下地）
- `e41be84` 使い方メモ.txt と 使い方.md を削除 / `008e6e0` docx を Git の対象外に

**覚えておくこと**

- **session 7 の項に書いてある短縮ハッシュ（`77f147e` など）はもう存在しない**。
  持ち出してはいけないファイルを消すために履歴を書き直したため。同じ内容は
  `9828d76` / `a2a1914` / `c7aa654` / `4cd3dee` / `e3d1c6e` / `971dabf` にある
- `pytest` は入っていない。確認は scratchpad に置いた単体のスクリプトで行う。
  走らせ方は `PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe <ファイル>`
  （**`python` 単体は使わない**）
- ヒアドキュメント（`<<'PY'`）の中に書いたバックスラッシュは、Python に届く前に
  潰れて本物の改行になる。`BS = chr(92)` と差し替え用の記号を使うか、行番号を
  指定して直す

**残っていること**
- GitHub リポジトリの公開設定（利用者が Settings → Danger Zone で切り替える。
  `gh` は入っていない）
- Render の `APP_USERS` の合言葉を今のものから変える
- 本番 Supabase への TSP 台帳の投入（session 6 から持ち越し）
- 先読み（prefetch）… `scraper._rate_limit` と MINC `_throttle` に鍵が要る
- `?sync_minc=` に合言葉を付ける（今は URL を知っていれば誰でも送れる）
- migration 管理の方針決め（今は `init_db` の中で毎回 DDL を流している）
- **ステップ6 UI構造（`st.navigation`）は利用者が取り下げた。着手しないこと**
- **台帳入れ替えの無停止化もやらない**（利用者判断）

### 最終更新: 2026-08-29（session 7）

**やったこと（session 7）** — ブランチ `feat/ui-foundation`（`master` にも反映済み）

テーマは **「保存でデータが消えないようにする」**。複数人が同時に触っても
壊れない形に直した。全部 `master` に push 済み＝本番に配られている。

- **MINC の Cookie 保存を「新しい方が残る」に**（`77f147e`）
  - `minc_state_put` に `WHERE minc_state.updated_at < EXCLUDED.updated_at` を
    足した。読む側（`musicforest._pull_state`）は「手元の方が新しければ
    触らない」なのに、書く側は無条件に上書きしていて向きが逆だった
  - 手元のPCとサーバーの時計がずれていると古い方が残りうる。そこは承知の上

- **キーの作り方を、書き方の揺れに強くした**（`95d4e8a`）
  - 旧 `make_keys` は `endswith` で判定していたため、`1AN-001` ＋ トラック
    `01` のような行が盤番号だけのキーになり、**同じ盤の別の曲がひとつに
    まとまる**恐れがあった（ステップ4で「残っている穴」と書いていたもの）
  - `number_parser` はどの経路でも `元管理番号 = 盤番号-トラック番号` を作り、
    `ライブラリ盤番号` を別に持っている。**この列を使う**。文字列の末尾を
    削って盤番号を推測してはいけない（`1AN001` → `1AN0` になる）
  - トラック番号は2桁に揃える（`1` と `01` を同じ曲にする）。**新しく作る
    キーは先頭の候補だけ**、昔の形は「探す候補」として残すので移行は不要

- **ファイル名キー `file_key` を足した**（`5762677`）
  - 「同じ曲名・同じトラック番号でも、CD が違えば別の曲」を分けるため。
    強さは **管理番号 → ファイル名 → 曲名＋トラック番号** の順
  - 使うのは `WAV一致ファイル名` / `MP3一致ファイル名`。**`NUENDOファイル名`
    は使わない**（案件ごとに付け替えられるため）。拡張子・全角半角・空白を
    そろえ、**6文字未満は捨てる**（`01.wav` で無関係の曲が繋がるため）
  - 曲名＋トラック番号で当てるときは、**ファイル名が食い違う行は捨てる**
    （`_track_hit`）。貯まっている側にファイル名が無ければ当て、あとで
    書き足す。これで昔の行もそのまま使える
  - `init_db` の SQLite 用 `ALTER TABLE` は **本体の DDL より先**に流すこと。
    後にすると、足したばかりの列を使う索引で「そんな列は無い」と落ちる

- **共有楽曲データの保存を1トランザクションに**（`9fe1a9d`）
  - `master_merge`（database.py）に「引く→混ぜる→書く」をまとめた。
    PostgreSQL では読むときに `SELECT ... FOR UPDATE` で行を押さえる
  - **まだ無い行は押さえられない**ので、`ux_song_master_mgmt` /
    `ux_song_master_file`（空でない行だけの UNIQUE）で DB 側でも止める。
    はじかれたらトランザクションごとやり直す（2周目は相手の行が見えるので
    混ぜる方に回る。最大3回）
  - **`track_key` に UNIQUE は張らない**。音源違いで重なるのが正常なため
  - 既に重なった行があって索引を張れないときは、起動を止めず今までどおり
    動かす。本番の重複は事前に数えて **0件**だった（26行中）

- **管理タブの手直しに、上書き前の確認を足した**（`115aa85`）
  - 編集画面は開いたときの内容を丸ごと書き戻すので、開いている間に他の人が
    直していると**その直しを黙って消していた**
  - 書く前に同じトランザクションで読み直し、更新時刻**と中身**が開いたときの
    ままか確かめる。違えば書かずに `STALE`(-1) を返し、画面に
    「⚠️ 他の人が直したか消しました」＋「🔄 読み直す」を出す
  - 更新時刻は秒までしか持っていないので、時刻だけでは同じ秒の衝突を
    見逃す。だから中身も見比べている

- **案件の保存にも同じ見張りを付けた**（このコミット）
  - 案件は**行の全置換**なので、同時に開かれると相手の行が丸ごと消える
  - `save_project(pid, songs, events, seen_at)` を新設。楽曲とイベントを
    1つのトランザクションで置き換え、**保存後の更新時刻を返す**。
    呼ぶ側はそれを「開いたときの時刻」として持ち直す
    （`st.session_state["project_seen_at"]`、`_remember_project`）
  - 先に保存されていたら `ProjectChanged` を出して**何も書かない**。
    💾ボタンでは「⚠️ それでも上書きする」「📂 相手の内容を読み込み直す」の
    2択を出す。自動保存（ダウンロード時・照合直後）は黙って見送り、
    画面に残っている内容は消さない
  - `seen_at` を渡さなければ今までどおり書ける（`scripts/migrate_db.py` 用）

- **試験**（scratchpad の `t_*.py` 11本、すべて NG=0）
  - 新規: `t_keys` / `t_master_io` / `t_migrate_filekey` / `t_master_tx` /
    `t_master_edit` / `t_project_save` / `t_minc_state`
  - 走らせ方: `PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe <ファイル>`
    （コンソールが cp932 なので指定が要る。**`python` 単体は使わない**）

- **触ってはいけないもの**（利用者からの指示）
  - `AllCD_DATA__Tbl_All_CD_Data.csv` / `TSP_CD_DATA__Tbl_a.csv` /
    `TSP_CD_DATA__Tbl_a_player.csv` / `TSP_CD_DATA__Tbl_a_kaisya.csv` は
    **上書きしない**。`cd_master` は読み取り専用の資料として扱う
  - `使い方メモ.txt` は利用者の私物。**コミットしない**

**やったこと（session 6）** — ブランチ `feat/ui-foundation`

進め方: 「土台を先、構造を後」。この順で1つずつ進めている。
**0. UI土台 ✅ / 1. ログイン＋所有者分け ✅ / 2. 共有楽曲データ song_master ✅ /
3. その管理タブ ✅ / 4. TSP CD 取り込み ✅（本番への投入はまだ） /
5. Render デプロイ ✅ / 6. UI構造（st.navigation）← 次**

**公開先: <https://cyosakuken-app.onrender.com>（`master` に push すると自動で配られる）**

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
- **ステップ1: ログイン＋案件の所有者分け**（`24fb802`, `f00cd90`, `372635c`）
  - `modules/auth.py` 新規。環境変数 `APP_USERS=ID:PASSWORD,ID2:PASSWORD2`。
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
  - 入力欄に **`autocomplete` を付けてはいけない**。ブラウザの自動入力は
    欄に文字を表示しても Streamlit に値が届かないことがあり、
    見た目は埋まっているのに「違います」になる（一度これで詰まった）。
    失敗時は受け取った**文字数だけ**を出して切り分けられるようにしてある
  - ID もプルダウンにしない。一覧から選ばせると、ログインしていない人にも
    全員の ID が見えてしまうため
  - **`.env` はプロセス起動時に一度だけ読む**。書き換えたら起動し直すこと

- **ステップ2: 共有楽曲データ song_master**（`fc5a762`）
  - `modules/song_master.py` 新規。案件をまたいで楽曲データを貯め、
    次からは自動で埋める。`song_master` テーブルは案件に属さない
  - 貯めるのは **「確定」「作曲者一致」「アーティスト一致」の行だけ**。
    要確認や未調査まで貯めると、間違いが全案件に広がる
  - 一致は **「管理番号が一致」または「トラック番号と曲名の両方が一致」**。
    曲名だけでは当てない（同じ曲名の別の曲が多いため）
  - 旧形式の管理番号は盤番号で終わることがある（`57A-0023`）。そのままだと
    同じ盤の別の曲がひとつにまとまるので、トラック番号を足して分ける
  - 出典に強さ（`SRC_RANK`）を付けて強い方が弱い方を上書きする。
    手入力・確定は何が来ても動かさない
  - ただし**画面側の songs_df は空欄だけを埋める**。どの値がどこから来たかを
    DataFrame が覚えていないので、埋まっている値には触らない
  - PostgreSQL は `JSONB`（`dict` が返る）、SQLite は `TEXT`（`str` が返る）。
    読むときに両方を受ける。`IN :param` は `bindparam(..., expanding=True)`
    が要る（付けないとタプルが1つの値として扱われる）
- **照合実行で ID3 の結果が消えていた穴**（`fc5a762`）
  - DB への保存が `_import_mp3finder_id3` より**先**に走っていたので、
    ID3 から取れた作曲者などが保存されずに消えていた。保存を最後に移した
- **Cue CSV から楽曲以外を拾わない**（`fc5a762`）
  - 書き出しによっては**全部の行に列数ぶんのカンマが付く**
    （`トラック - M2-1,,,,,,,,,,`）。トラック名の見出しを「カンマの無い行」で
    見分けていたため、見出しが曲として並んでいた。
    **末尾の空欄を落としてから**見て、中身が1つだけの行は見出しとして扱う
  - 「マ ー カ ー ト ラ ッ ク リ ス ト」のように**一文字ずつ空けた見出し**も
    拾い、空白を詰めた名前をトラック名にする（`_is_spaced_title`）
  - 構成・ノートパッド・マーカー・ビデオのトラックは既定で外す。
    選択欄には残すので必要なら戻せる。黙って減らさないため
  - イベント名に `1khz` が入る基準信号の行は落とす
  - 手元の Cue 4本で確認。`Cue2_楽曲情報付き.csv` は 47件→42件（見出し5件が
    消えた）、他の3本は結果が変わらないこと

- **ステップ3: 共有楽曲データの管理タブ**（`7d9f07b`）
  - 🗃️ タブで中身を見る・直す・消す。出典と更新時刻も出す
  - 検証は3段構え（SQLite の単体・本番 Supabase への往復・AppTest で
    空のときと入っているときの2画面）
  - **AppTest の落とし穴**: ログインが通ると `auth.py` が `login_password` を
    捨てて再実行するが、AppTest の部品表にはログイン欄が残ったままなので、
    次の `.run()` が `KeyError: login_id` で落ちる。試験では
    `APP_USERS=""` にしてログイン画面を出さないこと
  - `describe_backend()` が `sqlite:///` を渡しても既定のパスを出していたのを
    修正。試験で別の DB を指したのに気づけず、手元の DB を汚したかと
    誤解する元になる

- **ステップ4: 自社CD台帳（TSP 36万曲）**（`eaa3980`）
  - `cd_master` テーブル（読み取り専用）＋ `modules/tsp_import.py` ＋
    `scripts/import_tsp.py` ＋ `modules/cd_master.py`
  - **`song_master` に混ぜない**。人が育てる表と、丸ごと入れ替える資料は
    役割が違う。JSON にせず普通の列にしてある（手元の SQLite で 97MB）
  - **当てるのは管理番号だけ。曲名では当てない**。「トラック番号＋曲名」は
    台帳の中で1万3千種類も重なっており、別の盤の曲を掴む
  - 盤番号だけの古い形にはトラック番号を足した候補も投げる。
    **「末尾が一致するか」で判断してはいけない**（`1AN-001` ＋ トラック
    `01` が漏れる）。候補を全部投げて、台帳にあった方を採る
  - **`song_master.make_keys` には同じ穴が残っている**（`endswith` 判定）。
    直すと既に貯めたキーが変わるので、意図的に触っていない
  - 埋まり方: アーティスト・CD名 100% / 曲名・CD番号・レコード会社 99% /
    作曲者 98% / **JASRAC作品コード 64%**（残りはメドレー等でコードが無い）
  - 照合の順は 共有楽曲データ → 台帳。人が直した値を先に入れさせる
  - **本番 Supabase への投入はまだ**（`python scripts/import_tsp.py <フォルダ>`。
    10〜30分・約130MB増）

- **ステップ5: Render デプロイ**（`db7344f`, `e0922ec`）
  - 手順は `docs/Renderデプロイ手順.md`。`render.yaml` は使わない
    （画面から GitHub リポジトリを繋ぐ方式）
  - `.python-version` に `3.14`
  - Start Command は `$PORT` と `0.0.0.0` と `--server.headless true` が要る。
    Health Check Path は `/_stcore/health`
  - 秘密（`DATABASE_URL` / `APP_USERS`）は Render の画面にだけ入れる。
    **`APP_USERS` を入れ忘れると誰でも入れる状態で公開される**
  - `IS_LOCAL_WINDOWS = os.name == "nt"` で手元専用の機能を止める。
    MINC ログイン・Chrome 同期のボタンは無効化して理由を出し、
    フォルダスキャンは先に案内を出す。**落とすのではなく理由を出す**
  - 無料プラン: 15分で寝る（次の1回は起きるのに30〜60秒）・メモリ512MB・
    **サーバーに書いたファイルは再起動で消える**

- **MINC の Cookie を DB に置いた**（`25252e3`）
  - MINC は **reCAPTCHA があるので自動ログインできない**。サーバーで
    Playwright を動かしても解決しない。人が取った Cookie を運ぶしかない
  - `minc_state` テーブル（`name` / `state` / `updated_at`）。`name` は
    Cookie ファイルの名前。手元は利用者ごと、サーバーは全員で1つ、という
    使い分けがそのまま行の分かれ方になる。
    **裏を返すと、サーバーに同期しても手元のアプリには反映されない**
  - Render には `MINC_STATE_PATH=/tmp/minc_state.json` を入れてある
    （既定は `H:\PROGRAM\search_music\auth\state.json` 決め打ちのため）
  - 保存する2箇所で DB にも入れ、読むときに DB の方が新しければ書き戻す。
    書き戻したら **mtime を DB の時刻に合わせる**（画面の「45分前」が狂う）
  - DB が使えなくても例外にしない。ファイルには書けているので今つないで
    いる人はそのまま使える
  - 画面の状態表示は再実行のたびに通るので、DB に聞く間隔を60秒空ける。
    実際に MINC へ繋ぐ直前（`load_client` / `check_session`）だけ `force=True`
  - **3時間で切れるのは MINC 側の仕様**。貼り直しは無くならない
  - 再起動しても「接続済み」のままになることを実機で確認済み

- **Chrome 拡張を Render 対応に**（`738e668`）
  - 同期先を入力欄にした（既定は Render の URL、手元なら `localhost:8501`）。
    `chrome.storage.local` に覚える
  - 送る前に `/_stcore/health` を叩いて**最大90秒起こす**。今までは起きる前に
    タブを閉じて「同期しました」と嘘をついていた
  - 時間切れのときは成功と出さず「届いたか分からない」と出す
  - 受け取り口（`?sync_minc=`）は `_save_cookies_to_state` を呼ぶので、
    **拡張から届いた Cookie も自動で DB に入る**
  - **受け取り口はログインの前にある**（裏タブがログインを通れないため）。
    結果として URL を知っている人なら誰でも Cookie を送り込める。
    合言葉で塞ぐのは今後の課題
  - 拡張は GitHub からは更新されない。`chrome://extensions` で再読み込みが要る
  - 実機で同期成功を確認済み（2026-08-27）

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
