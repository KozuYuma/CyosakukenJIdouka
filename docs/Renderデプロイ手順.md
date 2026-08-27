# Render デプロイ手順

GitHub のリポジトリを Render に繋いで、ブラウザから使えるようにする手順。
`master` に入ったものが配られる。作業は別の枝でやり、固まったら合流させる。

---

## 1. Render の画面での設定

<https://dashboard.render.com> → **New +** → **Web Service** →
**Build and deploy from a Git repository** → `KozuYuma/CyosakukenJIdouka` を選ぶ。

（初回は GitHub との連携許可を求められる。このリポジトリだけ許可すればよい）

次の画面で埋める項目:

| 項目 | 値 |
|---|---|
| Name | `cyosakuken-app`（好きな名前。URL になる） |
| Region | **Singapore**（日本から一番近い） |
| Branch | **master** |
| Root Directory | 空欄のまま |
| Runtime | Python 3 |
| Build Command | `pip install -r requirements.txt` |
| Start Command | 下記 |
| Instance Type | Free（後から上げられる） |

**Start Command**（1行で貼る）

```
streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true
```

- `$PORT` は Render が決める。決め打ちにしないこと
- `0.0.0.0` で待たないと外から繋がらない
- `--server.headless true` が無いと、起動時にブラウザを開こうとして止まる

Python のバージョンはリポジトリの `.python-version`（3.14）が使われるので、
画面で指定しなくてよい。

**Advanced** を開いて設定するもの:

| 項目 | 値 |
|---|---|
| Health Check Path | `/_stcore/health` |

Streamlit が用意している生存確認の口。教えておくと、起動できていないときに
Render 側で気づける。

---

## 2. 環境変数（Environment）

**秘密なので、この2つは必ず Render の画面で入れる。ファイルには書かない。**

| キー | 値 |
|---|---|
| `DATABASE_URL` | Supabase の接続文字列（Session pooler の URI）。手元の `.env` と同じもの |
| `APP_USERS` | `ID:PASSWORD,ID:PASSWORD` … 手元の `.env` と同じもの |

`APP_USERS` を入れ忘れると**ログイン画面が出ず、誰でも入れる状態**で公開される。
必ず最初のデプロイの前に入れること。

`DATABASE_URL` が空だとサーバー内の SQLite を使う。Render の無料プランは
再起動のたびに中身が消えるので、必ず入れること。

---

## 3. 動いたか確かめる

1. デプロイのログに `You can now view your Streamlit app` が出る
2. `https://cyosakuken-app.onrender.com` を開く
3. ログイン画面が出る → ID と PASSWORD で入れる
4. 「🗃️ 共有楽曲データ」タブで、Supabase の中身が見える

---

## 4. 更新のしかた

`master` に push すると自動でデプロイされる。

```
git checkout master
git merge feat/xxxxx
git push origin master
```

---

## 無料プランで気をつけること

- **15分使わないと寝る。** 次に開いた人は起きるまで30〜60秒待つ
- **メモリ 512MB。** 200MB の CSV を上げるような使い方は落ちる。
  普段の Cue CSV / WAV 一覧（数MB）なら問題ない
- **保存先はサーバーに残らない。** DB は Supabase にあるので影響しないが、
  サーバー上に書いたファイルは再起動で消える

## クラウドでは動かない機能

手元の Windows でしか動かないものがある。エラーが出ても壊れてはいない。

- **📂 フォルダをスキャン** — サーバーには利用者のフォルダが無い。
  「📄 CSV をアップロード」の方を使う
- **MINC のブラウザログイン** — 手元の Chrome と `H:\PROGRAM\search_music`
  を呼ぶので、サーバーでは動かない
- **MINC の Cookie 読み取り** — 手元の Chrome から読む仕組みのため
