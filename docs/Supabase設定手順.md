# Supabase 設定手順

このアプリのデータ（プロジェクト・楽曲まとめ・NUENDOイベント）を
ローカルの SQLite ファイルではなく Supabase（PostgreSQL）に置くための手順。

Render にデプロイするとファイルは再起動のたびに消えてしまうため、
先に DB を外に出しておく必要がある。

---

## 仕組み（先に読むと迷わない）

接続先は環境変数 `DATABASE_URL` ひとつで切り替わる。

| `DATABASE_URL` | 接続先 |
|---|---|
| 未設定 / 空 | ローカルの SQLite（`data/cyosakuken.db`）＝ 従来どおり |
| PostgreSQL の接続文字列 | Supabase |

`app.py` は一切変更していない。`modules/database.py` の中だけで切り替えている。
つまり **設定しなければ今までと完全に同じ動作** で、いつでも戻せる。

---

## 手順

### 1. Supabase でプロジェクトを作る

1. https://supabase.com/dashboard を開く（GitHub アカウントでログイン済み）
2. **New project**
3. 入力するもの
   - **Name**: `cyosakuken`（何でもよい）
   - **Database Password**: 自動生成でよい。**このパスワードは後で使うので必ず控える**
     （後から見ることはできない。忘れた場合は再発行になる）
   - **Region**: `Northeast Asia (Tokyo)` を選ぶと速い
4. **Create new project** → 準備完了まで 1〜2 分待つ

### 2. 接続文字列を取得する

1. 画面上部の **Connect** ボタン
2. **Session pooler** のタブを選ぶ
   - Direct connection は IPv6 のみで、家庭・社内回線からは繋がらないことが多い
   - Transaction pooler は一部の SQL が使えないため、こちらは避ける
3. 表示される URI をコピーする。こんな形:

   ```
   postgresql://postgres.xxxxxxxxxxxx:[YOUR-PASSWORD]@aws-0-ap-northeast-1.pooler.supabase.com:5432/postgres
   ```

4. `[YOUR-PASSWORD]` を **1. で控えたパスワード** に置き換える

> パスワードに `@ : / ? # &` などの記号が含まれていると URL として壊れる。
> その場合は Supabase の Settings → Database → **Reset database password** で、
> 英数字だけのパスワードに変えてしまうのが早い。

### 3. `.env` に書く

プロジェクト直下（`app.py` と同じ場所）に `.env` を作る。

```
copy .env.example .env
```

`.env` を開いて `DATABASE_URL=` の後ろに 2. の文字列を貼る。

```
DATABASE_URL=postgresql://postgres.xxxx:実際のパスワード@aws-0-ap-northeast-1.pooler.supabase.com:5432/postgres
```

`.env` は `.gitignore` 済みなので GitHub には上がらない。
**接続文字列はパスワードそのものなので、チャットやメール、Slack に貼らないこと。**

### 4. 繋がるか確かめる

```
.venv\Scripts\python.exe scripts\check_db.py
```

`PostgreSQL（aws-0-ap-northeast-1.pooler.supabase.com:5432）` と出れば成功。
`SQLite（...）` と出たら `.env` が読まれていない（置き場所かファイル名を確認）。

### 5. 今のデータを Supabase へ移す

```
.venv\Scripts\python.exe scripts\migrate_db.py --from-sqlite --dry-run
.venv\Scripts\python.exe scripts\migrate_db.py --from-sqlite
```

ローカルの `data/cyosakuken.db` は消さずに残る。うまくいかなければ
`.env` の `DATABASE_URL` を空にするだけで元に戻る。

### 6. アプリを起動して確認

```
run.bat
```

プロジェクト一覧に今までのプロジェクトが出ていれば完了。

---

## 補足

- **無料枠**: 500MB。このアプリの用途（1プロジェクト数十行）なら当分埋まらない。
  ただし **7日間まったくアクセスが無いと一時停止** される。
  ダッシュボードから復帰できるが、Render に載せたあとは定期アクセスがあるので通常は止まらない。
- **テーブル**: `projects` / `song_rows` / `event_rows` の3つが自動で作られる。
  Supabase の Table Editor から中身を直接見られる。
- **行の持ち方**: 1行 = 1 JSON（`data` 列）。列名に日本語や記号があっても壊れず、
  アプリ側で列が増えても DB の作り直しが要らない。
- **Row Level Security**: 現状このアプリは接続文字列を持つ人だけがアクセスする前提。
  Supabase の RLS は使っていない。将来 Web 公開して複数人で使う場合は要検討。
