"""
データベースモジュール
プロジェクト・楽曲まとめ・イベント一覧の永続化。

接続先は環境変数 DATABASE_URL で切り替える:
    未設定  → SQLite（data/cyosakuken.db）。従来どおりローカルで完結する。
    設定あり → PostgreSQL（Supabase / Render）。

app.py から見える関数（init_db / list_projects / create_project /
delete_project / save_songs / load_songs / save_events / load_events）の
呼び出し方は従来と同じなので、app.py 側の変更は不要。

--- スキーマについて ---
旧実装は songs_df の列がそのままテーブルの列で、列が増えるたびに
ALTER TABLE ADD COLUMN していた。列名に日本語・記号が混ざるうえ、
バージョン間で列構成が変わるとテーブルが壊れやすい。
新実装では 1行 = 1 JSON にして、列構成をデータ側に持たせる。

旧テーブル（songs / nuendo_events）には触らない。移行は
scripts/migrate_db.py で行う。
"""
import json
import os
import threading
import unicodedata
from pathlib import Path

import pandas as pd
from sqlalchemy import bindparam, create_engine, text
from sqlalchemy.engine import Engine
from sqlalchemy.exc import IntegrityError

DB_PATH = Path(__file__).parent.parent / "data" / "cyosakuken.db"

_engine_lock = threading.Lock()
_engine: Engine | None = None


# ─── 接続 ──────────────────────────────────────────────

_env_loaded = False


def _load_local_env() -> None:
    """開発機の .env を読む。Render 等では環境変数が直接入る。

    以前は DATABASE_URL が既にあれば読まずに戻していたが、.env には
    ログインの合言葉（APP_USERS）など DB 以外の設定も入るようになった
    ので、一度だけ全部読むようにした。既にある環境変数は setdefault
    なので上書きしない（Render 側の設定が .env に負けることはない）。
    """
    global _env_loaded
    if _env_loaded:
        return
    _env_loaded = True
    for cand in (Path(__file__).parent.parent / ".env",
                 Path(r"H:\PROGRAM\search_music\.env")):
        if not cand.is_file():
            continue
        try:
            for line in cand.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
        except Exception:
            continue


def load_local_env() -> None:
    """.env を読む公開版。DB 以外（ログインの合言葉など）でも使う。"""
    _load_local_env()


def get_db_url() -> str:
    """接続文字列。空文字なら SQLite を使う。"""
    _load_local_env()
    url = os.environ.get("DATABASE_URL", "").strip()
    if not url:
        return ""
    # Supabase がコピーさせる形式は postgresql:// なので、
    # 使うドライバ（psycopg3）を明示した形に直す。
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    if url.startswith("postgresql://"):
        url = "postgresql+psycopg://" + url[len("postgresql://"):]
    return url


def is_postgres() -> bool:
    return get_db_url().startswith("postgresql")


def get_engine() -> Engine:
    """Engine は使い回す。Streamlit は再実行のたびにこの関数を通るため。"""
    global _engine
    if _engine is not None:
        return _engine
    with _engine_lock:
        if _engine is not None:
            return _engine
        url = get_db_url()
        if url:
            # pool_pre_ping: Supabase 側で切られた接続を掴んだまま
            # 実行してしまうのを防ぐ（放置後の最初の操作で効く）
            _engine = create_engine(url, pool_pre_ping=True, pool_size=3,
                                    max_overflow=2, pool_recycle=1800)
        else:
            DB_PATH.parent.mkdir(parents=True, exist_ok=True)
            _engine = create_engine(f"sqlite:///{DB_PATH}")
        return _engine


def describe_backend() -> str:
    """設定画面などに出す用の説明。接続文字列は秘密なので出さない。"""
    if not is_postgres():
        url = get_db_url()
        # 試験で sqlite:/// を直接指定することがある。その場合は
        # 既定の置き場所ではなく、実際に使うファイルを出す
        if url.startswith("sqlite"):
            return f"SQLite（{url.split('///')[-1]}）"
        return f"SQLite（{DB_PATH}）"
    host = get_db_url().split("@")[-1].split("/")[0] if "@" in get_db_url() else "?"
    return f"PostgreSQL（{host}）"


# ─── スキーマ ───────────────────────────────────────────

def _ddl() -> list[str]:
    if is_postgres():
        return [
            """CREATE TABLE IF NOT EXISTS projects (
                   id          SERIAL PRIMARY KEY,
                   name        TEXT NOT NULL,
                   description TEXT DEFAULT '',
                   owner       TEXT DEFAULT '',
                   created_at  TIMESTAMPTZ DEFAULT now(),
                   updated_at  TIMESTAMPTZ DEFAULT now()
               )""",
            # 既にテーブルがある場合の追加。IF NOT EXISTS があるので何度でも通る
            "ALTER TABLE projects ADD COLUMN IF NOT EXISTS owner TEXT DEFAULT ''",
            """CREATE TABLE IF NOT EXISTS song_rows (
                   project_id INTEGER NOT NULL,
                   row_no     INTEGER NOT NULL,
                   data       JSONB   NOT NULL,
                   PRIMARY KEY (project_id, row_no)
               )""",
            """CREATE TABLE IF NOT EXISTS event_rows (
                   project_id INTEGER NOT NULL,
                   row_no     INTEGER NOT NULL,
                   data       JSONB   NOT NULL,
                   PRIMARY KEY (project_id, row_no)
               )""",
            # 共有楽曲データ。案件に属さず、全員で使い回す
            """CREATE TABLE IF NOT EXISTS song_master (
                   id         SERIAL PRIMARY KEY,
                   mgmt_key   TEXT NOT NULL DEFAULT '',
                   track_key  TEXT NOT NULL DEFAULT '',
                   file_key   TEXT NOT NULL DEFAULT '',
                   title      TEXT NOT NULL DEFAULT '',
                   data       JSONB NOT NULL,
                   updated_at TIMESTAMPTZ DEFAULT now()
               )""",
            # 先に作られたテーブルにも足す
            "ALTER TABLE song_master ADD COLUMN IF NOT EXISTS "
            "file_key TEXT NOT NULL DEFAULT ''",
            # 空文字のキーは一致に使わないので、索引から外して軽くする
            "CREATE INDEX IF NOT EXISTS ix_song_master_mgmt "
            "ON song_master (mgmt_key) WHERE mgmt_key <> ''",
            "CREATE INDEX IF NOT EXISTS ix_song_master_track "
            "ON song_master (track_key) WHERE track_key <> ''",
            "CREATE INDEX IF NOT EXISTS ix_song_master_file "
            "ON song_master (file_key) WHERE file_key <> ''",
            # 自社CDの台帳（TSP）。人が育てる song_master とは別に持つ。
            # 元データを丸ごと入れ替えられるよう、JSON にせず普通の列にする
            """CREATE TABLE IF NOT EXISTS cd_master (
                   mgmt_key  TEXT PRIMARY KEY,
                   disc_key  TEXT NOT NULL DEFAULT '',
                   track_key TEXT NOT NULL DEFAULT '',
                   track_no  TEXT NOT NULL DEFAULT '',
                   title     TEXT NOT NULL DEFAULT '',
                   artist    TEXT NOT NULL DEFAULT '',
                   composer  TEXT NOT NULL DEFAULT '',
                   cd_name   TEXT NOT NULL DEFAULT '',
                   cd_no     TEXT NOT NULL DEFAULT '',
                   label     TEXT NOT NULL DEFAULT '',
                   jasrac    TEXT NOT NULL DEFAULT ''
               )""",
            # 作品コードから CD を引くための索引。これが無いと 36万行を
            # 端から見ることになる。作品コードの無い行は引かないので外す
            "CREATE INDEX IF NOT EXISTS ix_cd_master_jasrac "
            "ON cd_master (jasrac) WHERE jasrac <> ''",
            # MINC の Cookie。サーバーでは書いたファイルが再起動で消える
            # ため、DB を置き場所にする。updated_at はファイルの更新時刻
            # （epoch 秒）。どちらが新しいかを比べるのに使う
            """CREATE TABLE IF NOT EXISTS minc_state (
                   name       TEXT PRIMARY KEY,
                   state      TEXT NOT NULL,
                   updated_at DOUBLE PRECISION NOT NULL DEFAULT 0
               )""",
        ]
    return [
        """CREATE TABLE IF NOT EXISTS projects (
               id          INTEGER PRIMARY KEY AUTOINCREMENT,
               name        TEXT NOT NULL,
               description TEXT DEFAULT '',
               owner       TEXT DEFAULT '',
               created_at  TEXT DEFAULT (datetime('now','localtime')),
               updated_at  TEXT DEFAULT (datetime('now','localtime'))
           )""",
        """CREATE TABLE IF NOT EXISTS song_rows (
               project_id INTEGER NOT NULL,
               row_no     INTEGER NOT NULL,
               data       TEXT    NOT NULL,
               PRIMARY KEY (project_id, row_no)
           )""",
        """CREATE TABLE IF NOT EXISTS event_rows (
               project_id INTEGER NOT NULL,
               row_no     INTEGER NOT NULL,
               data       TEXT    NOT NULL,
               PRIMARY KEY (project_id, row_no)
           )""",
        """CREATE TABLE IF NOT EXISTS song_master (
               id         INTEGER PRIMARY KEY AUTOINCREMENT,
               mgmt_key   TEXT NOT NULL DEFAULT '',
               track_key  TEXT NOT NULL DEFAULT '',
               file_key   TEXT NOT NULL DEFAULT '',
               title      TEXT NOT NULL DEFAULT '',
               data       TEXT NOT NULL,
               updated_at TEXT DEFAULT (datetime('now','localtime'))
           )""",
        "CREATE INDEX IF NOT EXISTS ix_song_master_mgmt "
        "ON song_master (mgmt_key)",
        "CREATE INDEX IF NOT EXISTS ix_song_master_track "
        "ON song_master (track_key)",
        "CREATE INDEX IF NOT EXISTS ix_song_master_file "
        "ON song_master (file_key)",
        """CREATE TABLE IF NOT EXISTS cd_master (
               mgmt_key  TEXT PRIMARY KEY,
               disc_key  TEXT NOT NULL DEFAULT '',
               track_key TEXT NOT NULL DEFAULT '',
               track_no  TEXT NOT NULL DEFAULT '',
               title     TEXT NOT NULL DEFAULT '',
               artist    TEXT NOT NULL DEFAULT '',
               composer  TEXT NOT NULL DEFAULT '',
               cd_name   TEXT NOT NULL DEFAULT '',
               cd_no     TEXT NOT NULL DEFAULT '',
               label     TEXT NOT NULL DEFAULT '',
               jasrac    TEXT NOT NULL DEFAULT ''
           )""",
        "CREATE INDEX IF NOT EXISTS ix_cd_master_jasrac "
        "ON cd_master (jasrac)",
        """CREATE TABLE IF NOT EXISTS minc_state (
               name       TEXT PRIMARY KEY,
               state      TEXT NOT NULL,
               updated_at REAL NOT NULL DEFAULT 0
           )""",
    ]


def init_db() -> None:
    """起動時に呼ぶ。必要なテーブルがなければ作成する。"""
    if not is_postgres():
        # SQLite に ADD COLUMN IF NOT EXISTS は無い。既に列があれば
        # "duplicate column name"、テーブルがまだ無ければ "no such table"
        # で失敗するだけなので、それを握りつぶす。失敗しても他の DDL を
        # 巻き込まないよう、接続を分けて実行する。
        #
        # 本体の DDL より先に流す。あとに回すと、足したばかりの列を使う
        # 索引を作るところで「そんな列は無い」と言われてしまう
        for stmt in (
            "ALTER TABLE projects ADD COLUMN owner TEXT DEFAULT ''",
            "ALTER TABLE song_master ADD COLUMN file_key TEXT "
            "NOT NULL DEFAULT ''",
        ):
            try:
                with get_engine().begin() as conn:
                    conn.execute(text(stmt))
            except Exception:
                pass

    with get_engine().begin() as conn:
        for stmt in _ddl():
            conn.execute(text(stmt))

    _add_unique_indexes()


def _add_unique_indexes() -> None:
    """共有楽曲データのキーに、重なりを許さない索引を張る。

    同じ管理番号・同じファイル名の行が二重にできないよう、DB 側でも
    止める。二人が同時に同じ曲を貯めたときの取りこぼしを防ぐため。

    track_key には張らない。同じ曲名・同じトラック番号でも音源（CD）が
    違えば別の行にする決まりなので、重なるのが正常な状態のため。

    既に重なった行が入っていると失敗する。そのときは起動を止めず、
    今までどおりの索引のまま動かす（重なりはアプリ側で吸収できる）。
    """
    made = True
    for stmt in (
        "CREATE UNIQUE INDEX IF NOT EXISTS ux_song_master_mgmt "
        "ON song_master (mgmt_key) WHERE mgmt_key <> ''",
        "CREATE UNIQUE INDEX IF NOT EXISTS ux_song_master_file "
        "ON song_master (file_key) WHERE file_key <> ''",
    ):
        try:
            with get_engine().begin() as conn:
                conn.execute(text(stmt))
        except Exception:
            made = False
    if not made:
        return
    # 張れたら、同じ列を見ていた古い索引はもう要らない
    for stmt in ("DROP INDEX IF EXISTS ix_song_master_mgmt",
                 "DROP INDEX IF EXISTS ix_song_master_file"):
        try:
            with get_engine().begin() as conn:
                conn.execute(text(stmt))
        except Exception:
            pass


def _now_sql() -> str:
    return "now()" if is_postgres() else "datetime('now','localtime')"


# ─── DataFrame ⇔ 行 ────────────────────────────────────

def _jsonable(v):
    """欠損値を None（＝JSON の null）に直す。

    DataFrame ごと `where(notna)` で置き換える方法は使えない。列の中身が
    すべて空だと pandas はその列を float 型と見なし、None を入れても
    NaN に戻してしまうため。実際 `使用形態` などの未入力列がそうなり、
    JSON に `NaN` という不正な語が書き出されていた（SQLite は文字列として
    受け取るので気付けず、PostgreSQL の JSONB で初めて弾かれた）。
    そこで値ひとつずつ見る。
    """
    if v is None:
        return None
    if isinstance(v, float):
        # NaN と ±Inf は JSON に無い。どちらも「値なし」として扱う
        return None if (v != v or v in (float("inf"), float("-inf"))) else v
    try:
        # NaT や pd.NA もここで拾う。配列が入っていると例外になるので囲う
        if not isinstance(v, (list, tuple, dict, set)) and pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v


def _df_to_rows(df: pd.DataFrame, project_id: int) -> list[dict]:
    """DataFrame を 1行1JSON に変換する。欠損は None（＝JSON null）にする。"""
    rows = []
    for i, rec in enumerate(df.to_dict(orient="records")):
        rec = {k: _jsonable(v) for k, v in rec.items()}
        # 表の列名・値はそのまま保持する。ensure_ascii=False で日本語を潰さない。
        # allow_nan=False: 取りこぼしがあれば無言で不正な JSON を作らず例外にする
        rows.append({
            "project_id": project_id,
            "row_no": i,
            "data": json.dumps(rec, ensure_ascii=False, default=str,
                               allow_nan=False),
        })
    return rows


def _rows_to_df(records: list) -> pd.DataFrame | None:
    if not records:
        return None
    dicts = []
    for r in records:
        raw = r[0]
        # PostgreSQL の JSONB は dict で返り、SQLite の TEXT は文字列で返る
        dicts.append(raw if isinstance(raw, dict) else json.loads(raw))
    return pd.DataFrame(dicts)


def _replace_rows(table: str, project_id: int, df: pd.DataFrame) -> None:
    """プロジェクト単位で全置換する。"""
    rows = _df_to_rows(df, project_id)
    ins = text(f"INSERT INTO {table} (project_id, row_no, data) "
               f"VALUES (:project_id, :row_no, :data)")
    if is_postgres():
        # JSONB 列に文字列を入れるとエラーになるのでキャストする
        ins = text(f"INSERT INTO {table} (project_id, row_no, data) "
                   f"VALUES (:project_id, :row_no, CAST(:data AS JSONB))")
    with get_engine().begin() as conn:
        conn.execute(text(f"DELETE FROM {table} WHERE project_id = :pid"),
                     {"pid": project_id})
        if rows:
            conn.execute(ins, rows)


def _load_rows(table: str, project_id: int) -> pd.DataFrame | None:
    try:
        with get_engine().connect() as conn:
            res = conn.execute(
                text(f"SELECT data FROM {table} WHERE project_id = :pid "
                     f"ORDER BY row_no"),
                {"pid": project_id},
            ).fetchall()
        return _rows_to_df(res)
    except Exception:
        return None


# ─── プロジェクト ───────────────────────────────────────

def _as_text(v) -> str:
    """日時を "YYYY-MM-DD HH:MM:SS" の文字列にする。

    SQLite は文字列を返すが PostgreSQL は datetime を返す。app.py は
    `p["updated_at"][:10]` のように文字列として切り出しているので、
    ここで形を揃えて app.py 側を変えずに済ませる。
    """
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    return v.strftime("%Y-%m-%d %H:%M:%S")


def list_projects(owner: str | None = None) -> list[dict]:
    """プロジェクトを更新日時の降順で返す。

    owner を渡すとその人のものだけを返す。owner が空文字のプロジェクト
    （所有者を分ける前に作られたもの）は誰のものか分からないので、
    誰が見ても出す。取りこぼして「消えた」と思われる方が困るため。
    owner=None（既定）は絞り込みなし。
    """
    sql = ("SELECT id, name, description, owner, created_at, updated_at "
           "FROM projects")
    params: dict = {}
    if owner:
        sql += " WHERE owner = :owner OR owner IS NULL OR owner = ''"
        params["owner"] = owner
    sql += " ORDER BY updated_at DESC"
    try:
        with get_engine().connect() as conn:
            rows = conn.execute(text(sql), params).mappings().all()
        out = []
        for r in rows:
            d = dict(r)
            for k in ("created_at", "updated_at"):
                d[k] = _as_text(d.get(k))
            d["owner"] = d.get("owner") or ""
            out.append(d)
        return out
    except Exception:
        return []


def create_project(name: str, description: str = "", owner: str = "") -> int:
    """新規プロジェクトを作成してIDを返す。"""
    params = {"name": name.strip(), "description": description.strip(),
              "owner": (owner or "").strip()}
    with get_engine().begin() as conn:
        if is_postgres():
            return conn.execute(text(
                "INSERT INTO projects (name, description, owner) "
                "VALUES (:name, :description, :owner) RETURNING id"
            ), params).scalar_one()
        return conn.execute(text(
            "INSERT INTO projects (name, description, owner) "
            "VALUES (:name, :description, :owner)"
        ), params).lastrowid


def set_project_owner(project_id: int, owner: str) -> None:
    """所有者を付け替える。所有者が空のまま残ったものを引き取る用。"""
    with get_engine().begin() as conn:
        conn.execute(text("UPDATE projects SET owner = :owner WHERE id = :pid"),
                     {"owner": (owner or "").strip(), "pid": project_id})


def delete_project(project_id: int) -> None:
    """プロジェクトと関連データをすべて削除する。"""
    with get_engine().begin() as conn:
        for tbl in ("song_rows", "event_rows"):
            conn.execute(text(f"DELETE FROM {tbl} WHERE project_id = :pid"),
                         {"pid": project_id})
        conn.execute(text("DELETE FROM projects WHERE id = :pid"),
                     {"pid": project_id})


def _touch_project(project_id: int) -> None:
    with get_engine().begin() as conn:
        conn.execute(text(f"UPDATE projects SET updated_at = {_now_sql()} "
                          f"WHERE id = :pid"), {"pid": project_id})


# ─── 楽曲まとめ ─────────────────────────────────────────

def save_songs(project_id: int, songs_df: pd.DataFrame) -> None:
    """songs_df をDBに保存する（プロジェクト単位で全置換）。"""
    _replace_rows("song_rows", project_id, songs_df)
    _touch_project(project_id)


def load_songs(project_id: int) -> pd.DataFrame | None:
    """DBから songs_df を復元する。データがなければ None。"""
    return _load_rows("song_rows", project_id)


# ─── NUENDOイベント ──────────────────────────────────────

def save_events(project_id: int, events_df: pd.DataFrame) -> None:
    """events_df をDBに保存する（プロジェクト単位で全置換）。"""
    _replace_rows("event_rows", project_id, events_df)


def load_events(project_id: int) -> pd.DataFrame | None:
    """DBから events_df を復元する。"""
    return _load_rows("event_rows", project_id)


# ─── 共有楽曲データ ──────────────────────────────────────
# 案件に属さない。全員で貯めて全員で使う。中身の決め方は
# modules/song_master.py 側にある。ここは出し入れだけ。

# SELECT で並べる列。どこも同じ形で受け取れるように一箇所にまとめる
MASTER_COLUMNS = ("id, mgmt_key, track_key, file_key, title, data, "
                  "updated_at")


def _master_row(r) -> dict:
    """SELECT の1行を dict にする。data は PostgreSQL なら dict で返る。"""
    d = dict(r)
    raw = d.get("data")
    d["data"] = raw if isinstance(raw, dict) else json.loads(raw or "{}")
    d["updated_at"] = _as_text(d.get("updated_at"))
    return d


def _master_select(conn, mgmt: list[str], track: list[str], file_: list[str],
                   lock: bool = False) -> list[dict]:
    """どれかのキーに当たる行を、渡された接続で読む。

    lock=True なら、読んだ行を書き換えられないように押さえる
    （PostgreSQL の SELECT ... FOR UPDATE）。SQLite には無い書き方だが、
    そちらは書き込み自体が1つずつしか通らないので要らない。
    """
    where, params = [], {}
    if mgmt:
        where.append("mgmt_key IN :mgmt")
        params["mgmt"] = tuple(mgmt)
    if track:
        where.append("track_key IN :track")
        params["track"] = tuple(track)
    if file_:
        where.append("file_key IN :file")
        params["file"] = tuple(file_)
    if not where:
        return []

    sql = text(
        f"SELECT {MASTER_COLUMNS} "
        "FROM song_master WHERE " + " OR ".join(where)
        + (" FOR UPDATE" if lock else "")
    ).bindparams(*[
        # tuple をそのまま渡すと1個の値と見なされるので展開を指示する
        bindparam(k, expanding=True) for k in params
    ])
    return [_master_row(r) for r in conn.execute(sql, params).mappings().all()]


def _keys_of(records: list[dict], name: str) -> list[str]:
    """records から、候補キーを重複なく集める。"""
    out: list[str] = []
    for rec in records:
        for k in rec.get(name) or []:
            if k and k not in out:
                out.append(k)
    return out


def master_fetch(mgmt_keys: set[str], track_keys: set[str],
                 file_keys: set[str] | None = None) -> list[dict]:
    """どれかのキーに当たる行を返す。キーが空なら何も返さない。"""
    mgmt = [k for k in (mgmt_keys or set()) if k]
    track = [k for k in (track_keys or set()) if k]
    file_ = [k for k in (file_keys or set()) if k]
    if not mgmt and not track and not file_:
        return []
    try:
        with get_engine().connect() as conn:
            return _master_select(conn, mgmt, track, file_)
    except Exception:
        return []


def _master_write(conn, records: list[dict]) -> int:
    """渡された接続で、id があれば更新・なければ追加する。"""
    cast = "CAST(:data AS JSONB)" if is_postgres() else ":data"
    ins = text(
        f"INSERT INTO song_master (mgmt_key, track_key, file_key, title, data) "
        f"VALUES (:mgmt_key, :track_key, :file_key, :title, {cast})"
    )
    upd = text(
        f"UPDATE song_master SET mgmt_key = :mgmt_key, track_key = :track_key, "
        f"file_key = :file_key, title = :title, data = {cast}, "
        f"updated_at = {_now_sql()} WHERE id = :id"
    )
    n = 0
    for rec in records:
        p = {
            "mgmt_key": rec.get("mgmt_key") or "",
            "track_key": rec.get("track_key") or "",
            "file_key": rec.get("file_key") or "",
            "title": rec.get("title") or "",
            "data": json.dumps(rec.get("data") or {}, ensure_ascii=False,
                               default=str, allow_nan=False),
        }
        if rec.get("id"):
            conn.execute(upd, {**p, "id": rec["id"]})
        else:
            conn.execute(ins, p)
        n += 1
    return n


def master_upsert(records: list[dict]) -> int:
    """id があれば更新、なければ追加。書いた件数を返す。"""
    if not records:
        return 0
    with get_engine().begin() as conn:
        return _master_write(conn, records)


STALE = -1   # 読んでから書くまでの間に、他の人が先に直していた


def master_update_seen(record: dict, seen: dict) -> int:
    """読んだときから変わっていなければ書く。

    書けたら 1、他の人が先に直していたか行が消えていたら STALE(-1)。

    seen は編集画面を開いたときに読んだ行。更新時刻だけでなく中身も
    見比べる。更新時刻は秒までしか持っていないので、同じ秒に二人が
    直した場合を取りこぼさないため。
    """
    rid = record.get("id")
    if not rid:
        return 0
    sql = text(
        f"SELECT {MASTER_COLUMNS} FROM song_master WHERE id = :id"
        + (" FOR UPDATE" if is_postgres() else "")
    )
    with get_engine().begin() as conn:
        rows = conn.execute(sql, {"id": rid}).mappings().all()
        if not rows:
            return STALE
        cur = _master_row(rows[0])
        if (cur.get("updated_at") != _as_text(seen.get("updated_at"))
                or (cur.get("data") or {}) != (seen.get("data") or {})):
            return STALE
        return _master_write(conn, [record])


def master_merge(records: list[dict], decide, attempts: int = 3) -> int:
    """「引いて・混ぜて・書く」を1つのトランザクションで通す。

    records は song_master.collect が作る形（mgmt_cands などの候補を
    持つ）。decide(既にある行) が、書き込む行の一覧を返す。

    途中で他の人が同じ曲を書いても取りこぼさないよう、二段構えにする。

      1. 読むときに行を押さえる（PostgreSQL の FOR UPDATE）。
         押さえている間、その行は他の人から書き換えられない。
      2. まだ無い行は押さえようがないので、ほぼ同時に同じ曲が二度
         入ろうとすることがある。そこは DB の索引（ux_song_master_*）が
         はじくので、はじかれたら最初からやり直す。やり直したときは
         相手の書いた行が見えるので、今度は混ぜる方に回る。
    """
    if not records:
        return 0
    mgmt = _keys_of(records, "mgmt_cands")
    track = _keys_of(records, "track_cands")
    file_ = _keys_of(records, "file_cands")

    for attempt in range(attempts):
        try:
            with get_engine().begin() as conn:
                existing = _master_select(conn, mgmt, track, file_,
                                          lock=is_postgres())
                writes = decide(existing)
                if not writes:
                    return 0
                return _master_write(conn, writes)
        except IntegrityError:
            # 同じキーの行を、ほぼ同時に別の人が入れた
            if attempt == attempts - 1:
                raise
    return 0


def master_all() -> list[dict]:
    """全件。管理タブ（ステップ3）で使う。"""
    try:
        with get_engine().connect() as conn:
            rows = conn.execute(text(
                f"SELECT {MASTER_COLUMNS} "
                "FROM song_master ORDER BY updated_at DESC"
            )).mappings().all()
        return [_master_row(r) for r in rows]
    except Exception:
        return []


def master_search(keyword: str = "", limit: int = 300) -> list[dict]:
    """曲名・管理番号・中身の文字で探す。空なら新しい順に limit 件。

    件数が増えても全件を読まないよう、DB 側で絞って件数も切る。
    """
    limit = max(1, min(int(limit or 300), 2000))
    kw = str(keyword or "").strip()

    # JSONB はそのままだと LIKE が使えないので文字列にしてから見る
    data_as_text = "CAST(data AS TEXT)" if is_postgres() else "data"
    where = ""
    params: dict = {"lim": limit}
    if kw:
        where = (
            " WHERE title LIKE :kw OR mgmt_key LIKE :kw_id "
            f"OR track_key LIKE :kw OR file_key LIKE :kw "
            f"OR {data_as_text} LIKE :kw"
        )
        params["kw"] = f"%{kw}%"
        # 管理番号は記号を抜いた形で入っているので、探す方も揃える
        params["kw_id"] = f"%{_norm_key(kw)}%"

    try:
        with get_engine().connect() as conn:
            rows = conn.execute(text(
                f"SELECT {MASTER_COLUMNS} "
                "FROM song_master" + where +
                " ORDER BY updated_at DESC LIMIT :lim"
            ), params).mappings().all()
        return [_master_row(r) for r in rows]
    except Exception:
        return []


def _norm_key(value: str) -> str:
    """管理番号の検索語を、保存してある形（英数字だけ・大文字）に揃える。"""
    s = unicodedata.normalize("NFKC", str(value or "")).upper()
    return "".join(ch for ch in s if ch.isalnum())


def master_delete(ids) -> int:
    """共有楽曲データを消す。消した件数を返す。"""
    wanted = [int(i) for i in (ids or []) if str(i).strip()]
    if not wanted:
        return 0
    sql = text("DELETE FROM song_master WHERE id IN :ids").bindparams(
        bindparam("ids", expanding=True))
    try:
        with get_engine().begin() as conn:
            return int(conn.execute(sql, {"ids": tuple(wanted)}).rowcount or 0)
    except Exception:
        return 0


def master_count() -> int:
    """貯まっている曲数。"""
    try:
        with get_engine().connect() as conn:
            return int(conn.execute(
                text("SELECT COUNT(*) FROM song_master")).scalar_one())
    except Exception:
        return 0


# ─── 自社CDの台帳（cd_master） ─────────────────────────
#
# TSP から書き出した36万曲。人は直さない。元データが更新されたら
# scripts/import_tsp.py で丸ごと入れ替える。

CD_COLUMNS: tuple[str, ...] = (
    "mgmt_key", "disc_key", "track_key", "track_no", "title",
    "artist", "composer", "cd_name", "cd_no", "label", "jasrac",
)


def cd_count() -> int:
    """台帳に入っている曲数。テーブルがまだ無ければ 0。"""
    try:
        with get_engine().connect() as conn:
            return int(conn.execute(
                text("SELECT COUNT(*) FROM cd_master")).scalar_one())
    except Exception:
        return 0


def cd_clear() -> None:
    """台帳を空にする。入れ替えの前に呼ぶ。"""
    with get_engine().begin() as conn:
        conn.execute(text("DELETE FROM cd_master"))


def cd_insert(rows: list[dict]) -> int:
    """台帳に書き込む。同じ管理番号が来たら先に入っていた方を残す。

    元データには同じ固定管理番号が2回出てくることが少しだけある。
    どちらが正しいか判断できないので、黙って上書きせず先着を残す。
    """
    if not rows:
        return 0

    cols = ", ".join(CD_COLUMNS)
    vals = ", ".join(f":{c}" for c in CD_COLUMNS)
    # ON CONFLICT は PostgreSQL でも SQLite でも同じ書き方で通る
    sql = text(f"INSERT INTO cd_master ({cols}) VALUES ({vals}) "
               "ON CONFLICT (mgmt_key) DO NOTHING")

    payload = [{c: str(r.get(c) or "") for c in CD_COLUMNS} for r in rows]
    with get_engine().begin() as conn:
        conn.execute(sql, payload)
    return len(payload)


def cd_fetch(mgmt_keys: set[str]) -> list[dict]:
    """管理番号で台帳を引く。

    曲名＋トラック番号では引かない。台帳は36万曲あり、「1曲目・
    オープニング」のような組み合わせが1万3千種類も重なっているため、
    曲名で当てると別の盤の曲を掴んでしまう。
    """
    mgmt = [k for k in (mgmt_keys or set()) if k]
    if not mgmt:
        return []

    sql = text(
        f"SELECT {', '.join(CD_COLUMNS)} FROM cd_master "
        "WHERE mgmt_key IN :mgmt"
    ).bindparams(bindparam("mgmt", expanding=True))
    try:
        with get_engine().connect() as conn:
            return [dict(r) for r in
                    conn.execute(sql, {"mgmt": tuple(mgmt)}).mappings().all()]
    except Exception:
        return []


def jasrac_variants(code: str) -> list[str]:
    """作品コードの書き方の揺れを並べる。

    台帳の元データは「237-8679-5」とハイフン入りのことも、
    「23786795」と数字だけのこともある。どちらで来ても引けるよう、
    両方の形を作って照合する。8桁でなければ元の形と数字だけの形。
    """
    raw = str(code or "").strip()
    digits = "".join(ch for ch in raw if ch.isdigit())
    if not digits:
        return []
    out = [digits]
    if len(digits) == 8:
        # JASRAC の作品コードは 3-4-1 桁で区切って表示する
        out.append(f"{digits[:3]}-{digits[3:7]}-{digits[7:]}")
    if raw and raw not in out:
        out.append(raw)
    return out


def cd_fetch_by_jasrac(codes: set[str]) -> dict[str, list[dict]]:
    """JASRAC作品コードで台帳を引く。

    返すのは「数字だけにしたコード → その作品が入っている台帳の行」。
    同じ作品が複数の盤に入っていることがあるので、行は並べて返す。
    品番のある行を先に、その中は管理番号の順にする（毎回同じ順で
    返るようにするため。呼ぶ側が先頭を採っても揺れない）。

    台帳が空のとき・テーブルがまだ無いときは空を返す。ここで例外を
    投げると一括検索が丸ごと止まってしまうため。
    """
    want: dict[str, str] = {}   # 台帳に問い合わせる形 → 数字だけの形
    for c in codes or set():
        digits = "".join(ch for ch in str(c or "") if ch.isdigit())
        if not digits:
            continue
        for v in jasrac_variants(c):
            want[v] = digits
    if not want:
        return {}

    sql = text(
        f"SELECT {', '.join(CD_COLUMNS)} FROM cd_master "
        # jasrac <> '' は結果を変えない。PostgreSQL 側の索引が
        # 「作品コードのある行だけ」なので、条件を揃えて使わせる
        "WHERE jasrac <> '' AND jasrac IN :codes"
    ).bindparams(bindparam("codes", expanding=True))
    try:
        with get_engine().connect() as conn:
            rows = [dict(r) for r in
                    conn.execute(sql, {"codes": tuple(want)}).mappings().all()]
    except Exception:
        return {}

    out: dict[str, list[dict]] = {}
    for r in rows:
        key = want.get(str(r.get("jasrac", "")).strip())
        if not key:
            continue
        out.setdefault(key, []).append(r)
    for v in out.values():
        v.sort(key=lambda r: (not str(r.get("cd_no", "")).strip(),
                              str(r.get("mgmt_key", ""))))
    return out


# ─── MINC の Cookie（minc_state） ──────────────────────────
#
# MINC は reCAPTCHA があるので自動ログインできない。人がブラウザで
# ログインして取れた Cookie を持ち回るしかない。サーバーではファイルが
# 再起動で消えるので、DB を置き場所にする。
#
# name は Cookie ファイルの名前。手元は利用者ごとに別ファイル、サーバー
# は全員で1つ、という今の使い分けがそのまま行の分かれ方になる。

def minc_state_get(name: str) -> tuple[str, float] | None:
    """保存してある Cookie を (中身, 更新時刻) で返す。無ければ None。"""
    if not name:
        return None
    try:
        with get_engine().connect() as conn:
            row = conn.execute(
                text("SELECT state, updated_at FROM minc_state WHERE name = :n"),
                {"n": name},
            ).first()
    except Exception:
        return None
    if not row:
        return None
    return str(row[0]), float(row[1] or 0)


def minc_state_put(name: str, state: str, updated_at: float) -> None:
    """Cookie を保存する。同じ名前が既にあれば新しい方を残す。

    上書きは「今入っているものより新しいとき」だけにする。DB に入って
    いる値を読んでから書くまでの間に、別の人が貼り直していることが
    あるため。ここを素の上書きにすると、手元の古い Cookie が誰かの
    新しい Cookie を潰してしまう。持ち回りの向きを、読む側
    （musicforest._pull_state）と同じ「新しい方が勝つ」にそろえる。

    Cookie はログインの鍵そのものなので、失敗しても中身は例外に載せない。
    """
    if not name or not state:
        return
    with get_engine().begin() as conn:
        conn.execute(
            text("INSERT INTO minc_state (name, state, updated_at) "
                 "VALUES (:n, :s, :u) "
                 "ON CONFLICT (name) DO UPDATE SET "
                 "state = EXCLUDED.state, updated_at = EXCLUDED.updated_at "
                 "WHERE minc_state.updated_at < EXCLUDED.updated_at"),
            {"n": name, "s": state, "u": float(updated_at)},
        )
