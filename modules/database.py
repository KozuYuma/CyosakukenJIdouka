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
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

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
    ]


def init_db() -> None:
    """起動時に呼ぶ。必要なテーブルがなければ作成する。"""
    with get_engine().begin() as conn:
        for stmt in _ddl():
            conn.execute(text(stmt))

    if not is_postgres():
        # SQLite に ADD COLUMN IF NOT EXISTS は無い。既に列があれば
        # "duplicate column name" で失敗するだけなので、それを握りつぶす。
        # 失敗しても他の DDL を巻き込まないよう、接続を分けて実行する。
        try:
            with get_engine().begin() as conn:
                conn.execute(text(
                    "ALTER TABLE projects ADD COLUMN owner TEXT DEFAULT ''"
                ))
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
