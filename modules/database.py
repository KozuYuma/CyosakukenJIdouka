"""
SQLite データベースモジュール
プロジェクト・楽曲まとめ・イベント一覧の永続化
DB ファイル: data/cyosakuken.db（アプリルート直下）
"""
import sqlite3
from pathlib import Path

import pandas as pd

DB_PATH = Path(__file__).parent.parent / "data" / "cyosakuken.db"


def _conn() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """起動時に呼ぶ。必要なテーブルがなければ作成する。"""
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS projects (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                name        TEXT NOT NULL,
                description TEXT DEFAULT '',
                created_at  TEXT DEFAULT (datetime('now','localtime')),
                updated_at  TEXT DEFAULT (datetime('now','localtime'))
            )
        """)


# ─── プロジェクト ───────────────────────────────────────

def list_projects() -> list[dict]:
    """全プロジェクトを更新日時の降順で返す。"""
    try:
        with _conn() as conn:
            rows = conn.execute(
                "SELECT id, name, description, created_at, updated_at "
                "FROM projects ORDER BY updated_at DESC"
            ).fetchall()
            return [dict(r) for r in rows]
    except Exception:
        return []


def create_project(name: str, description: str = "") -> int:
    """新規プロジェクトを作成してIDを返す。"""
    with _conn() as conn:
        cur = conn.execute(
            "INSERT INTO projects (name, description) VALUES (?, ?)",
            (name.strip(), description.strip()),
        )
        return cur.lastrowid


def delete_project(project_id: int) -> None:
    """プロジェクトと関連データをすべて削除する。"""
    with _conn() as conn:
        for tbl in ("songs", "nuendo_events"):
            try:
                conn.execute(f"DELETE FROM {tbl} WHERE project_id = ?", (project_id,))
            except sqlite3.OperationalError:
                pass
        conn.execute("DELETE FROM projects WHERE id = ?", (project_id,))


def _touch_project(project_id: int) -> None:
    with _conn() as conn:
        conn.execute(
            "UPDATE projects SET updated_at = datetime('now','localtime') WHERE id = ?",
            (project_id,),
        )


# ─── 楽曲まとめ ─────────────────────────────────────────

def save_songs(project_id: int, songs_df: pd.DataFrame) -> None:
    """songs_df をDBに保存する（プロジェクト単位で全置換、列追加で自動スキーマ移行）。"""
    df = songs_df.copy()
    df.insert(0, "project_id", project_id)
    df = df.where(pd.notna(df), other=None)  # NaN → None (SQLite NULL)

    col_defs   = ", ".join(f'"{c}" TEXT' for c in df.columns)
    cols       = ", ".join(f'"{c}"' for c in df.columns)
    ph         = ", ".join("?" for _ in df.columns)
    insert_sql = f'INSERT INTO songs ({cols}) VALUES ({ph})'
    rows       = [tuple(row) for row in df.itertuples(index=False, name=None)]

    conn = _conn()
    try:
        # スキーマ確保（DDL は即時コミット）
        conn.execute(f'CREATE TABLE IF NOT EXISTS songs ({col_defs})')
        existing = {row[1] for row in conn.execute("PRAGMA table_info(songs)").fetchall()}
        for col in df.columns:
            if col not in existing:
                conn.execute(f'ALTER TABLE songs ADD COLUMN "{col}" TEXT')
        conn.commit()

        # 当プロジェクトを削除して再挿入
        conn.execute("DELETE FROM songs WHERE project_id = ?", (project_id,))
        conn.executemany(insert_sql, rows)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    _touch_project(project_id)


def load_songs(project_id: int) -> pd.DataFrame | None:
    """DBから songs_df を復元する。データがなければ None。"""
    try:
        with _conn() as conn:
            df = pd.read_sql(
                'SELECT * FROM songs WHERE project_id = ? ORDER BY "No"',
                conn,
                params=(project_id,),
            )
        return df.drop(columns=["project_id"], errors="ignore")
    except Exception:
        return None


# ─── NUENDOイベント ──────────────────────────────────────

def save_events(project_id: int, events_df: pd.DataFrame) -> None:
    """events_df をDBに保存する（プロジェクト単位で全置換）。"""
    df = events_df.copy()
    df.insert(0, "project_id", project_id)
    df = df.where(pd.notna(df), other=None)

    col_defs   = ", ".join(f'"{c}" TEXT' for c in df.columns)
    cols       = ", ".join(f'"{c}"' for c in df.columns)
    ph         = ", ".join("?" for _ in df.columns)
    insert_sql = f'INSERT INTO nuendo_events ({cols}) VALUES ({ph})'
    rows       = [tuple(row) for row in df.itertuples(index=False, name=None)]

    conn = _conn()
    try:
        conn.execute(f'CREATE TABLE IF NOT EXISTS nuendo_events ({col_defs})')
        existing = {row[1] for row in conn.execute("PRAGMA table_info(nuendo_events)").fetchall()}
        for col in df.columns:
            if col not in existing:
                conn.execute(f'ALTER TABLE nuendo_events ADD COLUMN "{col}" TEXT')
        conn.commit()

        conn.execute("DELETE FROM nuendo_events WHERE project_id = ?", (project_id,))
        conn.executemany(insert_sql, rows)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def load_events(project_id: int) -> pd.DataFrame | None:
    """DBから events_df を復元する。"""
    try:
        with _conn() as conn:
            df = pd.read_sql(
                "SELECT * FROM nuendo_events WHERE project_id = ?",
                conn,
                params=(project_id,),
            )
        return df.drop(columns=["project_id"], errors="ignore")
    except Exception:
        return None
