#!/usr/bin/env python3
"""
旧スキーマ（列がそのままテーブル列になっていた songs / nuendo_events）から
新スキーマ（1行1JSON の song_rows / event_rows）へデータを移す。

同じスクリプトで SQLite → Supabase の移行にも使う。移行先は
modules/database.py と同じく DATABASE_URL で決まる。

    # ① 旧テーブル → 新テーブル（同じ SQLite の中で）
    .venv\\Scripts\\python.exe scripts\\migrate_db.py

    # ② ローカル SQLite → Supabase（DATABASE_URL を設定した状態で）
    .venv\\Scripts\\python.exe scripts\\migrate_db.py --from-sqlite

    # 何が起きるか見るだけ
    .venv\\Scripts\\python.exe scripts\\migrate_db.py --dry-run

旧テーブルは削除しない。うまくいかなければ元のまま残っている。
"""
import argparse
import sqlite3
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from modules import database as db  # noqa: E402


def _read_legacy(sqlite_path: Path) -> tuple[list[dict], dict, dict]:
    """旧 SQLite から プロジェクト / 楽曲 / イベント を読む。"""
    if not sqlite_path.is_file():
        raise SystemExit(f"旧DBが見つかりません: {sqlite_path}")

    conn = sqlite3.connect(str(sqlite_path))
    conn.row_factory = sqlite3.Row
    try:
        tables = {r[0] for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'").fetchall()}

        projects = []
        if "projects" in tables:
            projects = [dict(r) for r in conn.execute(
                "SELECT id, name, description FROM projects ORDER BY id").fetchall()]

        songs, events = {}, {}
        for tbl, dest in (("songs", songs), ("nuendo_events", events)):
            if tbl not in tables:
                continue
            df = pd.read_sql(f"SELECT * FROM {tbl}", conn)
            if "project_id" not in df.columns or df.empty:
                continue
            for pid, grp in df.groupby("project_id"):
                dest[int(pid)] = grp.drop(columns=["project_id"])
        return projects, songs, events
    finally:
        conn.close()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--from-sqlite", action="store_true",
                    help="移行元を必ずローカル SQLite にする（移行先が Supabase のとき）")
    ap.add_argument("--dry-run", action="store_true", help="書き込まずに件数だけ出す")
    args = ap.parse_args()

    src = db.DB_PATH
    print(f"移行元: {src}")
    print(f"移行先: {db.describe_backend()}")
    if not args.from_sqlite and db.is_postgres():
        raise SystemExit(
            "移行先が PostgreSQL です。ローカル SQLite から移すなら --from-sqlite を付けてください。")

    projects, songs, events = _read_legacy(src)
    print(f"\nプロジェクト {len(projects)} 件 / "
          f"楽曲を持つ {len(songs)} 件 / イベントを持つ {len(events)} 件")
    for p in projects:
        s = len(songs.get(p["id"], []))
        e = len(events.get(p["id"], []))
        print(f"  [{p['id']}] {p['name']}  楽曲{s}行 / イベント{e}行")

    if args.dry_run:
        print("\n--dry-run のため書き込みません。")
        return

    db.init_db()

    # 移行先に既にあるプロジェクトを使い回して二重登録を避ける。
    # 同名のプロジェクトが複数ある実データがあるため（例: TEST111 が5件）、
    # 名前ではなく ID を先に見る。同じDB内の移行では projects テーブルを
    # そのまま使うので ID が一致し、確実に元の対応が保たれる。
    dst = db.list_projects()
    by_id   = {p["id"]: p for p in dst}
    by_name: dict[str, int] = {}
    for p in dst:
        by_name.setdefault(p["name"], p["id"])   # 同名は先頭のみ

    id_map: dict[int, int] = {}
    for p in projects:
        if p["id"] in by_id and by_id[p["id"]]["name"] == p["name"]:
            id_map[p["id"]] = p["id"]
            print(f"  同じID: [{p['id']}] {p['name']}")
        elif p["name"] in by_name:
            id_map[p["id"]] = by_name[p["name"]]
            print(f"  名前で対応: {p['name']} → id={id_map[p['id']]}")
        else:
            id_map[p["id"]] = db.create_project(p["name"], p.get("description") or "")
            print(f"  作成: {p['name']} → id={id_map[p['id']]}")

    moved = 0
    for old_id, df in songs.items():
        if old_id not in id_map:
            print(f"  [警告] project_id={old_id} のプロジェクト行が無いので楽曲{len(df)}行を飛ばします")
            continue
        db.save_songs(id_map[old_id], df)
        moved += len(df)
    for old_id, df in events.items():
        if old_id not in id_map:
            continue
        db.save_events(id_map[old_id], df)

    print(f"\n完了: 楽曲 {moved} 行を移しました。旧テーブルはそのまま残しています。")


if __name__ == "__main__":
    main()
