#!/usr/bin/env python3
"""
DB に繋がるか確かめるだけのスクリプト。

    .venv\\Scripts\\python.exe scripts\\check_db.py

接続先・テーブルの有無・プロジェクト件数を出す。
接続文字列にはパスワードが入っているので画面には出さない。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from modules import database as db  # noqa: E402


def main() -> None:
    print(f"接続先: {db.describe_backend()}")
    if not db.is_postgres():
        print("  → .env の DATABASE_URL が未設定か、読めていません。")
        print("     Supabase を使うつもりなら docs/Supabase設定手順.md の 3. を確認してください。")

    try:
        db.init_db()
    except Exception as e:
        # ここで落ちるのはほぼ接続の問題なので、よくある原因を添える
        print(f"\n[失敗] 接続できません: {type(e).__name__}: {e}")
        print("  よくある原因:")
        print("   - パスワードの [YOUR-PASSWORD] を置き換えていない")
        print("   - パスワードに @ : / などの記号が入っていて URL が壊れている")
        print("   - Direct connection（IPv6）を選んでいる → Session pooler にする")
        print("   - Supabase プロジェクトが一時停止している（7日間未使用）")
        raise SystemExit(1)

    projects = db.list_projects()
    print("テーブル: OK")
    print(f"プロジェクト: {len(projects)} 件")
    for p in projects:
        s = db.load_songs(p["id"])
        e = db.load_events(p["id"])
        print(f"  [{p['id']}] {p['name']}  "
              f"楽曲{0 if s is None else len(s)}行 / "
              f"イベント{0 if e is None else len(e)}行")
    print("\n接続 OK")


if __name__ == "__main__":
    main()
