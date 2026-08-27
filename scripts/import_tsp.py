"""
TSP の CSV を自社CDの台帳（cd_master）に取り込む。

使い方（プロジェクトの root で）:

    python scripts/import_tsp.py "C:/Users/User/Downloads/files"
    python scripts/import_tsp.py "C:/..." --dry-run   # 書かずに数えるだけ

接続先は DATABASE_URL に従う（未設定ならローカルの SQLite）。
何度実行しても同じ結果になる。中身は毎回まるごと入れ替える。
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from modules.database import (  # noqa: E402
    cd_clear,
    cd_count,
    cd_insert,
    describe_backend,
    init_db,
)
from modules.tsp_import import iter_records, missing_files  # noqa: E402

CHUNK = 5_000


def main(argv: list[str]) -> int:
    args = [a for a in argv[1:] if not a.startswith("--")]
    dry = "--dry-run" in argv
    if not args:
        print(__doc__)
        return 1

    folder = Path(args[0])
    if not folder.is_dir():
        print(f"フォルダがありません: {folder}")
        return 1
    lack = missing_files(folder)
    if lack:
        print("ファイルが足りません:", " / ".join(lack))
        return 1

    print("読み元:", folder)
    print("書き先:", describe_backend() if not dry else "（--dry-run なので書きません）")

    init_db()
    before = cd_count()
    print(f"いま台帳に入っている曲: {before:,}")

    if not dry:
        # 消してから入れる。古い曲が残り続けないようにするため
        cd_clear()

    t0 = time.time()
    total = 0
    buf: list[dict] = []
    for rec in iter_records(folder):
        buf.append(rec)
        if len(buf) >= CHUNK:
            if not dry:
                cd_insert(buf)
            total += len(buf)
            buf.clear()
            print(f"  {total:,} 曲 … {time.time() - t0:.0f} 秒", end="\r")
    if buf:
        if not dry:
            cd_insert(buf)
        total += len(buf)

    print(" " * 60, end="\r")
    print(f"読んだ曲: {total:,}（{time.time() - t0:.0f} 秒）")

    if dry:
        print("--dry-run なので書いていません。")
        return 0

    after = cd_count()
    print(f"台帳の曲: {after:,}")
    skipped = total - after
    if skipped:
        # 元データに同じ固定管理番号が2回出てくる分。先着を残している
        print(f"うち、管理番号が重なっていて入れなかった: {skipped:,} 曲")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
