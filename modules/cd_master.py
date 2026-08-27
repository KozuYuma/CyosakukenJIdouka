"""
自社CDの台帳（cd_master）を使って、楽曲まとめの空欄を埋める。

台帳は TSP から取り込んだ36万曲。人は直さない読み取り専用の資料。
育てていく共有楽曲データ（song_master）とは役割が違う。

当てるのは管理番号だけ。曲名では当てない。理由は database.cd_fetch に。
"""
from __future__ import annotations

import pandas as pd

from modules.database import cd_fetch
from modules.song_master import _cell, mark_status, norm_id

# 出典の名前。song_master.SRC_RANK に載せてある
SRC = "自社CD"

# 台帳の列 → 楽曲まとめの列
COLUMN_MAP: dict[str, str] = {
    "title": "曲名",
    "artist": "アーティスト",
    "composer": "作曲者",
    "cd_name": "CD名",
    "cd_no": "CD番号",
    "label": "レコード会社名",
    "jasrac": "JASRAC作品コード",
}


def keys_of(row) -> list[str]:
    """楽曲まとめの1行から、台帳を引く管理番号キーの候補を作る。

    台帳のキーは曲ごとの固定管理番号（1AN-001-01）。案件の CSV には
    盤番号だけ（1AN-001）で入っていることもあるので、トラック番号を
    足した形も候補にする。

    「盤番号の末尾がトラック番号と同じか」では判断しない。1AN-001 と
    トラック01 のように、盤番号の末尾がたまたま一致してしまうため。
    候補を全部投げて、台帳にあった方を使う。
    """
    mgmt = norm_id(row.get("元管理番号"))
    if not mgmt:
        return []

    out = [mgmt]
    track = norm_id(row.get("トラック番号"))
    if track:
        for t in (track, track.zfill(2)):
            cand = f"{mgmt}{t}"
            if cand not in out:
                out.append(cand)
    return out


def fill(songs_df: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    """台帳の値で空欄を埋める。(埋めた後の df, 当たった曲数, 埋めた欄数)。

    埋めるのは空欄だけ。既に入っている値は触らない。
    """
    if songs_df is None or songs_df.empty:
        return songs_df, 0, 0

    keys = [keys_of(row) for _, row in songs_df.iterrows()]
    found = cd_fetch({k for cands in keys for k in cands})
    if not found:
        return songs_df, 0, 0

    by_key = {r["mgmt_key"]: r for r in found}

    df = songs_df.copy()
    hit_rows = 0
    filled = 0
    for pos, cands in enumerate(keys):
        # 候補は書いてあったとおりの番号が先。台帳にある方を採る
        hit = next((by_key[k] for k in cands if k in by_key), None)
        if hit is None:
            continue
        idx = df.index[pos]
        touched = False
        # 当てているのは管理番号だけなので、当たった行はどの曲かが決まる
        if mark_status(df, idx):
            filled += 1
            touched = True
        for src_col, dst_col in COLUMN_MAP.items():
            if dst_col not in df.columns:
                continue
            if _cell(df.at[idx, dst_col]):
                continue
            val = _cell(hit.get(src_col))
            if not val:
                continue
            df.at[idx, dst_col] = val
            filled += 1
            touched = True
        if touched:
            hit_rows += 1

    return df, hit_rows, filled
