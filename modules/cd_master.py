"""
自社CDの台帳（cd_master）を使って、楽曲まとめの空欄を埋める。

台帳は TSP から取り込んだ36万曲。人は直さない読み取り専用の資料。
育てていく共有楽曲データ（song_master）とは役割が違う。

まず管理番号で当てる。番号は曲ごとに固有なので、当たればどの曲かが
決まる。番号で当たらなかった曲は、次にトラック番号＋曲名で当てにいく。
ただしこの組み合わせは曲を1つに決められない（「1曲目・オープニング」
のような重なりが1万3千種類ある）ので、当たりが1つに絞れた欄だけ埋め、
絞れなかった行には「複数候補あり」と書いて人に回す。
"""
from __future__ import annotations

import re

import pandas as pd

from modules.database import cd_fetch, cd_fetch_by_track
from modules.song_master import (
    LEDGER_OVERWRITABLE, _cell, mark_status, norm_id, norm_title,
    track_candidates,
)

# 出典の名前。song_master.SRC_RANK に載せてある
SRC = "自社CD"

#: トラック番号＋曲名で当たった行に書く確認ステータス。
#: 管理番号で当たった「台帳一致」より弱い当たりなので、後から
#: 見分けられるように分けてある
TRACK_STATUS = "台帳一致（曲名）"

#: 当たりが複数あって1つに決められなかった行に書く確認ステータス
AMBIGUOUS_STATUS = "複数候補あり"

#: 台帳から値を入れた行に書き残す出どころ。表では 🔵 の印になる
SOURCE_COL = "情報元"
SOURCE_LABEL = "TSP台帳"

#: この確認ステータスの行は人が確定させたものなので、台帳でも上書きしない
CONFIRMED_STATUS = "確定"

#: 1つのキーに対して見比べる台帳の行数の上限。同名の曲が何百枚のCDに
#: 入っているようなキーは、どのみち1つに決められないので早めに諦める
MAX_CANDIDATES = 60

#: 管理番号の英字2文字がこれならヴォーカル。TSP の Excel 版と同じ決まりで、
#: それ以外は全部インストとして扱う。JI/FI/ST/AN の歌モノは人が直す
VOCAL_SERIES = ("VO", "VJ")

#: ハイフンを落とした管理番号（1EX-545-04 → 1EX54504）。
#: トラック番号が無い盤番号だけの形も拾う
_SERIES_RE = re.compile(r"^\d([A-Z]{2})\d{3}(?:\d{2})?$")

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


def iv_of(mgmt) -> str:
    """管理番号から I/V区分 を決める。決められなければ空文字。

    台帳を引かずに番号だけで決まるので、台帳に無い新規入庫の曲でも書ける。
    VO・VJ をヴォーカル、それ以外をインストとするのは TSP の Excel 版と
    同じ決まり。例外（サントラの歌モノなど）は人が直す前提。
    """
    m = _SERIES_RE.match(norm_id(mgmt))
    if not m:
        return ""
    return "ヴォーカル" if m.group(1) in VOCAL_SERIES else "インスト"


def houyo_of(jasrac) -> str:
    """JASRACコードの2文字目から邦洋区分を決める（数字→邦楽、英字→洋楽）。

    これも TSP の Excel 版と同じ決まり。例外があるので人が直す前提。
    """
    code = _cell(jasrac)
    if len(code) < 2:
        return ""
    c = code[1]
    if c.isdigit():
        return "邦楽"
    if c.isalpha():
        return "洋楽"
    return ""


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


def _mark(df: pd.DataFrame, idx, status: str) -> bool:
    """確認ステータスを書く。手つかずの行だけ。書いたら True。"""
    if "確認ステータス" not in df.columns:
        return False
    if _cell(df.at[idx, "確認ステータス"]) not in LEDGER_OVERWRITABLE:
        return False
    df.at[idx, "確認ステータス"] = status
    return True


def _same(a: str, b: str) -> bool:
    """台帳の値どうしを比べる。書き方の揺れは吸収する。"""
    return norm_title(a) == norm_title(b)


def _narrow(rows: list[dict], row) -> list[dict]:
    """既に分かっているCD情報で、台帳の当たりを絞る。

    CD番号 → CD名 → アーティスト の順に、行に値が入っていて、かつ
    それで1件以上残るときだけ絞る。絞った結果が空になるなら、その
    条件は当てにならない（表記が違う等）とみなして絞らない。
    """
    for col, src in (("CD番号", "cd_no"), ("CD名", "cd_name"),
                     ("アーティスト", "artist")):
        want = _cell(row.get(col))
        if not want:
            continue
        kept = [r for r in rows if _same(r.get(src), want)]
        if kept:
            rows = kept
    return rows


def _agreed(rows: list[dict], src_col: str) -> tuple[str, bool]:
    """台帳の当たりが、その欄で同じことを言っているか。

    返すのは (値, 揃っているか)。空欄は数えない。中身のある値が
    2種類以上あれば揃っていないとみなす。
    """
    vals: list[str] = []
    for r in rows:
        v = _cell(r.get(src_col))
        if v and not any(_same(v, x) for x in vals):
            vals.append(v)
    if not vals:
        return "", True
    return vals[0], len(vals) == 1


def _apply(df: pd.DataFrame, idx, rows: list[dict],
           overwrite: bool = False) -> tuple[int, bool]:
    """当たった台帳の行で欄を埋める。(埋めた欄数, 迷った欄があるか)。

    当たりが複数あっても、その欄について全員が同じことを言っている
    なら埋めてよい。言うことが割れている欄だけ空けておく。

    overwrite を立てると、既に入っている値も台帳の値で置き換える。
    管理番号で当てたときだけ使う。番号は曲ごとに固有で、盤まわりの
    情報は台帳が元なので、外から拾ってきた値より確かなため。
    """
    filled = 0
    conflicted = False
    for src_col, dst_col in COLUMN_MAP.items():
        if dst_col not in df.columns:
            continue
        val, agreed = _agreed(rows, src_col)
        if not agreed:
            conflicted = True
            continue
        if not val:
            continue
        cur = _cell(df.at[idx, dst_col])
        if cur and not (overwrite and not _same(cur, val)):
            continue
        df.at[idx, dst_col] = val
        filled += 1
    return filled, conflicted


def _confirmed(df: pd.DataFrame, idx) -> bool:
    """人が確定させた行か。確定した行は台帳でも上書きしない。"""
    if "確認ステータス" not in df.columns:
        return False
    return _cell(df.at[idx, "確認ステータス"]) == CONFIRMED_STATUS


def _mark_source(df: pd.DataFrame, idx) -> None:
    """この行の値は台帳から入れた、と書き残す。表の 🔵 の印はこれを見る。"""
    if SOURCE_COL not in df.columns:
        df[SOURCE_COL] = ""
    df.at[idx, SOURCE_COL] = SOURCE_LABEL


def fill(songs_df: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    """台帳の値で欄を埋める。(埋めた後の df, 当たった曲数, 埋めた欄数)。

    管理番号で当たった行は、台帳の値を優先して入れ替える（人が
    「確定」にした行は触らない）。番号が無くてトラック番号＋曲名で
    当てた行は、今までどおり空欄だけを埋める。

    まず管理番号で当てる。当たらなかった行は、トラック番号＋曲名で
    もう一度当てにいく（管理番号が書かれていない曲を拾うため）。
    """
    if songs_df is None or songs_df.empty:
        return songs_df, 0, 0

    keys = [keys_of(row) for _, row in songs_df.iterrows()]
    found = cd_fetch({k for cands in keys for k in cands})
    by_key = {r["mgmt_key"]: r for r in found}

    df = songs_df.copy()
    hit_rows = 0
    filled = 0
    # 管理番号で当たらなかった行の位置。あとでトラック番号＋曲名を試す
    left: list[int] = []
    for pos, cands in enumerate(keys):
        # 候補は書いてあったとおりの番号が先。台帳にある方を採る
        hit = next((by_key[k] for k in cands if k in by_key), None)
        if hit is None:
            left.append(pos)
            continue
        idx = df.index[pos]
        touched = False
        # 当てているのは管理番号なので、当たった行はどの曲かが決まる
        if mark_status(df, idx):
            filled += 1
            touched = True
        # 管理番号で当たった行は、台帳の値を優先して入れ替える。
        # 人が確定させた行だけは触らない
        _f, _ = _apply(df, idx, [hit], overwrite=not _confirmed(df, idx))
        if _f:
            filled += _f
            touched = True
        if touched:
            _mark_source(df, idx)
            hit_rows += 1

    _h2, _f2 = _fill_by_track(df, left)
    _h3, _f3 = _fill_derived(df)
    return df, hit_rows + _h2 + _h3, filled + _f2 + _f3


def _fill_derived(df: pd.DataFrame) -> tuple[int, int]:
    """管理番号と JASRACコードから決まる欄を埋める。(当たった曲数, 埋めた欄数)。

    I/V区分 は管理番号だけで決まるので、台帳に当たらなかった行にも入れる。
    邦洋区分 は JASRACコードを見るので、台帳から番号が入った行で効く。

    どちらも決まりで機械的に決めているだけの推定なので、既に値が入って
    いる行は触らない（人が直した値を消さないため）。
    """
    hit_rows = 0
    filled = 0
    for pos in range(len(df)):
        idx = df.index[pos]
        row = df.iloc[pos]
        vals = (
            ("I/V区分", iv_of(row.get("元管理番号"))),
            ("邦洋区分", houyo_of(row.get("JASRAC作品コード"))),
        )
        n = 0
        for col, val in vals:
            if not val or col not in df.columns:
                continue
            if _cell(df.at[idx, col]):
                continue
            df.at[idx, col] = val
            n += 1
        if not n:
            continue
        filled += n
        # 台帳で既に当たっている行は、曲数を二重に数えない
        already = (SOURCE_COL in df.columns
                   and _cell(df.at[idx, SOURCE_COL]) == SOURCE_LABEL)
        if not already:
            hit_rows += 1
        _mark_source(df, idx)
    return hit_rows, filled


def _fill_by_track(df: pd.DataFrame, positions: list[int]) -> tuple[int, int]:
    """管理番号で当たらなかった行を、トラック番号＋曲名で埋める。

    (当たった曲数, 埋めた欄数)。df はその場で書き換える。

    このキーは曲を1つに決められないので、
      ・行に入っているCD番号・CD名・アーティストでまず絞る
      ・それでも複数残るときは、全員が同じことを言っている欄だけ埋める
      ・言うことが割れた欄が1つでもあれば「複数候補あり」と書いて残す
    という決まりにしてある。勝手に別の盤の曲を入れないため。
    """
    if not positions:
        return 0, 0

    # 行ごとのキー候補（桁をそろえた形と、そのままの形）
    cands: dict[int, list[str]] = {}
    for pos in positions:
        row = df.iloc[pos]
        ks = track_candidates(row)
        if ks:
            cands[pos] = ks
    if not cands:
        return 0, 0

    found = cd_fetch_by_track({k for ks in cands.values() for k in ks})
    if not found:
        return 0, 0

    hit_rows = 0
    filled = 0
    for pos, ks in cands.items():
        rows: list[dict] = []
        seen: set[str] = set()
        for k in ks:
            for r in found.get(k, []):
                mk = str(r.get("mgmt_key") or "")
                if mk not in seen:
                    seen.add(mk)
                    rows.append(r)
        if not rows or len(rows) > MAX_CANDIDATES:
            continue

        idx = df.index[pos]
        rows = _narrow(rows, df.iloc[pos])
        _f, conflicted = _apply(df, idx, rows)
        # 迷った欄があった行は、人がどれか選びに行く必要がある
        if _mark(df, idx, AMBIGUOUS_STATUS if conflicted else TRACK_STATUS):
            _f += 1
        if _f:
            _mark_source(df, idx)
            filled += _f
            hit_rows += 1

    return hit_rows, filled
