"""
共有楽曲データ（song_master）。

一度調べた曲の権利情報を貯めておき、次に同じ曲が来たら自動で埋める。
案件をまたいで全員で共有するので、誰かが調べた結果が全員に効く。

同じ曲と見なす条件（どちらか一方でも当たれば一致）:

  1. 管理番号が一致する
  2. トラック番号と曲名の両方が一致する

曲名だけでは当てない。同じ曲名の別物が多すぎるため。

値は列ごとに「出典」を持たせて貯める。強い出典の値が来たら弱い方を
上書きするが、人が手で入れた値（手入力・確定）は絶対に上書きしない。
"""
from __future__ import annotations

import unicodedata

import pandas as pd

from modules.database import (
    master_fetch,
    master_upsert,
)

# 人が管理タブで直したときの出典。PROTECTED_RANK なので機械に負けない
HAND_SRC = "手入力"

# 貯める列。使用形態・確認ステータス・メモは案件ごとに変わるので入れない
MASTER_FIELDS: tuple[str, ...] = (
    "作曲者",
    "作詞者",
    "編曲者",
    "訳詞者",
    "アーティスト",
    "レコード会社名",
    "CD番号",
    "CD名",
    "JASRAC作品コード",
    "NexTone管理番号",
    "委任者",
    "I/V区分",
    "邦洋区分",
    "原訳詞区分",
    "音源区分",
    "自社楽曲ID",
)

# 出典の強さ。大きいほど信用できる
SRC_RANK: dict[str, int] = {
    "手入力": 100,
    "確定": 100,
    # 自社の台帳。管理番号は元々ここから出ているので、盤まわりの情報
    # （CD番号・レコード会社名・収録曲）は外のどの資料より確かめやすい
    "自社CD": 90,
    "J-WID詳細": 80,
    "NexTone": 70,
    "MINC": 60,
    "MusicBrainz": 40,
    "曲名推定": 20,
}

# これ以上は上書きしない。人が手で入れた値を機械が壊さないため
PROTECTED_RANK = 100

# song_master に入れてよい確認ステータス。「一致」以上だけを貯める。
# 未調査や要確認のものを貯めると、間違いが全案件に広がる
TRUSTED_STATUS: dict[str, str] = {
    "確定": "確定",
    "作曲者一致": "J-WID詳細",
    "アーティスト一致": "MINC",
}

# 管理番号で台帳に当たった行に付ける確認ステータス。
#
# 管理番号は曲ごとに固有の番号なので、当たった時点でどの曲かは決まって
# いる。まだ「未調査」と出ていると、人が調べに行く先として毎回目に入って
# しまうので、当たったことを書いておく。
#
# TRUSTED_STATUS には入れない。台帳から入れた値を、そのまま台帳に貯め
# 直すことになるため（同じ値が出典だけ強くなって戻ってくる）。
LEDGER_STATUS = "台帳一致"

# 台帳の当たりで上書きしてよい確認ステータス。人や検索が何か書いた行は
# 触らない
LEDGER_OVERWRITABLE = ("", "未調査")


def rank(src: str) -> int:
    """出典の強さ。知らない出典は一番弱く扱う。"""
    return SRC_RANK.get(str(src or "").strip(), 10)


# ─── 正規化 ────────────────────────────────────────────

def norm_id(value) -> str:
    """管理番号を比べる形にする。全角・空白・ハイフンの揺れを吸収する。"""
    s = unicodedata.normalize("NFKC", str(value or "")).upper()
    return "".join(ch for ch in s if ch.isalnum())


def norm_title(value) -> str:
    """曲名を比べる形にする。

    全角半角と大文字小文字と空白だけを揃える。記号まで落とすと
    「A」と「A'」のような別物が同じ曲になってしまうのでやらない。
    """
    s = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return "".join(s.split())


def make_keys(row) -> tuple[str, str]:
    """(管理番号キー, トラック番号＋曲名キー)。使えない方は空文字。"""
    mgmt = norm_id(row.get("元管理番号"))
    track = norm_id(row.get("トラック番号"))

    # 旧形式は盤番号だけで終わることがある（57A-0023）。そのままだと
    # 同じ盤の別の曲が全部ひとつになってしまうので、トラック番号を足す。
    if mgmt and track and not mgmt.endswith(track):
        mgmt = f"{mgmt}{track}"

    title = norm_title(row.get("曲名"))
    track_key = f"{track}|{title}" if track and title else ""
    return mgmt, track_key


# ─── 貯める ────────────────────────────────────────────

def _cell(value) -> str:
    """DataFrame の値を文字列に。欠損は空文字。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    return "" if s.lower() in ("nan", "none") else s


def collect(songs_df: pd.DataFrame, user: str = "") -> list[dict]:
    """songs_df から song_master に入れる形を作る。

    「一致」以上の行だけを対象にする。列ごとに出典を付ける。
    """
    if songs_df is None or songs_df.empty:
        return []

    out: list[dict] = []
    for _, row in songs_df.iterrows():
        status = _cell(row.get("確認ステータス"))
        src = TRUSTED_STATUS.get(status)
        if src is None:
            continue

        mgmt_key, track_key = make_keys(row)
        if not mgmt_key and not track_key:
            continue

        data: dict[str, dict] = {}
        for col in MASTER_FIELDS:
            val = _cell(row.get(col))
            if not val:
                continue
            data[col] = {"v": val, "src": src, "by": user}

        if not data:
            continue

        out.append({
            "mgmt_key": mgmt_key,
            "track_key": track_key,
            "title": _cell(row.get("曲名")),
            "data": data,
        })
    return out


def merge_data(old: dict, new: dict) -> tuple[dict, bool]:
    """既存の値に新しい値を混ぜる。強い出典が勝つ。

    返り値は (混ぜた結果, 変わったか)。
    """
    merged = dict(old or {})
    changed = False
    for col, cell in (new or {}).items():
        cur = merged.get(col)
        if cur is None:
            merged[col] = cell
            changed = True
            continue
        # 人が入れた値は動かさない
        if rank(cur.get("src")) >= PROTECTED_RANK:
            continue
        if rank(cell.get("src")) < rank(cur.get("src")):
            continue
        if cur.get("v") == cell.get("v") and cur.get("src") == cell.get("src"):
            continue
        merged[col] = cell
        changed = True
    return merged, changed


def save(songs_df: pd.DataFrame, user: str = "") -> int:
    """songs_df の確定・一致行を song_master に貯める。入れた曲数を返す。"""
    records = collect(songs_df, user)
    if not records:
        return 0

    existing = master_fetch(
        {r["mgmt_key"] for r in records if r["mgmt_key"]},
        {r["track_key"] for r in records if r["track_key"]},
    )
    by_mgmt = {e["mgmt_key"]: e for e in existing if e["mgmt_key"]}
    by_track = {e["track_key"]: e for e in existing if e["track_key"]}

    writes: list[dict] = []
    for rec in records:
        hit = by_mgmt.get(rec["mgmt_key"]) or by_track.get(rec["track_key"])
        if hit is None:
            writes.append({**rec, "id": None})
            # 同じ実行の中で同じ曲が二度出てきても、二重に作らない
            if rec["mgmt_key"]:
                by_mgmt[rec["mgmt_key"]] = {**rec, "id": None}
            if rec["track_key"]:
                by_track[rec["track_key"]] = {**rec, "id": None}
            continue

        merged, changed = merge_data(hit.get("data") or {}, rec["data"])
        if not changed:
            continue
        hit["data"] = merged
        writes.append({
            "id": hit.get("id"),
            # 片方しかキーが無かった行に、もう片方のキーを足していく
            "mgmt_key": hit.get("mgmt_key") or rec["mgmt_key"],
            "track_key": hit.get("track_key") or rec["track_key"],
            "title": hit.get("title") or rec["title"],
            "data": merged,
        })

    if not writes:
        return 0
    return master_upsert(writes)


# ─── 使う ──────────────────────────────────────────────

def mark_status(df: pd.DataFrame, idx) -> bool:
    """管理番号で台帳に当たった行に LEDGER_STATUS を書く。書いたら True。

    自社CD台帳（cd_master）からも同じ形で呼ぶ。
    """
    if "確認ステータス" not in df.columns:
        return False
    if _cell(df.at[idx, "確認ステータス"]) not in LEDGER_OVERWRITABLE:
        return False
    df.at[idx, "確認ステータス"] = LEDGER_STATUS
    return True


def fill(songs_df: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    """song_master の値で songs_df の空欄を埋める。

    返り値は (埋めた後の df, 当たった曲数, 埋めた欄の数)。

    埋めるのは空欄だけ。既に入っている値は触らない。今そこにある値が
    どこから来たのか df 側は覚えていないので、上書きすると人が直した
    値を消しかねないため。
    """
    if songs_df is None or songs_df.empty:
        return songs_df, 0, 0

    keys = [make_keys(row) for _, row in songs_df.iterrows()]
    found = master_fetch(
        {k[0] for k in keys if k[0]},
        {k[1] for k in keys if k[1]},
    )
    if not found:
        return songs_df, 0, 0

    by_mgmt = {e["mgmt_key"]: e for e in found if e["mgmt_key"]}
    by_track = {e["track_key"]: e for e in found if e["track_key"]}

    df = songs_df.copy()
    hit_rows = 0
    filled = 0
    for pos, (mgmt_key, track_key) in enumerate(keys):
        by_number = bool(mgmt_key) and mgmt_key in by_mgmt
        hit = by_mgmt.get(mgmt_key) or by_track.get(track_key)
        if hit is None:
            continue
        idx = df.index[pos]
        touched = False
        # 管理番号で当たった行は、それだけで「どの曲か」が決まる。
        # 曲名＋トラック番号で当たった方は付けない（同名の別物が
        # 混ざりうるので、人が見に行く先として残す）
        if by_number and mark_status(df, idx):
            filled += 1
            touched = True
        for col, cell in (hit.get("data") or {}).items():
            if col not in df.columns:
                continue
            if _cell(df.at[idx, col]):
                continue
            val = cell.get("v") if isinstance(cell, dict) else cell
            if not val:
                continue
            df.at[idx, col] = val
            filled += 1
            touched = True
        if touched:
            hit_rows += 1

    return df, hit_rows, filled


# ─── 管理タブから直す・見る ────────────────────────────

def cell_of(record: dict, col: str) -> dict:
    """1件の中の1列を {v, src, by} の形で取り出す。無ければ空。"""
    cell = (record.get("data") or {}).get(col)
    if isinstance(cell, dict):
        return {"v": _cell(cell.get("v")),
                "src": _cell(cell.get("src")),
                "by": _cell(cell.get("by"))}
    # 昔の形（値だけ）が入っていても読めるようにしておく
    return {"v": _cell(cell), "src": "", "by": ""}


def edit(record: dict, values: dict, user: str = "") -> int:
    """管理タブでの手直しを保存する。書いた件数を返す。

    人が直した値は出典を「手入力」にする。これで、あとから機械が
    調べ直しても上書きされない（merge_data が PROTECTED_RANK で止める）。
    空にした列はその場から消す。「値が無い」と「空文字が入っている」を
    分けても得が無いため。
    """
    if not record:
        return 0

    data = dict(record.get("data") or {})
    changed = False
    for col in MASTER_FIELDS:
        if col not in values:
            continue
        new = _cell(values.get(col))
        old = cell_of(record, col)
        if new == old["v"]:
            continue
        if new:
            data[col] = {"v": new, "src": HAND_SRC, "by": user}
        else:
            data.pop(col, None)
        changed = True

    title = _cell(values.get("曲名")) or _cell(record.get("title"))
    if title != _cell(record.get("title")):
        changed = True

    if not changed:
        return 0

    return master_upsert([{
        "id": record.get("id"),
        "mgmt_key": record.get("mgmt_key") or "",
        "track_key": record.get("track_key") or "",
        "title": title,
        "data": data,
    }])


def to_frame(records: list[dict]) -> pd.DataFrame:
    """一覧に出す表を作る。中身の列は値だけ並べる（出典は編集画面で見る）。"""
    rows = []
    for rec in records or []:
        row = {
            "id": rec.get("id"),
            "曲名": _cell(rec.get("title")),
            "管理番号キー": _cell(rec.get("mgmt_key")),
            "トラックキー": _cell(rec.get("track_key")),
            "更新": _cell(rec.get("updated_at")),
        }
        for col in MASTER_FIELDS:
            row[col] = cell_of(rec, col)["v"]
        rows.append(row)

    cols = ["id", "曲名", "管理番号キー", "トラックキー", "更新",
            *MASTER_FIELDS]
    return pd.DataFrame(rows, columns=cols)


def sources_of(record: dict) -> str:
    """「誰が・どの出典で」を1行にまとめる。一覧の説明に使う。"""
    seen: list[str] = []
    for col in MASTER_FIELDS:
        c = cell_of(record, col)
        if not c["v"] or not c["src"]:
            continue
        label = f"{c['src']}（{c['by']}）" if c["by"] else c["src"]
        if label not in seen:
            seen.append(label)
    return " / ".join(seen)
