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


def norm_track(value) -> str:
    """トラック番号を比べる形にする。数字なら2桁にそろえる。

    同じ曲が「1」と「01」で書かれていることがある。桁をそろえないと
    別の曲になってしまう。数字でないもの（"A" など）はそのまま。
    """
    t = norm_id(value)
    return t.zfill(2) if t.isdigit() and len(t) < 2 else t


# 管理番号キーの中で「ここから先はトラック番号」を表す区切り。
# norm_id は英数字だけを残すので、ハイフンが元の番号から出てくることは
# ない。つまりこの区切りは盤番号とトラック番号の境目としか読めない
MGMT_SEP = "-"


def mgmt_candidates(row) -> list[str]:
    """1行から、管理番号キーの候補を作る。先頭が新しく作るときの形。

    キーにしたいのは「盤番号＋トラック番号」。曲ごとに固有の番号なので、
    これが当たれば曲が決まる。

    盤番号は、あればそれ（ライブラリ盤番号）を使う。番号を読み取った行
    には必ず入っており、元管理番号がどう書かれていても同じキーになる。

    盤番号が無いのは、人が手で番号を入れた行。元管理番号がトラック番号
    まで入った形（1AN-001-01）なのか、盤番号だけ（57A-0023）なのかを
    見分ける手立てが無い。以前は「末尾がトラック番号と同じなら入って
    いる」と決めていたが、57A-0023 のトラック23 のように末尾がたまたま
    一致すると、盤番号だけのキーになって同じ盤の他の曲と混ざってしまう。
    そこで、新しく作るときは区切り（-）を入れた形にする。norm_id は
    英数字しか残さないので、この区切りは境目としか読めない。
    昔の形も候補には残すので、既にある行は今までどおり見つかる。
    """
    mgmt = norm_id(row.get("元管理番号"))
    disc = norm_id(row.get("ライブラリ盤番号"))
    if not mgmt and not disc:
        return []

    track = norm_id(row.get("トラック番号"))
    t2 = norm_track(track)
    if not t2:
        return [mgmt or disc]

    if disc:
        # 盤番号が分かっているので迷いは無い。今まで貯めてきたのと
        # 同じ形（盤番号＋2桁トラック）をそのまま使う
        out = [f"{disc}{t2}", f"{disc}{track}"]
        # 元管理番号がそれと違う形で書かれていたときのための候補。
        # 盤番号そのもの（トラックの付かない形）は入れない。それを
        # 入れると、同じ盤の曲が全部ひとつになってしまう
        if mgmt and mgmt != disc:
            out.append(mgmt)
    else:
        out = [f"{mgmt}{MGMT_SEP}{t2}", f"{mgmt}{t2}", f"{mgmt}{track}"]
        if mgmt.endswith(t2) or (track and mgmt.endswith(track)):
            # 元管理番号にトラック番号まで入っていた場合の昔のキー。
            # 盤番号だけの行とぶつかることはあるが、それは以前から
            # 同じで、ここを外すと昔の行が見つからなくなる
            out.append(mgmt)

    seen: list[str] = []
    for cand in out:
        if cand and cand not in seen:
            seen.append(cand)
    return seen


def track_candidates(row) -> list[str]:
    """トラック番号＋曲名キーの候補。先頭が新しく作るときの形。

    桁をそろえる前（"1|曲名"）で貯めた行が既にあるので、そちらも候補に
    残す。貯め直さずに、当たった方へ合流させる。
    """
    title = norm_title(row.get("曲名"))
    if not title:
        return []
    track = norm_id(row.get("トラック番号"))
    t2 = norm_track(track)
    if not t2:
        return []
    out = [f"{t2}|{title}"]
    if track and track != t2:
        out.append(f"{track}|{title}")
    return out


def make_keys(row) -> tuple[str, str]:
    """(管理番号キー, トラック番号＋曲名キー)。使えない方は空文字。

    どちらも新しく作るときの形。既にある行を探すときは
    mgmt_candidates / track_candidates の候補を全部使う。
    """
    mgmt = mgmt_candidates(row)
    track = track_candidates(row)
    return (mgmt[0] if mgmt else ""), (track[0] if track else "")


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

        mgmt_cands = mgmt_candidates(row)
        track_cands = track_candidates(row)
        if not mgmt_cands and not track_cands:
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
            # 新しく作るときの形
            "mgmt_key": mgmt_cands[0] if mgmt_cands else "",
            "track_key": track_cands[0] if track_cands else "",
            # 既にある行を探すときの候補（書き方の揺れを吸収する）
            "mgmt_cands": mgmt_cands,
            "track_cands": track_cands,
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


def _first_hit(index: dict, keys) -> dict | None:
    """候補を先頭から見て、最初に当たった行を返す。"""
    for k in keys or []:
        got = index.get(k)
        if got is not None:
            return got
    return None


def save(songs_df: pd.DataFrame, user: str = "") -> int:
    """songs_df の確定・一致行を song_master に貯める。入れた曲数を返す。"""
    records = collect(songs_df, user)
    if not records:
        return 0

    existing = master_fetch(
        {k for r in records for k in r["mgmt_cands"]},
        {k for r in records for k in r["track_cands"]},
    )
    by_mgmt = {e["mgmt_key"]: e for e in existing if e["mgmt_key"]}
    by_track = {e["track_key"]: e for e in existing if e["track_key"]}

    writes: list[dict] = []
    for rec in records:
        hit = (_first_hit(by_mgmt, rec["mgmt_cands"])
               or _first_hit(by_track, rec["track_cands"]))
        if hit is None:
            # 同じ実行の中で同じ曲が二度出てきても、二重に作らない。
            # 書き方が違うだけの行も拾えるよう、候補を全部登録する。
            # 登録するのは writes に積んだのと同じ dict。あとで中身を
            # 足しても、積んである方に反映される
            new = {
                "id": None,
                "mgmt_key": rec["mgmt_key"],
                "track_key": rec["track_key"],
                "title": rec["title"],
                "data": dict(rec["data"]),
            }
            writes.append(new)
            for k in rec["mgmt_cands"]:
                by_mgmt.setdefault(k, new)
            for k in rec["track_cands"]:
                by_track.setdefault(k, new)
            continue

        merged, changed = merge_data(hit.get("data") or {}, rec["data"])
        if not changed:
            continue
        hit["data"] = merged
        # 片方しかキーが無かった行に、もう片方のキーを足していく
        hit["mgmt_key"] = hit.get("mgmt_key") or rec["mgmt_key"]
        hit["track_key"] = hit.get("track_key") or rec["track_key"]
        hit["title"] = hit.get("title") or rec["title"]
        # この実行で作ったばかりの行なら、既に積んである
        if any(hit is w for w in writes):
            continue
        writes.append({
            "id": hit.get("id"),
            "mgmt_key": hit["mgmt_key"],
            "track_key": hit["track_key"],
            "title": hit["title"],
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

    keys = [(mgmt_candidates(row), track_candidates(row))
            for _, row in songs_df.iterrows()]
    found = master_fetch(
        {k for cands, _ in keys for k in cands},
        {k for _, cands in keys for k in cands},
    )
    if not found:
        return songs_df, 0, 0

    by_mgmt = {e["mgmt_key"]: e for e in found if e["mgmt_key"]}
    by_track = {e["track_key"]: e for e in found if e["track_key"]}

    df = songs_df.copy()
    hit_rows = 0
    filled = 0
    for pos, (mgmt_cands, track_cands) in enumerate(keys):
        hit = _first_hit(by_mgmt, mgmt_cands)
        by_number = hit is not None
        hit = hit or _first_hit(by_track, track_cands)
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
