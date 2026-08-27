"""
TSP の CSV を読んで、自社CDの台帳（cd_master）の形に整える。

元になるファイルは4つ。Access から書き出したもの。

    AllCD_DATA__Tbl_All_CD_Data.csv   曲（36万件）
    TSP_CD_DATA__Tbl_a.csv            CD（2万枚）
    TSP_CD_DATA__Tbl_a_player.csv     アーティスト名の対応表
    TSP_CD_DATA__Tbl_a_kaisya.csv     レコード会社名の対応表

曲とCDは管理番号（1AN-001）で繋がる。CD側から CD番号・レコード会社名を
持ってくる。読み込むだけで、DB には書かない（書くのは scripts/import_tsp.py）。
"""
from __future__ import annotations

import csv
import sys
from pathlib import Path
from typing import Iterator

from modules.song_master import norm_id, norm_title

# Access 由来のファイルは1つの欄が長いことがあるので上限を上げておく
csv.field_size_limit(min(sys.maxsize, 2_000_000_000))

FILE_TRACKS = "AllCD_DATA__Tbl_All_CD_Data.csv"
FILE_CDS = "TSP_CD_DATA__Tbl_a.csv"
FILE_PLAYERS = "TSP_CD_DATA__Tbl_a_player.csv"
FILE_LABELS = "TSP_CD_DATA__Tbl_a_kaisya.csv"

REQUIRED = (FILE_TRACKS, FILE_CDS, FILE_PLAYERS, FILE_LABELS)


def missing_files(folder) -> list[str]:
    """足りないファイルの名前。全部あれば空。"""
    d = Path(folder)
    return [n for n in REQUIRED if not (d / n).is_file()]


def _rows(path: Path) -> Iterator[dict]:
    # Access の書き出しは UTF-8。BOM が付くことがあるので sig で読む
    with path.open(encoding="utf-8-sig", newline="") as f:
        yield from csv.DictReader(f)


def _clean(value) -> str:
    return str(value or "").strip()


def _lookup(path: Path, name_col: str) -> dict[str, str]:
    """ID → 名前 の対応表を作る。"""
    out: dict[str, str] = {}
    for r in _rows(path):
        key = _clean(r.get("ID"))
        if key:
            out[key] = _clean(r.get(name_col))
    return out


def load_cds(folder) -> dict[str, dict]:
    """管理番号（記号なし）→ CD の情報。"""
    d = Path(folder)
    players = _lookup(d / FILE_PLAYERS, "Tbl_a_player")
    labels = _lookup(d / FILE_LABELS, "Tbl_a_kaisya")

    out: dict[str, dict] = {}
    for r in _rows(d / FILE_CDS):
        key = norm_id(r.get("Tbl_a_ken"))
        if not key:
            continue
        out[key] = {
            "cd_name": _clean(r.get("Tbl_a_title")),
            "cd_no": _clean(r.get("Tbl_a_cdban")),
            "artist": players.get(_clean(r.get("Tbl_a_player_ID")), ""),
            "label": labels.get(_clean(r.get("Tbl_a_kaisya_ID")), ""),
        }
    return out


def iter_records(folder) -> Iterator[dict]:
    """台帳に入れる1曲分を順に返す。36万件あるので溜めずに流す。"""
    d = Path(folder)
    cds = load_cds(d)

    for r in _rows(d / FILE_TRACKS):
        mgmt_key = norm_id(r.get("Fld_mm_kotei"))
        if not mgmt_key:
            continue

        title = _clean(r.get("Fld_mm_title"))
        track_no = _clean(r.get("Fld_mm_tr"))
        disc_key = norm_id(r.get("Fld_mm_ban"))
        cd = cds.get(disc_key, {})

        # 曲ごとのアーティストが入っていればそれを使う。空ならCDの
        # アーティスト（V.A. など盤ぜんたいの名義）で補う
        artist = _clean(r.get("Fld_mm_player")) or cd.get("artist", "")

        # 作品コードの欄には「メドレー情報」のような但し書きも入る。
        # 数字とハイフンで出来ていなければ作品コードとして扱わない
        code = _clean(r.get("Fld_mm_code"))
        if code and not all(ch.isdigit() or ch == "-" for ch in code):
            code = ""

        tkey = norm_title(title)
        yield {
            "mgmt_key": mgmt_key,
            "disc_key": disc_key,
            # song_master の track_key と同じ作り方にしておく。
            # そうしないと管理番号が違う持ち込み盤に当たらない
            "track_key": f"{norm_id(track_no)}|{tkey}" if track_no and tkey else "",
            "track_no": track_no,
            "title": title,
            "artist": artist,
            "composer": _clean(r.get("Fld_mm_comp")),
            "cd_name": cd.get("cd_name", "") or _clean(r.get("Fld_mm_atitle")),
            "cd_no": cd.get("cd_no", ""),
            "label": cd.get("label", ""),
            "jasrac": code,
        }
