"""
フォルダスキャンモジュール
指定フォルダ内の WAV / MP3 ファイルを再帰的に取得して DataFrame を返す。
wave モジュール（標準ライブラリ）で WAV の再生時間を取得する。
MP3 は mutagen があれば再生時間と ID3 タグ（曲名・アーティスト・
アルバム・作曲者）も読む。無ければ従来どおり空欄のまま返す。
"""
import wave
from pathlib import Path

import pandas as pd

try:  # mutagen は任意。入っていない環境でもスキャン自体は動かす
    from mutagen.mp3 import MP3 as _MutagenMP3
except Exception:  # pragma: no cover - 環境依存
    _MutagenMP3 = None

#: ID3 のフレーム名 → 出力する列名。app.py の ID3 取り込みと同じ名前にする
_ID3_TAGS = {
    "TIT2": "タイトル(ID3)",
    "TPE1": "アーティスト(ID3)",
    "TALB": "アルバム(ID3)",
    "TCOM": "作曲者(ID3)",
}


def _wav_duration(path: Path) -> str:
    """WAV ファイルの再生時間を HH:MM:SS.mmm 形式で返す"""
    try:
        with wave.open(str(path), "rb") as w:
            secs = w.getnframes() / w.getframerate()
        h = int(secs // 3600)
        m = int((secs % 3600) // 60)
        s = secs % 60
        return f"{h:02d}:{m:02d}:{s:06.3f}"
    except Exception:
        return ""


def _mp3_info(path: Path) -> dict:
    """MP3 の再生時間と ID3 タグを読む。読めないときは空欄で返す。

    再生時間は WAV と同じ HH:MM:SS.mmm 形式にそろえる。こうしておくと
    MusicBrainz の尺絞り込みにそのまま渡せる。
    """
    info = {"Duration": "", **{c: "" for c in _ID3_TAGS.values()}}
    if _MutagenMP3 is None:
        return info
    try:
        audio = _MutagenMP3(str(path))
    except Exception:
        return info
    secs = getattr(getattr(audio, "info", None), "length", 0) or 0
    if secs:
        h = int(secs // 3600)
        m = int((secs % 3600) // 60)
        info["Duration"] = f"{h:02d}:{m:02d}:{secs % 60:06.3f}"
    tags = getattr(audio, "tags", None)
    for frame, col in _ID3_TAGS.items():
        try:
            val = tags.get(frame) if tags is not None else None
        except Exception:
            val = None
        if val is not None:
            info[col] = str(val).strip()
    return info


def scan_wav_folder(folder: str, recursive: bool = True) -> pd.DataFrame:
    """
    指定フォルダの WAV ファイルを一覧化する。

    Args:
        folder: スキャンするフォルダのパス
        recursive: サブフォルダも含めるか（デフォルト True）

    Returns:
        FileName / FilePath / FileSize / Duration の DataFrame
    """
    p = Path(folder)
    if not p.exists():
        raise FileNotFoundError(f"フォルダが見つかりません: {folder}")
    if not p.is_dir():
        raise NotADirectoryError(f"フォルダではありません: {folder}")

    pattern = "**/*" if recursive else "*"
    files = sorted(
        f for f in p.glob(pattern)
        if f.is_file() and f.suffix.lower() == ".wav"
    )

    rows = []
    for f in files:
        rows.append({
            "FileName": f.stem,
            "FilePath": str(f),
            "FileSize": f.stat().st_size,
            "Duration": _wav_duration(f),
        })

    if not rows:
        return pd.DataFrame(columns=["FileName", "FilePath", "FileSize", "Duration"])
    return pd.DataFrame(rows)


def scan_mp3_folder(folder: str, recursive: bool = True) -> pd.DataFrame:
    """
    指定フォルダの MP3 ファイルを一覧化する。
    mutagen が入っていれば再生時間と ID3 タグも読む。

    Args:
        folder: スキャンするフォルダのパス
        recursive: サブフォルダも含めるか（デフォルト True）

    Returns:
        FileName / FilePath / FileSize / Duration の DataFrame
    """
    p = Path(folder)
    if not p.exists():
        raise FileNotFoundError(f"フォルダが見つかりません: {folder}")
    if not p.is_dir():
        raise NotADirectoryError(f"フォルダではありません: {folder}")

    pattern = "**/*" if recursive else "*"
    files = sorted(
        f for f in p.glob(pattern)
        if f.is_file() and f.suffix.lower() == ".mp3"
    )

    rows = []
    for f in files:
        rows.append({
            "FileName": f.stem,
            "FilePath": str(f),
            "FileSize": f.stat().st_size,
            **_mp3_info(f),
        })

    if not rows:
        return pd.DataFrame(
            columns=["FileName", "FilePath", "FileSize", "Duration",
                     *_ID3_TAGS.values()])
    return pd.DataFrame(rows)
