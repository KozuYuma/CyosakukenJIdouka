"""
フォルダスキャンモジュール
指定フォルダ内の WAV / MP3 ファイルを再帰的に取得して DataFrame を返す。
wave モジュール（標準ライブラリ）で WAV の再生時間を取得する。
"""
import wave
from pathlib import Path

import pandas as pd


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
    Duration は空（mutagen 未導入のため）。

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
            "Duration": "",
        })

    if not rows:
        return pd.DataFrame(columns=["FileName", "FilePath", "FileSize", "Duration"])
    return pd.DataFrame(rows)
