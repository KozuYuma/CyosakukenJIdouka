"""
照合モジュール
NUENDO Cue イベントと WAV/MP3 ファイルの照合処理
"""
import pandas as pd

from .normalizer import normalize_for_match, detect_title_from_filename
from .number_parser import parse_number

# 照合ステータスの定義（優先度順）
MATCH_EXACT = "完全一致"
MATCH_NORMALIZED = "正規化一致"
MATCH_NUMBER = "管理番号一致"
MATCH_TITLE = "曲名一致"
MATCH_FILENAME = "ファイル名一致"
MATCH_PARTIAL = "部分一致"
MATCH_NONE = "WAV未照合"


def _build_lookup(df: pd.DataFrame) -> dict[str, dict]:
    """
    WAV/MP3 DataFrame から照合用ルックアップを構築する。
    キー: 正規化済みファイル名 → 値: 行データ（dict）
    """
    lookup: dict[str, dict] = {}
    for _, row in df.iterrows():
        filename = str(row.get("FileName", ""))
        if not filename or filename.lower() == "nan":
            continue
        # 完全一致用（拡張子あり）
        lookup[filename.lower()] = row.to_dict()
        # 正規化一致用（拡張子なし・記号正規化済み）
        norm = normalize_for_match(filename)
        if norm:
            lookup[norm] = row.to_dict()
    return lookup


def _match_single(
    event_name: str,
    nuendo_filename: str,
    lookup: dict[str, dict],
    audio_df: pd.DataFrame,
) -> tuple[dict | None, str]:
    """
    1 件の Cue イベントを WAV/MP3 DataFrame に照合する。
    Returns: (マッチした行データ or None, 照合ステータス文字列)
    """
    event_str = str(event_name).strip()
    file_str = str(nuendo_filename).strip()

    # --- 優先度1: イベント名とファイル名の完全一致 ---
    key_exact = event_str.lower()
    if key_exact in lookup:
        return lookup[key_exact], MATCH_EXACT

    # --- 優先度2: 正規化一致 ---
    event_norm = normalize_for_match(event_str)
    if event_norm and event_norm in lookup:
        return lookup[event_norm], MATCH_NORMALIZED

    # --- 優先度3: 管理番号一致 ---
    event_parsed = parse_number(event_str)
    event_mgmt_id = event_parsed.get("元管理番号", "")
    if event_mgmt_id:
        for _, row in audio_df.iterrows():
            wav_parsed = parse_number(str(row.get("FileName", "")))
            if wav_parsed.get("元管理番号") == event_mgmt_id:
                return row.to_dict(), MATCH_NUMBER

    # --- 優先度4: 曲名一致（管理番号除去後）---
    event_title = event_parsed.get("検出曲名", "")
    event_title_norm = normalize_for_match(event_title)
    if event_title_norm:
        for _, row in audio_df.iterrows():
            wav_title = detect_title_from_filename(str(row.get("FileName", "")))
            if event_title_norm == normalize_for_match(wav_title):
                return row.to_dict(), MATCH_TITLE

    # --- 優先度5: NUENDO ファイル名一致 ---
    file_norm = normalize_for_match(file_str)
    if file_norm and file_norm in lookup:
        return lookup[file_norm], MATCH_FILENAME

    # --- 優先度6: 部分一致 ---
    if event_norm:
        for _, row in audio_df.iterrows():
            wav_norm = normalize_for_match(str(row.get("FileName", "")))
            if wav_norm and (event_norm in wav_norm or wav_norm in event_norm):
                return row.to_dict(), MATCH_PARTIAL

    return None, MATCH_NONE


def build_song_list(
    cue_df: pd.DataFrame,
    wav_df: pd.DataFrame,
    mp3_df: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Cue 一覧 × WAV 一覧 を照合し、楽曲まとめと全イベント一覧を返す。
    WAV で照合できなかったものは MP3 一覧でも試みる（mp3_df が渡された場合）。

    Returns:
        songs_df  : 楽曲まとめ（イベント名単位でユニーク）
        events_df : イベント一覧（NUENDO イベント単位・重複あり）
    """
    wav_lookup = _build_lookup(wav_df) if wav_df is not None and len(wav_df) > 0 else {}
    mp3_lookup = _build_lookup(mp3_df) if mp3_df is not None and len(mp3_df) > 0 else {}

    events: list[dict] = []
    # イベント名をキーにして楽曲をまとめる（同じ曲が複数回使われる場合に対応）
    songs_dict: dict[str, dict] = {}

    for _, row in cue_df.iterrows():
        event_name = str(row.get("イベント名", "")).strip()
        nuendo_filename = str(row.get("ファイル名", "")).strip()
        track = str(row.get("トラック名", "")).strip()
        start_time = str(row.get("START TIME", "")).strip()
        end_time = str(row.get("終了時間", "")).strip()
        in_time = str(row.get("イン時間", "")).strip()
        out_time = str(row.get("アウト時間", "")).strip()
        duration = str(row.get("長さ", "")).strip()

        # --- WAV 照合 ---
        matched_wav, wav_status = _match_single(
            event_name, nuendo_filename, wav_lookup, wav_df if wav_df is not None else pd.DataFrame()
        )

        # --- WAV で照合できなかった場合は MP3 を補助として試みる ---
        matched_mp3 = None
        mp3_status = ""
        if wav_status == MATCH_NONE and mp3_lookup:
            matched_mp3, mp3_status = _match_single(
                event_name,
                nuendo_filename,
                mp3_lookup,
                mp3_df if mp3_df is not None else pd.DataFrame(),
            )

        # --- 管理番号分解（イベント名から） ---
        parsed = parse_number(event_name)

        # --- WAV 情報抽出 ---
        wav_filename = matched_wav.get("FileName", "") if matched_wav else ""
        wav_duration = matched_wav.get("Duration", "") if matched_wav else ""
        wav_title = (
            detect_title_from_filename(wav_filename)
            if wav_filename
            else parsed.get("検出曲名", "")
        )

        # --- MP3 情報抽出（補助） ---
        mp3_filename = matched_mp3.get("FileName", "") if matched_mp3 else ""
        mp3_duration = matched_mp3.get("Duration", "") if matched_mp3 else ""

        # --- 確認ステータス初期値 ---
        if wav_status == MATCH_NONE and not matched_mp3:
            confirm_status = "MP3補助確認"
        else:
            confirm_status = "未調査"

        # --- イベント一覧に追加 ---
        events.append(
            {
                "トラック": track,
                "イベント名": event_name,
                "ファイル名": nuendo_filename,
                "START TIME": start_time,
                "終了時間": end_time,
                "イン時間": in_time,
                "アウト時間": out_time,
                "使用尺": duration,
                "照合ステータス": wav_status,
            }
        )

        # --- 楽曲まとめ（イベント名でユニーク化） ---
        if event_name not in songs_dict:
            songs_dict[event_name] = {
                "管理番号種別": parsed.get("管理番号種別", ""),
                "元管理番号": parsed.get("元管理番号", ""),
                "ライブラリ盤番号": parsed.get("ライブラリ盤番号", ""),
                "トラック番号": parsed.get("トラック番号", ""),
                "曲名": parsed.get("検出曲名", ""),
                "イベント名": event_name,
                "NUENDOファイル名": nuendo_filename,
                "WAV一致ファイル名": wav_filename,
                "WAV検出タイトル": wav_title,
                "WAVフル尺": wav_duration,
                "WAV照合ステータス": wav_status,
                "MP3一致ファイル名": mp3_filename,
                "MP3フル尺": mp3_duration,
                "作曲者": "",
                "作詞者": "",
                "アーティスト": "",
                "CD番号": "",
                "JASRAC作品コード": "",
                "NexTone管理番号": "",
                "確認ステータス": confirm_status,
                "メモ": "",
            }

    songs_df = pd.DataFrame(list(songs_dict.values()))
    songs_df.insert(0, "No", range(1, len(songs_df) + 1))

    events_df = pd.DataFrame(events)

    return songs_df.reset_index(drop=True), events_df.reset_index(drop=True)
