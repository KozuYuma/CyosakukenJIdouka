"""
照合モジュール
NUENDO Cue イベントと WAV/MP3 ファイルの照合処理
"""
import pandas as pd

from .normalizer import normalize_for_match, detect_title_from_filename
from .number_parser import parse_number


def duration_to_min_sec(duration_str: str) -> tuple[int, int]:
    """
    "HH:MM:SS:FF"（タイムコード）/ "HH:MM:SS.mmm" / "MM:SS.mmm" / "SS.mmm" → (分, 秒)。
    HH:MM:SS:FF 形式はフレーム部分を切り捨て（フレームが1以上あれば1秒繰り上げ）。
    申告フォーマットの使用時間（分）・使用時間（秒）算出用。
    """
    s = str(duration_str).strip()
    if not s or s.lower() == "nan":
        return 0, 0
    try:
        parts = s.split(":")
        if len(parts) == 4:
            # HH:MM:SS:FF タイムコード形式
            hh, mm, ss, ff = int(parts[0]), int(parts[1]), int(parts[2]), int(parts[3])
            total_sec = hh * 3600 + mm * 60 + ss + (1 if ff > 0 else 0)
        elif len(parts) == 3:
            total_sec = int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
        elif len(parts) == 2:
            total_sec = int(parts[0]) * 60 + float(parts[1])
        else:
            total_sec = float(parts[0])
        total_int = int(total_sec)
        return total_int // 60, total_int % 60
    except (ValueError, TypeError):
        return 0, 0

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
    WAV CSV は "FileName" 列、mp3finder CSV は "ファイル名" 列を使用。
    """
    # FileName (WAV形式) or ファイル名 (mp3finder形式) を許容
    col = "FileName" if "FileName" in df.columns else "ファイル名"
    lookup: dict[str, dict] = {}
    for _, row in df.iterrows():
        filename = str(row.get(col, ""))
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

    # WAV CSV は "FileName"、mp3finder CSV は "ファイル名" 列を使用
    fn_col = "FileName" if (len(audio_df) > 0 and "FileName" in audio_df.columns) else "ファイル名"

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
            wav_parsed = parse_number(str(row.get(fn_col, "")))
            if wav_parsed.get("元管理番号") == event_mgmt_id:
                return row.to_dict(), MATCH_NUMBER

    # --- 優先度4: 曲名一致（管理番号除去後）---
    event_title = event_parsed.get("検出曲名", "")
    event_title_norm = normalize_for_match(event_title)
    if event_title_norm:
        for _, row in audio_df.iterrows():
            wav_title = detect_title_from_filename(str(row.get(fn_col, "")))
            if event_title_norm == normalize_for_match(wav_title):
                return row.to_dict(), MATCH_TITLE

    # --- 優先度5: NUENDO ファイル名一致 ---
    file_norm = normalize_for_match(file_str)
    if file_norm and file_norm in lookup:
        return lookup[file_norm], MATCH_FILENAME

    # --- 優先度6: 部分一致 ---
    if event_norm:
        for _, row in audio_df.iterrows():
            wav_norm = normalize_for_match(str(row.get(fn_col, "")))
            if wav_norm and (event_norm in wav_norm or wav_norm in event_norm):
                return row.to_dict(), MATCH_PARTIAL

    return None, MATCH_NONE


def build_song_list(
    cue_df: pd.DataFrame,
    wav_df: pd.DataFrame | None = None,
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

        # イベント名の検出曲名（管理番号除去済み）からさらにトラック番号プレフィックスを除去する
        # 例: "1-01 回レ!雪月花" → "回レ!雪月花"、"05 It's a Happening!" → "It's a Happening!"
        _raw_title = parsed.get("検出曲名", "")
        # イベント名は末尾マーカー除去不要（ファイル名用の ⑥⑦ をスキップ）
        _clean_title = detect_title_from_filename(_raw_title, strip_trailing=False) or _raw_title

        # --- WAV 情報抽出 ---
        wav_filename = matched_wav.get("FileName", "") if matched_wav else ""
        wav_duration = matched_wav.get("Duration", "") if matched_wav else ""
        wav_title = (
            detect_title_from_filename(wav_filename)
            if wav_filename
            else _clean_title
        )

        # --- MP3 情報抽出（補助） ---
        # WAV形式は "FileName"/"Duration"、mp3finder形式は "ファイル名"/"再生時間" を使用
        if matched_mp3:
            mp3_filename = matched_mp3.get("FileName") or matched_mp3.get("ファイル名", "")
            mp3_duration = matched_mp3.get("Duration") or matched_mp3.get("再生時間", "")
            # ID3 の曲名。無ければファイル名から拾う。検索語の控えになる
            mp3_title = str(matched_mp3.get("タイトル(ID3)", "") or "").strip()
            if not mp3_title or mp3_title.lower() == "nan":
                mp3_title = detect_title_from_filename(mp3_filename) if mp3_filename else ""
        else:
            mp3_filename = ""
            mp3_duration = ""
            mp3_title = ""

        # --- 確認ステータス初期値 ---
        if wav_status == MATCH_NONE and not matched_mp3:
            confirm_status = "MP3補助確認"
        else:
            confirm_status = "未調査"

        # --- イベント一覧に追加 ---
        _min, _sec = duration_to_min_sec(duration)
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
                "使用時間（分）": _min,
                "使用時間（秒）": _sec,
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
                "曲名": _clean_title,
                "イベント名": event_name,
                "NUENDOファイル名": nuendo_filename,
                "WAV一致ファイル名": wav_filename,
                "WAV検出タイトル": wav_title,
                "WAVフル尺": wav_duration,
                "WAV照合ステータス": wav_status,
                "MP3一致ファイル名": mp3_filename,
                "MP3フル尺": mp3_duration,
                "MP3検出タイトル": mp3_title,
                "使用形態": "背景",
                "音源区分": "CD",
                "I/V区分": "",
                "邦洋区分": "",
                "原訳詞区分": "",
                "作曲者": "",
                "作詞者": "",
                "編曲者": "",
                "訳詞者": "",
                "アーティスト": "",
                "レコード会社名": "",
                "CD番号": "",
                "CD名": "",
                "JASRAC作品コード": "",
                "NexTone管理番号": "",
                "委任者": "",
                # 管理状況から書き取る（○/△/×）。空欄はまだ引いていない。
                # 末尾 J/N は出どころ（JASRAC／NexTone）別の生の値で、
                # 表に出すのは両方を合わせた「放送」「配信」の方
                "放送": "",
                "配信": "",
                "放送J": "",
                "放送N": "",
                "配信J": "",
                "配信N": "",
                "確認ステータス": confirm_status,
                "自社楽曲ID": "",
                # 値の出どころ。自社CDの台帳から入れた行にだけ書く
                "情報元": "",
                "メモ": "",
            }

    songs_df = pd.DataFrame(list(songs_dict.values()))
    songs_df.insert(0, "No", range(1, len(songs_df) + 1))

    events_df = pd.DataFrame(events)

    return songs_df.reset_index(drop=True), events_df.reset_index(drop=True)
