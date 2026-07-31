"""
著作権調査パイプライン

検索語優先順位:
  1. song_title（楽曲まとめの「曲名」列 — 最優先）
  2. wav_detected_title（WAVファイル名から管理番号・トラック番号を除去したもの）
  3. イベント名から抽出した曲名（管理番号除去後）

実行順序:
  1. 検索語決定（曲名 > WAV検出タイトル > イベント名から抽出）
  2. Claude API で CD名・アーティスト・作曲者情報を取得（オプション）
  3. MusicBrainz で曲名 × 尺 → ISRC・正式タイトル・アーティスト取得
  4. Spotify で曲名 × 尺 → ISRC・アルバム・アーティスト取得（オプション）
  5. 正式タイトル（または検索語）で J-WID を検索 → JASRAC コード取得（composer で絞り込み）
  6. NexTone も同じ検索語・composer で検索
"""
from __future__ import annotations

from modules.musicbrainz import _hms_to_sec, search_recording
from modules.number_parser import parse_number
from modules.scraper import search_all, search_jwid


def run_pipeline(
    event_name: str,
    wav_full_duration: str = "",
    wav_detected_title: str = "",
    song_title: str = "",
    composer: str = "",
    jwid_artist: str = "",
    tolerance_sec: float = 15.0,
    mb_score_threshold: int = 80,
    mb_limit: int = 5,
    use_claude: bool = False,
    use_spotify: bool = False,
    sp_limit: int = 5,
) -> dict:
    """
    1件のイベントに対してパイプライン全体を実行する。

    Args:
        event_name:          NUENDO イベント名（管理番号 + 曲名）
        wav_full_duration:   WAV フル尺（"HH:MM:SS.mmm" 形式）。空の場合は尺絞りなし
        wav_detected_title:  WAV検出タイトル（WAVファイル名から抽出済み）
        song_title:          楽曲まとめの「曲名」列。最優先の検索語として使用する
        composer:            作曲者ヒント（songs_df["作曲者"] 等）。
                             J-WID / NexTone 結果のランキングと絞り込みに使用
        tolerance_sec:       尺絞り込み許容誤差（秒）— MusicBrainz / Spotify 共通
        mb_score_threshold:  MusicBrainz スコアがこれ以上なら正式タイトルで J-WID 検索
        mb_limit:            MusicBrainz 取得件数上限
        use_claude:          True のとき Claude API でCD名等を取得する（要 ANTHROPIC_API_KEY）
        use_spotify:         True のとき Spotify API でも検索する（要 SPOTIFY_CLIENT_ID/SECRET）
        sp_limit:            Spotify 取得件数上限

    Returns: {
        "event_name":        str,
        "search_title":      str,         # 実際に使用した検索語
        "wav_duration_sec":  float,
        "claude_result":     dict | None,
        "mb_results":        list[dict],
        "mb_best":           dict | None,
        "sp_results":        list[dict],
        "sp_best":           dict | None,
        "jwid_search_term":  str,
        "jwid_results":      dict,        # "composer_matched_count" キーを含む
        "nextone_results":   dict,
    }
    """
    # ── Step 1: 検索語決定（曲名 > WAV検出タイトル > イベント名から抽出）─────
    _song = str(song_title).strip()
    _song = "" if _song.lower() == "nan" else _song

    _wav = str(wav_detected_title).strip()
    _wav = "" if _wav.lower() == "nan" else _wav

    _comp = str(composer).strip()
    _comp = "" if _comp.lower() == "nan" else _comp

    if _song:
        search_title = _song          # 曲名（最優先）
    elif _wav:
        search_title = _wav           # WAV検出タイトル
    else:
        parsed = parse_number(str(event_name).strip())
        search_title = parsed.get("検出曲名", "") or str(event_name).strip()

    wav_sec = _hms_to_sec(wav_full_duration)

    # ── Step 2: Claude API ルックアップ（オプション）──────────────────
    claude_result: dict | None = None
    if use_claude and search_title:
        try:
            from modules.claude_lookup import lookup_music_info
            claude_result = lookup_music_info(search_title)
        except Exception as e:
            claude_result = {"error": str(e)}

    # ── Step 3: MusicBrainz ──────────────────────────────────────────
    mb_results = search_recording(
        title=search_title,
        duration_sec=wav_sec if wav_sec > 0 else None,
        tolerance_sec=tolerance_sec,
        limit=mb_limit,
    )

    mb_best: dict | None = None
    jwid_search_term = search_title

    valid_mb = [r for r in mb_results if "error" not in r]
    if valid_mb:
        mb_best = valid_mb[0]
        if mb_best.get("score", 0) >= mb_score_threshold:
            jwid_search_term = mb_best.get("title") or search_title

    # ── Step 4: Spotify（オプション）─────────────────────────────────
    sp_results: list[dict] = []
    sp_best: dict | None = None
    if use_spotify and search_title:
        try:
            from modules.spotify import search_track
            sp_results = search_track(
                title=search_title,
                duration_sec=wav_sec if wav_sec > 0 else None,
                tolerance_sec=tolerance_sec,
                limit=sp_limit,
            )
            valid_sp = [r for r in sp_results if "error" not in r]
            if valid_sp:
                sp_best = valid_sp[0]
                if sp_best and (mb_best is None or mb_best.get("score", 0) < mb_score_threshold):
                    if sp_best.get("score", 0) >= 70:
                        jwid_search_term = sp_best.get("title") or jwid_search_term
        except Exception as e:
            sp_results = [{"error": str(e)}]

    _jwid_artist = str(jwid_artist).strip()
    if _jwid_artist.lower() == "nan":
        _jwid_artist = ""

    # ── Step 5 & 6: J-WID / NexTone（composer で絞り込み）────────────
    search_results = search_all(jwid_search_term, composer=_comp, artist=_jwid_artist)

    return {
        "event_name":       event_name,
        "search_title":     search_title,
        "wav_duration_sec": wav_sec,
        "claude_result":    claude_result,
        "mb_results":       mb_results,
        "mb_best":          mb_best,
        "sp_results":       sp_results,
        "sp_best":          sp_best,
        "jwid_search_term": jwid_search_term,
        "jwid_results":     search_results["jwid"],
        "nextone_results":  search_results["nextone"],
    }
