"""
MusicBrainz 録音検索モジュール
曲名と再生時間で Recording を検索し、アーティスト・ISRC 等を返す。

MusicBrainz API (v2):
  https://musicbrainz.org/ws/2/recording?query=...&fmt=json
  Lucene フィールド: recording: / artist: / dur: (ms)
  レート制限: 1 req/sec (認証なし)
"""
import time
import urllib.parse

import requests

_BASE = "https://musicbrainz.org/ws/2"
_HEADERS = {
    "User-Agent": "CyosakukenJIdouka_app/1.0 (copyright-research-tool; contact: local)"
}
_RATE_LIMIT_SEC = 1.1

_last_call: float = 0.0


def _wait():
    global _last_call
    elapsed = time.time() - _last_call
    if elapsed < _RATE_LIMIT_SEC:
        time.sleep(_RATE_LIMIT_SEC - elapsed)
    _last_call = time.time()


def _hms_to_sec(hms: str) -> float:
    """
    "HH:MM:SS.mmm" または "MM:SS.mmm" → float 秒に変換。
    パース失敗時は 0.0 を返す。
    """
    s = str(hms).strip()
    if not s or s.lower() == "nan":
        return 0.0
    try:
        parts = s.split(":")
        if len(parts) == 3:
            h, m, sec = int(parts[0]), int(parts[1]), float(parts[2])
            return h * 3600 + m * 60 + sec
        if len(parts) == 2:
            m, sec = int(parts[0]), float(parts[1])
            return m * 60 + sec
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def search_recording(
    title: str,
    duration_sec: float | None = None,
    tolerance_sec: float = 15.0,
    limit: int = 10,
) -> list[dict]:
    """
    MusicBrainz で録音を検索する。

    Args:
        title:         曲名（管理番号除去済みを推奨）
        duration_sec:  再生時間（秒）。None の場合は尺による絞り込みなし
        tolerance_sec: 尺の許容誤差（秒）。デフォルト ±15 秒
        limit:         最大取得件数

    Returns:
        list of dict with keys:
          score, mb_id, title, artist, album, duration_sec, isrc
        エラー時は [{"error": "メッセージ"}]
    """
    if not title or str(title).strip() == "" or str(title).lower() == "nan":
        return []

    _wait()

    query_parts = [f'recording:"{title}"']
    if duration_sec is not None and duration_sec > 0:
        lo = max(0, int((duration_sec - tolerance_sec) * 1000))
        hi = int((duration_sec + tolerance_sec) * 1000)
        query_parts.append(f"dur:[{lo} TO {hi}]")

    query = " AND ".join(query_parts)

    try:
        resp = requests.get(
            f"{_BASE}/recording",
            params={"query": query, "fmt": "json", "limit": limit},
            headers=_HEADERS,
            timeout=10,
        )
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        return [{"error": str(e)}]

    results = []
    for rec in data.get("recordings", []):
        artist_credit = rec.get("artist-credit", [])
        artist = " / ".join(
            ac.get("artist", {}).get("name", "")
            for ac in artist_credit
            if isinstance(ac, dict)
        )

        releases = rec.get("releases", [])
        album = releases[0].get("title", "") if releases else ""

        length_ms = rec.get("length")
        dur = length_ms / 1000.0 if length_ms else 0.0

        isrcs = rec.get("isrcs", [])
        isrc = ", ".join(isrcs) if isrcs else ""

        mb_url = f"https://musicbrainz.org/recording/{rec.get('id', '')}"

        results.append({
            "score":        int(rec.get("score", 0)),
            "mb_id":        rec.get("id", ""),
            "title":        rec.get("title", ""),
            "artist":       artist,
            "album":        album,
            "duration_sec": dur,
            "isrc":         isrc,
            "mb_url":       mb_url,
        })

    return results


def mb_search_url(title: str) -> str:
    """MusicBrainz の手動検索 URL を返す"""
    q = urllib.parse.quote(title)
    return f"https://musicbrainz.org/search?query={q}&type=recording&method=advanced"
