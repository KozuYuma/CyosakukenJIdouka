"""
Spotify Web API による楽曲検索

Client Credentials フロー（ユーザー認証不要）を使用。
.env に SPOTIFY_CLIENT_ID / SPOTIFY_CLIENT_SECRET を設定する。

返却フォーマット（search_track の各要素）:
  {
    "title":       str,   # トラック名
    "artist":      str,   # アーティスト名（複数の場合はカンマ区切り）
    "album":       str,   # アルバム名
    "duration_sec": float,
    "isrc":        str,
    "release_date": str,
    "spotify_url": str,
    "score":       int,   # 尺一致スコア（100=完全一致、尺不明なら80固定）
  }
  エラー時は {"error": str}
"""
from __future__ import annotations

import os
import time
from pathlib import Path

import requests

_TOKEN_CACHE: dict = {}   # {"access_token": str, "expires_at": float}
_RATE_LIMIT_SEC = 0.1
_last_call: float = 0.0


def _load_env() -> tuple[str, str]:
    client_id     = os.environ.get("SPOTIFY_CLIENT_ID", "")
    client_secret = os.environ.get("SPOTIFY_CLIENT_SECRET", "")
    if not (client_id and client_secret):
        try:
            env_path = Path(__file__).parent.parent / ".env"
            if env_path.exists():
                for line in env_path.read_text(encoding="utf-8").splitlines():
                    line = line.strip()
                    if line.startswith("SPOTIFY_CLIENT_ID="):
                        client_id = line.split("=", 1)[1].strip().strip('"').strip("'")
                    elif line.startswith("SPOTIFY_CLIENT_SECRET="):
                        client_secret = line.split("=", 1)[1].strip().strip('"').strip("'")
        except Exception:
            pass
    return client_id, client_secret


def is_available() -> bool:
    cid, csec = _load_env()
    return bool(cid and csec)


def _get_token() -> str:
    global _TOKEN_CACHE
    now = time.monotonic()
    if _TOKEN_CACHE.get("access_token") and now < _TOKEN_CACHE.get("expires_at", 0):
        return _TOKEN_CACHE["access_token"]

    client_id, client_secret = _load_env()
    if not (client_id and client_secret):
        raise RuntimeError("SPOTIFY_CLIENT_ID / SPOTIFY_CLIENT_SECRET が未設定です")

    resp = requests.post(
        "https://accounts.spotify.com/api/token",
        data={"grant_type": "client_credentials"},
        auth=(client_id, client_secret),
        timeout=10,
    )
    resp.raise_for_status()
    data = resp.json()
    _TOKEN_CACHE = {
        "access_token": data["access_token"],
        "expires_at": now + data.get("expires_in", 3600) - 60,
    }
    return _TOKEN_CACHE["access_token"]


def search_track(
    title: str,
    duration_sec: float | None = None,
    tolerance_sec: float = 15.0,
    limit: int = 10,
) -> list[dict]:
    """
    Spotify でトラックを検索する。

    Args:
        title:        検索する曲名
        duration_sec: WAV フル尺（秒）。None の場合は尺フィルタなし
        tolerance_sec: 尺の許容誤差（秒）
        limit:        最大取得件数

    Returns:
        list[dict] — 見つかったトラック情報のリスト（尺近い順）。
        エラー時は [{"error": str}]
    """
    global _last_call

    _title = str(title).strip()
    if not _title:
        return [{"error": "タイトルが空です"}]

    if not is_available():
        return [{"error": "SPOTIFY_CLIENT_ID / SPOTIFY_CLIENT_SECRET が未設定です"}]

    # レート制限
    elapsed = time.monotonic() - _last_call
    if elapsed < _RATE_LIMIT_SEC:
        time.sleep(_RATE_LIMIT_SEC - elapsed)
    _last_call = time.monotonic()

    try:
        token = _get_token()
    except Exception as e:
        return [{"error": f"トークン取得失敗: {e}"}]

    headers = {"Authorization": f"Bearer {token}"}
    params = {
        "q": f'track:"{_title}"',
        "type": "track",
        "limit": min(limit * 2, 50),   # 尺フィルタ後に limit 件残るよう多めに取得
        "market": "JP",
    }

    try:
        resp = requests.get(
            "https://api.spotify.com/v1/search",
            headers=headers,
            params=params,
            timeout=10,
        )
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        return [{"error": str(e)}]

    tracks = data.get("tracks", {}).get("items", [])
    if not tracks:
        return []

    results: list[dict] = []
    for t in tracks:
        dur_ms = t.get("duration_ms") or 0
        dur_s  = dur_ms / 1000.0

        # 尺フィルタ
        if duration_sec is not None and duration_sec > 0 and dur_s > 0:
            if abs(dur_s - duration_sec) > tolerance_sec:
                continue

        artists = ", ".join(a["name"] for a in t.get("artists", []))
        album   = t.get("album", {}).get("name", "")
        rel_date = t.get("album", {}).get("release_date", "")
        isrc    = t.get("external_ids", {}).get("isrc", "")
        url     = t.get("external_urls", {}).get("spotify", "")

        # スコア: 尺一致度を 0〜100 で計算（尺不明なら 80）
        if duration_sec and dur_s:
            diff = abs(dur_s - duration_sec)
            score = max(0, int(100 - diff / tolerance_sec * 50))
        else:
            score = 80

        results.append({
            "title":        t.get("name", ""),
            "artist":       artists,
            "album":        album,
            "duration_sec": dur_s,
            "isrc":         isrc,
            "release_date": rel_date,
            "spotify_url":  url,
            "score":        score,
        })

        if len(results) >= limit:
            break

    # スコア降順ソート
    results.sort(key=lambda r: r["score"], reverse=True)
    return results


def spotify_search_url(title: str) -> str:
    import urllib.parse
    return f"https://open.spotify.com/search/{urllib.parse.quote(title)}"
