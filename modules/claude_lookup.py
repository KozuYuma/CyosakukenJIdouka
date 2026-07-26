"""
Claude API を使った楽曲情報ルックアップ

WAV検出タイトルをもとに Claude に問い合わせ、
CD名・アーティスト・作曲者・作詞者などを取得する。

依存: anthropic>=0.30.0
API キー: 環境変数 ANTHROPIC_API_KEY または .env ファイルの ANTHROPIC_API_KEY
"""
from __future__ import annotations

import json
import os
import re
import time

_client = None
_last_call: float = 0.0
_MIN_INTERVAL = 0.5  # 秒（過剰リクエスト防止）

_SYSTEM_PROMPT = """あなたは音楽著作権の調査アシスタントです。
ユーザーから楽曲タイトルを受け取り、その楽曲に関する情報を返してください。
必ず以下のJSONフォーマットのみで返答してください（コードブロックなし）:

{
  "official_title": "正式な楽曲タイトル（不明なら入力タイトルそのまま）",
  "artist": "アーティスト名（不明なら空文字）",
  "cd_name": "CD/アルバム名（不明なら空文字）",
  "composer": "作曲者名（不明なら空文字）",
  "lyricist": "作詞者名（不明なら空文字）",
  "isrc": "ISRC（不明なら空文字）",
  "notes": "補足・注意事項（なければ空文字）",
  "confidence": "high/medium/low（情報の確信度）"
}

- 楽曲が複数バージョンある場合は最もよく知られたものを返す
- 日本語と英語どちらでも正確な情報を優先する
- 知らない場合は推測せず空文字にすること
- CDやアルバム情報はできるだけ記入すること"""


def _get_client():
    global _client
    if _client is not None:
        return _client

    try:
        import anthropic
    except ImportError:
        return None

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        try:
            from pathlib import Path
            env_path = Path(__file__).parent.parent / ".env"
            if env_path.exists():
                for line in env_path.read_text(encoding="utf-8").splitlines():
                    if line.startswith("ANTHROPIC_API_KEY="):
                        api_key = line.split("=", 1)[1].strip().strip('"').strip("'")
                        break
        except Exception:
            pass

    if not api_key:
        return None

    _client = anthropic.Anthropic(api_key=api_key)
    return _client


def lookup_music_info(title: str) -> dict:
    """
    WAV検出タイトルから楽曲情報を取得する。

    Returns: {
        "official_title": str,
        "artist": str,
        "cd_name": str,
        "composer": str,
        "lyricist": str,
        "isrc": str,
        "notes": str,
        "confidence": "high"|"medium"|"low",
        "error": str  # エラー時のみ
    }
    """
    global _last_call

    _title = str(title).strip()
    if not _title or _title.lower() == "nan":
        return {"error": "タイトルが空です"}

    client = _get_client()
    if client is None:
        return {"error": "ANTHROPIC_API_KEY が設定されていないか anthropic パッケージが未インストールです"}

    # レート制限
    elapsed = time.monotonic() - _last_call
    if elapsed < _MIN_INTERVAL:
        time.sleep(_MIN_INTERVAL - elapsed)
    _last_call = time.monotonic()

    try:
        message = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=512,
            system=_SYSTEM_PROMPT,
            messages=[
                {"role": "user", "content": f"楽曲タイトル: {_title}"}
            ],
        )
        raw = message.content[0].text.strip()

        # JSONを抽出（コードブロックが混入した場合にも対応）
        json_match = re.search(r"\{[\s\S]*\}", raw)
        if not json_match:
            return {"error": f"JSONパース失敗: {raw[:200]}"}

        data = json.loads(json_match.group())
        data.setdefault("official_title", _title)
        data.setdefault("artist", "")
        data.setdefault("cd_name", "")
        data.setdefault("composer", "")
        data.setdefault("lyricist", "")
        data.setdefault("isrc", "")
        data.setdefault("notes", "")
        data.setdefault("confidence", "low")
        return data

    except Exception as e:
        return {"error": str(e)}


def is_available() -> bool:
    """Claude API が使用可能かどうか確認する。"""
    return _get_client() is not None
