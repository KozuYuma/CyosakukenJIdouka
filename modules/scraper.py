"""
J-WID / NexTone スクレイピングモジュール

【重要】
・個人の業務用途での使用を想定しています
・リクエスト間隔を設けてサーバー負荷を抑えています
・サイト改修により動作しなくなる可能性があります
・各サービスの利用規約を確認のうえ使用してください
"""

import re
import time
import urllib.parse

import requests
from bs4 import BeautifulSoup

# =====================================================================
# 設定
# =====================================================================

# リクエスト間隔（秒）- サーバーへの負荷軽減
RATE_LIMIT_SEC = 2.0

# タイムアウト（秒）
TIMEOUT = 15

# ブラウザに偽装したヘッダー
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}

# セッション（Cookie を保持するためのセッション管理）
_session = requests.Session()
_session.headers.update(HEADERS)

# ドメインごとの最終リクエスト時刻
_last_request: dict[str, float] = {}


def _rate_limit(domain: str) -> None:
    """同じドメインへの連続リクエストを抑制する"""
    elapsed = time.time() - _last_request.get(domain, 0)
    if elapsed < RATE_LIMIT_SEC:
        time.sleep(RATE_LIMIT_SEC - elapsed)
    _last_request[domain] = time.time()


def _get(url: str, params: dict | None = None, domain: str = "") -> requests.Response:
    """レート制限付き GET リクエスト"""
    _rate_limit(domain or url)
    return _session.get(url, params=params, timeout=TIMEOUT, allow_redirects=True)


def _find_val(data: dict, keys: list[str]) -> str:
    """複数のキー候補で辞書を検索し最初にヒットした値を返す"""
    for key in keys:
        for k, v in data.items():
            if key in k and v:
                return str(v).strip()
    return ""


# =====================================================================
# J-WID (JASRAC)
# =====================================================================

JWID_BASE = "https://www2.jasrac.or.jp/eJwid"
JWID_SEARCH = f"{JWID_BASE}/main"

# J-WID の検索結果テーブルで使われる列名候補
_JWID_WORK_CODE_KEYS  = ["作品コード", "WORKS_CD", "コード"]
_JWID_TITLE_KEYS      = ["作品名", "タイトル", "TITLE"]
_JWID_COMPOSER_KEYS   = ["作曲者", "作曲", "COMPOSER"]
_JWID_LYRICIST_KEYS   = ["作詞者", "作詞", "LYRICIST"]
_JWID_ARRANGER_KEYS   = ["編曲者", "編曲", "ARRANGER"]
_JWID_PUBLISHER_KEYS  = ["出版者", "出版", "PUBLISHER"]
_JWID_ARTIST_KEYS     = ["アーティスト", "歌手", "実演家"]


def search_jwid(title: str) -> dict:
    """
    J-WID（JASRAC）でタイトルを検索して結果を返す。

    Returns: {
        "source": "J-WID",
        "search_url": str,
        "results": [{"作品コード":..., "作品名":..., ...}, ...],
        "error": str | None,
        "debug_html": str,   # パース失敗時の確認用 HTML（先頭 3000 文字）
    }
    """
    params = {
        "trxID": "F10101",
        "TITLE_NAME": title,
        "TITLE_NAME_MATCH_KBN": "2",  # 部分一致
    }
    search_url = f"{JWID_SEARCH}?{urllib.parse.urlencode(params, encoding='euc-jp', errors='replace')}"

    out: dict = {
        "source": "J-WID",
        "search_url": search_url,
        "results": [],
        "error": None,
        "debug_html": "",
    }

    try:
        resp = _get(JWID_SEARCH, params=params, domain="jasrac.or.jp")
        resp.raise_for_status()

        # J-WID は EUC-JP エンコーディング
        resp.encoding = "euc-jp"
        html = resp.text
        out["debug_html"] = html[:3000]

        soup = BeautifulSoup(html, "lxml")

        # エラーメッセージを確認
        error_msgs = soup.find_all(string=re.compile(r"(該当.*なし|エラー|Error|0件)", re.I))
        if error_msgs and not _has_results_table(soup):
            out["error"] = f"検索結果なし（サイトメッセージ: {error_msgs[0].strip()[:80]}）"
            return out

        results = _parse_jwid_table(soup)
        out["results"] = results

    except requests.exceptions.ConnectionError:
        out["error"] = "接続エラー: J-WID サイトに接続できません。ネットワーク接続を確認してください。"
    except requests.exceptions.Timeout:
        out["error"] = f"タイムアウト（{TIMEOUT}秒）: J-WID の応答が遅すぎます。"
    except requests.exceptions.HTTPError as e:
        out["error"] = f"HTTP エラー: {e}"
    except Exception as e:
        out["error"] = f"予期しないエラー: {type(e).__name__}: {e}"

    return out


def _has_results_table(soup: BeautifulSoup) -> bool:
    """検索結果テーブルが存在するか確認"""
    for table in soup.find_all("table"):
        text = table.get_text()
        if "作品コード" in text or "作品名" in text:
            return True
    return False


def _parse_jwid_table(soup: BeautifulSoup) -> list[dict]:
    """J-WID 検索結果 HTML からデータ行を抽出する"""
    results = []

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue

        # ヘッダー行を取得
        header_cells = rows[0].find_all(["th", "td"])
        headers = [c.get_text(strip=True) for c in header_cells]

        # 作品コードか作品名がヘッダーに含まれるテーブルだけ処理
        if not any(k in h for k in _JWID_WORK_CODE_KEYS + _JWID_TITLE_KEYS for h in headers):
            continue

        for row in rows[1:]:
            cells = row.find_all("td")
            if not cells:
                continue
            cell_texts = [c.get_text(strip=True) for c in cells]

            # ヘッダーとセルをマッピング
            row_data = {}
            for i, h in enumerate(headers):
                if i < len(cell_texts):
                    row_data[h] = cell_texts[i]

            item = {
                "作品コード":  _find_val(row_data, _JWID_WORK_CODE_KEYS),
                "作品名":      _find_val(row_data, _JWID_TITLE_KEYS),
                "作曲者":      _find_val(row_data, _JWID_COMPOSER_KEYS),
                "作詞者":      _find_val(row_data, _JWID_LYRICIST_KEYS),
                "編曲者":      _find_val(row_data, _JWID_ARRANGER_KEYS),
                "出版者":      _find_val(row_data, _JWID_PUBLISHER_KEYS),
                "アーティスト": _find_val(row_data, _JWID_ARTIST_KEYS),
            }

            if any(v for v in item.values()):
                results.append(item)

    return results


# =====================================================================
# NexTone
# =====================================================================

NEXTONE_SEARCH_HTML = "https://search.nex-tone.co.jp/search"
# NexTone は Next.js 製。内部 API エンドポイントを複数試みる
_NEXTONE_API_CANDIDATES = [
    "https://search.nex-tone.co.jp/api/search",
    "https://search.nex-tone.co.jp/api/works",
    "https://api.nex-tone.co.jp/v1/search",
]

_NT_ID_KEYS       = ["管理番号", "id", "code", "management_id", "workId", "work_id"]
_NT_TITLE_KEYS    = ["作品名", "title", "work_title", "name", "song_title"]
_NT_COMPOSER_KEYS = ["作曲者", "composer", "music_author"]
_NT_LYRICIST_KEYS = ["作詞者", "lyricist", "words_author"]
_NT_ARTIST_KEYS   = ["アーティスト", "artist", "performer"]
_NT_ALBUM_KEYS    = ["アルバム", "album", "cd_title"]


def search_nextone(title: str) -> dict:
    """
    NexTone でタイトルを検索して結果を返す。
    JSON API → HTML の順で試みる。
    """
    search_url = f"{NEXTONE_SEARCH_HTML}?keyword={urllib.parse.quote(title)}"

    out: dict = {
        "source": "NexTone",
        "search_url": search_url,
        "results": [],
        "error": None,
        "debug_html": "",
    }

    # ---- JSON API を試みる ----
    api_results = _try_nextone_api(title)
    if api_results is not None:
        out["results"] = api_results
        return out

    # ---- HTML スクレイピング ----
    try:
        resp = _get(NEXTONE_SEARCH_HTML, params={"keyword": title}, domain="nex-tone.co.jp")
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"
        html = resp.text
        out["debug_html"] = html[:3000]

        soup = BeautifulSoup(html, "lxml")
        results = _parse_nextone_html(soup)
        out["results"] = results

        if not results:
            # 検索結果ゼロかパース失敗かを判別
            page_text = soup.get_text()
            if re.search(r"(0件|該当.*なし|not found)", page_text, re.I):
                out["error"] = "検索結果 0 件"
            else:
                out["error"] = (
                    "HTML パース失敗: サイトの構造が変わった可能性があります。"
                    " debug_html を確認して scraper.py の _parse_nextone_html を修正してください。"
                )

    except requests.exceptions.ConnectionError:
        out["error"] = "接続エラー: NexTone サイトに接続できません。"
    except requests.exceptions.Timeout:
        out["error"] = f"タイムアウト（{TIMEOUT}秒）"
    except requests.exceptions.HTTPError as e:
        out["error"] = f"HTTP エラー: {e}"
    except Exception as e:
        out["error"] = f"予期しないエラー: {type(e).__name__}: {e}"

    return out


def _try_nextone_api(title: str) -> list[dict] | None:
    """NexTone の内部 API エンドポイントを試みる"""
    for api_url in _NEXTONE_API_CANDIDATES:
        try:
            _rate_limit("nex-tone.co.jp")
            resp = _session.get(
                api_url,
                params={"keyword": title, "q": title, "page": 1, "limit": 20},
                timeout=TIMEOUT,
            )
            if resp.status_code != 200:
                continue
            ct = resp.headers.get("Content-Type", "")
            if "application/json" not in ct and "json" not in ct:
                continue

            data = resp.json()
            results = _parse_nextone_json(data)
            if results is not None:
                return results

        except Exception:
            continue

    return None


def _parse_nextone_json(data) -> list[dict] | None:
    """NexTone JSON レスポンスを整形する（構造が不明なため複数パターンを試みる）"""
    # リスト直接
    if isinstance(data, list):
        items = data
    # data.results / data.data / data.items / data.works のいずれか
    elif isinstance(data, dict):
        items = (
            data.get("results")
            or data.get("data")
            or data.get("items")
            or data.get("works")
            or []
        )
    else:
        return None

    if not isinstance(items, list):
        return None

    results = []
    for item in items:
        if not isinstance(item, dict):
            continue

        def pick(keys):
            for k in keys:
                # 完全一致
                if k in item and item[k]:
                    return str(item[k]).strip()
                # 大文字小文字無視の部分一致
                for ik, iv in item.items():
                    if k.lower() in ik.lower() and iv:
                        return str(iv).strip()
            return ""

        entry = {
            "管理番号":    pick(_NT_ID_KEYS),
            "作品名":      pick(_NT_TITLE_KEYS),
            "作曲者":      pick(_NT_COMPOSER_KEYS),
            "作詞者":      pick(_NT_LYRICIST_KEYS),
            "アーティスト": pick(_NT_ARTIST_KEYS),
            "アルバム":    pick(_NT_ALBUM_KEYS),
        }
        if any(v for v in entry.values()):
            results.append(entry)

    return results


def _parse_nextone_html(soup: BeautifulSoup) -> list[dict]:
    """NexTone HTML 検索結果を抽出する（テーブル→カードの順で試みる）"""
    results = []

    # テーブル形式
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        headers = [th.get_text(strip=True) for th in rows[0].find_all(["th", "td"])]
        if not any(k in h for k in ["作品名", "管理番号", "タイトル", "composer"] for h in headers):
            continue

        for row in rows[1:]:
            cells = [td.get_text(strip=True) for td in row.find_all("td")]
            if not cells:
                continue
            row_data = {h: cells[i] for i, h in enumerate(headers) if i < len(cells)}
            item = {
                "管理番号":    _find_val(row_data, _NT_ID_KEYS),
                "作品名":      _find_val(row_data, _NT_TITLE_KEYS),
                "作曲者":      _find_val(row_data, _NT_COMPOSER_KEYS),
                "作詞者":      _find_val(row_data, _NT_LYRICIST_KEYS),
                "アーティスト": _find_val(row_data, _NT_ARTIST_KEYS),
                "アルバム":    _find_val(row_data, _NT_ALBUM_KEYS),
            }
            if any(v for v in item.values()):
                results.append(item)

    # カード形式（div.result-item などのクラスを探す）
    if not results:
        card_selector = re.compile(r"(result|item|work|track|song|card)", re.I)
        for card in soup.find_all(["div", "li", "article"], class_=card_selector):
            # 各カード内のラベル: 値ペアを抽出
            item: dict[str, str] = {"作品名": "", "管理番号": "", "作曲者": "", "作詞者": "", "アーティスト": "", "アルバム": ""}
            for elem in card.find_all(["dt", "th", "span", "label"]):
                label = elem.get_text(strip=True)
                sibling = elem.find_next_sibling(["dd", "td", "span"])
                value = sibling.get_text(strip=True) if sibling else ""
                if not value:
                    continue
                for key, candidates in {
                    "作品名": _NT_TITLE_KEYS,
                    "管理番号": _NT_ID_KEYS,
                    "作曲者": _NT_COMPOSER_KEYS,
                    "作詞者": _NT_LYRICIST_KEYS,
                    "アーティスト": _NT_ARTIST_KEYS,
                    "アルバム": _NT_ALBUM_KEYS,
                }.items():
                    if any(c in label for c in candidates):
                        item[key] = value
            if any(v for v in item.values()):
                results.append(item)

    return results


# =====================================================================
# まとめて検索
# =====================================================================

def search_all(title: str) -> dict:
    """J-WID と NexTone を連続して検索し、両方の結果を返す"""
    return {
        "jwid": search_jwid(title),
        "nextone": search_nextone(title),
    }
