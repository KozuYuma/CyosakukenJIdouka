"""
MusicForest (www.minc.or.jp) 検索クライアント。

【ログイン手順】
  MusicForest は reCAPTCHA があるため requests だけでは自動ログインできません。
  初回（およびセッション切れ時）は以下の手順で Cookie を保存してください:

    cd H:/PROGRAM/search_music
    .venv/Scripts/python.exe src/login_browser.py

  -> H:/PROGRAM/search_music/auth/state.json に Cookie が保存されます（有効期限 約3時間）。

【環境変数】
  MINC_STATE_PATH  Cookie ファイルのパス（省略時は上記デフォルト）
"""

from __future__ import annotations

import json
import os
import re
import time
import urllib.parse
from pathlib import Path
from typing import Optional

import requests
from bs4 import BeautifulSoup

BASE_URL    = "https://www.minc.or.jp"
SEARCH_URL  = f"{BASE_URL}/music/list"
DETAIL_URL  = f"{BASE_URL}/saku/detail/"
_CHECK_URL  = f"{BASE_URL}/search"

_DEFAULT_STATE_PATH = Path(r"H:\PROGRAM\search_music\auth\state.json")

_RATE_LIMIT = 1.5
_TIMEOUT    = 30

_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


class MusicForestError(RuntimeError):
    """認証エラー・通信エラーなど"""


# =====================================================================
# クライアント
# =====================================================================

class MusicForestClient:
    def __init__(self, delay: float = _RATE_LIMIT):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": _USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        })
        self.delay = delay
        self._last_ts: float = 0.0

    # ---- 低レベル -------------------------------------------------------

    def _throttle(self) -> None:
        wait = self.delay - (time.monotonic() - self._last_ts)
        if wait > 0:
            time.sleep(wait)
        self._last_ts = time.monotonic()

    def _get(self, url: str, **kwargs) -> requests.Response:
        self._throttle()
        resp = self.session.get(url, timeout=_TIMEOUT, allow_redirects=True, **kwargs)
        resp.raise_for_status()
        return resp

    # ---- Cookie ロード --------------------------------------------------

    def load_state(self, state_path: Path) -> None:
        """Playwright が保存した state.json の Cookie をセッションに適用する。"""
        if not state_path.exists():
            raise MusicForestError(
                f"Cookie ファイルが見つかりません: {state_path}\n"
                "login_browser.py でブラウザログインを実行してください。"
            )
        data = json.loads(state_path.read_text(encoding="utf-8"))
        for c in data.get("cookies", []):
            self.session.cookies.set(
                c["name"],
                c["value"],
                domain=c.get("domain", "").lstrip("."),
                path=c.get("path", "/"),
            )

    # ---- 認証確認 -------------------------------------------------------

    def is_authenticated(self) -> bool:
        """認証必須ページにアクセスし、ログイン状態を確認する。"""
        try:
            resp = self._get(_CHECK_URL)
            if "/login" in resp.url:
                return False
            soup = BeautifulSoup(resp.text, "lxml")
            if soup.select_one('input[name="password"]'):
                return False
            # reCAPTCHA 要素があれば未認証
            if soup.select_one('.g-recaptcha, #recaptcha-element'):
                return False
            return True
        except Exception:
            return False

    # ---- 検索 -----------------------------------------------------------

    def search(
        self,
        title: str,
        author: str = "",
        match: int = 3,
    ) -> dict:
        """
        曲名・著作者名で MusicForest を検索する。

        match: 1=完全一致 / 2=前方一致 / 3=キーワード（部分一致）

        Returns: {
            "source": "MusicForest",
            "search_url": str,
            "results": [
                {
                    "作品名": str,
                    "アーティスト": str,
                    "収録CD": str,
                    "JASRAC作品コード": str,   # data-href の jcd から直接取得
                    "NexTone管理番号": str,    # data-href の ncd から直接取得
                    "_detail_href": str,
                    "_source_table": str,      # "収録曲" / "配信曲" / "作品"
                },
                ...
            ],
            "error": None | str,
            "truncated": bool,
            "debug_html": str,
        }
        """
        out: dict = {
            "source": "MusicForest",
            "search_url": SEARCH_URL,
            "results": [],
            "error": None,
            "truncated": False,
            "debug_html": "",
        }
        try:
            params = {
                "tr":    title,
                "ka":    author,
                "type":  "search-form-title",
                "match": str(match),
            }
            resp = self._get(SEARCH_URL, params=params)

            # ログインページへリダイレクトされたらセッション切れ
            if "/login" in resp.url:
                out["error"] = "セッションが切れています。再ログインしてください。"
                return out

            out["search_url"] = resp.url
            html = resp.text
            out["debug_html"] = html[:4000]

            soup = BeautifulSoup(html, "lxml")
            results = _parse_search_results(soup)
            out["results"] = results

            for tbl_id in ("#track-list", "#haishin-list", "#sakuhin-list"):
                tbl = soup.select_one(tbl_id)
                if tbl and len(tbl.select("tr")) >= 499:
                    out["truncated"] = True

        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else 0
            if status in (401, 403):
                out["error"] = "セッションが切れています。再ログインしてください。"
            else:
                out["error"] = f"HTTP エラー {status}: {e}"
        except MusicForestError as e:
            out["error"] = str(e)
        except requests.exceptions.ConnectionError:
            out["error"] = "接続エラー: MusicForest に接続できません。"
        except requests.exceptions.Timeout:
            out["error"] = f"タイムアウト（{_TIMEOUT}秒）"
        except Exception as e:
            out["error"] = f"予期しないエラー: {type(e).__name__}: {e}"

        return out

    # ---- CD 詳細（委任者/非委任者）-------------------------------------

    def fetch_product_detail(self, album_id: str, track_id: str) -> dict:
        """
        parts/product/detail から※集中管理の委任者/非委任者と IV 区分を取得する。
        Returns: {"集中管理": "委任者" | "非委任者" | "", "IV": "I" | "V" | "", "error": None | str}
        """
        out: dict = {"集中管理": "", "IV": "", "error": None}
        if not album_id or not track_id:
            out["error"] = "album_id / track_id が不明です"
            return out
        url = f"{BASE_URL}/parts/product/detail?album_id={album_id}&track_id={track_id}"
        try:
            resp = self._get(url)
            if "/login" in resp.url.lower():
                out["error"] = "セッションが切れています。再ログインしてください。"
                return out
            soup = BeautifulSoup(resp.text, "lxml")
            active_span = soup.select_one("span.delegation.active")
            if active_span:
                out["集中管理"] = active_span.get_text(strip=True)
            # IV 区分: <th>IV</th> の次 <td> が "I" または "V"
            for th in soup.find_all("th"):
                if th.get_text(strip=True) == "IV":
                    td = th.find_next_sibling("td")
                    if td:
                        iv_text = td.get_text(strip=True)
                        if iv_text in ("I", "V"):
                            out["IV"] = iv_text
                    break
        except requests.exceptions.ConnectionError:
            out["error"] = "接続エラー"
        except Exception as e:
            out["error"] = f"{type(e).__name__}: {e}"
        return out

    # ---- 詳細ページ -----------------------------------------------------

    def get_detail(self, data_href: str) -> dict:
        """
        作品詳細ページから作曲者・作詞者・編曲者などを取得する。

        data_href: button.saku-detail-link の data-href 属性値
                   例: "jcd=123-456-7&ncd=NT000123&refer=track-list"

        Returns: {
            "作品名": str,
            "作品コード": str,         # JASRAC (jcd)
            "NexTone管理番号": str,    # ncd
            "ISWC": str,
            "アーティスト": str,
            "作曲者": str,
            "作詞者": str,
            "編曲者": str,
            "error": None | str,
        }
        """
        out = {
            "作品名": "",
            "作品コード": "",
            "NexTone管理番号": "",
            "ISWC": "",
            "アーティスト": "",
            "作曲者": "",
            "作詞者": "",
            "編曲者": "",
            "error": None,
        }
        try:
            # URL パラメータから jcd / ncd を即時抽出（ページロード前でも取れる）
            params = dict(urllib.parse.parse_qsl(data_href))
            out["作品コード"]      = params.get("jcd", "")
            out["NexTone管理番号"] = params.get("ncd", "")

            url = f"{DETAIL_URL}?{data_href}"
            resp = self._get(url)
            if "/login" in resp.url:
                out["error"] = "セッションが切れています。再ログインしてください。"
                return out

            soup = BeautifulSoup(resp.text, "lxml")
            _parse_detail(soup, out)

        except requests.exceptions.ConnectionError:
            out["error"] = "接続エラー"
        except Exception as e:
            out["error"] = f"{type(e).__name__}: {e}"
        return out


# =====================================================================
# HTML パーサ
# =====================================================================

def _parse_record_company(raw: str) -> str:
    """
    '発売会社／販売会社' セル値からレコード会社名を返す。
    発売会社が有効値なら発売会社を、'-' または空なら販売会社を使用。
    書式例: "A社（発売）／B社（販売）"、"A社"、"-"
    """
    if not raw or raw.strip() in ("", "-", "－"):
        return ""
    # スラッシュ（全角・半角）で分割
    parts = [p.strip() for p in re.split(r"[/／]", raw) if p.strip()]

    def _clean(s: str) -> str:
        s = re.sub(r"[（(][発販]売[）)]", "", s)
        s = re.sub(r"^[発販]売[：:・]?\s*", "", s)
        return s.strip()

    if parts:
        first = _clean(parts[0])
        if first and first not in ("-", "－"):
            return first
    if len(parts) > 1:
        second = _clean(parts[1])
        return second if second not in ("-", "－") else ""
    return ""


def _parse_search_results(soup: BeautifulSoup) -> list[dict]:
    """
    検索結果ページの 3 テーブルから楽曲情報を抽出する。

    各テーブルの行に button.saku-detail-link があれば
    data-href の jcd / ncd から JASRAC / NexTone コードを即時取得する。
    ヘッダー列名をキーにして抽出するため、列順が変わっても安全。
    """
    results: list[dict] = []
    seen_href: set[str] = set()

    for tbl_id, source_label in [
        ("#track-list",   "収録曲"),
        ("#haishin-list", "配信曲"),
        ("#sakuhin-list", "作品"),
    ]:
        tbl = soup.select_one(tbl_id)
        if tbl is None:
            continue

        # ── ヘッダー列名→インデックス マップを構築 ──
        col_map: dict[str, int] = {}
        thead = tbl.find("thead")
        header_row = thead.find("tr") if thead else None
        if not header_row:
            first_tr = tbl.find("tr")
            if first_tr and first_tr.find("th"):
                header_row = first_tr
        if header_row:
            for idx, cell in enumerate(header_row.find_all(["th", "td"])):
                col_map[cell.get_text(strip=True)] = idx

        def _cell(row_cells: list, col_name: str) -> str:
            """列名の部分一致でセルテキストを返す。"""
            for name, idx in col_map.items():
                if col_name in name and idx < len(row_cells):
                    return row_cells[idx]
            return ""

        for row in tbl.select("tr"):
            btn = row.select_one("button.saku-detail-link")
            if btn is None:
                continue
            data_href = (btn.get("data-href") or "").strip()
            if not data_href or data_href in seen_href:
                continue
            seen_href.add(data_href)

            params = dict(urllib.parse.parse_qsl(data_href))
            jcd = params.get("jcd", "")
            ncd = params.get("ncd", "")

            # CD商品タイトル anchor から album_id / track_id を取得
            _cd_a = row.select_one("a.collapseDetail[data-target]")
            album_id = str(_cd_a.get("data-target", "")) if _cd_a else ""
            track_id = str(_cd_a.get("data-track",  "")) if _cd_a else ""

            # データ行の全セル（th=No列 + td を含む）
            row_cells = [c.get_text(" ", strip=True) for c in row.find_all(["td", "th"])]

            if col_map:
                title         = _cell(row_cells, "曲名")
                artist        = _cell(row_cells, "アーティスト")
                catalog       = _cell(row_cells, "品番")
                cd_title      = _cell(row_cells, "CD商品タイトル")
                publisher_raw = _cell(row_cells, "発売会社")   # "発売会社／販売会社" に部分一致
                isrc          = _cell(row_cells, "ISRC")
            else:
                # ヘッダー取得不可時の位置ベースフォールバック
                tds = [c.get_text(" ", strip=True) for c in row.find_all("td")]
                title         = tds[0] if tds else ""
                artist        = tds[1] if len(tds) > 1 else ""
                catalog       = ""
                cd_title      = ""
                publisher_raw = ""
                isrc          = ""

            if not title:
                title = btn.get_text(strip=True)

            # 品番が空の場合: collapseDetail アンカーのテキストから補完
            # 2枚組など複数 CD 行の場合、select_one が最初のアンカー = このトラック収録 CD を返す
            if not catalog and _cd_a:
                _anc = _cd_a.get_text(" ", strip=True)
                # 品番パターン: 英字 2〜5 文字 + ハイフン + 数字 (例: COCP-12345, SRCL-1234)
                _cat_m = re.match(r"^([A-Z]{2,5}-\d{3,7}(?:[~〜]\d+)?)\s*(.*)", _anc)
                if _cat_m:
                    catalog = _cat_m.group(1)
                    if not cd_title:
                        cd_title = _cat_m.group(2).strip()

            cd_display     = " ".join(x for x in [catalog, cd_title] if x).strip()
            record_company = _parse_record_company(publisher_raw)

            item: dict = {
                "作品名":           title,
                "アーティスト":     artist,
                "品番":             catalog,
                "CD商品タイトル":   cd_title,
                "収録CD":           cd_display,
                "ISRC":            isrc,
                "発売会社販売会社":  publisher_raw,
                "レコード会社名":    record_company,
                "JASRAC作品コード":  jcd,
                "NexTone管理番号":  ncd,
                "_detail_href":     data_href,
                "_source_table":    source_label,
                "_album_id":        album_id,
                "_track_id":        track_id,
            }
            if title or jcd or ncd:
                results.append(item)

    return results


def _parse_detail(soup: BeautifulSoup, out: dict) -> None:
    """
    作品詳細ページ HTML から基本情報と権利者情報を抽出して out に書き込む。

    MusicForest の詳細ページは dl>dt+dd と tr>th+td の両パターンが混在する。
    権利者テーブルでは <th> や最初の <td> に「作曲」「作詞」「編曲」ロールが入る。
    """
    # ── 基本情報 (dl/dt/dd パターン) ──────────────────────────────────
    for dl in soup.select("dl"):
        dts = dl.select("dt")
        dds = dl.select("dd")
        for dt, dd in zip(dts, dds):
            label = dt.get_text(strip=True)
            value = dd.get_text(" ", strip=True)
            _apply_basic(label, value, out)

    # ── 基本情報 (table/tr/th+td パターン) ────────────────────────────
    for row in soup.select("tr"):
        th = row.select_one("th")
        td = row.select_one("td")
        if th and td:
            _apply_basic(th.get_text(strip=True), td.get_text(" ", strip=True), out)

    # ── 権利者情報: 作曲 / 作詞 / 編曲 ロールを持つ行 ─────────────────
    composers: list[str] = []
    lyricists: list[str] = []
    arrangers: list[str] = []

    for row in soup.select("tr"):
        cells = row.select("td, th")
        if len(cells) < 2:
            continue
        role  = cells[0].get_text(strip=True)
        name  = cells[1].get_text(" ", strip=True)
        if not name or name.lower() == "nan":
            continue
        if "作曲" in role:
            composers.append(name)
        if "作詞" in role:   # elif → if: 「作詞作曲」は作曲者・作詞者両方に入れる
            lyricists.append(name)
        if "編曲" in role:
            arrangers.append(name)

    if composers and not out["作曲者"]:
        out["作曲者"] = "/".join(composers)
    if lyricists and not out["作詞者"]:
        out["作詞者"] = "/".join(lyricists)
    if arrangers and not out["編曲者"]:
        out["編曲者"] = "/".join(arrangers)


def _apply_basic(label: str, value: str, out: dict) -> None:
    """ラベル → フィールドのマッピングを out に反映する（空のフィールドのみ）。"""
    if not value or value.lower() == "nan":
        return
    if "作品名" in label and not out["作品名"]:
        out["作品名"] = value
    elif ("作品コード" in label or "JASRAC" in label) and not out["作品コード"]:
        # "作品コード" が "NexTone管理番号" の列を誤って取るのを防ぐ
        if "NexTone" not in label and "管理番号" not in label:
            out["作品コード"] = value
    elif "ISWC" in label and not out["ISWC"]:
        out["ISWC"] = value
    elif "アーティスト" in label and not out["アーティスト"]:
        out["アーティスト"] = value


# =====================================================================
# ファクトリ / ユーティリティ
# =====================================================================

def get_state_path() -> Path:
    """環境変数 MINC_STATE_PATH → デフォルトの search_music パスを返す。"""
    env = os.environ.get("MINC_STATE_PATH", "")
    return Path(env) if env else _DEFAULT_STATE_PATH


def load_client() -> MusicForestClient:
    """保存済み Cookie からクライアントを作成して返す。失敗時は MusicForestError。"""
    client = MusicForestClient()
    client.load_state(get_state_path())
    return client


def _session_age_str(state_path: Path) -> str:
    """state.json の mtime から経過時間を返す（例: '45 分前'）。"""
    try:
        age_sec = time.time() - state_path.stat().st_mtime
        age_min = int(age_sec / 60)
        if age_min < 60:
            return f"{age_min} 分前"
        h, m = divmod(age_min, 60)
        return f"{h} 時間 {m} 分前"
    except Exception:
        return ""


def check_session(client: Optional["MusicForestClient"] = None) -> tuple[bool, str]:
    """
    MusicForest のセッション状態を確認する。
    client を渡すと既存セッションを使い回す（セッションローテーション対策）。
    Returns: (ok: bool, message: str)
    """
    state_path = get_state_path()
    if not state_path.exists():
        return False, f"Cookie ファイルが見つかりません:\n{state_path}"
    age = _session_age_str(state_path)
    age_warn = ""
    try:
        age_sec = time.time() - state_path.stat().st_mtime
        if age_sec > 9000:  # 2.5時間超で警告
            age_warn = " ⚠️ まもなく期限切れの可能性"
    except Exception:
        pass
    try:
        if client is None:
            client = MusicForestClient()
            client.load_state(state_path)
        ok = client.is_authenticated()
        if ok:
            return True, f"ログイン済み（ログイン: {age}）{age_warn}  [{state_path.name}]"
        return False, f"セッション切れ（ログイン: {age}） — 再ログインが必要です"
    except MusicForestError as e:
        return False, str(e)
    except Exception as e:
        return False, f"確認エラー: {e}"
