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
import unicodedata
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


def _norm_title(s: str) -> str:
    """曲名照合用の正規化（NFKC・小文字化・空白/記号ゆれの除去）。"""
    _s = unicodedata.normalize("NFKC", str(s or "")).lower()
    return re.sub(r"[\s　・･'\"’“”~〜\-−ー]", "", _s)


def _product_detail_url(album_id: str, track_id: str = "") -> str:
    """
    CD商品詳細（モーダル）のURLを組み立てる。

    MINC の js/jmdScroll.js（.collapseDetail のクリックハンドラ）と同じ形式:
        track_id あり: /parts/product/detail/?album_id=<id>&track_id=<tid>
        track_id なし: /parts/product/detail/?album_id=<id>
    track_id を空文字で送ると収録曲テーブルが返らないため、必ず省略する。
    """
    _alb = urllib.parse.quote(str(album_id).strip(), safe="-")
    _trk = str(track_id or "").strip()
    url = f"{BASE_URL}/parts/product/detail/?album_id={_alb}"
    if _trk:
        url += f"&track_id={urllib.parse.quote(_trk, safe='-')}"
    return url


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

            # ページ全体の collapseDetail リンクをスキャン
            _all_page_cd: dict[tuple, dict] = {}
            for _a in soup.select("a.collapseDetail[data-target]"):
                _alb = str(_a.get("data-target", "")).strip()
                _trk = str(_a.get("data-track", "")).strip()
                if _alb and _alb.lstrip("-").isdigit() and _trk:
                    _k = (_alb, _trk)
                    if _k not in _all_page_cd:
                        _all_page_cd[_k] = {
                            "label":    _a.get_text(strip=True) or f"CD ({_alb})",
                            "album_id": _alb,
                            "track_id": _trk,
                        }

            # album_id なし結果の作品名セット（これに一致しない別曲リンクを除外）
            _no_alb_names: set[str] = {
                r["作品名"].strip() for r in results
                if not r.get("_album_id") and r.get("作品名")
            }
            # album_id あり結果の (album_id, track_id) → 作品名 マップ
            _has_alb_map: dict[tuple, str] = {
                (_r["_album_id"], _r["_track_id"]): _r.get("作品名", "").strip()
                for _r in results if _r.get("_album_id")
            }

            # フィルタリング:
            #   - results に album_id あり項目として存在し、かつ作品名が _no_alb_names に
            #     含まれない → 別曲の収録曲行由来 → 除外
            #   - results に存在しない（新規CDリンク）→ そのまま含める
            _page_cd_links: list[dict] = []
            for _k, _lnk in _all_page_cd.items():
                if _k in _has_alb_map and _has_alb_map[_k] not in _no_alb_names:
                    continue
                _page_cd_links.append(_lnk)
            out["_page_cd_links"] = _page_cd_links

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
        parts/product/detail から集中管理区分・IV区分・CD情報・ハイライトトラック情報を取得する。

        Returns: {
            "集中管理": "委任者" | "非委任者" | "",
            "IV": "I" | "V" | "",
            "CD商品タイトル": str,
            "品番": str,
            "アーティスト": str,
            "曲名": str,          # ハイライト行の曲名
            "トラック番号": str,   # ハイライト行のトラック番号
            "尺": str,             # ハイライト行の尺
            "error": None | str,
            "debug_html": str,
        }
        """
        out: dict = {
            "集中管理": "", "IV": "",
            "CD商品タイトル": "", "品番": "", "アーティスト": "", "レコード会社名": "",
            "曲名": "", "トラック番号": "", "尺": "",
            "error": None, "debug_html": "",
        }
        if not album_id or not track_id:
            out["error"] = "album_id / track_id が不明です"
            return out
        url = _product_detail_url(album_id, track_id)
        try:
            resp = self._get(url)
            if "/login" in resp.url.lower():
                out["error"] = "セッションが切れています。再ログインしてください。"
                return out
            html = resp.text
            out["debug_html"] = html[:4000]
            soup = BeautifulSoup(html, "lxml")

            # ── 委任者 ────────────────────────────────────────────────────
            active_span = soup.select_one("span.delegation.active")
            if active_span:
                out["集中管理"] = active_span.get_text(strip=True)

            # ── CD商品タイトル: modal-title ──────────────────────────────────
            modal_title = soup.select_one("h4.modal-title")
            if modal_title:
                out["CD商品タイトル"] = modal_title.get_text(strip=True)

            # ── 品番・発売会社: detail_data の div.col-sm-* テキスト ─────────
            for div in soup.select("div.detail_data div[class*='col-sm']"):
                text = div.get_text(" ", strip=True)
                if not out["品番"]:
                    _m = re.match(r"品番[：:]\s*(.+)", text)
                    if _m:
                        out["品番"] = _m.group(1).strip()
                if not out["レコード会社名"] and "発売会社" in text:
                    _m2 = re.search(r"発売会社[：:]\s*([^/\n]+)", text)
                    if _m2:
                        out["レコード会社名"] = _m2.group(1).strip()

            # ── CD メタデータ フォールバック（dl > dt+dd パターン）──────────
            for dl in soup.select("dl"):
                for dt, dd in zip(dl.select("dt"), dl.select("dd")):
                    key = dt.get_text(strip=True)
                    val = dd.get_text(" ", strip=True)
                    if not out["CD商品タイトル"] and any(kw in key for kw in ("タイトル", "商品名", "CD名", "アルバム")):
                        out["CD商品タイトル"] = val
                    if not out["アーティスト"] and "アーティスト" in key:
                        out["アーティスト"] = val
                    if not out["品番"] and any(kw in key for kw in ("品番", "カタログ")):
                        out["品番"] = val
                    if not out["レコード会社名"] and any(kw in key for kw in ("発売会社", "レコード")):
                        out["レコード会社名"] = _parse_record_company(val)

            # ── CD メタデータ フォールバック（th+td パターン）──────────────
            for row in soup.select("tr"):
                ths = row.select("th")
                tds = row.select("td")
                if not ths or not tds:
                    continue
                key = ths[0].get_text(strip=True)
                val = tds[0].get_text(" ", strip=True)
                if not out["CD商品タイトル"] and any(kw in key for kw in ("タイトル", "商品", "CD")):
                    out["CD商品タイトル"] = val
                if not out["アーティスト"] and "アーティスト" in key:
                    out["アーティスト"] = val
                if not out["品番"] and any(kw in key for kw in ("品番", "カタログ")):
                    out["品番"] = val
                if not out["レコード会社名"] and any(kw in key for kw in ("発売会社", "レコード")):
                    out["レコード会社名"] = _parse_record_company(val)
                # IV 区分（既存ロジックを継承）
                if key == "IV":
                    iv_text = val.strip()
                    if iv_text in ("I", "V"):
                        out["IV"] = iv_text

            # ── IV 区分（補完: th 隣接 td パターン）──────────────────────────
            if not out["IV"]:
                for th in soup.find_all("th"):
                    if th.get_text(strip=True) == "IV":
                        td = th.find_next_sibling("td")
                        if td:
                            iv_text = td.get_text(strip=True)
                            if iv_text in ("I", "V"):
                                out["IV"] = iv_text
                        break

            # ── ハイライト行（現在トラック）────────────────────────────────
            hl_row = soup.select_one("tr.high_light, tr.highLight")
            if hl_row:
                tbl = hl_row.find_parent("table")
                col_map: dict[str, int] = {}
                if tbl:
                    hdr_row = (
                        tbl.select_one("thead tr")
                        or next((tr for tr in tbl.select("tr") if tr.find("th")), None)
                    )
                    if hdr_row:
                        for idx, cell in enumerate(hdr_row.find_all(["th", "td"])):
                            col_map[cell.get_text(strip=True)] = idx

                hl_cells = [c.get_text(" ", strip=True) for c in hl_row.find_all(["td", "th"])]

                def _col(kws: list[str]) -> str:
                    for kw in kws:
                        for name, idx in col_map.items():
                            if kw in name and idx < len(hl_cells):
                                return hl_cells[idx]
                    return ""

                if col_map:
                    out["トラック番号"] = _col(["No", "トラック", "番号"])
                    out["曲名"]       = _col(["曲名", "タイトル", "楽曲"])
                    out["尺"]         = _col(["尺", "時間", "分", "length"])
                    # ハイライト行の IV が取れればそちらを優先
                    _hl_iv = _col(["IV"])
                    if _hl_iv in ("I", "V"):
                        out["IV"] = _hl_iv

                # フォールバック: 列ヘッダーなし
                if not out["トラック番号"]:
                    for c in hl_cells:
                        if c.isdigit():
                            out["トラック番号"] = c
                            break
                if not out["曲名"]:
                    cands = [
                        c for c in hl_cells
                        if len(c) >= 2 and c not in ("I", "V")
                        and not c.isdigit()
                        and not re.match(r"^\d+[':′]\d+", c)
                    ]
                    if cands:
                        out["曲名"] = max(cands, key=len)
                if not out["尺"]:
                    for c in hl_cells:
                        if re.match(r"^\d+[':′]\d+", c):
                            out["尺"] = c
                            break
                if not out["IV"]:
                    for c in hl_cells:
                        if c in ("I", "V"):
                            out["IV"] = c
                            break

        except requests.exceptions.ConnectionError:
            out["error"] = "接続エラー"
        except Exception as e:
            out["error"] = f"{type(e).__name__}: {e}"
        return out

    # ---- CD → 収録曲の逆引き ---------------------------------------------

    def find_track_id(self, album_id: str, title: str = "") -> str:
        """
        CD商品リスト（from_saku）には track_id が無いため、曲名検索の結果ページから
        同じ album_id を持つ collapseDetail リンクを探して track_id を拾う。

        見つからなければ空文字を返す。
        """
        _alb = str(album_id).strip()
        if not _alb or not str(title).strip():
            return ""
        try:
            sr = self.search(str(title).strip(), match=3)
        except Exception:
            return ""
        for _r in sr.get("results", []):
            if str(_r.get("_album_id", "")).strip() == _alb and _r.get("_track_id"):
                return str(_r["_track_id"]).strip()
        for _l in sr.get("_page_cd_links", []):
            if str(_l.get("album_id", "")).strip() == _alb and _l.get("track_id"):
                return str(_l["track_id"]).strip()
        return ""

    def fetch_track_list(self, album_id: str, track_id: str = "", title: str = "") -> dict:
        """
        CD商品の収録曲を全曲取得する（CDから曲を逆引きする用途）。

        parts/product/detail は table.cd-detail2-track-list に
        曲順／メドレー／曲名／IV／収録時間／アーティスト／ISRC／
        JASRAC作品コード／NexTone作品コード／著作権管理情報 を全曲分出力する。

        track_id はサーバ側で必須（省略すると
        "The track id field must have a value." が返る）だが、CD商品リストの
        行には track_id が無い。そのため以下の順に候補を試す:
          1. 呼び出し元から渡された track_id
          2. 曲名検索の結果ページから拾った同一 album_id の track_id
          3. track_id=0（ハイライト無しで開けるか）

        Returns: {
            "CD商品タイトル": str, "品番": str, "レコード会社名": str,
            "集中管理": str,
            "tracks": [...], "url": str, "attempts": [{"url","result"}, ...],
            "error": None | str, "debug_html": str,
        }
        """
        out: dict = {
            "CD商品タイトル": "", "品番": "", "レコード会社名": "", "集中管理": "",
            "tracks": [], "url": "", "attempts": [], "error": None, "debug_html": "",
        }
        if not album_id:
            out["error"] = "album_id が不明です"
            return out

        # 試す track_id の候補を組み立てる（重複除去・順序維持）
        _cands: list[str] = []
        for _c in (str(track_id or "").strip(),
                   self.find_track_id(album_id, title) if not str(track_id or "").strip() else "",
                   "0"):
            if _c and _c not in _cands:
                _cands.append(_c)
        if not _cands:
            _cands = ["0"]

        try:
            html = ""
            soup = None
            for _c in _cands:
                url = _product_detail_url(album_id, _c)
                out["url"] = url
                resp = self._get(url)
                if "/login" in resp.url.lower():
                    out["error"] = "セッションが切れています。再ログインしてください。"
                    return out
                html = resp.text
                soup = BeautifulSoup(html, "lxml")
                if soup.select_one("table.cd-detail2-track-list, table[class*='track-list']"):
                    out["attempts"].append({"url": url, "result": "OK"})
                    break
                # バリデーションエラー等のメッセージを拾って記録する
                _msg = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))[:120]
                out["attempts"].append({"url": url, "result": _msg or "収録曲テーブルなし"})
                soup = None

            out["debug_html"] = html[:4000]
            if soup is None:
                out["error"] = (
                    "収録曲テーブルが見つかりませんでした（track_id を特定できません）。"
                    + " ／ ".join(f"{a['url']} → {a['result']}" for a in out["attempts"])
                )
                return out

            _mt = soup.select_one("h4.modal-title")
            if _mt:
                out["CD商品タイトル"] = _mt.get_text(strip=True)
            _dlg = soup.select_one("span.delegation.active")
            if _dlg:
                out["集中管理"] = _dlg.get_text(strip=True)
            for div in soup.select("div.detail_data div[class*='col-sm']"):
                text = div.get_text(" ", strip=True)
                if not out["品番"]:
                    _m = re.match(r"品番[：:]\s*(.+)", text)
                    if _m:
                        out["品番"] = _m.group(1).strip()
                if not out["レコード会社名"] and "発売会社" in text:
                    _m2 = re.search(r"発売会社[：:]\s*([^/\n]+)", text)
                    if _m2:
                        out["レコード会社名"] = _m2.group(1).strip()

            table = soup.select_one("table.cd-detail2-track-list, table[class*='track-list']")

            # ヘッダー行から列位置を作る（列順の変更に耐えるため）
            hdr_row = (
                table.select_one("thead tr")
                or next((tr for tr in table.select("tr") if tr.find("th")), None)
            )
            col_map: dict[str, int] = {}
            if hdr_row:
                for idx, cell in enumerate(hdr_row.find_all(["th", "td"])):
                    col_map[re.sub(r"\s+", "", cell.get_text(" ", strip=True))] = idx

            def _pick(cells: list[str], *kws: str) -> str:
                for kw in kws:
                    for name, idx in col_map.items():
                        if kw in name and idx < len(cells):
                            return cells[idx]
                return ""

            for row in table.select("tr"):
                if row.find("th"):
                    continue
                cells = [c.get_text(" ", strip=True) for c in row.find_all("td")]
                if not cells:
                    continue
                _jcd_c = _pick(cells, "JASRAC")
                _ncd_c = _pick(cells, "NexTone")
                out["tracks"].append({
                    "曲順":            _pick(cells, "曲順", "No"),
                    "メドレー":         _pick(cells, "メドレー"),
                    "曲名":            _pick(cells, "曲名", "タイトル"),
                    "IV":             _pick(cells, "IV"),
                    "収録時間":         _pick(cells, "収録時間", "時間", "尺"),
                    "アーティスト":     _pick(cells, "アーティスト"),
                    "ISRC":           _pick(cells, "ISRC"),
                    "JASRAC作品コード":  "" if _jcd_c == "-" else _jcd_c,
                    "NexTone作品コード": "" if _ncd_c == "-" else _ncd_c,
                    "管理情報":         _pick(cells, "管理情報", "著作権"),
                })

            if not out["tracks"]:
                out["error"] = "収録曲を取得できませんでした。"

        except requests.exceptions.HTTPError as e:
            out["error"] = f"HTTP エラー: {e}"
        except requests.exceptions.ConnectionError:
            out["error"] = "接続エラー: MusicForest に接続できません。"
        except requests.exceptions.Timeout:
            out["error"] = f"タイムアウト（{_TIMEOUT}秒）"
        except Exception as e:
            out["error"] = f"{type(e).__name__}: {e}"
        return out

    # ---- JASRACコードでCDリスト取得 -------------------------------------

    def search_cds_by_jasrac(self, jcd: str, title: str = "", ncd: str = "") -> dict:
        """
        JASRACコードに紐づく収録CD商品リストを MINC から**全件**取得する。

        エンドポイント（MINC の「CD商品リスト」ページそのもの）:
            GET /product/list/from_saku?jcd=<8桁JCD>&ncd=&snm=<曲名>
        このページは table#cd-list に全件（ページング無し）を出力するため、
        1リクエストで 220 件などの完全なリストが得られる。

        手順:
          1. title 未指定なら作品詳細ページ（refer=music/list-product）から曲名を取得
          2. from_saku ページを GET → table#cd-list を全行パース

        Returns: {
            "作品名": str, "作品コード": str, "件数": int,
            "作曲者": str, "作詞者": str,
            "cds": [{
                "No", "品番", "CD商品タイトル", "アーティスト",
                "形態", "曲数", "発売日", "発売会社", "販売会社",
                "レコード会社名", "権利", "初回盤",
                "album_id", "track_id", "label",
            }, ...],
            "error": None | str, "debug_html": str,
        }
        """
        out: dict = {
            "作品名": "", "作品コード": "", "件数": 0,
            "作曲者": "", "作詞者": "",
            "編曲者": "", "訳詞者": "", "ISWC": "", "アーティスト": "",
            "NexTone管理番号": "",
            "cds": [], "配信": [], "search_url": "", "error": None, "debug_html": "",
        }
        cleaned = re.sub(r"[-\s]", "", str(jcd)).upper().strip()
        _ncd = re.sub(r"\s", "", str(ncd)).upper().strip()
        if not cleaned and not _ncd:
            out["error"] = "JASRACコードを入力してください。"
            return out

        # NexTone コード（NT始まり）が jcd 欄に入っている場合は ncd 側へ振り替える
        if re.match(r"^NT", cleaned):
            _ncd = _ncd or cleaned
            cleaned = ""

        if len(cleaned) >= 7:
            formatted = cleaned[:3] + "-" + cleaned[3:7] + "-" + cleaned[7:]
        else:
            formatted = cleaned
        out["作品コード"] = formatted
        out["NexTone管理番号"] = _ncd

        _title = str(title).strip()

        # タイトル未指定 → 作品詳細ページから取得（refer=music/list-product が正しい書式）
        if not _title:
            _detail = self.get_detail(f"jcd={cleaned}&ncd={_ncd}&refer=music/list-product")
            out["debug_html"] = _detail.get("debug_html", "")
            if _detail.get("error"):
                out["error"] = (
                    f"作品詳細の取得に失敗しました（{_detail['error']}）。"
                    "CD情報検索の「曲名」欄に曲名を入力して再検索してください。"
                )
                return out
            _title = _detail.get("作品名", "")
            for _k in ("作曲者", "作詞者", "編曲者", "訳詞者"):
                if _detail.get(_k):
                    out[_k] = _detail[_k]

        if not _title:
            out["error"] = (
                "曲名を取得できませんでした。CD情報検索の「曲名」欄に曲名を入力して再検索してください。"
            )
            return out

        out["作品名"] = _title

        # CD商品リストページを直接取得（ページング無し・全件）
        url = (
            f"{BASE_URL}/product/list/from_saku"
            f"?jcd={urllib.parse.quote(cleaned, safe='-')}"
            f"&ncd={urllib.parse.quote(_ncd, safe='-')}"
            f"&snm={urllib.parse.quote(_title)}"
        )
        out["search_url"] = url
        try:
            resp = self._get(url)
            if "/login" in resp.url.lower():
                out["error"] = "セッションが切れています。再ログインしてください。"
                return out

            html = resp.text
            out["debug_html"] = html[:4000]
            soup = BeautifulSoup(html, "lxml")

            # h1「<曲名>　CD商品リスト」から作品名を補正
            _h1 = soup.select_one("h1")
            if _h1:
                _h1t = re.sub(r"CD商品リスト\s*$", "", _h1.get_text(" ", strip=True)).strip()
                if _h1t:
                    out["作品名"] = _h1t

            # タブ「CD商品 (220件)」から件数
            _m_cnt = re.search(r"CD商品\s*[（(]\s*([\d,]+)\s*件", soup.get_text(" ", strip=True))
            if _m_cnt:
                out["件数"] = int(_m_cnt.group(1).replace(",", ""))

            table = soup.select_one("table#cd-list")
            if table is None:
                out["error"] = (
                    f"CD商品リストのテーブルを取得できませんでした"
                    f"（JASRACコード {cleaned} / 曲名「{_title}」）。"
                    "曲名が正しいか確認してください。"
                )
                return out

            cds: list[dict] = []
            for row in table.select("tbody tr"):
                tds = row.find_all("td")
                if len(tds) < 8:
                    continue

                _a = tds[3].select_one("a.collapseDetail[data-target]")
                album_id = str(_a.get("data-target", "")).strip() if _a else ""
                track_id = str(_a.get("data-track", "")).strip() if _a else ""

                # 権利表示アイコン（jasrac / nextone）と初回盤フラグ
                rights = [
                    c for sp in tds[2].select("span.icon.rights")
                    for c in sp.get("class", []) if c in ("jasrac", "nextone")
                ]
                shokaiban = bool(tds[2].select_one("span.icon.limited"))

                # 形態／曲数（<br> 区切り）
                _kt = [
                    s for s in tds[5].get_text("\n", strip=True).split("\n") if s
                ]
                keitai = _kt[0] if _kt else ""
                kyokusu = _kt[1] if len(_kt) > 1 else ""

                # 発売会社／販売会社（" / " 区切り）
                _cos = [
                    s.strip() for s in
                    tds[7].get_text(" ", strip=True).split("/")
                    if s.strip() and s.strip() != "-"
                ]
                hatsubai = _cos[0] if _cos else ""
                hanbai = _cos[1] if len(_cos) > 1 else ""

                hinban = tds[1].get_text(" ", strip=True)
                cd_title = tds[3].get_text(" ", strip=True)

                cds.append({
                    "No":             tds[0].get_text(" ", strip=True),
                    "品番":            hinban,
                    "CD商品タイトル":   cd_title,
                    "アーティスト":     tds[4].get_text(" ", strip=True),
                    "形態":            keitai,
                    "曲数":            kyokusu,
                    "発売日":          tds[6].get_text(" ", strip=True),
                    "発売会社":         hatsubai,
                    "販売会社":         hanbai,
                    "レコード会社名":    hatsubai,
                    "権利":            rights,
                    "初回盤":          shokaiban,
                    "album_id":       album_id,
                    "track_id":       track_id,
                    "detail_url":     _product_detail_url(album_id, track_id) if album_id else "",
                    "label":          " / ".join([x for x in (hinban, cd_title) if x]) or f"CD ({album_id})",
                })

            out["cds"] = cds
            if not out["件数"]:
                out["件数"] = len(cds)
            if not cds:
                # 作品コードに CD が紐付いていないケース（収録曲行に管理情報ボタンが無く、
                # 配信行にだけ作品コードが付いている等）は検索結果ページから拾い直す。
                _fb = self.search_fallback_by_title(cleaned, _ncd, _title)
                out["cds"] = _fb["cds"]
                out["配信"] = _fb["配信"]
                if _fb["cds"]:
                    out["_cd_fallback"] = True
                    out["件数"] = len(_fb["cds"])
                    out["error"] = (
                        f"作品コード {formatted} にはCD商品が紐付いていなかったため、"
                        f"曲名「{_title}」の検索結果からCD {len(_fb['cds'])} 件を取得しました。"
                    )
                elif _fb["配信"]:
                    out["error"] = (
                        f"JASRACコード {cleaned}（「{_title}」）にCD商品はありませんでした。"
                        f"配信音源が {len(_fb['配信'])} 件見つかりましたので下に表示します。"
                    )
                else:
                    out["error"] = (
                        f"JASRACコード {cleaned}（「{_title}」）に紐づくCD商品が見つかりませんでした。"
                    )

        except requests.exceptions.ConnectionError:
            out["error"] = "接続エラー: MusicForest に接続できません。"
        except requests.exceptions.Timeout:
            out["error"] = f"タイムアウト（{_TIMEOUT}秒）"
        except Exception as e:
            out["error"] = f"{type(e).__name__}: {e}"

        return out

    def search_fallback_by_title(
        self, jcd: str, ncd: str = "", title: str = ""
    ) -> dict:
        """
        作品コード起点の CD商品リストが 0 件だったときに、曲名の検索結果ページから
        収録曲（#track-list）と配信曲（#haishin-list）を拾い直す。

        MINC では CD の収録曲行に作品コードが紐付いていない登録があり
        （管理情報ボタンが無い行。配信曲行にだけ作品コードが付く）、
        その場合 /product/list/from_saku は 0 件を返す。

        行に作品コードがある場合は要求されたコードと一致する行だけを採用し、
        コードが無い行は曲名一致で採用する。

        Returns: {"cds": [search_cds_by_jasrac と同じ形式], "配信": [...]}
        """
        out: dict = {"cds": [], "配信": []}
        _title = str(title).strip()
        if not _title:
            return out
        _j = re.sub(r"[-\s]", "", str(jcd)).upper()
        _n = re.sub(r"[-\s]", "", str(ncd)).upper()
        _tnorm = _norm_title(_title)
        try:
            res = self.search(_title, match=3)
        except Exception:
            return out

        for it in res.get("results", []):
            _src = it.get("_source_table", "")
            if _src not in ("収録曲", "配信曲"):
                continue
            _ij = re.sub(r"[-\s]", "", str(it.get("JASRAC作品コード", ""))).upper()
            _in = re.sub(r"[-\s]", "", str(it.get("NexTone管理番号", ""))).upper()
            if _ij or _in:
                # コードを持つ行は一致必須（同名異作品を拾わない）
                if not ((_j and _ij == _j) or (_n and _in == _n)):
                    continue
            else:
                # コード無しの行（今回のケース）は曲名一致で採用
                _rt = _norm_title(it.get("作品名", ""))
                if not _rt or (_rt != _tnorm and _tnorm not in _rt):
                    continue

            if _src == "配信曲":
                out["配信"].append({
                    "曲名":            it.get("作品名", ""),
                    "アーティスト":     it.get("アーティスト", ""),
                    "アルバム名":       it.get("アルバム名", "") or it.get("CD商品タイトル", ""),
                    "ISRC":           it.get("ISRC", ""),
                    "配信日":          it.get("配信日", ""),
                    "JASRAC作品コード": it.get("JASRAC作品コード", ""),
                    "NexTone管理番号":  it.get("NexTone管理番号", ""),
                })
                continue

            _alb = str(it.get("_album_id", "") or "")
            if not _alb:
                continue
            _trk = str(it.get("_track_id", "") or "")
            _hinban = it.get("品番", "")
            _cdt = it.get("CD商品タイトル", "")
            _cos = [
                s.strip() for s in str(it.get("発売会社販売会社", "")).split("/")
                if s.strip() and s.strip() != "-"
            ]
            out["cds"].append({
                "No":            str(len(out["cds"]) + 1),
                "品番":           _hinban,
                "CD商品タイトル":  _cdt,
                "アーティスト":    it.get("アーティスト", ""),
                "形態":           "",
                "曲数":           "",
                "発売日":         it.get("配信日", ""),
                "発売会社":        _cos[0] if _cos else it.get("レコード会社名", ""),
                "販売会社":        _cos[1] if len(_cos) > 1 else "",
                "レコード会社名":   it.get("レコード会社名", ""),
                "権利":           [],
                "初回盤":         False,
                "ISRC":          it.get("ISRC", ""),
                "album_id":      _alb,
                "track_id":      _trk,
                "detail_url":    _product_detail_url(_alb, _trk) if _alb else "",
                "label":         " / ".join([x for x in (_hinban, _cdt) if x]) or f"CD ({_alb})",
            })
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
            "訳詞者": str,
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
            "訳詞者": "",
            "error": None,
        }
        try:
            # URL パラメータから jcd / ncd を即時抽出（ページロード前でも取れる）
            params = dict(urllib.parse.parse_qsl(data_href))
            jcd = params.get("jcd", "")
            ncd = params.get("ncd", "")
            # jcd に NexTone コード（NT 始まり）が入る場合は ncd へ移す
            if re.match(r"^NT", jcd, re.IGNORECASE):
                ncd = ncd or jcd
                jcd = ""
            out["作品コード"]      = jcd
            out["NexTone管理番号"] = ncd

            url = f"{DETAIL_URL}?{data_href}"
            resp = self._get(url)

            html = resp.text
            if "/login" in resp.url.lower():
                out["error"] = "セッションが切れています。再ログインしてください。"
                return out

            soup = BeautifulSoup(html, "lxml")
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

            # CD商品タイトル anchor から album_id / track_id を取得
            _cd_a = row.select_one("a.collapseDetail[data-target]")
            album_id = str(_cd_a.get("data-target", "")) if _cd_a else ""
            track_id = str(_cd_a.get("data-track",  "")) if _cd_a else ""

            if btn is None:
                # 著作権管理情報が紐付いていない行（作品コード無し）。
                # CD へのリンク（collapseDetail）を持つ行はCD情報として拾う価値があるので
                # 捨てずに、コード空のまま結果に含める。
                if not (album_id and album_id.lstrip("-").isdigit()):
                    continue
                data_href = ""
                jcd = ncd = ""
                _key = f"_nocode:{album_id}:{track_id}"
                if _key in seen_href:
                    continue
                seen_href.add(_key)
            else:
                data_href = (btn.get("data-href") or "").strip()
                if not data_href or data_href in seen_href:
                    continue
                seen_href.add(data_href)

                params = dict(urllib.parse.parse_qsl(data_href))
                jcd = params.get("jcd", "")
                ncd = params.get("ncd", "")
                # jcd に NexTone コード（NT 始まり）が入る場合は ncd へ移す
                if re.match(r"^NT", jcd, re.IGNORECASE):
                    ncd = ncd or jcd
                    jcd = ""

            # album_id が数字列でない場合（例: "#..." の Bootstrap data-target）はリセット
            if album_id and not album_id.lstrip("-").isdigit():
                album_id = ""
                track_id = ""

            # フォールバック①: data-album-id / data-track-id 属性パターン
            if not album_id:
                _alt = (row.select_one("[data-album-id]")
                        or row.select_one("[data-albumid]")
                        or row.select_one("[data-album_id]"))
                if _alt:
                    album_id = str(
                        _alt.get("data-album-id") or _alt.get("data-albumid") or _alt.get("data-album_id") or ""
                    )
                    track_id = str(
                        _alt.get("data-track-id") or _alt.get("data-trackid") or _alt.get("data-track_id") or ""
                    )

            # フォールバック②: 行のHTML全体から /parts/product/detail の URL パターンを正規表現で検索
            if not album_id:
                _row_html = str(row)
                _m_alb = re.search(r'album_id=(\d{10,})', _row_html)
                _m_trk = re.search(r'track_id=(\d+)', _row_html)
                if _m_alb and _m_trk:
                    album_id = _m_alb.group(1)
                    track_id = _m_trk.group(1)

            # データ行の全セル（th=No列 + td を含む）
            row_cells = [c.get_text(" ", strip=True) for c in row.find_all(["td", "th"])]

            if col_map:
                title         = _cell(row_cells, "曲名")
                artist        = _cell(row_cells, "アーティスト")
                catalog       = _cell(row_cells, "品番")
                cd_title      = _cell(row_cells, "CD商品タイトル")
                publisher_raw = _cell(row_cells, "発売会社")   # "発売会社／販売会社" に部分一致
                isrc          = _cell(row_cells, "ISRC")
                # 配信曲テーブルは「アルバム名」「配信日」列（CD商品タイトル／発売日は無い）
                album_name    = _cell(row_cells, "アルバム") or cd_title
                release_date  = _cell(row_cells, "配信") or _cell(row_cells, "発売日")
            else:
                # ヘッダー取得不可時の位置ベースフォールバック
                tds = [c.get_text(" ", strip=True) for c in row.find_all("td")]
                title         = tds[0] if tds else ""
                artist        = tds[1] if len(tds) > 1 else ""
                catalog       = ""
                cd_title      = ""
                publisher_raw = ""
                isrc          = ""
                album_name    = ""
                release_date  = ""

            if not title and btn is not None:
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

            # album_id なし行の中に埋め込まれた collapseDetail リンクを収集
            # （配信曲行が収録CDへの参照リンクを持つ場合に対応）
            _row_cd_links: list[dict] = []
            if not album_id:
                for _a in row.select("a.collapseDetail[data-target]"):
                    _ra = str(_a.get("data-target", "")).strip()
                    _rt = str(_a.get("data-track", "")).strip()
                    if _ra and _ra.lstrip("-").isdigit() and _rt:
                        _row_cd_links.append({
                            "label":    _a.get_text(strip=True) or f"CD ({_ra})",
                            "album_id": _ra,
                            "track_id": _rt,
                        })

            item: dict = {
                "作品名":           title,
                "アーティスト":     artist,
                "品番":             catalog,
                "CD商品タイトル":   cd_title,
                "収録CD":           cd_display,
                "アルバム名":        album_name,
                "配信日":           release_date,
                "ISRC":            isrc,
                "発売会社販売会社":  publisher_raw,
                "レコード会社名":    record_company,
                "JASRAC作品コード":  jcd,
                "NexTone管理番号":  ncd,
                "_detail_href":     data_href,
                "_source_table":    source_label,
                "_album_id":        album_id,
                "_track_id":        track_id,
                "_row_cd_links":    _row_cd_links,  # この行内の CD リンク（album_id なし行のみ）
                "_row_html":        str(row)[:1200],  # デバッグ用
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
    # ロール列は cells[0] とは限らない（空セル・番号列が先頭に来るレイアウトに対応）
    _ROLE_KEYWORDS = ["作曲", "作詞", "編曲", "訳詞"]
    _NAME_SKIP = _ROLE_KEYWORDS + ["権利", "種別", "割合", "比率", "コード", "番号", "登録"]

    composers: list[str] = []
    lyricists: list[str] = []
    arrangers: list[str] = []
    translators: list[str] = []

    for row in soup.select("tr"):
        cells = row.select("td, th")
        cell_texts = [c.get_text(" ", strip=True) for c in cells]
        if len(cell_texts) < 2:
            continue

        # ロールが入っているセルを探す
        role_idx = next(
            (ci for ci, t in enumerate(cell_texts)
             if any(kw in t for kw in _ROLE_KEYWORDS)),
            None,
        )
        if role_idx is None:
            continue
        role = cell_texts[role_idx]

        # 名前: ロール以外のセルから「2文字以上・数字だけでない・ヘッダーキーワードなし」を選ぶ
        name = next(
            (t for ci, t in enumerate(cell_texts)
             if ci != role_idx
             and t and t.lower() != "nan"
             and len(t.replace(" ", "").replace("　", "")) >= 2
             and not t.replace(" ", "").replace("　", "").isdigit()
             and not any(kw in t for kw in _NAME_SKIP)),
            "",
        )
        if not name:
            continue

        if "作曲" in role:
            composers.append(name)
        if "訳詞" in role:
            translators.append(name)
        elif "作詞" in role:
            lyricists.append(name)
        if "編曲" in role:
            arrangers.append(name)

    if composers and not out["作曲者"]:
        out["作曲者"] = "/".join(composers)
    if lyricists and not out["作詞者"]:
        out["作詞者"] = "/".join(lyricists)
    if arrangers and not out["編曲者"]:
        out["編曲者"] = "/".join(arrangers)
    if translators and not out.get("訳詞者"):
        out["訳詞者"] = "/".join(translators)


def _apply_basic(label: str, value: str, out: dict) -> None:
    """ラベル → フィールドのマッピングを out に反映する（空のフィールドのみ）。"""
    if not value or value.lower() == "nan":
        return
    if "作品名" in label and not out["作品名"]:
        out["作品名"] = value
    elif ("作品コード" in label or "JASRAC" in label) and not out["作品コード"]:
        # "作品コード" が "NexTone管理番号" の列を誤って取るのを防ぐ
        if "NexTone" not in label and "管理番号" not in label:
            # 値が NexTone コード形式（NT 始まり）なら JASRAC フィールドには入れない
            if not re.match(r"^NT", value, re.IGNORECASE):
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


def _read_chrome_minc_cookies_win32(domain_filter: str = "minc.or.jp") -> list[dict]:
    """
    pywin32 + pycryptodomex を使い Chrome が起動中でも Cookie を読み込む。
    管理者権限不要。FILE_SHARE_* フラグで共有アクセスし、AES-GCM で復号する。
    """
    import base64
    import sqlite3
    import tempfile

    try:
        import win32file  # type: ignore
        import win32crypt  # type: ignore
        from Cryptodome.Cipher import AES as _AES  # type: ignore
    except ImportError as exc:
        raise ImportError(f"pywin32 / pycryptodomex が必要: {exc}") from exc

    local_app  = os.environ.get("LOCALAPPDATA", "")
    chrome_dir = os.path.join(local_app, "Google", "Chrome", "User Data")
    cookie_src = os.path.join(chrome_dir, "Default", "Network", "Cookies")
    ls_path    = os.path.join(chrome_dir, "Local State")

    if not os.path.exists(cookie_src):
        raise FileNotFoundError(f"Chrome Cookie ファイルが見つかりません: {cookie_src}")
    if not os.path.exists(ls_path):
        raise FileNotFoundError(f"Chrome Local State が見つかりません: {ls_path}")

    # AES キーを Local State から取得（DPAPI で暗号化されている）
    with open(ls_path, encoding="utf-8") as _f:
        _ls = json.load(_f)
    _enc_key_b64 = _ls.get("os_crypt", {}).get("encrypted_key", "")
    if not _enc_key_b64:
        raise ValueError("encrypted_key が Local State に見つかりません")
    _dpapi_blob = base64.b64decode(_enc_key_b64)[5:]  # 先頭 5 byte は 'DPAPI' 識別子
    _aes_key    = win32crypt.CryptUnprotectData(_dpapi_blob, None, None, None, 0)[1]

    # Chrome Cookie ファイルを共有フラグ付きで開いて生バイト読み込み（ロック中でも可）
    _h = win32file.CreateFile(
        cookie_src,
        0x80000000,  # GENERIC_READ
        0x00000007,  # FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
        None,
        3,           # OPEN_EXISTING
        0x00000080,  # FILE_ATTRIBUTE_NORMAL
        None,
    )
    _file_size = os.path.getsize(cookie_src)
    _, _raw = win32file.ReadFile(_h, _file_size)
    _h.Close()

    # 一時 SQLite ファイルに書き出して読み込む（immutable モードで整合性エラーを回避）
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sqlite") as _tmp:
        _tmp.write(_raw)
        _tmp_path = _tmp.name

    result: list[dict] = []
    try:
        _conn = sqlite3.connect(f"file:{_tmp_path}?mode=ro&immutable=1", uri=True)
        _cur  = _conn.cursor()
        _cur.execute(
            "SELECT host_key, name, value, encrypted_value, path, expires_utc, is_secure "
            "FROM cookies WHERE host_key LIKE ?",
            (f"%{domain_filter}%",),
        )
        for _host, _name, _val, _enc_val, _path, _exp, _sec in _cur.fetchall():
            _decoded = _val or ""
            if not _decoded and _enc_val:
                try:
                    _ev = bytes(_enc_val) if not isinstance(_enc_val, bytes) else _enc_val
                    if _ev[:3] in (b"v10", b"v11", b"v20"):
                        _nonce  = _ev[3:15]
                        _ctext  = _ev[15:]
                        _cipher = _AES.new(_aes_key, _AES.MODE_GCM, nonce=_nonce)
                        _decoded = _cipher.decrypt(_ctext)[:-16].decode("utf-8", errors="replace")
                    elif _ev:
                        _decoded = win32crypt.CryptUnprotectData(_ev, None, None, None, 0)[1].decode("utf-8", errors="replace")
                except Exception:
                    _decoded = ""
            result.append({
                "name":    _name,
                "value":   _decoded,
                "domain":  _host,
                "path":    _path or "/",
                "expires": _exp,
                "secure":  bool(_sec),
            })
        _conn.close()
    finally:
        try:
            os.unlink(_tmp_path)
        except Exception:
            pass

    return result


def _save_cookies_to_state(cookies: list[dict], state_path: Path) -> tuple[int, bool]:
    """cookies リストを state.json に保存する。既存の非 minc Cookie は保持。"""
    existing: dict = {"cookies": [], "origins": []}
    if state_path.exists():
        try:
            existing = json.loads(state_path.read_text(encoding="utf-8"))
        except Exception:
            pass
    kept = [c for c in existing.get("cookies", []) if "minc.or.jp" not in c.get("domain", "")]
    for c in cookies:
        kept.append({
            "name":     c["name"],
            "value":    c["value"],
            "domain":   c["domain"],
            "path":     c.get("path", "/"),
            "expires":  c.get("expires", -1),
            "httpOnly": c.get("httpOnly", False),
            "secure":   c.get("secure", True),
            "sameSite": "Lax",
        })
    existing["cookies"] = kept
    state_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")
    sess_found = any(c["name"] == "_sess" for c in cookies)
    return len(cookies), sess_found


def sync_session_from_chrome() -> tuple[bool, str]:
    """
    Chrome の Cookie を読み込み state.json を更新する。

    優先順:
      1. win32file 直接読み込み（Chrome 起動中でも可・管理者権限不要）
      2. browser_cookie3（Chrome が閉じている場合のフォールバック）

    Returns: (ok: bool, message: str)
    """
    state_path = get_state_path()
    last_err: str = ""

    # ── 方法 1: win32file で直接読み込み（Chrome 起動中でも動作） ──
    try:
        cookies = _read_chrome_minc_cookies_win32("minc.or.jp")
        if not cookies:
            last_err = "Chrome に minc.or.jp の Cookie が見つかりませんでした。ブラウザで MINC にログインしてください。"
        else:
            cnt, sess_found = _save_cookies_to_state(cookies, state_path)
            label = "_sess あり ✅" if sess_found else "_sess なし（未ログインの可能性）"
            return True, f"Chrome から {cnt} 件の Cookie を同期しました（{label}）"
    except Exception as e1:
        last_err = f"win32 直接読み込みエラー: {type(e1).__name__}: {e1}"

    # ── 方法 2: browser_cookie3（Chrome が閉じている場合に有効） ──
    try:
        import browser_cookie3  # type: ignore
        cj = browser_cookie3.chrome(domain_name=".minc.or.jp")
        raw = list(cj)
        if not raw:
            raw = list(browser_cookie3.chrome(domain_name="minc.or.jp"))
        if raw:
            cookie_dicts = [
                {"name": c.name, "value": c.value, "domain": c.domain,
                 "path": c.path or "/", "expires": c.expires or -1,
                 "httpOnly": False, "secure": c.secure}
                for c in raw
            ]
            cnt, sess_found = _save_cookies_to_state(cookie_dicts, state_path)
            label = "_sess あり ✅" if sess_found else "_sess なし（未ログインの可能性）"
            return True, f"Chrome から {cnt} 件の Cookie を同期しました（{label}）"
        last_err = "Chrome に minc.or.jp の Cookie が見つかりませんでした。"
    except Exception as e2:
        last_err = f"{last_err}\nbrowser_cookie3 エラー: {type(e2).__name__}: {e2}"

    return False, f"自動同期に失敗しました。\n{last_err}\n\n→ 下の「手動入力」フォームをお使いください。"


def update_sess_cookie(sess_value: str, xsrf_value: str = "") -> None:
    """
    state.json の _sess (および XSRF-TOKEN) の値を手動入力値で上書きする。
    state.json が存在しない場合は最小限の構造で新規作成する。
    """
    state_path = get_state_path()
    existing: dict = {"cookies": [], "origins": []}
    if state_path.exists():
        try:
            existing = json.loads(state_path.read_text(encoding="utf-8"))
        except Exception:
            pass

    expires_ts = time.time() + 10800  # 3 時間後

    def _upsert(name: str, value: str, domain: str, http_only: bool) -> None:
        for c in existing["cookies"]:
            if c.get("name") == name and "minc.or.jp" in c.get("domain", ""):
                c["value"] = value
                c["expires"] = expires_ts
                return
        existing["cookies"].append({
            "name": name, "value": value,
            "domain": domain, "path": "/",
            "expires": expires_ts,
            "httpOnly": http_only, "secure": True, "sameSite": "Lax",
        })

    _upsert("_sess",      sess_value,  "www.minc.or.jp", True)
    if xsrf_value:
        _upsert("XSRF-TOKEN", xsrf_value, "www.minc.or.jp", False)

    state_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")


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
