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

RATE_LIMIT_SEC = 2.0
TIMEOUT = 15

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

_session = requests.Session()
_session.headers.update(HEADERS)

_last_request: dict[str, float] = {}


def _rate_limit(domain: str) -> None:
    elapsed = time.time() - _last_request.get(domain, 0)
    if elapsed < RATE_LIMIT_SEC:
        time.sleep(RATE_LIMIT_SEC - elapsed)
    _last_request[domain] = time.time()


def _get(url: str, params: dict | None = None, domain: str = "") -> requests.Response:
    _rate_limit(domain or url)
    return _session.get(url, params=params, timeout=TIMEOUT, allow_redirects=True)


def _post(url: str, data, domain: str = "", headers: dict | None = None) -> requests.Response:
    _rate_limit(domain or url)
    return _session.post(url, data=data, timeout=TIMEOUT, allow_redirects=True,
                         headers=headers or {})


# =====================================================================
# J-WID (JASRAC)
# =====================================================================

JWID_BASE = "https://www2.jasrac.or.jp/eJwid/"
JWID_SEARCH_URL = JWID_BASE + "main?trxID=A00401-3"

# 利用規約同意済みフラグ（プロセス起動後一度だけ同意すれば OK）
_jwid_agreed: bool = False


def _jwid_agree() -> bool:
    """J-WID 利用規約に同意してセッションを初期化する。成功時 True を返す。"""
    global _jwid_agreed
    try:
        _rate_limit("jasrac.or.jp")
        _session.get(JWID_BASE, timeout=TIMEOUT)          # トップページ取得（Cookie 初期化）
        _rate_limit("jasrac.or.jp")
        resp = _session.post(JWID_BASE + "main?trxID=F00100", timeout=TIMEOUT)
        resp.raise_for_status()
        # "検索画面" が含まれていれば同意成功
        html = resp.content.decode("ms932", errors="replace")
        _jwid_agreed = "検索画面" in html
        return _jwid_agreed
    except Exception:
        _jwid_agreed = False
        return False


def search_jwid(title: str, author: str = "") -> dict:
    """
    J-WID（JASRAC）でタイトル（必要に応じて著作者名）を検索して結果を返す。

    Args:
        title:  検索する曲名
        author: 著作者名ヒント（空でも可）。指定すると J-WID 側で絞り込みを行う

    Returns: {
        "source": "J-WID",
        "search_url": str,
        "results": [{"作品コード":..., "作品名":..., "作曲者":..., "アーティスト":...}, ...],
        "error": str | None,
        "debug_html": str,
    }
    """
    global _jwid_agreed
    out: dict = {
        "source": "J-WID",
        "search_url": JWID_SEARCH_URL,
        "results": [],
        "error": None,
        "debug_html": "",
    }

    try:
        # 必要に応じて利用規約に同意（一度だけ）
        if not _jwid_agreed:
            if not _jwid_agree():
                out["error"] = "J-WID 利用規約への同意に失敗しました。"
                return out

        # フォームデータを MS932 エンコードで送信
        search_params: dict[str, str] = {
            "IN_WORKS_TITLE_NAME1":             title,
            "IN_WORKS_TITLE_OPTION1":           "2",   # 中間一致
            "IN_WORKS_TITLE_TYPE1":             "0",   # 全て
            "IN_KEN_NAME1":                     author,
            "IN_KEN_NAME_JOB1":                 "0",
            "IN_KEN_NAME_OPTION1":              "0",   # 前方一致
            "IN_KEN_NAME2":                     "",
            "IN_KEN_NAME_JOB2":                 "1",
            "IN_KEN_NAME_OPTION2":              "0",
            "IN_ARTIST_NAME1":                  "",
            "IN_ARTIST_NAME_OPTION1":           "0",
            "IN_DEFAULT_SEARCH_WORKS_NAIGAI":   "0",   # 全て
            "CMD_SEARCH":                       "",
            "IN_DEFAULT_WORKS_KOUHO_MAX":       "20",
            "IN_DEFAULT_WORKS_KOUHO_SEQ":       "1",
            "RESULT_CURRENT_PAGE":              "1",
            "OLD_KENRISYA_DISPLAY_TYPE":        "",
            "IN_WORKS_TITLE_CONDITION":         "0",
            "IN_WORKS_TITLE_NAME2":             "",
            "IN_WORKS_TITLE_TYPE2":             "0",
            "IN_WORKS_TITLE_OPTION2":           "0",
            "IN_KEN_NAME_CONDITION":            "0",
            "IN_ARTIST_NAME_CONDITION":         "0",
            "IN_ARTIST_NAME2":                  "",
            "IN_ARTIST_NAME_OPTION2":           "0",
        }
        # MS932 で URL エンコード（サーバーが MS932 を期待するため）
        body = urllib.parse.urlencode(search_params, encoding="ms932").encode("ascii")

        _rate_limit("jasrac.or.jp")
        resp = _session.post(
            JWID_SEARCH_URL,
            data=body,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Referer": JWID_BASE + "main?trxID=F00100",
            },
            timeout=TIMEOUT,
        )
        resp.raise_for_status()

        html = resp.content.decode("ms932", errors="replace")
        out["debug_html"] = html[:3000]

        # セッション切れ（エラー021 / ValidatorApplicationErrException）→ 再同意して1回リトライ
        if "エラー番号021" in html or "ValidatorApplicationErrException" in html or "エラー" in html[:500]:
            _jwid_agreed = False
            if _jwid_agree():
                _rate_limit("jasrac.or.jp")
                resp = _session.post(JWID_SEARCH_URL, data=body,
                                     headers={"Content-Type": "application/x-www-form-urlencoded"},
                                     timeout=TIMEOUT)
                html = resp.content.decode("ms932", errors="replace")
                out["debug_html"] = html[:3000]

        if "検索条件に該当するデータが見つかりませんでした" in html:
            return out  # 0件（正常）

        soup = BeautifulSoup(html, "lxml")
        out["results"] = _parse_jwid_table(soup)

    except requests.exceptions.ConnectionError:
        out["error"] = "接続エラー: J-WID サイトに接続できません。"
    except requests.exceptions.Timeout:
        out["error"] = f"タイムアウト（{TIMEOUT}秒）"
    except requests.exceptions.HTTPError as e:
        out["error"] = f"HTTP エラー: {e}"
    except Exception as e:
        out["error"] = f"予期しないエラー: {type(e).__name__}: {e}"

    return out


def _parse_jwid_table(soup: BeautifulSoup) -> list[dict]:
    """J-WID 検索結果 HTML からデータ行を抽出する。

    対象テーブル: <table class="search-result">
    列構造:
      data-role="result-code"   → 作品コード
      data-role="result-title"  → 作品名
      data-role="result-author" → 著作者名（作曲者 / 作詞者 混在）
      data-role="result-artist" → アーティスト名
    詳細リンク: 各行末の <a class="btn btn-detail AUTO_JUMP">
    """
    results = []

    tbl = soup.find("table", class_="search-result")
    if tbl is None:
        return results

    for row in tbl.find_all("tr"):
        code_td   = row.find("td", attrs={"data-role": "result-code"})
        title_td  = row.find("td", attrs={"data-role": "result-title"})
        author_td = row.find("td", attrs={"data-role": "result-author"})
        artist_td = row.find("td", attrs={"data-role": "result-artist"})

        if not title_td:
            continue

        # 詳細リンク: <a class="btn btn-detail AUTO_JUMP" href="main?trxID=F20101&WORKS_CD=...">
        detail_url = ""
        works_cd = ""
        detail_a = row.find("a", class_="btn-detail")
        if detail_a and detail_a.get("href"):
            href = detail_a["href"]
            if not href.startswith("http"):
                href = JWID_BASE + href.lstrip("./")
            detail_url = href
            m = re.search(r"WORKS_CD=([^&]+)", href)
            if m:
                works_cd = m.group(1)

        item = {
            "作品コード":   code_td.get_text(strip=True)   if code_td   else "",
            "WORKS_CD":     works_cd,
            "作品名":       title_td.get_text(strip=True)  if title_td  else "",
            "著作者名":     author_td.get_text(strip=True) if author_td else "",
            "アーティスト": artist_td.get_text(strip=True) if artist_td else "",
            "_detail_url":  detail_url,
        }
        if any(v for k, v in item.items() if not k.startswith("_") and k != "WORKS_CD"):
            results.append(item)

    return results


def _parse_management_status(soup: BeautifulSoup) -> dict[str, str]:
    """
    J-WID 詳細ページの 管理状況(利用分野) セクションから ○△× を抽出する。

    SVG形状: circle → ○, polygon → △, line/path → ×, 要素なし → ×

    Returns: {"演奏会等": "○", "上映/BGM": "○", "放送": "○", ...}
    """
    status: dict[str, str] = {}
    mgmt = soup.find("div", class_="management")
    if not mgmt:
        return status

    for dl in mgmt.find_all("dl"):
        dd = dl.find("dd")
        if not dd:
            continue
        for li in dd.find_all("li"):
            a = li.find("a", class_="field")
            if not a:
                continue

            # フィールド名: 2番目の span (balloon 内 full name) または 1番目
            spans = a.find_all("span")
            if len(spans) >= 2:
                field_name = spans[1].get_text(strip=True)
            elif spans:
                field_name = spans[0].get_text(strip=True)
            else:
                continue
            if not field_name:
                continue

            # SVG形状で管理状況を判定
            svg = li.find("svg")
            if svg:
                child_names = [c.name for c in svg.children if hasattr(c, "name") and c.name]
                if "circle" in child_names:
                    icon = "○"
                elif "polygon" in child_names:
                    icon = "△"
                else:
                    icon = "×"
            else:
                icon = "×"

            status[field_name] = icon

    return status


def fetch_jwid_detail(detail_url: str) -> dict:
    """
    J-WID 作品詳細ページから作曲者・作詞者・編曲者・管理状況を取得する。

    Returns: {
        "作曲者": str, "作詞者": str, "編曲者": str,
        "著作者リスト": [{"役割": str, "氏名": str, "所属団体": str}, ...],
        "管理状況": {"演奏会等": "○", ...},
        "error": str | None,
    }
    """
    out: dict = {
        "作曲者": "", "作詞者": "", "編曲者": "", "訳詞者": "",
        "著作者リスト": [],
        "管理状況": {},
        "error": None,
    }
    if not detail_url:
        out["error"] = "detail_url が空"
        return out
    global _jwid_agreed
    if not _jwid_agreed:
        _jwid_agree()
    try:
        _rate_limit("jasrac.or.jp")
        resp = _session.get(detail_url, timeout=TIMEOUT)
        resp.raise_for_status()
        html = resp.content.decode("ms932", errors="replace")

        # セッション切れ → 再同意して1回リトライ
        if "ValidatorApplicationErrException" in html or "エラー番号021" in html or "エラー" in html[:500]:
            _jwid_agreed = False
            if _jwid_agree():
                _rate_limit("jasrac.or.jp")
                resp = _session.get(detail_url, timeout=TIMEOUT)
                html = resp.content.decode("ms932", errors="replace")

        soup = BeautifulSoup(html, "lxml")

        composers, lyricists, arrangers, translators = [], [], [], []

        # 著作者テーブルは各利用分野タブに重複するため「未選択」デフォルトタブのみ参照
        # div#tab-def > div.PC > table.detail: No. | 著作者名 | 識別 | 契約 | 所属団体 | 特記
        tab_def = soup.find("div", id="tab-def") or soup
        pc_div  = tab_def.find("div", class_="PC") or tab_def
        for tbl in pc_div.find_all("table", class_="detail"):
            for row in tbl.find_all("tr"):
                tds = row.find_all("td")
                if len(tds) < 5:
                    continue
                no_text   = tds[0].get_text(strip=True)
                if not no_text.isdigit():
                    continue
                name_text = tds[1].get_text(strip=True)
                role_text = tds[2].get_text(strip=True)
                society   = tds[4].get_text(strip=True)
                if not name_text:
                    continue
                out["著作者リスト"].append({"役割": role_text, "氏名": name_text, "所属団体": society})
                if "作曲" in role_text:
                    composers.append(name_text)
                if "訳詞" in role_text:   # 訳詞は作詞より先にチェック
                    translators.append(name_text)
                elif "作詞" in role_text:  # 「作詞作曲」も作詞者に入る（作曲は上の if で捕捉済み）
                    lyricists.append(name_text)
                if "編曲" in role_text:
                    arrangers.append(name_text)

        out["作曲者"] = "/".join(composers)
        out["作詞者"] = "/".join(lyricists)
        out["編曲者"] = "/".join(arrangers)
        out["訳詞者"] = "/".join(translators)
        out["管理状況"] = _parse_management_status(soup)

    except Exception as e:
        out["error"] = f"{type(e).__name__}: {e}"
    return out


def fetch_jwid_rights_by_code(jasrac_code: str) -> dict:
    """
    JASRAC作品コード（例: 052-1235-9）から管理状況・著作者情報を取得する。

    先行 search_jwid() なしで単独利用可能。
    """
    global _jwid_agreed
    if not _jwid_agreed:
        _jwid_agree()

    works_cd = jasrac_code.replace("-", "").replace("－", "").strip()
    if not works_cd:
        return {
            "作曲者": "", "作詞者": "", "編曲者": "",
            "著作者リスト": [], "管理状況": {},
            "error": "作品コードが空です",
        }

    detail_url = (
        JWID_BASE
        + f"main?trxID=F20101&WORKS_CD={works_cd}"
        + "&subSessionID=001&subSession=start"
    )
    return fetch_jwid_detail(detail_url)


# =====================================================================
# NexTone
# =====================================================================

_NEXTONE_BASE     = "https://search.nex-tone.co.jp/"
_NEXTONE_TERMS    = _NEXTONE_BASE
_NEXTONE_LIST     = _NEXTONE_BASE + "list"

# 初期化状態（利用規約同意後のフォームフィールドを保持）
_nextone_form_action: str = ""
_nextone_form_defaults: dict[str, str] = {}


def _nextone_init() -> bool:
    """NexTone 利用規約に同意し、検索フォームのデフォルト値を収集する。成功時 True。"""
    global _nextone_form_action, _nextone_form_defaults
    try:
        # 1. 利用規約ページを GET してアクション URL を取得
        _rate_limit("nex-tone.co.jp")
        r1 = _session.get(_NEXTONE_TERMS, timeout=TIMEOUT)
        action_m = re.search(r'action="(.*?)"', r1.text)
        if not action_m:
            return False
        agree_url = _NEXTONE_BASE + action_m.group(1).lstrip("./")

        # 2. 同意 POST → /list?1 へリダイレクト
        _rate_limit("nex-tone.co.jp")
        r2 = _session.post(agree_url, data={"id1_hf_0": "", "accept": ""},
                           timeout=TIMEOUT, allow_redirects=True)

        # list?1 ページのフォームフィールドを収集
        soup = BeautifulSoup(r2.text, "lxml")
        form = soup.find("form", id=re.compile(r"id\d+"))
        if not form:
            return False

        form_action_rel = form.get("action", "")
        _nextone_form_action = _NEXTONE_BASE + form_action_rel.lstrip("./")

        defaults: dict[str, str] = {}
        for inp in form.find_all(["input", "select"]):
            name = inp.get("name", "")
            if not name:
                continue
            typ = inp.get("type", inp.name)
            if typ == "select":
                sel_opt = inp.find("option", selected=True)
                defaults[name] = sel_opt["value"] if sel_opt else ""
            elif typ in ("radio", "checkbox"):
                if inp.get("checked"):
                    defaults[name] = inp.get("value", "on")
            else:
                defaults[name] = inp.get("value", "")

        _nextone_form_defaults = defaults
        return bool(_nextone_form_action and defaults)

    except Exception:
        return False


def search_nextone(title: str) -> dict:
    """
    NexTone でタイトルを検索して結果を返す。

    Returns: {
        "source": "NexTone",
        "search_url": str,
        "results": [{"管理番号":..., "作品名":..., "作曲者":..., "アーティスト":...}, ...],
        "error": str | None,
        "debug_html": str,
    }
    """
    out: dict = {
        "source": "NexTone",
        "search_url": _NEXTONE_LIST,
        "results": [],
        "error": None,
        "debug_html": "",
    }

    try:
        # 未初期化なら利用規約同意
        if not _nextone_form_action or not _nextone_form_defaults:
            if not _nextone_init():
                out["error"] = "NexTone 利用規約への同意または初期化に失敗しました。"
                return out

        # フォームデータを組み立てる
        form_data = dict(_nextone_form_defaults)   # デフォルト値をコピー
        form_data["freeWord"] = title
        form_data["search"]   = ""                 # 検索ボタン

        _rate_limit("nex-tone.co.jp")
        resp = _session.post(_nextone_form_action, data=form_data,
                             timeout=TIMEOUT, allow_redirects=True)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        html = resp.text
        out["debug_html"] = html[:3000]

        soup = BeautifulSoup(html, "lxml")
        results = _parse_nextone_table(soup)
        out["results"] = results

        # 次回検索用に form action を更新（Wicket はページバージョンが変わる）
        _update_nextone_form_action(soup)

        if not results:
            page_text = soup.get_text()
            if re.search(r"(0件|0 件|該当.*なし|not found)", page_text, re.I):
                out["error"] = "検索結果 0 件"
            else:
                out["error"] = (
                    "HTML パース失敗または検索結果なし。"
                    " debug_html を確認してください。"
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


def _update_nextone_form_action(soup: BeautifulSoup) -> None:
    """検索レスポンスから次回用の Wicket フォームアクション URL を更新する。"""
    global _nextone_form_action
    new_form = soup.find("form", id=re.compile(r"id\d+"))
    if new_form and new_form.get("action"):
        _nextone_form_action = _NEXTONE_BASE + new_form["action"].lstrip("./")


def _parse_nextone_table(soup: BeautifulSoup) -> list[dict]:
    """NexTone 検索結果 HTML（新形式）からデータ行を抽出する。

    対象テーブル: <table class="pieces-table">
    列構造:
      td.piece-table-col-piece-cd → span.result-piece-cd-value → 管理番号
      td.piece-table-col [2]      → span.result-value          → 作品名
      td.piece-table-col [3]      → span.result-value          → 著作者名（作曲者）
      td.piece-table-col [4]      → span.result-value          → アーティスト名
    """
    results = []

    tbl = soup.find("table", class_="pieces-table")
    if tbl is None:
        return results

    tbody = tbl.find("tbody")
    if tbody is None:
        return results

    for row in tbody.find_all("tr"):
        # 管理番号
        cd_td = row.find("td", class_="piece-table-col-piece-cd")
        mgmt_no = ""
        if cd_td:
            cd_span = cd_td.find("span", class_="result-piece-cd-value")
            mgmt_no = cd_span.get_text(strip=True) if cd_span else ""

        # 残りのデータ列（作品名・著作者名・アーティスト名）
        data_tds = row.find_all("td", class_="piece-table-col")
        def _col(idx: int) -> str:
            if idx >= len(data_tds):
                return ""
            span = data_tds[idx].find("span", class_="result-value")
            return span.get_text(strip=True) if span else ""

        item = {
            "管理番号":     mgmt_no,
            "作品名":       _col(0),
            "作曲者":       _col(1),
            "作詞者":       "",
            "アーティスト": _col(2),
        }
        if any(v for v in item.values()):
            results.append(item)

    return results


# =====================================================================
# 作曲者ヒントによる結果ランキング
# =====================================================================

def _composer_partial_match(result_composer: str, hint: str) -> bool:
    """
    結果の作曲者名がヒントと部分一致するか（姓単位・大文字小文字無視）。

    ヒントの各単語（2文字以上）が結果の作曲者名に含まれれば一致とみなす。
    例: hint="Alan Silvestri" → "alan" or "silvestri" が結果に含まれるか
    例: hint="加藤達也" → "加藤達也" が結果に含まれるか
    """
    if not hint or not result_composer:
        return False
    h = hint.strip().lower()
    r = result_composer.strip().lower()
    if h in r or r in h:
        return True
    return any(part in r for part in h.split() if len(part) > 1)


def _rank_by_composer(results: list[dict], composer_hint: str) -> list[dict]:
    """作曲者ヒントに一致する結果を先頭に並び替え、一致件数も返す"""
    if not composer_hint or not results:
        return results
    matched   = [r for r in results if _composer_partial_match(r.get("作曲者", ""), composer_hint)]
    unmatched = [r for r in results if not _composer_partial_match(r.get("作曲者", ""), composer_hint)]
    return matched + unmatched


# =====================================================================
# まとめて検索
# =====================================================================

def search_all(title: str, composer: str = "") -> dict:
    """
    J-WID と NexTone を連続して検索し、両方の結果を返す。

    Args:
        title:    検索する曲名
        composer: 作曲者ヒント（空でも可）。
                  指定した場合は結果を作曲者名で並び替え、
                  各サービスの結果辞書に ``composer_matched_count`` キーを付与する。

    Returns:
        {"jwid": {..., "composer_matched_count": int}, "nextone": {...}}
    """
    _composer = str(composer).strip()
    if _composer.lower() == "nan":
        _composer = ""

    # J-WID はタイトルだけで検索し、作曲者は後段のランキングで使う
    # （author を渡すと J-WID 側でも絞り込まれ、ヒント誤り時に全件ゼロになるリスクがある）
    jwid_result    = search_jwid(title)
    nextone_result = search_nextone(title)

    for result in [jwid_result, nextone_result]:
        items = result.get("results") or []
        if _composer and items:
            result["results"] = _rank_by_composer(items, _composer)
            result["composer_matched_count"] = sum(
                1 for r in result["results"]
                if _composer_partial_match(r.get("作曲者", ""), _composer)
            )
        else:
            result["composer_matched_count"] = 0

    return {"jwid": jwid_result, "nextone": nextone_result}
