"""
検索補助モジュール
J-WID / NexTone / Google 検索用の検索語を生成する
"""
import urllib.parse

import pandas as pd


# J-WID・NexTone の検索エントリポイント
# ※ 実際の検索 URL パラメータは各サイトの仕様変更により変わる可能性があります
JWID_BASE = "https://www2.jasrac.or.jp/eJwid/"
NEXTONE_SEARCH_BASE = "https://search.nex-tone.co.jp/search"


def _make_google_url(term: str) -> str:
    q = urllib.parse.quote(f"{term} 著作権 JASRAC NexTone")
    return f"https://www.google.com/search?q={q}"


def _make_nextone_url(term: str) -> str:
    q = urllib.parse.quote(term)
    return f"{NEXTONE_SEARCH_BASE}?keyword={q}"


def generate_search_terms(songs_df: pd.DataFrame) -> pd.DataFrame:
    """
    楽曲まとめ DataFrame から J-WID / NexTone / Google 検索用の
    検索語一覧 DataFrame を生成して返す。
    """
    records: list[dict] = []

    for _, row in songs_df.iterrows():
        # 検索語の優先候補を収集（重複排除しつつ順序維持）
        term_candidates: list[tuple[str, str]] = []  # (ラベル, 検索語)

        if row.get("WAV検出タイトル") and str(row["WAV検出タイトル"]).strip():
            term_candidates.append(("WAV検出タイトル", str(row["WAV検出タイトル"]).strip()))

        if row.get("正式タイトル") and str(row["正式タイトル"]).strip():
            term_candidates.append(
                ("正式タイトル", str(row["正式タイトル"]).strip()))

        if row.get("MP3検出タイトル") and str(row["MP3検出タイトル"]).strip():
            term_candidates.append(
                ("MP3検出タイトル", str(row["MP3検出タイトル"]).strip()))

        if row.get("曲名") and str(row["曲名"]).strip():
            term_candidates.append(("管理番号除去後曲名", str(row["曲名"]).strip()))

        if row.get("ライブラリ盤番号") and str(row["ライブラリ盤番号"]).strip():
            term_candidates.append(("ライブラリ盤番号", str(row["ライブラリ盤番号"]).strip()))

        if row.get("CD番号") and str(row["CD番号"]).strip():
            term_candidates.append(("CD番号", str(row["CD番号"]).strip()))

        # 重複を除去（同じ文字列が複数ラベルで来た場合）
        seen: set[str] = set()
        for label, term in term_candidates:
            if term in seen:
                continue
            seen.add(term)

            records.append(
                {
                    "No": row.get("No", ""),
                    "イベント名": row.get("イベント名", ""),
                    "検索語ラベル": label,
                    "検索語": term,
                    "J-WID": JWID_BASE,            # J-WID はトップから手動検索
                    "NexTone検索URL": _make_nextone_url(term),
                    "Google検索URL": _make_google_url(term),
                }
            )

        # 候補がゼロの場合はイベント名をそのまま使う
        if not seen:
            event_name = str(row.get("イベント名", "")).strip()
            if event_name:
                records.append(
                    {
                        "No": row.get("No", ""),
                        "イベント名": event_name,
                        "検索語ラベル": "イベント名（候補なし）",
                        "検索語": event_name,
                        "J-WID": JWID_BASE,
                        "NexTone検索URL": _make_nextone_url(event_name),
                        "Google検索URL": _make_google_url(event_name),
                    }
                )

    return pd.DataFrame(records)
