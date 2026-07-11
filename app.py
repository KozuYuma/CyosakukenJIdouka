"""
著作権調査支援ツール - メインアプリ
NUENDO Cue CSV × WAV 一覧照合 → 権利情報管理 → Excel 出力

起動方法: run.bat をダブルクリック
        または: streamlit run app.py
"""
import urllib.parse

import pandas as pd
import streamlit as st

from modules.csv_reader import (
    normalize_cue_columns,
    read_csv_auto,
    validate_cue_csv,
    validate_wav_csv,
)
from modules.excel_exporter import export_to_excel
from modules.matcher import build_song_list
from modules.musicbrainz import _hms_to_sec, mb_search_url, search_recording
from modules.pipeline import run_pipeline
from modules.scraper import search_all
from modules.search_helper import JWID_BASE, generate_search_terms
from modules.spotify import is_available as spotify_available, spotify_search_url

# =====================================================================
# アプリ設定
# =====================================================================
st.set_page_config(
    page_title="著作権調査支援ツール",
    page_icon="🎵",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# 確認ステータスの選択肢（全画面共通）
CONFIRM_STATUS_OPTIONS = [
    "未調査",
    "候補あり",
    "確定",
    "要確認",
    "J-WID要確認",
    "NexTone要確認",
    "ライブラリ元確認",
    "Audiostock確認",
    "MP3補助確認",
]

# 編集可能列（data_editor で直接入力できる列）
EDITABLE_COLS = [
    "作曲者",
    "作詞者",
    "アーティスト",
    "CD番号",
    "JASRAC作品コード",
    "NexTone管理番号",
    "確認ステータス",
    "メモ",
]


# =====================================================================
# セッション状態初期化
# =====================================================================
def _init_session() -> None:
    defaults: dict = {
        "cue_df": None,
        "wav_df": None,
        "mp3_df": None,
        "mp3_is_finder": False,   # mp3_df が nuendo_mp3_finder 出力形式かどうか
        "songs_df": None,
        "events_df": None,
        "search_df": None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


_init_session()

# =====================================================================
# nuendo_mp3_finder CSV ヘルパー
# =====================================================================
_MP3FINDER_KEY_COLS = {"イベント名", "ファイル名"}
_MP3FINDER_MATCH_COLS = {"一致種別", "照合種別"}
_MP3FINDER_ID3_COLS = ["タイトル(ID3)", "アーティスト(ID3)", "アルバム(ID3)", "作曲者(ID3)"]


def _is_mp3finder_csv(df: pd.DataFrame) -> bool:
    """nuendo_mp3_finder.py の出力 CSV かどうかを判定する"""
    cols = set(df.columns)
    return _MP3FINDER_KEY_COLS.issubset(cols) and bool(cols & _MP3FINDER_MATCH_COLS)


def _import_mp3finder_id3(
    mp3finder_df: pd.DataFrame,
    songs_df: pd.DataFrame,
) -> tuple[pd.DataFrame, int]:
    """
    nuendo_mp3_finder CSV の ID3 タグ情報を songs_df に補完する。
    完全一致 > タイトル一致 > 部分一致 の優先順で最良行を選び、
    空フィールドのみ上書きする（既存データは保持）。
    """
    match_col = "一致種別" if "一致種別" in mp3finder_df.columns else "照合種別"
    PRIORITY = {"完全一致": 3, "タイトル一致": 2, "部分一致": 1, "該当なし": 0}
    id3_cols_present = [c for c in _MP3FINDER_ID3_COLS if c in mp3finder_df.columns]

    updated = 0
    for idx in songs_df.index:
        event = str(songs_df.at[idx, "イベント名"]).strip()
        if not event or event.lower() == "nan":
            continue

        hits = mp3finder_df[mp3finder_df["イベント名"] == event].copy()
        if hits.empty:
            continue

        hits["_pri"] = hits[match_col].map(lambda x: PRIORITY.get(str(x), 0))
        hits = hits[hits["_pri"] > 0]
        if hits.empty:
            continue

        # ID3 フィールドが多い行を優先
        hits["_id3"] = hits[id3_cols_present].apply(
            lambda row: sum(1 for v in row if str(v).strip() not in ("", "nan")),
            axis=1,
        )
        best = hits.sort_values(["_pri", "_id3"], ascending=[False, False]).iloc[0]

        def _fill(src: str, dst: str) -> bool:
            if src not in best.index or dst not in songs_df.columns:
                return False
            v = str(best[src]).strip()
            if not v or v.lower() == "nan":
                return False
            cur = str(songs_df.at[idx, dst]).strip()
            if cur and cur.lower() != "nan":
                return False  # 既存データは上書きしない
            songs_df.at[idx, dst] = v
            return True

        changed = any([
            _fill("アーティスト(ID3)", "アーティスト"),
            _fill("作曲者(ID3)",       "作曲者"),
            _fill("アルバム(ID3)",     "CD番号"),
            _fill("ファイル名",        "MP3一致ファイル名"),
            _fill("再生時間",          "MP3フル尺"),
        ])
        if changed:
            updated += 1

    return songs_df, updated

# =====================================================================
# ヘッダー
# =====================================================================
st.title("🎵 著作権調査支援ツール")
st.caption(
    "NUENDO Cue CSV × WAV 一覧照合 → 権利情報手入力 → Excel 出力 | "
    "音響効果・選曲業務 Cue Sheet 作成補助"
)

tabs = st.tabs(
    [
        "📁 ファイル読み込み",
        "🎵 楽曲まとめ",
        "📋 イベント一覧",
        "🔍 検索補助",
        "📊 Excel 出力",
    ]
)


# =====================================================================
# タブ 1: ファイル読み込み
# =====================================================================
with tabs[0]:
    st.header("ファイル読み込み")

    col_left, col_right = st.columns(2)

    # ---- Cue CSV ----
    with col_left:
        st.subheader("① NUENDO Cue CSV（必須）")
        st.caption("NUENDO から書き出した Cue CSV をアップロードしてください。")
        cue_file = st.file_uploader(
            "Cue CSV を選択", type=["csv"], key="upload_cue"
        )
        if cue_file:
            try:
                df, enc = read_csv_auto(cue_file)
                df = normalize_cue_columns(df)
                missing = validate_cue_csv(df)
                if missing:
                    st.error(
                        f"❌ 必須列が不足しています: {missing}\n\n"
                        f"検出された列: {list(df.columns)}"
                    )
                else:
                    st.session_state.cue_df = df
                    st.success(f"✅ 読み込み完了（{enc}）: {len(df)} 件")
                    with st.expander("プレビュー（先頭 5 行）"):
                        st.dataframe(df.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"❌ 読み込みエラー: {e}")

    # ---- WAV（フォルダスキャン または CSV アップロード） ----
    with col_right:
        st.subheader("② WAV ファイル一覧（必須）")

        wav_tab_scan, wav_tab_csv = st.tabs(["📂 フォルダをスキャン", "📄 CSV をアップロード"])

        with wav_tab_scan:
            st.caption("Audio フォルダのパスを貼り付けてスキャンします。PowerShell 不要です。")
            wav_folder = st.text_input(
                "WAV フォルダパス",
                placeholder=r"例: H:\プロジェクト名\Audio",
                key="wav_folder_path",
            )
            wav_recursive = st.checkbox("サブフォルダも含める", value=True, key="wav_recursive")
            if st.button("🔍 WAV をスキャン", key="scan_wav", use_container_width=True):
                if not wav_folder:
                    st.warning("フォルダパスを入力してください。")
                else:
                    with st.spinner("スキャン中..."):
                        try:
                            from modules.folder_scanner import scan_wav_folder
                            df = scan_wav_folder(wav_folder, wav_recursive)
                            if len(df) == 0:
                                st.warning("WAV ファイルが見つかりませんでした。パスを確認してください。")
                            else:
                                st.session_state.wav_df = df
                                st.success(f"✅ {len(df)} 件の WAV ファイルを検出しました")
                                with st.expander("プレビュー（先頭 5 行）"):
                                    st.dataframe(df.head(5), use_container_width=True)
                        except Exception as e:
                            st.error(f"❌ エラー: {e}")

        with wav_tab_csv:
            st.caption("Get-WavList.ps1 で作成した CSV をアップロードします。")
            wav_file = st.file_uploader(
                "WAV 一覧 CSV を選択", type=["csv"], key="upload_wav"
            )
            if wav_file:
                try:
                    df, enc = read_csv_auto(wav_file)
                    missing = validate_wav_csv(df)
                    if missing:
                        st.error(
                            f"❌ 必須列が不足しています: {missing}\n\n"
                            f"検出された列: {list(df.columns)}"
                        )
                    else:
                        st.session_state.wav_df = df
                        st.success(f"✅ 読み込み完了（{enc}）: {len(df)} 件")
                        with st.expander("プレビュー（先頭 5 行）"):
                            st.dataframe(df.head(5), use_container_width=True)
                except Exception as e:
                    st.error(f"❌ 読み込みエラー: {e}")

        if st.session_state.wav_df is not None:
            st.info(f"現在の WAV 一覧: {len(st.session_state.wav_df)} 件読み込み済み")

    st.divider()

    # ---- ③ MusicBrainz 自動調査 ----
    st.subheader("③ MusicBrainz 自動調査（照合実行後に有効）")
    st.caption(
        "照合実行後、**曲名** × WAV尺で MusicBrainz を検索し ISRC・アーティストを取得。"
        " ④ MP3 の作曲者情報が入力済みの場合、J-WID / NexTone の結果を作曲者で絞り込みます。"
    )

    if st.session_state.songs_df is None:
        st.info("先に ① Cue CSV と ② WAV ファイル一覧を読み込み、照合実行してください。")
    else:
        _mb_songs = st.session_state.songs_df
        _mb_total = len(_mb_songs)
        _mb_unresolved = _mb_songs[
            _mb_songs["確認ステータス"].isin(["未調査", "MP3補助確認"])
        ] if "確認ステータス" in _mb_songs.columns else _mb_songs

        mb_bulk_col1, mb_bulk_col2, mb_bulk_col3 = st.columns(3)
        with mb_bulk_col1:
            mb_bulk_target = st.radio(
                "検索対象",
                ["未調査のみ", "全曲"],
                key="mb_bulk_target",
                horizontal=True,
            )
        with mb_bulk_col2:
            mb_bulk_tol = st.slider("尺の許容誤差(秒)", 5, 60, 15, 5, key="mb_bulk_tol")
        with mb_bulk_col3:
            mb_bulk_thresh = st.number_input("MBスコア閾値", 0, 100, 80, 5, key="mb_bulk_thresh")

        mb_bulk_opt1, mb_bulk_opt2 = st.columns(2)
        with mb_bulk_opt1:
            _claude_avail = False
            try:
                from modules.claude_lookup import is_available as _claude_check
                _claude_avail = _claude_check()
            except Exception:
                pass
            mb_bulk_use_claude = st.checkbox(
                "Claude API も使う",
                value=False,
                key="mb_bulk_use_claude",
                disabled=not _claude_avail,
                help="ANTHROPIC_API_KEY が設定済みのとき有効。CD名・作曲者情報をClaudeに問い合わせます。" if _claude_avail else "ANTHROPIC_API_KEY が未設定です（.env ファイルに設定してください）",
            )
        with mb_bulk_opt2:
            _sp_avail = spotify_available()
            mb_bulk_use_spotify = st.checkbox(
                "Spotify API も使う",
                value=False,
                key="mb_bulk_use_spotify",
                disabled=not _sp_avail,
                help="Spotifyでアーティスト・アルバム・ISRCを取得します。" if _sp_avail else "SPOTIFY_CLIENT_ID / SPOTIFY_CLIENT_SECRET が未設定です（.env ファイルに設定してください）",
            )

        _mb_targets = _mb_unresolved if mb_bulk_target == "未調査のみ" else _mb_songs
        _mb_target_count = len(_mb_targets)

        if st.button(
            f"🔄 MusicBrainz → J-WID / NexTone 一括実行（{_mb_target_count} 件）",
            key="mb_bulk_run",
            type="primary",
            use_container_width=True,
            disabled=_mb_target_count == 0,
        ):
            _mb_bar = st.progress(0)
            _mb_status = st.empty()
            _mb_stats = {"MB命中": 0, "JWID命中": 0, "NexTone命中": 0, "エラー": 0}

            for _mb_i, (_mb_idx, _mb_row) in enumerate(_mb_targets.iterrows()):
                _mb_name = str(_mb_row.get("イベント名", ""))
                _mb_status.text(f"[{_mb_i+1}/{_mb_target_count}] {_mb_name[:40]}")
                _mb_bar.progress((_mb_i + 1) / _mb_target_count)
                try:
                    _pip = run_pipeline(
                        event_name=_mb_name,
                        wav_full_duration=str(_mb_row.get("WAVフル尺", "")),
                        wav_detected_title=str(_mb_row.get("WAV検出タイトル", "")),
                        song_title=str(_mb_row.get("曲名", "")),
                        composer=str(_mb_row.get("作曲者", "")),
                        tolerance_sec=float(mb_bulk_tol),
                        mb_score_threshold=int(mb_bulk_thresh),
                        use_claude=bool(mb_bulk_use_claude),
                        use_spotify=bool(mb_bulk_use_spotify),
                    )
                    _mb_best = _pip.get("mb_best")
                    if _mb_best and _mb_best.get("score", 0) >= int(mb_bulk_thresh):
                        _mb_stats["MB命中"] += 1
                        if _mb_best.get("artist"):
                            st.session_state.songs_df.at[_mb_idx, "アーティスト"] = _mb_best["artist"]

                    # Spotify 結果で補完（MusicBrainz 未取得 or 低スコアのとき）
                    _sp_best = _pip.get("sp_best")
                    if _sp_best and not _sp_best.get("error"):
                        if not st.session_state.songs_df.at[_mb_idx, "アーティスト"] and _sp_best.get("artist"):
                            st.session_state.songs_df.at[_mb_idx, "アーティスト"] = _sp_best["artist"]
                        if _sp_best.get("album") and "CD番号" in st.session_state.songs_df.columns:
                            if not st.session_state.songs_df.at[_mb_idx, "CD番号"]:
                                st.session_state.songs_df.at[_mb_idx, "CD番号"] = _sp_best["album"]

                    # Claude API 結果で補完（J-WID / MB / Spotify より後に上書きしない）
                    _cl = _pip.get("claude_result") or {}
                    if not _cl.get("error") and _cl.get("confidence") in ("high", "medium"):
                        if _cl.get("artist") and not st.session_state.songs_df.at[_mb_idx, "アーティスト"]:
                            st.session_state.songs_df.at[_mb_idx, "アーティスト"] = _cl["artist"]
                        if _cl.get("composer") and not st.session_state.songs_df.at[_mb_idx, "作曲者"]:
                            st.session_state.songs_df.at[_mb_idx, "作曲者"] = _cl["composer"]
                        if _cl.get("cd_name") and "CD番号" in st.session_state.songs_df.columns:
                            if not st.session_state.songs_df.at[_mb_idx, "CD番号"]:
                                st.session_state.songs_df.at[_mb_idx, "CD番号"] = _cl["cd_name"]

                    _jw = _pip["jwid_results"]
                    _jw_r = _jw.get("results") or []
                    _jw_comp_n = _jw.get("composer_matched_count", 0)
                    # 1件 OR 作曲者一致が1件のみ → 自動適用（results は一致順にソート済み）
                    if _jw_r and (len(_jw_r) == 1 or _jw_comp_n == 1):
                        _mb_stats["JWID命中"] += 1
                        _r = _jw_r[0]
                        if _r.get("作曲者"):
                            st.session_state.songs_df.at[_mb_idx, "作曲者"] = _r["作曲者"]
                        if _r.get("作詞者"):
                            st.session_state.songs_df.at[_mb_idx, "作詞者"] = _r["作詞者"]
                        if _r.get("作品コード"):
                            st.session_state.songs_df.at[_mb_idx, "JASRAC作品コード"] = _r["作品コード"]
                        st.session_state.songs_df.at[_mb_idx, "確認ステータス"] = "候補あり"
                    elif _jw_r:  # 複数候補（絞り込めない）
                        st.session_state.songs_df.at[_mb_idx, "確認ステータス"] = "候補あり"

                    _nt = _pip["nextone_results"]
                    _nt_r = _nt.get("results") or []
                    _nt_comp_n = _nt.get("composer_matched_count", 0)
                    if _nt_r and (len(_nt_r) == 1 or _nt_comp_n == 1):
                        _mb_stats["NexTone命中"] += 1
                        _rn = _nt_r[0]
                        if not st.session_state.songs_df.at[_mb_idx, "作曲者"] and _rn.get("作曲者"):
                            st.session_state.songs_df.at[_mb_idx, "作曲者"] = _rn["作曲者"]
                        if _rn.get("管理番号"):
                            st.session_state.songs_df.at[_mb_idx, "NexTone管理番号"] = _rn["管理番号"]
                        if st.session_state.songs_df.at[_mb_idx, "確認ステータス"] == "未調査":
                            st.session_state.songs_df.at[_mb_idx, "確認ステータス"] = "候補あり"
                    elif _nt_r:
                        if st.session_state.songs_df.at[_mb_idx, "確認ステータス"] == "未調査":
                            st.session_state.songs_df.at[_mb_idx, "確認ステータス"] = "候補あり"
                except Exception:
                    _mb_stats["エラー"] += 1

            _mb_status.empty()
            _mb_bar.empty()
            st.success(
                f"完了！　MB命中: {_mb_stats['MB命中']}件　"
                f"J-WID: {_mb_stats['JWID命中']}件　"
                f"NexTone: {_mb_stats['NexTone命中']}件　"
                f"エラー: {_mb_stats['エラー']}件"
            )
            st.info("「楽曲まとめ」タブで結果を確認・修正してください。")

    st.divider()

    col_left2, col_right2 = st.columns(2)

    # ---- MP3（フォルダスキャン または CSV アップロード） ----
    with col_left2:
        st.subheader("④ MP3 ファイル一覧（任意・補助）")
        st.caption("WAV で照合できなかった場合の補助データです。")

        mp3_tab_scan, mp3_tab_csv = st.tabs(["📂 フォルダをスキャン", "📄 CSV をアップロード"])

        with mp3_tab_scan:
            mp3_folder = st.text_input(
                "MP3 フォルダパス",
                placeholder=r"例: H:\MP3ライブラリ",
                key="mp3_folder_path",
            )
            mp3_recursive = st.checkbox("サブフォルダも含める", value=True, key="mp3_recursive")
            if st.button("🔍 MP3 をスキャン", key="scan_mp3", use_container_width=True):
                if not mp3_folder:
                    st.warning("フォルダパスを入力してください。")
                else:
                    with st.spinner("スキャン中..."):
                        try:
                            from modules.folder_scanner import scan_mp3_folder
                            df = scan_mp3_folder(mp3_folder, mp3_recursive)
                            if len(df) == 0:
                                st.warning("MP3 ファイルが見つかりませんでした。パスを確認してください。")
                            else:
                                st.session_state.mp3_df = df
                                st.success(f"✅ {len(df)} 件の MP3 ファイルを検出しました")
                                with st.expander("プレビュー（先頭 5 行）"):
                                    st.dataframe(df.head(5), use_container_width=True)
                        except Exception as e:
                            st.error(f"❌ エラー: {e}")

        with mp3_tab_csv:
            mp3_file = st.file_uploader(
                "MP3 一覧 CSV を選択（任意）", type=["csv"], key="upload_mp3"
            )
            if mp3_file:
                try:
                    df, enc = read_csv_auto(mp3_file)
                    st.session_state.mp3_df = df
                    if _is_mp3finder_csv(df):
                        st.session_state.mp3_is_finder = True
                        n_events = df["イベント名"].nunique()
                        st.success(
                            f"✅ nuendo_mp3_finder 出力を検出（{enc}）："
                            f"{len(df)} 行 / {n_events} イベント"
                        )
                        st.caption("イベント名・照合種別・ID3タグ情報が含まれています。")
                    else:
                        st.session_state.mp3_is_finder = False
                        st.success(f"✅ 読み込み完了（{enc}）: {len(df)} 件")
                    with st.expander("プレビュー（先頭 5 行）"):
                        st.dataframe(df.head(5), use_container_width=True)
                except Exception as e:
                    st.error(f"❌ 読み込みエラー: {e}")

            # nuendo_mp3_finder 形式ならば ID3 取り込みセクションを表示
            if st.session_state.get("mp3_is_finder") and st.session_state.mp3_df is not None:
                _mf_df = st.session_state.mp3_df
                _mf_n_evt = _mf_df["イベント名"].nunique()
                st.divider()
                st.markdown(f"**📋 nuendo_mp3_finder 出力（{_mf_n_evt} イベント）が読み込まれています**")
                if st.session_state.songs_df is not None:
                    if st.button(
                        "💿 ID3タグ情報を楽曲まとめに取り込む",
                        key="import_mp3finder_csv",
                        type="primary",
                        use_container_width=True,
                    ):
                        _new_songs, _n_upd = _import_mp3finder_id3(
                            _mf_df, st.session_state.songs_df.copy()
                        )
                        st.session_state.songs_df = _new_songs
                        st.success(f"✅ {_n_upd} 件のイベントに ID3 情報を補完しました。")
                        st.caption("「楽曲まとめ」タブで結果を確認してください。")
                else:
                    st.info("「照合実行」後に取り込みボタンが表示されます。")

    # ---- 既存 Excel（任意・将来実装） ----
    with col_right2:
        st.subheader("⑤ 既存調査 Excel（任意）")
        st.caption("前回の調査 Excel を読み込み、作家情報・作品コードを引き継ぎます。")
        excel_file = st.file_uploader(
            "既存調査 Excel を選択（任意）", type=["xlsx"], key="upload_excel"
        )
        if excel_file:
            st.info("⚠️ 既存 Excel 読み込み機能は将来バージョンで実装予定です。")

    st.divider()

    # ---- 読み込み状況サマリー ----
    st.subheader("読み込み状況")
    c1, c2, c3, c4 = st.columns(4)
    cue_n = len(st.session_state.cue_df) if st.session_state.cue_df is not None else 0
    wav_n = len(st.session_state.wav_df) if st.session_state.wav_df is not None else 0
    mp3_n = len(st.session_state.mp3_df) if st.session_state.mp3_df is not None else 0
    song_n = len(st.session_state.songs_df) if st.session_state.songs_df is not None else 0

    c1.metric("Cue イベント", f"{cue_n} 件")
    c2.metric("WAV ファイル", f"{wav_n} 件")
    c3.metric("MP3 ファイル", f"{mp3_n} 件")
    c4.metric("照合済み楽曲", f"{song_n} 件")

    # ---- 照合実行ボタン ----
    st.divider()
    can_match = (
        st.session_state.cue_df is not None
        and st.session_state.wav_df is not None
    )

    if not can_match:
        st.info(
            "Cue CSV と WAV 一覧 CSV を両方読み込んでから「照合実行」ボタンを押してください。"
        )

    if st.button(
        "🔄 照合実行",
        disabled=not can_match,
        type="primary",
        use_container_width=True,
    ):
        with st.spinner("照合処理中..."):
            songs_df, events_df = build_song_list(
                st.session_state.cue_df,
                st.session_state.wav_df,
                st.session_state.mp3_df,
            )
            st.session_state.songs_df = songs_df
            st.session_state.events_df = events_df
            st.session_state.search_df = generate_search_terms(songs_df)

        st.success(
            f"✅ 照合完了！楽曲まとめ: {len(songs_df)} 件 ／ "
            f"イベント: {len(events_df)} 件"
        )

        # ステータス別件数を表示
        if "WAV照合ステータス" in songs_df.columns:
            status_counts = (
                songs_df["WAV照合ステータス"]
                .value_counts()
                .reset_index()
            )
            status_counts.columns = ["照合ステータス", "件数"]
            st.dataframe(status_counts, use_container_width=False, hide_index=True)

        st.info("「楽曲まとめ」タブで内容を確認・編集してください。")
        if st.session_state.get("mp3_is_finder") and st.session_state.mp3_df is not None:
            st.info(
                "💡 ④ MP3 CSV タブの「ID3タグ情報を楽曲まとめに取り込む」ボタンで"
                " MP3 のメタデータ（アーティスト・作曲者・アルバム）を補完できます。"
            )


# =====================================================================
# タブ 2: 楽曲まとめ
# =====================================================================
with tabs[1]:
    st.header("楽曲まとめ")

    if st.session_state.songs_df is None:
        st.info("「ファイル読み込み」タブで照合を実行してください。")
    else:
        songs_df: pd.DataFrame = st.session_state.songs_df

        # ---- フィルター ----
        fc1, fc2 = st.columns([2, 2])
        with fc1:
            status_opts = sorted(songs_df["確認ステータス"].dropna().unique().tolist())
            status_filter = st.multiselect(
                "確認ステータスで絞り込み",
                options=status_opts,
                default=status_opts,
                key="songs_status_filter",
            )
        with fc2:
            type_opts = sorted(songs_df["管理番号種別"].dropna().unique().tolist())
            type_filter = st.multiselect(
                "管理番号種別で絞り込み",
                options=["すべて"] + type_opts,
                default=["すべて"],
                key="songs_type_filter",
            )

        # フィルター適用
        mask = songs_df["確認ステータス"].isin(status_filter)
        if "すべて" not in type_filter and type_filter:
            mask &= songs_df["管理番号種別"].isin(type_filter)
        filtered_df = songs_df[mask]

        st.caption(f"表示: {len(filtered_df)} 件 ／ 全 {len(songs_df)} 件")

        # ---- 一括自動検索 ----
        with st.expander("🔍 一括自動検索（J-WID / NexTone）", expanded=False):
            bulk_target = st.radio(
                "検索対象",
                ["未調査のみ", "全曲"],
                horizontal=True,
                key="bulk_search_target",
            )
            target_mask = (
                st.session_state.songs_df["確認ステータス"].isin(["未調査", "MP3補助確認"])
                if bulk_target == "未調査のみ"
                else pd.Series([True] * len(st.session_state.songs_df),
                               index=st.session_state.songs_df.index)
            )
            target_count = int(target_mask.sum())
            est_min = max(1, round(target_count * 4 / 60))

            _has_composer_info = (
                "作曲者" in st.session_state.songs_df.columns
                and st.session_state.songs_df["作曲者"].astype(str).str.strip().ne("").any()
            )
            st.info(
                f"対象: **{target_count} 件** ／ 推定所要時間: 約 {est_min} 分  \n"
                "曲名で J-WID / NexTone を検索します（管理番号は使いません）。  \n"
                + ("✅ 作曲者情報あり — 検索結果を作曲者で絞り込みます。  \n" if _has_composer_info else "")
                + "1件 or 作曲者一致が1件 → 自動入力（ステータス: 候補あり）  \n"
                  "複数ヒット → 自動入力せずマークのみ（Tab4 で手動確認）"
            )

            if st.button(
                f"🔍 一括自動検索を実行（{target_count} 件）",
                key="bulk_search_btn",
                type="primary",
                disabled=target_count == 0,
            ):
                target_indices = st.session_state.songs_df[target_mask].index.tolist()
                total = len(target_indices)
                progress_bar = st.progress(0)
                status_ph = st.empty()
                stats: dict[str, int] = {
                    "自動入力": 0, "複数候補": 0, "ヒットなし": 0, "エラー": 0
                }

                for i, idx in enumerate(target_indices):
                    row = st.session_state.songs_df.loc[idx]
                    event_name = str(row.get("イベント名", ""))

                    # 検索語: 曲名 → WAV検出タイトル → イベント名（管理番号は使わない）
                    search_term = (
                        str(row.get("曲名", "")).strip()
                        or str(row.get("WAV検出タイトル", "")).strip()
                        or event_name
                    )
                    if search_term.lower() == "nan":
                        search_term = event_name

                    # 作曲者ヒント（MP3 ID3 取り込み済みの場合に有効）
                    composer_hint = str(row.get("作曲者", "")).strip()
                    if composer_hint.lower() == "nan":
                        composer_hint = ""

                    status_ph.caption(f"({i + 1}/{total}) 検索中: {search_term[:50]}")
                    progress_bar.progress((i + 1) / total)

                    try:
                        result = search_all(search_term, composer=composer_hint)
                    except Exception:
                        stats["エラー"] += 1
                        continue

                    jwid_r      = result.get("jwid", {}).get("results", []) or []
                    jwid_comp_n = result.get("jwid", {}).get("composer_matched_count", 0)
                    nt_r        = result.get("nextone", {}).get("results", []) or []
                    nt_comp_n   = result.get("nextone", {}).get("composer_matched_count", 0)

                    # 自動適用条件: 1件のみ OR 作曲者一致が1件（results は一致順にソート済み）
                    def _auto_apply(results, comp_n):
                        return bool(results) and (len(results) == 1 or comp_n == 1)

                    updates: dict = {}

                    if _auto_apply(jwid_r, jwid_comp_n):
                        r = jwid_r[0]
                        if r.get("作曲者"):     updates["作曲者"] = r["作曲者"]
                        if r.get("作詞者"):     updates["作詞者"] = r["作詞者"]
                        if r.get("作品コード"): updates["JASRAC作品コード"] = r["作品コード"]

                    if _auto_apply(nt_r, nt_comp_n):
                        r = nt_r[0]
                        if r.get("作曲者") and not updates.get("作曲者"):
                            updates["作曲者"] = r["作曲者"]
                        if r.get("管理番号"):
                            updates["NexTone管理番号"] = r["管理番号"]

                    if updates:
                        updates["確認ステータス"] = "候補あり"
                        for col, val in updates.items():
                            if col in st.session_state.songs_df.columns:
                                st.session_state.songs_df.at[idx, col] = val
                        stats["自動入力"] += 1
                    elif jwid_r or nt_r:
                        st.session_state.songs_df.at[idx, "確認ステータス"] = "候補あり"
                        stats["複数候補"] += 1
                    else:
                        stats["ヒットなし"] += 1

                progress_bar.empty()
                status_ph.empty()

                result_msg = (
                    f"✅ 完了: 自動入力 {stats['自動入力']} 件 ／ "
                    f"複数候補 {stats['複数候補']} 件 ／ "
                    f"ヒットなし {stats['ヒットなし']} 件"
                )
                if stats["エラー"]:
                    result_msg += f" ／ エラー {stats['エラー']} 件"
                st.success(result_msg)
                st.rerun()

        # ---- 編集不可列の設定 ----
        disabled_cols = [c for c in songs_df.columns if c not in EDITABLE_COLS]

        # ---- data_editor ----
        edited_df = st.data_editor(
            filtered_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            disabled=disabled_cols,
            column_config={
                "No": st.column_config.NumberColumn("No", width="small"),
                "管理番号種別": st.column_config.TextColumn("管理番号種別", width="medium"),
                "元管理番号": st.column_config.TextColumn("元管理番号", width="medium"),
                "ライブラリ盤番号": st.column_config.TextColumn("ライブラリ盤番号", width="medium"),
                "トラック番号": st.column_config.TextColumn("トラック番号", width="small"),
                "曲名": st.column_config.TextColumn("曲名", width="large"),
                "イベント名": st.column_config.TextColumn("イベント名", width="large"),
                "WAV照合ステータス": st.column_config.TextColumn("WAV照合ステータス", width="medium"),
                "WAVフル尺": st.column_config.TextColumn("WAVフル尺", width="small"),
                "確認ステータス": st.column_config.SelectboxColumn(
                    "確認ステータス",
                    options=CONFIRM_STATUS_OPTIONS,
                    width="medium",
                ),
                "JASRAC作品コード": st.column_config.TextColumn("JASRAC作品コード", width="medium"),
                "NexTone管理番号": st.column_config.TextColumn("NexTone管理番号", width="medium"),
                "CD番号": st.column_config.TextColumn("CD番号", width="medium"),
                "メモ": st.column_config.TextColumn("メモ", width="large"),
            },
            key="songs_editor",
        )

        # 編集内容をセッション状態に反映（フィルター後の行のみ更新）
        if edited_df is not None:
            for col in EDITABLE_COLS:
                if col in edited_df.columns:
                    st.session_state.songs_df.loc[filtered_df.index, col] = (
                        edited_df[col].values
                    )


# =====================================================================
# タブ 3: イベント一覧
# =====================================================================
with tabs[2]:
    st.header("イベント一覧（NUENDO イベント単位）")

    if st.session_state.events_df is None:
        st.info("「ファイル読み込み」タブで照合を実行してください。")
    else:
        events_df: pd.DataFrame = st.session_state.events_df

        status_opts_e = sorted(events_df["照合ステータス"].dropna().unique().tolist())
        ev_filter = st.multiselect(
            "照合ステータスで絞り込み",
            options=status_opts_e,
            default=status_opts_e,
            key="event_status_filter",
        )

        filtered_ev = events_df[events_df["照合ステータス"].isin(ev_filter)]
        st.caption(f"表示: {len(filtered_ev)} 件 ／ 全 {len(events_df)} 件")
        st.dataframe(filtered_ev, use_container_width=True, hide_index=True)


# =====================================================================
# タブ 4: 検索補助
# =====================================================================
with tabs[3]:
    st.header("検索補助（J-WID / NexTone 自動調査）")

    if st.session_state.songs_df is None:
        st.info("「ファイル読み込み」タブで照合を実行してください。")
    else:
        songs_df = st.session_state.songs_df

        # ---- 楽曲選択 ----
        song_labels = (songs_df["No"].astype(str) + ". " + songs_df["イベント名"]).tolist()
        selected_label = st.selectbox("調査する楽曲を選択", options=song_labels, key="search_song_select")

        if selected_label:
            selected_no = int(selected_label.split(".")[0])
            row_idx = songs_df[songs_df["No"] == selected_no].index[0]
            row = songs_df.loc[row_idx]

            # 検索語を収集（優先度順）
            term_candidates: list[tuple[str, str]] = []
            for field, label in [
                ("WAV検出タイトル", "WAV検出タイトル"),
                ("曲名",           "管理番号除去後曲名"),
                ("ライブラリ盤番号", "ライブラリ盤番号"),
                ("CD番号",         "CD番号"),
            ]:
                val = str(row.get(field, "")).strip()
                if val and val.lower() != "nan":
                    term_candidates.append((label, val))
            if not term_candidates:
                term_candidates.append(("イベント名", str(row.get("イベント名", "")).strip()))

            main_term = term_candidates[0][1]
            encoded = urllib.parse.quote(main_term)

            st.subheader(f"No.{selected_no} ／ {row.get('イベント名', '')}")

            # ---- 検索語と手動リンク ----
            col_terms, col_links = st.columns([3, 2])
            with col_terms:
                st.markdown("**検索語候補**（クリックして選択＆コピー）")
                for label, term in term_candidates:
                    st.text_input(label, value=term, key=f"term_{selected_no}_{label}")

            with col_links:
                st.markdown("**手動検索リンク**")
                st.link_button("🔍 J-WID トップ", JWID_BASE, use_container_width=True)
                st.link_button(
                    "🔍 NexTone で検索",
                    f"https://search.nex-tone.co.jp/search?keyword={encoded}",
                    use_container_width=True,
                )
                st.link_button(
                    "🔍 Google で検索",
                    f"https://www.google.com/search?q={encoded}+著作権+JASRAC+NexTone",
                    use_container_width=True,
                )
                st.link_button(
                    "🎵 MusicBrainz で検索",
                    mb_search_url(main_term),
                    use_container_width=True,
                )
                st.link_button(
                    "🎧 Spotify で検索",
                    spotify_search_url(main_term),
                    use_container_width=True,
                )

            st.divider()

            # ---- 全自動パイプライン ----
            st.markdown("#### 🔄 全自動調査パイプライン（MusicBrainz → J-WID / NexTone）")
            st.caption("① 曲名抽出 → ② MusicBrainz で正式タイトル & ISRC 取得 → ③ J-WID / NexTone で JASRAC コード取得")

            pip_col1, pip_col2 = st.columns([2, 1])
            with pip_col1:
                pip_tolerance = st.slider(
                    "尺の許容誤差（秒）— MusicBrainz / Spotify 共通",
                    min_value=5, max_value=60, value=15, step=5,
                    key=f"pip_tol_{selected_no}",
                )
            with pip_col2:
                pip_threshold = st.number_input(
                    "MB スコア閾値（以上で正式タイトル使用）",
                    min_value=0, max_value=100, value=80, step=5,
                    key=f"pip_thresh_{selected_no}",
                )

            pip_opt1, pip_opt2 = st.columns(2)
            with pip_opt1:
                _pip_claude_avail = False
                try:
                    from modules.claude_lookup import is_available as _pip_claude_check
                    _pip_claude_avail = _pip_claude_check()
                except Exception:
                    pass
                pip_use_claude = st.checkbox(
                    "Claude API も使う",
                    value=False,
                    key=f"pip_use_claude_{selected_no}",
                    disabled=not _pip_claude_avail,
                    help="CD名・作曲者等をClaudeに問い合わせます" if _pip_claude_avail else "ANTHROPIC_API_KEY が未設定",
                )
            with pip_opt2:
                _pip_sp_avail = spotify_available()
                pip_use_spotify = st.checkbox(
                    "Spotify API も使う",
                    value=False,
                    key=f"pip_use_spotify_{selected_no}",
                    disabled=not _pip_sp_avail,
                    help="アーティスト・アルバム・ISRCをSpotifyから取得します" if _pip_sp_avail else "SPOTIFY_CLIENT_ID / SECRET が未設定",
                )

            wav_dur_raw = str(row.get("WAVフル尺", "")).strip()
            wav_dur_sec_val = _hms_to_sec(wav_dur_raw)
            if wav_dur_raw and wav_dur_raw.lower() != "nan":
                st.caption(f"WAV フル尺: {wav_dur_raw}（{wav_dur_sec_val:.1f} 秒）をMusicBrainz の尺絞り込みに使用します。")
            else:
                st.caption("WAV フル尺が未取得のため MusicBrainz は尺絞り込みなしで検索します。")

            if st.button(
                "🔄 全自動調査パイプラインを実行",
                key=f"pipeline_btn_{selected_no}",
                type="primary",
                use_container_width=True,
            ):
                with st.spinner("① 検索語決定 → ② Claude/MusicBrainz/Spotify → ③ J-WID / NexTone 検索 中..."):
                    pip_result = run_pipeline(
                        event_name=str(row.get("イベント名", "")),
                        wav_full_duration=wav_dur_raw,
                        wav_detected_title=str(row.get("WAV検出タイトル", "")),
                        tolerance_sec=float(pip_tolerance),
                        mb_score_threshold=int(pip_threshold),
                        use_claude=bool(pip_use_claude),
                        use_spotify=bool(pip_use_spotify),
                    )
                st.session_state[f"pipeline_result_{selected_no}"] = pip_result

            # ---- パイプライン結果表示 ----
            pip_result = st.session_state.get(f"pipeline_result_{selected_no}")
            if pip_result:
                st.markdown("---")

                # ステップ 1: 検索語
                st.markdown(f"**① 検索タイトル:** `{pip_result['search_title']}`")

                # ステップ 2: Claude API（オプション）
                _cl_res = pip_result.get("claude_result")
                if _cl_res is not None:
                    if _cl_res.get("error"):
                        st.warning(f"② Claude: エラー — {_cl_res['error']}")
                    else:
                        _conf = _cl_res.get("confidence", "")
                        _conf_icon = {"high": "🟢", "medium": "🟡", "low": "🔴"}.get(_conf, "⚪")
                        with st.expander(
                            f"② Claude 情報 {_conf_icon} 確信度:{_conf}  "
                            f"『{_cl_res.get('official_title','')}』/ {_cl_res.get('artist','')}",
                            expanded=True,
                        ):
                            cl_c1, cl_c2 = st.columns(2)
                            cl_c1.text_input("正式タイトル", value=_cl_res.get("official_title",""), key=f"cl_title_{selected_no}", disabled=True)
                            cl_c1.text_input("アーティスト", value=_cl_res.get("artist",""),         key=f"cl_art_{selected_no}",   disabled=True)
                            cl_c1.text_input("CD / アルバム名", value=_cl_res.get("cd_name",""),     key=f"cl_cd_{selected_no}",    disabled=True)
                            cl_c2.text_input("作曲者",        value=_cl_res.get("composer",""),      key=f"cl_comp_{selected_no}",  disabled=True)
                            cl_c2.text_input("作詞者",        value=_cl_res.get("lyricist",""),      key=f"cl_lyric_{selected_no}", disabled=True)
                            cl_c2.text_input("ISRC",          value=_cl_res.get("isrc",""),          key=f"cl_isrc_{selected_no}",  disabled=True)
                            if _cl_res.get("notes"):
                                st.caption(f"備考: {_cl_res['notes']}")
                            if st.button("✅ Claude 情報を適用", key=f"cl_apply_{selected_no}", use_container_width=True):
                                for col, val in {
                                    "アーティスト": _cl_res.get("artist",""),
                                    "作曲者":       _cl_res.get("composer",""),
                                    "作詞者":       _cl_res.get("lyricist",""),
                                    "CD番号":       _cl_res.get("cd_name",""),
                                    "確認ステータス": "候補あり",
                                }.items():
                                    if val and col in st.session_state.songs_df.columns:
                                        st.session_state.songs_df.at[row_idx, col] = val
                                st.success("楽曲まとめに反映しました。")

                # ステップ 3: MusicBrainz
                mb_best = pip_result.get("mb_best")
                mb_all = [r for r in pip_result.get("mb_results", []) if "error" not in r]
                mb_err = next((r["error"] for r in pip_result.get("mb_results", []) if "error" in r), None)

                if mb_err:
                    st.warning(f"③ MusicBrainz エラー: {mb_err}")
                elif not mb_all:
                    st.warning("③ MusicBrainz: 該当なし — 検索タイトルで J-WID を検索しました")
                else:
                    score = mb_best.get("score", 0) if mb_best else 0
                    dur_disp = f"{mb_best['duration_sec']:.1f}s" if mb_best and mb_best.get("duration_sec") else ""
                    isrc_disp = mb_best.get("isrc", "") if mb_best else ""
                    official_title = mb_best.get("title", "") if mb_best else ""
                    st.success(
                        f"③ MusicBrainz: **{official_title}** ／ {mb_best.get('artist','') if mb_best else ''}"
                        f"  スコア:{score}  尺:{dur_disp}  ISRC:{isrc_disp or '(なし)'}"
                    )
                    if len(mb_all) > 1:
                        with st.expander(f"MusicBrainz 全候補（{len(mb_all)}件）"):
                            for i, r in enumerate(mb_all):
                                d = f"{r['duration_sec']:.1f}s" if r.get("duration_sec") else ""
                                st.markdown(
                                    f"`[{r['score']}]` **{r['title']}** / {r['artist']}  {d}  "
                                    f"ISRC:{r.get('isrc','(なし)')}  "
                                    f"[MBリンク]({r.get('mb_url','')})"
                                )

                # ステップ 4: Spotify（オプション）
                _sp_res = pip_result.get("sp_results", [])
                _sp_best = pip_result.get("sp_best")
                if _sp_res:
                    _sp_err = next((r["error"] for r in _sp_res if "error" in r), None)
                    if _sp_err:
                        st.warning(f"④ Spotify エラー: {_sp_err}")
                    elif not [r for r in _sp_res if "error" not in r]:
                        st.info("④ Spotify: 該当なし")
                    else:
                        _sp_valid = [r for r in _sp_res if "error" not in r]
                        _spb = _sp_best or _sp_valid[0]
                        st.success(
                            f"④ Spotify: **{_spb.get('title','')}** ／ {_spb.get('artist','')}  "
                            f"アルバム:{_spb.get('album','')}  "
                            f"尺:{_spb.get('duration_sec',0):.1f}s  "
                            f"ISRC:{_spb.get('isrc','(なし)')}"
                        )
                        _sp_apply_col, _sp_link_col = st.columns([3, 1])
                        with _sp_link_col:
                            if _spb.get("spotify_url"):
                                st.link_button("🎵 Spotify で開く", _spb["spotify_url"], use_container_width=True)
                        with _sp_apply_col:
                            if st.button("✅ Spotify 情報を適用", key=f"sp_apply_pip_{selected_no}", use_container_width=True):
                                for col, val in {
                                    "アーティスト": _spb.get("artist", ""),
                                    "CD番号":       _spb.get("album", ""),
                                    "確認ステータス": "候補あり",
                                }.items():
                                    if val and col in st.session_state.songs_df.columns:
                                        st.session_state.songs_df.at[row_idx, col] = val
                                st.success("楽曲まとめに反映しました。")
                        if len(_sp_valid) > 1:
                            with st.expander(f"Spotify 全候補（{len(_sp_valid)}件）"):
                                for i, r in enumerate(_sp_valid):
                                    d = f"{r['duration_sec']:.1f}s" if r.get("duration_sec") else ""
                                    st.markdown(
                                        f"`[{r['score']}]` **{r['title']}** / {r['artist']}  "
                                        f"『{r.get('album','')}』  {d}  "
                                        f"ISRC:{r.get('isrc','(なし)')}  "
                                        f"[Spotify]({r.get('spotify_url','')})"
                                    )

                # ステップ 5: 検索語
                st.markdown(f"**⑤ J-WID / NexTone 検索語:** `{pip_result['jwid_search_term']}`")

                # ステップ 5: J-WID 結果
                jwid_r = pip_result["jwid_results"]
                ntone_r = pip_result["nextone_results"]

                pip_tab_j, pip_tab_n = st.tabs(["📋 J-WID 結果", "📋 NexTone 結果"])

                with pip_tab_j:
                    st.caption(f"検索URL: {jwid_r.get('search_url','')}")
                    if jwid_r.get("error"):
                        st.error(f"エラー: {jwid_r['error']}")
                    elif not jwid_r.get("results"):
                        st.warning("J-WID: 該当なし")
                        with st.expander("デバッグ HTML"):
                            st.code(jwid_r.get("debug_html", "")[:3000], language="html")
                    else:
                        st.success(f"{len(jwid_r['results'])} 件")
                        for i, item in enumerate(jwid_r["results"]):
                            with st.expander(
                                f"候補{i+1}: {item.get('作品名','')} ／ {item.get('作品コード','')}",
                                expanded=(i == 0),
                            ):
                                pc1, pc2 = st.columns(2)
                                pc1.text_input("作品コード", value=item.get("作品コード",""), key=f"pip_j_code_{selected_no}_{i}", disabled=True)
                                pc1.text_input("作品名",    value=item.get("作品名",""),    key=f"pip_j_title_{selected_no}_{i}", disabled=True)
                                pc1.text_input("作曲者",    value=item.get("作曲者",""),    key=f"pip_j_comp_{selected_no}_{i}",  disabled=True)
                                pc2.text_input("作詞者",    value=item.get("作詞者",""),    key=f"pip_j_lyric_{selected_no}_{i}", disabled=True)
                                pc2.text_input("編曲者",    value=item.get("編曲者",""),    key=f"pip_j_arr_{selected_no}_{i}",   disabled=True)
                                if st.button("✅ 適用", key=f"pip_apply_j_{selected_no}_{i}", use_container_width=True):
                                    for col, val in {
                                        "作曲者": item.get("作曲者",""),
                                        "作詞者": item.get("作詞者",""),
                                        "JASRAC作品コード": item.get("作品コード",""),
                                        "アーティスト": mb_best.get("artist","") if mb_best else "",
                                        "確認ステータス": "候補あり",
                                    }.items():
                                        if val:
                                            st.session_state.songs_df.at[row_idx, col] = val
                                    st.success("楽曲まとめに反映しました。")

                with pip_tab_n:
                    st.caption(f"検索URL: {ntone_r.get('search_url','')}")
                    if ntone_r.get("error"):
                        st.error(f"エラー: {ntone_r['error']}")
                    elif not ntone_r.get("results"):
                        st.warning("NexTone: 該当なし")
                    else:
                        st.success(f"{len(ntone_r['results'])} 件")
                        for i, item in enumerate(ntone_r["results"]):
                            with st.expander(
                                f"候補{i+1}: {item.get('作品名','')} ／ {item.get('管理番号','')}",
                                expanded=(i == 0),
                            ):
                                nc1, nc2 = st.columns(2)
                                nc1.text_input("管理番号",    value=item.get("管理番号",""),    key=f"pip_n_id_{selected_no}_{i}",    disabled=True)
                                nc1.text_input("作品名",      value=item.get("作品名",""),      key=f"pip_n_title_{selected_no}_{i}", disabled=True)
                                nc1.text_input("作曲者",      value=item.get("作曲者",""),      key=f"pip_n_comp_{selected_no}_{i}",  disabled=True)
                                nc2.text_input("作詞者",      value=item.get("作詞者",""),      key=f"pip_n_lyric_{selected_no}_{i}", disabled=True)
                                nc2.text_input("アーティスト", value=item.get("アーティスト",""), key=f"pip_n_art_{selected_no}_{i}",   disabled=True)
                                if st.button("✅ 適用", key=f"pip_apply_n_{selected_no}_{i}", use_container_width=True):
                                    for col, val in {
                                        "作曲者": item.get("作曲者",""),
                                        "作詞者": item.get("作詞者",""),
                                        "NexTone管理番号": item.get("管理番号",""),
                                        "アーティスト": item.get("アーティスト","") or (mb_best.get("artist","") if mb_best else ""),
                                        "確認ステータス": "候補あり",
                                    }.items():
                                        if val:
                                            st.session_state.songs_df.at[row_idx, col] = val
                                    st.success("楽曲まとめに反映しました。")

            st.divider()

            # ---- 個別手動調査セクション ----
            st.markdown("#### 🤖 J-WID / NexTone 自動調査（個別）")

            # 検索語を選択できるようにする
            search_term = st.selectbox(
                "使用する検索語",
                options=[t for _, t in term_candidates],
                key=f"auto_term_{selected_no}",
            )

            if st.button(
                f'🔍 「{search_term}」で J-WID / NexTone を自動検索',
                key=f"auto_search_{selected_no}",
                type="primary",
                use_container_width=True,
            ):
                with st.spinner("J-WID / NexTone を検索中... (各サイト約2秒待機します)"):
                    results = search_all(search_term)
                    st.session_state[f"scrape_results_{selected_no}"] = results

            # ---- 結果表示 ----
            results = st.session_state.get(f"scrape_results_{selected_no}")
            if results:
                jwid_res   = results["jwid"]
                nextone_res = results["nextone"]

                tab_j, tab_n = st.tabs(["📋 J-WID 結果", "📋 NexTone 結果"])

                # ---- J-WID 結果 ----
                with tab_j:
                    st.caption(f"検索URL: {jwid_res['search_url']}")

                    if jwid_res["error"]:
                        st.error(f"エラー: {jwid_res['error']}")
                        with st.expander("デバッグ: 取得 HTML（先頭 3000 文字）"):
                            st.code(jwid_res.get("debug_html", ""), language="html")

                    elif not jwid_res["results"]:
                        st.warning("J-WID: 該当なし（または HTML パース失敗）")
                        with st.expander("デバッグ: 取得 HTML（先頭 3000 文字）"):
                            st.code(jwid_res.get("debug_html", ""), language="html")

                    else:
                        st.success(f"{len(jwid_res['results'])} 件見つかりました")
                        for i, item in enumerate(jwid_res["results"]):
                            with st.expander(
                                f"候補 {i+1}: {item.get('作品名', '(作品名不明)')} ／ {item.get('作品コード', '')}",
                                expanded=(i == 0),
                            ):
                                c1, c2 = st.columns(2)
                                c1.text_input("作品コード（JASRAC）", value=item.get("作品コード", ""), key=f"j_code_{selected_no}_{i}", disabled=True)
                                c1.text_input("作品名",   value=item.get("作品名", ""),  key=f"j_title_{selected_no}_{i}",    disabled=True)
                                c1.text_input("作曲者",   value=item.get("作曲者", ""),  key=f"j_comp_{selected_no}_{i}",     disabled=True)
                                c2.text_input("作詞者",   value=item.get("作詞者", ""),  key=f"j_lyric_{selected_no}_{i}",    disabled=True)
                                c2.text_input("編曲者",   value=item.get("編曲者", ""),  key=f"j_arr_{selected_no}_{i}",      disabled=True)
                                c2.text_input("出版者",   value=item.get("出版者", ""),  key=f"j_pub_{selected_no}_{i}",      disabled=True)

                                if st.button(
                                    "✅ このデータを楽曲まとめに適用",
                                    key=f"apply_jwid_{selected_no}_{i}",
                                    use_container_width=True,
                                ):
                                    _apply_fields = {
                                        "作曲者":         item.get("作曲者", ""),
                                        "作詞者":         item.get("作詞者", ""),
                                        "アーティスト":   item.get("アーティスト", ""),
                                        "JASRAC作品コード": item.get("作品コード", ""),
                                        "確認ステータス": "J-WID要確認" if not item.get("作品コード") else "候補あり",
                                    }
                                    for col, val in _apply_fields.items():
                                        if val:
                                            st.session_state.songs_df.at[row_idx, col] = val
                                    st.success("楽曲まとめに反映しました。「楽曲まとめ」タブで確認してください。")

                # ---- NexTone 結果 ----
                with tab_n:
                    st.caption(f"検索URL: {nextone_res['search_url']}")

                    if nextone_res["error"]:
                        st.error(f"エラー: {nextone_res['error']}")
                        with st.expander("デバッグ: 取得 HTML（先頭 3000 文字）"):
                            st.code(nextone_res.get("debug_html", ""), language="html")

                    elif not nextone_res["results"]:
                        st.warning("NexTone: 該当なし（または HTML パース失敗）")
                        with st.expander("デバッグ: 取得 HTML（先頭 3000 文字）"):
                            st.code(nextone_res.get("debug_html", ""), language="html")

                    else:
                        st.success(f"{len(nextone_res['results'])} 件見つかりました")
                        for i, item in enumerate(nextone_res["results"]):
                            with st.expander(
                                f"候補 {i+1}: {item.get('作品名', '(作品名不明)')} ／ {item.get('管理番号', '')}",
                                expanded=(i == 0),
                            ):
                                c1, c2 = st.columns(2)
                                c1.text_input("管理番号（NexTone）", value=item.get("管理番号", ""),    key=f"n_id_{selected_no}_{i}",    disabled=True)
                                c1.text_input("作品名",              value=item.get("作品名", ""),      key=f"n_title_{selected_no}_{i}", disabled=True)
                                c1.text_input("作曲者",              value=item.get("作曲者", ""),      key=f"n_comp_{selected_no}_{i}",  disabled=True)
                                c2.text_input("作詞者",              value=item.get("作詞者", ""),      key=f"n_lyric_{selected_no}_{i}", disabled=True)
                                c2.text_input("アーティスト",        value=item.get("アーティスト", ""), key=f"n_art_{selected_no}_{i}",   disabled=True)
                                c2.text_input("アルバム",            value=item.get("アルバム", ""),    key=f"n_alb_{selected_no}_{i}",   disabled=True)

                                if st.button(
                                    "✅ このデータを楽曲まとめに適用",
                                    key=f"apply_nt_{selected_no}_{i}",
                                    use_container_width=True,
                                ):
                                    _apply_fields = {
                                        "作曲者":          item.get("作曲者", ""),
                                        "作詞者":          item.get("作詞者", ""),
                                        "アーティスト":    item.get("アーティスト", ""),
                                        "NexTone管理番号": item.get("管理番号", ""),
                                        "確認ステータス":  "NexTone要確認" if not item.get("管理番号") else "候補あり",
                                    }
                                    for col, val in _apply_fields.items():
                                        if val:
                                            st.session_state.songs_df.at[row_idx, col] = val
                                    st.success("楽曲まとめに反映しました。「楽曲まとめ」タブで確認してください。")

        # ---- MusicBrainz セクション ----
        st.divider()
        st.markdown("#### 🎵 MusicBrainz 調査（曲名 + 尺）")
        st.caption("国際楽曲データベースで曲名・アーティスト・ISRC を検索します。日本のライブラリ音楽は収録が薄い場合があります。")

        mb_col_term, mb_col_dur = st.columns([3, 2])
        with mb_col_term:
            mb_search_term = st.selectbox(
                "検索に使用する曲名",
                options=[t for _, t in term_candidates],
                key=f"mb_term_{selected_no}",
            )
        with mb_col_dur:
            wav_dur_raw = str(row.get("WAVフル尺", "")).strip()
            wav_dur_sec = _hms_to_sec(wav_dur_raw)
            use_duration = st.checkbox(
                f"尺で絞り込む（{wav_dur_raw or '未取得'}）",
                value=(wav_dur_sec > 0),
                key=f"mb_use_dur_{selected_no}",
                disabled=(wav_dur_sec <= 0),
            )
            tolerance = st.slider(
                "許容誤差（秒）",
                min_value=5, max_value=60, value=15, step=5,
                key=f"mb_tol_{selected_no}",
                disabled=not use_duration,
            )

        if st.button(
            f'🎵 「{mb_search_term}」で MusicBrainz を検索',
            key=f"mb_search_{selected_no}",
            use_container_width=True,
        ):
            dur_arg = wav_dur_sec if use_duration else None
            with st.spinner("MusicBrainz を検索中..."):
                mb_results = search_recording(
                    title=mb_search_term,
                    duration_sec=dur_arg,
                    tolerance_sec=tolerance,
                    limit=10,
                )
            st.session_state[f"mb_results_{selected_no}"] = mb_results

        mb_results = st.session_state.get(f"mb_results_{selected_no}")
        if mb_results:
            if mb_results and "error" in mb_results[0]:
                st.error(f"MusicBrainz エラー: {mb_results[0]['error']}")
            elif not mb_results:
                st.warning("MusicBrainz: 該当なし")
            else:
                st.success(f"{len(mb_results)} 件見つかりました")
                for i, item in enumerate(mb_results):
                    score = item.get("score", 0)
                    label = f"候補 {i+1} [{score}点]: {item.get('title', '')}  ／  {item.get('artist', '')}"
                    with st.expander(label, expanded=(i == 0)):
                        mc1, mc2 = st.columns(2)
                        mc1.text_input("タイトル",  value=item.get("title", ""),        key=f"mb_title_{selected_no}_{i}", disabled=True)
                        mc1.text_input("アーティスト", value=item.get("artist", ""),    key=f"mb_art_{selected_no}_{i}",   disabled=True)
                        mc1.text_input("アルバム",  value=item.get("album", ""),        key=f"mb_alb_{selected_no}_{i}",   disabled=True)
                        dur_disp = f"{item['duration_sec']:.1f} 秒" if item.get("duration_sec") else ""
                        mc2.text_input("再生時間",  value=dur_disp,                     key=f"mb_dur_{selected_no}_{i}",   disabled=True)
                        mc2.text_input("ISRC",      value=item.get("isrc", ""),         key=f"mb_isrc_{selected_no}_{i}",  disabled=True)
                        mc2.text_input("スコア",    value=str(score),                   key=f"mb_score_{selected_no}_{i}", disabled=True)

                        st.link_button(
                            "🔗 MusicBrainz で詳細を確認",
                            item.get("mb_url", "https://musicbrainz.org"),
                            use_container_width=True,
                        )

                        if st.button(
                            "✅ アーティスト情報を楽曲まとめに適用",
                            key=f"apply_mb_{selected_no}_{i}",
                            use_container_width=True,
                        ):
                            _mb_apply = {
                                "アーティスト": item.get("artist", ""),
                            }
                            for col, val in _mb_apply.items():
                                if val:
                                    st.session_state.songs_df.at[row_idx, col] = val
                            st.success("アーティスト情報を反映しました。「楽曲まとめ」タブで確認してください。")

        # ---- 全検索語一覧 ----
        st.divider()
        st.subheader("全検索語一覧")
        if st.session_state.search_df is not None and len(st.session_state.search_df) > 0:
            st.dataframe(
                st.session_state.search_df[["No", "イベント名", "検索語ラベル", "検索語"]],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("照合実行後に表示されます。")


# =====================================================================
# タブ 5: Excel 出力
# =====================================================================
with tabs[4]:
    st.header("Excel 出力")

    if st.session_state.songs_df is None:
        st.info("「ファイル読み込み」タブで照合を実行してください。")
    else:
        st.markdown(
            """
            以下の 6 シートを含む Excel ファイルを生成します：

            | シート名 | 内容 |
            |---------|------|
            | 楽曲まとめ | 権利調査用メインシート（手入力情報含む） |
            | イベント一覧 | NUENDO イベント単位の全一覧 |
            | WAV 一覧 | 読み込んだ WAV ファイル一覧 |
            | MP3 一覧 | 読み込んだ MP3 ファイル一覧 |
            | 検索語一覧 | J-WID / NexTone 検索用語一覧 |
            | 確認メモ | 処理ルール・ステータス凡例 |
            """
        )

        out_filename = st.text_input(
            "出力ファイル名",
            value="著作権調査.xlsx",
            key="excel_filename",
        )

        if st.button(
            "📊 Excel ファイルを生成", type="primary", use_container_width=True
        ):
            with st.spinner("Excel 生成中..."):
                excel_bytes = export_to_excel(
                    st.session_state.songs_df,
                    st.session_state.events_df,
                    st.session_state.wav_df,
                    st.session_state.mp3_df,
                    st.session_state.search_df,
                )

            st.download_button(
                label="⬇️ ダウンロード",
                data=excel_bytes,
                file_name=out_filename,
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
                use_container_width=True,
                type="primary",
            )
            st.success("✅ Excel ファイルの準備ができました！「ダウンロード」ボタンを押してください。")

        st.divider()
        st.subheader("現在の楽曲まとめ（確認用）")
        st.dataframe(
            st.session_state.songs_df[
                ["No", "イベント名", "確認ステータス", "WAV照合ステータス", "作曲者", "JASRAC作品コード", "NexTone管理番号"]
            ],
            use_container_width=True,
            hide_index=True,
        )
