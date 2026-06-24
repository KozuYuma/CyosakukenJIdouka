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
from modules.scraper import search_all
from modules.search_helper import JWID_BASE, generate_search_terms

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
        "songs_df": None,
        "events_df": None,
        "search_df": None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


_init_session()

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

    # ---- WAV CSV ----
    with col_right:
        st.subheader("② WAV 一覧 CSV（必須）")
        st.caption(
            "PowerShell スクリプト（scripts/Get-WavList.ps1）で取得した"
            " Audio フォルダの WAV 一覧 CSV をアップロードしてください。"
        )
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

    st.divider()

    col_left2, col_right2 = st.columns(2)

    # ---- MP3 CSV（任意） ----
    with col_left2:
        st.subheader("③ MP3 一覧 CSV（任意・補助）")
        st.caption(
            "WAV で照合できなかった場合の補助データです。"
            " WAV で全件照合できる場合は不要です。"
        )
        mp3_file = st.file_uploader(
            "MP3 一覧 CSV を選択（任意）", type=["csv"], key="upload_mp3"
        )
        if mp3_file:
            try:
                df, enc = read_csv_auto(mp3_file)
                st.session_state.mp3_df = df
                st.success(f"✅ 読み込み完了（{enc}）: {len(df)} 件")
                with st.expander("プレビュー（先頭 5 行）"):
                    st.dataframe(df.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"❌ 読み込みエラー: {e}")

    # ---- 既存 Excel（任意・将来実装） ----
    with col_right2:
        st.subheader("④ 既存調査 Excel（任意）")
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

            st.divider()

            # ---- 自動調査セクション ----
            st.markdown("#### 🤖 J-WID / NexTone 自動調査")

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
