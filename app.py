"""
著作権調査支援ツール - メインアプリ
NUENDO Cue CSV × WAV 一覧照合 → 権利情報管理 → Excel 出力

起動方法: run.bat をダブルクリック
        または: streamlit run app.py
"""
import json
import os
import re
import threading
import time
import unicodedata
import urllib.parse
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as _stc

from modules.auth import current_user, is_enabled as auth_enabled, logout_button, require_login
from modules.csv_reader import (
    drop_non_music_events,
    music_tracks,
    normalize_cue_columns,
    read_csv_auto,
    validate_cue_csv,
    validate_wav_csv,
)
from modules.database import (
    cd_count,
    cd_fetch_by_jasrac,
    create_project,
    delete_project,
    describe_backend,
    init_db,
    list_projects,
    load_events,
    load_songs,
    master_count,
    master_delete,
    master_search,
    save_events,
    save_songs,
    set_project_owner,
)
from modules.cd_master import fill as cd_fill
from modules.excel_exporter import export_to_excel, build_shinkok_df, _SHINKOK_RENAME
from modules.matcher import build_song_list
from modules.song_master import (
    MASTER_FIELDS,
    cell_of as master_cell,
    edit as master_edit,
    fill as master_fill,
    save as master_learn,
    sources_of as master_sources,
    to_frame as master_frame,
)
from modules.musicbrainz import _hms_to_sec, mb_search_url, search_recording
from modules.musicforest import (
    MusicForestClient,
    MusicForestError,
    check_session,
    get_state_path,
    load_client,
    sync_session_from_chrome,
    update_sess_cookie,
)
from modules.normalizer import normalize_for_match
from modules.pipeline import run_pipeline
from modules.scraper import search_all
from modules.search_helper import JWID_BASE, generate_search_terms
from modules.spotify import is_available as spotify_available, spotify_search_url
from modules.ui import (
    ISSUE_MARK_LEGEND,
    count_done,
    inject_css,
    issue_mark,
    status_bar,
)

# =====================================================================
# アプリ設定
# =====================================================================
st.set_page_config(
    page_title="著作権調査支援ツール",
    page_icon="🎵",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# 配色・書体は .streamlit/config.toml。ここでは余白と共通部品だけ入れる
inject_css()

# 手元の Windows で動いているか。サーバー（Render）では False。
#
# ブラウザを開く・利用者のフォルダを読む・Chrome の Cookie を読む、と
# いった機能は手元でしか動かない。サーバーで押されたときに黙って例外を
# 出すのではなく、なぜ使えないかを画面に出すために使う。
IS_LOCAL_WINDOWS = os.name == "nt"

# セクションへ移動するときのスクロール。`__ID__` を移動先の id に置き換えて使う。
#
# scrollIntoView() は目印を画面のてっぺんに合わせるが、Streamlit の見出し帯は
# 画面に貼り付いたまま上に重なるので、見たい行がその下に隠れて「少し行き過ぎた」
# ように見える。そこで帯の高さの分だけ手前で止める。
#
# もう一つ、押した直後は表がまだ描き上がっておらず、あとから背が伸びて位置が
# ずれる。一度きりでは合わないので、間を空けて数回だけ直す。ただし利用者が
# 自分でスクロールを始めたら、そこで手を引く（勝手に動かされると鬱陶しいため）。
_SCROLL_JS = """
<script>
(function () {
  var doc = parent.document, win = parent;
  var id = "__ID__", cancelled = false;
  ["wheel", "touchstart", "keydown"].forEach(function (ev) {
    win.addEventListener(ev, function () { cancelled = true; },
                         { passive: true, once: true });
  });
  // 画面に貼り付いている見出し帯の高さ。隠れているときは 0
  function headBottom() {
    var h = doc.querySelector('header[data-testid="stHeader"]');
    if (!h) return 0;
    var st = win.getComputedStyle(h);
    if (st.position !== "fixed" && st.position !== "sticky") return 0;
    if (st.display === "none" || st.visibility === "hidden") return 0;
    return h.getBoundingClientRect().bottom;
  }
  // 目印を実際に動かしている入れ物。窓そのもののこともある
  function scroller(e) {
    var n = e.parentElement;
    while (n && n !== doc.body) {
      var oy = win.getComputedStyle(n).overflowY;
      if ((oy === "auto" || oy === "scroll") && n.scrollHeight > n.clientHeight + 4)
        return n;
      n = n.parentElement;
    }
    return null;
  }
  function go(smooth) {
    if (cancelled) return;
    var e = doc.getElementById(id);
    if (!e) return;
    var sc = scroller(e);
    var scTop = sc ? sc.getBoundingClientRect().top : 0;
    // 帯の下、または入れ物の上端。どちらか低い方の少し下に置く
    var want = Math.max(scTop, headBottom()) + 12;
    var delta = e.getBoundingClientRect().top - want;
    if (Math.abs(delta) < 4) return;
    (sc || win).scrollBy({ top: delta, behavior: smooth ? "smooth" : "auto" });
  }
  [450, 1000, 1700].forEach(function (ms, i) {
    setTimeout(function () { go(i === 0); }, ms);
  });
})();
</script>
"""

# 確認ステータスの選択肢（全画面共通）
CONFIRM_STATUS_OPTIONS = [
    "未調査",
    "該当なし",
    "作曲者一致",
    "アーティスト一致",
    "台帳一致",
    "候補あり",
    "複数候補あり",
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
    "曲名",
    "使用形態",
    "音源区分",
    "I/V区分",
    "邦洋区分",
    "原訳詞区分",
    "作曲者",
    "作詞者",
    "編曲者",
    "訳詞者",
    "アーティスト",
    "レコード会社名",
    "CD番号",
    "CD名",
    "JASRAC作品コード",
    "NexTone管理番号",
    "委任者",
    "確認ステータス",
    "自社楽曲ID",
    "メモ",
]

# 申告フォーマット列 → songs_df 列の逆引きマッピング（on_change コールバック用）
_SHINKOK_RENAME_REV: dict[str, str] = {v: k for k, v in _SHINKOK_RENAME.items()}
# 申告フォーマットに固有の追加列（_SHINKOK_RENAME に含まれないが songs_df に直接対応する列）
_SHINKOK_EXTRA_COLS = ("確認ステータス", "委任者", "CD名")

#: 表に出す J-WID 管理状況の項目。名前は J-WID の 管理状況 dict のキーと
#: 同じ。管理状況は全部で18項目あるが、申告のときに毎回 J-WID を開いて
#: 見に行っていたのはこの2つなので、ここだけ楽曲の欄として持ち回る。
JWID_MGMT_COLS = ("放送", "配信")

#: 放送・配信を出どころ別に持っておく欄。J は JASRAC（J-WID）、
#: N は NexTone。同じ曲が両方に載っていて、JASRAC では × でも NexTone
#: では ○ ということがある。片方だけ見ると「管理なし」に見えてしまうので、
#: 生の値は別々に持ち、表に出す「放送」「配信」は両方を合わせて作る。
MGMT_SOURCE_COLS = {c: (f"{c}J", f"{c}N") for c in JWID_MGMT_COLS}
_MGMT_RAW_COLS = tuple(c for pair in MGMT_SOURCE_COLS.values() for c in pair)

# JASRACコード変更時にクリアすべき songs_df 列
_CLEAR_ON_JCD_CHANGE = [
    "CD番号", "CD名", "レコード会社名", "委任者",
    "邦洋区分", "原訳詞区分", "I/V区分",
    "作詞者", "作曲者", "編曲者", "訳詞者",
    # 管理状況は作品コードに紐づくので、コードが変われば意味を失う。
    # 消すのは JASRAC 由来の値だけ。NexTone 側は管理番号の話なので残す
    *(pair[0] for pair in MGMT_SOURCE_COLS.values()),
]


def _normalize_jcd(s: str) -> str:
    """JASRACコードをハイフン・空白除去 + 大文字化した正規化文字列で返す（比較用）。"""
    return re.sub(r"[-\s]", "", str(s)).upper().strip()


def _apply_clear_on_jcd_change(row_idx: int, new_jcd: str) -> None:
    """JASRACコードが変わる場合に関連フィールドを空文字クリアする。"""
    songs = st.session_state.songs_df
    if "JASRAC作品コード" not in songs.columns:
        return
    cur_norm = _normalize_jcd(str(songs.at[row_idx, "JASRAC作品コード"]))
    new_norm = _normalize_jcd(new_jcd)
    if new_norm and cur_norm and new_norm != cur_norm:
        for col in _CLEAR_ON_JCD_CHANGE:
            if col in songs.columns:
                songs.at[row_idx, col] = ""
        # JASRAC 側を消したので、表に出す欄は NexTone 側だけで作り直す
        for col, (_, nt_col) in MGMT_SOURCE_COLS.items():
            if col in songs.columns and nt_col in songs.columns:
                songs.at[row_idx, col] = _combine_mgmt(
                    "", songs.at[row_idx, nt_col])


def _stale_multi_mask(songs: pd.DataFrame | None):
    """昔「候補あり」と書かれた行のうち、本当は「複数候補あり」の行を選ぶ。

    一括検索は昔、複数ヒットして自動入力しなかった行にも「候補あり」と
    書いていた。人がどれか選びに行くべき行なので、今は「複数候補あり」と
    書く。既に保存されている行はそのままなので、ここで拾って直す。

    見分け方は「何も入っていないこと」。自動入力できた行には作品コードか
    管理番号（か、少なくとも作曲者・作詞者）が入る。複数ヒットの行は印を
    付けただけで何も入れていない。取りこぼす方が安全なので、全部空の行に
    限る（人が手で「候補あり」に戻した行を巻き込まないため）。
    """
    if songs is None or songs.empty or "確認ステータス" not in songs.columns:
        return None
    mask = songs["確認ステータス"].astype(str).str.strip() == "候補あり"
    for col in ("JASRAC作品コード", "NexTone管理番号", "作曲者", "作詞者"):
        if col in songs.columns:
            mask &= songs[col].astype(str).str.strip().isin(("", "nan", "None"))
    return mask


def _mgmt_missing_mask(songs: pd.DataFrame | None):
    """放送・配信をまだ引いていない行を選ぶ。

    管理状況は JASRAC作品コードで J-WID を、NexTone管理番号で NexTone を
    引いて取る。番号が入っている方だけを数に入れる。両方に番号がある曲は
    両方引く（JASRAC で × でも NexTone で ○ のことがあるため）。
    台帳で当たった行（台帳一致）は詳細を引いていないので、ここに並ぶ。
    """
    if songs is None or songs.empty:
        return None
    blank = ("", "nan", "None")
    need = None
    for code_col, side in (("JASRAC作品コード", 0), ("NexTone管理番号", 1)):
        if code_col not in songs.columns:
            continue
        has_code = ~songs[code_col].astype(str).str.strip().isin(blank)
        missing = None
        for pair in MGMT_SOURCE_COLS.values():
            src_col = pair[side]
            if src_col not in songs.columns:
                continue
            col_blank = songs[src_col].astype(str).str.strip().isin(blank)
            missing = col_blank if missing is None else (missing | col_blank)
        if missing is None:
            continue
        part = has_code & missing
        need = part if need is None else (need | part)
    return need


def _norm_hinban(s: str) -> str:
    """品番（CD番号）を比べるための形にする。JASRACコードと同じ直し方でよい。"""
    return _normalize_jcd(s)


def _delegate_missing_mask(songs: pd.DataFrame | None):
    """委任者がまだ入っていない行を選ぶ。作品の番号がある行だけ。

    委任者は MINC の CD商品詳細にしか出ていない。そこへ辿り着くには
    作品の番号（JASRAC作品コード か NexTone管理番号）が要るので、
    番号の無い行は数に入れない。データベース（台帳）で当たった行は
    MINC を引いていないので、ここに並ぶことになる。
    """
    if songs is None or songs.empty or "委任者" not in songs.columns:
        return None
    blank = ("", "nan", "None")
    has_code = None
    for col in ("JASRAC作品コード", "NexTone管理番号"):
        if col not in songs.columns:
            continue
        part = ~songs[col].astype(str).str.strip().isin(blank)
        has_code = part if has_code is None else (has_code | part)
    if has_code is None:
        return None
    return has_code & songs["委任者"].astype(str).str.strip().isin(blank)


def _pick_shinkok_row(state: dict) -> None:
    """申告フォーマットの「選択」欄を1行だけに保つ。

    st.data_editor には st.dataframe のような行選択が無いので、チェック欄を
    自前で1列持っている。ただの列なので何行でもチェックできてしまう。
    そこで新しくチェックされた行だけを残し、他の印は消す。

    消すのは表示用の DataFrame ではなく data_editor が覚えている編集差分の
    方。ここを直さないと、次に描き直しても前の印が復活してしまう。
    """
    edited = state.get("edited_rows") or {}
    checked = [int(p) for p, ch in edited.items()
               if isinstance(ch, dict) and ch.get("選択") is True]
    if not checked and not any(
            isinstance(ch, dict) and "選択" in ch for ch in edited.values()):
        return  # チェック欄は触られていない

    prev = st.session_state.get("_shinkok_sel")
    # 前回と違う行が新しく押された方。無ければ最後に見つかったもの
    _new = [p for p in checked if p != prev]
    keep = _new[-1] if _new else (checked[-1] if checked else None)

    for p, ch in edited.items():
        if isinstance(ch, dict) and "選択" in ch and int(p) != keep:
            ch["選択"] = False
    st.session_state["_shinkok_sel"] = keep


def _sync_shinkok_to_songs() -> None:
    """申告フォーマット data_editor の編集内容を songs_df へ即時反映するコールバック。

    data_editor の session_state 値は DataFrame ではなく編集差分の dict
    （{"edited_rows": {行番号: {列名: 値}}, "added_rows": [...], "deleted_rows": [...]}）。
    行番号は表示に使った DataFrame の位置なので、その DataFrame（_shinkok_src）と
    突き合わせてイベント名を引き、songs_df の該当行を特定する。
    """
    state = st.session_state.get("shinkok_editor")
    src: pd.DataFrame | None = st.session_state.get("_shinkok_src")
    songs: pd.DataFrame | None = st.session_state.get("songs_df")
    if not isinstance(state, dict) or src is None or songs is None:
        return

    _pick_shinkok_row(state)
    if "イベント名" not in src.columns or "イベント名" not in songs.columns:
        return

    _editable = set(_SHINKOK_RENAME_REV) | set(_SHINKOK_EXTRA_COLS)
    for _pos, _changes in (state.get("edited_rows") or {}).items():
        try:
            _p = int(_pos)
        except (TypeError, ValueError):
            continue
        if not (0 <= _p < len(src)):
            continue
        ev_name = str(src.iloc[_p].get("イベント名", "")).strip()
        if not ev_name:
            continue
        mask = songs["イベント名"] == ev_name
        if not mask.any():
            continue
        song_idx = songs.index[mask][0]
        for sh_col, val in (_changes or {}).items():
            if sh_col not in _editable:
                continue
            s_col = _SHINKOK_RENAME_REV.get(sh_col, sh_col)
            if s_col not in songs.columns:
                continue
            songs.at[song_idx, s_col] = "" if val is None or pd.isna(val) else val


def _show_cd_panel(
    jcd: str, row_idx: int, key_prefix: str, title: str = "", artist: str = ""
) -> None:
    """JASRACコードに紐づくCDリストを折りたたみなしでインライン表示し、申告フォーマットに反映できるようにする。"""
    if not jcd or str(jcd).strip().lower() in ("", "nan"):
        return
    _cp_res_key = f"cpanel_res_{key_prefix}"
    _cp_btn_key = f"cpanel_btn_{key_prefix}"

    if st.button("🔍 このJASRACコードのCDを検索（MINC）", key=_cp_btn_key, use_container_width=True):
        with st.spinner("CDリストを取得中..."):
            try:
                _cp_r = _get_mf_client().search_cds_by_jasrac(jcd, title=title)
                st.session_state[_cp_res_key] = _cp_r
                for _ck in [k for k in st.session_state if k.startswith(f"cpanel_det_{key_prefix}_")]:
                    del st.session_state[_ck]
            except MusicForestError as _cp_e:
                st.session_state[_cp_res_key] = {"cds": [], "error": str(_cp_e)}

    _render_cd_results(st.session_state.get(_cp_res_key), row_idx, key_prefix, artist)


def _cp_current_houyo(row_idx: int) -> str:
    """songs_df の現在の邦洋区分（空扱いの値は "" として返す）。"""
    if "邦洋区分" not in st.session_state.songs_df.columns:
        return ""
    _v = str(st.session_state.songs_df.at[row_idx, "邦洋区分"]).strip()
    return "" if _v.lower() in ("", "nan", "none") else _v


def _cp_row_jcd(row_idx: int) -> str:
    """songs_df の現在のJASRAC作品コード（ハイフン・空白を除いた比較用）。"""
    if "JASRAC作品コード" not in st.session_state.songs_df.columns:
        return ""
    _v = str(st.session_state.songs_df.at[row_idx, "JASRAC作品コード"]).strip()
    if _v.lower() in ("", "nan", "none"):
        return ""
    return re.sub(r"[-\s]", "", _v).upper()


def _render_delivery_rows(_cp_res: dict, row_idx: int, key_prefix: str) -> None:
    """CD商品が無い作品向けに、MINCの「配信曲」情報を一覧＋反映UIとして描画する。"""
    _dl = _cp_res.get("配信") or []
    if not _dl:
        return
    st.markdown(f"**🎧 配信音源（{len(_dl)}件）**")
    st.dataframe(
        pd.DataFrame([
            {
                "曲名":          d.get("曲名", ""),
                "アーティスト":   d.get("アーティスト", ""),
                "アルバム名":     d.get("アルバム名", ""),
                "ISRC":         d.get("ISRC", ""),
                "配信日":        d.get("配信日", ""),
            }
            for d in _dl
        ]),
        use_container_width=True,
        hide_index=True,
        height=min(300, 40 + 35 * len(_dl)),
    )
    _dl_sel = st.selectbox(
        "反映する配信音源を選択",
        options=list(range(len(_dl))),
        format_func=lambda i: (
            f"{_dl[i].get('アルバム名', '') or '(アルバム名なし)'}"
            f"｜{_dl[i].get('アーティスト', '')}｜{_dl[i].get('ISRC', '')}"
        ),
        key=f"cpanel_dlsel_{key_prefix}",
    )
    _dl_item = _dl[_dl_sel]
    if st.button(
        "✅ 配信音源の情報を反映（音源区分=配信）",
        key=f"cpanel_dlapply_{key_prefix}",
        use_container_width=True,
    ):
        _dl_apply = {
            "曲名":            _dl_item.get("曲名", "") or _cp_res.get("作品名", ""),
            "JASRAC作品コード": _dl_item.get("JASRAC作品コード", "") or _cp_res.get("作品コード", ""),
            "アーティスト":     _dl_item.get("アーティスト", ""),
            "CD名":            _dl_item.get("アルバム名", ""),
            "ISRC":           _dl_item.get("ISRC", ""),
            "音源区分":         "配信",
        }
        _dl_jcd = re.sub(r"[-\s]", "", str(_dl_apply["JASRAC作品コード"])).upper()
        _dl_hy = _infer_houyo(_dl_jcd)
        if _dl_hy and (not _cp_current_houyo(row_idx) or _dl_jcd != _cp_row_jcd(row_idx)):
            _dl_apply["邦洋区分"] = _dl_hy
        for _dl_col, _dl_val in _dl_apply.items():
            if _dl_val and _dl_col in st.session_state.songs_df.columns:
                st.session_state.songs_df.at[row_idx, _dl_col] = _dl_val
        st.session_state["_apply_msg"] = (
            f"配信音源「{_dl_item.get('アルバム名', '')}」の情報を反映しました"
            f"（音源区分=配信／CD商品なし）。"
        )
        st.session_state.pop("songs_editor", None)
        st.rerun()
    st.caption(
        "※ この作品はMINCにCD商品が登録されていないため、検索結果の「配信曲」から取得しています。"
        "CD番号・レコード会社名は取得できません。"
    )


def _render_cd_results(
    _cp_res: dict | None, row_idx: int, key_prefix: str, artist_filter: str = ""
) -> None:
    """search_cds_by_jasrac の結果（CD商品リスト全件）を一覧＋反映UIとして描画する。"""
    if not _cp_res:
        return

    # 収録曲取得の結果メッセージ（取得直後に rerun するのでここで出す）
    _cp_toast_key = f"cpanel_toast_{key_prefix}"
    if st.session_state.get(_cp_toast_key):
        _cp_tmsg, _cp_ticon = st.session_state.pop(_cp_toast_key)
        st.toast(_cp_tmsg, icon=_cp_ticon)
    if not _cp_res.get("cds"):
        if _cp_res.get("配信"):
            st.info(_cp_res.get("error") or "CD商品はありませんでした（配信のみの音源です）。")
        elif _cp_res.get("error"):
            st.error(f"CD検索エラー: {_cp_res['error']}")
        else:
            st.warning("収録CDが見つかりませんでした。")
        _render_delivery_rows(_cp_res, row_idx, key_prefix)
        if _cp_res.get("search_url"):
            st.caption(f"🔗 MINCで直接開く: [{_cp_res['search_url']}]({_cp_res['search_url']})")
        return
    if _cp_res.get("_cd_fallback"):
        # 作品コードにCDが紐付いていない登録（収録曲行に管理情報ボタンが無い）のため
        # 検索結果ページから拾い直したケース。エラーではないので info で出す。
        st.info(f"💡 {_cp_res['error']}")
    elif _cp_res.get("error"):
        st.warning(f"⚠️ 一部エラー: {_cp_res['error']}")

    _cp_items = _cp_res["cds"]
    _cp_head = [f"💿 「{_cp_res.get('作品名', '')}」（{_cp_res.get('作品コード', '')}）"]
    for _cp_k in ("作曲者", "作詞者"):
        if _cp_res.get(_cp_k):
            _cp_head.append(f"{_cp_k}: {_cp_res[_cp_k]}")
    st.caption(
        "　／　".join(_cp_head)
        + f"　／　CD商品 **{_cp_res.get('件数', len(_cp_items))} 件**"
    )

    # ── 検索時に指定されたアーティストで事前に絞り込む ──────────────────────
    #   0件になる場合（オムニバス盤は (V.A.) 表記）は全件表示にフォールバック
    _cp_art = str(artist_filter or "").strip()
    if _cp_art.lower() == "nan":
        _cp_art = ""
    if _cp_art:
        _cp_artl = _cp_art.lower()
        _cp_hit = [c for c in _cp_items if _cp_artl in c.get("アーティスト", "").lower()]
        if _cp_hit:
            st.caption(f"🎤 アーティスト「{_cp_art}」で絞り込み: **{len(_cp_hit)}** / {len(_cp_items)} 件")
            _cp_items = _cp_hit
        else:
            st.caption(
                f"🎤 アーティスト「{_cp_art}」に一致するCDが無いため全件表示しています"
                "（オムニバス盤は (V.A.) 表記です）"
            )

    # ── 絞り込み（品番／CD商品タイトル／アーティスト／会社名の部分一致）────────
    _cp_q = st.text_input(
        "絞り込み（品番・CDタイトル・アーティスト・会社名）",
        key=f"cpanel_q_{key_prefix}",
        placeholder="例: TOCT / ベスト / チューリップ",
    ).strip()
    if _cp_q:
        _cp_ql = _cp_q.lower()
        _cp_view = [
            c for c in _cp_items
            if _cp_ql in " ".join([
                c.get("品番", ""), c.get("CD商品タイトル", ""),
                c.get("アーティスト", ""), c.get("発売会社", ""), c.get("販売会社", ""),
            ]).lower()
        ]
    else:
        _cp_view = _cp_items

    if not _cp_view:
        st.info(f"「{_cp_q}」に一致するCDはありません。")
        return

    # ── 全件一覧（行クリックで下の「反映するCDを選択」に連動）──────────────
    _cp_sel_key  = f"cpanel_sel_{key_prefix}"
    _cp_last_key = f"cpanel_tblrow_{key_prefix}"

    # 絞り込みで件数が変わった場合に備え、保持中の選択インデックスを丸める
    if not isinstance(st.session_state.get(_cp_sel_key), int) or \
            not (0 <= st.session_state.get(_cp_sel_key, -1) < len(_cp_view)):
        st.session_state[_cp_sel_key] = 0

    _cp_ev = st.dataframe(
        pd.DataFrame([
            {
                "No":            c.get("No", ""),
                "品番":           c.get("品番", ""),
                "CD商品タイトル":  c.get("CD商品タイトル", ""),
                "アーティスト":    c.get("アーティスト", ""),
                "形態":           c.get("形態", ""),
                "曲数":           c.get("曲数", ""),
                "発売日":         c.get("発売日", ""),
                "発売会社":        c.get("発売会社", ""),
                "販売会社":        c.get("販売会社", ""),
                "権利":           "/".join(c.get("権利", [])),
                "初回盤":         "○" if c.get("初回盤") else "",
            }
            for c in _cp_view
        ]),
        use_container_width=True,
        hide_index=True,
        height=min(400, 40 + 35 * len(_cp_view)),
        on_select="rerun",
        selection_mode="single-row",
        key=f"cpanel_tbl_{key_prefix}",
    )
    st.caption("表の行をクリックすると、下の「反映するCDを選択」に反映されます。")

    # 表で選ばれた行 → セレクトボックスへ反映（変化したときだけ上書きし、
    # セレクトボックス側の手動変更を潰さないようにする）
    try:
        _cp_rows = list(_cp_ev.selection.rows)
    except AttributeError:
        _cp_rows = list((_cp_ev or {}).get("selection", {}).get("rows", []))
    if _cp_rows:
        _cp_row = int(_cp_rows[0])
        if 0 <= _cp_row < len(_cp_view) and st.session_state.get(_cp_last_key) != _cp_row:
            st.session_state[_cp_last_key] = _cp_row
            st.session_state[_cp_sel_key] = _cp_row
    else:
        st.session_state.pop(_cp_last_key, None)

    # ── 1枚選んで申告フォーマットへ反映 ──────────────────────────────────
    _cp_sel = st.selectbox(
        "反映するCDを選択",
        options=list(range(len(_cp_view))),
        format_func=lambda i: (
            f"{_cp_view[i].get('品番', '')}｜{_cp_view[i].get('CD商品タイトル', '')}"
            f"｜{_cp_view[i].get('発売日', '')}｜{_cp_view[i].get('発売会社', '')}"
        ),
        key=_cp_sel_key,
    )
    _cp_item = _cp_view[_cp_sel]
    _cp_a = _cp_item.get("album_id", "")
    _cp_t = _cp_item.get("track_id", "")

    _cp_det_key = f"cpanel_det_{key_prefix}_{_cp_a}"
    _cp_det = st.session_state.get(_cp_det_key, {})
    _cp_dlg = _cp_det.get("集中管理", "")

    _cp_dc1, _cp_dc2, _cp_dc3 = st.columns(3)
    _cp_dc1.text_input("品番",         value=_cp_item.get("品番", ""),           key=f"cpanel_cat_{key_prefix}", disabled=True)
    _cp_dc2.text_input("レコード会社", value=_cp_item.get("レコード会社名", ""), key=f"cpanel_rco_{key_prefix}", disabled=True)
    _cp_dc3.text_input("委任者区分",   value=_cp_dlg or "(詳細取得で確認)",      key=f"cpanel_dlg_{key_prefix}", disabled=True)

    _cp_b1, _cp_b2 = st.columns(2)
    with _cp_b1:
        if st.button(
            "🎵 収録曲を表示（このCDから曲を逆引き）",
            key=f"cpanel_fetch_{key_prefix}",
            use_container_width=True,
            disabled=not _cp_a,
            help="CD商品詳細から全収録曲（曲順・曲名・IV・収録時間・ISRC・JASRACコード）と委任者区分を取得します",
        ):
            with st.spinner("収録曲を取得中..."):
                try:
                    _cp_fd = _get_mf_client().fetch_track_list(
                        _cp_a, _cp_t, title=_cp_res.get("作品名", "")
                    )
                    st.session_state[_cp_det_key] = _cp_fd
                    if _cp_fd.get("error"):
                        st.session_state[_cp_toast_key] = (f"エラー: {_cp_fd['error']}", "❌")
                    else:
                        st.session_state[_cp_toast_key] = (
                            f"{len(_cp_fd.get('tracks', []))}曲を取得  "
                            f"委任者={_cp_fd.get('集中管理', '(なし)')}",
                            "✅",
                        )
                    # 収録曲リストや委任者欄はこのボタンより前で描画済みなので、
                    # 取得結果を今の画面に出すには再実行が必要（2回押さないと出ない対策）
                    st.rerun()
                except MusicForestError as _cp_fe:
                    st.toast(f"エラー: {_cp_fe}", icon="❌")
    with _cp_b2:
        if st.button(
            "✅ CD番号・レコード会社を反映",
            key=f"cpanel_apply_{key_prefix}",
            use_container_width=True,
            disabled=not (_cp_item.get("品番") or _cp_item.get("レコード会社名")),
        ):
            _cp_apply = {
                "CD番号":         _cp_item.get("品番", ""),
                "CD名":           _cp_item.get("CD商品タイトル", ""),
                "レコード会社名": _cp_item.get("レコード会社名", ""),
                "委任者":         _cp_dlg,
            }
            for _cp_col, _cp_val in _cp_apply.items():
                if _cp_val and _cp_col in st.session_state.songs_df.columns:
                    st.session_state.songs_df.at[row_idx, _cp_col] = _cp_val
            _cp_iv_str = {"I": "インスト", "V": "ヴォーカル"}.get(_cp_det.get("IV", ""), "")
            _cp_cur_iv = str(
                st.session_state.songs_df.at[row_idx, "I/V区分"]
                if "I/V区分" in st.session_state.songs_df.columns else ""
            ).strip()
            if not _cp_iv_str:
                # MINCのCD詳細にI/V表記が無い場合は作詞者の有無から決める
                # （作家名が取得済みの行に限る）
                _cp_row_lyr = str(st.session_state.songs_df.at[row_idx, "作詞者"]
                                  if "作詞者" in st.session_state.songs_df.columns else "")
                _cp_row_cmp = str(st.session_state.songs_df.at[row_idx, "作曲者"]
                                  if "作曲者" in st.session_state.songs_df.columns else "")
                if not (_is_blank(_cp_row_lyr) and _is_blank(_cp_row_cmp)):
                    _cp_iv_str = _infer_iv(_cp_row_lyr)
            if _cp_iv_str and not _cp_cur_iv:
                st.session_state.songs_df.at[row_idx, "I/V区分"] = _cp_iv_str
            # 邦洋区分（JASRACコード2文字目）
            # 行のJASRACコードと違う作品を反映する場合は上書きする
            _cp_hy0_src = re.sub(r"[-\s]", "", str(_cp_res.get("作品コード", ""))).upper()
            _cp_hy0 = _infer_houyo(_cp_hy0_src)
            if _cp_hy0 and "邦洋区分" in st.session_state.songs_df.columns and (
                not _cp_current_houyo(row_idx) or _cp_hy0_src != _cp_row_jcd(row_idx)
            ):
                st.session_state.songs_df.at[row_idx, "邦洋区分"] = _cp_hy0
            st.session_state["_apply_msg"] = (
                f"CD番号・レコード会社名を反映しました。（{_cp_item.get('品番', '')}）"
            )
            st.session_state.pop("songs_editor", None)
            st.rerun()

    # ── 収録曲（CD → 曲の逆引き）────────────────────────────────────────
    _cp_tracks = _cp_det.get("tracks") or []
    if _cp_det.get("error") and not _cp_tracks:
        st.warning(f"収録曲を取得できませんでした: {_cp_det['error']}")
        if _cp_det.get("attempts"):
            with st.expander("🔍 試したURLと応答", expanded=False):
                for _cp_at in _cp_det["attempts"]:
                    st.markdown(f"- [{_cp_at['url']}]({_cp_at['url']}) → `{_cp_at['result']}`")
    if _cp_tracks:
        st.markdown(
            f"**🎵 収録曲（{len(_cp_tracks)}曲）** — "
            f"{_cp_det.get('CD商品タイトル', '') or _cp_item.get('CD商品タイトル', '')}"
        )
        # 行クリック → 下の「この曲を申告フォーマットに反映」に連動（CD一覧と同じ操作）
        _cp_tsel_key  = f"cpanel_tsel_{key_prefix}"
        _cp_tlast_key = f"cpanel_tblrow_trk_{key_prefix}"
        if not isinstance(st.session_state.get(_cp_tsel_key), int) or \
                not (0 <= st.session_state.get(_cp_tsel_key, -1) < len(_cp_tracks)):
            st.session_state[_cp_tsel_key] = 0

        _cp_tev = st.dataframe(
            pd.DataFrame([
                {
                    "曲順":            t.get("曲順", ""),
                    "曲名":            t.get("曲名", ""),
                    "IV":             t.get("IV", ""),
                    "収録時間":         t.get("収録時間", ""),
                    "アーティスト":     t.get("アーティスト", ""),
                    "ISRC":           t.get("ISRC", ""),
                    "JASRAC作品コード":  t.get("JASRAC作品コード", ""),
                    "NexTone作品コード": t.get("NexTone作品コード", ""),
                }
                for t in _cp_tracks
            ]),
            use_container_width=True,
            hide_index=True,
            height=min(400, 40 + 35 * len(_cp_tracks)),
            on_select="rerun",
            selection_mode="single-row",
            key=f"cpanel_ttbl_{key_prefix}",
        )
        st.caption("表の行をクリックすると、下の「この曲を申告フォーマットに反映」に反映されます。")

        try:
            _cp_trows = list(_cp_tev.selection.rows)
        except AttributeError:
            _cp_trows = list((_cp_tev or {}).get("selection", {}).get("rows", []))
        if _cp_trows:
            _cp_trow = int(_cp_trows[0])
            if 0 <= _cp_trow < len(_cp_tracks) and st.session_state.get(_cp_tlast_key) != _cp_trow:
                st.session_state[_cp_tlast_key] = _cp_trow
                st.session_state[_cp_tsel_key] = _cp_trow
        else:
            st.session_state.pop(_cp_tlast_key, None)

        _cp_tsel = st.selectbox(
            "この曲を申告フォーマットに反映",
            options=list(range(len(_cp_tracks))),
            format_func=lambda i: (
                f"{_cp_tracks[i].get('曲順', '')}. {_cp_tracks[i].get('曲名', '')}"
                f"（{_cp_tracks[i].get('収録時間', '')}／{_cp_tracks[i].get('IV', '')}）"
            ),
            key=_cp_tsel_key,
        )
        _cp_trk = _cp_tracks[_cp_tsel]
        _cp_want_cred = st.checkbox(
            "作詞者・作曲者もMINCから取得して反映する（作品詳細に1回アクセスします）",
            value=True,
            key=f"cpanel_tcred_{key_prefix}",
            help="収録曲一覧には作家名が載っていないため、JASRAC作品コードから作品詳細を引いて補います。",
        )
        if st.button(
            "✅ この曲＋CD情報を反映",
            key=f"cpanel_tapply_{key_prefix}",
            use_container_width=True,
            disabled=not _cp_trk.get("曲名"),
        ):
            _cp_tapply = {
                "曲名":             _cp_trk.get("曲名", ""),
                "JASRAC作品コード":  _cp_trk.get("JASRAC作品コード", ""),
                "アーティスト":      _cp_trk.get("アーティスト", ""),
                "CD番号":           _cp_item.get("品番", "") or _cp_det.get("品番", ""),
                "CD名":             _cp_det.get("CD商品タイトル", "") or _cp_item.get("CD商品タイトル", ""),
                "レコード会社名":     _cp_item.get("レコード会社名", "") or _cp_det.get("レコード会社名", ""),
                "委任者":           _cp_det.get("集中管理", ""),
                "I/V区分":          {"I": "インスト", "V": "ヴォーカル"}.get(_cp_trk.get("IV", ""), ""),
            }

            # ── 作家名（収録曲一覧には無いので作品詳細から補う）──────────────
            _cp_cred_msg = ""
            if _cp_want_cred:
                _cp_tjcd = re.sub(r"[-\s]", "", _cp_trk.get("JASRAC作品コード", "")).upper()
                _cp_tncd = re.sub(r"[-\s]", "", _cp_trk.get("NexTone作品コード", "")).upper()
                _cp_rjcd = re.sub(r"[-\s]", "", str(_cp_res.get("作品コード", ""))).upper()
                _cp_cred: dict = {}
                # 検索中の作品と同じコードなら、取得済みの情報を先に使う（通信なし）。
                # ただし曲名を指定して検索した場合は作家名が空なので、その時は取得しに行く。
                if _cp_tjcd and _cp_tjcd == _cp_rjcd:
                    _cp_cred = {
                        k: _cp_res.get(k, "")
                        for k in ("作曲者", "作詞者", "編曲者", "訳詞者")
                        if _cp_res.get(k)
                    }
                if not _cp_cred:
                    if _cp_tjcd or _cp_tncd:
                        with st.spinner("作詞者・作曲者を取得中..."):
                            try:
                                _cp_cred = _get_mf_client().get_detail(
                                    f"jcd={_cp_tjcd}&ncd={_cp_tncd}&refer=music/list-product"
                                ) or {}
                            except MusicForestError as _cp_ce:
                                _cp_cred_msg = f"（作家名の取得に失敗: {_cp_ce}）"
                        if _cp_cred.get("error"):
                            _cp_cred_msg = f"（作家名の取得に失敗: {_cp_cred['error']}）"
                        elif not any(_cp_cred.get(k) for k in ("作曲者", "作詞者")):
                            _cp_cred_msg = (
                                f"（MINCの作品詳細に作家名がありませんでした"
                                f"／コード {_cp_tjcd or _cp_tncd}）"
                            )
                    elif _mf_norm_name(_cp_trk.get("曲名", "")) == _mf_norm_name(
                        _cp_res.get("作品名", "")
                    ):
                        # 作品コードが無い曲。検索した作品そのものなら、検索結果の
                        # 「作詞／作曲」列から拾った作家名を使える（同名判定なので
                        # CD内の別の無コード曲には適用しない）。
                        _cp_cred = {
                            k: _cp_res.get(k, "")
                            for k in ("作曲者", "作詞者", "編曲者", "訳詞者")
                            if _cp_res.get(k)
                        }
                        if not _cp_cred:
                            _cp_cred_msg = "（作品コードが無いため作家名は取得できません）"
                    else:
                        _cp_cred_msg = "（作品コードが無いため作家名は取得できません）"
                for _cp_ck in ("作曲者", "作詞者", "編曲者", "訳詞者"):
                    if (_cp_cred or {}).get(_cp_ck):
                        _cp_tapply[_cp_ck] = _cp_cred[_cp_ck]

                # I/V区分: MINCの収録曲表にI/V表記が無い場合は作詞者の有無で決める。
                # 作品詳細を引けた（＝作詞者がいないと確認できた）ときだけ判定する。
                if not _cp_tapply.get("I/V区分") and any(
                    (_cp_cred or {}).get(k) for k in ("作曲者", "作詞者", "編曲者", "訳詞者")
                ):
                    _cp_tapply["I/V区分"] = _infer_iv((_cp_cred or {}).get("作詞者", ""))

            # 邦洋区分（JASRACコード2文字目: 数字→邦楽、英字→洋楽）
            # 別の作品を反映するときは古い値が残ると誤りになるので上書きする。
            # 同じ作品を反映し直すときは、手で直した値を潰さないよう空欄のときだけ補う。
            _cp_hy_src = (
                re.sub(r"[-\s]", "", _cp_trk.get("JASRAC作品コード", ""))
                or re.sub(r"[-\s]", "", str(_cp_res.get("作品コード", "")))
            ).upper()
            _cp_hy = _infer_houyo(_cp_hy_src)
            if _cp_hy and (
                not _cp_current_houyo(row_idx) or _cp_hy_src != _cp_row_jcd(row_idx)
            ):
                _cp_tapply["邦洋区分"] = _cp_hy

            for _cp_col, _cp_val in _cp_tapply.items():
                if _cp_val and _cp_col in st.session_state.songs_df.columns:
                    st.session_state.songs_df.at[row_idx, _cp_col] = _cp_val
            st.session_state["_apply_msg"] = (
                f"{_cp_trk.get('曲順', '')}曲目「{_cp_trk.get('曲名', '')}」"
                f"（{_cp_trk.get('収録時間', '')}）とCD情報を反映しました。"
                + (f" 作曲: {_cp_tapply.get('作曲者', '')}" if _cp_tapply.get("作曲者") else "")
                + _cp_cred_msg
            )
            st.session_state.pop("songs_editor", None)
            st.rerun()

    # ── 参照したMINCのURL ───────────────────────────────────────────────
    _cp_urls = []
    if _cp_res.get("search_url"):
        _cp_urls.append(f"[CD商品リスト]({_cp_res['search_url']})")
    if _cp_item.get("detail_url"):
        _cp_urls.append(f"[選択中のCDの商品詳細]({_cp_item['detail_url']})")
    if _cp_urls:
        st.caption("🔗 MINC: " + "　／　".join(_cp_urls))


# DB 読み込み時に不足する可能性がある列とそのデフォルト値
_SONG_DEFAULTS: dict[str, str] = {
    "使用形態": "背景",
    "音源区分": "CD",
    "I/V区分": "",
    "邦洋区分": "",
    "原訳詞区分": "",
    "編曲者": "",
    "訳詞者": "",
    "レコード会社名": "",
    "CD名": "",
    "委任者": "",
    "放送": "",
    "配信": "",
    # 出どころ別の生の値（表には出さない）
    "放送J": "",
    "放送N": "",
    "配信J": "",
    "配信N": "",
    "自社楽曲ID": "",
}


def _ensure_song_defaults(df: pd.DataFrame) -> pd.DataFrame:
    """DB から読み込んだ songs_df に新規列が不足していれば既定値で補完する。"""
    for col, default in _SONG_DEFAULTS.items():
        if col not in df.columns:
            df[col] = default
    _migrate_mgmt_sources(df)
    return df


def _migrate_mgmt_sources(df: pd.DataFrame) -> None:
    """先に配った「放送」「配信」を、出どころ別の欄へ写す。

    最初は J-WID の値しか入れていなかったので、出どころ別の欄がどちらも
    空で表示欄だけ埋まっている行は、その値を JASRAC 由来として写す。
    もう写してある行は触らない。
    """
    blank = ("", "nan", "None")
    for col, (jw_col, nt_col) in MGMT_SOURCE_COLS.items():
        if not {col, jw_col, nt_col} <= set(df.columns):
            continue
        cur = df[col].astype(str).str.strip()
        need = (df[jw_col].astype(str).str.strip().isin(blank)
                & df[nt_col].astype(str).str.strip().isin(blank)
                & ~cur.isin(blank))
        if need.any():
            df.loc[need, jw_col] = cur[need]


# J-WID 管理状況カテゴリ定義（利用分野の表示用）
_JWID_MGMT_GROUPS = [
    ("演奏",   ["演奏会等", "上映/BGM", "社交場/ｶﾗｵｹ"]),
    ("複製",   ["録音", "出版", "貸与", "ビデオ", "映画"]),
    ("複合",   ["放送", "配信", "通カラ"]),
    ("広告",   ["広告/CM送録", "広告/映録", "広告/録音", "広告/ビデオ", "広告/出版"]),
    ("ゲーム", ["ゲーム/録音", "ゲーム/ビデオ"]),
]


def _format_management_status(mgmt: dict) -> str:
    """管理状況dict → カテゴリ別マークダウン文字列。"""
    lines = []
    for cat, fields in _JWID_MGMT_GROUPS:
        parts = []
        for f in fields:
            icon = mgmt.get(f, "?")
            short = f.split("/")[-1]
            parts.append(f"{short}:{icon}")
        lines.append(f"**{cat}**: {', '.join(parts)}")
    return "  \n".join(lines)


#: 管理の強さ。○（管理あり）＞△（一部管理）＞×（管理なし）
_MGMT_RANK = {"○": 2, "△": 1, "×": 0}


def _combine_mgmt(jw: str, nt: str) -> str:
    """JASRAC と NexTone の管理状況を、表に出す1つの値にまとめる。

    どちらかが ○ なら使える曲なので ○ にする。NexTone 側を採ったときだけ
    「(N)」を付けて出どころを残す（JASRAC で引くと × に見えるのに実際は
    NexTone が管理している、というのが分かるように）。
    空文字はまだ引いていないという意味。
    """
    jw = _mgmt_val(jw)
    nt = _mgmt_val(nt)
    if not nt:
        return jw
    if not jw:
        return f"{nt}(N)"
    if _MGMT_RANK.get(nt, -1) > _MGMT_RANK.get(jw, -1):
        return f"{nt}(N)"
    return jw


def _mgmt_val(v) -> str:
    """管理状況の1マスを文字列にする。空・nan・「?」は空文字にそろえる。"""
    s = str(v or "").strip()
    return "" if s in ("nan", "None", "?") else s


def _apply_management_status(mgmt: dict, updates: dict,
                             source: str = "J", row=None) -> int:
    """管理状況から「放送」「配信」を updates に入れる。入れた数を返す。

    J-WID／NexTone を1曲ずつ開いて見に行っていた欄なので、当たった時点で
    書き取っておく。生の値は出どころ別の欄（放送J・放送N…）に置き、表に
    出す欄はもう片方の出どころも見て作る。

    source は "J"（JASRAC／J-WID）か "N"（NexTone）。row には今の行を
    渡す（Series でも dict でも可）。もう片方の出どころに何が入っている
    かを見るために使う。updates に既に入っていればそちらを優先する。
    """
    if not mgmt:
        return 0
    is_nt = str(source).upper().startswith("N")
    n = 0
    for col in JWID_MGMT_COLS:
        val = _mgmt_val(mgmt.get(col, ""))
        if not val:
            continue
        jw_col, nt_col = MGMT_SOURCE_COLS[col]
        own_col, other_col = (nt_col, jw_col) if is_nt else (jw_col, nt_col)
        updates[own_col] = val
        other = updates.get(other_col)
        if other is None and row is not None:
            try:
                other = row[other_col]
            except (KeyError, IndexError, TypeError):
                other = ""
        other = _mgmt_val(other)
        updates[col] = (_combine_mgmt(other, val) if is_nt
                        else _combine_mgmt(val, other))
        n += 1
    return n


def _write_mgmt_from_fetch(songs: pd.DataFrame, idx, mgmt: dict,
                           source: str = "J") -> int:
    """引いてきた管理状況を songs_df の1行に書き込む。書いた欄の数を返す。

    同じ出どころの値が既に入っていれば引き直さない（同じ番号を引いた結果
    なので上書きしても変わらない）。表に出す欄は、もう片方の出どころも
    見て作り直す。
    """
    if not mgmt:
        return 0
    is_nt = str(source).upper().startswith("N")
    keep = {}
    for col in JWID_MGMT_COLS:
        own_col = MGMT_SOURCE_COLS[col][1 if is_nt else 0]
        if own_col in songs.columns and _mgmt_val(songs.at[idx, own_col]):
            continue  # もう入っている
        if _mgmt_val(mgmt.get(col, "")):
            keep[col] = mgmt[col]
    if not keep:
        return 0
    updates: dict = {}
    _apply_management_status(keep, updates, source, songs.loc[idx])
    n = 0
    for col, val in updates.items():
        if col not in songs.columns:
            continue
        if str(songs.at[idx, col]) != val:
            songs.at[idx, col] = val
            if col in _MGMT_RAW_COLS:   # 数えるのは中身の欄だけ
                n += 1
    return n


def _infer_houyo(jasrac: str) -> str:
    """JASRACコードの2文字目から邦洋区分を推定する（数字→邦楽、英字→洋楽）。"""
    if len(jasrac) >= 2:
        c = jasrac[1]
        if c.isdigit():
            return "邦楽"
        if c.isalpha():
            return "洋楽"
    return ""


def _is_blank(v) -> bool:
    """空欄扱いの値（空文字 / nan / none）かどうか。"""
    return str(v).strip().lower() in ("", "nan", "none")


def _mf_norm_name(s: str) -> str:
    """アーティスト名照合用の正規化（NFKC・小文字化・空白除去）。"""
    if _is_blank(s):
        return ""
    return re.sub(r"[\s　]", "", unicodedata.normalize("NFKC", str(s))).lower()


def _composer_matches(hint: str, candidate: str) -> bool:
    """作曲者名が一致するか（表記ゆれを吸収した照合）。

    MP3のID3タグは「中西圭三」、MINCは「中西　圭三」のように表記が揺れるため、
    _mf_norm_name（NFKC＋空白除去＋小文字化）で正規化してから比べる。
    複数名は "/" 等の区切りで分割し、1人でも共通すれば一致とみなす。

    注意: ローマ字表記のタグ（"Keizo Nakanishi"）は一致しない。不一致は
    「別人」ではなく「決め手が無い」という意味なので、不一致を理由に
    候補を捨ててはいけない（呼び出し側は優先度付けにのみ使うこと）。
    """
    if _is_blank(hint) or _is_blank(candidate):
        return False
    _split = lambda s: {
        _mf_norm_name(x) for x in re.split(r"[/／・、,]", str(s)) if _mf_norm_name(x)
    }
    _h, _c = _split(hint), _split(candidate)
    if not _h or not _c:
        return False
    if _h & _c:
        return True
    # 姓のみ等の部分表記も拾う。短い名前での誤一致を避けるため3文字以上に限る。
    return any(
        (a in b or b in a) and len(a) >= 3 and len(b) >= 3
        for a in _h for b in _c
    )


#: アーティスト欄に入る「特定の実演家を指していない」表記。照合に使ってはいけない。
_ARTIST_GENERIC = {
    "va", "v.a.", "(v.a.)", "variousartists", "various", "オムニバス",
    "オリジナルサウンドトラック", "サウンドトラック", "unknownartist", "unknown",
}


def _artist_matches(hint: str, candidate: str) -> bool:
    """アーティスト（実演家）名が一致するか。

    ID3タグは「ZOO」「EXILE TRIBE feat. ...」のように付随表記が付くため、
    feat./with 以降と末尾の括弧書きを落としてから _composer_matches の
    照合ロジック（NFKC＋空白除去＋部分一致）に渡す。

    注意: 作曲者と違い、アーティストは作品を一意に決めない。同じ作品の
    カバー音源は作曲者が同じでアーティストだけが違う（実測: 心の旅
    jcd=03175928 に「チューリップ」と「前田憲男」の2行）。よって
    「どの作品か」ではなく「どのCD／音源か」の決め手として使うこと。
    """
    def _clean(s: str) -> str:
        s = re.sub(r"(?i)\s*(feat\.?|featuring|with)\s.+$", "", str(s))
        s = re.sub(r"[（(\[].*?[)）\]]\s*$", "", s).strip()
        return s

    _h, _c = _clean(hint), _clean(candidate)
    if _mf_norm_name(_h) in _ARTIST_GENERIC or _mf_norm_name(_c) in _ARTIST_GENERIC:
        return False
    return _composer_matches(_h, _c)


def _infer_iv(lyricist: str) -> str:
    """作詞者の有無から I/V区分を推定する（作詞者あり→ヴォーカル、なし→インスト）。

    作家名を取得できた（＝作詞者が本当にいないと分かる）場合にのみ使うこと。
    未取得の空欄をインストと決めつけないため、呼び出し側で判断する。
    """
    return "インスト" if _is_blank(lyricist) else "ヴォーカル"


def _apply_iv_from_credits(apply: dict) -> None:
    """反映内容（apply dict）に作家名が含まれていれば I/V区分 を決めて書き足す。

    その作品の作家名を取得できたということは作詞者の有無が確定しているので、
    反映する作品の値として上書きする（既に I/V区分 が入っている apply は触らない）。
    """
    if apply.get("I/V区分"):
        return
    if not any(not _is_blank(apply.get(k, "")) for k in ("作曲者", "作詞者", "編曲者", "訳詞者")):
        return
    apply["I/V区分"] = _infer_iv(apply.get("作詞者", ""))


def _get_mf_client() -> MusicForestClient:
    """session_state にキャッシュされた MusicForestClient を返す。
    同一セッション内で同じインスタンスを使い回すことで、サーバー側のセッション
    ローテーション（新クッキー）を引き継いだままにし、ログイン状態が切れるのを防ぐ。
    認証確認ボタン押下時は st.session_state['mf_client'] を削除してから呼ぶこと。
    """
    _c = st.session_state.get("mf_client")
    # modules/musicforest.py を編集するとモジュールが再読み込みされ、クラスオブジェクト
    # が差し替わる。session_state に残った旧クラスのインスタンスは古いメソッド定義を
    # 持ったままなので（TypeError: unexpected keyword argument の原因）、
    # 型が一致しない場合は作り直す。
    if not isinstance(_c, MusicForestClient):
        _c = load_client()
        st.session_state["mf_client"] = _c
    return _c


# =====================================================================
# セッション状態初期化
# =====================================================================
def _init_session() -> None:
    defaults: dict = {
        "cue_df": None,
        "wav_df": None,
        "mp3_df": None,
        "mp3_is_finder": False,
        "master_db_df": None,
        "songs_df": None,
        "events_df": None,
        "search_df": None,
        "project_id": None,
        "project_name": "",
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


_init_session()
init_db()

# =====================================================================
# Chrome 拡張機能からの MINC セッション同期（クエリパラメータ受信）
# =====================================================================
_sync_param = st.query_params.get("sync_minc", "")
if _sync_param:
    import base64 as _b64
    try:
        _cookie_list = json.loads(_b64.b64decode(_sync_param.encode()).decode("utf-8"))
        _cookie_dicts = [
            {
                "name":     c["name"],
                "value":    c["value"],
                "domain":   c.get("domain", "www.minc.or.jp"),
                "path":     c.get("path", "/"),
                "expires":  c.get("expires", -1),
                "httpOnly": False,
                "secure":   c.get("secure", True),
            }
            for c in _cookie_list if c.get("value")
        ]
        from modules.musicforest import _save_cookies_to_state, get_state_path as _gsp
        _cnt, _sess_ok = _save_cookies_to_state(_cookie_dicts, _gsp())
        st.session_state.pop("mf_client", None)
        st.session_state.pop("mf_auth_state", None)
        st.session_state["_ext_sync_msg"] = (
            f"Chrome 拡張機能から {_cnt} 件の Cookie を同期しました。"
            + (" _sess あり ✅" if _sess_ok else " _sess なし（要ログイン）")
        )
    except Exception as _sync_e:
        st.session_state["_ext_sync_msg"] = f"同期エラー: {_sync_e}"
    st.query_params.clear()

# =====================================================================
# ログイン（APP_USERS 未設定なら素通り）
# =====================================================================
# Cookie 同期より後に置くこと。Chrome 拡張機能は裏で新しいセッションを
# 開くので、先に止めると同期が届かなくなる。
CURRENT_USER = require_login()

# =====================================================================
# nuendo_mp3_finder CSV ヘルパー
# =====================================================================
_MP3FINDER_KEY_COLS = {"イベント名", "ファイル名"}
_MP3FINDER_MATCH_COLS = {"一致種別", "照合種別", "マッチ種別"}
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
    _match_col_candidates = ["一致種別", "照合種別", "マッチ種別"]
    match_col = next((c for c in _match_col_candidates if c in mp3finder_df.columns), None)
    if match_col is None:
        return songs_df, 0
    PRIORITY = {
        "完全一致": 5, "正規化一致": 4, "管理番号一致": 3,
        "タイトル一致": 2, "部分一致": 1, "該当なし": 0,
    }
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
            _fill("アルバム(ID3)",     "CD名"),
            _fill("ファイル名",        "MP3一致ファイル名"),
            _fill("再生時間",          "MP3フル尺"),
        ])
        if changed:
            updated += 1

    return songs_df, updated

def _import_master_db(
    master_df: pd.DataFrame,
    songs_df: pd.DataFrame,
) -> tuple[pd.DataFrame, int]:
    """
    マスターDB CSV の情報を songs_df に補完する。
    空フィールドのみ上書き（既存データは保持）。
    曲名で正規化マッチング → 一致したレコードのフィールドを補完する。
    """
    # 曲名列を検出
    title_col = next((c for c in ["曲名", "タイトル", "作品名"] if c in master_df.columns), None)
    if title_col is None:
        return songs_df, 0

    # マスターDB ルックアップを構築
    master_lookup: dict[str, dict] = {}
    for _, mrow in master_df.iterrows():
        raw = str(mrow.get(title_col, "")).strip()
        if not raw or raw.lower() == "nan":
            continue
        master_lookup[raw.lower()] = mrow.to_dict()
        norm = normalize_for_match(raw)
        if norm:
            master_lookup[norm] = mrow.to_dict()

    # songs_df の列名 → master_df の可能な列名（優先順）
    field_map: dict[str, list[str]] = {
        "JASRAC作品コード": ["JASRAC作品コード", "作品コード", "JASRAC"],
        "NexTone管理番号":  ["NexTone管理番号", "NexTone管理コード", "NexTone"],
        "作曲者":           ["作曲者", "Composer"],
        "作詞者":           ["作詞者", "Lyricist"],
        "アーティスト":     ["アーティスト", "アーティスト名", "Artist"],
        "CD番号":           ["CD番号", "品番", "CatalogNo"],
    }

    updated = 0
    for idx in songs_df.index:
        song_title = str(songs_df.at[idx, "曲名"]).strip()
        if not song_title or song_title.lower() == "nan":
            continue

        norm = normalize_for_match(song_title)
        master_row = master_lookup.get(norm) or master_lookup.get(song_title.lower())
        if master_row is None:
            continue

        changed = False
        for dst_col, src_candidates in field_map.items():
            if dst_col not in songs_df.columns:
                continue
            cur = str(songs_df.at[idx, dst_col]).strip()
            if cur and cur.lower() != "nan":
                continue  # 既存データは上書きしない
            for src_col in src_candidates:
                val = str(master_row.get(src_col, "")).strip()
                if val and val.lower() != "nan":
                    songs_df.at[idx, dst_col] = val
                    changed = True
                    break

        if changed:
            updated += 1

    return songs_df, updated


def _default_project_name() -> str:
    """Cue CSV のファイル名から案件名を作る。"""
    _name = str(st.session_state.get("cue_file_name") or "").strip()
    _stem = os.path.splitext(_name)[0].strip()
    return _stem[:120] if _stem else "無題の案件"


def _autosave_to_db(note: str = "") -> bool:
    """今の内容をDBに保存する。案件が無ければ何もしない。

    ダウンロードボタンの on_click から呼ぶ。照合の後の作業（一括検索の
    結果や手入力の修正）は💾ボタンでしか保存されないので、書き出した
    ファイルとDBの中身がずれないよう、書き出す瞬間にも保存しておく。

    コールバックの中で st.success を出しても表示前に再実行が走って
    消えるため、伝言を session_state に置いて本体側で出す。
    """
    _pid = st.session_state.get("project_id")
    if not _pid or st.session_state.get("songs_df") is None:
        return False
    try:
        save_songs(_pid, st.session_state.songs_df)
        if st.session_state.get("events_df") is not None:
            save_events(_pid, st.session_state.events_df)
        st.session_state["_autosave_msg"] = (
            f"💾 「{st.session_state.get('project_name', '')}」に保存しました{note}"
        )
        return True
    except Exception as _e:
        st.session_state["_autosave_msg"] = f"⚠️ 自動保存に失敗しました: {_e}"
        return False


# 件数を数えるだけの問い合わせを、しばらく使い回す間隔（秒）
_COUNT_TTL = 60.0
# 自社CDの台帳は取り込みスクリプトでしか変わらないので、長めでよい
_CD_COUNT_TTL = 600.0


def _cached_count(key: str, fn, ttl: float = _COUNT_TTL) -> int:
    """件数を数えるだけの問い合わせを、しばらく使い回す。

    この画面はボタンを押すたびに丸ごと動き直す作りなので、素直に数えると
    1クリックごとにクラウドDBへ往復する（実測 175ミリ秒。しかも上の
    ステータスバーと管理タブで2回）。件数は目安として出しているだけで、
    数秒古くても困らない。中身を書き換えたときは _forget_count で捨てる。

    失敗したときは、前に数えた値があればそれを返す（0 に化けて「まだ何も
    無い」と表示されるのを避けるため）。
    """
    box = st.session_state.setdefault("_count_cache", {})
    hit = box.get(key)
    now = time.monotonic()
    if hit and (now - hit[1]) < ttl:
        return hit[0]
    try:
        n = int(fn())
    except Exception:
        n = hit[0] if hit else 0
    box[key] = (n, now)
    return n


def _forget_count(key: str) -> None:
    """数え直しが要るときに、覚えている件数を捨てる。"""
    (st.session_state.get("_count_cache") or {}).pop(key, None)


def _master_learn() -> int:
    """調べ終わった曲を共有楽曲データに貯める。貯めた曲数を返す。

    貯めるのは「確定」「作曲者一致」「アーティスト一致」の行だけ。
    要確認や未調査まで貯めると、間違いが全案件に広がってしまう。
    失敗しても本来の作業は止めない（貯めるのはおまけなので）。
    """
    _df = st.session_state.get("songs_df")
    if _df is None or _df.empty:
        return 0
    try:
        _n = master_learn(_df, CURRENT_USER)
        if _n:
            _forget_count("master")   # 増えた分をすぐ出す
        return _n
    except Exception as _e:
        st.warning(f"⚠️ 共有楽曲データへの保存に失敗しました: {_e}")
        return 0


# =====================================================================
# ヘッダー
# =====================================================================
def _minc_state() -> tuple[str, str]:
    """MINC ログインの状態を (表示文字, 調子) で返す。

    check_session() は MINC に実際に接続しに行くので、毎回の再実行で
    呼ぶわけにはいかない（画面が数秒止まる）。ここは state.json を見る
    だけの判定にとどめ、本当に生きているかの確認は従来どおり各タブの
    ボタンに任せる。
    """
    try:
        _p = get_state_path()
        # DB に新しい Cookie があれば取り寄せる。ここは毎回の再実行で
        # 通るので、間隔を空けて聞く方（force なし）を使う
        from modules.musicforest import sync_state_from_db
        sync_state_from_db(_p)
        if not _p.is_file():
            return ("未接続", "off")
        _has_sess = '"_sess"' in _p.read_text(encoding="utf-8")
        from modules.musicforest import _session_age_str
        _age = _session_age_str(_p)
        if not _has_sess:
            return ("要ログイン", "warn")
        return (f"接続済み（{_age}）" if _age else "接続済み", "ok")
    except Exception:
        return ("不明", "off")


_bar_songs = st.session_state.get("songs_df")
_minc_txt, _minc_tone = _minc_state()
# describe_backend() は SQLite のフルパスまで返して長いので、バーには
# どちらに繋がっているかだけ出す（接続文字列は秘密なので元から出ない）
_db_txt = ("クラウド（Supabase）" if describe_backend().startswith("PostgreSQL")
           else "ローカル")
# 件数を数えるだけのクエリ。軽そうに見えてクラウドDBへの往復なので、
# 数えた値をしばらく使い回す
_master_n = _cached_count("master", master_count)

status_bar(
    "🎵 著作権調査支援ツール",
    [
        # ログインしていないときは項目ごと出さない（status_bar が空を飛ばす）
        ("ログイン中", CURRENT_USER if auth_enabled() else "", "ok"),
        # 案件が無いとDBに保存されないので、未選択は「注意」の色で出す
        ("案件", st.session_state.get("project_name") or "未選択（未保存）",
         "" if st.session_state.get("project_id") else "warn"),
        ("保存先", _db_txt, ""),
        ("MINC", _minc_txt, _minc_tone),
        # 貯まっている共有楽曲データの件数。増えていくのが見えるように出す
        ("共有楽曲", f"{_master_n} 曲" if _master_n else "", ""),
    ],
    progress=count_done(_bar_songs),
)

# ダウンロード時の自動保存はコールバックの中で走るので、結果はここで出す
if "_autosave_msg" in st.session_state:
    st.toast(st.session_state.pop("_autosave_msg"))

# ログアウトは滅多に押さないので、本文を狭めないようサイドバーへ
if auth_enabled():
    with st.sidebar:
        st.markdown(f"ログイン中: **{CURRENT_USER}**")
        logout_button()

tabs = st.tabs(
    [
        "📋 申告フォーム作成",
        "📊 イベント一覧",
        "🗃️ 共有楽曲データ",
    ]
)


# =====================================================================
# ⚙️ ファイル読み込み・設定（tabs[0] 先頭）
# =====================================================================
with tabs[0]:
    st.header("⚙️ ファイル読み込み・設定")

    # ── プロジェクト管理 ──────────────────────────────────
    with st.expander(
        "🗄️ プロジェクト管理（DB保存・読み込み・マスターDB）",
        # マスターDBも入れて縦に長くなったので、常に畳んでおく。
        # 未選択のときは下に警告を出すので、開きっぱなしにしなくてよい。
        expanded=False,
    ):
        # 自分の案件 ＋ 所有者が空の案件（分ける前に作られたもの）だけ出す
        _projects = list_projects(CURRENT_USER)

        _pm_col1, _pm_col2 = st.columns(2)

        # 新規作成
        with _pm_col1:
            st.markdown("**新規プロジェクト**")
            _new_name = st.text_input(
                "プロジェクト名",
                placeholder="例: ◯◯劇場版 BGM調査",
                key="new_project_name",
            )
            _new_desc = st.text_input(
                "説明（任意）",
                placeholder="放送日・依頼元など",
                key="new_project_desc",
            )
            if st.button("➕ 新規作成", key="btn_create_project", use_container_width=True):
                if _new_name.strip():
                    _pid = create_project(_new_name.strip(), _new_desc.strip(),
                                          owner=CURRENT_USER)
                    st.session_state.project_id = _pid
                    st.session_state.project_name = _new_name.strip()
                    st.success(f"✅ プロジェクト「{_new_name.strip()}」を作成しました（ID: {_pid}）")
                    st.rerun()
                else:
                    st.warning("プロジェクト名を入力してください。")

        # 既存読み込み
        with _pm_col2:
            st.markdown("**既存プロジェクトを読み込む**")
            if _projects:
                # 所有者が空のものは印を付ける。読み込むと自分のものになる
                _proj_labels = [
                    f"[{p['id']}] {p['name']}  （{p['updated_at'][:10]}）"
                    + ("　🔓 所有者なし" if not p.get("owner") else "")
                    for p in _projects
                ]
                _sel_label = st.selectbox(
                    "プロジェクトを選択",
                    options=_proj_labels,
                    key="select_existing_project",
                )
                _sel_proj = _projects[_proj_labels.index(_sel_label)]

                _load_col, _del_col = st.columns([3, 1])
                with _load_col:
                    if st.button("📂 読み込む", key="btn_load_project", use_container_width=True):
                        _pid = _sel_proj["id"]
                        _loaded_songs = load_songs(_pid)
                        _loaded_events = load_events(_pid)
                        if _loaded_songs is not None and len(_loaded_songs) > 0:
                            st.session_state.songs_df = _ensure_song_defaults(_loaded_songs)
                            st.session_state.events_df = _loaded_events
                            st.session_state.project_id = _pid
                            st.session_state.project_name = _sel_proj["name"]
                            # 所有者が空のものは、読み込んだ人のものにする。
                            # これで所有者を分ける前の案件が自然に片付き、
                            # 移行用のスクリプトを別に用意しなくて済む。
                            if not _sel_proj.get("owner"):
                                set_project_owner(_pid, CURRENT_USER)
                            st.success(
                                f"✅ 「{_sel_proj['name']}」を読み込みました"
                                f"（楽曲 {len(_loaded_songs)} 件）"
                            )
                            st.rerun()
                        else:
                            st.warning("このプロジェクトにはまだデータが保存されていません。")
                with _del_col:
                    if st.button("🗑️", key="btn_del_project", help="プロジェクトを削除", use_container_width=True):
                        delete_project(_sel_proj["id"])
                        if st.session_state.project_id == _sel_proj["id"]:
                            st.session_state.project_id = None
                            st.session_state.project_name = ""
                        st.success(f"削除しました: {_sel_proj['name']}")
                        st.rerun()
            else:
                st.info("まだプロジェクトがありません。先に新規作成してください。")

        # 現在のプロジェクト表示 ＋ 保存ボタン
        st.divider()
        if st.session_state.project_id:
            _save_col, _info_col = st.columns([2, 3])
            with _info_col:
                st.markdown(
                    f"**現在:** 🗂️ {st.session_state.project_name}"
                    f"　（ID: {st.session_state.project_id}）"
                )
            with _save_col:
                if st.button(
                    "💾 楽曲まとめをDBに保存",
                    key="btn_save_db",
                    type="primary",
                    use_container_width=True,
                    disabled=st.session_state.songs_df is None,
                ):
                    save_songs(st.session_state.project_id, st.session_state.songs_df)
                    if st.session_state.events_df is not None:
                        save_events(st.session_state.project_id, st.session_state.events_df)
                    n = len(st.session_state.songs_df)
                    st.success(f"✅ {n} 件を保存しました。")
                    # 手で直して「確定」にしたものが一番強い出典なので、
                    # ここで貯めるのが共有データにとって一番価値が高い
                    _learned = _master_learn()
                    if _learned:
                        st.info(
                            f"🗃️ 共有楽曲データに {_learned} 曲を貯めました。"
                        )
        else:
            st.info("プロジェクトを作成または読み込むと、DBへの保存が有効になります。")

        st.divider()

        # マスターDB CSV。プロジェクトの読み書きと同じ「既にある情報を
        # 持ってくる」操作なので、ファイル読み込みではなくここに置く。
        st.markdown("**📚 マスターDB CSV（任意）**")
        st.caption(
            "Access DB などからエクスポートした曲情報 CSV を読み込みます。"
            "曲名で照合し、JASRAC作品コード・NexTone管理番号・作曲者などを空フィールドに補完します。"
        )
        st.caption("列名の例: `曲名`, `JASRAC作品コード`, `NexTone管理番号`, `作曲者`, `作詞者`, `アーティスト`, `CD番号`")
        master_file = st.file_uploader(
            "マスターDB CSV を選択（任意）", type=["csv"], key="upload_master_db"
        )
        if master_file:
            try:
                df, enc = read_csv_auto(master_file)
                st.session_state.master_db_df = df
                _detected_title = next((c for c in ["曲名", "タイトル", "作品名"] if c in df.columns), None)
                if _detected_title:
                    st.success(f"✅ 読み込み完了（{enc}）: {len(df)} 件  曲名列: 「{_detected_title}」")
                else:
                    st.warning(f"⚠️ 読み込みましたが「曲名」「タイトル」「作品名」列が見つかりません。")
                with st.expander("プレビュー（先頭 5 行）"):
                    st.dataframe(df.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"❌ 読み込みエラー: {e}")

        if st.session_state.master_db_df is not None:
            _mdb = st.session_state.master_db_df
            st.info(f"マスターDB: {len(_mdb)} 件読み込み済み")
            if st.session_state.songs_df is not None:
                if st.button(
                    "📚 マスターDBから楽曲まとめに補完",
                    key="import_master_db_btn",
                    type="primary",
                    use_container_width=True,
                ):
                    _new_songs, _n_upd = _import_master_db(
                        _mdb, st.session_state.songs_df.copy()
                    )
                    st.session_state.songs_df = _new_songs
                    st.success(f"✅ {_n_upd} 件を補完しました。「楽曲まとめ」タブで確認してください。")
            else:
                st.info("「照合実行」後に補完ボタンが表示されます。")

    # DBへの保存は project_id がある時だけ走る（照合後の自動保存も💾ボタンも）。
    # 案件を作らずに作業すると再読み込みで全部消えるので、畳んだ状態でも
    # 気付けるようにここで知らせる。
    if not st.session_state.project_id:
        st.info(
            "🗄️ 案件が未選択です。このまま「🔄 照合実行」すると、"
            "**Cue CSV のファイル名で案件を自動作成**して保存します。"
            "続きをやる場合や名前を決めたい場合は、上の「🗄️ プロジェクト管理」から"
            "作成・読み込みしてください。"
        )

    st.divider()

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
                    # 案件を自動で作るときの名前に使う
                    st.session_state.cue_file_name = cue_file.name

                    # 基準信号（1KHZ-20DB-04 など）は楽曲ではないので落とす
                    df, _dropped_ev = drop_non_music_events(df)

                    # NUENDO の書き出しには「構成」「ノートパッド」のような
                    # 音楽以外のトラックが混ざる。既定で外しておくが、
                    # 選択欄には残して戻せるようにする（黙って消さない）。
                    _music, _other = music_tracks(df)
                    _tracks = _music + _other
                    if len(_tracks) > 1:
                        _use = st.multiselect(
                            "使うトラック",
                            options=_tracks,
                            default=_music,
                            key="cue_tracks",
                            help="構成・ノートパッド・マーカー・ビデオの各"
                                 "トラックは楽曲ではないので既定で外しています。"
                                 "必要なら選び直せます。",
                        )
                        if _use:
                            df = df[df["トラック名"].isin(_use)].reset_index(
                                drop=True)

                    st.session_state.cue_df = df
                    st.success(f"✅ 読み込み完了（{enc}）: {len(df)} 件")
                    if _tracks:
                        st.caption("使用中のトラック: " + " / ".join(
                            f"{t}（{int((df['トラック名'] == t).sum())}件）"
                            for t in _tracks
                            if (df["トラック名"] == t).any()
                        ))
                    _skipped = []
                    if _other:
                        _skipped.append("音楽以外のトラック "
                                        + "・".join(_other))
                    if _dropped_ev:
                        _skipped.append(f"基準信号など {_dropped_ev} 件")
                    if _skipped:
                        st.caption("除外: " + " ／ ".join(_skipped))
                    with st.expander("プレビュー（先頭 5 行）"):
                        st.dataframe(df.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"❌ 読み込みエラー: {e}")

    # ---- ② MP3（フォルダスキャン または CSV アップロード） ----
    with col_right:
        st.subheader("② MP3 ファイル一覧")
        st.caption("WAV の補助、または WAV なしで作曲者・フル尺などを補完できます。")

        # CSV（nuendo_mp3_finder の出力）を使うことがほとんどなので先に置く。
        # 先頭のタブが初期表示になる。
        mp3_tab_csv, mp3_tab_scan = st.tabs(["📄 CSV をアップロード", "📂 フォルダをスキャン"])

        with mp3_tab_scan:
            if not IS_LOCAL_WINDOWS:
                # サーバーから利用者のパソコンの中は見えない。押しても
                # 「見つかりません」になるだけなので、先に理由を出す
                st.info(
                    "ℹ️ この画面はサーバーで動いているため、お使いのパソコンの"
                    "フォルダは読めません。「📄 CSV をアップロード」の方を"
                    "使ってください。"
                )
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
                        # 照合実行済みの場合は即座に ID3 自動補完を実行
                        if st.session_state.songs_df is not None:
                            _auto_songs, _auto_upd = _import_mp3finder_id3(
                                df,
                                st.session_state.songs_df.copy(),
                            )
                            st.session_state.songs_df = _auto_songs
                            if _auto_upd > 0:
                                st.success(
                                    f"💿 MP3 ID3タグ情報を {_auto_upd} 件自動補完しました。"
                                )
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

    st.divider()

    # ---- ③ WAV（フォルダスキャン または CSV アップロード） ----
    st.subheader("③ WAV ファイル一覧（任意）")

    wav_tab_scan, wav_tab_csv = st.tabs(["📂 フォルダをスキャン", "📄 CSV をアップロード"])

    with wav_tab_scan:
        st.caption("Audio フォルダのパスを貼り付けてスキャンします。PowerShell 不要です。")
        if not IS_LOCAL_WINDOWS:
            # サーバーから利用者のパソコンの中は見えない。押しても
            # 「見つかりません」になるだけなので、先に理由を出す
            st.info(
                "ℹ️ この画面はサーバーで動いているため、お使いのパソコンの"
                "フォルダは読めません。「📄 CSV をアップロード」の方を"
                "使ってください。"
            )
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
    can_match = st.session_state.cue_df is not None

    if not can_match:
        st.info("① Cue CSV を読み込んでから「照合実行」ボタンを押してください。")

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

        # 照合結果サマリーを表示
        _summary_cols = st.columns(2)
        if "WAV照合ステータス" in songs_df.columns and st.session_state.wav_df is not None:
            with _summary_cols[0]:
                st.caption("WAV照合")
                wav_counts = (
                    songs_df["WAV照合ステータス"].value_counts().reset_index()
                )
                wav_counts.columns = ["ステータス", "件数"]
                st.dataframe(wav_counts, use_container_width=True, hide_index=True)
        if st.session_state.get("mp3_is_finder") and st.session_state.mp3_df is not None:
            with _summary_cols[1]:
                st.caption("MP3照合")
                _match_col_cands = ["マッチ種別", "一致種別", "照合種別"]
                _mc = next((c for c in _match_col_cands if c in st.session_state.mp3_df.columns), None)
                if _mc:
                    mp3_counts = (
                        st.session_state.mp3_df[_mc].value_counts().reset_index()
                    )
                    mp3_counts.columns = ["マッチ種別", "件数"]
                    st.dataframe(mp3_counts, use_container_width=True, hide_index=True)

        st.info("「楽曲まとめ」タブで内容を確認・編集してください。")

        # 案件が無ければここで作る。データが生まれるのはこの瞬間なので、
        # ここで作らないと以降の自動保存がすべて素通りし、再読み込みで
        # 全部消えてしまう。
        if not st.session_state.project_id:
            _auto_name = _default_project_name()
            try:
                st.session_state.project_id = create_project(
                    _auto_name, "照合実行時に自動作成", owner=CURRENT_USER
                )
                st.session_state.project_name = _auto_name
                st.info(f"🗄️ 案件「{_auto_name}」を自動で作成しました。")
            except Exception as _e:
                st.warning(
                    f"⚠️ 案件の自動作成に失敗しました: {_e}\n\n"
                    "「🗄️ プロジェクト管理」から手動で作成してください。"
                )

        if st.session_state.get("mp3_is_finder") and st.session_state.mp3_df is not None:
            _auto_songs, _auto_upd = _import_mp3finder_id3(
                st.session_state.mp3_df,
                st.session_state.songs_df.copy(),
            )
            st.session_state.songs_df = _auto_songs
            if _auto_upd > 0:
                st.success(
                    f"💿 MP3 ID3タグ情報（アーティスト・作曲者・アルバム）を"
                    f" {_auto_upd} 件自動補完しました。"
                )

        # 共有楽曲データで空欄を埋める。誰かが前に調べた曲なら、ここで
        # もう入る。ID3 の補完より後に置くこと（先に置くと ID3 の値で
        # 埋まる欄まで共有データを見に行くことになり、無駄が出る）
        try:
            _m_df, _m_hits, _m_filled = master_fill(st.session_state.songs_df)
            if _m_filled:
                st.session_state.songs_df = _m_df
                st.success(
                    f"🗃️ 共有楽曲データから {_m_hits} 曲・{_m_filled} 欄を"
                    f"自動補完しました。"
                )
        except Exception as _e:
            st.warning(f"⚠️ 共有楽曲データの参照に失敗しました: {_e}")

        # 自社CDの台帳で、まだ空いている欄を埋める。共有楽曲データより
        # 後に置くこと。人が手で直した値のほうが強いので、先に入れさせる
        try:
            _c_df, _c_hits, _c_filled = cd_fill(st.session_state.songs_df)
            if _c_filled:
                st.session_state.songs_df = _c_df
                st.success(
                    f"💿 自社CDの台帳から {_c_hits} 曲・{_c_filled} 欄を"
                    f"自動補完しました。"
                )
        except Exception as _e:
            st.warning(f"⚠️ 自社CDの台帳の参照に失敗しました: {_e}")

        # 保存はここまで全部終えてから。以前は照合直後に保存していたので、
        # ID3 の補完結果が保存されずに消えていた
        if st.session_state.project_id:
            save_songs(st.session_state.project_id, st.session_state.songs_df)
            save_events(st.session_state.project_id, st.session_state.events_df)
            st.info(f"💾 プロジェクト「{st.session_state.project_name}」に自動保存しました。")


# =====================================================================
# 申告フォーマット + 楽曲まとめ（tabs[0] 2ブロック目）
# =====================================================================
with tabs[0]:
    st.divider()

    if st.session_state.songs_df is None:
        st.info("上の「⚙️ ファイル読み込み・設定」で Cue.csv を読み込んで照合を実行してください。")
    else:
        songs_df: pd.DataFrame = st.session_state.songs_df

        # ── 申告フォーマット ───────────────────────────────────
        st.subheader("📋 申告フォーマット")

        # ---- MINC ログイン設定 ----
        with st.expander("🔑 MINC ログイン設定", expanded=False):
            # Chrome 拡張機能からの同期結果表示
            if "_ext_sync_msg" in st.session_state:
                _msg = st.session_state.pop("_ext_sync_msg")
                if "エラー" in _msg:
                    st.error(_msg)
                else:
                    st.success(_msg)

            _mc1, _mc2, _mc3 = st.columns([3, 3, 2])
            with _mc1:
                _minc_mail_input = st.text_input(
                    "MINCメールアドレス",
                    key="minc_login_mail",
                    placeholder="example@email.com",
                    help="ユーザーごとに入力してください。空白の場合は .env の値を使用します。",
                )
            with _mc2:
                _minc_pass_input = st.text_input(
                    "MINCパスワード",
                    key="minc_login_pass",
                    type="password",
                    help="空白の場合は .env の値を使用します。",
                )
            with _mc3:
                st.write("")
                # ブラウザを開く仕組みなので、手元の Windows でしか動かない。
                # サーバーで押すと存在しない exe を起動しようとして落ちる
                if st.button(
                    "🔑 MINC ログイン",
                    use_container_width=True,
                    disabled=not IS_LOCAL_WINDOWS,
                    help=("Playwright でブラウザを開きます。ログイン後ブラウザを"
                          "閉じると Cookie が保存されます。"
                          if IS_LOCAL_WINDOWS else
                          "この機能は手元のパソコンでしか使えません。"
                          "サーバーにはブラウザがないためです。"),
                ):
                    import subprocess, re as _re
                    _login_py = r"H:\PROGRAM\search_music\src\login_browser.py"
                    _python   = r"H:\PROGRAM\search_music\.venv\Scripts\python.exe"
                    if not Path(_python).is_file():
                        st.error(
                            "❌ ログイン用のプログラムが見つかりません。\n\n"
                            f"探した場所: {_python}\n\n"
                            "この機能は手元のパソコン（search_music を置いて"
                            "あるところ）でだけ使えます。"
                        )
                        st.stop()
                    _env = os.environ.copy()
                    _mail_val = _minc_mail_input.strip()
                    _pass_val = _minc_pass_input.strip()
                    if _mail_val:
                        _env["MINC_MAIL_ADDRESS"] = _mail_val
                        _safe = _re.sub(r"[^\w@.-]", "_", _mail_val)
                        _env["MINC_STATE_PATH"]  = rf"H:\PROGRAM\search_music\auth\state_{_safe}.json"
                        _env["MINC_PROFILE_DIR"] = rf"H:\PROGRAM\search_music\auth\chrome-profile_{_safe}"
                        os.environ["MINC_STATE_PATH"] = _env["MINC_STATE_PATH"]
                    if _pass_val:
                        _env["MINC_PASSWORD"] = _pass_val
                    st.info("ブラウザが開きます。ログイン後ブラウザを閉じてください…")
                    subprocess.run([_python, _login_py], env=_env, check=False)
                    if "mf_auth_state" in st.session_state:
                        del st.session_state["mf_auth_state"]
                    st.rerun()

            st.divider()

            # ---- Chrome セッション同期 ----
            st.caption("**ブラウザと同じセッションを使う（二重ログアウト防止）**")
            st.caption(
                "Chrome で minc.or.jp にログイン済みの場合、そのセッションをスクレイパーに同期することで"
                "「同一デバイスからの二重アクセス」によるログアウトを防げます。"
            )

            # 方法 A: Chrome 拡張機能（最も簡単・Chrome 起動中でも動作）
            with st.expander("🧩 方法 A（推奨）: Chrome 拡張機能で同期（1クリック）", expanded=True):
                st.markdown(
                    "**初回のみ: 拡張機能をインストール**\n"
                    "1. Chrome で `chrome://extensions` を開く\n"
                    "2. 右上の「デベロッパーモード」をオン\n"
                    "3. 「パッケージ化されていない拡張機能を読み込む」をクリック\n"
                    f"4. `H:\\PROGRAM\\CyosakukenJIdouka_app\\chrome-extension` フォルダを選択\n\n"
                    "**同期するとき:**\n"
                    "1. Chrome で minc.or.jp にログインした状態で、右上の拡張機能アイコン（🧩）をクリック\n"
                    "2. 「MINC Session Sync」→「著作権アプリに同期する」を押す\n"
                    "3. アプリのタブが開いたら同期完了 ✅"
                )

            # 方法 B: Chrome を閉じて自動同期
            _sync_col1, _sync_col2 = st.columns([2, 2])
            with _sync_col1:
                with st.expander("方法 B: Chrome を閉じて自動同期"):
                    st.caption("Chrome をすべて終了してから押してください。")
                    # 手元の Chrome から Cookie を読む仕組みなので、
                    # サーバーでは読む相手がいない
                    if st.button("🔗 Chromeから自動同期", use_container_width=True,
                                 key="minc_chrome_sync",
                                 disabled=not IS_LOCAL_WINDOWS,
                                 help=None if IS_LOCAL_WINDOWS else
                                 "この機能は手元のパソコンでしか使えません。"
                                 "サーバーには Chrome がないためです。"):
                        _sync_ok, _sync_msg = sync_session_from_chrome()
                        if _sync_ok:
                            st.session_state.pop("mf_client", None)
                            st.session_state.pop("mf_auth_state", None)
                            st.success(_sync_msg)
                            st.rerun()
                        else:
                            st.warning(_sync_msg)

            with _sync_col2:
                with st.expander("方法 C: _sess を手動コピペ"):
                    st.caption(
                        "F12 → Application → Cookies → www.minc.or.jp → `_sess` の Value をコピー"
                    )
                    _sess_input = st.text_input(
                        "_sess Cookie 値",
                        key="minc_sess_manual",
                        type="password",
                        placeholder="_sess の Value をここに貼り付け",
                    )
                    _xsrf_input = st.text_input(
                        "XSRF-TOKEN（任意）",
                        key="minc_xsrf_manual",
                        type="password",
                        placeholder="XSRF-TOKEN の Value（省略可）",
                    )
                    if st.button("💾 セッションを更新", key="minc_sess_save",
                                 use_container_width=True):
                        if not _sess_input.strip():
                            st.warning("_sess の値を入力してください。")
                        else:
                            update_sess_cookie(_sess_input.strip(), _xsrf_input.strip())
                            st.session_state.pop("mf_client", None)
                            st.session_state.pop("mf_auth_state", None)
                            st.success("セッションを更新しました。")
                            st.rerun()

        # ---- MINC セッション状態表示 ----
        _minc_status_col, _minc_recheck_col = st.columns([5, 1])
        with _minc_recheck_col:
            if st.button("🔄 確認", key="minc_recheck_top", use_container_width=True):
                st.session_state.pop("mf_auth_state", None)
                st.session_state.pop("mf_client", None)  # 再ログイン後は新クライアントを作成
        with _minc_status_col:
            if "mf_auth_state" not in st.session_state:
                try:
                    _mf_ok2, _mf_msg2 = check_session(_get_mf_client())
                except MusicForestError as _e:
                    _mf_ok2, _mf_msg2 = False, str(_e)
                st.session_state["mf_auth_state"] = (_mf_ok2, _mf_msg2)
            _mf_ok2, _mf_msg2 = st.session_state["mf_auth_state"]
            if _mf_ok2:
                st.success(f"✅ MINC: {_mf_msg2}")
            else:
                st.warning(f"⚠️ MINC: {_mf_msg2}")

        # ── 一括検索（楽曲まとめ・フィルター）────────────────

        # ---- フィルター ----
        fc1, fc2 = st.columns([2, 2])
        with fc1:
            status_opts = sorted(songs_df["確認ステータス"].dropna().unique().tolist())
            # 保存済みフィルター値の整合チェック
            _stored_sf = st.session_state.get("songs_status_filter")
            if _stored_sf is not None:
                if not any(s in status_opts for s in _stored_sf):
                    # 全部無効化 → リセット
                    del st.session_state["songs_status_filter"]
                else:
                    # 反映などで新ステータスが追加された場合、フィルターに自動追加
                    _new_opts = [s for s in status_opts if s not in _stored_sf]
                    if _new_opts:
                        st.session_state["songs_status_filter"] = list(_stored_sf) + _new_opts
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

        # フィルター適用（選択なし = 全件表示）
        if status_filter:
            mask = songs_df["確認ステータス"].isin(status_filter)
        else:
            mask = pd.Series([True] * len(songs_df), index=songs_df.index)
        if "すべて" not in type_filter and type_filter:
            mask &= songs_df["管理番号種別"].isin(type_filter)
        filtered_df = songs_df[mask]

        st.caption(f"表示: {len(filtered_df)} 件 ／ 全 {len(songs_df)} 件")

        # ---- 一括検索 ----
        # expander の中のウィジェットを操作すると再実行が走り、expanded 引数の値
        # （False）に戻って畳まれてしまう。操作時にコールバックでフラグを立て、
        # 開いたままにする（コールバックは再実行前に走るので次の描画に間に合う）。
        def _keep_bulk_open() -> None:
            st.session_state["bulk_search_open"] = True

        def _keep_bulk_test_open() -> None:
            st.session_state["bulk_search_open"] = True
            st.session_state["bulk_test_open"] = True

        with st.expander(
            "🔍 一括検索（MINC / J-WID / NexTone）詳細設定",
            expanded=bool(st.session_state.get("bulk_search_open")),
        ):
            bulk_target = st.radio(
                "検索対象",
                ["未調査のみ", "全曲"],
                horizontal=True,
                key="bulk_search_target",
                on_change=_keep_bulk_open,
            )
            target_mask = (
                st.session_state.songs_df["確認ステータス"].isin(["未調査", "MP3補助確認"])
                if bulk_target == "未調査のみ"
                else pd.Series([True] * len(st.session_state.songs_df),
                               index=st.session_state.songs_df.index)
            )
            target_count = int(target_mask.sum())
            _mf_ok_info, _ = st.session_state.get("mf_auth_state", (False, ""))
            # 1曲あたりの目安。3つのサイトを同時に調べるようにした後の実測
            # （MINCあり 6.0秒／3曲平均）。MINC を使わない分は少し速い
            _sec_per_song = 6 if _mf_ok_info else 4
            est_min = max(1, round(target_count * _sec_per_song / 60))

            _has_composer_info = (
                "作曲者" in st.session_state.songs_df.columns
                and st.session_state.songs_df["作曲者"].astype(str).str.strip().ne("").any()
            )
            st.info(
                f"対象: **{target_count} 件** ／ 推定所要時間: 約 {est_min} 分  \n"
                "**J-WID / NexTone**: 作曲者・作詞者・編曲者・訳詞者・I/V区分・作品コードを取得。  \n"
                + ("**MINC**: ✅ ログイン済み — レコード会社名・CD番号・CD名・委任者も取得します。  \n" if _mf_ok_info else "**MINC**: ⚠️ 未ログイン — レコード会社名/CD名/CD番号を取得するには「🔑 MINC ログイン」が必要です。  \n")
                + ("✅ 作曲者情報あり — 検索結果を作曲者で絞り込みます。  \n" if _has_composer_info else "")
                + "1件 or 作曲者一致が1件 → 自動入力（ステータス: 候補あり）  \n"
                  "複数ヒット → 自動入力せずマークのみ（下の補完検索で手動確認）"
            )

            # ---- テスト検索（診断用） ----
            with st.expander(
                "🧪 テスト検索（1曲で動作確認）",
                expanded=bool(st.session_state.get("bulk_test_open")),
            ):
                _test_title = st.text_input(
                    "曲名", key="test_search_title", placeholder="例: 風よ運んでいいよ",
                    on_change=_keep_bulk_test_open,
                )
                _test_composer = st.text_input(
                    "作曲者ヒント（任意）", key="test_search_composer",
                    on_change=_keep_bulk_test_open,
                )
                if st.button("テスト検索を実行", key="test_search_btn", on_click=_keep_bulk_test_open):
                    if _test_title.strip():
                        with st.spinner("検索中..."):
                            _test_r = search_all(_test_title.strip(), composer=_test_composer.strip())
                        _jt = _test_r.get("jwid", {})
                        _nt = _test_r.get("nextone", {})
                        st.markdown("**J-WID**")
                        if _jt.get("error"):
                            st.error(f"エラー: {_jt['error']}")
                        else:
                            st.write(f"{len(_jt.get('results') or [])} 件ヒット")
                            if _jt.get("results"):
                                st.dataframe(_jt["results"], use_container_width=True)
                        with st.expander("J-WID debug HTML（先頭2000字）"):
                            st.code(_jt.get("debug_html", "")[:2000])
                        st.markdown("**NexTone**")
                        if _nt.get("error"):
                            st.error(f"エラー: {_nt['error']}")
                        else:
                            st.write(f"{len(_nt.get('results') or [])} 件ヒット")
                            if _nt.get("results"):
                                # 管理状況は入れ子の dict なのでそのままでは表に
                                # ならない。見たいのは放送・配信なのでそこだけ開く
                                _nt_view = pd.DataFrame(_nt["results"])
                                if "管理状況" in _nt_view.columns:
                                    for _nt_c in JWID_MGMT_COLS:
                                        _nt_view[_nt_c] = [
                                            (_m or {}).get(_nt_c, "")
                                            for _m in _nt_view["管理状況"]]
                                    _nt_view = _nt_view.drop(columns=["管理状況"])
                                st.dataframe(_nt_view, use_container_width=True)
                        with st.expander("NexTone debug HTML（先頭2000字）"):
                            st.code(_nt.get("debug_html", "")[:2000])
                    else:
                        st.warning("曲名を入力してください。")

            if st.button(
                f"🔍 一括検索を実行（{target_count} 件）",
                key="bulk_search_btn",
                type="primary",
                disabled=target_count == 0,
                on_click=_keep_bulk_open,
            ):
                target_indices = st.session_state.songs_df[target_mask].index.tolist()
                total = len(target_indices)
                progress_bar = st.progress(0)
                status_ph = st.empty()
                stats: dict[str, int] = {
                    "自動入力": 0, "作曲者一致": 0, "アーティスト一致": 0,
                    "作曲者アーティスト一致": 0, "複数候補": 0,
                    "ヒットなし": 0, "エラー": 0, "MINCエラー": 0,
                    "台帳CD": 0, "MINC引き直しCD": 0,
                }
                # 自社CD台帳の引き当て結果。同じ作品コードが何度も出て
                # くるので、1回の実行の中では覚えておいて DB への往復を
                # 減らす（Supabase は1往復 0.2秒ほどかかる）
                _cd_ledger_cache: dict[str, list[dict]] = {}
                # 台帳で埋まらず MINC に聞き直した分。同じ作品コードを
                # 二度引かないために覚えておく（1回が最大12秒かかる）
                _cd_minc_cache: dict[str, list[dict]] = {}

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

                    # 作曲者ヒント（空の場合はアーティスト名でフォールバック）
                    composer_hint = str(row.get("作曲者", "")).strip()
                    if not composer_hint or composer_hint.lower() == "nan":
                        composer_hint = str(row.get("アーティスト", "")).strip()
                    if composer_hint.lower() == "nan":
                        composer_hint = ""

                    # 候補の絞り込み用。MP3のID3タグ等で既に入っている作曲者名だけを使う
                    # （アーティスト名で代用すると別人の同名曲を掴むため）
                    _comp_known = str(row.get("作曲者", "")).strip()
                    if _comp_known.lower() == "nan":
                        _comp_known = ""
                    # 同上。MP3のID3タグは作曲者が空でもアーティストは入っていることが多い
                    _art_known = str(row.get("アーティスト", "")).strip()
                    if _art_known.lower() == "nan":
                        _art_known = ""

                    status_ph.caption(f"({i + 1}/{total}) 検索中: {search_term[:50]}")
                    progress_bar.progress((i + 1) / total)

                    # ---- MINC は先に投げておく ----
                    #
                    # MINC は J-WID / NexTone とは別のサイトなので、順番に
                    # 待つ理由がない。ここで投げておいて、J-WID を調べ終えた
                    # 下の方で受け取る。1曲あたりの待ち時間が「合計」から
                    # 「一番遅い1つ」に縮む。
                    # 相手の1台から見た間隔は変わらない（待ち時間の管理は
                    # サイトごとに別勘定のため）。
                    # 別スレッドから st.session_state は触れないので、
                    # client はここ（本体側）で取っておいて渡す。
                    _mf_ok_bulk, _ = st.session_state.get("mf_auth_state", (False, ""))
                    _mf_box: dict = {}
                    _mf_th = None
                    if _mf_ok_bulk:
                        try:
                            _mf_c = _get_mf_client()
                        except Exception as _e:
                            _mf_c = None
                            _mf_box["e"] = _e
                        if _mf_c is not None:
                            def _mf_run(_c=_mf_c, _t=search_term, _box=_mf_box):
                                try:
                                    _box["r"] = _c.search(_t, match=3)
                                except Exception as _e2:
                                    _box["e"] = _e2
                            _mf_th = threading.Thread(target=_mf_run, daemon=True)
                            _mf_th.start()

                    try:
                        result = search_all(search_term, composer=composer_hint)
                    except Exception as _e:
                        stats["エラー"] += 1
                        stats.setdefault("_last_error", str(_e))
                        # 投げっぱなしにすると次の曲の検索と重なるので待つ
                        if _mf_th is not None:
                            _mf_th.join()
                        continue

                    _jwid_err = result.get("jwid", {}).get("error")
                    if _jwid_err:
                        stats["エラー"] += 1
                        stats.setdefault("_last_error", f"J-WID: {_jwid_err}")

                    jwid_r      = result.get("jwid", {}).get("results", []) or []
                    jwid_comp_n = result.get("jwid", {}).get("composer_matched_count", 0)
                    nt_r        = result.get("nextone", {}).get("results", []) or []
                    nt_comp_n   = result.get("nextone", {}).get("composer_matched_count", 0)

                    # 自動適用条件: 1件のみ OR 作曲者一致が1件（results は一致順にソート済み）
                    def _auto_apply(results, comp_n):
                        return bool(results) and (len(results) == 1 or comp_n == 1)

                    updates: dict = {}
                    _jwid_detail_ok = False  # I/V区分「インスト」判定用
                    _minc_iv = ""            # MINC から取得した IV 値（"I" or "V"）

                    if _auto_apply(jwid_r, jwid_comp_n):
                        r = jwid_r[0]
                        if r.get("作品コード"):
                            updates["JASRAC作品コード"] = r["作品コード"]
                        # J-WID 検索結果には著作者名しかなく、作曲者/作詞者は詳細ページから取得
                        _jw_durl = r.get("_detail_url", "")
                        if _jw_durl:
                            try:
                                from modules.scraper import fetch_jwid_detail as _fetch_jwd_b
                                _jwd_b = _fetch_jwd_b(_jw_durl)
                                if not _jwd_b.get("error"):
                                    _jwid_detail_ok = True
                                    if _jwd_b.get("作曲者"): updates["作曲者"] = _jwd_b["作曲者"]
                                    if _jwd_b.get("作詞者"): updates["作詞者"] = _jwd_b["作詞者"]
                                    if _jwd_b.get("編曲者"): updates["編曲者"] = _jwd_b["編曲者"]
                                    if _jwd_b.get("訳詞者"): updates["訳詞者"] = _jwd_b["訳詞者"]
                                    # 放送・配信。詳細ページはもう取ってある
                                    # ので、ここで書き取るのに追加の通信は要らない
                                    _apply_management_status(
                                        _jwd_b.get("管理状況") or {},
                                        updates, "J", row)
                            except Exception:
                                pass

                    if _auto_apply(nt_r, nt_comp_n):
                        r = nt_r[0]
                        if r.get("作曲者") and not updates.get("作曲者"):
                            updates["作曲者"] = r["作曲者"]
                        if r.get("作詞者") and not updates.get("作詞者"):
                            updates["作詞者"] = r["作詞者"]
                        if r.get("管理番号"):
                            updates["NexTone管理番号"] = r["管理番号"]
                        # 放送・配信。NexTone は検索結果の表に支分権が
                        # そのまま出ているので、追加の通信は要らない
                        _apply_management_status(
                            r.get("管理状況") or {}, updates, "N", row)
                        if r.get("アーティスト") and not updates.get("アーティスト"):
                            updates["アーティスト"] = r["アーティスト"]

                    # MINC 検索（セッション有効時のみ）。上で先に投げてある
                    # ので、ここでは結果を受け取るだけ
                    _mf_multi_match = False
                    _mf_comp_matched = False   # 作曲者まで一致した候補を採用したか
                    _mf_art_matched  = False   # アーティストまで一致した候補を採用したか
                    if _mf_ok_bulk:
                        try:
                            if _mf_th is not None:
                                _mf_th.join()
                            if _mf_box.get("e") is not None:
                                raise _mf_box["e"]
                            _mf_bulk = _mf_box.get("r") or {}
                            _mf_bulk_items = _mf_bulk.get("results", []) or []
                            # 候補の絞り込み:
                            #   ① 作品名が曲名と完全一致する候補に限定（1件だけなら無条件採用）
                            #   ② その中で作曲者が一致する候補を優先（＝どの作品か）
                            #   ③ さらにアーティストも一致する行を優先（＝どのCD／音源か）
                            # MINCの検索結果には「作詞／作曲」列も「アーティスト」列も
                            # 含まれるので、②③の照合に追加の通信は発生しない。
                            # 作曲者は作品ごとに一意だが、アーティストは同じ作品でも
                            # 音源（カバー等）ごとに変わるため、②より弱い決め手として扱う。
                            _mfr = None
                            if len(_mf_bulk_items) == 1:
                                _mf_named = _mf_bulk_items
                            else:
                                _song_n = normalize_for_match(search_term)
                                _mf_named = [
                                    _mi for _mi in _mf_bulk_items
                                    if normalize_for_match(_mi.get("作品名", "")) == _song_n
                                ]
                            _mf_by_comp = [
                                _mi for _mi in _mf_named
                                if _composer_matches(_comp_known, _mi.get("作曲者", ""))
                            ]
                            _mf_by_art = [
                                _mi for _mi in _mf_named
                                if _artist_matches(_art_known, _mi.get("アーティスト", ""))
                            ]
                            _mf_art_ids = {id(_mi) for _mi in _mf_by_art}

                            # 同じ作品で複数行あるとき、どの行を採るか。
                            #
                            # 欲しいのは CD の情報なので、まず「品番のある行
                            # ＝CD の行」を見る。MINC では CD を持つ収録曲の行
                            # はアーティスト欄が空のことが多く、アーティスト
                            # 一致だけで選ぶと、CD を持たない配信曲の行が勝っ
                            # てしまう（「ゾートロープの光の小人」）。空欄は
                            # 「不一致」ではなく「不明」として扱う。
                            #
                            # ただしアーティストが食い違う行の CD は別の盤な
                            # ので、それよりは「CD は無いがアーティストが合う
                            # 行」を採る。間違った盤を書くより空の方がまし。
                            def _mf_pick_rank(_mi: dict) -> int:
                                _has_cd = not _is_blank(_mi.get("品番", ""))
                                _art_ok = id(_mi) in _mf_art_ids
                                _art_blank = _is_blank(_mi.get("アーティスト", ""))
                                if _has_cd and _art_ok:
                                    return 0
                                if _has_cd and _art_blank:
                                    return 1
                                if _art_ok:
                                    return 2
                                if _has_cd:
                                    return 3
                                return 4

                            def _mf_pick(_cands: list) -> dict | None:
                                # min は同点なら先に出てきた方を残す（＝元の順）
                                return min(_cands, key=_mf_pick_rank) if _cands else None

                            if _mf_by_comp:
                                # 作曲者まで一致 → 同名異曲ではないと確認できたので確信度が高い。
                                _mfr = _mf_pick(_mf_by_comp)
                                _mf_comp_matched = True
                                _mf_art_matched = id(_mfr) in _mf_art_ids
                            elif _mf_by_art:
                                # 作曲者が空（ID3タグに無い等）でもアーティストで絞れた場合。
                                # カバー音源を掴む可能性が残るので作曲者一致より一段弱い扱い。
                                #
                                # 選ぶ相手はアーティストが当たった行だけでなく、曲名が
                                # 一致する行すべてにする。アーティスト欄が空の行は「別人」
                                # ではなく「不明」で、CD が付いているのはたいていその行
                                # だからで、当たった行だけを見ると CD を拾えない
                                _mfr = _mf_pick(_mf_named)
                                _mf_art_matched = id(_mfr) in _mf_art_ids
                                if not _mf_art_matched:
                                    # アーティストで当てた行ではない＝決め手は曲名だけ
                                    _mf_multi_match = len(_mf_bulk_items) > 1
                            elif _mf_named:
                                _mfr = _mf_pick(_mf_named)
                                # 曲名一致だけで複数候補から選んだ場合は要確認
                                _mf_multi_match = len(_mf_bulk_items) > 1
                            if _mfr:
                                if _mfr.get("JASRAC作品コード") and not updates.get("JASRAC作品コード"):
                                    updates["JASRAC作品コード"] = _mfr["JASRAC作品コード"]
                                if _mfr.get("NexTone管理番号") and not updates.get("NexTone管理番号"):
                                    updates["NexTone管理番号"] = _mfr["NexTone管理番号"]
                                if _mfr.get("アーティスト") and not updates.get("アーティスト"):
                                    updates["アーティスト"] = _mfr["アーティスト"]
                                if _mfr.get("品番") and not updates.get("CD番号"):
                                    updates["CD番号"] = _mfr["品番"]
                                if _mfr.get("CD商品タイトル") and not updates.get("CD名"):
                                    updates["CD名"] = _mfr["CD商品タイトル"]
                                if _mfr.get("レコード会社名") and not updates.get("レコード会社名"):
                                    updates["レコード会社名"] = _mfr["レコード会社名"]
                                # 委任者（MINC CD詳細から取得）
                                _m_alb = _mfr.get("_album_id", "")
                                _m_trk = _mfr.get("_track_id", "")
                                # album_id なし（配信曲/作品テーブル）→ ページ内 CD リンクが1件なら自動補完
                                if not (_m_alb and _m_trk):
                                    _pg_lnks = _mf_bulk.get("_page_cd_links", [])
                                    if len(_pg_lnks) == 1:
                                        _m_alb = _pg_lnks[0]["album_id"]
                                        _m_trk = _pg_lnks[0]["track_id"]
                                if _m_alb and _m_trk and not updates.get("委任者"):
                                    try:
                                        _delg_b = _mf_c.fetch_product_detail(_m_alb, _m_trk)
                                        _chuukanri = _delg_b.get("集中管理", "")
                                        if _chuukanri in ("委任者", "非委任者"):
                                            updates["委任者"] = _chuukanri
                                        _iv_minc = _delg_b.get("IV", "")
                                        if _iv_minc in ("I", "V"):
                                            _minc_iv = _iv_minc
                                    except Exception:
                                        pass
                                # MINC 作品詳細から作曲者・作詞者・編曲者・訳詞者を補完
                                _mfr_dhref = _mfr.get("_detail_href", "")
                                if _mfr_dhref:
                                    try:
                                        _mf_d = _mf_c.get_detail(_mfr_dhref)
                                        if not _mf_d.get("error"):
                                            if _mf_d.get("作曲者") and not updates.get("作曲者"): updates["作曲者"] = _mf_d["作曲者"]
                                            if _mf_d.get("作詞者") and not updates.get("作詞者"): updates["作詞者"] = _mf_d["作詞者"]
                                            if _mf_d.get("編曲者") and not updates.get("編曲者"): updates["編曲者"] = _mf_d["編曲者"]
                                            if _mf_d.get("訳詞者") and not updates.get("訳詞者"): updates["訳詞者"] = _mf_d["訳詞者"]
                                    except Exception:
                                        pass
                                # 詳細ページを引けなかった／作家名が載っていなかった分は
                                # 検索結果の「作詞／作曲」列で埋める（通信なし）
                                for _ak in ("作曲者", "作詞者", "編曲者", "訳詞者"):
                                    if _mfr.get(_ak) and not updates.get(_ak):
                                        updates[_ak] = _mfr[_ak]
                        except Exception as _me:
                            stats["MINCエラー"] += 1
                            stats.setdefault("_minc_last_error", f"{type(_me).__name__}: {_me}")

                    # ---- CD の情報を自社CD台帳（TSP）から埋める ----
                    #
                    # MINC に聞き直すと1曲ごとに通信が1回増える（応答は実測
                    # 0.6〜12秒）。台帳はこちらの DB なので、作品コードさえ
                    # 分かれば通信なしで品番・CD名・レコード会社名が入る。
                    # 台帳が空（まだ取り込んでいない）なら何も起きない。
                    _cd_code = (updates.get("JASRAC作品コード")
                                or str(row.get("JASRAC作品コード", ""))).strip()
                    _cd_filled = not _is_blank(
                        updates.get("CD番号") or row.get("CD番号", ""))
                    if _cd_code and not _cd_filled:
                        _cd_dig = "".join(c for c in _cd_code if c.isdigit())
                        if _cd_dig not in _cd_ledger_cache:
                            try:
                                _cd_ledger_cache.update(
                                    cd_fetch_by_jasrac({_cd_dig})
                                    or {_cd_dig: []})
                            except Exception:
                                _cd_ledger_cache[_cd_dig] = []
                        _cd_cands = _cd_ledger_cache.get(_cd_dig) or []
                        # 同じ作品が複数の盤に入っていることがある。手元に
                        # アーティスト名があれば、それに合う盤を選ぶ
                        _cd_hit = next(
                            (_c for _c in _cd_cands
                             if _artist_matches(_art_known, _c.get("artist", ""))),
                            _cd_cands[0] if _cd_cands else None,
                        )
                        if _cd_hit:
                            for _col, _key in (("CD番号", "cd_no"),
                                               ("CD名", "cd_name"),
                                               ("レコード会社名", "label")):
                                if _cd_hit.get(_key) and not updates.get(_col) \
                                        and _is_blank(row.get(_col, "")):
                                    updates[_col] = _cd_hit[_key]
                            if updates.get("CD番号"):
                                stats["台帳CD"] += 1

                    # ---- それでも埋まらなかった分だけ MINC に聞く ----
                    #
                    # 台帳は JASRAC作品コードでしか引けないので、NexTone
                    # だけで管理されている曲は当たらない。検索結果の行に
                    # 品番が載っていないこともある。そういう残りだけ、作品
                    # コードから CD商品リストを引き直す。
                    #
                    # ここは1曲につき通信が1回増える（実測 0.6〜12秒）。
                    # 埋め残しだけが対象なので、実測では33曲中4曲だった。
                    _cd_filled2 = not _is_blank(
                        updates.get("CD番号") or row.get("CD番号", ""))
                    _cd_ncd = (updates.get("NexTone管理番号")
                               or str(row.get("NexTone管理番号", ""))).strip()
                    if _is_blank(_cd_ncd):
                        _cd_ncd = ""
                    if _mf_ok_bulk and not _cd_filled2 and (_cd_code or _cd_ncd):
                        _mk = f"{_cd_code}|{_cd_ncd}"
                        if _mk not in _cd_minc_cache:
                            try:
                                _cd_minc_cache[_mk] = (
                                    _get_mf_client().search_cds_by_jasrac(
                                        _cd_code,
                                        title=search_term,
                                        ncd=_cd_ncd,
                                    ).get("cds") or []
                                )
                            except Exception as _ce:
                                _cd_minc_cache[_mk] = []
                                stats["MINCエラー"] += 1
                                stats.setdefault(
                                    "_minc_last_error",
                                    f"{type(_ce).__name__}: {_ce}")
                        _mc = _cd_minc_cache.get(_mk) or []
                        # 台帳と同じ選び方。手元にアーティスト名があれば
                        # それに合う盤、無ければ先頭
                        _mc_hit = next(
                            (_c for _c in _mc
                             if _artist_matches(_art_known,
                                                _c.get("アーティスト", ""))),
                            _mc[0] if _mc else None,
                        )
                        if _mc_hit:
                            for _col, _key in (("CD番号", "品番"),
                                               ("CD名", "CD商品タイトル"),
                                               ("レコード会社名", "レコード会社名")):
                                if _mc_hit.get(_key) and not updates.get(_col) \
                                        and _is_blank(row.get(_col, "")):
                                    updates[_col] = _mc_hit[_key]
                            if updates.get("CD番号"):
                                stats["MINC引き直しCD"] += 1

                    # I/V区分 自動判定
                    #   ① MINC の CD情報に I/V 表記があればそれを使う
                    #   ② 無ければ作詞者の有無で判定（作詞者あり→ヴォーカル／なし→インスト）
                    # ②は作家名を取得できた行に限る（未取得の空欄をインストにしないため）
                    _BLANK = ("", "nan", "none")
                    _new_lyr   = updates.get("作詞者", "").strip()
                    _exist_lyr = str(row.get("作詞者", "")).strip()
                    _lyr = _new_lyr or ("" if _exist_lyr.lower() in _BLANK else _exist_lyr)
                    _cred_known = bool(
                        _jwid_detail_ok
                        or _lyr
                        or updates.get("作曲者", "").strip()
                        or str(row.get("作曲者", "")).strip().lower() not in _BLANK
                    )
                    _iv_set    = str(row.get("I/V区分", "")).strip().lower() not in _BLANK
                    if not _iv_set:
                        if _minc_iv == "I":
                            updates["I/V区分"] = "インスト"
                        elif _minc_iv == "V":
                            updates["I/V区分"] = "ヴォーカル"
                        elif _cred_known:
                            updates["I/V区分"] = _infer_iv(_lyr)

                    # 原訳詞区分 自動判定（作詞者ありで未設定なら "原詞"）
                    if (_new_lyr or (_exist_lyr and _exist_lyr.lower() not in _BLANK)):
                        if str(row.get("原訳詞区分", "")).strip().lower() in _BLANK and not updates.get("原訳詞区分"):
                            updates["原訳詞区分"] = "原詞"

                    # 邦洋区分 自動判定（JASRACコード 2文字目: 数字→邦楽、英字→洋楽）
                    _jasrac = (updates.get("JASRAC作品コード") or str(row.get("JASRAC作品コード", ""))).strip()
                    if _jasrac and len(_jasrac) >= 2 and str(row.get("邦洋区分", "")).strip().lower() in _BLANK and not updates.get("邦洋区分"):
                        _c2 = _jasrac[1]
                        if _c2.isdigit():
                            updates["邦洋区分"] = "邦楽"
                        elif _c2.isalpha():
                            updates["邦洋区分"] = "洋楽"

                    if updates:
                        # 作曲者一致 > アーティスト一致 > 曲名のみ一致 > 複数候補から選択
                        if _mf_comp_matched:
                            updates["確認ステータス"] = "作曲者一致"
                            stats["作曲者一致"] += 1
                            if _mf_art_matched:
                                stats["作曲者アーティスト一致"] += 1
                        elif _mf_art_matched:
                            updates["確認ステータス"] = "アーティスト一致"
                            stats["アーティスト一致"] += 1
                        else:
                            updates["確認ステータス"] = "複数候補あり" if _mf_multi_match else "候補あり"
                        for col, val in updates.items():
                            if col in st.session_state.songs_df.columns:
                                st.session_state.songs_df.at[idx, col] = val
                        stats["自動入力"] += 1
                    elif jwid_r or nt_r:
                        # 当たりはあったが自動では入れなかった＝どれか選ぶ
                        # 必要がある行。人が見に行く先なので、そう書く
                        st.session_state.songs_df.at[idx, "確認ステータス"] = "複数候補あり"
                        stats["複数候補"] += 1
                    else:
                        st.session_state.songs_df.at[idx, "確認ステータス"] = "該当なし"
                        stats["ヒットなし"] += 1

                progress_bar.empty()
                status_ph.empty()

                result_msg = (
                    f"✅ 完了: 自動入力 {stats['自動入力']} 件"
                    + (f"（うち作曲者まで一致 {stats['作曲者一致']} 件"
                       + (f"／うちアーティストも一致 {stats['作曲者アーティスト一致']} 件"
                          if stats["作曲者アーティスト一致"] else "")
                       + "）"
                       if stats["作曲者一致"] else "")
                    + (f"（うちアーティストのみ一致 {stats['アーティスト一致']} 件）"
                       if stats["アーティスト一致"] else "")
                    + (f"（うち CD を台帳から {stats['台帳CD']} 件）"
                       if stats["台帳CD"] else "")
                    + (f"（うち CD を MINC から {stats['MINC引き直しCD']} 件）"
                       if stats["MINC引き直しCD"] else "")
                    + f" ／ 複数候補 {stats['複数候補']} 件 ／ "
                    f"ヒットなし {stats['ヒットなし']} 件"
                )
                if stats["エラー"]:
                    result_msg += f" ／ エラー {stats['エラー']} 件"
                if stats.get("_last_error"):
                    result_msg += f"  \nJ-WID/NexToneエラー: {stats['_last_error']}"
                if stats["MINCエラー"]:
                    result_msg += (
                        f"  \n⚠️ MINC エラー {stats['MINCエラー']} 件: "
                        f"{stats.get('_minc_last_error','')}"
                    )
                # 直後に rerun するとこの場で出したメッセージは消えるため持ち越す
                st.session_state["_apply_msg"] = result_msg
                # 何分もかけて取った結果を取り直さずに済むよう、ここで保存する。
                # 結果は songs_df に書き込んだだけで、DBにはまだ入っていない
                _autosave_to_db("（一括検索の完了時）")
                # 調べがついた曲を共有データに貯める。次に同じ曲が来たら
                # 照合実行の時点で自動的に入る
                _learned = _master_learn()
                if _learned:
                    st.session_state["_apply_msg"] += (
                        f"  \n🗃️ 共有楽曲データに {_learned} 曲を貯めました。"
                    )
                # 検索が終わったら詳細設定は畳んでよい
                st.session_state.pop("bulk_search_open", None)
                st.rerun()

        # ---- MusicBrainz / Spotify / Claude 一括補完 ----
        # 元は「⚙️ ファイル読み込み・設定」にあったが、照合実行後でないと
        # 動かない機能なので、同じく照合後に使う一括検索の隣へ移した。
        with st.expander(
            "🎵 MusicBrainz / Spotify / Claude で一括補完",
            expanded=False,
        ):
            st.caption(
                "**曲名** × WAV尺で MusicBrainz を検索し ISRC・アーティストを取得します。"
                " MP3 の作曲者情報が入力済みなら、J-WID / NexTone の結果を作曲者で絞り込みます。"
            )

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
                        _mb_composer = str(_mb_row.get("作曲者", "")).strip()
                        if not _mb_composer or _mb_composer.lower() == "nan":
                            _mb_composer = str(_mb_row.get("アーティスト", "")).strip()
                        if _mb_composer.lower() == "nan":
                            _mb_composer = ""
                        _pip = run_pipeline(
                            event_name=_mb_name,
                            wav_full_duration=str(_mb_row.get("WAVフル尺", "")),
                            wav_detected_title=str(_mb_row.get("WAV検出タイトル", "")),
                            song_title=str(_mb_row.get("曲名", "")),
                            composer=_mb_composer,
                            tolerance_sec=float(mb_bulk_tol),
                            mb_score_threshold=int(mb_bulk_thresh),
                            use_claude=bool(mb_bulk_use_claude),
                            use_spotify=bool(mb_bulk_use_spotify),
                        )
                        _mb_best = _pip.get("mb_best")
                        if _mb_best and _mb_best.get("score", 0) >= int(mb_bulk_thresh):
                            _mb_stats["MB命中"] += 1

                        # Spotify 結果で補完（MusicBrainz 未取得 or 低スコアのとき）
                        _sp_best = _pip.get("sp_best")
                        if _sp_best and not _sp_best.get("error"):
                            if not st.session_state.songs_df.at[_mb_idx, "アーティスト"] and _sp_best.get("artist"):
                                st.session_state.songs_df.at[_mb_idx, "アーティスト"] = _sp_best["artist"]
                            if _sp_best.get("album") and "CD名" in st.session_state.songs_df.columns:
                                if not st.session_state.songs_df.at[_mb_idx, "CD名"]:
                                    st.session_state.songs_df.at[_mb_idx, "CD名"] = _sp_best["album"]

                        # Claude API 結果で補完（J-WID / MB / Spotify より後に上書きしない）
                        _cl = _pip.get("claude_result") or {}
                        if not _cl.get("error") and _cl.get("confidence") in ("high", "medium"):
                            if _cl.get("artist") and not st.session_state.songs_df.at[_mb_idx, "アーティスト"]:
                                st.session_state.songs_df.at[_mb_idx, "アーティスト"] = _cl["artist"]
                            if _cl.get("composer") and not st.session_state.songs_df.at[_mb_idx, "作曲者"]:
                                st.session_state.songs_df.at[_mb_idx, "作曲者"] = _cl["composer"]
                            if _cl.get("cd_name") and "CD名" in st.session_state.songs_df.columns:
                                if not st.session_state.songs_df.at[_mb_idx, "CD名"]:
                                    st.session_state.songs_df.at[_mb_idx, "CD名"] = _cl["cd_name"]

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
                            if _r.get("訳詞者"):
                                st.session_state.songs_df.at[_mb_idx, "訳詞者"] = _r["訳詞者"]
                            if _r.get("作品コード"):
                                _pip_jcd = _r["作品コード"]
                                st.session_state.songs_df.at[_mb_idx, "JASRAC作品コード"] = _pip_jcd
                                _hy = _infer_houyo(_pip_jcd)
                                if _hy and not str(st.session_state.songs_df.at[_mb_idx, "邦洋区分"] if "邦洋区分" in st.session_state.songs_df.columns else "").strip():
                                    if "邦洋区分" in st.session_state.songs_df.columns:
                                        st.session_state.songs_df.at[_mb_idx, "邦洋区分"] = _hy
                            st.session_state.songs_df.at[_mb_idx, "確認ステータス"] = "候補あり"
                        elif _jw_r:  # 当たりはあったが絞り込めなかった
                            # 人がどれか選びに行く先なので、1件だけ当たった
                            # 行（候補あり）とは分けて書く
                            st.session_state.songs_df.at[_mb_idx, "確認ステータス"] = "複数候補あり"

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
                            # 放送・配信は検索結果に出ているので通信は増えない
                            _write_mgmt_from_fetch(
                                st.session_state.songs_df, _mb_idx,
                                _rn.get("管理状況") or {}, "N")
                            if st.session_state.songs_df.at[_mb_idx, "確認ステータス"] in ("未調査", "MP3補助確認"):
                                st.session_state.songs_df.at[_mb_idx, "確認ステータス"] = "候補あり"
                        elif _nt_r:
                            if st.session_state.songs_df.at[_mb_idx, "確認ステータス"] in ("未調査", "MP3補助確認"):
                                st.session_state.songs_df.at[_mb_idx, "確認ステータス"] = "複数候補あり"
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
                # 一括検索と同じ理由でここでも保存する。この後に rerun しない
                # ので、伝言は持ち越さずその場で出す
                if _autosave_to_db("（一括補完の完了時）"):
                    st.info(st.session_state.pop("_autosave_msg"))
                _learned = _master_learn()
                if _learned:
                    st.info(f"🗃️ 共有楽曲データに {_learned} 曲を貯めました。")

        # ---- 申告フォーマット プレビュー（提出用・イベント行単位）----
        # 反映ボタン後の成功メッセージ（反映結果の表が見えるようここへ自動スクロール）
        st.markdown('<a id="sec-shinkok"></a>', unsafe_allow_html=True)
        _apply_done = "_apply_msg" in st.session_state
        if _apply_done:
            st.success(st.session_state.pop("_apply_msg"))

        _shinkok_songs  = st.session_state.songs_df
        _shinkok_events = st.session_state.events_df
        _shinkok_df: pd.DataFrame | None = None
        if _shinkok_events is not None and len(_shinkok_events) > 0:
            _shinkok_df = build_shinkok_df(_shinkok_songs, _shinkok_events)
            _preview_cols = [c for c in _shinkok_df.columns if c not in {"トラック", "START TIME", "使用尺"}]

            # スクロールJS（ナビゲーションボタン押下後の rerun で実行）
            # 反映ボタン直後は申告フォーマットの表まで自動で移動する。
            # 明示的に指定されたスクロール先（各セクションへのジャンプ）が優先。
            if _apply_done:
                st.session_state.setdefault("_scroll_target", "sec-shinkok")
            _sh_scroll = st.session_state.pop("_scroll_target", None)
            if _sh_scroll:
                _stc.html(_SCROLL_JS.replace("__ID__", _sh_scroll), height=0)

            _SHINKOK_COL_CFG = {
                "使用時間（分）": st.column_config.NumberColumn("分", width="small", format="%d"),
                "使用時間（秒）": st.column_config.NumberColumn("秒", width="small", format="%d"),
                "使用形態":      st.column_config.TextColumn("使用形態", width="small"),
                "音源区分":      st.column_config.TextColumn("音源区分", width="small"),
                "I/V区分":      st.column_config.TextColumn("I/V区分",  width="small"),
                "邦・洋区分":    st.column_config.TextColumn("邦・洋区分", width="small"),
                "原・訳詞区分":  st.column_config.TextColumn("原・訳詞区分", width="small"),
                "確認ステータス": st.column_config.TextColumn("確認ステータス", width="medium"),
                "委任者":        st.column_config.TextColumn("委任者", width="small"),
                # 放送・配信は J-WID と NexTone の管理状況を写した欄。人が
                # 書き換えるものではないので編集できないようにしてある
                "放送": st.column_config.TextColumn(
                    "放送", width="small",
                    help="管理状況「放送」。○ 管理あり／△ 一部管理／× 管理なし。"
                         "(N) は NexTone 側の値（JASRAC が × でも NexTone が"
                         "管理していることがある）。空欄はまだ引いていない",
                ),
                "配信": st.column_config.TextColumn(
                    "配信", width="small",
                    help="管理状況「配信」。○ 管理あり／△ 一部管理／× 管理なし。"
                         "(N) は NexTone 側の値（JASRAC が × でも NexTone が"
                         "管理していることがある）。空欄はまだ引いていない",
                ),
                "CD名":          st.column_config.TextColumn("CD名", width="medium"),
            }

            st.caption(
                f"申告フォーマット：{len(_shinkok_df)} 行 ／ {_shinkok_songs['イベント名'].nunique()} 曲"
                "　ダブルクリックで直接編集できます（楽曲まとめに自動保存）。"
                "　左端にチェックを入れると、下に補完検索へ移動するボタンが出ます。"
            )

            # 見る・選ぶ・直すを1つの表でまかなう。
            #
            # st.data_editor には st.dataframe のような行選択が無いので、
            # 左端に「選択」というチェック欄を自前で足している。1行だけに
            # 保つ番は _pick_shinkok_row が持つ。
            #
            # 同じ理由で、この表には確認ステータスの色を敷けない
            # （data_editor は Styler を受け取れない）。色の代わりに、
            # 隣に段階を表す1文字の「状態」列を置く。こちらは編集できない。
            #
            # 印の意味は見出しに添える。表のセル1つ1つに説明を出す仕組みは
            # Streamlit の表には無い（中身が canvas で描かれていて、セルが
            # HTML の要素になっていないため）。正確な言葉が要るときは、
            # 同じ表の「確認ステータス」列にそのまま出ている。
            _shinkok_src = _shinkok_df[_preview_cols].copy()
            # 印は「どこまで調べたか」ではなく「このまま出せるか」で付ける。
            # 調べ方の途中経過は隣の確認ステータス列にそのまま出ている
            _shinkok_src.insert(0, "状態", [
                issue_mark(r) for _, r in _shinkok_src.iterrows()
            ])
            _shinkok_src.insert(0, "選択", False)
            # 表示に使った DataFrame の位置で編集差分が返ってくるので、
            # 行番号→イベント名 を引けるように控えておく
            st.session_state["_shinkok_src"] = _shinkok_src

            _SHINKOK_COL_CFG = {
                **_SHINKOK_COL_CFG,
                "選択": st.column_config.CheckboxColumn("選択", width="small"),
                "状態": st.column_config.TextColumn(
                    "状態", width="small",
                    help="このまま申告に出せるかどうかの印。\n\n"
                         "🔴 空欄あり … レコード会社名・レコード番号・"
                         "邦洋・I/V・作曲・アーティストのどれかが空、"
                         "ヴォーカルなのに作詞が空、JASRAC も NexTone も"
                         "番号が無い\n\n"
                         "🟡 権利状態注意 … 非委任者、または放送・配信が"
                         "○ではない（まだ引いていない空欄も含む）\n\n"
                         "印なし … 空欄も無く、権利も○。そのまま出せる\n\n"
                         "どこまで調べたかは隣の「確認ステータス」列に出ています",
                ),
            }
            _edited_shinkok = st.data_editor(
                _shinkok_src,
                use_container_width=True,
                hide_index=True,
                height=460,
                key="shinkok_editor",
                column_config=_SHINKOK_COL_CFG,
                disabled=["状態", *JWID_MGMT_COLS],
                on_change=_sync_shinkok_to_songs,
            )

            # チェックされた行のナビゲーションボタンを即表示。
            # 表のすぐ下に置く（印の説明よりも先）。行き先を選ぶのは表を
            # 見ながらなので、スクロールしないと両方見えない位置だと使えない
            _sel_pos = st.session_state.get("_shinkok_sel")
            _sel_rows = ([_sel_pos]
                         if isinstance(_sel_pos, int) and 0 <= _sel_pos < len(_shinkok_df)
                         else [])
            if _sel_rows:
                _sel_ev = str(_shinkok_df.iloc[_sel_rows[0]].get("イベント名", "")).strip()
                _smatch2 = (
                    _shinkok_songs[_shinkok_songs["イベント名"] == _sel_ev]
                    if _sel_ev else pd.DataFrame()
                )
                if not _smatch2.empty:
                    _sm2 = _smatch2.iloc[0]
                    _no2 = int(_sm2["No"])
                    _name2 = str(_sm2.get("曲名", _sel_ev) or _sel_ev).strip() or _sel_ev
                    _status2 = str(_sm2.get("確認ステータス", "未調査") or "未調査").strip()
                    _sel_label2 = f"{_no2}. [{_status2}] {_sel_ev}"
                    _gbc1, _gbc2, _gbc3, _gbc4 = st.columns([2.5, 2.5, 2.5, 3.5])
                    with _gbc1:
                        if st.button("🌲 MINC で調査", key="shin_goto_minc", use_container_width=True):
                            st.session_state["search_song_select"] = _sel_label2
                            st.session_state["_scroll_target"] = "sec-minc-individual"
                            st.rerun()
                    with _gbc2:
                        if st.button("🔄 パイプラインで調査", key="shin_goto_pip", use_container_width=True):
                            st.session_state["search_song_select"] = _sel_label2
                            st.session_state["_scroll_target"] = "sec-pipeline"
                            st.rerun()
                    with _gbc3:
                        if st.button("💿 CD情報を検索", key="shin_goto_cds", use_container_width=True):
                            st.session_state["search_song_select"] = _sel_label2
                            st.session_state["_scroll_target"] = "sec-cd-search"
                            st.rerun()
                    with _gbc4:
                        st.caption(f"📍 **{_name2}** [{_status2}]")

            st.caption(
                f"状態の印: {ISSUE_MARK_LEGEND}"
                "　🔴 は申告に要る欄が空いている行、🟡 は非委任者または"
                "放送・配信が○でない行。どこまで調べたかは"
                "「確認ステータス」列に出ています。"
            )

            # 放送・配信を引く。
            #
            # 表に空欄が見えている、その場で引けるようにする。同じことは
            # 「書き出し」タブの管理状況 CSV でもできるが、空欄に気づくのは
            # この表なので、気づいた場所に導線を置く。
            # 引く行があるときだけ出す
            _mg_mask = _mgmt_missing_mask(_shinkok_songs)
            if _mg_mask is not None and int(_mg_mask.sum()):
                _mg_n = int(_mg_mask.sum())
                with st.expander(f"📡 放送・配信を引く（{_mg_n} 曲）"):
                    st.caption(
                        "JASRAC作品コードで J-WID を、NexTone管理番号で NexTone を"
                        "引いて、放送・配信の管理状況（○△×）を入れます。"
                        "両方に番号がある曲は両方引き、どちらかが ○ なら ○ と出します"
                        "（NexTone 側を採ったときは「○(N)」）。"
                        "1曲あたり 1〜2 秒かかります。既に入っている欄は引き直しません。"
                        "　※管理状況は変わるものなので、共有楽曲データには貯めて"
                        "いません。最新を知りたいときはここから引き直してください。"
                    )
                    st.dataframe(
                        _shinkok_songs.loc[
                            _mg_mask,
                            [c for c in ("No", "曲名", "JASRAC作品コード",
                                         "NexTone管理番号",
                                         "確認ステータス", "放送", "配信")
                             if c in _shinkok_songs.columns]
                        ],
                        use_container_width=True, hide_index=True, height=200,
                    )
                    if st.button(f"引く（{_mg_n} 曲）",
                                 key="shinkok_fetch_mgmt", type="primary"):
                        from modules.scraper import (
                            fetch_jwid_rights_by_code as _fetch_rights,
                            fetch_nextone_rights_by_code as _fetch_nrights)
                        _mg_songs = st.session_state.songs_df
                        _mg_idxs = list(_shinkok_songs.index[_mg_mask])
                        _mg_prog = st.progress(0)
                        _mg_stat = st.empty()
                        _mg_cache: dict[tuple, dict] = {}
                        _mg_filled = _mg_err = 0
                        for _mg_i, _mg_idx in enumerate(_mg_idxs):
                            _mg_name = str(_mg_songs.at[_mg_idx, "曲名"]).strip()
                            _mg_stat.text(f"取得中: {_mg_i + 1}/{_mg_n} — {_mg_name}")
                            # JASRAC・NexTone の順に、番号がある方だけ引く
                            for _mg_src, _mg_col_code, _mg_fn in (
                                    ("J", "JASRAC作品コード", _fetch_rights),
                                    ("N", "NexTone管理番号", _fetch_nrights)):
                                if _mg_col_code not in _mg_songs.columns:
                                    continue
                                _mg_code = _mgmt_val(
                                    _mg_songs.at[_mg_idx, _mg_col_code])
                                if not _mg_code:
                                    continue
                                # この行でその出どころが埋まっているなら引かない
                                if all(_mgmt_val(_mg_songs.at[_mg_idx, _p[
                                            1 if _mg_src == "N" else 0]])
                                       for _p in MGMT_SOURCE_COLS.values()):
                                    continue
                                # 同じ番号の曲が複数あっても引くのは1回
                                _mg_key = (_mg_src, _mg_code)
                                if _mg_key not in _mg_cache:
                                    try:
                                        _mg_cache[_mg_key] = _mg_fn(_mg_code)
                                    except Exception as _e:
                                        _mg_cache[_mg_key] = {"error": str(_e),
                                                              "管理状況": {}}
                                _mg_res = _mg_cache[_mg_key]
                                if _mg_res.get("error"):
                                    _mg_err += 1
                                _mg_filled += _write_mgmt_from_fetch(
                                    _mg_songs, _mg_idx,
                                    _mg_res.get("管理状況") or {}, _mg_src)
                            _mg_prog.progress((_mg_i + 1) / _mg_n)
                        _mg_prog.empty()
                        _mg_stat.empty()
                        if _mg_err:
                            st.warning(f"⚠️ {_mg_err} 件は取得できませんでした。")
                        if not _mg_filled:
                            st.info("埋めるものはありませんでした。")
                        elif _autosave_to_db("（放送・配信の取り込み）"):
                            st.session_state["_autosave_msg"] = (
                                f"放送・配信を {_mg_filled} 欄ぶん入れました"
                            )
                        st.rerun()

            # 委任者を MINC から引く。
            #
            # 委任者は MINC の CD商品詳細にしか出ていない。作品の番号で
            # 収録CDの一覧を取り、その中からこの曲の CD を選んで詳細を
            # 見る、という2手が要る。データベースで当たった行は MINC を
            # 引いていないので、ここが委任者の入り口になる。
            # 引く行があるときだけ出す
            _dg_mask = _delegate_missing_mask(_shinkok_songs)
            if _dg_mask is not None and int(_dg_mask.sum()):
                _dg_n = int(_dg_mask.sum())
                with st.expander(f"🧾 委任者を MINC から引く（{_dg_n} 曲）"):
                    st.caption(
                        "作品コードで MINC の収録CD一覧を引き、CD番号が同じ CD の"
                        "商品詳細から委任者区分（委任者／非委任者）を入れます。"
                        "同じ CD が見つからない曲は飛ばします（別の盤の委任者を"
                        "入れてしまわないため）。ついでに空いていれば"
                        "レコード会社名・CD名・I/V区分も同じ画面から入れます。"
                        "　1曲につき MINC を3回引くので、1曲あたり 4〜6 秒"
                        "かかります（20曲で2分ほど）。MINC にログインしている"
                        "必要があります。"
                    )
                    st.dataframe(
                        _shinkok_songs.loc[
                            _dg_mask,
                            [c for c in ("No", "曲名", "CD番号",
                                         "JASRAC作品コード", "NexTone管理番号",
                                         "確認ステータス")
                             if c in _shinkok_songs.columns]
                        ],
                        use_container_width=True, hide_index=True, height=200,
                    )
                    if st.button(f"MINC から引く（{_dg_n} 曲）",
                                 key="shinkok_fetch_delegate", type="primary"):
                        _dg_c = _get_mf_client()
                        _dg_ok, _dg_msg = check_session(_dg_c)
                        if not _dg_ok:
                            st.error(
                                f"MINC のセッションが使えません（{_dg_msg}）。"
                                "「検索補助」タブの 🔑 MINC ログインからつなぎ直して"
                                "ください。"
                            )
                        else:
                            _dg_songs = st.session_state.songs_df
                            _dg_idxs = list(_shinkok_songs.index[_dg_mask])
                            _dg_prog = st.progress(0)
                            _dg_stat = st.empty()
                            _dg_cds: dict[tuple, dict] = {}
                            _dg_det: dict[tuple, dict] = {}
                            _dg_filled = _dg_skip = _dg_err = 0
                            for _dg_i, _dg_idx in enumerate(_dg_idxs):
                                _dg_row = _dg_songs.loc[_dg_idx]
                                _dg_title = str(_dg_row.get("曲名", "")).strip()
                                _dg_jcd = _mgmt_val(_dg_row.get("JASRAC作品コード"))
                                _dg_ncd = _mgmt_val(_dg_row.get("NexTone管理番号"))
                                _dg_cdno = _mgmt_val(_dg_row.get("CD番号"))
                                _dg_stat.text(
                                    f"取得中: {_dg_i + 1}/{_dg_n} — {_dg_title}")
                                try:
                                    # 同じ作品の曲が複数あっても一覧を引くのは1回
                                    _dg_key = (_dg_jcd, _dg_ncd, _dg_title)
                                    if _dg_key not in _dg_cds:
                                        _dg_cds[_dg_key] = _dg_c.search_cds_by_jasrac(
                                            _dg_jcd, _dg_title, _dg_ncd)
                                    _dg_list = _dg_cds[_dg_key].get("cds") or []
                                    # この曲の CD を選ぶ。委任者は盤ごとの話なので、
                                    # CD番号が合う盤が無ければ入れない
                                    _dg_pick = None
                                    if _dg_cdno:
                                        _dg_want = _norm_hinban(_dg_cdno)
                                        _dg_pick = next(
                                            (c for c in _dg_list
                                             if _norm_hinban(c.get("品番", "")) == _dg_want),
                                            None)
                                    elif len(_dg_list) == 1:
                                        # CD番号がまだ無い曲。収録CDが1枚しか
                                        # 無いなら、選びようがないので確定できる
                                        _dg_pick = _dg_list[0]
                                    if not _dg_pick or not _dg_pick.get("album_id"):
                                        _dg_skip += 1
                                        _dg_prog.progress((_dg_i + 1) / _dg_n)
                                        continue
                                    # CD商品リストの行には track_id が入って
                                    # いない。fetch_track_list は曲名から
                                    # track_id を探すところまでやってくれる
                                    _dg_dkey = (_dg_pick["album_id"], _dg_title)
                                    if _dg_dkey not in _dg_det:
                                        _dg_det[_dg_dkey] = _dg_c.fetch_track_list(
                                            _dg_pick["album_id"],
                                            _dg_pick.get("track_id", ""),
                                            _dg_title)
                                    _dg_d = _dg_det[_dg_dkey]
                                    if _dg_d.get("error"):
                                        _dg_err += 1
                                        _dg_prog.progress((_dg_i + 1) / _dg_n)
                                        continue
                                except Exception:
                                    _dg_err += 1
                                    _dg_prog.progress((_dg_i + 1) / _dg_n)
                                    continue

                                # 埋めるのは空いている欄だけ。人が入れた値は動かさない
                                _dg_u: dict = {}
                                if _dg_d.get("集中管理") in ("委任者", "非委任者"):
                                    _dg_u["委任者"] = _dg_d["集中管理"]
                                # 収録曲の中からこの曲を探して I/V を見る
                                _dg_key2 = normalize_for_match(_dg_title)
                                _dg_tr = next(
                                    (t for t in (_dg_d.get("tracks") or [])
                                     if normalize_for_match(t.get("曲名", "")) == _dg_key2),
                                    None) if _dg_key2 else None
                                if _dg_tr and _dg_tr.get("IV") == "I":
                                    _dg_u["I/V区分"] = "インスト"
                                elif _dg_tr and _dg_tr.get("IV") == "V":
                                    _dg_u["I/V区分"] = "ヴォーカル"
                                if _dg_tr and _dg_tr.get("アーティスト"):
                                    _dg_u["アーティスト"] = _dg_tr["アーティスト"]
                                _dg_u["CD番号"] = (_dg_d.get("品番")
                                                   or _dg_pick.get("品番", ""))
                                _dg_u["CD名"] = (_dg_d.get("CD商品タイトル")
                                                 or _dg_pick.get("CD商品タイトル", ""))
                                _dg_u["レコード会社名"] = (
                                    _dg_d.get("レコード会社名")
                                    or _dg_pick.get("レコード会社名", ""))
                                _dg_u = {k: v for k, v in _dg_u.items() if v}
                                if not _dg_u.get("委任者"):
                                    _dg_skip += 1
                                for _dg_col, _dg_val in _dg_u.items():
                                    if _dg_col not in _dg_songs.columns:
                                        continue
                                    if _mgmt_val(_dg_songs.at[_dg_idx, _dg_col]):
                                        continue
                                    _dg_songs.at[_dg_idx, _dg_col] = _dg_val
                                    _dg_filled += 1
                                _dg_prog.progress((_dg_i + 1) / _dg_n)
                            _dg_prog.empty()
                            _dg_stat.empty()
                            if _dg_err:
                                st.warning(f"⚠️ {_dg_err} 曲は取得できませんでした。")
                            if _dg_skip:
                                st.info(
                                    f"{_dg_skip} 曲は委任者を決められませんでした"
                                    "（CD番号の合う盤が MINC に無い、または商品詳細に"
                                    "委任者区分が出ていない）。"
                                )
                            if not _dg_filled:
                                st.info("埋めるものはありませんでした。")
                            elif _autosave_to_db("（委任者の取り込み）"):
                                st.session_state["_autosave_msg"] = (
                                    f"委任者ほかを {_dg_filled} 欄ぶん入れました"
                                )
                            st.rerun()

            # データベース（共有楽曲データ・自社CDの台帳）の当たりを書き直す。
            #
            # 読み込みのときに一度引いているので、普段は押す必要がない。
            # 押すのは、そのあとデータベースを更新したときだけ。だから
            # 放送・配信より下に置いてある。埋めるのは空欄だけ・上書き
            # するのは「未調査」の確認ステータスだけなので、何度押しても害はない
            with st.expander("🗃️ データベースで空欄を埋める（通信なし・一瞬）"):
                st.caption(
                    "**CSV を読み込んだあとにデータベースを更新した場合だけ押してください。**"
                    "　読み込みのときに一度引いているので、普段は押す必要がありません。"
                    "　共有楽曲データと自社CDの台帳を管理番号で引き直し、"
                    "レコード会社名・CD番号・作家名などの空欄を埋めます。"
                    "空いている欄だけを埋め、確認ステータスは「未調査」の行だけ"
                    "「台帳一致」に変えます。人が入れた値は動きません。"
                )
                if st.button("データベースで埋める", key="shinkok_refill_ledger"):
                    _rf_hits = _rf_filled = 0
                    _rf_err = ""
                    for _rf_fn, _rf_name in ((master_fill, "共有楽曲データ"),
                                             (cd_fill, "自社CDの台帳")):
                        try:
                            _rf_df, _rf_h, _rf_f = _rf_fn(st.session_state.songs_df)
                            if _rf_f:
                                st.session_state.songs_df = _rf_df
                                _rf_hits += _rf_h
                                _rf_filled += _rf_f
                        except Exception as _e:
                            _rf_err = f"{_rf_name}の参照に失敗しました: {_e}"
                    if _rf_err:
                        st.warning(f"⚠️ {_rf_err}")
                    elif not _rf_filled:
                        st.info("埋めるものはありませんでした。")
                    if _rf_filled and _autosave_to_db("（データベースで埋め直し）"):
                        st.session_state["_autosave_msg"] = (
                            f"データベースから {_rf_hits} 曲・{_rf_filled} 欄を埋めました"
                        )
                        st.rerun()

            # 昔の書き方で保存された行の付け直し。
            # 直す行があるときだけ出す（普段は畳んだ帯すら出さない）。
            # 勝手に書き換えず、中身を見せてからボタンで直す
            _stale_mask = _stale_multi_mask(_shinkok_songs)
            if _stale_mask is not None and int(_stale_mask.sum()):
                _stale_n = int(_stale_mask.sum())
                with st.expander(f"🧹 古い「候補あり」を付け直す（{_stale_n} 曲）"):
                    st.caption(
                        "一括検索は昔、複数ヒットして自動入力しなかった行にも"
                        "「候補あり」と書いていました。どれか選びに行くべき行なので、"
                        "今は「複数候補あり」と書きます。"
                        "下の行は、何も入っていないので昔の書き方だと分かるものです。"
                    )
                    st.dataframe(
                        _shinkok_songs.loc[
                            _stale_mask,
                            [c for c in ("No", "曲名", "アーティスト", "確認ステータス")
                             if c in _shinkok_songs.columns]
                        ],
                        use_container_width=True, hide_index=True, height=200,
                    )
                    if st.button(f"「複数候補あり」に付け直す（{_stale_n} 曲）",
                                 key="shinkok_fix_stale", type="primary"):
                        st.session_state.songs_df.loc[
                            _stale_mask, "確認ステータス"] = "複数候補あり"
                        # 保存できたときだけ伝言を差し替える。失敗した
                        # ときは _autosave_to_db が入れた警告を残す
                        if _autosave_to_db("（状態の付け直し）"):
                            st.session_state["_autosave_msg"] = (
                                f"{_stale_n} 曲を「複数候補あり」に付け直しました"
                            )
                        st.rerun()

            # CSV ダウンロード。「選択」「状態」は画面を操作するためだけの
            # 列なので、書き出す表からは外す
            _sh_dl_col, _sh_gap = st.columns([2, 3])
            with _sh_dl_col:
                _shinkok_out = _edited_shinkok.drop(
                    columns=[c for c in ("選択", "状態")
                             if c in _edited_shinkok.columns])
                _shinkok_csv = _shinkok_out.to_csv(index=False, encoding="utf-8-sig")
                st.download_button(
                    label="⬇️ 申告フォーマット CSV（編集済み）",
                    data=_shinkok_csv.encode("utf-8-sig"),
                    file_name="申告フォーマット.csv",
                    mime="text/csv",
                    use_container_width=True,
                    # 書き出した内容とDBの中身がずれないよう、
                    # ダウンロードと同時に保存する
                    on_click=_autosave_to_db,
                    args=("（CSV書き出し時）",),
                )


# =====================================================================
# タブ 2: イベント一覧
# =====================================================================
with tabs[1]:
    st.header("イベント一覧（NUENDO イベント単位）")

    if st.session_state.events_df is None:
        st.info("「📋 申告フォーム作成」タブの設定セクションで照合を実行してください。")
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

        # 使用時間（分・秒）が events_df にない旧データを補完
        def _dur_to_min_sec(s):
            s = str(s).strip()
            if not s or s.lower() == "nan":
                return 0, 0
            try:
                parts = s.split(":")
                if len(parts) == 3:
                    t = int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
                elif len(parts) == 2:
                    t = int(parts[0]) * 60 + float(parts[1])
                else:
                    t = float(parts[0])
                t = int(t)
                return t // 60, t % 60
            except (ValueError, TypeError):
                return 0, 0

        if "使用時間（分）" not in filtered_ev.columns:
            if "使用尺" in filtered_ev.columns:
                filtered_ev = filtered_ev.copy()
                filtered_ev[["使用時間（分）", "使用時間（秒）"]] = (
                    filtered_ev["使用尺"]
                    .apply(lambda x: pd.Series(_dur_to_min_sec(x), index=["使用時間（分）", "使用時間（秒）"]))
                )

        st.dataframe(
            filtered_ev,
            use_container_width=True,
            hide_index=True,
            column_config={
                "使用尺":       st.column_config.TextColumn("使用尺", width="small"),
                "使用時間（分）": st.column_config.NumberColumn("分", width="small", format="%d"),
                "使用時間（秒）": st.column_config.NumberColumn("秒", width="small", format="%d"),
                "START TIME":  st.column_config.TextColumn("START TIME", width="small"),
                "終了時間":     st.column_config.TextColumn("終了時間",   width="small"),
                "イン時間":     st.column_config.TextColumn("イン時間",   width="small"),
                "アウト時間":   st.column_config.TextColumn("アウト時間", width="small"),
            },
        )


# =====================================================================
# 🗃️ 共有楽曲データ（ステップ3）
# =====================================================================
with tabs[2]:
    st.header("🗃️ 共有楽曲データ")
    st.caption(
        "案件をまたいで貯めている楽曲データです。ここで直した値は"
        "「手入力」として残るので、あとから機械が調べ直しても"
        "上書きされません。"
    )

    # 自社CDの台帳（TSP）はここでは直せない読み取り専用の資料。
    # 入っているかどうかだけ分かるようにしておく
    _cd_total = _cached_count("cd", cd_count, _CD_COUNT_TTL)
    if _cd_total:
        st.caption(
            f"💿 自社CDの台帳: {_cd_total:,} 曲（読み取り専用）。"
            "照合のとき、管理番号が一致した曲の空欄をここから埋めます。"
        )

    _total = _cached_count("master", master_count)
    if not _total:
        st.info(
            "まだ何も貯まっていません。照合して「確定」「作曲者一致」"
            "「アーティスト一致」になった曲が、保存のときに貯まります。"
        )
    else:
        _c_kw, _c_n = st.columns([4, 1])
        with _c_kw:
            _kw = st.text_input(
                "検索",
                key="master_kw",
                placeholder="曲名 / 管理番号 / 作曲者 / アーティスト …",
                label_visibility="collapsed",
            )
        with _c_n:
            st.markdown(f"全 **{_total}** 曲")

        _LIMIT = 300
        _recs = master_search(_kw, limit=_LIMIT)

        if not _recs:
            st.warning("見つかりませんでした。別の言葉で探してください。")
        else:
            _view = master_frame(_recs)
            if len(_recs) >= _LIMIT:
                # 黙って切らない。全部出ていると思われると困るため
                st.caption(
                    f"多いので新しい順に {_LIMIT} 件まで出しています。"
                    "検索で絞ってください。"
                )

            _sel = st.dataframe(
                _view.drop(columns=["id"]),
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="multi-row",
                key="master_table",
                column_config={
                    "曲名":         st.column_config.TextColumn("曲名", width="medium"),
                    "管理番号キー":  st.column_config.TextColumn("管理番号キー", width="small"),
                    "トラックキー":  st.column_config.TextColumn("トラックキー", width="small"),
                    "更新":         st.column_config.TextColumn("更新", width="small"),
                },
            )

            _rows = list(_sel.selection.rows) if _sel and _sel.selection else []
            _ids = [int(_view.iloc[i]["id"]) for i in _rows]

            if not _ids:
                st.caption("左端をクリックして選ぶと、直したり消したりできます。")
            else:
                _by_id = {int(r["id"]): r for r in _recs}

                # ── 直す（1件だけ選んだとき）──────────────────
                if len(_ids) == 1:
                    _rec = _by_id[_ids[0]]
                    _src_line = master_sources(_rec)
                    with st.expander(
                        f"✏️ 「{_rec.get('title') or '（曲名なし）'}」を直す",
                        expanded=True,
                    ):
                        st.caption(
                            f"最終更新: {_rec.get('updated_at') or '—'}"
                            + (f" ／ 出典: {_src_line}" if _src_line else "")
                        )
                        with st.form(f"master_edit_{_ids[0]}"):
                            _vals = {
                                "曲名": st.text_input(
                                    "曲名", value=_rec.get("title") or "",
                                    key=f"me_title_{_ids[0]}",
                                )
                            }
                            _cols = st.columns(2)
                            for _i, _f in enumerate(MASTER_FIELDS):
                                _c = master_cell(_rec, _f)
                                _help = "まだ何も入っていません"
                                if _c["v"]:
                                    _help = f"出典: {_c['src'] or '不明'}"
                                    if _c["by"]:
                                        _help += f" ／ 入れた人: {_c['by']}"
                                with _cols[_i % 2]:
                                    _vals[_f] = st.text_input(
                                        _f, value=_c["v"], help=_help,
                                        key=f"me_{_f}_{_ids[0]}",
                                    )
                            _save = st.form_submit_button(
                                "💾 この曲を保存", type="primary",
                                use_container_width=True,
                            )
                        if _save:
                            _n = master_edit(_rec, _vals, CURRENT_USER)
                            if _n:
                                st.success("✅ 保存しました。")
                                st.rerun()
                            else:
                                st.info("変わったところがありません。")

                # ── 消す ────────────────────────────────────
                st.divider()
                st.markdown(f"🗑️ **{len(_ids)} 曲**を選んでいます")
                _ok = st.checkbox(
                    "消すと元に戻せません。全員の共有データから消えます。",
                    key="master_del_ok",
                )
                if st.button("🗑️ 選んだ曲を消す", disabled=not _ok,
                             key="master_del_btn"):
                    _n = master_delete(_ids)
                    if _n:
                        _forget_count("master")   # 減った分をすぐ出す
                        st.success(f"🗑️ {_n} 曲を消しました。")
                        # 選択が残っていると消えた行を指したままになる。
                        # 代入ではなく削除にすること。ウィジェットを作った
                        # あとで代入すると Streamlit が例外を出す
                        st.session_state.pop("master_table", None)
                        st.session_state.pop("master_del_ok", None)
                        st.rerun()
                    else:
                        st.error("❌ 消せませんでした。")


# =====================================================================
# タブ 4: 検索補助
# =====================================================================
with tabs[0]:
    st.divider()
    st.subheader("🔍 補完検索（J-WID / MINC / NexTone）")

    if st.session_state.songs_df is None:
        st.info("上の設定セクションで照合を実行してください。")
    else:
        songs_df = st.session_state.songs_df

        # ---- 楽曲選択 ----
        _tab4_status_filter = st.radio(
            "表示するステータス",
            ["すべて", "未確定のみ"],
            horizontal=True,
            key="tab4_status_filter",
            help="「未確定のみ」は未調査・該当なし・作曲者一致・アーティスト一致・候補あり・複数候補あり・MP3補助確認のみ表示します。",
        )
        if _tab4_status_filter == "未確定のみ":
            _tab4_df = songs_df[songs_df["確認ステータス"].isin(
                ["未調査", "該当なし", "作曲者一致", "アーティスト一致", "候補あり", "複数候補あり", "MP3補助確認", "J-WID要確認", "NexTone要確認", "要確認"]
            )]
        else:
            _tab4_df = songs_df
        song_labels = (
            _tab4_df["No"].astype(str) + ". ["
            + _tab4_df["確認ステータス"] + "] "
            + _tab4_df["イベント名"]
        ).tolist()
        # 反映でステータスが変わるとラベル文字列も変わり選択がリセットされるため、
        # 古い選択ラベルを No. で探し直して最新ラベルに更新する
        _stored_sel = st.session_state.get("search_song_select")
        if _stored_sel and _stored_sel not in song_labels:
            try:
                _stored_no = int(_stored_sel.split(".")[0])
                _matched = next((l for l in song_labels if int(l.split(".")[0]) == _stored_no), None)
                if _matched:
                    st.session_state["search_song_select"] = _matched
            except (ValueError, StopIteration):
                pass
        selected_label = st.selectbox(
            f"調査する楽曲を選択（{len(_tab4_df)} 件）",
            options=song_labels,
            key="search_song_select",
        )

        if selected_label:
            selected_no = int(selected_label.split(".")[0])
            row_idx = songs_df[songs_df["No"] == selected_no].index[0]
            row = songs_df.loc[row_idx]

            # 選択曲名 + セクション移動ボタン
            _sel_song_name = str(row.get("曲名") or row.get("イベント名", "")).strip()
            if _sel_song_name and _sel_song_name.lower() != "nan":
                st.caption(f"🎵 {_sel_song_name}")
            st.markdown(
                """
                <div style="display:flex;gap:8px;margin:4px 0 8px">
                  <a href="#sec-minc-individual"
                     style="flex:1;text-align:center;padding:7px 4px;
                            background:#1565C0;color:#fff;border-radius:6px;
                            text-decoration:none;font-size:13px;font-weight:600">
                    🌲 MINC 楽曲検索（個別）
                  </a>
                  <a href="#sec-cd-search"
                     style="flex:1;text-align:center;padding:7px 4px;
                            background:#1565C0;color:#fff;border-radius:6px;
                            text-decoration:none;font-size:13px;font-weight:600">
                    💿 CD情報検索
                  </a>
                  <a href="#sec-pipeline"
                     style="flex:1;text-align:center;padding:7px 4px;
                            background:#1565C0;color:#fff;border-radius:6px;
                            text-decoration:none;font-size:13px;font-weight:600">
                    🔄 全自動パイプライン
                  </a>
                </div>
                """,
                unsafe_allow_html=True,
            )

            # 選択曲が変わったとき、Excel確認タブのフィルターを自動更新
            if st.session_state.get("_ex_auto_from_no") != selected_no:
                st.session_state["_ex_auto_from_no"] = selected_no
                _auto_composer = str(row.get("作曲者", "")).strip()
                _auto_artist   = str(row.get("アーティスト", "")).strip()
                if _auto_composer.lower() in ("", "nan"):
                    _auto_composer = ""
                if _auto_artist.lower() in ("", "nan"):
                    _auto_artist = ""
                if _auto_composer:
                    st.session_state["ex_filter_composer"] = _auto_composer
                    st.session_state["ex_filter_artist"]   = ""
                elif _auto_artist:
                    st.session_state["ex_filter_composer"] = ""
                    st.session_state["ex_filter_artist"]   = _auto_artist

            # 検索語を収集（優先度順・重複除去）
            term_candidates: list[tuple[str, str]] = []
            _tc_seen: set[str] = set()
            for field, label in [
                ("WAV検出タイトル", "WAV検出タイトル"),
                ("曲名",           "曲名"),
                ("アーティスト",   "アーティスト"),
                ("ライブラリ盤番号", "ライブラリ盤番号"),
                ("CD番号",         "CD番号"),
                ("イベント名",     "イベント名"),
            ]:
                val = str(row.get(field, "")).strip()
                if val and val.lower() != "nan" and val not in _tc_seen:
                    term_candidates.append((label, val))
                    _tc_seen.add(val)
            if not term_candidates:
                term_candidates.append(("イベント名", str(row.get("イベント名", "")).strip()))

            # col_terms の text_input で編集済みであればその値を使う
            _main_label = term_candidates[0][0]
            main_term = st.session_state.get(
                f"term_{selected_no}_{_main_label}", term_candidates[0][1]
            )
            encoded = urllib.parse.quote(main_term)

            st.subheader(f"No.{selected_no} ／ {row.get('イベント名', '')}")
        # ---- MINC 楽曲検索（個別・保険）----
        st.divider()
        st.markdown('<a id="sec-minc-individual"></a>', unsafe_allow_html=True)
        st.markdown("#### 🌲 MINC 楽曲検索（個別）")
        st.caption(
            "まずここで検索して候補を確認してください。"
            "　minc.or.jp にログイン済みの Cookie を使って作曲者・作詞者・JASRAC コード・CD情報・I/V区分・委任者を取得します。"
        )

        # 認証状態バー（アプリ全体で1つのキーを使いまわす）
        _mf_auth_col, _mf_btn_col = st.columns([4, 1])
        with _mf_btn_col:
            _mf_check = st.button("🔄 認証確認", key=f"mf_check_{selected_no}", use_container_width=True)
        if _mf_check:
            st.session_state.pop("mf_auth_state", None)
            st.session_state.pop("mf_client", None)  # 再ログイン後は新クライアントを作成
        if "mf_auth_state" not in st.session_state:
            try:
                _mf_ok, _mf_msg = check_session(_get_mf_client())
            except MusicForestError as _e:
                _mf_ok, _mf_msg = False, str(_e)
            st.session_state["mf_auth_state"] = (_mf_ok, _mf_msg)
        _mf_ok, _mf_msg = st.session_state["mf_auth_state"]
        with _mf_auth_col:
            if _mf_ok:
                st.success(f"✅ {_mf_msg}")
            else:
                st.warning(
                    f"⚠️ {_mf_msg}\n\n"
                    f"ログイン: `.venv\\Scripts\\python.exe "
                    f"H:\\PROGRAM\\search_music\\src\\login_browser.py`"
                )

        # アーティスト絞り込み（検索結果をこの名前で絞る。検索前に入れておける）
        _mf_art_key = f"mf_artist_{selected_no}"
        if _mf_art_key not in st.session_state:
            _mf_row_art = str(row.get("アーティスト", "")).strip()
            st.session_state[_mf_art_key] = "" if _is_blank(_mf_row_art) else _mf_row_art

        if _mf_ok:
            _mf_s1, _mf_s2, _mf_s2b, _mf_s3 = st.columns([3, 2, 2, 1])
            with _mf_s1:
                _mf_term_opts = [f"[{lbl}]  {val}" for lbl, val in term_candidates]
                _mf_term_sel = st.selectbox(
                    "検索語候補",
                    options=_mf_term_opts,
                    key=f"mf_title_{selected_no}",
                )
                _mf_title_val = term_candidates[_mf_term_opts.index(_mf_term_sel)][1]
            with _mf_s2:
                # 他セクション（パイプライン等）から著作者名を流し込む場合は
                # ウィジェット生成「前」のここで反映する
                # （生成後に session_state を書き換えると StreamlitAPIException）
                _mf_author_pending = st.session_state.pop(f"mf_author_pending_{selected_no}", None)
                if _mf_author_pending is not None:
                    st.session_state[f"mf_author_{selected_no}"] = _mf_author_pending
                _mf_author_val = st.text_input(
                    "著作者名（任意・絞り込み用）",
                    value="",
                    key=f"mf_author_{selected_no}",
                    placeholder=str(row.get("作曲者", "")).strip() or "例: 加藤達也",
                )
            with _mf_s2b:
                st.text_input(
                    "アーティスト（任意・絞り込み用）",
                    key=_mf_art_key,
                    placeholder="例: ZOO",
                    help=(
                        "MINCの検索結果をこのアーティスト名で絞り込みます（部分一致）。"
                        "楽曲まとめのアーティストを初期値に入れています。"
                        "一致する候補が無いときは全件表示に戻します。"
                    ),
                )
            with _mf_s3:
                _mf_match = st.selectbox(
                    "一致方式",
                    options=["2: 前方一致", "3: キーワード", "1: 完全一致"],
                    key=f"mf_match_{selected_no}",
                    help="match=1 は MINC 側でキーワード扱いになり別の曲が返ることがあります。前方一致が最も安定します。",
                )
            _mf_match_int = int(_mf_match[0])

            # 選択語を編集できるinput（候補が変わったときリセット）
            _mf_edit_key = f"mf_term_edit_{selected_no}"
            _mf_prev_key = f"mf_term_prev_{selected_no}"
            if st.session_state.get(_mf_prev_key) != _mf_title_val:
                if _mf_edit_key in st.session_state:
                    del st.session_state[_mf_edit_key]
                st.session_state[_mf_prev_key] = _mf_title_val
            _mf_search_term = st.text_input(
                "検索語（編集可）",
                key=_mf_edit_key,
                value=_mf_title_val,
                help="候補から自動入力。不要な語を削除するなど自由に編集できます。",
            )

            if st.button(
                f"🌲 MINC で「{(_mf_search_term or _mf_title_val)[:20]}」を検索",
                key=f"mf_search_{selected_no}",
                type="primary",
                use_container_width=True,
            ):
                with st.spinner("MINC を検索中... （1.5秒/リクエスト）"):
                    try:
                        _mf_client = _get_mf_client()
                        _mf_result = _mf_client.search(
                            _mf_search_term or _mf_title_val,
                            author=_mf_author_val,
                            match=_mf_match_int,
                        )
                        st.session_state[f"mf_results_{selected_no}"] = _mf_result
                    except MusicForestError as e:
                        st.session_state[f"mf_results_{selected_no}"] = {"error": str(e), "results": []}

        # ---- MusicForest 検索結果表示 ----
        _mf_res = st.session_state.get(f"mf_results_{selected_no}")
        if _mf_res:
            if _mf_res.get("error"):
                st.error(f"MINC エラー: {_mf_res['error']}")
                with st.expander("デバッグ HTML"):
                    st.code(_mf_res.get("debug_html", "")[:3000], language="html")
            elif not _mf_res.get("results"):
                st.warning("MusicForest: 該当なし")
                with st.expander("デバッグ HTML"):
                    st.code(_mf_res.get("debug_html", "")[:3000], language="html")
            else:
                _mf_items = _mf_res["results"]
                if _mf_res.get("truncated"):
                    st.warning("⚠️ 検索結果が 500件上限に達しました。検索語を絞り込んでください。")
                st.success(f"🌲 MINC: {len(_mf_items)} 件見つかりました")
                st.caption(f"検索URL: {_mf_res.get('search_url','')}")

                # MINC詳細取得・J-WID取得の結果を永続バナーで表示（rerun後も消えない）
                _mf_flash_key = f"mf_flash_{selected_no}"
                if _mf_flash_key in st.session_state:
                    _fl = st.session_state.pop(_mf_flash_key)
                    if _fl.get("error"):
                        st.error(_fl["error"])
                    else:
                        st.success(_fl["message"])

                # ── アーティストで絞り込み ────────────────────────────────
                # 候補番号・session_state キーは元の並び順のインデックスを使い続ける
                # （絞り込みで番号がずれると取得済みの詳細が別候補に付いてしまう）
                _mf_pairs = list(enumerate(_mf_items))
                _mf_art_q = _mf_norm_name(st.session_state.get(_mf_art_key, ""))
                if _mf_art_q:
                    _mf_hit = [
                        (_i, _it) for _i, _it in _mf_pairs
                        if _mf_art_q in _mf_norm_name(_it.get("アーティスト", ""))
                    ]
                    if _mf_hit:
                        st.caption(
                            f"🎤 アーティスト「{st.session_state[_mf_art_key].strip()}」で絞り込み: "
                            f"**{len(_mf_hit)}** / {len(_mf_pairs)} 件"
                        )
                        _mf_pairs = _mf_hit
                    else:
                        st.caption(
                            f"🎤 アーティスト「{st.session_state[_mf_art_key].strip()}」に一致する候補が"
                            "無いため全件表示しています（表記ゆれの可能性があります）"
                        )

                for _mf_disp_i, (_mf_i, _mf_item) in enumerate(_mf_pairs[:20]):
                    _mf_label = (
                        f"候補{_mf_i+1} [{_mf_item['_source_table']}]: "
                        f"{_mf_item.get('作品名','')} ／ {_mf_item.get('アーティスト','')} "
                        f"  JASRAC:{_mf_item.get('JASRAC作品コード','(なし)')}  "
                        f"NexTone:{_mf_item.get('NexTone管理番号','(なし)')}"
                    )
                    with st.expander(_mf_label, expanded=(_mf_disp_i == 0), key=f"mf_exp_{selected_no}_{_mf_i}"):
                        _mf_c1, _mf_c2 = st.columns(2)
                        _mf_c1.text_input("作品名",         value=_mf_item.get("作品名",""),          key=f"mf_name_{selected_no}_{_mf_i}", disabled=True)
                        _mf_c1.text_input("アーティスト",   value=_mf_item.get("アーティスト",""),    key=f"mf_art_{selected_no}_{_mf_i}",  disabled=True)
                        _mf_c1.text_input("品番（CD番号）",  value=_mf_item.get("品番",""),            key=f"mf_cat_{selected_no}_{_mf_i}",  disabled=True)
                        _mf_c1.text_input("CD商品タイトル",  value=_mf_item.get("CD商品タイトル",""),  key=f"mf_cdtitle_{selected_no}_{_mf_i}", disabled=True)
                        _mf_c2.text_input("JASRAC作品コード", value=_mf_item.get("JASRAC作品コード",""), key=f"mf_jcd_{selected_no}_{_mf_i}",  disabled=True)
                        _mf_c2.text_input("NexTone管理番号", value=_mf_item.get("NexTone管理番号",""), key=f"mf_ncd_{selected_no}_{_mf_i}",  disabled=True)
                        _mf_c2.text_input("レコード会社名",  value=_mf_item.get("レコード会社名",""),  key=f"mf_label_{selected_no}_{_mf_i}", disabled=True)
                        _mf_c2.text_input("発売会社／販売会社（生）", value=_mf_item.get("発売会社販売会社",""), key=f"mf_pub_{selected_no}_{_mf_i}", disabled=True)

                        # キー定義（ハンドラ・詳細フィールド共用）
                        _mf_detail_key = f"mf_detail_{selected_no}_{_mf_i}"
                        _jwid_minc_key = f"mf_jwid_{selected_no}_{_mf_i}"
                        _mf_delg_key   = f"mf_delg_{selected_no}_{_mf_i}"
                        _mf_jcd    = _mf_item.get("JASRAC作品コード", "")
                        _mf_alb_id = _mf_item.get("_album_id", "")
                        _mf_trk_id = _mf_item.get("_track_id", "")

                        # ---- ボタン列（詳細フィールドより先に描画）----
                        # ハンドラが session_state を更新した後に下の詳細フィールドが描画されるため
                        # st.rerun() 不要でデータが即時反映される
                        _mf_btn1, _mf_btn_jwid, _mf_btn2 = st.columns(3)

                        with _mf_btn1:
                            if st.button(
                                "💿 MINC CD情報取得",
                                key=f"mf_detail_btn_{selected_no}_{_mf_i}",
                                use_container_width=True,
                                help=(
                                    "アルバムIDのある候補（収録曲テーブル）のみI/V区分・委任者を取得できます。\n"
                                    "アルバムIDがない候補（配信曲・作品テーブル）は下の「💿 CD情報取得」セクションをご利用ください。"
                                ),
                            ):
                                with st.spinner("MINC 情報取得中..."):
                                    # Step 1: 作品詳細（作曲者/作詞者）
                                    _d_err = None
                                    try:
                                        _d = _get_mf_client().get_detail(_mf_item["_detail_href"])
                                        st.session_state[_mf_detail_key] = _d
                                        if _d.get("error"):
                                            _d_err = _d["error"]
                                    except MusicForestError as _e:
                                        _d_err = str(_e)
                                        st.session_state[_mf_detail_key] = {"error": _d_err}

                                    # Step 2: CD詳細（I/V区分・委任者）
                                    if _mf_alb_id and _mf_trk_id:
                                        try:
                                            _cd = _get_mf_client().fetch_product_detail(_mf_alb_id, _mf_trk_id)
                                            st.session_state[_mf_delg_key] = _cd
                                            _iv_label = {"I": "インスト", "V": "ヴォーカル"}.get(_cd.get("IV", ""), "不明")
                                            _cd_err = _cd.get("error")
                                            if _cd_err:
                                                st.toast(f"CD情報エラー: {_cd_err}", icon="❌")
                                            elif _d_err:
                                                st.toast(
                                                    f"CD情報取得完了（候補{_mf_i+1}）: I/V={_iv_label} / 委任者={_cd.get('集中管理','不明')} ※作品詳細エラー: {_d_err}",
                                                    icon="⚠️",
                                                )
                                            else:
                                                _d_data = st.session_state.get(_mf_detail_key, {})
                                                st.toast(
                                                    f"取得完了（候補{_mf_i+1}）: "
                                                    f"作曲者={_d_data.get('作曲者','') or '(なし)'}　"
                                                    f"I/V={_iv_label}　"
                                                    f"委任者={_cd.get('集中管理','不明')}",
                                                    icon="✅",
                                                )
                                        except MusicForestError as _e:
                                            st.session_state[_mf_delg_key] = {"集中管理": "", "IV": "", "error": str(_e)}
                                            st.toast(f"CD情報取得エラー: {_e}", icon="❌")
                                    else:
                                        # 収録CD情報なし（作品テーブルの場合）→ CD情報取得不可
                                        if _d_err:
                                            st.toast(f"MINC詳細エラー: {_d_err}", icon="❌")
                                        else:
                                            _d_data = st.session_state.get(_mf_detail_key, {})
                                            st.toast(
                                                f"作家情報取得完了（候補{_mf_i+1}）: "
                                                f"作曲者={_d_data.get('作曲者','') or '(なし)'}　"
                                                f"※CD情報（I/V・委任者）は収録CD情報がないため取得できません",
                                                icon="⚠️",
                                            )

                        with _mf_btn_jwid:
                            if st.button(
                                "📋 J-WID作家情報",
                                key=f"mf_jwid_btn_{selected_no}_{_mf_i}",
                                use_container_width=True,
                                disabled=not _mf_jcd,
                                help=f"JASRACコード {_mf_jcd} でJ-WIDを直接引き当て。作家情報＋管理状況を取得します",
                            ):
                                with st.spinner("J-WID から取得中..."):
                                    from modules.scraper import fetch_jwid_rights_by_code as _fetch_by_code
                                    _jw = _fetch_by_code(_mf_jcd)
                                    st.session_state[_jwid_minc_key] = _jw
                                    if _jw.get("error"):
                                        st.toast(f"J-WID エラー: {_jw['error']}", icon="❌")
                                    else:
                                        st.toast(
                                            f"J-WID 取得完了（候補{_mf_i+1}）: "
                                            f"作曲者: {_jw['作曲者'] or '(なし)'}　"
                                            f"作詞者: {_jw['作詞者'] or '(なし)'}　"
                                            f"訳詞者: {_jw['訳詞者'] or '(なし)'}", icon="✅"
                                        )

                        with _mf_btn2:
                            if st.button(
                                "✅ 申告フォーマットに反映",
                                key=f"mf_apply_{selected_no}_{_mf_i}",
                                use_container_width=True,
                            ):
                                _jw_d = st.session_state.get(_jwid_minc_key, {})
                                _detail_now = st.session_state.get(_mf_detail_key, {})
                                # 作曲者・作詞者が未取得なら MINC 詳細を自動フェッチ
                                if _mf_item.get("_detail_href") and not (
                                        _jw_d.get("作曲者") or _detail_now.get("作曲者") or
                                        _jw_d.get("作詞者") or _detail_now.get("作詞者")):
                                    try:
                                        _ad = _get_mf_client().get_detail(_mf_item["_detail_href"])
                                        if not _ad.get("error"):
                                            st.session_state[_mf_detail_key] = _ad
                                            _detail_now = _ad
                                    except Exception:
                                        pass
                                # J-WID を優先、なければ MINC 詳細、
                                # それも無ければ検索結果の「作詞／作曲」列
                                _composer   = _jw_d.get("作曲者") or _detail_now.get("作曲者") or _mf_item.get("作曲者", "")
                                _lyricist   = _jw_d.get("作詞者") or _detail_now.get("作詞者") or _mf_item.get("作詞者", "")
                                _translator = _jw_d.get("訳詞者") or _detail_now.get("訳詞者") or _mf_item.get("訳詞者", "")
                                _arranger   = _jw_d.get("編曲者") or _detail_now.get("編曲者") or _mf_item.get("編曲者", "")
                                _delg_r = st.session_state.get(_mf_delg_key, {})
                                _委任者 = _delg_r.get("集中管理","")
                                _iv_raw = _delg_r.get("IV","")
                                _iv_apply = {"I": "インスト", "V": "ヴォーカル"}.get(_iv_raw, "")
                                _mf_jcd2 = _mf_item.get("JASRAC作品コード","") or _detail_now.get("作品コード","")
                                # JASRACコードが変わる場合は先に関連フィールドをクリア
                                _apply_clear_on_jcd_change(row_idx, _mf_jcd2)
                                _mf_apply = {
                                    "曲名":            _mf_item.get("作品名",""),
                                    "作曲者":          _composer,
                                    "作詞者":          _lyricist,
                                    "訳詞者":          _translator,
                                    "編曲者":          _arranger,
                                    "アーティスト":    _mf_item.get("アーティスト","") or _delg_r.get("アーティスト",""),
                                    "CD番号":          _mf_item.get("品番","") or _delg_r.get("品番",""),
                                    "CD名":            _mf_item.get("CD商品タイトル","") or _delg_r.get("CD商品タイトル",""),
                                    "レコード会社名":  _mf_item.get("レコード会社名",""),
                                    "JASRAC作品コード": _mf_jcd2,
                                    "NexTone管理番号": _mf_item.get("NexTone管理番号","") or _detail_now.get("NexTone管理番号",""),
                                    "委任者":          _委任者,
                                    "確認ステータス":  "候補あり",
                                }
                                _apply_management_status(
                                    _jw_d.get("管理状況") or {}, _mf_apply,
                                    "J", st.session_state.songs_df.loc[row_idx])
                                _hy = _infer_houyo(_mf_jcd2)
                                _cur_hy_mf = str(st.session_state.songs_df.at[row_idx, "邦洋区分"] if "邦洋区分" in st.session_state.songs_df.columns else "").strip()
                                if _hy and not _cur_hy_mf:
                                    _mf_apply["邦洋区分"] = _hy
                                _cur_iv_mf = str(st.session_state.songs_df.at[row_idx, "I/V区分"] if "I/V区分" in st.session_state.songs_df.columns else "").strip()
                                if not _iv_apply and not _is_blank(_composer + _lyricist + _arranger):
                                    # CD詳細にI/V表記が無ければ作詞者の有無で判定
                                    _iv_apply = _infer_iv(_lyricist)
                                if _iv_apply and not _cur_iv_mf:
                                    _mf_apply["I/V区分"] = _iv_apply
                                for _col, _val in _mf_apply.items():
                                    if _val and _col in st.session_state.songs_df.columns:
                                        st.session_state.songs_df.at[row_idx, _col] = _val
                                st.session_state["_apply_msg"] = "楽曲まとめ・申告フォーマットに反映しました。"
                                st.session_state.pop("songs_editor", None)
                                st.rerun()

                        # ---- 詳細フィールド（ボタンハンドラ実行後に描画 → 取得データが即反映）----

                        _mf_detail = st.session_state.get(_mf_detail_key, {})
                        _mf_dc1, _mf_dc2, _mf_dc3 = st.columns(3)
                        _mf_dc1.text_input("作曲者（MINC詳細）", value=_mf_detail.get("作曲者",""), key=f"mf_comp_{selected_no}_{_mf_i}", disabled=True, placeholder="詳細取得で確認")
                        _mf_dc2.text_input("作詞者（MINC詳細）", value=_mf_detail.get("作詞者",""), key=f"mf_lyric_{selected_no}_{_mf_i}", disabled=True, placeholder="詳細取得で確認")
                        _mf_dc3.text_input("編曲者（MINC詳細）", value=_mf_detail.get("編曲者",""), key=f"mf_arr_{selected_no}_{_mf_i}", disabled=True)

                        # J-WID 直接引き当て（MINCのJASRACコードを使用）
                        _jwid_minc = st.session_state.get(_jwid_minc_key, {})
                        if _jwid_minc and not _jwid_minc.get("error"):
                            _jw_c1, _jw_c2, _jw_c3, _jw_c4 = st.columns(4)
                            _jw_c1.text_input("作曲者（J-WID）", value=_jwid_minc.get("作曲者",""), key=f"mf_j_comp_{selected_no}_{_mf_i}", disabled=True)
                            _jw_c2.text_input("作詞者（J-WID）", value=_jwid_minc.get("作詞者",""), key=f"mf_j_lyric_{selected_no}_{_mf_i}", disabled=True)
                            _jw_c3.text_input("訳詞者（J-WID）", value=_jwid_minc.get("訳詞者",""), key=f"mf_j_tran_{selected_no}_{_mf_i}", disabled=True)
                            _jw_c4.text_input("編曲者（J-WID）", value=_jwid_minc.get("編曲者",""), key=f"mf_j_arr_{selected_no}_{_mf_i}", disabled=True)
                            _mgmt_minc = _jwid_minc.get("管理状況", {})
                            if _mgmt_minc:
                                st.markdown("**管理状況（JASRAC）:**  \n" + _format_management_status(_mgmt_minc))
                        elif _jwid_minc.get("error"):
                            st.error(f"J-WID 取得エラー: {_jwid_minc['error']}")

                        # CD情報取得結果（I/V区分・委任者・CDメタデータ）
                        _mf_delg = st.session_state.get(_mf_delg_key, {})
                        if _mf_delg:
                            if _mf_delg.get("error"):
                                st.error(f"CD情報取得エラー: {_mf_delg['error']}")
                            else:
                                _iv_raw_d = _mf_delg.get("IV", "")
                                _iv_disp = {"I": "インスト", "V": "ヴォーカル"}.get(_iv_raw_d, "")
                                _delg_status = _mf_delg.get("集中管理", "")
                                _cd_info_c1, _cd_info_c2 = st.columns(2)
                                _cd_info_c1.text_input(
                                    "I/V区分（MINC）",
                                    value=_iv_disp or "（取得できず）",
                                    key=f"mf_iv_disp_{selected_no}_{_mf_i}",
                                    disabled=True,
                                )
                                _cd_info_c2.text_input(
                                    "委任者（MINC）",
                                    value=_delg_status or "（取得できず）",
                                    key=f"mf_delg_disp_{selected_no}_{_mf_i}",
                                    disabled=True,
                                )
                                # CDタイトル・トラック情報（fetch_product_detail 拡張版で取得した場合）
                                _d_cd_title  = _mf_delg.get("CD商品タイトル", "")
                                _d_cd_cat    = _mf_delg.get("品番", "")
                                _d_cd_trk    = _mf_delg.get("トラック番号", "")
                                _d_cd_name   = _mf_delg.get("曲名", "")
                                _d_cd_dur    = _mf_delg.get("尺", "")
                                _d_cd_art    = _mf_delg.get("アーティスト", "")
                                _d_cd_rec_co = _mf_delg.get("レコード会社名", "")
                                if _d_cd_title or _d_cd_trk or _d_cd_name:
                                    _cd_meta_parts = []
                                    if _d_cd_title:
                                        _cd_meta_parts.append(f"**CD:** {_d_cd_title}" + (f"（{_d_cd_cat}）" if _d_cd_cat else ""))
                                    if _d_cd_rec_co:
                                        _cd_meta_parts.append(f"**レコード会社:** {_d_cd_rec_co}")
                                    if _d_cd_art:
                                        _cd_meta_parts.append(f"**アーティスト:** {_d_cd_art}")
                                    if _d_cd_trk or _d_cd_name:
                                        _cd_meta_parts.append(
                                            f"**トラック{_d_cd_trk}:** {_d_cd_name}" + (f"（{_d_cd_dur}）" if _d_cd_dur else "")
                                        )
                                    st.markdown("  \n".join(_cd_meta_parts))
                                if _delg_status == "委任者":
                                    st.success(f"※集中管理: **{_delg_status}**（送信可能化権が日本レコード協会に集中管理委任済み）")
                                elif _delg_status == "非委任者":
                                    st.warning(f"※集中管理: **{_delg_status}**（送信可能化権は集中管理されていません）")

                        # CD情報検索パネル（JASRACコードで収録CDリストを取得）
                        if _mf_jcd:
                            st.divider()
                            _show_cd_panel(
                                _mf_jcd, row_idx, f"mf_{selected_no}_{_mf_i}",
                                title=_mf_item.get("作品名", ""),
                            )

                        st.link_button(
                            "🌲 MINC で詳細を確認",
                            f"https://www.minc.or.jp/saku/detail/?{_mf_item['_detail_href']}",
                            use_container_width=True,
                        )
                        if _mf_item.get("_row_html"):
                            with st.expander("🔍 デバッグ: 検索結果行の HTML（album_id が取れない場合に確認）"):
                                st.code(_mf_item["_row_html"], language="html")

                # ---- CD情報取得（収録CD情報なし候補向け） ----
                _no_cd_indices = [i for i, it in enumerate(_mf_items[:20]) if not it.get("_album_id")]
                _page_cd_links = _mf_res.get("_page_cd_links", [])

                if _no_cd_indices:
                    _exp_label = "💿 CD情報取得" if _page_cd_links else "💿 CD情報 手動取得（URLから直接指定）"
                    with st.expander(_exp_label, expanded=bool(_page_cd_links)):
                        # 2つのボタンが共有する変数（どちらかのボタンが押されたときにセットされる）
                        _m_alb = ""
                        _m_trk = ""
                        _man_target_idx = None
                        _m_cd_result = None
                        _m_fetch_err = ""

                        _man_cand_opts = [f"候補{i+1}: {_mf_items[i].get('作品名','')}" for i in _no_cd_indices]

                        # ── ① ページ内CDリスト（推奨） ──────────────────────────
                        if _page_cd_links:
                            st.markdown("##### 📋 検索結果ページ内のCDリスト")
                            st.caption(
                                "MINCの検索結果ページに出ていたCDリンクを全部拾った"
                                "ものです。候補に紐付けてください。"
                            )
                            # 件数だけ見て「こちらの方が揃っている」と思われがちだが、
                            # 中身は検索に引っかかった別作品のCDも混ざった寄せ集めで、
                            # しかもページ自体が500行で切られる。実測では「ハナミズキ」
                            # の検索で 500件中この作品のCDは402件、作品コードから
                            # 引けば459件だった（＝こちらの方が57枚足りない）
                            st.warning(
                                "⚠️ このリストは**同名・別の曲のCDも混ざります**"
                                "（検索結果ページに出ていたCDを全部拾っているため）。"
                                "その曲の収録CDを漏れなく見るには、候補の"
                                "「🔍 このJASRACコードのCDを検索（MINC）」の方を"
                                "使ってください。ここは作品コードが取れない候補の"
                                "ための逃げ道です。"
                            )
                            if _mf_res.get("truncated"):
                                st.caption(
                                    "※MINCの検索結果が500行で打ち切られているため、"
                                    "このリストにも入り切っていないCDがあります。"
                                    "曲名や作者名で検索を絞ると拾える件数が増えます。"
                                )
                            # 数が多いと選びにくいので、CD名の一部で絞れるようにする
                            _pg_q = st.text_input(
                                "絞り込み（CD名の一部）",
                                key=f"mf_pg_q_{selected_no}",
                                placeholder="例: ベスト / TOCT",
                            ).strip()
                            _pg_pool = _page_cd_links
                            if _pg_q:
                                _pg_ql = _pg_q.lower()
                                _pg_hit = [lnk for lnk in _page_cd_links
                                           if _pg_ql in lnk["label"].lower()]
                                if _pg_hit:
                                    st.caption(
                                        f"「{_pg_q}」で絞り込み: "
                                        f"**{len(_pg_hit)}** / {len(_page_cd_links)} 件")
                                    _pg_pool = _pg_hit
                                else:
                                    # 0件のまま選ばせても何も選べないので全件に戻す
                                    st.caption(
                                        f"「{_pg_q}」に一致するCDが無いため"
                                        f"全 {len(_page_cd_links)} 件を出しています")
                            else:
                                st.caption(f"CD **{len(_page_cd_links)} 件**")
                            _pg_labels = [lnk["label"] for lnk in _pg_pool]
                            # 絞り込みで選択中のCDが消えると選びようがなくなるので、
                            # 残っていないときは選び直し（先頭）に戻す
                            _pg_cd_key = f"mf_pg_cd_{selected_no}"
                            if st.session_state.get(_pg_cd_key) not in _pg_labels:
                                st.session_state.pop(_pg_cd_key, None)
                            _pg_sel_label = st.selectbox(
                                "CD",
                                options=_pg_labels,
                                key=_pg_cd_key,
                            )
                            _pg_sel_lnk = _pg_pool[_pg_labels.index(_pg_sel_label)]
                            _pg_cand_sel = st.selectbox(
                                "適用する候補",
                                options=_man_cand_opts,
                                key=f"mf_pg_cand_{selected_no}",
                            )
                            _pg_target_idx = _no_cd_indices[_man_cand_opts.index(_pg_cand_sel)]
                            if st.button("💿 このCDで情報を取得", key=f"mf_pg_fetch_{selected_no}"):
                                _m_alb = _pg_sel_lnk["album_id"]
                                _m_trk = _pg_sel_lnk["track_id"]
                                _man_target_idx = _pg_target_idx
                                with st.spinner("CD情報を取得中..."):
                                    try:
                                        _m_cd_result = _get_mf_client().fetch_product_detail(_m_alb, _m_trk)
                                    except Exception as _pe:
                                        _m_fetch_err = str(_pe)

                            st.divider()
                            st.markdown("##### 🔗 URLから直接指定（上のリストにない場合）")
                        else:
                            st.caption(
                                "MINCの **検索URL**（/music/list?tr=…）、"
                                "**CD商品リストURL**（/product/list/?dn=品番…）、"
                                "**CD詳細URL**（/parts/product/detail?album_id=…）"
                                "のいずれかを貼り付けてください。"
                            )

                        # ── ② URL入力 ─────────────────────────────────────────
                        _man_url = st.text_input(
                            "MINC URL（検索URL・CD商品リストURL・CD詳細URL）",
                            key=f"mf_manual_url_{selected_no}",
                            placeholder="https://www.minc.or.jp/music/list?tr=... / /product/list/?dn=品番 / /parts/product/detail?album_id=...",
                        )
                        _man_sel = st.selectbox(
                            "適用する候補" if not _page_cd_links else "適用する候補（URL指定用）",
                            options=_man_cand_opts,
                            key=f"mf_manual_cand_{selected_no}",
                        )
                        _man_url_target_idx = _no_cd_indices[_man_cand_opts.index(_man_sel)] if _man_sel else None
                        if st.button(
                            "💿 URLからCD情報を取得",
                            key=f"mf_manual_fetch_{selected_no}",
                            disabled=not _man_url,
                        ):
                            try:
                                _m_parsed = urllib.parse.urlparse(_man_url.strip())
                                _m_q = dict(urllib.parse.parse_qsl(_m_parsed.query))
                                if "product/detail" in _m_parsed.path:
                                    _m_alb = _m_q.get("album_id", "")
                                    _m_trk = _m_q.get("track_id", "")
                                    if not _m_alb or not _m_trk:
                                        _m_fetch_err = "URLから album_id / track_id を取得できませんでした。"
                                elif "music/list" in _m_parsed.path:
                                    _m_title = _m_q.get("tr", "")
                                    _m_author = _m_q.get("ka", "")
                                    _m_match = int(_m_q.get("match", "2"))
                                    if not _m_title:
                                        _m_fetch_err = "URLから検索語（tr パラメータ）を取得できませんでした。"
                                    else:
                                        with st.spinner(f"MINC で「{_m_title}」を検索中..."):
                                            _m_res = _get_mf_client().search(_m_title, author=_m_author, match=_m_match)
                                        _m_hits = [it for it in _m_res.get("results", []) if it.get("_album_id")]
                                        if _m_hits:
                                            _m_alb = _m_hits[0]["_album_id"]
                                            _m_trk = _m_hits[0]["_track_id"]
                                        else:
                                            _m_fetch_err = "CD情報付き結果が見つかりませんでした。CD詳細URLを直接貼り付けてください。"
                                elif "product/list" in _m_parsed.path:
                                    # CD商品リスト（品番検索 /product/list/?dn=… や
                                    # 作品コードからの from_saku）。この表の行には
                                    # track_id が無いので、候補の曲名から拾い直す
                                    _m_dn = (_m_q.get("dn", "")
                                             or _m_q.get("jcd", "")).strip()
                                    _m_cand_title = (
                                        _mf_items[_man_url_target_idx].get("作品名", "")
                                        if _man_url_target_idx is not None else "")
                                    if not _m_dn:
                                        _m_fetch_err = (
                                            "URLから品番（dn パラメータ）を取得できませんでした。")
                                    else:
                                        with st.spinner(f"MINC で品番「{_m_dn}」を検索中..."):
                                            _m_cds = _get_mf_client().search_cds_by_hinban(_m_dn)
                                        _m_list = _m_cds.get("cds") or []
                                        if not _m_list:
                                            _m_fetch_err = (
                                                _m_cds.get("error")
                                                or f"品番 {_m_dn} のCDが見つかりませんでした。")
                                        else:
                                            _m_alb = _m_list[0].get("album_id", "")
                                            _m_trk = _m_list[0].get("track_id", "")
                                            if _m_alb and not _m_trk:
                                                with st.spinner("トラックを特定中..."):
                                                    _m_trk = _get_mf_client().find_track_id(
                                                        _m_alb, _m_cand_title) or "0"
                                            if not _m_alb:
                                                _m_fetch_err = (
                                                    f"品番 {_m_dn} のCDから album_id を"
                                                    "取得できませんでした。")
                                else:
                                    _m_fetch_err = (
                                        "MINCの検索URL（/music/list）、CD商品リストURL"
                                        "（/product/list）、CD詳細URL（/parts/product/detail）"
                                        "のいずれかを貼り付けてください。"
                                    )
                                if not _m_fetch_err and _m_alb and _m_trk:
                                    with st.spinner("CD情報を取得中..."):
                                        _m_cd_result = _get_mf_client().fetch_product_detail(_m_alb, _m_trk)
                                    _man_target_idx = _man_url_target_idx
                            except Exception as _me:
                                _m_fetch_err = str(_me)

                        if _m_fetch_err:
                            st.error(_m_fetch_err)

                        # ── 共通: 取得結果の処理 ────────────────────────────────
                        if _m_cd_result is not None:
                            if _m_cd_result.get("error"):
                                st.error(f"エラー: {_m_cd_result['error']}")
                            else:
                                if _man_target_idx is not None:
                                    st.session_state[f"mf_delg_{selected_no}_{_man_target_idx}"] = _m_cd_result
                                    _mf_res_ss_u = st.session_state.get(f"mf_results_{selected_no}", {})
                                    if "results" in _mf_res_ss_u and _man_target_idx < len(_mf_res_ss_u["results"]):
                                        _ss_item = _mf_res_ss_u["results"][_man_target_idx]
                                        for _fk, _fv in {
                                            "品番":           _m_cd_result.get("品番", ""),
                                            "CD商品タイトル": _m_cd_result.get("CD商品タイトル", ""),
                                            "アーティスト":   _m_cd_result.get("アーティスト", ""),
                                            "_album_id":      _m_alb,
                                            "_track_id":      _m_trk,
                                        }.items():
                                            if _fv and not _ss_item.get(_fk):
                                                _ss_item[_fk] = _fv
                                        for _wk in (
                                            f"mf_cat_{selected_no}_{_man_target_idx}",
                                            f"mf_cdtitle_{selected_no}_{_man_target_idx}",
                                            f"mf_art_{selected_no}_{_man_target_idx}",
                                        ):
                                            st.session_state.pop(_wk, None)
                                _cd_iv_raw  = _m_cd_result.get("IV", "")
                                _cd_iv_appl = {"I": "インスト", "V": "ヴォーカル"}.get(_cd_iv_raw, "")
                                _cd_title_d = _m_cd_result.get("CD商品タイトル", "")
                                _cd_cat_d   = _m_cd_result.get("品番", "")
                                _cd_art_d   = _m_cd_result.get("アーティスト", "")
                                _cd_trk_d   = _m_cd_result.get("トラック番号", "")
                                _cd_name_d  = _m_cd_result.get("曲名", "")
                                _cd_dur_d   = _m_cd_result.get("尺", "")
                                _cd_delg_d  = _m_cd_result.get("集中管理", "")
                                _cd_rec_co  = _m_cd_result.get("レコード会社名", "")
                                _m_direct = {
                                    "委任者":         _cd_delg_d,
                                    "CD番号":         _cd_cat_d,
                                    "CD名":           _cd_title_d,
                                    "アーティスト":   _cd_art_d,
                                    "レコード会社名": _cd_rec_co,
                                }
                                if not _cd_iv_appl and not (
                                    _is_blank(row.get("作曲者", "")) and _is_blank(row.get("作詞者", ""))
                                ):
                                    # CD詳細にI/V表記が無ければ、取得済みの作詞者の有無で判定
                                    _cd_iv_appl = _infer_iv(str(row.get("作詞者", "")))
                                if _cd_iv_appl and not str(row.get("I/V区分", "")).strip():
                                    _m_direct["I/V区分"] = _cd_iv_appl
                                for _col, _val in _m_direct.items():
                                    if _val and _col in st.session_state.songs_df.columns:
                                        st.session_state.songs_df.at[row_idx, _col] = _val
                                st.session_state.pop("songs_editor", None)
                                _msg_lines = [
                                    f"✅ MINC CD情報を楽曲まとめに反映しました  "
                                    f"IV={_cd_iv_appl or '(取得不可)'}  委任者={_cd_delg_d or '(取得不可)'}"
                                ]
                                if _cd_title_d or _cd_cat_d:
                                    _msg_lines.append(f"CD: {_cd_title_d}" + (f"（{_cd_cat_d}）" if _cd_cat_d else ""))
                                if _cd_rec_co:
                                    _msg_lines.append(f"レコード会社: {_cd_rec_co}")
                                if _cd_trk_d or _cd_name_d:
                                    _msg_lines.append(
                                        f"トラック{_cd_trk_d}: {_cd_name_d}" + (f"（{_cd_dur_d}）" if _cd_dur_d else "")
                                    )
                                st.session_state["_apply_msg"] = "  \n".join(_msg_lines)
                                st.session_state[f"mf_manual_dbg_{selected_no}"] = _m_cd_result.get("debug_html", "")[:3000]
                                st.rerun()

                # デバッグ HTML 表示（手動取得で品番・CD名が取れなかった場合、rerun後も表示）
                _mf_manual_dbg = st.session_state.get(f"mf_manual_dbg_{selected_no}", "")
                if _mf_manual_dbg:
                    with st.expander("🔍 デバッグ HTML（品番・CD名が取得できなかった場合）", expanded=True):
                        st.caption("このHTMLを確認すると、ページ構造から正確なパーサーを書けます。")
                        st.code(_mf_manual_dbg, language="html")
                        if st.button("デバッグHTML削除", key=f"mf_dbg_clear_{selected_no}"):
                            del st.session_state[f"mf_manual_dbg_{selected_no}"]
                            st.rerun()

        # ---- CD情報検索 ----
        st.divider()
        st.markdown('<a id="sec-cd-search"></a>', unsafe_allow_html=True)
        st.markdown("#### 💿 CD情報検索")
        st.caption(
            "JASRACコードまたは曲名でMINCを検索し、収録CDリストから品番・レコード会社名を申告フォーマットに反映します。"
        )
        if not _mf_ok:
            st.info("⚠️ MINCにログインするとCD情報検索が使えます。")
        else:
            _cds_jcd_default = str(row.get("JASRAC作品コード", "")).strip()
            _cds_jcd_default = "" if _cds_jcd_default.lower() == "nan" else _cds_jcd_default
            _cds_title_default = str(row.get("曲名", "")).strip()
            _cds_title_default = "" if _cds_title_default.lower() == "nan" else _cds_title_default

            _cds_artist_default = str(row.get("アーティスト", "")).strip()
            _cds_artist_default = "" if _cds_artist_default.lower() == "nan" else _cds_artist_default

            _cds_c1, _cds_c2, _cds_c3 = st.columns(3)
            with _cds_c1:
                _cds_jcd_input = st.text_input(
                    "JASRACコード",
                    value=_cds_jcd_default,
                    key=f"cds_jcd_{selected_no}",
                    placeholder="例: 123-4567-8",
                    help="JASRACコードがあれば優先して検索します",
                )
            with _cds_c2:
                _cds_title_input = st.text_input(
                    "曲名（JASRACコードがない場合）",
                    value=_cds_title_default,
                    key=f"cds_title_{selected_no}",
                    placeholder="曲名で検索",
                )
            with _cds_c3:
                _cds_artist_input = st.text_input(
                    "アーティスト（任意・絞り込み用）",
                    value=_cds_artist_default,
                    key=f"cds_artist_{selected_no}",
                    placeholder="例: EXILE",
                    help=(
                        "CD商品リストをこのアーティストで絞り込みます。"
                        "曲名検索のときは、同名曲から作品を選ぶ手がかりにも使います。"
                        "一致が0件のときは全件表示します（オムニバス盤は (V.A.) 表記）。"
                    ),
                )

            if st.button("🔍 CDリストを検索", key=f"cds_search_{selected_no}", type="primary", use_container_width=True):
                with st.spinner("MINCからCDリストを取得中..."):
                    try:
                        _cds_client = _get_mf_client()
                        if _cds_jcd_input.strip():
                            _cds_raw = _cds_client.search_cds_by_jasrac(
                                _cds_jcd_input.strip(),
                                title=_cds_title_input.strip() or _cds_title_default,
                            )
                        else:
                            _cds_term = _cds_title_input.strip() or _cds_title_default
                            if not _cds_term:
                                _cds_raw = {"cds": [], "error": "JASRACコードまたは曲名を入力してください。"}
                            else:
                                # 曲名のみ → まず作品を検索し、ヒットしたJASRACコードで
                                #            CD商品リスト（全件）を取得する
                                _cds_sr = _cds_client.search(_cds_term)
                                _cds_hits = _cds_sr.get("results") or [{}]
                                _cds_ai = _cds_artist_input.strip().lower()

                                # 候補を「絞る」のではなく「並べ替える」。
                                #
                                # 同じ曲が複数の作品コードで登録されていること
                                # があり、CD を持っているのは収録曲の行、配信
                                # だけの行は別のコードを持つ、という形になる。
                                # 収録曲の行はアーティスト欄が空のことが多い
                                # ので、アーティストで絞ると CD を持つ方が真っ
                                # 先に消えてしまう（「ゾートロープの光の小人」）。
                                #
                                # 欲しいのは CD なので、まず行の種類で並べ、
                                # 同じ種類の中でアーティストの一致を見る。
                                # 空欄は「不一致」ではなく「不明」として真ん中。
                                _CDS_SRC_ORDER = {"収録曲": 0, "作品": 1, "配信曲": 2}

                                def _cds_rank(_h: dict) -> tuple[int, int]:
                                    _src = _CDS_SRC_ORDER.get(
                                        str(_h.get("_source_table", "")), 1)
                                    _a = str(_h.get("アーティスト", "")).strip().lower()
                                    if not _cds_ai:
                                        _art = 1
                                    elif _cds_ai in _a:
                                        _art = 0
                                    elif not _a:
                                        _art = 1
                                    else:
                                        _art = 2
                                    return (_src, _art)

                                _cds_hits = sorted(_cds_hits, key=_cds_rank)

                                # 上から順に引き、CD が見つかった時点で止める。
                                # 1件目に CD が無くても、別のコードに付いて
                                # いることがあるため。MINC への負担を考えて
                                # 3件まで。同じコードは二度引かない
                                _cds_raw = None
                                _cds_tried: list[str] = []
                                for _cds_h in _cds_hits:
                                    _cds_fjcd = str(_cds_h.get("JASRAC作品コード", "")).strip()
                                    _cds_fncd = str(_cds_h.get("NexTone管理番号", "")).strip()
                                    if not _cds_fjcd and not _cds_fncd:
                                        continue
                                    _cds_ckey = f"{_cds_fjcd}|{_cds_fncd}"
                                    if _cds_ckey in _cds_tried:
                                        continue
                                    _cds_tried.append(_cds_ckey)
                                    _cds_try = _cds_client.search_cds_by_jasrac(
                                        _cds_fjcd,
                                        title=_cds_h.get("作品名", "") or _cds_term,
                                        ncd=_cds_fncd,
                                    )
                                    # 最初に引いた結果は、CD が無くても残す。
                                    # 配信しかない曲は、その配信を出したいため
                                    if _cds_raw is None:
                                        _cds_raw = _cds_try
                                    if _cds_try.get("cds"):
                                        _cds_raw = _cds_try
                                        break
                                    if len(_cds_tried) >= 3:
                                        break

                                if _cds_raw is None:
                                    _cds_raw = {
                                        "cds": [],
                                        "error": (
                                            _cds_sr.get("error")
                                            or f"「{_cds_term}」に一致する作品が見つかりませんでした。"
                                        ),
                                    }
                        # 検索時点のアーティスト指定を結果と一緒に保持する
                        _cds_raw["_artist_filter"] = _cds_artist_input.strip()
                        st.session_state[f"cds_results_{selected_no}"] = _cds_raw
                        for _cds_ck in list(st.session_state.keys()):
                            if _cds_ck.startswith(f"cds_detail_{selected_no}_"):
                                del st.session_state[_cds_ck]
                    except MusicForestError as _cds_ce:
                        st.session_state[f"cds_results_{selected_no}"] = {"cds": [], "error": str(_cds_ce)}

            _cds_res_cur = st.session_state.get(f"cds_results_{selected_no}")
            _render_cd_results(
                _cds_res_cur,
                row_idx,
                f"cds_{selected_no}",
                artist_filter=(_cds_res_cur or {}).get("_artist_filter", ""),
            )

        if selected_label:
            st.divider()

            # ---- 全自動パイプライン ----
            st.markdown('<a id="sec-pipeline"></a>', unsafe_allow_html=True)
            st.markdown("#### 🔄 全自動調査パイプライン（MusicBrainz → MINC / J-WID / NexTone）")
            st.caption("① 曲名抽出 → ② MusicBrainz で正式タイトル & ISRC 取得 → ③ MINC / J-WID / NexTone で著作権情報取得")

            _pip_mf_ok, _ = st.session_state.get("mf_auth_state", (False, ""))

            pip_col1, pip_col2, pip_col3 = st.columns([2, 1, 1])
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
            with pip_col3:
                _pip_mf_match_sel = st.selectbox(
                    "MINC 一致方式",
                    options=["2: 前方一致", "3: キーワード", "1: 完全一致"],
                    key=f"pip_mf_match_{selected_no}",
                    disabled=not _pip_mf_ok,
                    help="パイプライン実行時の MINC 検索方式",
                )
            _pip_mf_match_int = int(_pip_mf_match_sel.split(":")[0])

            pip_opt1, pip_opt2, pip_opt3 = st.columns(3)
            with pip_opt1:
                pip_use_minc = st.checkbox(
                    "MINC も使う",
                    value=_pip_mf_ok,
                    key=f"pip_use_minc_{selected_no}",
                    disabled=not _pip_mf_ok,
                    help="CD番号・レコード会社名・委任者をMINCから取得します" if _pip_mf_ok else "MINCにログインしてください",
                )
            with pip_opt2:
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
            with pip_opt3:
                _pip_sp_avail = spotify_available()
                pip_use_spotify = st.checkbox(
                    "Spotify API も使う",
                    value=False,
                    key=f"pip_use_spotify_{selected_no}",
                    disabled=not _pip_sp_avail,
                    help="アーティスト・アルバム・ISRCをSpotifyから取得します" if _pip_sp_avail else "SPOTIFY_CLIENT_ID / SECRET が未設定",
                )

            st.radio(
                "詳細（作家名・管理状況）の自動取得",
                options=["先頭候補のみ", "取得しない", "全候補"],
                index=0,
                horizontal=True,
                key=f"pip_auto_detail_{selected_no}",
                help=(
                    "J-WID / MINC の詳細は候補ごとに別リクエストのため、全候補を取ると時間がかかります。"
                    "取得しなくてもタブ表示・候補一覧・作品コードは出ます。"
                    "各候補の「🔄 詳細を再取得」または結果タブ内の「まとめて取得」でいつでも取得できます。"
                ),
            )

            wav_dur_raw = str(row.get("WAVフル尺", "")).strip()
            wav_dur_sec_val = _hms_to_sec(wav_dur_raw)
            if wav_dur_raw and wav_dur_raw.lower() != "nan":
                st.caption(f"WAV フル尺: {wav_dur_raw}（{wav_dur_sec_val:.1f} 秒）をMusicBrainz の尺絞り込みに使用します。")
            else:
                st.caption("WAV フル尺が未取得のため MusicBrainz は尺絞り込みなしで検索します。")

            # 検索語候補セレクト
            if len(term_candidates) > 1:
                _pip_term_opts = [f"[{label}] {val}" for label, val in term_candidates]
                _pip_term_sel = st.radio(
                    "検索語候補",
                    options=_pip_term_opts,
                    key=f"pip_term_{selected_no}",
                    horizontal=True,
                )
                _pip_song_title = term_candidates[_pip_term_opts.index(_pip_term_sel)][1]
            else:
                _pip_song_title = term_candidates[0][1] if term_candidates else ""

            # 選択語を編集できるinput（候補が変わったときリセット）
            _pip_edit_key  = f"pip_term_edit_{selected_no}"
            _pip_prev_key  = f"pip_term_prev_{selected_no}"
            if st.session_state.get(_pip_prev_key) != _pip_song_title:
                if _pip_edit_key in st.session_state:
                    del st.session_state[_pip_edit_key]
                st.session_state[_pip_prev_key] = _pip_song_title
            _pip_search_term = st.text_input(
                "検索語（編集可）",
                key=_pip_edit_key,
                value=_pip_song_title,
                help="不要な語を削除するなど自由に編集できます。候補を変えると自動リセットされます。",
            )

            st.caption("🔍 J-WID 事前絞り込み（検索時に J-WID サーバーへ渡します。空欄 = 全件取得）")
            _pf1, _pf2, _pf3, _pf4 = st.columns(4)
            _pf1.text_input("作品名の一部",        placeholder="曲名キーワード",  key=f"jf_title_{selected_no}")
            _pf2.text_input("アーティスト名",       placeholder="例: EXILE",       key=f"jf_artist_{selected_no}",
                            help="J-WID の IN_ARTIST_NAME1 に渡してサーバー側で絞り込みます")
            _pf3.text_input("著作者名",             placeholder="作曲者・作詞者",  key=f"jf_author_{selected_no}")
            _pf4.text_input("JASRAC コード",        placeholder="作品コード",      key=f"jf_code_{selected_no}")

            if st.button(
                "🔄 全自動調査パイプラインを実行",
                key=f"pipeline_btn_{selected_no}",
                type="primary",
                use_container_width=True,
            ):
                with st.spinner("① 検索語決定 → ② Claude/MusicBrainz/Spotify → ③ MINC / J-WID / NexTone 検索 中..."):
                    pip_result = run_pipeline(
                        event_name=str(row.get("イベント名", "")),
                        wav_full_duration=wav_dur_raw,
                        wav_detected_title=str(row.get("WAV検出タイトル", "")),
                        song_title=_pip_search_term or _pip_song_title,
                        jwid_artist=st.session_state.get(f"jf_artist_{selected_no}", ""),
                        tolerance_sec=float(pip_tolerance),
                        mb_score_threshold=int(pip_threshold),
                        use_claude=bool(pip_use_claude),
                        use_spotify=bool(pip_use_spotify),
                    )
                st.session_state[f"pipeline_result_{selected_no}"] = pip_result
                # 再実行のたびに自動フェッチフラグをリセット（新規ページ結果を取得し直すため）
                st.session_state.pop(f"pip_j_auto_fetched_{selected_no}", None)
                st.session_state.pop(f"pip_mf_auto_fetched_{selected_no}", None)

                # MINC 検索（セッション有効かつチェックONのみ）
                if pip_use_minc and _pip_mf_ok:
                    _pip_minc_term = pip_result.get("search_title", _pip_search_term or _pip_song_title)
                    _pip_mb_artist = (pip_result.get("mb_best") or {}).get("artist", "")
                    with st.spinner(f"MINC で「{_pip_minc_term[:30]}」を検索中..."):
                        try:
                            _pip_mf_c = _get_mf_client()
                            _pip_mf_r = _pip_mf_c.search(_pip_minc_term, match=_pip_mf_match_int)
                            # 全結果が "作品" テーブル（_album_id なし）のとき、
                            # mb_best アーティスト名を加えて前方一致で再検索し CD 情報を補完する
                            if (
                                _pip_mb_artist
                                and _pip_mf_r.get("results")
                                and all(not it.get("_album_id") for it in _pip_mf_r["results"])
                            ):
                                _pip_mf_r2 = _pip_mf_c.search(_pip_minc_term, author=_pip_mb_artist, match=2)
                                if any(it.get("_album_id") for it in _pip_mf_r2.get("results", [])):
                                    _pip_mf_r2["_cd_fallback_artist"] = _pip_mb_artist
                                    _pip_mf_r = _pip_mf_r2
                            st.session_state[f"pipeline_minc_{selected_no}"] = _pip_mf_r
                        except MusicForestError as _me:
                            st.session_state[f"pipeline_minc_{selected_no}"] = {"error": str(_me), "results": []}
                else:
                    st.session_state.pop(f"pipeline_minc_{selected_no}", None)

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
                                    "CD名":         _cl_res.get("cd_name",""),
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
                                    "CD名":         _spb.get("album", ""),
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

                # タブラベル: J-WID は先頭結果のアーティスト/著作者、MINC は先頭結果の作品名
                _j_first = (jwid_r.get("results") or [{}])[0]
                _j_lbl_hint = _j_first.get("アーティスト", "") or _j_first.get("著作者名", "")
                _j_pages = jwid_r.get("pages_fetched", 1)
                _j_cnt = len(jwid_r.get("results") or [])
                _j_cnt_str = f"{_j_cnt}件" + (f"/{_j_pages}p" if _j_pages > 1 else "")
                _j_tab_lbl = "📋 J-WID" + (f" {_j_lbl_hint[:15]}" if _j_lbl_hint else "") + (f" [{_j_cnt_str}]" if _j_cnt else "")
                _mf_r_pip = st.session_state.get(f"pipeline_minc_{selected_no}") or {}
                _mf_first = (_mf_r_pip.get("results") or [{}])[0]
                _mf_lbl_hint = _mf_first.get("作品名", "") or _mf_first.get("アーティスト", "")
                _mf_tab_lbl = "🌲 MINC" + (f" {_mf_lbl_hint[:18]}" if _mf_lbl_hint else "")

                pip_tab_j, pip_tab_n, pip_tab_mf = st.tabs([_j_tab_lbl, "📋 NexTone 結果", _mf_tab_lbl])

                def _nkfc(s: str) -> str:
                    return unicodedata.normalize("NFKC", s).lower()

                with pip_tab_j:
                    st.caption(f"検索URL: {jwid_r.get('search_url','')}")
                    if jwid_r.get("error"):
                        st.error(f"エラー: {jwid_r['error']}")
                    elif not jwid_r.get("results"):
                        st.warning("J-WID: 該当なし")
                        with st.expander("デバッグ HTML"):
                            st.code(jwid_r.get("debug_html", "")[:3000], language="html")
                    else:
                        _j_pages_disp = jwid_r.get("pages_fetched", 1)
                        _j_total_disp = len(jwid_r["results"])
                        _j_page_str = f"（{_j_pages_disp} ページ分）" if _j_pages_disp > 1 else ""
                        st.success(f"{_j_total_disp} 件{_j_page_str}")
                        _mb_artist_hint = mb_best.get("artist", "") if mb_best else ""
                        if _mb_artist_hint:
                            st.caption(f"💡 MusicBrainz アーティスト参照: **{_mb_artist_hint}**")

                        # フィルター値はパイプラインボタン上部の入力から読む
                        _jf_title  = st.session_state.get(f"jf_title_{selected_no}",  "")
                        _jf_artist = st.session_state.get(f"jf_artist_{selected_no}", "")
                        _jf_author = st.session_state.get(f"jf_author_{selected_no}", "")
                        _jf_code   = st.session_state.get(f"jf_code_{selected_no}",   "")

                        # フィルター済みインデックスを先に計算して件数表示
                        _jf_indices = [
                            i for i, it in enumerate(jwid_r["results"])
                            if (not _jf_title  or _nkfc(_jf_title)  in _nkfc(it.get("作品名",      "")))
                            and (not _jf_artist or _nkfc(_jf_artist) in _nkfc(it.get("アーティスト","")))
                            and (not _jf_author or _nkfc(_jf_author) in _nkfc(it.get("著作者名",    "")))
                            and (not _jf_code   or _nkfc(_jf_code)   in _nkfc(it.get("作品コード",  "")))
                        ]
                        _jf_total = len(jwid_r["results"])
                        if any([_jf_title, _jf_artist, _jf_author, _jf_code]):
                            st.caption(f"絞り込み結果: {len(_jf_indices)} / {_jf_total} 件")
                            if not _jf_indices:
                                _art_vals = sorted({it.get("アーティスト","") for it in jwid_r["results"] if it.get("アーティスト","")})
                                st.warning("フィルター条件に一致する候補がありません。")
                                if _art_vals:
                                    st.caption("このJ-WID結果に含まれるアーティスト: " + "　/　".join(_art_vals))

                        # 詳細（作家名・管理状況）は候補ごとに別リクエスト。
                        # 既定は先頭候補のみ自動取得し、残りは任意のタイミングで取得する。
                        _pip_dmode = st.session_state.get(f"pip_auto_detail_{selected_no}", "先頭候補のみ")
                        _jwid_auto_key = f"pip_j_auto_fetched_{selected_no}"
                        _j_todo = [
                            i for i in _jf_indices
                            if not st.session_state.get(f"jwid_detail_{selected_no}_{i}")
                            and jwid_r["results"][i].get("_detail_url", "")
                        ]
                        _j_want: list[int] = []
                        if not st.session_state.get(_jwid_auto_key):
                            if _pip_dmode == "全候補":
                                _j_want = _j_todo
                            elif _pip_dmode == "先頭候補のみ":
                                _j_want = _j_todo[:1]
                            st.session_state[_jwid_auto_key] = True
                        if not _j_want and _j_todo and st.button(
                            f"⬇️ 表示中の候補の詳細をまとめて取得（{len(_j_todo)}件）",
                            key=f"pip_j_fetchall_{selected_no}",
                            use_container_width=True,
                            help="作曲者・作詞者・管理状況を取得します（1件あたり数秒かかります）",
                        ):
                            _j_want = _j_todo
                        if _j_want:
                            from modules.scraper import fetch_jwid_detail as _fetch_detail_auto
                            with st.spinner(f"J-WID 詳細を取得中（{len(_j_want)} 件）…"):
                                for _ji in _j_want:
                                    st.session_state[f"jwid_detail_{selected_no}_{_ji}"] = \
                                        _fetch_detail_auto(jwid_r["results"][_ji].get("_detail_url", ""))
                            st.rerun()

                        for _disp_idx, i in enumerate(_jf_indices):
                            item = jwid_r["results"][i]
                            _j_art = item.get("アーティスト","")
                            with st.expander(
                                f"候補{_disp_idx + 1}: {item.get('作品名','')}"
                                + (f"  アーティスト:{_j_art}" if _j_art else "")
                                + f"  著作者:{item.get('著作者名','')}  コード:{item.get('作品コード','(なし)')}",
                                expanded=(_disp_idx == 0),
                            ):
                                # 詳細ページから作曲者/作詞者を取得済みかチェック
                                _detail_key = f"jwid_detail_{selected_no}_{i}"
                                _detail = st.session_state.get(_detail_key, {})

                                # ウィジェット描画前に session_state を同期（描画後は変更不可）
                                if _detail:
                                    st.session_state[f"pip_j_comp_{selected_no}_{i}"] = _detail.get("作曲者", "")
                                    st.session_state[f"pip_j_lyric_{selected_no}_{i}"] = _detail.get("作詞者", "")
                                    st.session_state[f"pip_j_arr_{selected_no}_{i}"] = _detail.get("編曲者", "")
                                    st.session_state[f"pip_j_tran_{selected_no}_{i}"] = _detail.get("訳詞者", "")

                                pc1, pc2 = st.columns(2)
                                pc1.text_input("作品コード",      value=item.get("作品コード",""),   key=f"pip_j_code_{selected_no}_{i}",   disabled=True)
                                pc1.text_input("作品名",          value=item.get("作品名",""),       key=f"pip_j_title_{selected_no}_{i}",  disabled=True)
                                pc1.text_input("アーティスト名",  value=item.get("アーティスト",""), key=f"pip_j_artist_{selected_no}_{i}", disabled=True)
                                pc1.text_input("著作者名（一覧）", value=item.get("著作者名",""),    key=f"pip_j_auth_{selected_no}_{i}",   disabled=True)
                                _comp_disp = _detail.get("作曲者","（詳細取得で確認）") if not _detail else _detail.get("作曲者","")
                                _lyric_disp = _detail.get("作詞者","（詳細取得で確認）") if not _detail else _detail.get("作詞者","")
                                pc2.text_input("作曲者", value=_comp_disp, key=f"pip_j_comp_{selected_no}_{i}", disabled=True)
                                pc2.text_input("作詞者", value=_lyric_disp, key=f"pip_j_lyric_{selected_no}_{i}", disabled=True)
                                pc2.text_input("編曲者", value=_detail.get("編曲者",""),    key=f"pip_j_arr_{selected_no}_{i}", disabled=True)
                                pc2.text_input("訳詞者", value=_detail.get("訳詞者",""),    key=f"pip_j_tran_{selected_no}_{i}", disabled=True)

                                # 管理状況を表示（詳細取得済みの場合）
                                _mgmt_pip = _detail.get("管理状況", {})
                                if _mgmt_pip:
                                    st.markdown("**管理状況（JASRAC）:**  \n" + _format_management_status(_mgmt_pip))
                                if _detail and not _detail.get("作曲者"):
                                    with st.expander("⚠️ デバッグ情報（作曲者が取得できなかった場合）"):
                                        st.write(f"**error**: `{_detail.get('error')}`")
                                        st.write(f"**著作者リスト**: {_detail.get('著作者リスト')}")
                                        if _detail.get("debug_html"):
                                            st.code(_detail["debug_html"][:5000], language="html")
                                        else:
                                            st.warning("debug_html が空です（セッションエラーまたはdetail_urlが空の可能性）")

                                btn_col1, btn_col2 = st.columns(2)
                                with btn_col1:
                                    if st.button("🔄 J-WID詳細を再取得", key=f"pip_detail_j_{selected_no}_{i}", use_container_width=True):
                                        with st.spinner("詳細ページ取得中..."):
                                            from modules.scraper import fetch_jwid_detail as _fetch_detail
                                            _d = _fetch_detail(item.get("_detail_url", ""))
                                        st.session_state[_detail_key] = _d
                                        if _d.get("作曲者"):
                                            # MINC個別検索の著作者欄はこの時点で生成済みのため
                                            # 直接代入せず、次回の描画時に反映させる
                                            st.session_state[f"mf_author_pending_{selected_no}"] = _d["作曲者"].strip()
                                        if _d.get("error"):
                                            st.error(f"詳細取得エラー: {_d['error']}")
                                        elif not _d.get("作曲者") and not _d.get("作詞者"):
                                            st.warning("作曲者・作詞者が取得できませんでした。下の「デバッグ HTML」を確認してください。")
                                        else:
                                            st.success(f"作曲者: {_d['作曲者']} ／ 作詞者: {_d['作詞者']}")
                                        st.rerun()
                                with btn_col2:
                                    if st.button("✅ 申告フォーマットに反映", key=f"pip_apply_j_{selected_no}_{i}", use_container_width=True):
                                        _pip_j_jcd = item.get("作品コード","")
                                        # JASRACコードが変わる場合は先に関連フィールドをクリア
                                        _apply_clear_on_jcd_change(row_idx, _pip_j_jcd)
                                        _pip_j_apply = {
                                            "曲名":           item.get("作品名",""),
                                            "作曲者":         _detail.get("作曲者") or item.get("著作者名",""),
                                            "作詞者":         _detail.get("作詞者",""),
                                            "編曲者":         _detail.get("編曲者",""),
                                            "訳詞者":         _detail.get("訳詞者",""),
                                            "JASRAC作品コード": _pip_j_jcd,
                                            "アーティスト":   item.get("アーティスト","") or (mb_best.get("artist","") if mb_best else ""),
                                            "確認ステータス": "確定",
                                        }
                                        _apply_management_status(
                                            _detail.get("管理状況") or {}, _pip_j_apply,
                                            "J", st.session_state.songs_df.loc[row_idx])
                                        _hy = _infer_houyo(_pip_j_jcd)
                                        _cur_hy = str(st.session_state.songs_df.at[row_idx, "邦洋区分"] if "邦洋区分" in st.session_state.songs_df.columns else "").strip()
                                        if _hy and not _cur_hy:
                                            _pip_j_apply["邦洋区分"] = _hy
                                        _apply_iv_from_credits(_pip_j_apply)
                                        for col, val in _pip_j_apply.items():
                                            if val and col in st.session_state.songs_df.columns:
                                                st.session_state.songs_df.at[row_idx, col] = val
                                        st.session_state["_apply_msg"] = "楽曲まとめ・申告フォーマットに反映しました。"
                                        st.session_state.pop("songs_editor", None)
                                        st.rerun()

                                # CD情報検索パネル
                                _pipj_jcd = item.get("作品コード", "")
                                if _pipj_jcd and _mf_ok:
                                    st.divider()
                                    _show_cd_panel(
                                        _pipj_jcd, row_idx, f"pipj_{selected_no}_{i}",
                                        title=item.get("作品名", ""),
                                    )

                with pip_tab_n:
                    st.caption(f"検索URL: {ntone_r.get('search_url','')}")
                    if ntone_r.get("error"):
                        st.error(f"エラー: {ntone_r['error']}")
                    elif not ntone_r.get("results"):
                        st.warning("NexTone: 該当なし")
                    else:
                        st.success(f"{len(ntone_r['results'])} 件")
                        _nfc1, _nfc2 = st.columns(2)
                        _nf_title  = _nfc1.text_input("曲名で絞り込み",          placeholder="作品名の一部",    key=f"nf_title_{selected_no}")
                        _nf_artist = _nfc2.text_input("アーティスト名で絞り込み", placeholder="アーティスト名",  key=f"nf_artist_{selected_no}")
                        _nf_disp = 0
                        for i, item in enumerate(ntone_r["results"]):
                            if _nf_title  and _nkfc(_nf_title)  not in _nkfc(item.get("作品名",       "")): continue
                            if _nf_artist and _nkfc(_nf_artist) not in _nkfc(item.get("アーティスト", "")): continue
                            _nf_disp += 1
                            with st.expander(
                                f"候補{_nf_disp}: {item.get('作品名','')} ／ {item.get('管理番号','')}",
                                expanded=(_nf_disp == 1),
                            ):
                                nc1, nc2 = st.columns(2)
                                nc1.text_input("管理番号",    value=item.get("管理番号",""),    key=f"pip_n_id_{selected_no}_{i}",    disabled=True)
                                nc1.text_input("作品名",      value=item.get("作品名",""),      key=f"pip_n_title_{selected_no}_{i}", disabled=True)
                                nc1.text_input("作曲者",      value=item.get("作曲者",""),      key=f"pip_n_comp_{selected_no}_{i}",  disabled=True)
                                nc2.text_input("作詞者",      value=item.get("作詞者",""),      key=f"pip_n_lyric_{selected_no}_{i}", disabled=True)
                                nc2.text_input("アーティスト", value=item.get("アーティスト",""), key=f"pip_n_art_{selected_no}_{i}",   disabled=True)
                                if st.button("✅ 申告フォーマットに反映", key=f"pip_apply_n_{selected_no}_{i}", use_container_width=True):
                                    _pip_n_apply = {
                                        "作曲者": item.get("作曲者",""),
                                        "作詞者": item.get("作詞者",""),
                                        "NexTone管理番号": item.get("管理番号",""),
                                        "アーティスト": item.get("アーティスト","") or (mb_best.get("artist","") if mb_best else ""),
                                        "確認ステータス": "候補あり",
                                    }
                                    # 放送・配信は検索結果に出ているので通信は増えない
                                    _apply_management_status(
                                        item.get("管理状況") or {}, _pip_n_apply,
                                        "N", st.session_state.songs_df.loc[row_idx])
                                    _apply_iv_from_credits(_pip_n_apply)
                                    for col, val in _pip_n_apply.items():
                                        if val and col in st.session_state.songs_df.columns:
                                            st.session_state.songs_df.at[row_idx, col] = val
                                    st.session_state["_apply_msg"] = "楽曲まとめ・申告フォーマットに反映しました。"
                                    st.session_state.pop("songs_editor", None)
                                    st.rerun()

                with pip_tab_mf:
                    _pip_mf_res = st.session_state.get(f"pipeline_minc_{selected_no}")
                    if _pip_mf_res is None:
                        if _pip_mf_ok:
                            st.info("パイプライン実行時に「MINC も使う」をチェックすると自動で検索されます。")
                        else:
                            st.warning("MINCにログインしてから「MINC も使う」チェックを入れて実行してください。")
                    elif _pip_mf_res.get("error"):
                        st.error(f"MINC エラー: {_pip_mf_res['error']}")
                        with st.expander("デバッグ HTML"):
                            st.code(_pip_mf_res.get("debug_html","")[:3000], language="html")
                    elif not _pip_mf_res.get("results"):
                        st.warning("MINC: 該当なし")
                        with st.expander("デバッグ HTML"):
                            st.code(_pip_mf_res.get("debug_html","")[:3000], language="html")
                    else:
                        _pip_mf_items = _pip_mf_res["results"]
                        st.success(f"🌲 MINC: {len(_pip_mf_items)} 件")
                        st.caption(f"検索URL: {_pip_mf_res.get('search_url','')}")
                        if _pip_mf_res.get("_cd_fallback_artist"):
                            st.info(f"💡 タイトルのみでは全結果が「作品」テーブル（CD情報なし）だったため、アーティスト「{_pip_mf_res['_cd_fallback_artist']}」を追加して再検索しました。")

                        # 詳細は候補ごとに別リクエスト（キー未存在＝未取得、エラー時もキーをセットして無限ループ防止）
                        # 既定は先頭候補のみ自動取得し、残りは任意のタイミングで取得する。
                        _pip_dmode_mf = st.session_state.get(f"pip_auto_detail_{selected_no}", "先頭候補のみ")
                        _mf_auto_key = f"pip_mf_auto_fetched_{selected_no}"
                        _mf_todo = [
                            (_auto_pmi, _auto_it) for _auto_pmi, _auto_it in enumerate(_pip_mf_items[:10])
                            if f"pip_mf_ddetail_{selected_no}_{_auto_pmi}" not in st.session_state
                            and _auto_it.get("_detail_href", "")
                        ]
                        _mf_need_fetch: list = []
                        if not st.session_state.get(_mf_auto_key):
                            if _pip_dmode_mf == "全候補":
                                _mf_need_fetch = _mf_todo
                            elif _pip_dmode_mf == "先頭候補のみ":
                                _mf_need_fetch = _mf_todo[:1]
                            st.session_state[_mf_auto_key] = True
                        if not _mf_need_fetch and _mf_todo and st.button(
                            f"⬇️ 候補の詳細をまとめて取得（{len(_mf_todo)}件）",
                            key=f"pip_mf_fetchall_{selected_no}",
                            use_container_width=True,
                            help="作曲者・作詞者・編曲者を取得します（1件あたり数秒かかります）",
                        ):
                            _mf_need_fetch = _mf_todo
                        if _mf_need_fetch:
                            _mf_auto_c = _get_mf_client()
                            with st.spinner(f"MINC 詳細を取得中（{len(_mf_need_fetch)} 件）…"):
                                for _auto_pmi, _auto_it in _mf_need_fetch:
                                    _auto_dkey = f"pip_mf_ddetail_{selected_no}_{_auto_pmi}"
                                    try:
                                        st.session_state[_auto_dkey] = _mf_auto_c.get_detail(_auto_it["_detail_href"])
                                    except Exception as _auto_e:
                                        st.session_state[_auto_dkey] = {"error": str(_auto_e), "debug_html": f"例外: {type(_auto_e).__name__}: {_auto_e}"}
                            st.rerun()

                        for _pmi, _pm_item in enumerate(_pip_mf_items[:10]):
                            _pm_label = (
                                f"候補{_pmi+1}: {_pm_item.get('作品名','')} ／ {_pm_item.get('アーティスト','')} "
                                f"  CD:{_pm_item.get('CD商品タイトル','') or '(CD情報なし)'}  "
                                f"JASRAC:{_pm_item.get('JASRAC作品コード','(なし)')}  "
                                f"品番:{_pm_item.get('品番','(なし)')}"
                            )
                            _pm_detail_key = f"pip_mf_ddetail_{selected_no}_{_pmi}"
                            _pm_detail_data = st.session_state.get(_pm_detail_key, {})

                            with st.expander(_pm_label, expanded=(_pmi == 0)):
                                _pm_c1, _pm_c2 = st.columns(2)
                                _pm_c1.text_input("作品名",        value=_pm_item.get("作品名",""),         key=f"pip_mf_name_{selected_no}_{_pmi}", disabled=True)
                                _pm_c1.text_input("アーティスト",  value=_pm_item.get("アーティスト",""),   key=f"pip_mf_art_{selected_no}_{_pmi}",  disabled=True)
                                _pm_c1.text_input("品番（CD番号）", value=_pm_item.get("品番",""),           key=f"pip_mf_cat_{selected_no}_{_pmi}",  disabled=True)
                                _pm_c1.text_input("CD商品タイトル", value=_pm_item.get("CD商品タイトル",""), key=f"pip_mf_cdtitle_{selected_no}_{_pmi}", disabled=True)
                                _pm_c2.text_input("JASRAC作品コード", value=_pm_item.get("JASRAC作品コード",""), key=f"pip_mf_jcd_{selected_no}_{_pmi}", disabled=True)
                                _pm_c2.text_input("NexTone管理番号",  value=_pm_item.get("NexTone管理番号",""), key=f"pip_mf_ncd_{selected_no}_{_pmi}", disabled=True)
                                _pm_c2.text_input("レコード会社名",   value=_pm_item.get("レコード会社名",""),  key=f"pip_mf_label_{selected_no}_{_pmi}", disabled=True)

                                # 詳細フィールド（作曲者/作詞者/編曲者/訳詞者）
                                # session_state をウィジェット描画前に同期
                                if _pm_detail_data:
                                    st.session_state[f"pip_mf_comp_{selected_no}_{_pmi}"]  = _pm_detail_data.get("作曲者","")
                                    st.session_state[f"pip_mf_lyric_{selected_no}_{_pmi}"] = _pm_detail_data.get("作詞者","")
                                    st.session_state[f"pip_mf_arr2_{selected_no}_{_pmi}"]  = _pm_detail_data.get("編曲者","")
                                    st.session_state[f"pip_mf_tran_{selected_no}_{_pmi}"]  = _pm_detail_data.get("訳詞者","")
                                _pm_d4 = st.columns(4)
                                _pm_d4[0].text_input("作曲者",
                                    value=_pm_detail_data.get("作曲者","") if _pm_detail_data else "（詳細取得で確認）",
                                    key=f"pip_mf_comp_{selected_no}_{_pmi}", disabled=True)
                                _pm_d4[1].text_input("作詞者",
                                    value=_pm_detail_data.get("作詞者","") if _pm_detail_data else "（詳細取得で確認）",
                                    key=f"pip_mf_lyric_{selected_no}_{_pmi}", disabled=True)
                                _pm_d4[2].text_input("編曲者",
                                    value=_pm_detail_data.get("編曲者","") if _pm_detail_data else "",
                                    key=f"pip_mf_arr2_{selected_no}_{_pmi}", disabled=True)
                                _pm_d4[3].text_input("訳詞者",
                                    value=_pm_detail_data.get("訳詞者","") if _pm_detail_data else "",
                                    key=f"pip_mf_tran_{selected_no}_{_pmi}", disabled=True)

                                _pm_btn1, _pm_btn2 = st.columns(2)
                                with _pm_btn1:
                                    if st.button("🔄 MINC詳細を再取得", key=f"pip_mf_dget_{selected_no}_{_pmi}", use_container_width=True):
                                        with st.spinner("MINC詳細取得中..."):
                                            try:
                                                _pm_dd = _get_mf_client().get_detail(_pm_item.get("_detail_href",""))
                                                st.session_state[_pm_detail_key] = _pm_dd
                                            except Exception as _pm_e:
                                                st.session_state[_pm_detail_key] = {"error": str(_pm_e), "debug_html": f"例外: {type(_pm_e).__name__}: {_pm_e}"}
                                        st.rerun()
                                # 詳細取得エラーまたはデバッグ情報
                                # エラー時のみデバッグ情報を表示
                                if _pm_detail_data and _pm_detail_data.get("error"):
                                    st.error(f"MINC詳細エラー: {_pm_detail_data['error']}")
                                    _dbg = _pm_detail_data.get("debug_html", "")
                                    if _dbg:
                                        with st.expander("🔍 デバッグ HTML"):
                                            st.code(_dbg[:3000], language="html")

                                with _pm_btn2:
                                    if st.button("✅ MINC情報を申告フォーマットに反映", key=f"pip_mf_apply_{selected_no}_{_pmi}", use_container_width=True):
                                        _pm_jcd = _pm_item.get("JASRAC作品コード","")
                                        # JASRACコードが変わる場合は先に関連フィールドをクリア
                                        _apply_clear_on_jcd_change(row_idx, _pm_jcd)
                                        _pm_apply = {
                                            "曲名":            _pm_item.get("作品名",""),
                                            "アーティスト":    _pm_item.get("アーティスト",""),
                                            "CD番号":          _pm_item.get("品番",""),
                                            "CD名":            _pm_item.get("CD商品タイトル",""),
                                            "レコード会社名":  _pm_item.get("レコード会社名",""),
                                            "JASRAC作品コード": _pm_jcd,
                                            "NexTone管理番号": _pm_item.get("NexTone管理番号",""),
                                            "確認ステータス":  "候補あり",
                                        }
                                        _hy = _infer_houyo(_pm_jcd)
                                        _cur_hy_pm = str(st.session_state.songs_df.at[row_idx, "邦洋区分"] if "邦洋区分" in st.session_state.songs_df.columns else "").strip()
                                        if _hy and not _cur_hy_pm:
                                            _pm_apply["邦洋区分"] = _hy
                                        # 検索結果の「作詞／作曲」列 → 詳細取得済みならそちらで上書き
                                        # （どちらも無ければ空。ここでは自動フェッチはしない）
                                        for _ak in ["作曲者","作詞者","編曲者","訳詞者"]:
                                            if _pm_item.get(_ak): _pm_apply[_ak] = _pm_item[_ak]
                                        _cached = st.session_state.get(_pm_detail_key, {})
                                        if _cached and not _cached.get("error"):
                                            for _ak in ["作曲者","作詞者","編曲者","訳詞者"]:
                                                if _cached.get(_ak): _pm_apply[_ak] = _cached[_ak]
                                        _apply_iv_from_credits(_pm_apply)
                                        for _col, _val in _pm_apply.items():
                                            if _val and _col in st.session_state.songs_df.columns:
                                                st.session_state.songs_df.at[row_idx, _col] = _val
                                        st.session_state["_apply_msg"] = "楽曲まとめ・申告フォーマットに反映しました。"
                                        st.session_state.pop("songs_editor", None)
                                        st.rerun()

                                # CD情報検索パネル
                                _pipmf_jcd = _pm_item.get("JASRAC作品コード", "")
                                if _pipmf_jcd:
                                    st.divider()
                                    _show_cd_panel(
                                        _pipmf_jcd, row_idx, f"pipmf_{selected_no}_{_pmi}",
                                        title=_pm_item.get("作品名", ""),
                                    )


            # ---- 検索語と手動リンク ----
            col_terms, col_links = st.columns([3, 2])
            with col_terms:
                st.markdown("**検索語候補**（クリックして選択＆コピー）")
                for label, term in term_candidates:
                    st.text_input(label, value=term, key=f"term_{selected_no}_{label}")

            with col_links:
                st.markdown("**手動検索リンク**")
                st.caption("検索語（右のアイコンでコピー）")
                st.code(main_term, language=None)
                # J-WID: POST 送信のため URL に検索語を含められない → 承認後の検索フォームへ
                st.link_button(
                    "🔍 J-WID で検索",
                    "https://www2.jasrac.or.jp/eJwid/main?trxID=F00100",
                    use_container_width=True,
                )
                st.caption("↑ コピーした検索語を「作品タイトル」欄に貼り付けてください")
                # NexTone: 利用規約同意が必要なため直接検索URLへの誘導は不可 → トップページを開く
                st.link_button(
                    "🔍 NexTone で検索",
                    f"https://search.nex-tone.co.jp/",
                    use_container_width=True,
                )
                st.caption("↑ 利用規約に同意後、コピーした検索語で検索してください")
                # Google: 曲名 + 著作権者名（作曲者またはアーティスト）
                _rights_holder = str(row.get("作曲者", "")).strip()
                if not _rights_holder or _rights_holder.lower() == "nan":
                    _rights_holder = str(row.get("アーティスト", "")).strip()
                if _rights_holder and _rights_holder.lower() != "nan":
                    _google_q = urllib.parse.quote(f"{main_term} {_rights_holder}")
                else:
                    _google_q = encoded
                st.link_button(
                    "🔍 Google で検索",
                    f"https://www.google.com/search?q={_google_q}",
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
                st.caption("MINC（要ログイン）")
                _mf_link_term = st.session_state.get(f"mf_term_edit_{selected_no}") or main_term
                _mf_enc = urllib.parse.quote(_mf_link_term)
                st.link_button(
                    "🌲 MINC タイトル検索",
                    f"https://www.minc.or.jp/music/list/?tr={_mf_enc}&ka=&type=search-form-title&match=2",
                    use_container_width=True,
                )
                _mf_composer = str(row.get("作曲者", "")).strip()
                if not _mf_composer or _mf_composer.lower() == "nan":
                    _mf_composer = str(row.get("アーティスト", "")).strip()
                if _mf_composer and _mf_composer.lower() != "nan":
                    _mf_comp_enc = urllib.parse.quote(_mf_composer)
                    st.link_button(
                        f"🌲 MINC タイトル+著作者検索",
                        f"https://www.minc.or.jp/music/list/?tr={_mf_enc}&ka={_mf_comp_enc}&type=search-form-title&match=2",
                        use_container_width=True,
                    )

            # ---- J-WID 手動検索 → 反映 ----
            _jwid_manual_key = f"jwid_manual_{selected_no}"
            with st.expander("🔍 J-WID 手動検索コード入力 → 反映", expanded=False):
                st.caption(
                    f"上の「🔍 J-WID 作品検索」リンクで **{main_term[:30]}** を検索し、"
                    "見つかったJASRACコードを以下に入力して「反映」してください。"
                )
                _jw_col1, _jw_col2 = st.columns([3, 1])
                with _jw_col1:
                    _jw_manual_code = st.text_input(
                        "JASRACコード（例: 0M010710）",
                        key=f"jwid_manual_code_{selected_no}",
                        placeholder="作品コードを入力",
                    )
                with _jw_col2:
                    st.write("")
                    _jw_fetch_btn = st.button(
                        "📋 J-WID情報取得",
                        key=f"jwid_manual_fetch_{selected_no}",
                        use_container_width=True,
                        disabled=not _jw_manual_code.strip(),
                    )
                if _jw_fetch_btn and _jw_manual_code.strip():
                    with st.spinner("J-WID から取得中..."):
                        from modules.scraper import fetch_jwid_rights_by_code as _fwrm
                        _jw_m = _fwrm(_jw_manual_code.strip())
                        st.session_state[_jwid_manual_key] = _jw_m
                    if _jw_m.get("error"):
                        st.error(f"J-WID エラー: {_jw_m['error']}")
                    else:
                        st.success(
                            f"作曲者: {_jw_m.get('作曲者','(なし)')}  "
                            f"作詞者: {_jw_m.get('作詞者','(なし)')}  "
                            f"訳詞者: {_jw_m.get('訳詞者','(なし)')}"
                        )
                _jw_m_data = st.session_state.get(_jwid_manual_key, {})
                if _jw_m_data and not _jw_m_data.get("error"):
                    _jw_disp_cols = st.columns(3)
                    _jw_disp_cols[0].text_input("作曲者", value=_jw_m_data.get("作曲者",""), disabled=True, key=f"jw_m_comp_{selected_no}")
                    _jw_disp_cols[1].text_input("作詞者", value=_jw_m_data.get("作詞者",""), disabled=True, key=f"jw_m_lyric_{selected_no}")
                    _jw_disp_cols[2].text_input("訳詞者", value=_jw_m_data.get("訳詞者",""), disabled=True, key=f"jw_m_trans_{selected_no}")
                    _jw_disp_cols2 = st.columns(3)
                    _jw_disp_cols2[0].text_input("編曲者", value=_jw_m_data.get("編曲者",""), disabled=True, key=f"jw_m_arr_{selected_no}")
                    if st.button("✅ J-WID情報を申告フォーマットに反映", key=f"jwid_manual_apply_{selected_no}", use_container_width=True, type="primary"):
                        _jw_apply = {
                            "作曲者":          _jw_m_data.get("作曲者",""),
                            "作詞者":          _jw_m_data.get("作詞者",""),
                            "訳詞者":          _jw_m_data.get("訳詞者",""),
                            "編曲者":          _jw_m_data.get("編曲者",""),
                            "JASRAC作品コード": _jw_manual_code.strip(),
                            "確認ステータス":  "候補あり",
                        }
                        _hy = _infer_houyo(_jw_manual_code.strip())
                        if _hy and not str(row.get("邦洋区分","")).strip():
                            _jw_apply["邦洋区分"] = _hy
                        _apply_iv_from_credits(_jw_apply)
                        for _col, _val in _jw_apply.items():
                            if _val and _col in st.session_state.songs_df.columns:
                                st.session_state.songs_df.at[row_idx, _col] = _val
                        st.session_state["_apply_msg"] = "楽曲まとめ・申告フォーマットに反映しました。"
                        st.session_state.pop("songs_editor", None)
                        st.rerun()

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
# Excel 出力 + J-WID CSV（tabs[0] 末尾）
# =====================================================================
with tabs[0]:
    st.divider()
    st.subheader("📊 Excel 出力")

    if st.session_state.songs_df is None:
        st.info("上の設定セクションで照合を実行してください。")
    else:
        _shinkok_df = None
        if st.session_state.events_df is not None and len(st.session_state.events_df) > 0:
            _shinkok_df = build_shinkok_df(st.session_state.songs_df, st.session_state.events_df)

        # ---- Excel 出力 ----
        st.markdown(
            """
            以下のシートを含む Excel ファイルを生成します：

            | シート名 | 内容 |
            |---------|------|
            | 申告フォーマット | **イベント単位・申告列名** の提出用シート（オレンジヘッダー） |
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
                    shinkok_df=_shinkok_df,
                )

            # 書き出した内容とDBの中身がずれないよう、ここでも保存する。
            # 下のダウンロードボタンは押すと再実行で消えてしまうので、
            # 確実に一度だけ通る生成側に置く。
            if _autosave_to_db("（Excel生成時）"):
                st.info(st.session_state.pop("_autosave_msg"))

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

        # ---- J-WID 管理状況 CSV ----
        st.divider()
        st.subheader("📋 J-WID 管理状況 CSV")
        st.caption(
            "JASRAC作品コードが入力されている楽曲の管理状況（○△×）を J-WID から取得し、CSV でダウンロードします。"
            "　取得は 1件あたり約 2秒かかります。"
        )

        # JASRAC作品コードがある楽曲を抽出
        _rts_df = st.session_state.songs_df
        _rts_mask = _rts_df.get("JASRAC作品コード", pd.Series(dtype=str)).apply(
            lambda v: bool(str(v).strip()) and str(v).strip().lower() != "nan"
        )
        _rts_songs = _rts_df[_rts_mask][["No", "曲名", "JASRAC作品コード"]].copy() if _rts_mask.any() else pd.DataFrame()

        if _rts_songs.empty:
            st.info("JASRAC作品コードが入力されている楽曲がありません。「検索補助」タブで J-WID 検索後に「適用」してください。")
        else:
            st.dataframe(_rts_songs, use_container_width=True, hide_index=True)
            st.caption(f"{len(_rts_songs)} 件の楽曲に JASRAC 作品コードが設定されています。")

            if st.button("🔍 J-WID 管理状況を一括取得", type="primary", use_container_width=True):
                from modules.scraper import fetch_jwid_rights_by_code as _fwr
                _rights_results: dict = {}
                _prog = st.progress(0)
                _stat = st.empty()
                _n = len(_rts_songs)
                for _ri, (_, _rs_row) in enumerate(_rts_songs.iterrows()):
                    _code = str(_rs_row.get("JASRAC作品コード", "")).strip()
                    _name = str(_rs_row.get("曲名", "")).strip()
                    _stat.text(f"取得中: {_ri+1}/{_n} — {_name} ({_code})")
                    try:
                        _d = _fwr(_code)
                        _rights_results[_code] = _d
                    except Exception as _e:
                        _rights_results[_code] = {"error": str(_e), "管理状況": {}}
                    _prog.progress((_ri + 1) / _n)
                st.session_state["jwid_rights_batch"] = _rights_results
                _prog.empty()
                _stat.empty()

                # 取れた放送・配信は楽曲側にも書き取る。CSV を落とさなくても
                # 申告フォーマットの表でそのまま見られるようにするため。
                # 台帳で当たった行は J-WID の詳細を引いていないので、ここが
                # 放送・配信の入り口になる
                _rts_all = st.session_state.songs_df
                _rts_upd = 0
                for _rts_idx in _rts_all.index:
                    _rts_code = str(_rts_all.at[_rts_idx, "JASRAC作品コード"]).strip()
                    if not _rts_code:
                        continue
                    _rts_upd += _write_mgmt_from_fetch(
                        _rts_all, _rts_idx,
                        (_rights_results.get(_rts_code) or {}).get("管理状況") or {},
                        "J",
                    )

                st.success(f"✅ {_n} 件の取得完了")
                if _rts_upd and _autosave_to_db("（放送・配信の取り込み）"):
                    st.session_state["_autosave_msg"] = (
                        f"放送・配信を {_rts_upd} 欄ぶん楽曲に書き取りました"
                    )
                st.rerun()

            # 取得済み結果の表示と CSV ダウンロード
            _batch = st.session_state.get("jwid_rights_batch", {})
            if _batch:
                # 全管理状況フィールド
                _ALL_FIELDS = [
                    "演奏会等", "上映/BGM", "社交場/ｶﾗｵｹ",
                    "録音", "出版", "貸与", "ビデオ", "映画",
                    "放送", "配信", "通カラ",
                    "広告/CM送録", "広告/映録", "広告/録音", "広告/ビデオ", "広告/出版",
                    "ゲーム/録音", "ゲーム/ビデオ",
                ]
                _rts_rows = []
                for _, _rs_row in _rts_songs.iterrows():
                    _code = str(_rs_row.get("JASRAC作品コード", "")).strip()
                    _result = _batch.get(_code, {})
                    _mgmt = _result.get("管理状況", {})
                    _row_out = {
                        "No":              _rs_row.get("No", ""),
                        "曲名":            _rs_row.get("曲名", ""),
                        "JASRAC作品コード": _code,
                        "作曲者":          _result.get("作曲者", ""),
                        "作詞者":          _result.get("作詞者", ""),
                    }
                    for _f in _ALL_FIELDS:
                        _row_out[_f] = _mgmt.get(_f, "?")
                    if _result.get("error"):
                        _row_out["エラー"] = _result["error"]
                    _rts_rows.append(_row_out)

                _rts_result_df = pd.DataFrame(_rts_rows)
                st.dataframe(_rts_result_df, use_container_width=True, hide_index=True)

                _rts_csv = _rts_result_df.to_csv(index=False, encoding="utf-8-sig")
                st.download_button(
                    label="⬇️ 管理状況 CSV ダウンロード",
                    data=_rts_csv.encode("utf-8-sig"),
                    file_name="JASRAC管理状況.csv",
                    mime="text/csv",
                    use_container_width=True,
                    type="primary",
                )
