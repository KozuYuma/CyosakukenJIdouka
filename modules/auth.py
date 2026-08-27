"""
ログイン。ID と PASSWORD を人ごとに配る方式。

環境変数 APP_USERS に「ID:PASSWORD」をカンマ区切りで並べる:

    APP_USERS=yuma:hoge1234,taro:fuga5678

  * 未設定なら認証しない。手元で動かすときは今までどおり素通りする。
  * 誰がログインしているか分かるので、案件を人ごとに分けられる。
  * 人の入れ替えは環境変数を書き換えるだけ。ユーザー表も登録画面も要らない。

PASSWORD は .env に書くこと。ソースには絶対に書かない（.env は
.gitignore に入っている）。
"""
from __future__ import annotations

import hmac
import os

import streamlit as st

from modules.database import load_local_env

# APP_USERS が未設定のときの所有者名。手元で動かす場合はこれになる
LOCAL_USER = "local"

_STATE_KEY = "_auth_user"


def get_users() -> dict[str, str]:
    """{ID: PASSWORD}。APP_USERS が未設定なら空。"""
    load_local_env()
    raw = os.environ.get("APP_USERS", "").strip()
    if not raw:
        return {}
    users: dict[str, str] = {}
    for pair in raw.split(","):
        if ":" not in pair:
            continue
        user_id, pw = pair.split(":", 1)
        user_id, pw = user_id.strip(), pw.strip()
        if user_id and pw:
            users[user_id] = pw
    return users


def is_enabled() -> bool:
    """ログインを要求するかどうか。"""
    return bool(get_users())


def _verify(user_id: str, password: str) -> bool:
    """ID と PASSWORD を確かめる。

    == ではなく compare_digest を使う。== は先頭から順に比べて違った
    時点で止まるので、返答の速さの差から PASSWORD を1文字ずつ
    言い当てられる。
    """
    expected = get_users().get(user_id)
    if expected is None:
        return False
    return hmac.compare_digest(expected, password)


def current_user() -> str:
    """ログイン中の ID。認証を使っていなければ LOCAL_USER。"""
    if not is_enabled():
        return LOCAL_USER
    return st.session_state.get(_STATE_KEY) or ""


def _login_form() -> None:
    st.title("🎵 著作権調査支援ツール")
    st.caption("ID と PASSWORD を入れてください。"
               "分からない場合は管理者に聞いてください。")

    # ID もプルダウンではなく手入力にする。一覧から選ばせると、
    # ログインしていない人にも全員の ID が見えてしまうため。
    #
    # autocomplete は指定しない。ブラウザの自動入力は欄に文字を表示しても
    # Streamlit 側に値が届かないことがあり、見た目は埋まっているのに
    # 「違います」になる。手で打てば確実に届く。
    with st.form("login_form"):
        user_id = st.text_input("ID", key="login_id")
        pw = st.text_input("PASSWORD", type="password", key="login_password")
        ok = st.form_submit_button("ログイン", type="primary",
                                   use_container_width=True)
    if ok:
        if _verify(user_id.strip(), pw):
            st.session_state[_STATE_KEY] = user_id.strip()
            # PASSWORD を session_state に残さない
            st.session_state.pop("login_password", None)
            st.rerun()
        else:
            # どちらが違うかは言わない。ID だけ当てられるのを避けるため
            st.error("ID または PASSWORD が違います。")
            # 受け取った「文字数」だけ出す。中身は出さないので漏れない。
            # 欄が埋まって見えるのに 0 文字なら、ブラウザの自動入力が
            # 表示だけしていて値が届いていない。手で打ち直せば直る。
            st.caption(
                f"受け取った内容: ID {len(user_id)} 文字 / "
                f"PASSWORD {len(pw)} 文字"
            )


def require_login() -> str:
    """ログインしていなければログイン画面を出してそこで止める。

    返り値はログイン中の ID。案件の所有者として使う。
    """
    if not is_enabled():
        return LOCAL_USER
    if current_user():
        return current_user()
    _login_form()
    st.stop()


def logout_button() -> None:
    """ログアウト。作業中の内容も消す（次の人に見せないため）。"""
    if not is_enabled():
        return
    if st.button("ログアウト", key="btn_logout", use_container_width=True):
        for k in (_STATE_KEY, "songs_df", "events_df", "cue_df", "wav_df",
                  "mp3_df", "search_df", "master_db_df",
                  "project_id", "project_name"):
            st.session_state.pop(k, None)
        st.rerun()
