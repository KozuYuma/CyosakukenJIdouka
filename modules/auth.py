"""
ログイン。合言葉を人ごとに配る方式。

環境変数 APP_USERS に「名前:合言葉」をカンマ区切りで並べる:

    APP_USERS=ゆま:hoge1234,たろう:fuga5678

  * 未設定なら認証しない。手元で動かすときは今までどおり素通りする。
  * 誰がログインしているか分かるので、案件を人ごとに分けられる。
  * 人の入れ替えは環境変数を書き換えるだけ。ユーザー表も登録画面も要らない。

合言葉を .env に書くこと。ソースには絶対に書かない（.env は
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
    """{名前: 合言葉}。APP_USERS が未設定なら空。"""
    load_local_env()
    raw = os.environ.get("APP_USERS", "").strip()
    if not raw:
        return {}
    users: dict[str, str] = {}
    for pair in raw.split(","):
        if ":" not in pair:
            continue
        name, pw = pair.split(":", 1)
        name, pw = name.strip(), pw.strip()
        if name and pw:
            users[name] = pw
    return users


def is_enabled() -> bool:
    """ログインを要求するかどうか。"""
    return bool(get_users())


def _verify(name: str, password: str) -> bool:
    """合言葉を確かめる。

    == ではなく compare_digest を使う。== は先頭から順に比べて違った
    時点で止まるので、返答の速さの差から合言葉を1文字ずつ言い当てられる。
    """
    expected = get_users().get(name)
    if expected is None:
        return False
    return hmac.compare_digest(expected, password)


def current_user() -> str:
    """ログイン中の名前。認証を使っていなければ LOCAL_USER。"""
    if not is_enabled():
        return LOCAL_USER
    return st.session_state.get(_STATE_KEY) or ""


def _login_form() -> None:
    st.title("🎵 著作権調査支援ツール")
    st.caption("合言葉を入れてください。分からない場合は管理者に聞いてください。")

    names = sorted(get_users())
    with st.form("login_form"):
        name = st.selectbox("名前", options=names, key="login_name")
        pw = st.text_input("合言葉", type="password", key="login_password")
        ok = st.form_submit_button("ログイン", type="primary",
                                   use_container_width=True)
    if ok:
        if _verify(name, pw):
            st.session_state[_STATE_KEY] = name
            # 合言葉を session_state に残さない
            st.session_state.pop("login_password", None)
            st.rerun()
        else:
            st.error("名前か合言葉が違います。")


def require_login() -> str:
    """ログインしていなければログイン画面を出してそこで止める。

    返り値はログイン中の名前。案件の所有者として使う。
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
