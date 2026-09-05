"""
見た目の共通部品。

色・書体そのものは .streamlit/config.toml で決めている。ここには
config.toml では書けないもの（余白の詰め、上部ステータスバー、表の
状態色）だけを置く。

方針:
  * 藍をアクセントにする。緑・黄・橙・赤は「確認ステータス」の意味に
    予約しているので、飾りには使わない。
  * 色は Streamlit のテーマ変数（--primary-color など）を経由する。
    直に色を書くとダークテーマで読めなくなるため。
  * 内部の class 名は Streamlit の更新で変わるので当てにしない。
    data-testid と、こちらで付けた class だけを狙う。
"""
from __future__ import annotations

import html

import pandas as pd
import streamlit as st

# 「確認ステータス」の値 → 意味の段階。
# 表の色と、上部バーの進捗の数え方の両方でこれを使う。
_OK = "ok"        # 人が確定させた
_HIT = "hit"      # 照合で当たった。裏取りはまだ
_LEDGER = "ledger"  # 管理番号で自社の台帳に当たった
_MAYBE = "maybe"  # 当たりが1つだけ出て、そのまま入った
_MULTI = "multi"  # 当たりが複数。人がどれか選ぶ必要がある
_TODO = "todo"    # 人が見に行く必要がある
_NONE = "none"    # 当たらなかった
_IDLE = "idle"    # 手つかず

_STATUS_TONE: dict[str, str] = {
    "確定": _OK,
    "作曲者一致": _HIT,
    "アーティスト一致": _HIT,
    # 管理番号で自社の台帳に当たった。番号は曲ごとに固有なので、
    # どの曲かは決まっている
    "台帳一致": _LEDGER,
    # 管理番号が無い曲を、トラック番号＋曲名で台帳に当てた。番号ほど
    # 強くはないが、当たりが1つに絞れたときだけ書いている
    "台帳一致（曲名）": _MAYBE,
    "候補あり": _MAYBE,
    "複数候補あり": _MULTI,
    "要確認": _TODO,
    "J-WID要確認": _TODO,
    "NexTone要確認": _TODO,
    "ライブラリ元確認": _TODO,
    "Audiostock確認": _TODO,
    "MP3補助確認": _TODO,
    "該当なし": _NONE,
    "未調査": _IDLE,
}

# 表のセルに敷く色。半透明にしてあるのは、下地（明/暗）に馴染ませて
# 文字色をテーマ任せにできるようにするため。文字色は指定しない。
_CELL_BG: dict[str, str] = {
    _OK:     "rgba(27, 122, 86, 0.16)",   # 緑
    _HIT:    "rgba(47, 75, 143, 0.14)",   # 藍
    _LEDGER: "rgba(47, 75, 143, 0.14)",   # 藍。当たり方が違うだけで段階は同じ
    _MAYBE: "rgba(168, 115, 11, 0.10)",   # 薄い黄。1件だけなので弱く
    _MULTI: "rgba(168, 115, 11, 0.16)",   # 黄
    _TODO:  "rgba(184, 92, 30, 0.16)",    # 橙
    _NONE:  "rgba(179, 38, 30, 0.13)",    # 赤
    _IDLE:  "",                            # 手つかずは無色。数が多いので
}


def status_tone(value) -> str:
    """確認ステータスの値を段階に直す。未知の値は「手つかず」扱い。"""
    if value is None:
        return _IDLE
    return _STATUS_TONE.get(str(value).strip(), _IDLE)


#: 状態欄の印の読み方。表の見出しと下に添えて出す
ISSUE_MARK_LEGEND = (
    "🔵 TSP台帳から反映　🔴 空欄あり　🟡 権利状態注意　（印なし＝そのまま出せる）"
)

#: 自社CDの台帳（TSP）から値を入れた行に付ける印。検索で引いてきた
#: 行と見分けるためのもので、空欄や権利の印とは別に頭へ付ける
LEDGER_MARK = "🔵"

#: 空とみなす値。DB を通ると空欄が nan や None の字で戻ってくる
_ISSUE_BLANKS = ("", "nan", "none")


def _issue_val(row, *names) -> str:
    """行から、名前の挙がっている欄のうち最初に見つかったものを文字列で返す。

    申告フォーマットの表（レコード番号・作詞…）と楽曲まとめ（CD番号・
    作詞者…）で欄の名前が違うので、どちらでも引けるようにしてある。
    空とみなす値は空文字にそろえる。
    """
    for name in names:
        try:
            v = row[name]
        except (KeyError, IndexError, TypeError):
            continue
        s = str(v).strip() if v is not None else ""
        return "" if s.lower() in _ISSUE_BLANKS else s
    return ""


def issue_mark(row) -> str:
    """申告フォーマット1行の中身を見て、状態欄の印を返す。

    🔴 空欄あり  … 申告に要る欄が埋まっていない。まだ手を入れる行。
    🟡 権利状態注意 … 欄は埋まっているが、権利の状態に気をつける行。
    印なし        … そのまま出せる行。

    放送・配信は「○かどうか」で見るので、赤ではなく黄の側に入れてある。
    まだ引いていない（空の）ときも○ではないので黄になる。表の下の
    「📡 放送・配信を引く」から埋められる。
    """
    iv = _issue_val(row, "I/V区分")
    holes = (
        not _issue_val(row, "レコード会社名"),
        not _issue_val(row, "レコード番号", "CD番号"),
        not _issue_val(row, "邦・洋区分", "邦洋区分"),
        not iv,
        # ヴォーカルなら作詞者がいるはず。インストなら空でよい
        iv.startswith(("ヴォーカル", "V", "v"))
        and not _issue_val(row, "作詞", "作詞者"),
        # どちらの管理団体にも番号が無い＝まだ突き止められていない
        not _issue_val(row, "JASRACコード", "JASRAC作品コード")
        and not _issue_val(row, "NexTone管理番号"),
        not _issue_val(row, "作曲", "作曲者"),
        not _issue_val(row, "アーティスト"),
    )
    if any(holes):
        return "🔴"
    cautions = (
        _issue_val(row, "委任者") == "非委任者",
        not _issue_val(row, "放送").startswith("○"),
        not _issue_val(row, "配信").startswith("○"),
    )
    if any(cautions):
        return "🟡"
    return ""


def source_mark(source) -> str:
    """データの出どころの印。自社CDの台帳（TSP）から入れた行は 🔵。

    値は cd_master が行に書き残す「情報元」欄。検索で引いてきた行には
    何も書かれないので、印も付かない。
    """
    s = str(source or "").strip()
    return LEDGER_MARK if s and s.lower() not in _ISSUE_BLANKS else ""


def inject_css() -> None:
    """全画面共通のCSS。st.set_page_config の直後に一度だけ呼ぶ。"""
    st.markdown(
        """
<style>
  /* 上の余白を詰める。既定値は縦に大きく空きすぎて、
     1画面に収めたい表が押し出されてしまう */
  [data-testid="stMainBlockContainer"] {
      padding-top: 2.2rem;
      padding-bottom: 3rem;
      max-width: 1500px;
  }

  /* 「◯◯へ移動」で飛ぶ先の目印。画面に貼り付いている見出し帯の下に
     隠れてしまうので、その分だけ手前で止める */
  a[id^="sec-"] {
      display: block;
      scroll-margin-top: 5rem;
  }

  /* 管理番号・JASRAC作品コード・尺など、桁が縦に揃ってほしい値。
     等幅にしないと目視で照合できない */
  [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
      font-variant-numeric: tabular-nums;
  }

  /* タブ。既定は文字だけで、今どこにいるか分かりにくい */
  [data-testid="stTabs"] [role="tablist"] {
      gap: 0.25rem;
      border-bottom: 1px solid var(--border-color, rgba(128,128,128,0.25));
  }
  [data-testid="stTabs"] [role="tab"] {
      padding: 0.35rem 0.9rem;
  }

  /* ── 上部ステータスバー ───────────────────────────── */
  .cj-bar {
      display: flex;
      flex-wrap: wrap;
      align-items: baseline;
      gap: 0.4rem 1.4rem;
      padding: 0.55rem 0.9rem;
      margin-bottom: 1rem;
      border: 1px solid var(--border-color, rgba(128,128,128,0.25));
      border-left: 3px solid var(--primary-color, #2F4B8F);
      border-radius: 6px;
      background: var(--secondary-background-color, rgba(128,128,128,0.08));
  }
  .cj-bar-title {
      font-weight: 700;
      letter-spacing: 0.02em;
      margin-right: 0.4rem;
  }
  .cj-item {
      display: flex;
      align-items: baseline;
      gap: 0.35rem;
      font-size: 0.86rem;
      min-width: 0;
  }
  .cj-label {
      opacity: 0.6;
      font-size: 0.76rem;
      letter-spacing: 0.04em;
      white-space: nowrap;
  }
  .cj-value {
      font-weight: 600;
      /* 長い番組名でバーを折り返させない */
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      max-width: 22rem;
  }
  /* 状態を色だけで示すと分かりにくいので、印（●）も併せて出す */
  .cj-value::before { content: ""; }
  .cj-ok::before,
  .cj-warn::before,
  .cj-off::before {
      content: "\\25CF ";
      font-size: 0.7em;
      vertical-align: 0.15em;
  }
  .cj-ok   { color: var(--green-color,  #1B7A56); }
  .cj-warn { color: var(--yellow-color, #A8730B); }
  .cj-off  { color: var(--gray-color,   #6B7186); }

  /* 進捗の細いバー。数字だけより、残りの量が一目で分かる */
  .cj-gauge {
      flex: 0 0 8rem;
      height: 5px;
      border-radius: 3px;
      background: rgba(128,128,128,0.25);
      overflow: hidden;
      align-self: center;
  }
  .cj-gauge > span {
      display: block;
      height: 100%;
      background: var(--primary-color, #2F4B8F);
  }
</style>
""",
        unsafe_allow_html=True,
    )


def status_bar(title: str, items: list[tuple[str, str, str]],
               progress: tuple[int, int] | None = None) -> None:
    """画面上部の1行ステータスバー。

    items は (見出し, 値, 調子) の並び。調子は "ok" / "warn" / "off" /
    "" のいずれか。値が空のものは出さない（項目が増えても、まだ関係の
    ない画面ではバーが伸びないようにするため）。
    progress は (済み, 全体)。
    """
    parts = [f'<span class="cj-bar-title">{html.escape(title)}</span>']
    for label, value, tone in items:
        if not value:
            continue
        cls = f"cj-value cj-{tone}" if tone else "cj-value"
        parts.append(
            f'<span class="cj-item">'
            f'<span class="cj-label">{html.escape(label)}</span>'
            f'<span class="{cls}" title="{html.escape(str(value))}">'
            f'{html.escape(str(value))}</span></span>'
        )
    if progress and progress[1]:
        done, total = progress
        pct = max(0, min(100, round(done * 100 / total)))
        parts.append(
            f'<span class="cj-item"><span class="cj-label">進捗</span>'
            f'<span class="cj-value">{done}/{total}</span></span>'
            f'<span class="cj-gauge" title="{pct}%"><span style="width:{pct}%">'
            f'</span></span>'
        )
    st.markdown(f'<div class="cj-bar">{"".join(parts)}</div>',
                unsafe_allow_html=True)


def style_status(df: pd.DataFrame, col: str = "確認ステータス"):
    """確認ステータス列に色を敷いた Styler を返す。

    st.dataframe はこの Styler を受け取れるが、st.data_editor は受け取れ
    ない（編集用の表には色を付けられない）。列が無い場合や色付けに失敗
    した場合は元の DataFrame をそのまま返し、表示自体は必ず通す。
    """
    if col not in df.columns:
        return df
    try:
        return df.style.map(
            lambda v: f"background-color: {_CELL_BG[status_tone(v)]}"
                      if _CELL_BG[status_tone(v)] else "",
            subset=[col],
        )
    except Exception:
        return df


def count_done(df: pd.DataFrame | None, col: str = "確認ステータス") -> tuple[int, int]:
    """(調べ終わった行数, 全行数)。手つかず以外を「済み」と数える。"""
    if df is None or df.empty or col not in df.columns:
        return (0, 0 if df is None else len(df))
    done = sum(1 for v in df[col] if status_tone(v) != _IDLE)
    return (done, len(df))
