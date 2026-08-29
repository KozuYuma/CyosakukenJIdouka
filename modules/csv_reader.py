"""
CSV読み込みモジュール
UTF-8 / UTF-8 BOM / Shift_JIS に自動対応
区切り文字（カンマ / タブ / セミコロン）も自動判定

NUENDOのCue CSVは独自のマルチセクション形式のため専用パーサーを使用する。
構造:
  [プロジェクト情報]  2列のメタデータ
  [ファイルリスト]    2列の名前-パス一覧
  [オーディオトラックリスト]
    トラック - M1-1
      イベント名,ファイル名,START TIME,...  ← 各トラックに個別のヘッダー
      データ行...
    トラック - M2-1
      イベント名,...
      データ行...
"""
import io
import unicodedata

import chardet
import pandas as pd

# ─── 楽曲として拾わないもの ──────────────────────────────
# NUENDO の書き出しには音楽以外のトラックも入る。トラック名で外す。
# 部分一致で見るので「マーカートラックリスト」でも「マーカー」でも当たる。
NON_MUSIC_TRACKS: tuple[str, ...] = (
    "構成",
    "ノートパッド",
    "マーカー",
    "ビデオ",
)

# イベント名にこれが入っていたら楽曲ではない。1KHZ-20DB-04 のような
# 基準信号が各トラックの先頭に入っているため
NON_MUSIC_EVENTS: tuple[str, ...] = (
    "1khz",
)


def _norm(value) -> str:
    """全角半角・大文字小文字・空白の揺れを均した比較用の文字列。"""
    s = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return "".join(s.split())


def is_music_track(name) -> bool:
    """楽曲を並べるトラックかどうか。M1-1 などは True。"""
    n = _norm(name)
    if not n:
        return True     # トラック名が無い形式は今までどおり全部使う
    return not any(word in n for word in map(_norm, NON_MUSIC_TRACKS))


def is_music_event(name) -> bool:
    """楽曲のイベントかどうか。基準信号などは False。"""
    n = _norm(name)
    return not any(word in n for word in NON_MUSIC_EVENTS)


def drop_non_music_events(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """基準信号などのイベント行を落とす。(残った df, 落とした件数)。"""
    if df is None or df.empty or "イベント名" not in df.columns:
        return df, 0
    keep = df["イベント名"].map(is_music_event)
    dropped = int((~keep).sum())
    if not dropped:
        return df, 0
    return df[keep].reset_index(drop=True), dropped


def music_tracks(df: pd.DataFrame) -> tuple[list[str], list[str]]:
    """Cue の中のトラックを (楽曲用, それ以外) に分ける。"""
    if df is None or df.empty or "トラック名" not in df.columns:
        return [], []
    names = [str(t) for t in df["トラック名"].dropna().unique() if str(t).strip()]
    return ([t for t in names if is_music_track(t)],
            [t for t in names if not is_music_track(t)])


def detect_encoding(file_bytes: bytes) -> str:
    """chardet でバイト列の文字コードを推定する"""
    result = chardet.detect(file_bytes)
    enc = result.get("encoding") or "utf-8"
    if enc.lower().replace("-", "_") in ("shift_jis", "sjis", "cp932", "ms932"):
        return "cp932"
    return enc


def _detect_separator(text: str) -> str:
    """先頭10行を見て区切り文字を推定する"""
    head = "\n".join(text.splitlines()[:10])
    counts = {"\t": head.count("\t"), ",": head.count(","), ";": head.count(";")}
    return max(counts, key=counts.get)


def _is_cue_header(line: str) -> bool:
    """
    Cueデータのヘッダー行かどうか判定する。
    「イベント名」と「ファイル名」が両方含まれる行のみを対象にする。
    「長さ」単独などのメタデータ行を誤検出しないようにするため
    2つのキーワードを必須とする。
    """
    return "イベント名" in line and "ファイル名" in line


def _is_spaced_title(value: str) -> bool:
    """"マ ー カ ー ト ラ ッ ク リ ス ト" のような一文字ずつ空けた見出しか。

    NUENDO はセクションの見出しをこの書き方で出す。全部の塊が一文字
    なら見出しとみなす。「ノートパッド」のような普通の語は当たらない。
    """
    parts = value.split()
    return len(parts) >= 2 and all(len(p) == 1 for p in parts)


def _parse_nuendo_multitrack(text: str, sep: str) -> pd.DataFrame | None:
    """
    NUENDOのマルチトラックCue CSV専用パーサー。
    複数の「トラック - Mxx」セクションを統合して1つのDataFrameを返す。
    どのセクションにもCueデータがなければ None を返す。
    """
    lines = text.splitlines()
    all_rows: list[dict] = []
    current_headers: list[str] | None = None
    current_track: str = ""
    n_expected: int = 0

    for line in lines:
        stripped = line.strip()

        # 書き出しによっては全部の行に列数ぶんのカンマが付く。
        # 例: "トラック - M2-1,,,,,,,,,,"
        # 末尾の空欄を落としてから見ないと、見出し行をデータ行と
        # 間違えて「トラック - M2-1」が曲として並んでしまう
        fields_all = [f.strip() for f in stripped.split(sep)]
        while fields_all and fields_all[-1] == "":
            fields_all.pop()

        if not fields_all:
            continue

        # --- 中身が1つだけの行＝見出し。データ行ではない ---
        if len(fields_all) == 1:
            only = fields_all[0]

            # トラック名行。例: "トラック - M1-1"
            if " - " in only:
                candidate = only.split(" - ", 1)[1].strip()
                # スペース区切りのセクションタイトル（文字間にスペース）は除外
                # 例: "ト ラ ッ ク  リ ス ト" は除外、"M1-1" は採用
                if candidate and " " not in candidate.replace("-", ""):
                    current_track = candidate
                    current_headers = None
                    continue

            # セクション見出し。例: "マ ー カ ー ト ラ ッ ク リ ス ト"
            # 一文字ずつ空けて書かれるので、そこで見分ける。
            # 空白を詰めた名前をトラック名にしておけば、あとで
            # 「マーカー」「ビデオ」として外せる
            if _is_spaced_title(only):
                current_track = "".join(only.split())
                current_headers = None
                continue

            # 「ノートパッド」のようなトラックの付随情報。曲ではない
            continue

        # --- Cueヘッダー行の検出 ---
        if _is_cue_header(stripped):
            current_headers = [f.strip() for f in stripped.split(sep)]
            n_expected = len(current_headers)
            continue

        # --- データ行の処理 ---
        # ヘッダーが出る前の行（"構成,2" など）はここに来ない。
        # 区切り文字はファイルごとに違う（NUENDO はタブで書き出すことがある）。
        # ここをカンマ決め打ちにすると、タブ区切りのときに1行も拾えず、
        # 「ふつうのCSV」として最後まで読まれてしまい、2本目以降の
        # 見出し行（トラック - M2-1 / 構成 / ノートパッド / 2回目の
        # イベント名 / 末尾のセクション見出し）が曲として並んでしまう
        if current_headers and stripped and sep in stripped:
            # n_expected 列に合わせて分割（最後の列にカンマが含まれても壊れないよう maxsplit 指定）
            fields = [f.strip() for f in stripped.split(sep, n_expected - 1)]

            # 列数が少なすぎる行（メタデータ行等）はスキップ
            if len(fields) < n_expected - 1:
                continue

            row: dict = {}
            for i, h in enumerate(current_headers):
                row[h] = fields[i] if i < len(fields) else ""

            # トラック名を付加
            if current_track:
                row["トラック名"] = current_track

            # イベント名が空でない行だけ採用
            first_col = current_headers[0]
            if row.get(first_col, "").strip():
                all_rows.append(row)

    if not all_rows:
        return None

    df = pd.DataFrame(all_rows)

    # トラック名列を先頭に移動
    if "トラック名" in df.columns:
        cols = ["トラック名"] + [c for c in df.columns if c != "トラック名"]
        df = df[cols]

    return df.reset_index(drop=True)


def _columns_look_valid(columns) -> bool:
    """
    列名が文字化けしていないか確認する。
    CueCSVなら「イベント名+ファイル名」、WAV一覧なら「FileName」が含まれるはず。
    """
    col_str = " ".join(str(c) for c in columns)
    return (
        ("イベント名" in col_str and "ファイル名" in col_str)
        or "FileName" in col_str
        or "Event" in col_str
    )


def read_csv_auto(file) -> tuple[pd.DataFrame, str]:
    """
    アップロードされたファイルを文字コード・区切り文字・ヘッダー行を
    すべて自動判定して読み込む。
    NUENDOのマルチセクション形式にも対応。

    Returns: (DataFrame, 判定内容の説明文字列)
    """
    file_bytes = file.read()
    enc_guess = detect_encoding(file_bytes)

    # latin-1 は最後（日本語ファイルでは必ず文字化けするため）
    enc_candidates = ["utf-8-sig", "utf-8", "cp932", "shift_jis", "euc-jp", "latin-1"]
    if enc_guess not in enc_candidates:
        enc_candidates.insert(0, enc_guess)

    last_error = None

    for enc in enc_candidates:
        try:
            text = file_bytes.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue

        sep = _detect_separator(text)

        # --- NUENDOマルチセクション形式を最初に試みる ---
        if _is_cue_header_in_text(text):
            df = _parse_nuendo_multitrack(text, sep)
            if df is not None and len(df.columns) > 1:
                sep_label = {",": "カンマ", "\t": "タブ", ";": "セミコロン"}.get(sep, sep)
                return df, f"{enc} / {sep_label}区切り / NUENDOマルチセクション形式"

        # --- 通常のCSVとして読む ---
        skip_candidates = [_find_simple_header_row(text.splitlines()), 0]
        skip_candidates = list(dict.fromkeys(skip_candidates))  # 重複除去

        for sep_try in [sep, ",", "\t", ";"]:
            for skip in skip_candidates:
                try:
                    df = pd.read_csv(
                        io.BytesIO(file_bytes),
                        encoding=enc,
                        sep=sep_try,
                        skiprows=skip,
                        engine="python",
                        on_bad_lines="warn",
                    )
                    if len(df.columns) <= 1:
                        continue
                    if not _columns_look_valid(df.columns):
                        continue
                    sep_label2 = {",": "カンマ", "\t": "タブ", ";": "セミコロン"}.get(sep_try, sep_try)
                    desc = f"{enc} / {sep_label2}区切り"
                    if skip:
                        desc += f" / 先頭{skip}行スキップ"
                    return df, desc
                except Exception as e:
                    last_error = e

    raise ValueError(
        "CSV を読み込めませんでした。\n\n"
        "【確認してください】\n"
        "・NUENDOから書き出したCSVをそのまま使っているか\n"
        "・Excelで開いて「名前を付けて保存」→「CSV UTF-8（コンマ区切り）」で保存し直す\n"
        f"\n詳細: {last_error}"
    )


def _is_cue_header_in_text(text: str) -> bool:
    """テキスト全体にCueヘッダー行が含まれるか確認する"""
    for line in text.splitlines():
        if _is_cue_header(line):
            return True
    return False


def _find_simple_header_row(lines: list[str]) -> int:
    """
    通常CSV用: イベント名とファイル名が共存するヘッダー行を探す。
    見つからなければ0を返す。
    """
    for i, line in enumerate(lines):
        if _is_cue_header(line):
            return i
        if "FileName" in line:
            return i
    return 0


# ---- 列バリデーション ----

CUE_REQUIRED = ["イベント名", "ファイル名"]
WAV_REQUIRED = ["FileName"]


def validate_cue_csv(df: pd.DataFrame) -> list[str]:
    """Cue CSV の必須列チェック。不足列名リストを返す（空なら OK）"""
    return [c for c in CUE_REQUIRED if c not in df.columns]


def validate_wav_csv(df: pd.DataFrame) -> list[str]:
    """WAV 一覧 CSV の必須列チェック。不足列名リストを返す（空なら OK）"""
    return [c for c in WAV_REQUIRED if c not in df.columns]


def normalize_cue_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    NUENDOバージョン差異による列名ゆれを吸収するリネーム。
    既存列名はそのまま保持し、見つかった別名だけ正規名に変換する。
    """
    alias_map = {
        "トラック名": ["トラック", "Track Name", "Track"],
        "イベント名": ["Event Name", "Name", "クリップ名"],
        "ファイル名": ["File Name", "Source File", "Audio File"],
        "START TIME": ["スタートタイム", "Start", "開始時間", "Position"],
        "終了時間":   ["End Time", "End"],
        "イン時間":   ["Clip In", "In"],
        "アウト時間": ["Clip Out", "Out"],
        "長さ":       ["Duration", "デュレーション", "Length"],
    }
    rename = {}
    for canonical, aliases in alias_map.items():
        if canonical not in df.columns:
            for alias in aliases:
                if alias in df.columns:
                    rename[alias] = canonical
                    break
    return df.rename(columns=rename)
