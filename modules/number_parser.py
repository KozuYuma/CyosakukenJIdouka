"""
管理番号分解モジュール
ライブラリ管理番号・Audiostock 管理番号の検出と分解
"""
import re

# --- 有効なライブラリ系列の定義 ---
# キー: 英字2文字, 値: 使用可能な数字（先頭1桁）の集合
VALID_LIBRARY_SERIES: dict[str, set[int]] = {
    "ST": {1, 2, 3, 4, 5, 6, 7},
    "AN": {1, 2, 3, 4, 5},
    "VO": {1, 2},
    "VJ": {1, 2},
}

# ライブラリ管理番号パターン（厳密版）: 数字1桁 + 英字2文字 - 3桁 - 2桁
# 例: 6ST-653-09
LIBRARY_PATTERN = re.compile(
    r"\b([1-7])(ST|AN|VO|VJ)-(\d{3})-(\d{2})\b",
    re.IGNORECASE,
)

# 広義ライブラリ管理番号パターン: 数字1桁 + 大文字英字2文字 - 3桁数字 (+ 省略可能な -2桁)
# 例: 1AN-146, 2MX-033-04, 3BG-512 など系列問わず
# ファイル名からのタイトル抽出用（parse_number では厳密版を使う）
BROAD_LIBRARY_PATTERN = re.compile(
    r"\b\d[A-Z]{2}-\d{3}(?:-\d{2})?\b",
)

# BPM・コード サフィックスパターン: _BPM_KEY (例: _113_C, _120_Am, _95_F#m)
_BPM_KEY_SUFFIX = re.compile(r"_\d+_[A-Ga-g][#b]?m?$")

# Audiostock 管理番号パターン
# 例: audiostock_856447（末尾は _ または文字列終端 or 非数字文字）
# \b を末尾に使うと数字の後に _ が続く場合にマッチしないため使わない
AUDIOSTOCK_PATTERN = re.compile(
    r"(audiostock_(\d+))(?=[_\s]|$)",
    re.IGNORECASE,
)


def _strip_id(text: str, id_str: str) -> str:
    """テキストから管理番号文字列を除去して曲名部分だけ返す"""
    cleaned = text.replace(id_str, "")
    title = re.sub(r"^[\s\-_]+|[\s\-_]+$", "", cleaned).strip()
    return _BPM_KEY_SUFFIX.sub("", title).strip()


def parse_library_number(text: str) -> dict | None:
    """
    テキストからライブラリ管理番号を検出・分解する。
    見つからない / 無効な系列なら None を返す。
    """
    m = LIBRARY_PATTERN.search(text)
    if not m:
        return None

    digit = int(m.group(1))
    series_code = m.group(2).upper()
    disc_num = m.group(3)
    track_num = m.group(4)

    if series_code not in VALID_LIBRARY_SERIES:
        return None
    if digit not in VALID_LIBRARY_SERIES[series_code]:
        return None

    management_id = f"{digit}{series_code}-{disc_num}-{track_num}"
    series = f"{digit}{series_code}"
    disc_id = f"{digit}{series_code}-{disc_num}"

    return {
        "管理番号種別": "ライブラリ管理番号",
        "元管理番号": management_id,
        "管理系列": series,
        "ライブラリ盤番号": disc_id,
        "トラック番号": track_num,
        "Audiostock管理番号": "",
        "検出曲名": _strip_id(text, management_id),
    }


def parse_audiostock_number(text: str) -> dict | None:
    """
    テキストから Audiostock 管理番号を検出・分解する。
    見つからなければ None を返す。
    """
    m = AUDIOSTOCK_PATTERN.search(text)
    if not m:
        return None

    full_id = m.group(1)          # "audiostock_856447"
    number = m.group(2)           # "856447"

    return {
        "管理番号種別": "Audiostock管理番号",
        "元管理番号": full_id,
        "管理系列": "",
        "ライブラリ盤番号": "",
        "トラック番号": "",
        "Audiostock管理番号": number,
        "検出曲名": _strip_id(text, full_id),
    }


def parse_number(text: str) -> dict:
    """
    テキストから管理番号を自動判別して分解する。
    どちらにも該当しない場合は空フィールド＋テキストそのものを曲名として返す。
    """
    if not text or str(text).strip() == "" or str(text).lower() == "nan":
        return _empty_result()

    text = str(text).strip()

    # Audiostock を先に試す（ライブラリ番号より識別が明確）
    result = parse_audiostock_number(text)
    if result:
        return result

    result = parse_library_number(text)
    if result:
        return result

    return _empty_result(title=text)


def _empty_result(title: str = "") -> dict:
    return {
        "管理番号種別": "",
        "元管理番号": "",
        "管理系列": "",
        "ライブラリ盤番号": "",
        "トラック番号": "",
        "Audiostock管理番号": "",
        "検出曲名": title,
    }
