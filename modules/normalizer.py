"""
文字列正規化モジュール
照合精度向上のための表記ゆれ対応
（全角/半角・大文字/小文字・記号・拡張子など）
"""
import re
import unicodedata

from .number_parser import LIBRARY_PATTERN, AUDIOSTOCK_PATTERN


def remove_extension(filename: str) -> str:
    """音声ファイルの拡張子を除去する"""
    return re.sub(r"\.(wav|mp3|aiff?|flac|m4a|ogg)$", "", filename, flags=re.IGNORECASE)


def normalize_for_match(text: str) -> str:
    """
    照合比較用に文字列を正規化して返す。
    実際のファイル名や曲名の保持には使わず、比較専用に使うこと。
    """
    if not text or str(text).strip() == "" or str(text).lower() == "nan":
        return ""

    text = str(text).strip()

    # 拡張子を除去
    text = remove_extension(text)

    # 全角文字を半角へ（NFKC 正規化）
    text = unicodedata.normalize("NFKC", text)

    # 小文字化
    text = text.lower()

    # アンダースコア・ハイフンをスペースに統一
    text = re.sub(r"[_\-]+", " ", text)

    # 括弧類を除去
    text = re.sub(r"[\(\)\[\]【】「」『』（）〔〕｛｝]", " ", text)

    # 記号を除去（句読点・中点など）
    text = re.sub(r"[!！?？。、・…～〜♪♫]", " ", text)

    # 連続スペースを1つに圧縮
    text = re.sub(r"\s+", " ", text).strip()

    return text


def detect_title_from_filename(filename: str) -> str:
    """
    WAV/MP3 ファイル名から曲タイトル候補を検出する。
    管理番号部分（ライブラリ管理番号・Audiostock）を除去した残りを曲名とする。
    """
    if not filename or str(filename).lower() == "nan":
        return ""

    name = str(filename).strip()
    name = remove_extension(name)

    # ライブラリ管理番号を除去
    name = LIBRARY_PATTERN.sub("", name)

    # Audiostock 管理番号を除去（flags は compile 時に設定済み）
    name = AUDIOSTOCK_PATTERN.sub("", name)

    # 先頭・末尾の区切り文字を除去
    name = re.sub(r"^[\s\-_]+|[\s\-_]+$", "", name).strip()

    return name
