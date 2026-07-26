"""
文字列正規化モジュール
照合精度向上のための表記ゆれ対応
（全角/半角・大文字/小文字・記号・拡張子など）
"""
import re
import unicodedata

from .number_parser import BROAD_LIBRARY_PATTERN, LIBRARY_PATTERN, AUDIOSTOCK_PATTERN


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


def detect_title_from_filename(filename: str, strip_trailing: bool = True) -> str:
    """
    WAV/MP3 ファイル名（またはイベント名）から曲タイトル候補を検出する。

    除去ルール（順番通りに適用）:
      ①  広義ライブラリ管理番号（数字+大文字2文字+3桁、例: 1AN-146, 3BG-512-04）
      ②  Audiostock 管理番号（audiostock_XXXXXX）
      ③  先頭のディスク-トラック番号
            "01-03 Song"  / "01-03 - Song"  / "1-03 Song"
      ④  先頭の「2桁数字が2つ並ぶ」パターン
            "01 02 Song"  / "01_02 Song"
      ⑤  iTunes・CD リップ由来のトラック番号プレフィックス
            "01. Song" / "01 - Song" / "01 Song" / "Track 01 Song"
      ⑥  末尾のバリアント/バージョンマーカー（strip_trailing=True のときのみ）
            "-01" "_113_C" "_C"
      ⑦  末尾の年号・注釈（strip_trailing=True のときのみ）
            "(2005)" "[Bonus Track]" "[Live]"

    Args:
        strip_trailing: False にすると ⑥⑦ をスキップ。
                        イベント名を処理する場合は False を渡すと安全。
    """
    if not filename or str(filename).lower() == "nan":
        return ""

    name = str(filename).strip()
    name = remove_extension(name)

    # ① 広義ライブラリ管理番号を除去（系列問わず: 数字+2大文字-3桁+省略可-2桁）
    name = BROAD_LIBRARY_PATTERN.sub("", name)

    # ② Audiostock 管理番号を除去
    name = AUDIOSTOCK_PATTERN.sub("", name)

    # 先頭・末尾の空白や区切り文字を一時整理
    name = re.sub(r"^[\s\-_]+|[\s\-_]+$", "", name).strip()

    # ③ 先頭のディスク-トラック番号を除去
    #    "01-03 - Song" / "01-03 Song" / "1-03 - Song" / "1-03 Song"
    #    ※ライブラリ管理番号（1AN-146 等）は①で除去済み
    name = re.sub(r"^\d{1,2}-\d{2}(?:\s*-\s+|\s+)", "", name)

    # ④ 先頭の「2桁数字が2つ連続」パターン（スペース・ハイフン・アンダースコアで区切り）
    #    "01 02 Song" / "01_02 Song" / "01-02 Song"（ディスク-トラックでない場合の補完）
    name = re.sub(r"^\d{2}[\s_-]+\d{2}(?:\s*-\s+|\s+)", "", name)

    # ⑤ iTunes / CD リップ由来のトラック番号プレフィックス
    #    優先度順: "Track 01 - ", "01. ", "01 - ", "01 "
    name = re.sub(r"^[Tt]rack\s*\d{1,3}\s*[-\.\s]\s*", "", name)   # "Track 01 - "
    name = re.sub(r"^トラック\s*\d{1,3}\s*[-\.\s]\s*", "", name)    # "トラック01 "
    name = re.sub(r"^\d{1,3}\.\s+", "", name)                        # "01. Song"
    name = re.sub(r"^\d{1,3}\s+-\s+", "", name)                      # "01 - Song"
    name = re.sub(r"^\d{1,3}\s+", "", name)                          # "01 Song"

    if strip_trailing:
        # ⑥ 末尾のトラック番号・バリアントマーカーを除去
        name = re.sub(r"[-_]\d{2}$", "", name)          # "-01" "_09"
        name = re.sub(r"_\d+_[A-Za-z]+$", "", name)     # "_113_C"
        name = re.sub(r"_[A-Za-z]{1,2}$", "", name)     # "_C" "_AB"

        # ⑦ 末尾の年号・注釈（括弧付き）を除去
        #    "(2005)" "[Bonus Track]" "[Live]" "[Instrumental]" など
        name = re.sub(r"\s*[\(\[]\d{4}[\)\]]$", "", name)         # (2005) [2005]
        name = re.sub(
            r"\s*[\(\[](bonus\s*track|live|instrumental|remix|edit|ver\.?|version)[\)\]]$",
            "",
            name,
            flags=re.IGNORECASE,
        )

    # 最終的な先頭・末尾の整理
    name = re.sub(r"^[\s\-_]+|[\s\-_]+$", "", name).strip()

    return name
