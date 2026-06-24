"""
CSV読み込みモジュール
UTF-8 / UTF-8 BOM / Shift_JIS に自動対応
区切り文字（カンマ / タブ / セミコロン）も自動判定
NUENDOのCue CSVに含まれるメタデータ行を自動スキップ
"""
import io
import chardet
import pandas as pd


# ヘッダー行の判定に使うキーワード（日本語・英語）
_HEADER_KEYWORDS = [
    "イベント名", "ファイル名", "トラック名", "トラック",
    "Event Name", "File Name", "Track Name", "Track",
    "START TIME", "Duration", "Length",
]


def detect_encoding(file_bytes: bytes) -> str:
    """chardet でバイト列の文字コードを推定する"""
    result = chardet.detect(file_bytes)
    enc = result.get("encoding") or "utf-8"
    if enc.lower().replace("-", "_") in ("shift_jis", "sjis", "cp932", "ms932"):
        return "cp932"
    return enc


def _detect_separator(text: str) -> str:
    """先頭数行を見て区切り文字を推定する"""
    head = "\n".join(text.splitlines()[:10])
    counts = {"\t": head.count("\t"), ",": head.count(","), ";": head.count(";")}
    return max(counts, key=counts.get)


def _find_header_row(lines: list[str], sep: str) -> int:
    """
    実際のデータが始まるヘッダー行の行番号を返す。

    NUENDOのCue CSVは先頭に「プロジェクト名,値」のような
    2列のメタデータが数百行続いた後に本体テーブルが始まる場合がある。
    キーワードマッチ → フィールド数の変化点の順で探す。
    """
    # 1. ヘッダーキーワードを含む行を優先（最も確実）
    for i, line in enumerate(lines):
        if any(kw in line for kw in _HEADER_KEYWORDS):
            return i

    # 2. フィールド数が急増する行を探す（メタデータ終端 → データ開始）
    prev_count = 0
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        count = len(stripped.split(sep))
        if count >= 3 and count > prev_count + 1:
            return i
        prev_count = count

    return 0


def read_csv_auto(file) -> tuple[pd.DataFrame, str]:
    """
    アップロードされたファイルを文字コード・区切り文字・ヘッダー行を
    すべて自動判定して読み込む。

    Returns: (DataFrame, 判定内容の説明文字列)
    """
    file_bytes = file.read()
    enc_guess = detect_encoding(file_bytes)

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
        lines = text.splitlines()
        skiprows = _find_header_row(lines, sep)

        sep_label = {",": "カンマ", "\t": "タブ", ";": "セミコロン"}.get(sep, sep)
        desc = f"{enc} / {sep_label}区切り / {skiprows}行スキップ"

        for sep_try in [sep, ",", "\t", ";"]:
            for skip in ([skiprows] if skiprows > 0 else [0]):
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
                    # 空列・無名列が多すぎる場合は除外
                    named = [c for c in df.columns if not str(c).startswith("Unnamed")]
                    if len(named) < 2:
                        continue
                    sep_label2 = {",": "カンマ", "\t": "タブ", ";": "セミコロン"}.get(sep_try, sep_try)
                    return df, f"{enc} / {sep_label2}区切り" + (f" / 先頭{skip}行スキップ" if skip else "")
                except Exception as e:
                    last_error = e

    raise ValueError(
        "CSV を読み込めませんでした。\n\n"
        "【確認してください】\n"
        "・NUENDOから書き出したCSVをそのまま使っているか\n"
        "・Excelで開いて「名前を付けて保存」→「CSV UTF-8（コンマ区切り）」で保存し直す\n"
        f"\n詳細: {last_error}"
    )


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
