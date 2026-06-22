"""
CSV読み込みモジュール
UTF-8 / UTF-8 BOM / Shift_JIS に自動対応
"""
import io
import chardet
import pandas as pd


def detect_encoding(file_bytes: bytes) -> str:
    """chardet でバイト列の文字コードを推定する"""
    result = chardet.detect(file_bytes)
    enc = result.get("encoding") or "utf-8"
    # cp932 / shift_jis 系を統一
    if enc.lower().replace("-", "_") in ("shift_jis", "sjis", "cp932", "ms932"):
        return "cp932"
    return enc


def read_csv_auto(file) -> tuple[pd.DataFrame, str]:
    """
    アップロードされたファイルを文字コード自動判定で読み込む。
    Returns: (DataFrame, 判定した文字コード名)
    """
    file_bytes = file.read()
    enc = detect_encoding(file_bytes)

    # 推定コードで試みて、失敗したらフォールバック順に再試行
    fallbacks = ["utf-8-sig", "utf-8", "cp932", "shift_jis", "latin-1"]
    if enc not in fallbacks:
        fallbacks.insert(0, enc)

    last_error = None
    for candidate in fallbacks:
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=candidate)
            return df, candidate
        except (UnicodeDecodeError, pd.errors.ParserError) as e:
            last_error = e

    raise ValueError(
        f"CSVの文字コードを判定できませんでした。"
        f"UTF-8 または Shift_JIS で保存されているか確認してください。\n詳細: {last_error}"
    )


# ---- 列バリデーション ----

# NUENDOのCSVは書き出しバージョンで列名が微妙に異なる場合があるため、
# 「絶対必要」な列だけ厳格にチェックし、後処理で別名を吸収する。
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
        # 正規名: [別名候補...]
        "トラック名": ["トラック", "Track Name", "Track"],
        "イベント名": ["Event Name", "Name", "クリップ名"],
        "ファイル名": ["File Name", "Source File", "Audio File"],
        "START TIME": ["スタートタイム", "Start", "開始時間", "Position"],
        "終了時間": ["End Time", "End"],
        "イン時間": ["Clip In", "In"],
        "アウト時間": ["Clip Out", "Out"],
        "長さ": ["Duration", "デュレーション", "Length"],
    }
    rename = {}
    for canonical, aliases in alias_map.items():
        if canonical not in df.columns:
            for alias in aliases:
                if alias in df.columns:
                    rename[alias] = canonical
                    break
    return df.rename(columns=rename)
