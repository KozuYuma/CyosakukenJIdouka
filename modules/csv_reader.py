"""
CSV読み込みモジュール
UTF-8 / UTF-8 BOM / Shift_JIS に自動対応
区切り文字（カンマ / タブ / セミコロン）も自動判定
"""
import io
import chardet
import pandas as pd


def detect_encoding(file_bytes: bytes) -> str:
    """chardet でバイト列の文字コードを推定する"""
    result = chardet.detect(file_bytes)
    enc = result.get("encoding") or "utf-8"
    if enc.lower().replace("-", "_") in ("shift_jis", "sjis", "cp932", "ms932"):
        return "cp932"
    return enc


def _detect_separator(text: str) -> str:
    """
    先頭数行を見て区切り文字を推定する。
    NUENDO はバージョンによりタブ区切りで書き出す場合がある。
    """
    head = "\n".join(text.splitlines()[:5])
    tab_count   = head.count("\t")
    comma_count = head.count(",")
    semi_count  = head.count(";")

    if tab_count >= comma_count and tab_count >= semi_count:
        return "\t"
    if semi_count > comma_count:
        return ";"
    return ","


def read_csv_auto(file) -> tuple[pd.DataFrame, str]:
    """
    アップロードされたファイルを文字コード・区切り文字ともに自動判定して読み込む。
    Returns: (DataFrame, "文字コード / 区切り文字" の説明文字列)
    """
    file_bytes = file.read()
    enc = detect_encoding(file_bytes)

    enc_candidates = ["utf-8-sig", "utf-8", "cp932", "shift_jis", "euc-jp", "latin-1"]
    if enc not in enc_candidates:
        enc_candidates.insert(0, enc)

    last_error = None

    for candidate_enc in enc_candidates:
        # まずデコードできるか確認
        try:
            text = file_bytes.decode(candidate_enc)
        except (UnicodeDecodeError, LookupError):
            continue

        sep = _detect_separator(text)

        # sep と python エンジンで読む（C エンジンより寛容）
        for sep_try in [sep, ",", "\t", ";"]:
            try:
                df = pd.read_csv(
                    io.BytesIO(file_bytes),
                    encoding=candidate_enc,
                    sep=sep_try,
                    engine="python",      # C エンジンより列数ゆれに強い
                    on_bad_lines="warn",  # 壊れた行はスキップして続行
                )
                # 1列しか取れなかった場合は区切り文字ミス
                if len(df.columns) <= 1:
                    continue
                sep_label = {"," : "カンマ", "\t": "タブ", ";": "セミコロン"}.get(sep_try, sep_try)
                return df, f"{candidate_enc} / {sep_label}区切り"
            except Exception as e:
                last_error = e

    raise ValueError(
        f"CSV を読み込めませんでした。\n"
        f"・NUENDO から書き出した CSV をそのまま使っているか確認してください\n"
        f"・Excel で開いて保存し直すと解決する場合があります\n"
        f"詳細: {last_error}"
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
