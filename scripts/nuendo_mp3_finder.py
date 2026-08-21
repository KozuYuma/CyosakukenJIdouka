#!/usr/bin/env python3
"""
nuendo_mp3_finder.py

NUENDO Cue CSV からイベント名を抜き出し、指定フォルダ以下の
マッチする MP3 ファイルを収集してプロパティを表示する。

Usage:
    python nuendo_mp3_finder.py <cue.csv> <mp3_folder> [options]

    --workers N       並列スレッド数 (default: 8)
    --output FILE     結果を CSV に出力（省略時はコンソール出力のみ）
    --no-partial      部分一致を除外
    --verbose, -v     各ファイルの詳細プロパティを表示

Dependencies:
    pip install mutagen   # MP3 音声情報・ID3タグ取得（なくてもファイル情報のみ取得可）
"""

import argparse
import csv
import re
import sys
import unicodedata

# Windows コンソールを UTF-8 に設定（文字化け防止）
# GUI(--noconsole)の exe では sys.stdout/stderr が None になるため存在を確認する
for _stream in (sys.stdout, sys.stderr):
    if _stream is not None and getattr(_stream, "encoding", "") and \
            _stream.encoding.lower() in ("cp932", "shift_jis", "mbcs"):
        _stream.reconfigure(encoding="utf-8", errors="replace")
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

# mutagen はオプション依存
try:
    from mutagen.mp3 import MP3
    from mutagen.id3 import ID3, ID3NoHeaderError
    HAS_MUTAGEN = True
except ImportError:
    HAS_MUTAGEN = False


# =====================================================================
# データクラス
# =====================================================================

@dataclass
class MatchResult:
    """1件のマッチ結果"""
    event_name:     str
    match_type:     str
    mp3_path:       Path | None = None
    # ファイル情報
    file_size_bytes: int  = 0
    modified_at:    str  = ""
    # 音声情報（mutagen がある場合のみ）
    duration_sec:   float = 0.0
    bitrate_kbps:   int  = 0
    sample_rate_hz: int  = 0
    channels:       int  = 0
    # ID3 タグ
    tag_title:      str  = ""
    tag_artist:     str  = ""
    tag_album:      str  = ""
    tag_composer:   str  = ""
    tag_genre:      str  = ""
    tag_year:       str  = ""
    tag_track:      str  = ""
    tag_comment:    str  = ""


# =====================================================================
# NUENDO CSV 読み込み（マルチセクション形式対応）
# =====================================================================

def _decode(raw: bytes) -> tuple[str, str]:
    """複数エンコーディングを順番に試して最初に成功したものを返す"""
    for enc in ["utf-8-sig", "utf-8", "cp932", "shift_jis", "euc-jp", "latin-1"]:
        try:
            return raw.decode(enc), enc
        except (UnicodeDecodeError, LookupError):
            continue
    raise ValueError("対応できる文字コードが見つかりませんでした")


def _detect_sep(text: str) -> str:
    """先頭15行の区切り文字（カンマ・タブ・セミコロン）を推定する"""
    head = "\n".join(text.splitlines()[:15])
    counts = {",": head.count(","), "\t": head.count("\t"), ";": head.count(";")}
    return max(counts, key=counts.get)


def _is_cue_header(line: str) -> bool:
    """イベント名とファイル名が両方ある行 = Cue データのヘッダー行"""
    return "イベント名" in line and "ファイル名" in line


def read_event_names(csv_path: Path, verbose: bool = False) -> list[str]:
    """
    NUENDO Cue CSV を読み込み、イベント名の一覧（順序保持・重複除去）を返す。
    NUENDO の特殊なマルチセクション形式（トラックごとにヘッダーが繰り返される）に対応。
    """
    raw = csv_path.read_bytes()
    text, enc = _decode(raw)
    sep = _detect_sep(text)

    if verbose:
        print(f"   文字コード: {enc} / 区切り: {repr(sep)}")

    lines = text.splitlines()
    event_names: list[str] = []
    seen: set[str] = set()
    headers: list[str] | None = None
    n_expected: int = 0
    event_col: int = 0

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # ヘッダー行検出
        if _is_cue_header(stripped):
            headers = [h.strip() for h in stripped.split(sep)]
            n_expected = len(headers)
            event_col = next((i for i, h in enumerate(headers) if "イベント名" in h), 0)
            continue

        # データ行
        if headers and sep in stripped:
            fields = [f.strip() for f in stripped.split(sep, n_expected - 1)]
            if len(fields) > event_col:
                name = fields[event_col]
                if name and name not in seen:
                    event_names.append(name)
                    seen.add(name)

    return event_names


# =====================================================================
# 正規化・管理番号抽出
# =====================================================================

_LIB_RE = re.compile(
    r"([1-7]ST|[1-5]AN|[1-2]VO|[1-2]VJ)-\d{3}-\d{2}",
    re.IGNORECASE,
)
_AUDIOSTOCK_RE = re.compile(r"audiostock_\d+", re.IGNORECASE)


def _normalize(text: str) -> str:
    """照合比較用の正規化（全角半角統一・小文字・記号除去）"""
    t = unicodedata.normalize("NFKC", text)
    t = t.lower()
    t = re.sub(r"\.(mp3|wav|aiff?|flac|m4a)$", "", t, flags=re.IGNORECASE)
    t = re.sub(r"[_\-]+", " ", t)
    t = re.sub(r"[\(\)\[\]【】「」『』（）]", " ", t)
    t = re.sub(r"[!！?？。、・…～〜♪]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _lib_number(text: str) -> str:
    """ライブラリ管理番号を抽出する（例: '7ST-042-22'）"""
    m = _LIB_RE.search(text)
    return m.group(0).upper() if m else ""


def _clean_title(text: str) -> str:
    """管理番号・トラック番号プレフィックスを除去した曲名部分を返す"""
    t = _LIB_RE.sub("", text)
    t = _AUDIOSTOCK_RE.sub("", t)
    t = re.sub(r"^\d{1,2}-\d{1,2}\s+", "", t)   # "1-01 " 形式
    t = re.sub(r"^\d{1,3}\s+", "", t)             # "01 " 形式
    t = re.sub(r"-\d{2}$", "", t)                 # 末尾 "-01"
    t = re.sub(r"_\d+_[A-Za-z]+$", "", t)         # 末尾 "_113_C"
    t = re.sub(r"_[A-Za-z]{1,2}$", "", t)         # 末尾 "_C"
    return re.sub(r"^[\s\-_]+|[\s\-_]+$", "", t).strip()


# =====================================================================
# MP3 スキャン
# =====================================================================

# スキャン除外パターン（ファイル名のステム部分に対して適用）
_EXCLUDE_PATTERNS = [
    # テスト信号: 1KHZ-20DB, 1khz-20db-01 など
    re.compile(r"^\d+khz[-_]?\d+db", re.IGNORECASE),
    # 1〜2桁の数字のみ（スペース含む）: "1", "01", "1 ", "20 " など
    re.compile(r"^\d{1,2}\s*$"),
]


def _is_excluded(stem: str) -> bool:
    """除外すべきファイル名（ステム）かどうか判定する"""
    return any(pat.match(stem) for pat in _EXCLUDE_PATTERNS)


def scan_mp3_files(folder: Path, on_progress=None) -> list[Path]:
    """
    フォルダ以下の MP3 ファイルを再帰的に収集する（除外フィルター・進捗表示付き）。

    on_progress: 50件ごとに found 件数を渡すコールバック。
                 省略時はコンソールに上書き表示する（GUI からは差し替える）。
    """
    found: set[Path] = set()
    skipped = 0
    for ext in ("*.mp3", "*.MP3"):
        for f in folder.rglob(ext):
            if _is_excluded(f.stem):
                skipped += 1
                continue
            found.add(f)
            if len(found) % 50 == 0:
                if on_progress:
                    on_progress(len(found))
                else:
                    print(f"\r   スキャン中... {len(found)} 件発見", end="", flush=True)
    if on_progress:
        on_progress(len(found))
    else:
        print(f"\r   MP3: {len(found)} 件  (除外: {skipped} 件)                    ")
    return sorted(found)


# =====================================================================
# マッチング
# =====================================================================

MATCH_EXACT      = "完全一致"
MATCH_NORMALIZED = "正規化一致"
MATCH_NUMBER     = "管理番号一致"
MATCH_TITLE      = "タイトル一致"
MATCH_PARTIAL    = "部分一致"

# 優先度（数値が小さいほど優先）
_PRIORITY = {
    MATCH_EXACT: 0,
    MATCH_NORMALIZED: 1,
    MATCH_NUMBER: 2,
    MATCH_TITLE: 3,
    MATCH_PARTIAL: 4,
}


def match_event(
    event_name: str,
    mp3_files: list[Path],
    allow_partial: bool = True,
) -> list[tuple[Path, str]]:
    """
    1つのイベント名に対してマッチする MP3 ファイルとマッチ種別を返す。
    同一ファイルに複数マッチした場合は最高優先度のものだけを残す。
    """
    event_norm         = _normalize(event_name)
    event_lib          = _lib_number(event_name)
    event_title_norm   = _normalize(_clean_title(event_name))

    # ファイルごとに最高優先度のマッチを記録
    best: dict[Path, tuple[int, str]] = {}

    for mp3 in mp3_files:
        stem = mp3.stem

        def _update(mtype: str) -> None:
            pri = _PRIORITY[mtype]
            if mp3 not in best or pri < best[mp3][0]:
                best[mp3] = (pri, mtype)

        # 完全一致
        if stem.lower() == event_name.lower():
            _update(MATCH_EXACT)
            continue  # 完全一致があればそれ以上は不要

        # 正規化一致
        stem_norm = _normalize(stem)
        if event_norm and stem_norm == event_norm:
            _update(MATCH_NORMALIZED)
            continue

        # 管理番号一致
        if event_lib and _lib_number(stem) == event_lib:
            _update(MATCH_NUMBER)
            continue

        # タイトル一致
        stem_title_norm = _normalize(_clean_title(stem))
        if event_title_norm and stem_title_norm and stem_title_norm == event_title_norm:
            _update(MATCH_TITLE)
            continue

        # 部分一致（オプション）
        if allow_partial and event_norm and stem_norm:
            if event_norm in stem_norm or stem_norm in event_norm:
                _update(MATCH_PARTIAL)

    return [(mp3, mtype) for mp3, (_, mtype) in best.items()]


# =====================================================================
# MP3 プロパティ取得
# =====================================================================

def read_properties(mp3_path: Path, event_name: str, match_type: str) -> MatchResult:
    """MP3 ファイルのプロパティを読み取って MatchResult を返す（ファイルは変更しない）"""
    stat = mp3_path.stat()
    r = MatchResult(
        event_name      = event_name,
        match_type      = match_type,
        mp3_path        = mp3_path,
        file_size_bytes = stat.st_size,
        modified_at     = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
    )

    if not HAS_MUTAGEN:
        return r

    # 音声情報
    try:
        audio = MP3(str(mp3_path))
        r.duration_sec   = audio.info.length
        r.bitrate_kbps   = audio.info.bitrate // 1000
        r.sample_rate_hz = audio.info.sample_rate
        r.channels       = getattr(audio.info, "channels", 0)
    except Exception:
        pass

    # ID3 タグ
    try:
        tags = ID3(str(mp3_path))

        def _tag(key: str) -> str:
            f = tags.get(key)
            return str(f.text[0]) if f and hasattr(f, "text") and f.text else ""

        r.tag_title    = _tag("TIT2")
        r.tag_artist   = _tag("TPE1")
        r.tag_album    = _tag("TALB")
        r.tag_composer = _tag("TCOM")
        r.tag_genre    = _tag("TCON")
        r.tag_year     = _tag("TDRC") or _tag("TYER")
        r.tag_track    = _tag("TRCK")
        # コメントは COMM フレームを検索
        comments = [str(v.text[0]) for k, v in tags.items()
                    if k.startswith("COMM") and hasattr(v, "text") and v.text]
        r.tag_comment = comments[0] if comments else ""

    except (ID3NoHeaderError, Exception):
        pass

    return r


# =====================================================================
# コンソール出力
# =====================================================================

def _fmt_duration(sec: float) -> str:
    if sec <= 0:
        return ""
    h, rem = divmod(int(sec), 3600)
    m, s = divmod(rem, 60)
    frac = sec - int(sec)
    return f"{h:02d}:{m:02d}:{s:02d}.{int(frac*100):02d}"


def _fmt_size(b: int) -> str:
    if b >= 1024 ** 2:
        return f"{b/1024**2:.1f} MB"
    if b >= 1024:
        return f"{b/1024:.1f} KB"
    return f"{b} B"


def print_summary_table(results: list[MatchResult]) -> None:
    """一覧テーブルをコンソールに表示する"""
    if not results:
        print("  (マッチなし)")
        return

    EW, TW, FW = 40, 8, 45
    header = f"{'No':>3}  {'種別':{TW}}  {'イベント名':{EW}}  {'MP3ファイル名':{FW}}"
    print(header)
    print("-" * len(header))
    for i, r in enumerate(results, 1):
        ename = r.event_name[:EW]
        fname = r.mp3_path.name[:FW] if r.mp3_path else "-"
        print(f"{i:3}  {r.match_type:{TW}}  {ename:{EW}}  {fname:{FW}}")


def print_detail(r: MatchResult) -> None:
    """1件のプロパティ詳細を表示する"""
    print(f"\n  [{r.event_name}]")
    print(f"    種別     : {r.match_type}")
    if r.mp3_path is None:
        return
    print(f"    ファイル : {r.mp3_path.name}")
    print(f"    パス     : {r.mp3_path}")
    print(f"    サイズ   : {_fmt_size(r.file_size_bytes)}")
    print(f"    更新日時 : {r.modified_at}")
    if r.duration_sec > 0:
        print(f"    再生時間 : {_fmt_duration(r.duration_sec)}")
        print(f"    ビットレート: {r.bitrate_kbps} kbps")
        print(f"    サンプルレート: {r.sample_rate_hz} Hz  チャンネル: {r.channels}")
    tag_items = [
        ("タイトル",     r.tag_title),
        ("アーティスト", r.tag_artist),
        ("アルバム",     r.tag_album),
        ("作曲者",       r.tag_composer),
        ("ジャンル",     r.tag_genre),
        ("年",           r.tag_year),
        ("トラック",     r.tag_track),
        ("コメント",     r.tag_comment),
    ]
    tag_lines = [(k, v) for k, v in tag_items if v]
    if tag_lines:
        print("    [ID3タグ]")
        for k, v in tag_lines:
            print(f"      {k:10}: {v}")


# =====================================================================
# CSV 出力
# =====================================================================

def export_csv(results: list[MatchResult], output_path: Path, log=print) -> None:
    """照合結果を CSV ファイルに書き出す（log は GUI から差し替える）"""
    if not results:
        log("出力するデータがありません（マッチ件数が 0 件）。")
        return

    # フォルダパスだけ渡された場合はファイル名を自動付与
    if output_path.is_dir():
        output_path = output_path / "mp3_finder_result.csv"

    # 親フォルダが存在しない場合は作成
    output_path.parent.mkdir(parents=True, exist_ok=True)

    log(f"\nCSV 書き出し先: {output_path}")

    fieldnames = [
        "イベント名", "マッチ種別", "ファイル名", "フルパス",
        "ファイルサイズ(bytes)", "ファイルサイズ", "更新日時",
        "再生時間", "ビットレート(kbps)", "サンプルレート(Hz)", "チャンネル数",
        "タイトル(ID3)", "アーティスト(ID3)", "アルバム(ID3)", "作曲者(ID3)",
        "ジャンル(ID3)", "年(ID3)", "トラック(ID3)", "コメント(ID3)",
    ]

    try:
        with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for r in results:
                writer.writerow({
                    "イベント名":           r.event_name,
                    "マッチ種別":           r.match_type,
                    "ファイル名":           r.mp3_path.name if r.mp3_path else "",
                    "フルパス":             str(r.mp3_path) if r.mp3_path else "",
                    "ファイルサイズ(bytes)": r.file_size_bytes if r.mp3_path else "",
                    "ファイルサイズ":        _fmt_size(r.file_size_bytes) if r.mp3_path else "",
                    "更新日時":             r.modified_at,
                    "再生時間":             _fmt_duration(r.duration_sec) if r.mp3_path else "",
                    "ビットレート(kbps)":   r.bitrate_kbps if r.mp3_path else "",
                    "サンプルレート(Hz)":   r.sample_rate_hz if r.mp3_path else "",
                    "チャンネル数":         r.channels if r.mp3_path else "",
                    "タイトル(ID3)":        r.tag_title,
                    "アーティスト(ID3)":    r.tag_artist,
                    "アルバム(ID3)":        r.tag_album,
                    "作曲者(ID3)":          r.tag_composer,
                    "ジャンル(ID3)":        r.tag_genre,
                    "年(ID3)":              r.tag_year,
                    "トラック(ID3)":        r.tag_track,
                    "コメント(ID3)":        r.tag_comment,
                })
        log(f"[OK] CSV 出力完了: {output_path}  ({len(results)} 件)")
    except Exception as e:
        log(f"[ERROR] CSV 書き出し失敗: {e}")
        log(f"        パスを確認してください: {output_path}")


# =====================================================================
# エントリポイント
# =====================================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description="NUENDO Cue CSV のイベント名と一致する MP3 を収集してプロパティを表示する",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("csv_file",    help="NUENDO Cue CSV ファイルパス")
    parser.add_argument("folder",      help="MP3 を検索するフォルダパス")
    parser.add_argument("--workers",   type=int, default=8, metavar="N",
                        help="並列スレッド数 (default: 8)")
    parser.add_argument("--output",    metavar="FILE",
                        help="結果を CSV に出力するファイルパス")
    parser.add_argument("--no-partial", action="store_true",
                        help="部分一致を除外する")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="各ファイルのプロパティ詳細を表示する")
    args = parser.parse_args()

    csv_path    = Path(args.csv_file)
    folder      = Path(args.folder)
    allow_partial = not args.no_partial

    # 入力チェック
    if not csv_path.exists():
        print(f"エラー: CSV が見つかりません: {csv_path}", file=sys.stderr)
        sys.exit(1)
    if not folder.is_dir():
        print(f"エラー: フォルダが見つかりません: {folder}", file=sys.stderr)
        sys.exit(1)

    if not HAS_MUTAGEN:
        print("[WARN]  mutagen 未インストール: 音声情報・ID3タグは取得されません")
        print("   pip install mutagen でインストールできます\n")

    # ① イベント名抽出
    print(f"[CSV] CSV 読み込み: {csv_path.name}")
    event_names = read_event_names(csv_path, verbose=args.verbose)
    if not event_names:
        print("エラー: イベント名を取得できませんでした。CSV 形式を確認してください。",
              file=sys.stderr)
        sys.exit(1)
    print(f"   イベント名: {len(event_names)} 件\n")

    # ② MP3 スキャン
    print(f"[SCAN] MP3 スキャン: {folder}")
    mp3_files = scan_mp3_files(folder)
    print(f"   MP3: {len(mp3_files)} 件\n")

    if not mp3_files:
        print("   MP3 ファイルが見つかりません。全イベントを「該当なし」として出力します。\n")
        results: list[MatchResult] = [
            MatchResult(event_name=name, match_type="該当なし")
            for name in event_names
        ]
        print_summary_table(results)
        if args.output:
            export_csv(results, Path(args.output))
        else:
            print("\n[TIP] --output result.csv を指定すると CSV に出力できます")
        sys.exit(0)

    # ③ 並列マッチング（イベント名ごとに並列）
    print(f"[MATCH] マッチング中 ({args.workers} スレッド) ...")
    raw_matches: list[tuple[str, Path, str]] = []

    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futures = {
            ex.submit(match_event, name, mp3_files, allow_partial): name
            for name in event_names
        }
        for future in as_completed(futures):
            event_name = futures[future]
            for mp3_path, match_type in future.result():
                raw_matches.append((event_name, mp3_path, match_type))

    # 元の順序に並び替え
    order = {name: i for i, name in enumerate(event_names)}
    raw_matches.sort(key=lambda x: (order.get(x[0], 9999), x[1]))

    matched_events = len({m[0] for m in raw_matches})
    unmatched_events = len(event_names) - matched_events
    print(f"   マッチ: {matched_events} / {len(event_names)} イベント  "
          f"({len(raw_matches)} ファイル)  該当なし: {unmatched_events} 件\n")

    # ④ 並列プロパティ取得（マッチしたファイルごとに並列）
    print(f"[PROP] プロパティ取得中 ({args.workers} スレッド) ...")
    results: list[MatchResult] = []

    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futures2 = {
            ex.submit(read_properties, mp3, ev, mt): (ev, str(mp3))
            for ev, mp3, mt in raw_matches
        }
        for future in as_completed(futures2):
            results.append(future.result())

    # 順序を元に戻す
    match_order = {(ev, str(mp3)): i for i, (ev, mp3, _) in enumerate(raw_matches)}
    results.sort(key=lambda r: match_order.get((r.event_name, str(r.mp3_path) if r.mp3_path else ""), 9999))

    # 該当なし：マッチしなかったイベント名を末尾に追加（元順序を保持）
    matched_event_names = {r.event_name for r in results}
    for name in event_names:
        if name not in matched_event_names:
            results.append(MatchResult(event_name=name, match_type="該当なし"))

    # ⑤ コンソール出力
    print()
    print("=" * 80)
    print("  照合結果")
    print("=" * 80)
    print_summary_table(results)

    if args.verbose and results:
        print()
        print("=" * 80)
        print("  プロパティ詳細")
        print("=" * 80)
        for r in results:
            print_detail(r)

    # ⑥ サマリー
    unmatched_results = [r for r in results if r.match_type == "該当なし"]
    type_counts: dict[str, int] = {}
    for r in results:
        type_counts[r.match_type] = type_counts.get(r.match_type, 0) + 1

    print()
    print("=" * 80)
    print("  サマリー")
    print("=" * 80)
    print(f"  照合対象 : {len(event_names)} 件")
    matched_counts = {k: v for k, v in type_counts.items() if k != "該当なし"}
    print(f"  照合成功 : {matched_events} 件  "
          + "  ".join(f"{k}:{v}" for k, v in matched_counts.items()))
    print(f"  該当なし : {len(unmatched_results)} 件")
    if unmatched_results:
        print("\n  該当なしのイベント名:")
        for r in unmatched_results:
            print(f"    - {r.event_name}")

    # ⑦ CSV 出力（--output 指定時）
    if args.output:
        export_csv(results, Path(args.output))
    else:
        print("\n[TIP] --output result.csv を指定すると CSV に出力できます")


if __name__ == "__main__":
    main()
