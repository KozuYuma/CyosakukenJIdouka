"""
Excel 出力モジュール
openpyxl で整形済み Excel ファイルを生成する
"""
import io

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --- 確認ステータスの色定義（背景色 RGB 16進） ---
STATUS_COLOR_MAP: dict[str, str] = {
    "未調査":           "FFFFFF",  # 白
    "候補あり":         "FFF9C4",  # 薄黄
    "確定":             "C8E6C9",  # 薄緑
    "要確認":           "FFCDD2",  # 薄赤
    "J-WID要確認":      "FFE0B2",  # 薄オレンジ
    "NexTone要確認":    "F3E5F5",  # 薄紫
    "ライブラリ元確認": "E1F5FE",  # 薄水色
    "Audiostock確認":  "F1F8E9",  # 薄黄緑
    "MP3補助確認":      "FFF3E0",  # 薄アンバー
}

HEADER_FILL = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
DEFAULT_FONT = Font(size=10)

# Excel 文字列強制列（数値として解釈させない列名）
STRING_FORCE_COLS = {
    "JASRAC作品コード",
    "NexTone管理番号",
    "CD番号",
    "元管理番号",
    "ライブラリ盤番号",
    "トラック番号",
    "Audiostock管理番号",
}


def _write_sheet(ws, df: pd.DataFrame) -> None:
    """DataFrame を 1 ワークシートに書き込む（ヘッダー固定・フィルター・列幅調整つき）"""
    if df is None or len(df) == 0:
        ws.cell(row=1, column=1, value="データなし")
        return

    col_names = list(df.columns)

    # ---- ヘッダー行（1行目） ----
    for col_idx, col_name in enumerate(col_names, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ---- データ行 ----
    status_col_idx = col_names.index("確認ステータス") + 1 if "確認ステータス" in col_names else None

    for row_idx, (_, series) in enumerate(df.iterrows(), 2):
        for col_idx, col_name in enumerate(col_names, 1):
            val = series[col_name]

            # NaN・None を空文字に
            if pd.isna(val):
                val = ""

            # 文字列強制列はアポストロフィを付けて文字列として保存
            if col_name in STRING_FORCE_COLS:
                val = str(val) if val != "" else ""

            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = DEFAULT_FONT

        # 確認ステータス列の背景色
        if status_col_idx:
            status_val = str(series.get("確認ステータス", ""))
            color = STATUS_COLOR_MAP.get(status_val, "FFFFFF")
            ws.cell(row=row_idx, column=status_col_idx).fill = PatternFill(
                start_color=color, end_color=color, fill_type="solid"
            )

    # ---- フィルター & ウィンドウ固定 ----
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    # ---- 列幅自動調整（最大 50 文字） ----
    for col_idx, col_name in enumerate(col_names, 1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col_name))
        sample_rows = min(ws.max_row, 200)  # サンプル 200 行で計算
        for r in range(2, sample_rows + 1):
            val = ws.cell(row=r, column=col_idx).value
            if val:
                max_len = max(max_len, min(len(str(val)), 50))
        ws.column_dimensions[col_letter].width = max_len + 2


def _write_notes_sheet(ws) -> None:
    """処理ルール・凡例を確認メモシートに書き込む"""
    notes = [
        ("【処理ルール】", ""),
        ("", "・NUENDOの使用尺は編集後の使用尺です。曲のフル尺ではありません。"),
        ("", "・曲特定にはプロジェクト Audio フォルダ内の WAV フル尺を使います。"),
        ("", "・WAV で特定できない場合のみ MP3 を補助として使います。"),
        ("", "・MP3 タグ情報は前提にしません（タグなしファイルにも対応）。"),
        ("", "・NUENDO のファイル名より、イベント名を原曲特定の優先キーとします。"),
        ("", ""),
        ("【照合優先順位】", ""),
        ("", "1. イベント名と WAV ファイル名の完全一致"),
        ("", "2. イベント名から拡張子・余分な記号を除いた正規化一致"),
        ("", "3. 管理番号一致"),
        ("", "4. 曲名一致（管理番号除去後）"),
        ("", "5. NUENDO ファイル名一致"),
        ("", "6. 部分一致"),
        ("", ""),
        ("【管理番号形式】", ""),
        ("", "ライブラリ管理番号: 数字1桁 + 英字2文字 - 3桁 - 2桁"),
        ("", "  例: 6ST-653-09 GO! GO!  →  管理番号: 6ST-653-09 / 盤番号: 6ST-653 / トラック: 09"),
        ("", "  有効系列: 1ST〜7ST / 1AN〜5AN / 1VO〜2VO / 1VJ〜2VJ"),
        ("", ""),
        ("", "Audiostock 管理番号: audiostock_数字"),
        ("", "  例: audiostock_856447_残念なシーンのジングル  →  番号: 856447"),
        ("", ""),
        ("【確認ステータス凡例】", ""),
        ("未調査",           "まだ調査していない"),
        ("候補あり",         "候補は特定できているが確認中"),
        ("確定",             "権利情報確定"),
        ("要確認",           "何らかの問題あり"),
        ("J-WID要確認",      "J-WIDでの確認が必要"),
        ("NexTone要確認",    "NexToneでの確認が必要"),
        ("ライブラリ元確認", "ライブラリ元への確認が必要"),
        ("Audiostock確認",  "Audiostockでの確認が必要"),
        ("MP3補助確認",      "WAVで照合できず、MP3での補助確認が必要"),
    ]

    bold_font = Font(bold=True, size=10)
    normal_font = Font(size=10)

    for row_idx, (col_a, col_b) in enumerate(notes, 1):
        cell_a = ws.cell(row=row_idx, column=1, value=col_a)
        ws.cell(row=row_idx, column=2, value=col_b)
        if col_a.startswith("【") or col_a in STATUS_COLOR_MAP:
            cell_a.font = bold_font
            if col_a in STATUS_COLOR_MAP:
                color = STATUS_COLOR_MAP[col_a]
                cell_a.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        else:
            cell_a.font = normal_font

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 65


def export_to_excel(
    songs_df: pd.DataFrame | None,
    events_df: pd.DataFrame | None,
    wav_df: pd.DataFrame | None,
    mp3_df: pd.DataFrame | None,
    search_df: pd.DataFrame | None,
) -> bytes:
    """
    各 DataFrame を Excel ファイルとして書き出す。
    Returns: Excel バイト列（st.download_button に直接渡せる）
    """
    wb = Workbook()
    wb.remove(wb.active)  # デフォルト空シートを削除

    _write_sheet(wb.create_sheet("楽曲まとめ"), songs_df)
    _write_sheet(wb.create_sheet("イベント一覧"), events_df)
    _write_sheet(wb.create_sheet("WAV一覧"), wav_df)
    _write_sheet(wb.create_sheet("MP3一覧"), mp3_df)
    _write_sheet(wb.create_sheet("検索語一覧"), search_df)
    _write_notes_sheet(wb.create_sheet("確認メモ"))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
