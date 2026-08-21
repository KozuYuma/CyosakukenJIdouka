#!/usr/bin/env python3
"""
mp3_finder_gui.py

nuendo_mp3_finder のGUI版。非エンジニアが exe をダブルクリックして使えるよう、
「Cue CSV を選ぶ」「MP3フォルダを選ぶ」「実行」だけの画面にしてある。

照合ロジックは nuendo_mp3_finder をそのまま呼ぶ（重複実装しない）。

単体起動:
    python scripts\\mp3_finder_gui.py
exe 化:
    scripts\\build_mp3_finder_exe.ps1
"""

import queue
import subprocess
import sys
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# PyInstaller の onefile 展開先でも scripts/ を import できるようにする
sys.path.insert(0, str(Path(__file__).resolve().parent))

from nuendo_mp3_finder import (  # noqa: E402
    HAS_MUTAGEN,
    MatchResult,
    export_csv,
    match_event,
    read_event_names,
    read_properties,
    scan_mp3_files,
)

APP_TITLE = "NUENDO MP3 Finder"
WORKERS = 8


# =====================================================================
# 処理本体（ワーカースレッドで実行。UI には触らず log() だけ呼ぶ）
# =====================================================================

def run_finder(csv_path: Path, mp3_dir: Path, out_path: Path,
               allow_partial: bool, log) -> Path | None:
    """照合を実行して CSV に書き出す。戻り値は出力した CSV のパス。"""
    if not HAS_MUTAGEN:
        log("[警告] mutagen が無いため、再生時間とID3タグは取得されません。")

    # ① イベント名
    log(f"[1/4] CSV 読み込み: {csv_path.name}")
    event_names = read_event_names(csv_path)
    if not event_names:
        raise ValueError(
            "CSV からイベント名を取得できませんでした。\n"
            "NUENDO の Cue CSV（「イベント名」「ファイル名」列があるもの）か確認してください。"
        )
    log(f"      イベント名 {len(event_names)} 件")

    # ② MP3 スキャン
    log(f"[2/4] MP3 スキャン: {mp3_dir}")
    mp3_files = scan_mp3_files(
        mp3_dir, on_progress=lambda n: log(f"      スキャン中... {n} 件", replace=True)
    )
    log(f"      MP3 {len(mp3_files)} 件", replace=True)

    if not mp3_files:
        log("      MP3 が1件も見つかりません。全件「該当なし」として出力します。")
        results = [MatchResult(event_name=n, match_type="該当なし") for n in event_names]
        export_csv(results, out_path, log=log)
        return out_path

    # ③ マッチング
    log("[3/4] マッチング中 ...")
    raw: list[tuple[str, Path, str]] = []
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futures = {ex.submit(match_event, n, mp3_files, allow_partial): n
                   for n in event_names}
        for done, fut in enumerate(as_completed(futures), 1):
            for mp3_path, mtype in fut.result():
                raw.append((futures[fut], mp3_path, mtype))
            log(f"      {done}/{len(event_names)} 件", replace=True)

    order = {n: i for i, n in enumerate(event_names)}
    raw.sort(key=lambda x: (order.get(x[0], 9999), x[1]))
    matched = len({m[0] for m in raw})
    log(f"      マッチ {matched}/{len(event_names)} イベント "
        f"（{len(raw)} ファイル）", replace=True)

    # ④ プロパティ取得
    log("[4/4] プロパティ取得中 ...")
    results: list[MatchResult] = []
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futures2 = {ex.submit(read_properties, mp3, ev, mt): (ev, str(mp3))
                    for ev, mp3, mt in raw}
        for done, fut in enumerate(as_completed(futures2), 1):
            results.append(fut.result())
            log(f"      {done}/{len(futures2)} 件", replace=True)

    match_order = {(ev, str(mp3)): i for i, (ev, mp3, _) in enumerate(raw)}
    results.sort(key=lambda r: match_order.get(
        (r.event_name, str(r.mp3_path) if r.mp3_path else ""), 9999))

    # 該当なしを末尾に追加（元順序を保持）
    hit_names = {r.event_name for r in results}
    unmatched = [n for n in event_names if n not in hit_names]
    results.extend(MatchResult(event_name=n, match_type="該当なし") for n in unmatched)

    log(f"\n照合対象 {len(event_names)} 件 ／ 照合成功 {matched} 件 ／ "
        f"該当なし {len(unmatched)} 件")
    if unmatched:
        log("\n該当なしのイベント名:")
        for n in unmatched[:50]:
            log(f"  - {n}")
        if len(unmatched) > 50:
            log(f"  ... 他 {len(unmatched) - 50} 件")

    export_csv(results, out_path, log=log)
    return out_path


# =====================================================================
# GUI
# =====================================================================

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(720, 460)

        self.var_csv     = tk.StringVar()
        self.var_dir     = tk.StringVar()
        self.var_out     = tk.StringVar()
        self.var_partial = tk.BooleanVar(value=True)
        self.var_status  = tk.StringVar(value="Cue CSV と MP3 フォルダを選んで「実行」を押してください")

        # ワーカースレッド → UI へのメッセージ受け渡し
        self._q: queue.Queue = queue.Queue()
        self._running = False
        self._last_out: Path | None = None
        self._prog_idx: str | None = None   # 上書き対象の進捗行の開始位置

        self._build()
        self.after(100, self._drain_queue)

    # ---- 画面構築 -----------------------------------------------------
    def _build(self) -> None:
        pad = dict(padx=10, pady=6)
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        self._row(frm, 0, "Cue CSV", self.var_csv, self._pick_csv)
        self._row(frm, 1, "MP3 フォルダ", self.var_dir, self._pick_dir)
        self._row(frm, 2, "出力先 CSV", self.var_out, self._pick_out)

        ttk.Checkbutton(frm, text="部分一致も含める（曲名の一部が一致するファイルも拾う）",
                        variable=self.var_partial).grid(
            row=3, column=1, columnspan=2, sticky="w", **pad)

        self.btn_run = ttk.Button(frm, text="実　行", command=self._on_run)
        self.btn_run.grid(row=4, column=1, sticky="ew", **pad)
        self.btn_open = ttk.Button(frm, text="出力先を開く", command=self._open_out,
                                   state="disabled")
        self.btn_open.grid(row=4, column=2, sticky="ew", **pad)

        self.prog = ttk.Progressbar(frm, mode="determinate", maximum=100)
        self.prog.grid(row=5, column=0, columnspan=3, sticky="ew", **pad)

        ttk.Label(frm, textvariable=self.var_status, foreground="#555").grid(
            row=6, column=0, columnspan=3, sticky="w", padx=10)

        frm.rowconfigure(7, weight=1)
        box = ttk.Frame(frm)
        box.grid(row=7, column=0, columnspan=3, sticky="nsew", padx=10, pady=(4, 10))
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)
        self.log = tk.Text(box, height=14, wrap="none", state="disabled",
                           font=("Consolas", 9))
        self.log.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.log.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=sb.set)

    def _row(self, parent, r: int, label: str, var: tk.StringVar, cmd) -> None:
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(parent, textvariable=var).grid(row=r, column=1, sticky="ew", pady=6)
        ttk.Button(parent, text="参照...", command=cmd).grid(
            row=r, column=2, sticky="ew", padx=10, pady=6)

    # ---- ファイル選択 -------------------------------------------------
    def _pick_csv(self) -> None:
        p = filedialog.askopenfilename(
            title="NUENDO Cue CSV を選択",
            filetypes=[("CSV ファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )
        if not p:
            return
        self.var_csv.set(p)
        # 出力先が空なら CSV と同じ場所に自動で決めておく（利用者に考えさせない）
        if not self.var_out.get():
            src = Path(p)
            self.var_out.set(str(src.with_name(f"{src.stem}_mp3情報.csv")))

    def _pick_dir(self) -> None:
        p = filedialog.askdirectory(title="MP3 が入っているフォルダを選択")
        if p:
            self.var_dir.set(p)

    def _pick_out(self) -> None:
        p = filedialog.asksaveasfilename(
            title="出力先 CSV", defaultextension=".csv",
            initialfile=Path(self.var_out.get()).name if self.var_out.get() else "mp3_finder_result.csv",
            filetypes=[("CSV ファイル", "*.csv")],
        )
        if p:
            self.var_out.set(p)

    def _open_out(self) -> None:
        if self._last_out and self._last_out.exists():
            # 出力した CSV をエクスプローラーで選択状態にして開く。
            # /select, はパスと連結して1引数で渡す必要があるため文字列で起動する。
            subprocess.Popen(f'explorer /select,"{self._last_out}"')

    # ---- 実行 ---------------------------------------------------------
    def _on_run(self) -> None:
        if self._running:
            return
        csv_path = Path(self.var_csv.get().strip('" '))
        mp3_dir  = Path(self.var_dir.get().strip('" '))
        out_txt  = self.var_out.get().strip('" ')

        if not csv_path.is_file():
            messagebox.showerror(APP_TITLE, "Cue CSV が見つかりません。\nもう一度選び直してください。")
            return
        if not mp3_dir.is_dir():
            messagebox.showerror(APP_TITLE, "MP3 フォルダが見つかりません。\nもう一度選び直してください。")
            return
        if not out_txt:
            out_txt = str(csv_path.with_name(f"{csv_path.stem}_mp3情報.csv"))
            self.var_out.set(out_txt)

        self._running = True
        self._last_out = None
        self.btn_run.configure(state="disabled", text="実行中...")
        self.btn_open.configure(state="disabled")
        self.prog.configure(mode="indeterminate")
        self.prog.start(12)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        threading.Thread(
            target=self._worker,
            args=(csv_path, mp3_dir, Path(out_txt), self.var_partial.get()),
            daemon=True,
        ).start()

    def _worker(self, csv_path: Path, mp3_dir: Path, out_path: Path,
                allow_partial: bool) -> None:
        def log(msg: str, replace: bool = False) -> None:
            self._q.put(("log", (str(msg), replace)))
        try:
            result = run_finder(csv_path, mp3_dir, out_path, allow_partial, log)
            self._q.put(("done", result))
        except Exception as e:
            self._q.put(("error", (f"{type(e).__name__}: {e}", traceback.format_exc())))

    # ---- ワーカーからのメッセージを UI に反映 --------------------------
    def _drain_queue(self) -> None:
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "log":
                    self._append(*payload)
                elif kind == "done":
                    self._finish(payload)
                elif kind == "error":
                    self._fail(*payload)
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def _append(self, msg: str, replace: bool) -> None:
        self.log.configure(state="normal")
        # replace=True は進捗行。直前の進捗行を消して上書きし、ログが流れないようにする
        if replace and self._prog_idx:
            self.log.delete(self._prog_idx, "end-1c")
        start = self.log.index("end-1c")
        self.log.insert("end", msg + "\n")
        self._prog_idx = start if replace else None
        self.log.see("end")
        self.log.configure(state="disabled")
        self.var_status.set(msg.strip().splitlines()[0][:90] if msg.strip() else "")

    def _reset(self) -> None:
        self._running = False
        self.prog.stop()
        self.prog.configure(mode="determinate", value=0)
        self.btn_run.configure(state="normal", text="実　行")

    def _finish(self, out: Path | None) -> None:
        self._reset()
        self._last_out = out
        if out and out.exists():
            self.btn_open.configure(state="normal")
            self.var_status.set(f"完了: {out}")
            messagebox.showinfo(APP_TITLE, f"完了しました。\n\n出力先:\n{out}")
        else:
            self.var_status.set("完了しましたが、CSV は出力されませんでした")

    def _fail(self, short: str, detail: str) -> None:
        self._reset()
        self._append(f"\n[エラー] {short}", False)
        self._append(detail, False)
        self.var_status.set(f"エラー: {short}")
        messagebox.showerror(APP_TITLE, f"処理に失敗しました。\n\n{short}")


def main() -> None:
    # 引数なし（＝ダブルクリック）なら GUI。
    # --run は画面を出さずに実行する動作確認・バッチ用の隠しオプション。
    argv = sys.argv[1:]
    if argv and argv[0] == "--run":
        if len(argv) < 4:
            print("usage: --run <cue.csv> <mp3_folder> <out.csv> [--no-partial]")
            raise SystemExit(2)
        out = Path(argv[3])
        lines: list[str] = []

        def log(msg: str, replace: bool = False) -> None:
            lines.append(str(msg))
            print(msg)   # windowed exe では stdout が無いので黙って捨てられる

        try:
            run_finder(Path(argv[1]), Path(argv[2]), out,
                       "--no-partial" not in argv, log)
        finally:
            # console=False の exe では標準出力が無いため、ログはファイルに残す
            out.with_suffix(".log").write_text("\n".join(lines), encoding="utf-8")
        return
    App().mainloop()


if __name__ == "__main__":
    main()
