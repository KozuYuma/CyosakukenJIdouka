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
from tkinter import filedialog, font as tkfont, messagebox, ttk

# ドラッグ＆ドロップ（tkinterdnd2）。無くても参照ボタンで使えるので必須にはしない
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _BASE_TK = TkinterDnD.Tk
    HAS_DND = True
except Exception:
    DND_FILES = None
    _BASE_TK = tk.Tk
    HAS_DND = False

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


class Cancelled(Exception):
    """[中止]が押されたときに処理を抜けるための例外。エラー扱いにはしない。"""


def enable_dpi_awareness() -> None:
    """
    高DPI環境で文字が滲む（太って見える）のを防ぐ。

    DPI非対応のまま起動すると Windows が 96dpi で描いた画面を
    引き伸ばすため、全体がぼやけて太字のように見える。
    Tk のウィンドウを作る前に呼ぶ必要がある。
    """
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # System DPI aware
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()    # 旧Windows向け
        except Exception:
            pass


# =====================================================================
# 処理本体（ワーカースレッドで実行。UI には触らず log() だけ呼ぶ）
# =====================================================================

#: 各段階が進捗バーのどこからどこまでを受け持つか（％）。
#: 件数が分かる段階だけ実測で進める。合計が100になるようにしてある。
_PHASE = {
    "read":  (0, 5),     # CSV 読み込み
    "scan":  (5, 15),    # MP3 スキャン（総数が分からないので流し表示）
    "match": (15, 60),   # マッチング
    "prop":  (60, 95),   # プロパティ取得
    "write": (95, 100),  # CSV 書き出し
}


def run_finder(csv_path: Path, mp3_dir: Path, out_path: Path,
               allow_partial: bool, log, should_stop=None,
               on_progress=None) -> Path | None:
    """
    照合を実行して CSV に書き出す。戻り値は出力した CSV のパス。

    should_stop: 中止されたかを返す関数。要所で見て、真なら Cancelled を送出する。
                 スレッドを強制終了する手段は無いので、この方式で協調的に止める。
    on_progress: 進捗率(0-100)を渡すコールバック。総数が分からない段階では
                 None を渡す（受け手はバーを流し表示にする）。
    """
    def _check() -> None:
        if should_stop and should_stop():
            raise Cancelled

    def _prog(phase: str, done: int = 1, total: int = 1) -> None:
        """段階の中の進み具合を全体の％に直して渡す。"""
        if not on_progress:
            return
        lo, hi = _PHASE[phase]
        on_progress(lo + (hi - lo) * (done / total if total else 1))

    _prog("read", 0)

    if not HAS_MUTAGEN:
        log("[警告] mutagen が無いため、再生時間とID3タグは取得されません。")

    # ① イベント名
    log(f"[1/4] CSV 読み込み: {csv_path.name}")
    event_names = read_event_names(csv_path)
    _check()
    if not event_names:
        raise ValueError(
            "CSV からイベント名を取得できませんでした。\n"
            "NUENDO の Cue CSV（「イベント名」「ファイル名」列があるもの）か確認してください。"
        )
    log(f"      イベント名 {len(event_names)} 件")
    _prog("read")

    # ② MP3 スキャン
    log(f"[2/4] MP3 スキャン: {mp3_dir}")
    # 総数が事前に分からない段階なので、％ではなく流し表示にしてもらう
    if on_progress:
        on_progress(None)

    def _on_scan(n: int) -> None:
        # 大量のフォルダを掘っている最中でも中止できるよう、進捗の度に見る
        _check()
        log(f"      スキャン中... {n} 件", replace=True)

    mp3_files = scan_mp3_files(mp3_dir, on_progress=_on_scan)
    log(f"      MP3 {len(mp3_files)} 件", replace=True)
    _prog("scan")
    _check()

    if not mp3_files:
        log("      MP3 が1件も見つかりません。全件「該当なし」として出力します。")
        results = [MatchResult(event_name=n, match_type="該当なし") for n in event_names]
        export_csv(results, out_path, log=log)
        _prog("write")
        return out_path

    # ③ マッチング
    log("[3/4] マッチング中 ...")
    raw: list[tuple[str, Path, str]] = []
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futures = {ex.submit(match_event, n, mp3_files, allow_partial): n
                   for n in event_names}
        for done, fut in enumerate(as_completed(futures), 1):
            if should_stop and should_stop():
                # 未着手の分だけ取り消す。動き出している分は短いので待つ
                for f in futures:
                    f.cancel()
                raise Cancelled
            for mp3_path, mtype in fut.result():
                raw.append((futures[fut], mp3_path, mtype))
            log(f"      {done}/{len(event_names)} 件", replace=True)
            _prog("match", done, len(event_names))

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
            if should_stop and should_stop():
                for f in futures2:
                    f.cancel()
                raise Cancelled
            results.append(fut.result())
            log(f"      {done}/{len(futures2)} 件", replace=True)
            _prog("prop", done, len(futures2))

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

    _check()   # 中止直後に中途半端な CSV を書かないよう、書き出す直前にも見る
    export_csv(results, out_path, log=log)
    _prog("write")
    return out_path


# =====================================================================
# GUI
# =====================================================================

class App(_BASE_TK):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(720, 480)
        self._setup_fonts()

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
        # [中止]の合図。ワーカースレッドが要所で見る（強制終了はできないため）
        self._stop_flag = threading.Event()

        self._build()
        self._setup_dnd()
        self._prefill_from_argv(sys.argv[1:])
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._drain_queue)

    # ---- フォント -----------------------------------------------------
    def _setup_fonts(self) -> None:
        """
        画面の文字を実画面のDPIに合わせ、太字にならないよう明示する。

        DPI対応にしただけでは Tk の論理DPIが 96 のままなので、
        実画面のDPIに合わせて scaling を入れ直す（これをやらないと
        高DPI環境で文字が小さくなる）。
        """
        try:
            self.tk.call("tk", "scaling", self.winfo_fpixels("1i") / 72.0)
        except Exception:
            pass

        # 既定フォントが太字設定になっている環境があるため normal を明示する
        for name in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont"):
            try:
                tkfont.nametofont(name).configure(weight="normal")
            except Exception:
                pass

        # ログ欄も「Cue CSV」「MP3 フォルダ」などのラベルと同じ字にする。
        # Text ウィジェットの既定は等幅の TkFixedFont で、日本語が痩せたり
        # 太ったりして見えるため、画面の既定フォントをそのまま複製して使う。
        self.font_log = tkfont.Font(font=tkfont.nametofont("TkDefaultFont"))

    # ---- ドラッグ＆ドロップ -------------------------------------------
    def _setup_dnd(self) -> None:
        """ウィンドウにファイル／フォルダをドロップできるようにする。"""
        if not HAS_DND:
            return
        for widget in (self, self.ent_csv, self.ent_dir, self.ent_out):
            try:
                widget.drop_target_register(DND_FILES)
            except Exception:
                continue
        # 各入力欄に落とした場合はその欄へ、それ以外はファイル種別で振り分ける
        self.ent_csv.dnd_bind("<<Drop>>", lambda e: self._on_drop(e, "csv"))
        self.ent_dir.dnd_bind("<<Drop>>", lambda e: self._on_drop(e, "dir"))
        self.ent_out.dnd_bind("<<Drop>>", lambda e: self._on_drop(e, "out"))
        self.dnd_bind("<<Drop>>", lambda e: self._on_drop(e, None))

    def _on_drop(self, event, target: str | None):
        """
        ドロップされたパスを該当欄に入れる。

        event.data は Tcl のリスト形式（空白を含むパスは {} で括られる）なので
        splitlist で分解する。
        """
        try:
            paths = [Path(p) for p in self.tk.splitlist(event.data)]
        except Exception:
            paths = [Path(str(event.data).strip("{}"))]
        self._apply_paths(paths, target)
        return event.action

    def _apply_paths(self, paths: list[Path], target: str | None = None) -> None:
        """パスの種類を見て Cue CSV / MP3フォルダ / 出力先 に振り分ける。"""
        for p in paths:
            if target == "out":
                self.var_out.set(str(p))
            elif p.is_dir():
                # フォルダはどの欄に落としても MP3 フォルダとして扱う
                self.var_dir.set(str(p))
            elif target == "dir":
                # MP3フォルダ欄にファイルを落とされたら、その入れ物を採る
                self.var_dir.set(str(p.parent))
            elif p.suffix.lower() == ".csv" or target == "csv":
                self._set_csv(p)
            else:
                # CSV 以外のファイル（MP3等）はその親フォルダを MP3 フォルダに
                self.var_dir.set(str(p.parent))

    def _prefill_from_argv(self, argv: list[str]) -> None:
        """exe のアイコンにファイルをドロップして起動した場合の取り込み。"""
        paths = [Path(a) for a in argv if not a.startswith("-")]
        if paths:
            self._apply_paths([p for p in paths if p.exists()])

    def _set_csv(self, p: Path) -> None:
        self.var_csv.set(str(p))
        # 出力先が空なら CSV と同じ場所に自動で決めておく（利用者に考えさせない）
        if not self.var_out.get():
            self.var_out.set(str(p.with_name(f"{p.stem}_mp3情報.csv")))

    # ---- 画面構築 -----------------------------------------------------
    def _build(self) -> None:
        pad = dict(padx=10, pady=6)
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        self.ent_csv = self._row(frm, 0, "Cue CSV", self.var_csv, self._pick_csv)
        self.ent_dir = self._row(frm, 1, "MP3 フォルダ", self.var_dir, self._pick_dir)
        self.ent_out = self._row(frm, 2, "出力先 CSV", self.var_out, self._pick_out)

        if HAS_DND:
            ttk.Label(frm, text="※ Cue CSV や MP3 フォルダは、この画面に直接ドラッグ＆ドロップでも指定できます",
                      foreground="#666").grid(
                row=3, column=1, columnspan=2, sticky="w", padx=10)

        ttk.Checkbutton(frm, text="部分一致も含める（曲名の一部が一致するファイルも拾う）",
                        variable=self.var_partial).grid(
            row=4, column=1, columnspan=2, sticky="w", **pad)

        # 実行と中止は並べて置く（実行中しか押せないボタンを別行にすると見失うため）
        btns = ttk.Frame(frm)
        btns.grid(row=5, column=1, sticky="ew", **pad)
        btns.columnconfigure(0, weight=3)
        btns.columnconfigure(1, weight=1)
        self.btn_run = ttk.Button(btns, text="実　行", command=self._on_run)
        self.btn_run.grid(row=0, column=0, sticky="ew")
        self.btn_stop = ttk.Button(btns, text="中　止", command=self._on_stop,
                                   state="disabled")
        self.btn_stop.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        self.btn_open = ttk.Button(frm, text="出力先を開く", command=self._open_out,
                                   state="disabled")
        self.btn_open.grid(row=5, column=2, sticky="ew", **pad)

        self.prog = ttk.Progressbar(frm, mode="determinate", maximum=100)
        self.prog.grid(row=6, column=0, columnspan=3, sticky="ew", **pad)

        ttk.Label(frm, textvariable=self.var_status, foreground="#555").grid(
            row=7, column=0, columnspan=3, sticky="w", padx=10)

        frm.rowconfigure(8, weight=1)
        box = ttk.Frame(frm)
        box.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=10, pady=(4, 10))
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)
        # 等幅ではなくなったので、長いパスは折り返して見せる
        self.log = tk.Text(box, height=14, wrap="word", state="disabled",
                           font=self.font_log)
        self.log.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.log.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=sb.set)

    def _row(self, parent, r: int, label: str, var: tk.StringVar, cmd) -> ttk.Entry:
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=10, pady=6)
        ent = ttk.Entry(parent, textvariable=var)
        ent.grid(row=r, column=1, sticky="ew", pady=6)
        ttk.Button(parent, text="参照...", command=cmd).grid(
            row=r, column=2, sticky="ew", padx=10, pady=6)
        return ent

    # ---- ファイル選択 -------------------------------------------------
    def _pick_csv(self) -> None:
        p = filedialog.askopenfilename(
            title="NUENDO Cue CSV を選択",
            filetypes=[("CSV ファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )
        if p:
            self._set_csv(Path(p))

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
        self._stop_flag.clear()
        self.btn_run.configure(state="disabled", text="実行中...")
        self.btn_stop.configure(state="normal", text="中　止")
        self.btn_open.configure(state="disabled")
        self._set_progress(0)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        threading.Thread(
            target=self._worker,
            args=(csv_path, mp3_dir, Path(out_txt), self.var_partial.get()),
            daemon=True,
        ).start()

    def _on_stop(self) -> None:
        """[中止]。処理は要所でしか止まれないので、押した直後に反応だけ返す。"""
        if not self._running:
            return
        self._stop_flag.set()
        self.btn_stop.configure(state="disabled", text="中止中...")
        self.var_status.set("中止しています... 実行中の分が終わり次第止まります")

    def _on_close(self) -> None:
        """実行中に×で閉じられた場合。閉じる前に処理へ中止を伝える。"""
        if self._running:
            if not messagebox.askyesno(APP_TITLE, "実行中です。中止して終了しますか？"):
                return
            self._stop_flag.set()
        self.destroy()

    def _worker(self, csv_path: Path, mp3_dir: Path, out_path: Path,
                allow_partial: bool) -> None:
        def log(msg: str, replace: bool = False) -> None:
            self._q.put(("log", (str(msg), replace)))

        last = [-1.0]

        def progress(pct) -> None:
            # 同じ値を送り続けても意味が無いので、1%以上動いた時だけ流す
            if pct is not None and abs(pct - last[0]) < 1.0:
                return
            last[0] = -1.0 if pct is None else pct
            self._q.put(("prog", pct))

        try:
            result = run_finder(csv_path, mp3_dir, out_path, allow_partial, log,
                                should_stop=self._stop_flag.is_set,
                                on_progress=progress)
            self._q.put(("done", result))
        except Cancelled:
            self._q.put(("cancelled", None))
        except Exception as e:
            self._q.put(("error", (f"{type(e).__name__}: {e}", traceback.format_exc())))

    # ---- ワーカーからのメッセージを UI に反映 --------------------------
    def _drain_queue(self) -> None:
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "log":
                    self._append(*payload)
                elif kind == "prog":
                    self._set_progress(payload)
                elif kind == "done":
                    self._finish(payload)
                elif kind == "cancelled":
                    self._cancelled()
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

    def _set_progress(self, pct) -> None:
        """
        進捗バーを進める。

        pct が None の段階（MP3スキャン中）は総数が分からないので、
        止まって見えないよう流し表示（indeterminate）に切り替える。
        """
        if pct is None:
            if str(self.prog.cget("mode")) != "indeterminate":
                self.prog.configure(mode="indeterminate")
                self.prog.start(12)
            return
        if str(self.prog.cget("mode")) != "determinate":
            self.prog.stop()
            self.prog.configure(mode="determinate", maximum=100)
        self.prog.configure(value=max(0.0, min(100.0, float(pct))))

    def _reset(self) -> None:
        self._running = False
        self.prog.stop()
        self.prog.configure(mode="determinate", value=0)
        self.btn_run.configure(state="normal", text="実　行")
        self.btn_stop.configure(state="disabled", text="中　止")

    def _cancelled(self) -> None:
        """中止された。エラーではないので警告ダイアログは出さない。"""
        self._reset()
        self._append("\n[中止] 処理を中止しました。CSV は出力していません。", False)
        self.var_status.set("中止しました")

    def _finish(self, out: Path | None) -> None:
        self._reset()
        self._last_out = out
        if out and out.exists():
            self._set_progress(100)   # 終わったことが見て分かるよう満たしておく
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
    if argv and argv[0] == "--selftest":
        # exe に固めた状態で DPI設定・フォント・D&D が生きているか確認する。
        # windowed exe には標準出力が無いのでファイルに書く。
        enable_dpi_awareness()
        app = App()
        app.withdraw()
        app.update()
        lines = [
            f"tkinterdnd2   : {HAS_DND}",
            f"tkdnd         : {app.tk.call('package', 'require', 'tkdnd') if HAS_DND else '-'}",
            f"mutagen       : {HAS_MUTAGEN}",
            f"screen dpi    : {app.winfo_fpixels('1i'):.1f}",
            f"tk scaling    : {float(app.tk.call('tk', 'scaling')):.3f}",
            f"log font      : {app.font_log.actual('family')} "
            f"{app.font_log.actual('size')} {app.font_log.actual('weight')}",
            f"label font    : {tkfont.nametofont('TkDefaultFont').actual('family')} "
            f"{tkfont.nametofont('TkDefaultFont').actual('size')} "
            f"{tkfont.nametofont('TkDefaultFont').actual('weight')}",
            f"stop button   : {app.btn_stop.cget('state')}",
        ]
        app.destroy()
        Path(argv[1] if len(argv) > 1 else "selftest.log").write_text(
            "\n".join(lines), encoding="utf-8")
        return
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

    enable_dpi_awareness()   # Tk のウィンドウを作る前に呼ぶ
    App().mainloop()


if __name__ == "__main__":
    main()
