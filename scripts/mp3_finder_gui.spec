# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec: NUENDO MP3 Finder (GUI)

build:
    scripts\\build_mp3_finder_exe.ps1
出力:
    dist\\NUENDO_MP3_Finder.exe （単体で配布できる1ファイル）
"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_all

# spec 実行時は __file__ が無いので SPECPATH（PyInstaller が定義）を使う
SCRIPTS = Path(SPECPATH)  # noqa: F821

# ドラッグ＆ドロップ用の tkdnd は Tcl のパッケージ実体（.dll と .tcl）なので、
# import 解析では拾われない。collect_all で丸ごと同梱する。
dnd_datas, dnd_bins, dnd_hidden = collect_all("tkinterdnd2")  # noqa: F821

a = Analysis(
    [str(SCRIPTS / "mp3_finder_gui.py")],
    pathex=[str(SCRIPTS)],          # nuendo_mp3_finder を import できるように
    binaries=dnd_bins,
    datas=dnd_datas,
    hiddenimports=["nuendo_mp3_finder"] + dnd_hidden,
    hookspath=[],
    runtime_hooks=[],
    # 配布サイズを抑えるため、このツールが使わない重い依存を明示的に外す
    excludes=[
        "streamlit", "pandas", "numpy", "matplotlib", "openpyxl",
        "bs4", "lxml", "requests", "PIL", "pytest",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)  # noqa: F821

exe = EXE(  # noqa: F821
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="NUENDO_MP3_Finder",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,        # コンソール窓を出さない（GUIのみ）
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
