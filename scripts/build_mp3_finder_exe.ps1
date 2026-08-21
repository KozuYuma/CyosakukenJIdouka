<#
    NUENDO MP3 Finder の exe をビルドする。

    使い方（PowerShell で）:
        .\scripts\build_mp3_finder_exe.ps1

    出力: dist\NUENDO_MP3_Finder.exe

    ビルドには専用の仮想環境 .venv-build を使う。アプリ本体の .venv には
    streamlit や pandas が入っており、混ざると exe が無用に大きくなるため。
#>

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$BuildVenv = Join-Path $Root ".venv-build"
$Py        = Join-Path $BuildVenv "Scripts\python.exe"

if (-not (Test-Path $Py)) {
    Write-Host "[1/3] ビルド用の仮想環境を作成: $BuildVenv"
    # 本体の .venv と同じ Python でビルドする
    & (Join-Path $Root ".venv\Scripts\python.exe") -m venv $BuildVenv
    if (-not (Test-Path $Py)) {
        Write-Host "仮想環境の作成に失敗しました: $BuildVenv" -ForegroundColor Red
        Write-Host "既存の .venv-build を削除してから、もう一度実行してください。"
        exit 1
    }
} else {
    Write-Host "[1/3] ビルド用の仮想環境を再利用: $BuildVenv"
}

Write-Host "[2/3] 依存をインストール (pyinstaller, mutagen)"
& $Py -m pip install --upgrade pip --quiet
& $Py -m pip install --upgrade pyinstaller mutagen --quiet

Write-Host "[3/3] ビルド中..."
# 前回ビルドした exe が起動したままだと上書きできない
Get-Process NUENDO_MP3_Finder -ErrorAction SilentlyContinue | Stop-Process -Force

& $Py -m PyInstaller --noconfirm --clean "scripts\mp3_finder_gui.spec"
$BuildOk = ($LASTEXITCODE -eq 0)

$Exe = Join-Path $Root "dist\NUENDO_MP3_Finder.exe"
if ($BuildOk -and (Test-Path $Exe)) {
    $Size = [math]::Round((Get-Item $Exe).Length / 1MB, 1)
    Write-Host ""
    Write-Host "完了: $Exe  ($Size MB)" -ForegroundColor Green
    Write-Host "この exe 1つだけを配布すれば動きます（Python のインストール不要）。"
} else {
    Write-Host "ビルドに失敗しました。上のログを確認してください。" -ForegroundColor Red
    exit 1
}
