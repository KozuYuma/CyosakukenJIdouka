@echo off
chcp 65001 > nul
cd /d %~dp0..

echo ============================================
echo  NUENDO MP3 Finder
echo ============================================
echo.

set /p CUE_CSV="[1] Cue CSV path (drag and drop OK): "
set CUE_CSV=%CUE_CSV:"=%

if not exist "%CUE_CSV%" (
    echo ERROR: File not found: %CUE_CSV%
    goto END
)

echo.
set /p MP3_DIR="[2] MP3 folder path (drag and drop OK): "
set MP3_DIR=%MP3_DIR:"=%

if not exist "%MP3_DIR%" (
    echo ERROR: Folder not found: %MP3_DIR%
    goto END
)

echo.
set /p VERBOSE_INPUT="[3] Show detail properties? [y/N]: "
set VERBOSE_OPT=
if /i "%VERBOSE_INPUT%"=="y" set VERBOSE_OPT=--verbose

echo.
set /p CSV_OUT="[4] Output CSV path (press Enter to skip): "
set CSV_OUT=%CSV_OUT:"=%
set OUTPUT_OPT=
if not "%CSV_OUT%"=="" set OUTPUT_OPT=--output "%CSV_OUT%"

echo.
echo ============================================
echo  Running...
echo ============================================
echo.

.venv\Scripts\python.exe scripts\nuendo_mp3_finder.py "%CUE_CSV%" "%MP3_DIR%" %VERBOSE_OPT% %OUTPUT_OPT%

:END
echo.
echo ============================================
echo  Done. Press any key to close.
echo ============================================
pause > nul