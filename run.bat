@echo off
chcp 65001 > nul
cd /d %~dp0

echo ========================================
echo  著作権調査支援ツール 起動中...
echo ========================================

REM 仮想環境がなければ作成
if not exist ".venv" (
    echo 初回セットアップ: 仮想環境を作成しています...
    python -m venv .venv
    if errorlevel 1 (
        echo エラー: Pythonが見つかりません。python.org からインストールしてください。
        pause
        exit /b 1
    )
)

REM 仮想環境を有効化
call .venv\Scripts\activate.bat

REM ライブラリをインストール
echo ライブラリを確認・インストール中...
pip install -r requirements.txt -q

REM アプリ起動
echo.
echo ブラウザで http://localhost:8501 が開きます
echo 終了するには Ctrl+C を押してください
echo.
streamlit run app.py

pause
