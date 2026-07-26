@echo off
cd /d %~dp0

if not exist ".venv\Scripts\python.exe" (
    echo .venv not found.
    pause
    exit /b 1
)

echo Starting... http://localhost:8501
echo Press Ctrl+C to stop.
echo.

.venv\Scripts\python.exe -m streamlit run app.py

pause
