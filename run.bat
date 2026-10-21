@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo TranslateGemma Markdown Batch Translator
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: Python virtual environment was not found.
    echo.
    echo Run setup.bat first.
    echo.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" run_translate_all.py

echo.
echo ============================================================
echo Finished.
echo ============================================================
pause
