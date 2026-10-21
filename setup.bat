@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "TRANSLATE_MODEL=translategemma:4b"
set "REQUIREMENTS_FILE=requirements.txt"
set "OLLAMA_EXE="
set "OLLAMA_JUST_INSTALLED="

echo ============================================================
echo TranslateGemma Markdown Batch Translator Setup
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    where py >nul 2>nul
    if errorlevel 1 (
        echo ERROR: Python launcher "py" was not found.
        echo Install Python first, then run setup.bat again.
        echo.
        pause
        exit /b 1
    )

    echo Creating virtual environment...
    py -m venv .venv
    if errorlevel 1 goto :error
)

echo Updating pip...
".venv\Scripts\python.exe" -m pip install -U pip
if errorlevel 1 goto :error

if not exist "%REQUIREMENTS_FILE%" (
    echo ERROR: Requirements file was not found: "%REQUIREMENTS_FILE%"
    goto :error
)

echo Installing Python dependencies...
".venv\Scripts\python.exe" -m pip install -r "%REQUIREMENTS_FILE%"
if errorlevel 1 goto :error

echo Verifying Python imports...
".venv\Scripts\python.exe" -c "import ollama, win32com.client, pythoncom"
if errorlevel 1 goto :error

echo.
echo Checking Ollama runtime...
call :find_ollama
if defined OLLAMA_EXE goto :ollama_found

echo Ollama runtime was not found.
echo Downloading and installing Ollama for Windows...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference = 'Stop'; Invoke-RestMethod 'https://ollama.com/install.ps1' | Invoke-Expression"
if errorlevel 1 (
    echo ERROR: Ollama installation failed.
    goto :error
)
set "OLLAMA_JUST_INSTALLED=1"

call :find_ollama
if not defined OLLAMA_EXE (
    echo ERROR: Ollama was installed, but ollama.exe could not be found.
    echo Close this window, open a new terminal, and run setup.bat again.
    goto :error
)

:ollama_found
echo Using Ollama: "%OLLAMA_EXE%"

"%OLLAMA_EXE%" list >nul 2>nul
if not errorlevel 1 goto :ollama_ready

if defined OLLAMA_JUST_INSTALLED (
    echo Waiting for the newly installed Ollama app to start...
    for /L %%I in (1,1,15) do (
        timeout /t 1 /nobreak >nul 2>nul
        "%OLLAMA_EXE%" list >nul 2>nul
        if not errorlevel 1 goto :ollama_ready
    )
)

echo Ollama server is not responding. Starting it...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference = 'Stop'; Start-Process -FilePath $env:OLLAMA_EXE -ArgumentList 'serve' -WindowStyle Hidden"
if errorlevel 1 (
    echo ERROR: Could not start the Ollama server.
    goto :error
)

echo Waiting for the Ollama server...
for /L %%I in (1,1,60) do (
    "%OLLAMA_EXE%" list >nul 2>nul
    if not errorlevel 1 goto :ollama_ready
    timeout /t 1 /nobreak >nul 2>nul
)

echo ERROR: The Ollama server did not become ready within 60 seconds.
echo Try starting Ollama from the Start menu, then run setup.bat again.
echo To see the server error directly, run: "%OLLAMA_EXE%" serve
goto :error

:ollama_ready
echo Ollama server is ready.

call :ensure_model "%TRANSLATE_MODEL%"
if errorlevel 1 goto :error

echo.
echo ============================================================
echo SETUP COMPLETE
echo ============================================================
echo.
echo Next:
echo   Put Markdown files in input, then run run.bat
echo.
pause
exit /b 0

:error
echo.
echo SETUP FAILED.
pause
exit /b 1

:ensure_model
set "MODEL_TO_CHECK=%~1"
echo Checking model %MODEL_TO_CHECK%...
"%OLLAMA_EXE%" show "%MODEL_TO_CHECK%" >nul 2>nul
if not errorlevel 1 (
    echo Model %MODEL_TO_CHECK% is ready.
    exit /b 0
)

echo Model was not found. Downloading %MODEL_TO_CHECK%...
echo This is a large download and may take a while.
"%OLLAMA_EXE%" pull "%MODEL_TO_CHECK%"
if errorlevel 1 (
    echo ERROR: Could not download model %MODEL_TO_CHECK%.
    exit /b 1
)

"%OLLAMA_EXE%" show "%MODEL_TO_CHECK%" >nul 2>nul
if errorlevel 1 (
    echo ERROR: Model verification failed after download: %MODEL_TO_CHECK%
    exit /b 1
)

echo Model %MODEL_TO_CHECK% is ready.
exit /b 0

:find_ollama
set "OLLAMA_EXE="
for /F "delims=" %%I in ('where.exe ollama.exe 2^>nul') do if not defined OLLAMA_EXE set "OLLAMA_EXE=%%~fI"
if defined OLLAMA_EXE exit /b 0

if defined OLLAMA_INSTALL_DIR (
    if exist "%OLLAMA_INSTALL_DIR%\ollama.exe" set "OLLAMA_EXE=%OLLAMA_INSTALL_DIR%\ollama.exe"
)
if defined OLLAMA_EXE exit /b 0

if exist "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" (
    set "OLLAMA_EXE=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"
)
exit /b 0
