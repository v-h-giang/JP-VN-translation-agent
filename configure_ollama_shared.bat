@echo off
setlocal

echo ============================================================
echo Configure Ollama for shared use on one RAM-limited machine
echo ============================================================
echo.

setx OLLAMA_MAX_LOADED_MODELS 1 >nul
setx OLLAMA_NUM_PARALLEL 1 >nul

echo OLLAMA_MAX_LOADED_MODELS=1
echo OLLAMA_NUM_PARALLEL=1
echo.
echo Configuration saved for future Ollama processes.
echo Fully quit Ollama from the Windows system tray and start it again
echo for these settings to take effect.
echo.
pause
