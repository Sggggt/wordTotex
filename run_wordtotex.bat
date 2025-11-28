@echo off
setlocal

rem Change to repository root
cd /d "%~dp0"

set "PYTHON=python"

rem Ensure environment and dependencies in current Python
%PYTHON% "%~dp0setup_env.py"
if errorlevel 1 (
    echo Environment setup failed.
    pause
    exit /b 1
)

rem Configure host/port (adjust if firewall blocks defaults)
set "WORDTOTEX_HOST=127.0.0.1"
set "WORDTOTEX_PORT=8000"

rem Launch Flask app and open browser to UI
echo Starting web UI at http://%WORDTOTEX_HOST%:%WORDTOTEX_PORT% ...
start "" "http://%WORDTOTEX_HOST%:%WORDTOTEX_PORT%"
%PYTHON% "%~dp0app.py"

pause

endlocal
