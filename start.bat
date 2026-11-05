@echo off
chcp 65001 >nul
REM ============================================================
REM  Launch from source (no exe involved, so SmartScreen never shows up)
REM
REM  Use this when Python 3.9+ is installed locally (dev / own use).
REM  There are NO third-party dependencies: the Microsoft Graph backend
REM  uses the Python standard library only. To hand the app to someone
REM  without Python, run build.bat first.
REM
REM  Want full error output? Open a terminal in this folder and run:
REM      py -3 app.py
REM
REM  NOTE: keep this file ASCII-only. cmd.exe parses .bat bytes with the
REM  local ANSI codepage, so non-ASCII comments get mangled into commands.
REM ============================================================
cd /d "%~dp0"

echo Checking Python ...
py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1
if errorlevel 1 (
    echo   Python 3.9+ was not found. Install it from https://www.python.org/downloads/
    goto :fail
)

echo Pre-flight import check ...
py -3 -c "import core.config, core.ai, core.classifier, core.summarizer, core.graph, gui.main_window" 2>&1
if errorlevel 1 goto :fail

echo Starting the interface ...
REM pyw is the windowless launcher, so no leftover console window.
where pyw >nul 2>&1
if not errorlevel 1 (
    start "" pyw -3 app.py
) else (
    py -3 app.py
)
exit /b 0

:fail
echo.
echo Startup failed - open a terminal in this folder and run: py -3 app.py
pause
exit /b 1
