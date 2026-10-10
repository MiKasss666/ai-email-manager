@echo off
chcp 65001 >nul
REM ============================================================
REM  Launch from source (no exe involved, so SmartScreen never shows up)
REM
REM  Use this when Python + pywin32 are installed locally (dev / own use).
REM  To hand the app to someone without Python, run build.bat first.
REM
REM  Want full error output? Open a terminal in this folder and run:
REM      py -3 app.py
REM
REM  NOTE: keep this file ASCII-only. cmd.exe parses .bat bytes with the
REM  local ANSI codepage, so non-ASCII comments get mangled into commands.
REM ============================================================
cd /d "%~dp0"

echo Checking dependencies ...
py -3 -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo   pywin32 is missing, installing ...
    py -3 -m pip install pywin32
    if errorlevel 1 goto :fail
)

echo Pre-flight import check ...
py -3 -c "import core.config, core.ai, core.classifier, core.summarizer, core.outlook, gui.main_window" 2>&1
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
