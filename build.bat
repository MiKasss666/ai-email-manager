@echo off
chcp 65001 >nul
REM ============================================================
REM  Build a distributable folder (GUI). Run the exe inside it.
REM
REM  Output: dist\EmailManager\   ->  copied to  .\EmailManager\
REM          run  .\EmailManager\EmailManager.exe
REM
REM  WHY A FOLDER AND NOT A SINGLE-FILE EXE:
REM    A one-file exe unpacks itself into %TEMP% on EVERY launch.
REM    When that fails (antivirus, restricted temp, locked-down
REM    policy) you get "Could not create temporary directory!" and
REM    the app simply will not start. A folder build never touches
REM    %TEMP%, starts faster, and is far less likely to be flagged.
REM
REM    Want a real single file anyway? Change --onedir to --onefile
REM    and rebuild - but TEST IT on the target machine first.
REM
REM  To see error output, change --windowed to --console.
REM
REM  The app has no third-party runtime dependencies (Python standard
REM  library only), so no hidden imports or collect-submodules are needed.
REM
REM  NOTE: keep this file ASCII-only AND CRLF. cmd.exe parses .bat
REM  bytes with the local ANSI codepage (so non-ASCII comments turn
REM  into commands) and mis-parses LF-only line endings.
REM ============================================================
cd /d "%~dp0"

echo [1/4] Checking PyInstaller ...
py -3 -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo       Not installed, installing ...
    py -3 -m pip install pyinstaller
    if errorlevel 1 goto :fail
)

echo [2/4] Cleaning previous build output ...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist EmailManager rmdir /s /q EmailManager

echo [3/4] Building (about 1-2 minutes on the first run) ...
py -3 -m PyInstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --name EmailManager ^
    --icon EmailManager.ico ^
    app.py
if errorlevel 1 goto :fail

echo [4/4] Copying the folder next to this script ...
xcopy /e /i /y /q "dist\EmailManager" "EmailManager" >nul
if errorlevel 1 goto :fail

echo.
echo Done. Run: %~dp0EmailManager\EmailManager.exe
echo Ship the whole EmailManager folder - config.json appears inside it on first run.
pause
exit /b 0

:fail
echo.
echo Build failed - see the error output above.
pause
exit /b 1
