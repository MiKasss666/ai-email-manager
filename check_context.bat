@echo off
chcp 65001 >nul
REM ============================================================
REM  Email Manager - integrity context check
REM
REM  DOUBLE-CLICK THIS FILE FROM YOUR NORMAL DESKTOP.
REM  It only reads and prints - it changes nothing.
REM
REM  Why: a process running at Low integrity CANNOT start or
REM  reach Outlook through COM. The symptom is
REM  "com_error: (-2146959355, Server execution failed)" plus
REM  Outlook popping up "Cannot start Microsoft Outlook.".
REM  This script tells you which integrity level you get, and
REM  whether the built exe carries a Low label of its own.
REM ============================================================
cd /d "%~dp0"
echo ============================================================
echo  [1] Integrity level of THIS window  (want: Medium)
echo ============================================================
whoami /groups | findstr /i "Mandatory"
echo.
echo ============================================================
echo  [2] Mandatory label on the built exe  (want: NO Low line)
echo ============================================================
if exist "EmailManager\EmailManager.exe" (
    icacls "EmailManager\EmailManager.exe" | findstr /i "Mandatory"
    echo    ^(no output above = no label = good^)
) else (
    echo    EmailManager\EmailManager.exe not found - build it first
)
echo.
echo ============================================================
echo  [3] Mandatory label on this script itself
echo ============================================================
icacls "%~f0" | findstr /i "Mandatory"
echo    ^(no output above = no label = good^)
echo.
echo ============================================================
echo  [4] Is Outlook running?
echo ============================================================
tasklist /fi "imagename eq outlook.exe" /nh
echo.
echo ============================================================
echo  How to read this
echo ============================================================
echo   [1] Medium + [2] no Low   -^> context is fine, the app can drive Outlook
echo   [1] Low                   -^> THIS window is sandboxed. Launch the app
echo                                from a normal Explorer / Start menu window.
echo   [2] Low                   -^> the exe file itself is tainted. Rebuild by
echo                                double-clicking build.bat from a normal
echo                                window, or run start.bat (source) instead.
echo.
pause
