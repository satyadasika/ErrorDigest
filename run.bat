@echo off
REM ---------------------------------------------------------------------
REM  ErrorDigest - normal run. Reads new email, updates the workbook,
REM  and opens an UNSENT draft in Outlook for you to review and send.
REM ---------------------------------------------------------------------
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Setup has not been run yet. Double-click setup.bat first.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" error_digest.py %*
echo.
pause
