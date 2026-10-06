@echo off
REM ---------------------------------------------------------------------
REM  ErrorDigest - one-time setup. Double-click this once.
REM
REM  Creates a private Python environment inside this folder and installs
REM  the two packages the script needs. Nothing is installed system-wide,
REM  nothing is downloaded again afterwards, and nothing costs anything.
REM ---------------------------------------------------------------------
setlocal
cd /d "%~dp0"

echo.
echo === ErrorDigest setup ===
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on this PC.
    echo Install it from https://www.python.org/downloads/windows/
    echo IMPORTANT: tick "Add python.exe to PATH" on the first screen.
    echo.
    pause
    exit /b 1
)

echo Creating the private environment in .venv ...
python -m venv .venv
if errorlevel 1 goto failed

echo Installing openpyxl and pywin32 ...
".venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
".venv\Scripts\python.exe" -m pip install -r requirements.txt --quiet
if errorlevel 1 goto failed

echo Creating the workbook ...
".venv\Scripts\python.exe" error_digest.py --init
if errorlevel 1 goto failed

echo.
echo === Done ===
echo.
echo Next: open error_digest.xlsx, go to the Config sheet, and set
echo   folders    - the Outlook folder(s) to scan, e.g.  Inbox\Alerts
echo   draft_to   - who the digest draft should be addressed to
echo.
echo Then double-click run.bat
echo.
pause
exit /b 0

:failed
echo.
echo Setup failed. Read the message above, or send it to whoever set this up.
pause
exit /b 1
