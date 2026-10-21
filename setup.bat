@echo off
REM ---------------------------------------------------------------------
REM  ErrorDigest - one-time setup. Double-click this once.
REM
REM  Creates a private Python environment inside this folder and installs
REM  the two packages the script needs. Nothing is installed system-wide,
REM  nothing is downloaded again afterwards, and nothing costs anything.
REM
REM  The pip steps are deliberately NOT quiet. The first version used
REM  --quiet and the operator reasonably asked "what should I install, I
REM  did not see anything to install in setup.bat" - a setup script that
REM  does its main job in silence looks like one that does nothing.
REM ---------------------------------------------------------------------
setlocal
cd /d "%~dp0"

echo.
echo ================================================
echo   ErrorDigest setup
echo ================================================
echo.

echo [1/4] Looking for Python...
where python >nul 2>nul
if errorlevel 1 goto nopython
python --version
echo.

echo [2/4] Creating the private environment in .venv ...
python -m venv .venv
if errorlevel 1 goto failed
echo       done.
echo.

echo [3/4] Installing openpyxl and pywin32 ...
echo       (this is the only download, and it is free)
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo.

echo [4/4] Checking the packages actually work ...
REM  "pip said OK" is not the same as "the package imports". pywin32 in
REM  particular can install and still fail to load its DLLs, which would
REM  otherwise only show up later as a confusing Outlook error.
".venv\Scripts\python.exe" -c "import openpyxl; print('      openpyxl', openpyxl.__version__, 'OK')"
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -c "import win32com.client; print('      pywin32 OK')"
if errorlevel 1 goto nopywin32

echo.
echo Creating the workbook ...
".venv\Scripts\python.exe" error_digest.py --init
if errorlevel 1 goto failed

echo.
echo ================================================
echo   Setup finished
echo ================================================
echo.
echo Next:
echo   1. Open error_digest.xlsx
echo   2. On the Config sheet set:
echo        folders    - e.g.  Inbox\Alerts
echo        draft_to   - your email address
echo   3. CLOSE Excel  (the script cannot write the file while it is open)
echo   4. Double-click run.bat
echo.
pause
exit /b 0

:nopython
echo.
echo   Python was not found on this PC.
echo.
echo   Install it from:  https://www.python.org/downloads/windows/
echo   IMPORTANT: tick "Add python.exe to PATH" on the FIRST screen.
echo   Then close this window, open a new one, and run setup.bat again.
echo.
pause
exit /b 1

:nopywin32
echo.
echo   pywin32 installed but will not load.
echo   Try this, then run setup.bat again:
echo       .venv\Scripts\python.exe -m pywin32_postinstall -install
echo.
pause
exit /b 1

:failed
echo.
echo   Setup failed. The error is in the output above.
echo.
pause
exit /b 1
