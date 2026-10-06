@echo off
REM ---------------------------------------------------------------------
REM  ErrorDigest - scheduled run. Collects errors into the workbook and
REM  creates NO draft and NO window. This is what Task Scheduler calls.
REM  Run run.bat by hand later when you want the digest email.
REM ---------------------------------------------------------------------
setlocal
cd /d "%~dp0"
".venv\Scripts\python.exe" error_digest.py --no-draft >> digest_log.txt 2>&1
