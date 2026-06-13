@echo off
REM run_sync_validate.bat -- guest-side runner for the copy under-sync Tier-2 check.
REM Launched by an interactive (/it) scheduled task so COM can reach OneNote (NEVER plain SSH).
REM Writes test-results\sync_validate.log + test-results\sync_validate_exit.txt (poll sentinel).
setlocal
cd /d C:\onenote-mcp
if not exist test-results mkdir test-results
del /q test-results\sync_validate_exit.txt 2>nul

call uv run python scripts\tier2_sync_validate.py > test-results\sync_validate.log 2>&1
set RC=%ERRORLEVEL%
>test-results\sync_validate_exit.txt echo %RC%
endlocal
