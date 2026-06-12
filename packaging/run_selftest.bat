@echo off
REM run_selftest.bat -- guest-side COM smoke for the FROZEN exe (Phase 6 Stage 5).
REM Launched by the `onenote-selftest` scheduled task (/it) so it runs in the AUTOLOGON
REM interactive session -- the only session where COM can activate OneNote (a bare SSH
REM command lands in session 0 and CoCreateInstance fails). Do NOT run directly over SSH.
REM
REM Runs the frozen server's --selftest (bind OneNote via COM, list notebooks) and writes:
REM   test-results\selftest.log        full output
REM   test-results\selftest_exit.txt   exit code sentinel (remote_build.sh polls + mirrors)
setlocal
cd /d C:\onenote-mcp
if not exist test-results mkdir test-results
del /q test-results\selftest_exit.txt 2>nul

dist\OneNoteMCP\OneNoteMCP.exe --selftest > test-results\selftest.log 2>&1
set RC=%ERRORLEVEL%
REM redirect-first (the run_tier2.bat `0>` footgun): write the value reliably
>test-results\selftest_exit.txt echo %RC%
endlocal
