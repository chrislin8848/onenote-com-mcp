@echo off
REM run_tier2.bat -- guest-side Tier 2 runner (SPEC §2.4, Phase 3).
REM
REM Launched by the `onenote-tier2` scheduled task (schtasks /run) which remote_test.sh
REM triggers over SSH. The task is registered with /it so this runs in the AUTOLOGON
REM interactive session -- the only session where COM can activate OneNote (a bare SSH
REM command lands in session 0 and CoCreateInstance fails). Do NOT run this directly over SSH.
REM
REM Writes three artifacts under test-results\ for remote_test.sh to collect:
REM   tier2.xml       JUnit results
REM   tier2.log       full pytest + uv output
REM   exit_code.txt   the pytest exit code (the sentinel remote_test.sh polls + mirrors)
setlocal
cd /d C:\onenote-mcp
if not exist test-results mkdir test-results
del /q test-results\exit_code.txt 2>nul

call uv sync > test-results\tier2.log 2>&1
call uv run pytest -m windows --junitxml=test-results\tier2.xml -ra >> test-results\tier2.log 2>&1
REM capture pytest's exit code immediately (next command would overwrite ERRORLEVEL)
set RC=%ERRORLEVEL%
echo %RC%> test-results\exit_code.txt
endlocal
