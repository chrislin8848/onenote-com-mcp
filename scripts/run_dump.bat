@echo off
REM run_dump.bat -- guest-side fixture-dump runner (manual step, Phase 5b intake).
REM
REM Launched by the `onenote-dump` scheduled task (schtasks /run) which remote_dump.sh
REM triggers over SSH; registered with /it so it runs in the AUTOLOGON interactive session
REM (COM cannot activate OneNote from session 0). Do NOT run this directly over SSH.
REM
REM Dumps go to test-results\dump (NOT tests\fixtures) so already-sanitized fixtures are
REM never clobbered; the host inspects and moves files over by hand (PII lesson).
REM Page NAMES travel via scripts\dump_pages.txt (UTF-8) -- never on this command line,
REM where cmd.exe codepage would mangle non-ASCII.
setlocal
cd /d C:\onenote-mcp
if not exist test-results mkdir test-results
del /q test-results\dump_exit_code.txt 2>nul

call uv sync > test-results\dump.log 2>&1
call uv run python scripts/dump_fixtures.py ^
  --scope "{C94E632E-9829-45FF-914E-5E4031B2439D}{1}{B0}" ^
  --names-file scripts/dump_pages.txt ^
  --out test-results/dump >> test-results\dump.log 2>&1
set RC=%ERRORLEVEL%
REM redirect-first (digit+`>` would parse as stdin redirect and leave the file empty)
>test-results\dump_exit_code.txt echo %RC%
endlocal
