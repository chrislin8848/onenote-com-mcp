@echo off
REM run_spike.bat -- guest-side runner for the FROZEN resvg_py spike (insert_svg_image gate).
REM COM-free (pure rasterization), so unlike run_selftest.bat this is safe over plain SSH.
REM Runs the frozen spike exe; it rasterizes a Traditional-Chinese SVG and writes spike_out.png
REM into the CWD. We cd into test-results so the PNG + exit sentinel land where the host collects.
setlocal
cd /d C:\onenote-mcp
if not exist test-results mkdir test-results
del /q test-results\spike_exit.txt 2>nul
del /q test-results\spike_out.png 2>nul

cd /d C:\onenote-mcp\test-results
..\packaging\spike_svg\dist\spike_resvg\spike_resvg.exe > spike.log 2>&1
set RC=%ERRORLEVEL%
REM redirect-first (the run_tier2.bat `0>` footgun): write the value reliably
>spike_exit.txt echo %RC%
endlocal
