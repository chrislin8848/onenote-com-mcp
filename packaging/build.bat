@echo off
REM build.bat -- produce OneNoteMCP-Setup.exe on the Windows VM (SPEC §8, Phase 6 Stage 5).
REM PyInstaller does NOT cross-compile, so the freeze + installer build run here, not on the
REM Linux host. Run from the repo root (C:\onenote-mcp):  scripts\..\packaging\build.bat
REM
REM Prereqs on the VM:
REM   - uv (already present, used by the Tier-2 loop)
REM   - Inno Setup 6 (iscc.exe) on PATH -- NOT pip-installable; install it once (see README.md)
setlocal
cd /d C:\onenote-mcp

echo === syncing build deps (PyInstaller) ===
call uv sync --group packaging || exit /b 1

echo === freezing server (PyInstaller, onedir) ===
call uv run pyinstaller packaging\onenote-mcp.spec --noconfirm --distpath dist --workpath build || exit /b 1

echo === smoke-checking the frozen exe (--configure is a safe, non-server invocation) ===
dist\OneNoteMCP\OneNoteMCP.exe --configure || echo (configure returned non-zero; check Claude Desktop presence)

echo === building installer (Inno Setup) ===
where iscc >nul 2>nul
if errorlevel 1 (
  echo iscc.exe not on PATH -- install Inno Setup 6 on the VM, then re-run. Skipping installer.
  exit /b 2
)
iscc packaging\onenote-mcp.iss || exit /b 1

echo === done: dist\installer\OneNoteMCP-Setup.exe ===
endlocal
