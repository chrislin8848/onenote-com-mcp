@echo off
REM build.bat -- produce OneNoteMCP-Setup.exe on the Windows VM (SPEC §8, Phase 6 Stage 5).
REM PyInstaller does NOT cross-compile, so the freeze + installer build run here, not on the
REM Linux host. Run from the repo root (C:\onenote-mcp):  scripts\..\packaging\build.bat
REM
REM Prereqs on the VM:
REM   - uv (already present, used by the Tier-2 loop)
REM   - Inno Setup 6 -- NOT pip-installable; install once (see README.md). `winget install
REM     JRSoftware.InnoSetup` lands it per-user at %LocalAppData%\Programs\Inno Setup 6 (NOT on
REM     PATH), so we resolve iscc.exe from PATH first, then that known winget location.
setlocal
cd /d C:\onenote-mcp

echo === syncing build deps (PyInstaller) ===
call uv sync --group packaging || exit /b 1

echo === freezing server (PyInstaller, onedir) ===
call uv run pyinstaller packaging\onenote-mcp.spec --noconfirm --distpath dist --workpath build || exit /b 1

echo === smoke-checking the frozen exe (--configure is a safe, non-server invocation) ===
dist\OneNoteMCP\OneNoteMCP.exe --configure || echo (configure returned non-zero; check Claude Desktop presence)

echo === building installer (Inno Setup) ===
set "ISCC="
for /f "delims=" %%i in ('where iscc 2^>nul') do set "ISCC=%%i"
if not defined ISCC if exist "%LocalAppData%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC (
  echo iscc.exe not found -- install Inno Setup 6 on the VM, then re-run. Skipping installer.
  exit /b 2
)
echo using "%ISCC%"
"%ISCC%" packaging\onenote-mcp.iss || exit /b 1

echo === done: dist\installer\OneNoteMCP-Setup_<version>.exe ===
endlocal
