# packaging — freeze + installer (Phase 6 Stage 5, SPEC §8)

Produces `OneNoteMCP-Setup.exe`: a per-user Windows installer that drops the frozen MCP server
and registers it with every detected MCP client — Claude Desktop (regular + Store) and Antigravity
(CLI/IDE, one shared `.gemini` config) (SPEC §8). **Everything here is built on the Windows VM** —
PyInstaller does not cross-compile, so it cannot run on the Linux dev host.

## Files

- `onenote_mcp_entry.py` — PyInstaller entry script (runs `server.main`).
- `onenote-mcp.spec` — PyInstaller spec (onedir, console exe, win32com makepy bundled).
- `rthook_win32com_gen_py.py` — runtime hook: a writable gen_py cache so `gencache.EnsureModule`
  can generate the OneNote typelib module at runtime (OneNote can't be late-bound — Phase 0b).
- `onenote-mcp.iss` — Inno Setup script → `dist/installer/OneNoteMCP-Setup.exe`.
- `build.bat` — orchestrates freeze → smoke → installer on the VM.

## The freeze gotcha (why it's not the SPEC's "use late-bound Dispatch")

SPEC §8 suggested late-bound `Dispatch` for freeze compatibility, but Phase 0b proved OneNote
**cannot** be late-bound (`GetIDsOfNames` can't resolve `GetHierarchy`). So the backend keeps
`gencache.EnsureModule` + coclass instantiation. The freeze instead bundles the makepy machinery
(spec `hiddenimports`) and redirects win32com's gen_py cache to a writable temp dir (runtime
hook), letting `EnsureModule` regenerate the typelib module on first call. The target machine has
OneNote installed (typelib registered), so this works without bundling a pre-generated module.

## VM prerequisites

- **uv** — already present (the Tier-2 loop uses it).
- **Inno Setup 6** — provides `iscc.exe`. NOT pip-installable; install once on the VM:
  `winget install --id JRSoftware.InnoSetup --source winget --accept-package-agreements
  --accept-source-agreements` (avoid the `msstore` source — its agreement prompt hangs a
  non-interactive SSH session). winget installs it **per-user, NOT on PATH**, at
  `%LocalAppData%\Programs\Inno Setup 6\ISCC.exe`; `build.bat` resolves `iscc` from PATH first,
  then that known location (and `%ProgramFiles(x86)%`), and only skips the installer step
  (exit code 2) if none is found.

## Build

On the VM, from `C:\onenote-mcp`:

```
packaging\build.bat
```

Steps: `uv sync --group packaging` → `pyinstaller packaging/onenote-mcp.spec` →
`OneNoteMCP.exe --configure` smoke → `iscc packaging/onenote-mcp.iss`. Output:
`dist/OneNoteMCP/` (frozen app) and `dist/installer/OneNoteMCP-Setup_<version>.exe`
(the installer filename carries the `AppVersion` from the .iss, e.g. `_0.9.9`).

## Validation (still VM-gated — author-then-verify)

After building, confirm on the VM:
1. `dist/OneNoteMCP/OneNoteMCP.exe` starts, binds OneNote via COM, and answers a read tool
   (the freeze's gen_py regeneration is the risk — watch for `BackendUnavailableError`).
2. Run the installer; check the `onenote` entry lands in EVERY detected client config —
   `claude_desktop_config.json` (regular and/or Store path) and Antigravity's
   `%USERPROFILE%\.gemini\config\mcp_config.json` — and that the client(s) launch the server.
   With no client installed, `--configure` should write nothing and say so (install still ok).
3. The Start Menu "OneNoteMCP — 重新偵測並設定" shortcut re-runs `--configure` (for a client
   installed after setup).
4. stdout carries only JSON-RPC (no stray prints) — else the client can't parse the stream.
