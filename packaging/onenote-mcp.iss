; Inno Setup script — OneNoteMCP-Setup.exe (SPEC §8).
; Compiled on the VM with the Inno Setup compiler:  iscc packaging\onenote-mcp.iss
; (iscc.exe is NOT pip-installable — install Inno Setup 6 on the build VM first; see README.md.)
;
; Per-user install (no admin): each employee runs this like a normal app (SPEC §1.3 — no MDM/
; GPO push). After copying files it runs `OneNoteMCP.exe --configure`, which independently detects
; every supported MCP client — Claude Desktop (regular + Microsoft Store) and Antigravity
; (CLI/IDE, one shared `.gemini` config) — and registers the server into each one installed (§8).

#define AppName "OneNote MCP Server"
#define AppPublisher "Chris Lin"
#define AppAuthor "Chris Lin"
#define AppVersion "1.2.2"
#define ExeName "OneNoteMCP.exe"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppCopyright=Author: {#AppAuthor}
; Per-user install — no administrator rights required.
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\OneNoteMCP
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; Paths are resolved relative to this script's dir ({#SourcePath}); the freeze output and the
; installer both live under the repo-root dist\ (one level up from packaging\).
OutputDir={#SourcePath}..\dist\installer
OutputBaseFilename=OneNoteMCP-Setup_{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; Post-install info page (Traditional Chinese): reminds the user to paste the recommended
; Claude Desktop custom instruction (account-level, can't be auto-applied by an installer).
; Shown on a normal install; skipped by /VERYSILENT. UTF-8 (with BOM) so CJK renders.
InfoAfterFile={#SourcePath}post_install_zh-TW.txt
; The server needs the 64-bit OneNote desktop COM server; ship a 64-bit build.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
; The whole PyInstaller onedir output (OneNoteMCP.exe + its dependency folder).
Source: "{#SourcePath}..\dist\OneNoteMCP\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#ExeName}"
; "Re-detect and configure" shortcut (SPEC §8): a user who installs a client (Claude Desktop or
; Antigravity) AFTER this installer just clicks it to register the server with the new client —
; no reinstall. Launched via cmd /k so the result stays visible (the exe prints + exits at once).
Name: "{group}\OneNoteMCP — 重新偵測並設定"; Filename: "{cmd}"; \
    Parameters: "/k ""{app}\{#ExeName}"" --configure"; IconFilename: "{app}\{#ExeName}"

[Run]
; Register the server with every detected MCP client (Claude Desktop regular + Store, Antigravity)
; right after install. Writes nothing if none is installed (the install still completes; the user
; re-runs the shortcut above once a client is present).
Filename: "{app}\{#ExeName}"; Parameters: "--configure"; \
    StatusMsg: "Registering with Claude Desktop / Antigravity..."; \
    Flags: runhidden waituntilterminated

[UninstallRun]
; Best-effort: leave Claude config alone on uninstall (the entry points at a now-removed exe;
; harmless, and removing it would need a --deconfigure we deliberately don't ship this round).

; NOTE: the broken-OneNote-typelib case (a stale version subkey with no win32/win64 mapping,
; which poisons LoadRegTypeLib → TYPE_E_LIBNOTREGISTERED on a cold OneNote launch) is handled
; automatically by the post-install `--configure` step above: it writes a per-user (HKCU) shim
; pointing the broken version at the healthy typelib file — no admin, reversible, no user action.
; VM-reproduced + validated 2026-06-12. The runtime `--selftest` still detects+reports it as a
; safety net if the auto-repair ever can't run.
