; Inno Setup script — OneNoteMCP-Setup.exe (SPEC §8).
; Compiled on the VM with the Inno Setup compiler:  iscc packaging\onenote-mcp.iss
; (iscc.exe is NOT pip-installable — install Inno Setup 6 on the build VM first; see README.md.)
;
; Per-user install (no admin): each employee runs this like a normal app (SPEC §1.3 — no MDM/
; GPO push). After copying files it runs `OneNoteMCP.exe --configure`, which detects BOTH the
; regular and Microsoft Store Claude Desktop config locations and registers the server (§8).

#define AppName "OneNote MCP Server"
#define AppPublisher "Chris Lin"
#define AppVersion "0.9.9"
#define ExeName "OneNoteMCP.exe"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
; Per-user install — no administrator rights required.
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\OneNoteMCP
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=dist\installer
OutputBaseFilename=OneNoteMCP-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; The server needs the 64-bit OneNote desktop COM server; ship a 64-bit build.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
; The whole PyInstaller onedir output (OneNoteMCP.exe + its dependency folder).
Source: "dist\OneNoteMCP\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#ExeName}"

[Run]
; Register the server in Claude Desktop's config (both regular + Store) right after install.
Filename: "{app}\{#ExeName}"; Parameters: "--configure"; \
    StatusMsg: "Registering with Claude Desktop..."; Flags: runhidden waituntilterminated

[UninstallRun]
; Best-effort: leave Claude config alone on uninstall (the entry points at a now-removed exe;
; harmless, and removing it would need a --deconfigure we deliberately don't ship this round).
