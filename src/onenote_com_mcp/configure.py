"""``--configure``: register this server in every detected MCP client's config (SPEC §8).

The installer (and a user re-running the exe with ``--configure``) calls this to add an
``onenote`` entry to each supported client's config file, pointing at THIS executable. The schema
is identical across clients: a single ``mcpServers`` object whose ``onenote`` key carries a stdio
``command``/``args`` entry. We support THREE independently-detected client variants (SPEC §8):

- **Claude Desktop, regular (per-user install):** config at ``%APPDATA%\\Claude\\
  claude_desktop_config.json``; detected by the install itself — the program dir
  ``%LOCALAPPDATA%\\Programs\\Claude`` or an Uninstall registry key (NOT by the config file, which
  doesn't exist before first launch).
- **Claude Desktop, Microsoft Store (MSIX):** config at ``%LOCALAPPDATA%\\Packages\\<Claude
  package>\\LocalCache\\Roaming\\Claude\\claude_desktop_config.json`` — the package family name is
  DETECTED by globbing (its exact value varies), never hardcoded; the package dir IS the install.
- **Antigravity (CLI + IDE share ONE config):** config at ``%USERPROFILE%\\.gemini\\config\\
  mcp_config.json`` — writing this single file covers both the CLI and the IDE. Detected by the
  install — CONFIRMED on a real machine 2026-06-16: the program dir
  ``%LOCALAPPDATA%\\Programs\\antigravity`` (lowercase; no ``agy`` CLI on PATH), with an
  ``agy``/``antigravity`` CLI on PATH or a ``~/.gemini`` dir as weaker forward-compat fallbacks.

Detection is per-variant and additive — every variant that is installed gets written (install two
clients, both get configured); it is NOT "pick one + fallback". When NO supported client is
detected we write NOTHING (a config file at a guessed path = installed-but-the-tool-is-invisible);
the caller informs the user and the install still completes. A user who installs a client later
re-runs ``--configure`` (Start Menu shortcut) to pick it up.

It merges (never clobbers) existing config: other ``mcpServers`` and top-level keys are kept.
It deliberately does NOT write ``ONENOTE_MCP_LOG_LEVEL`` — diagnostic logging stays OFF by
default (§7); a user adds that to the entry's ``env`` block when debugging. (Antigravity has a
known day-one bug where global-MCP ``env`` passing is unreliable; if logging is ever needed there,
prefer the ``ONENOTE_MCP_LOG_FILE`` default path over an ``env`` var — VM-pending, SPEC §8.)

Pure path/JSON logic — host-testable by pointing APPDATA/LOCALAPPDATA/USERPROFILE at temp dirs.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

SERVER_NAME = "onenote"
_CLAUDE_CONFIG_FILENAME = "claude_desktop_config.json"
_ANTIGRAVITY_CONFIG_FILENAME = "mcp_config.json"

# OneNote 15.0 type library + its Application coclass. Win32ComBackend binds this libid.
_ONENOTE_LIBID = "{0EA692EE-BB50-4E3C-AEF0-356D91732725}"
_ONENOTE_COCLSID = "{DC67E480-C3CB-49F8-8232-60B0C2056C8E}"


def _reg_default(root: int, subkey: str) -> str | None:
    """Read a registry key's default value from the 64-bit view, or None. Guarded."""
    try:
        import winreg  # noqa: PLC0415 — Windows-only; configure.py must import on Linux too

        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as k:
            value, _ = winreg.QueryValueEx(k, "")
            return value or None
    except (OSError, ImportError):
        return None


def repair_onenote_typelib() -> list[str]:
    """Shim any broken OneNote typelib version subkey under HKCU (no admin). Returns notes.

    FIELD BUG (VM-reproduced 2026-06-12): a stale version subkey under the OneNote libid with NO
    ``0\\win32`` or ``0\\win64`` mapping (a PIA-only leftover, e.g. ``...\\1.0``) poisons
    ``LoadRegTypeLib`` for the whole libid, so the COM call fails with TYPE_E_LIBNOTREGISTERED
    (0x8002801D) when OneNote is launched fresh by CoCreateInstance. We repair it WITHOUT admin by
    writing the missing mapping under ``HKCU\\Software\\Classes`` (which overrides HKLM in the
    merged HKCR view), pointing the broken version at the same typelib file a healthy version
    uses. Per-user and reversible. No-op off Windows or when the registration is healthy; fully
    guarded — a repair failure must never break --configure's real job (registering Claude)."""
    try:
        import winreg  # noqa: PLC0415
    except ImportError:
        return []

    notes: list[str] = []
    try:
        base = rf"TypeLib\{_ONENOTE_LIBID}"
        try:
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT, base, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY
            ) as root_key:
                versions, i = [], 0
                while True:
                    try:
                        versions.append(winreg.EnumKey(root_key, i))
                    except OSError:
                        break
                    i += 1
        except OSError:
            return []  # OneNote typelib not registered at all — nothing for us to repair

        good_win32 = good_win64 = None
        broken: list[str] = []
        for ver in versions:
            w32 = _reg_default(winreg.HKEY_CLASSES_ROOT, rf"{base}\{ver}\0\win32")
            w64 = _reg_default(winreg.HKEY_CLASSES_ROOT, rf"{base}\{ver}\0\win64")
            good_win32 = good_win32 or w32
            good_win64 = good_win64 or w64
            if not w32 and not w64:
                broken.append(ver)

        if not broken:
            return []

        # Fall back to deriving the typelib file from the coclass LocalServer32 (+ resource \3,
        # OneNote's typelib resource) when no healthy version exists to copy from.
        if not good_win32 and not good_win64:
            server = _reg_default(
                winreg.HKEY_CLASSES_ROOT, rf"CLSID\{_ONENOTE_COCLSID}\LocalServer32"
            )
            if server:
                good_win32 = server.strip('"').rstrip("\\") + r"\3"

        if not good_win32 and not good_win64:
            notes.append(
                f"OneNote typelib version(s) {broken} are broken (no win32/win64 mapping) but no "
                "healthy mapping was found to shim — needs a manual fix (delete the stale subkey)."
            )
            return notes

        for ver in broken:
            for platform, path in (("win32", good_win32), ("win64", good_win64)):
                if not path:
                    continue
                try:
                    with winreg.CreateKeyEx(
                        winreg.HKEY_CURRENT_USER,
                        rf"SOFTWARE\Classes\{base}\{ver}\0\{platform}",
                        0,
                        winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY,
                    ) as wk:
                        winreg.SetValueEx(wk, "", 0, winreg.REG_SZ, path)
                    notes.append(
                        rf"repaired OneNote typelib {_ONENOTE_LIBID}\{ver}: per-user HKCU shim "
                        rf"{platform} -> {path}"
                    )
                except OSError as exc:
                    notes.append(rf"could not shim typelib {ver}\{platform}: {exc}")
    except Exception as exc:  # noqa: BLE001 — repair must never break --configure
        notes.append(f"OneNote typelib repair skipped (unexpected error: {exc!r})")
    return notes


def _looks_like_claude(package_name: str) -> bool:
    low = package_name.lower()
    return "claude" in low or "anthropicclaude" in low or "anthropic.claude" in low


def _claude_regular_installed_via_registry() -> bool:
    """True if a per-user Claude Desktop uninstall entry exists in HKCU. Guarded; False off
    Windows or on any error (the filesystem program-dir check is the primary signal)."""
    try:
        import winreg  # noqa: PLC0415 — Windows-only; configure.py must import on Linux too
    except ImportError:
        return False
    try:
        sub = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub) as root:
            i = 0
            while True:
                try:
                    name = winreg.EnumKey(root, i)
                except OSError:
                    break
                i += 1
                if "claude" in name.lower():
                    return True
                try:
                    with winreg.OpenKey(root, name) as k:
                        display, _ = winreg.QueryValueEx(k, "DisplayName")
                        if "claude" in str(display).lower():
                            return True
                except OSError:
                    continue
    except OSError:
        return False
    return False


def _claude_regular_config_path(env: dict[str, str]) -> Path | None:
    """Config path for the regular (non-Store) Claude Desktop IF it is installed, else None.

    Detection looks at the INSTALL — the program dir ``%LOCALAPPDATA%\\Programs\\Claude`` or an
    Uninstall registry key (SPEC §8: not the config file, which is absent before first launch).
    A macOS app bundle / support dir is a dev-only convenience."""
    appdata = env.get("APPDATA")
    if appdata:
        # Windows context: decide purely on Windows install signals (never fall through to the
        # macOS convenience, so a not-installed Windows profile yields None).
        localappdata = env.get("LOCALAPPDATA")
        installed = (
            bool(localappdata and (Path(localappdata) / "Programs" / "Claude").is_dir())
            or _claude_regular_installed_via_registry()
        )
        return Path(appdata) / "Claude" / _CLAUDE_CONFIG_FILENAME if installed else None

    # macOS dev convenience (the production target is Windows; APPDATA is unset off Windows)
    mac = Path.home() / "Library" / "Application Support" / "Claude" / _CLAUDE_CONFIG_FILENAME
    if Path("/Applications/Claude.app").exists() or mac.parent.is_dir():
        return mac
    return None


def claude_config_paths(environ: dict[str, str] | None = None) -> list[Path]:
    """Detected Claude Desktop config-file paths (regular if installed + each Store package).

    Every path corresponds to an INSTALLED Claude variant: the regular path only when the regular
    install is detected, and one path per ``*Claude*`` package dir found under
    ``%LOCALAPPDATA%\\Packages``. ``environ`` overrides ``os.environ`` (for tests)."""
    env = os.environ if environ is None else environ
    paths: list[Path] = []

    regular = _claude_regular_config_path(env)
    if regular:
        paths.append(regular)

    localappdata = env.get("LOCALAPPDATA")
    if localappdata:
        packages = Path(localappdata) / "Packages"
        if packages.is_dir():
            for pkg in sorted(packages.iterdir()):
                if pkg.is_dir() and _looks_like_claude(pkg.name):
                    paths.append(
                        pkg / "LocalCache" / "Roaming" / "Claude" / _CLAUDE_CONFIG_FILENAME
                    )

    # de-dupe, preserve order
    seen: set[Path] = set()
    return [p for p in paths if not (p in seen or seen.add(p))]


def _antigravity_home(env: dict[str, str]) -> Path:
    """The user profile dir that holds Antigravity's ``.gemini`` config tree."""
    userprofile = env.get("USERPROFILE")
    return Path(userprofile) if userprofile else Path.home()


def antigravity_installed(env: dict[str, str]) -> bool:
    """Antigravity install detection. CONFIRMED on a real install 2026-06-16: it lands in
    ``%LOCALAPPDATA%\\Programs\\antigravity`` (lowercase) and ships NO ``agy`` CLI on PATH. We check
    that program dir (both casings — Windows is case-insensitive, but the suite runs on a
    case-sensitive host), then fall back to an ``agy``/``antigravity`` CLI on PATH or a
    ``~/.gemini`` dir (weaker signals, kept for forward-compat)."""
    localappdata = env.get("LOCALAPPDATA")
    if localappdata:
        programs = Path(localappdata) / "Programs"
        if (programs / "antigravity").is_dir() or (programs / "Antigravity").is_dir():
            return True
    if shutil.which("agy") or shutil.which("antigravity"):
        return True
    return (_antigravity_home(env) / ".gemini").is_dir()


def antigravity_config_paths(environ: dict[str, str] | None = None) -> list[Path]:
    """The Antigravity config path (single ``.gemini/config/mcp_config.json``) IF Antigravity is
    detected, else empty. One file covers both the CLI and the IDE (SPEC §8)."""
    env = os.environ if environ is None else environ
    if not antigravity_installed(env):
        return []
    return [_antigravity_home(env) / ".gemini" / "config" / _ANTIGRAVITY_CONFIG_FILENAME]


def mcp_client_targets(environ: dict[str, str] | None = None) -> list[Path]:
    """Config-file paths for EVERY detected supported MCP client (Claude Desktop variants +
    Antigravity), de-duplicated. Empty when no supported client is installed."""
    env = os.environ if environ is None else environ
    targets = claude_config_paths(env) + antigravity_config_paths(env)
    seen: set[Path] = set()
    return [p for p in targets if not (p in seen or seen.add(p))]


def server_command() -> dict[str, object]:
    """The ``command``/``args`` a client should launch this server with (same for every client).

    Frozen (PyInstaller exe): the exe itself, no args. Source/dev: the Python interpreter
    running ``-m onenote_com_mcp``."""
    if getattr(sys, "frozen", False):
        return {"command": sys.executable}
    return {"command": sys.executable, "args": ["-m", "onenote_com_mcp"]}


def merge_server_entry(config: dict, entry: dict, name: str = SERVER_NAME) -> dict:
    """Return ``config`` with ``mcpServers[name] = entry``, preserving everything else."""
    merged = dict(config)
    servers = dict(merged.get("mcpServers") or {})
    servers[name] = entry
    merged["mcpServers"] = servers
    return merged


def _load(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def configure_mcp_clients(
    environ: dict[str, str] | None = None, entry: dict | None = None
) -> list[Path]:
    """Write/merge the ``onenote`` server entry into EVERY detected client config (Claude Desktop
    variants + Antigravity). Returns the paths written — possibly EMPTY when no supported client is
    detected (SPEC §8: write nothing rather than a guessed, invisible config; the caller informs
    the user and the install still completes)."""
    entry = entry if entry is not None else server_command()
    written: list[Path] = []
    for path in mcp_client_targets(environ):
        merged = merge_server_entry(_load(path), entry)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(merged, indent=2, ensure_ascii=False), encoding="utf-8")
        written.append(path)
    return written
