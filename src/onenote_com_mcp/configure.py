"""``--configure``: register this server in Claude Desktop's config (SPEC §8).

The installer (and a user re-running the exe with ``--configure``) calls this to add an
``onenote`` entry to Claude Desktop's ``claude_desktop_config.json``, pointing at THIS
executable. It handles BOTH Claude Desktop install flavors independently (SPEC §8):

- **Regular (per-user install):** ``%APPDATA%\\Claude\\claude_desktop_config.json``
- **Microsoft Store (MSIX):** ``%LOCALAPPDATA%\\Packages\\<Claude package>\\LocalCache\\Roaming
  \\Claude\\claude_desktop_config.json`` — the package family name is DETECTED by globbing (its
  exact value varies), never hardcoded.

It merges (never clobbers) existing config: other ``mcpServers`` and top-level keys are kept.
It deliberately does NOT write ``ONENOTE_MCP_LOG_LEVEL`` — diagnostic logging stays OFF by
default (§7); a user adds that to the entry's ``env`` block when debugging.

Pure path/JSON logic — host-testable by pointing APPDATA/LOCALAPPDATA at temp dirs.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

SERVER_NAME = "onenote"
_CONFIG_FILENAME = "claude_desktop_config.json"

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


def claude_config_paths(environ: dict[str, str] | None = None) -> list[Path]:
    """Detected Claude Desktop config-file paths (regular + any Store packages), de-duplicated.

    A path is included when its Claude config DIRECTORY can exist for us to write into — the
    regular ``%APPDATA%\\Claude`` whenever APPDATA is set, and each Store package dir found by
    glob. ``environ`` overrides ``os.environ`` (for tests)."""
    env = os.environ if environ is None else environ
    paths: list[Path] = []

    appdata = env.get("APPDATA")
    if appdata:
        paths.append(Path(appdata) / "Claude" / _CONFIG_FILENAME)

    localappdata = env.get("LOCALAPPDATA")
    if localappdata:
        packages = Path(localappdata) / "Packages"
        if packages.is_dir():
            for pkg in sorted(packages.iterdir()):
                if pkg.is_dir() and _looks_like_claude(pkg.name):
                    paths.append(pkg / "LocalCache" / "Roaming" / "Claude" / _CONFIG_FILENAME)

    # macOS dev convenience (the production target is Windows)
    mac = Path.home() / "Library" / "Application Support" / "Claude" / _CONFIG_FILENAME
    if mac.parent.parent.is_dir() and mac not in paths:
        paths.append(mac)

    # de-dupe, preserve order
    seen: set[Path] = set()
    return [p for p in paths if not (p in seen or seen.add(p))]


def server_command() -> dict[str, object]:
    """The ``command``/``args`` Claude Desktop should launch this server with.

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


def configure_claude_desktop(
    environ: dict[str, str] | None = None, entry: dict | None = None
) -> list[Path]:
    """Write/merge the ``onenote`` server entry into every detected Claude config. Returns the
    paths written. If none are detected but APPDATA is set, the regular path is created as the
    best-guess fallback (the installer runs after Claude is installed)."""
    entry = entry if entry is not None else server_command()
    targets = claude_config_paths(environ)
    if not targets:
        env = os.environ if environ is None else environ
        appdata = env.get("APPDATA")
        if appdata:
            targets = [Path(appdata) / "Claude" / _CONFIG_FILENAME]

    written: list[Path] = []
    for path in targets:
        merged = merge_server_entry(_load(path), entry)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(merged, indent=2, ensure_ascii=False), encoding="utf-8")
        written.append(path)
    return written
