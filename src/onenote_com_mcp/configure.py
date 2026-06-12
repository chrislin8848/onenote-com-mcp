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
