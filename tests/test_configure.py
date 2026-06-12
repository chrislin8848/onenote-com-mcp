"""Phase 6 Stage 5: --configure logic (SPEC §8). Host-testable by pointing APPDATA /
LOCALAPPDATA at temp dirs — pure path detection + JSON merge, no freeze needed.
"""

from __future__ import annotations

import json

import pytest

from onenote_com_mcp import configure


@pytest.fixture
def fake_env(tmp_path):
    appdata = tmp_path / "AppData" / "Roaming"
    localappdata = tmp_path / "AppData" / "Local"
    appdata.mkdir(parents=True)
    localappdata.mkdir(parents=True)
    return {"APPDATA": str(appdata), "LOCALAPPDATA": str(localappdata)}


# --- path detection ----------------------------------------------------------------------


def test_detects_regular_appdata_path(fake_env):
    paths = configure.claude_config_paths(fake_env)
    regular = next(p for p in paths if "Roaming" in str(p) and "Packages" not in str(p))
    assert regular.name == "claude_desktop_config.json"
    assert regular.parent.name == "Claude"


def test_detects_store_package_by_glob(fake_env):
    # a Store/MSIX Claude package directory — the family name is detected, not hardcoded
    pkg = (
        configure.Path(fake_env["LOCALAPPDATA"])
        / "Packages"
        / "AnthropicClaude_abc123def"
        / "LocalCache"
        / "Roaming"
        / "Claude"
    )
    pkg.mkdir(parents=True)
    paths = configure.claude_config_paths(fake_env)
    store = [p for p in paths if "Packages" in str(p)]
    assert len(store) == 1
    assert "AnthropicClaude_abc123def" in str(store[0])


def test_ignores_unrelated_packages(fake_env):
    for name in ("Microsoft.WindowsCalculator_x", "SomeOther_y"):
        (configure.Path(fake_env["LOCALAPPDATA"]) / "Packages" / name).mkdir(parents=True)
    assert not [p for p in configure.claude_config_paths(fake_env) if "Packages" in str(p)]


# --- merge semantics ---------------------------------------------------------------------


def test_merge_preserves_other_servers_and_keys():
    existing = {
        "mcpServers": {"other": {"command": "x"}},
        "globalShortcut": "Ctrl+Q",
    }
    merged = configure.merge_server_entry(existing, {"command": "onenote.exe"})
    assert merged["mcpServers"]["other"] == {"command": "x"}  # untouched
    assert merged["mcpServers"]["onenote"] == {"command": "onenote.exe"}
    assert merged["globalShortcut"] == "Ctrl+Q"  # unrelated top-level key kept


def test_server_command_source_mode():
    cmd = configure.server_command()  # not frozen under pytest
    assert (
        cmd["command"].endswith(("python", "python.exe", "python3")) or "python" in cmd["command"]
    )
    assert cmd["args"] == ["-m", "onenote_com_mcp"]


# --- end-to-end write --------------------------------------------------------------------


def test_configure_writes_merged_config_to_all_targets(fake_env):
    # one Store package + the regular path
    (
        configure.Path(fake_env["LOCALAPPDATA"])
        / "Packages"
        / "AnthropicClaude_xyz"
        / "LocalCache"
        / "Roaming"
        / "Claude"
    ).mkdir(parents=True)

    written = configure.configure_claude_desktop(fake_env, entry={"command": "onenote.exe"})

    assert len(written) >= 2  # regular + store
    for path in written:
        data = json.loads(path.read_text(encoding="utf-8"))
        assert data["mcpServers"]["onenote"] == {"command": "onenote.exe"}


def test_configure_does_not_write_log_level(fake_env):
    [
        *_,
    ] = configure.configure_claude_desktop(fake_env, entry={"command": "onenote.exe"})
    for path in configure.claude_config_paths(fake_env):
        if path.exists():
            text = path.read_text(encoding="utf-8")
            assert "ONENOTE_MCP_LOG_LEVEL" not in text  # logging stays OFF by default (§7)


def test_configure_is_idempotent(fake_env):
    configure.configure_claude_desktop(fake_env, entry={"command": "onenote.exe"})
    configure.configure_claude_desktop(fake_env, entry={"command": "onenote.exe"})
    regular = next(p for p in configure.claude_config_paths(fake_env) if "Packages" not in str(p))
    servers = json.loads(regular.read_text(encoding="utf-8"))["mcpServers"]
    assert list(servers).count("onenote") == 1  # no duplication


def test_configure_merges_into_existing_file(fake_env):
    regular = next(p for p in configure.claude_config_paths(fake_env) if "Packages" not in str(p))
    regular.parent.mkdir(parents=True, exist_ok=True)
    regular.write_text(json.dumps({"mcpServers": {"keep": {"command": "k"}}}), encoding="utf-8")

    configure.configure_claude_desktop(fake_env, entry={"command": "onenote.exe"})

    servers = json.loads(regular.read_text(encoding="utf-8"))["mcpServers"]
    assert servers["keep"] == {"command": "k"}
    assert servers["onenote"] == {"command": "onenote.exe"}
