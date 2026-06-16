"""Phase 6 Stage 5 + SPEC §8 (0616) --configure logic: multi-client detection + JSON merge.

Host-testable by pointing APPDATA / LOCALAPPDATA / USERPROFILE at temp dirs — pure path detection
+ JSON merge, no freeze needed. Three supported clients, each detected by INSTALL presence (not by
its config file, which is absent before first launch): Claude Desktop regular, Claude Desktop Store
(MSIX), and Antigravity (CLI/IDE share one ``.gemini`` config).
"""

from __future__ import annotations

import json

import pytest

from onenote_com_mcp import configure


@pytest.fixture
def fake_env(tmp_path):
    appdata = tmp_path / "AppData" / "Roaming"
    localappdata = tmp_path / "AppData" / "Local"
    userprofile = tmp_path / "Profile"
    for d in (appdata, localappdata, userprofile):
        d.mkdir(parents=True)
    return {
        "APPDATA": str(appdata),
        "LOCALAPPDATA": str(localappdata),
        "USERPROFILE": str(userprofile),
    }


# --- install helpers (create the INSTALL markers, never the config file) -----------------


def _install_claude_regular(env):
    (configure.Path(env["LOCALAPPDATA"]) / "Programs" / "Claude").mkdir(parents=True)


def _install_claude_store(env, family="AnthropicClaude_abc123def"):
    (
        configure.Path(env["LOCALAPPDATA"])
        / "Packages"
        / family
        / "LocalCache"
        / "Roaming"
        / "Claude"
    ).mkdir(parents=True)


def _install_antigravity(env):
    # real install dir, confirmed 2026-06-16: %LOCALAPPDATA%\Programs\antigravity (lowercase)
    (configure.Path(env["LOCALAPPDATA"]) / "Programs" / "antigravity").mkdir(parents=True)


# --- Claude path detection ---------------------------------------------------------------


def test_claude_regular_detected_when_installed(fake_env):
    _install_claude_regular(fake_env)
    paths = configure.claude_config_paths(fake_env)
    regular = next(p for p in paths if "Roaming" in str(p) and "Packages" not in str(p))
    assert regular.name == "claude_desktop_config.json"
    assert regular.parent.name == "Claude"


def test_claude_regular_NOT_detected_without_install(fake_env):
    # SPEC §8: detect the INSTALL, not the config file. No program dir / registry → no regular path.
    assert not [p for p in configure.claude_config_paths(fake_env) if "Packages" not in str(p)]


def test_detects_store_package_by_glob(fake_env):
    # a Store/MSIX Claude package directory — the family name is detected, not hardcoded
    _install_claude_store(fake_env)
    paths = configure.claude_config_paths(fake_env)
    store = [p for p in paths if "Packages" in str(p)]
    assert len(store) == 1
    assert "AnthropicClaude_abc123def" in str(store[0])


def test_ignores_unrelated_packages(fake_env):
    for name in ("Microsoft.WindowsCalculator_x", "SomeOther_y"):
        (configure.Path(fake_env["LOCALAPPDATA"]) / "Packages" / name).mkdir(parents=True)
    assert not [p for p in configure.claude_config_paths(fake_env) if "Packages" in str(p)]


# --- Antigravity detection ---------------------------------------------------------------


def test_antigravity_detected_via_program_dir(fake_env):
    _install_antigravity(fake_env)
    paths = configure.antigravity_config_paths(fake_env)
    assert len(paths) == 1
    p = paths[0]
    assert p.name == "mcp_config.json"
    assert p.parent.name == "config"
    assert p.parent.parent.name == ".gemini"
    # one .gemini config under the user profile covers BOTH the CLI and the IDE
    assert str(p).startswith(fake_env["USERPROFILE"])


def test_antigravity_detected_via_gemini_dir(fake_env):
    # the .gemini dir alone (created on first run) is a best-effort install signal
    (configure.Path(fake_env["USERPROFILE"]) / ".gemini").mkdir(parents=True)
    assert configure.antigravity_installed(fake_env)
    assert len(configure.antigravity_config_paths(fake_env)) == 1


def test_antigravity_not_detected(fake_env):
    # no program dir, no .gemini dir, and no agy/antigravity CLI on the test machine's PATH
    assert not configure.antigravity_installed(fake_env)
    assert configure.antigravity_config_paths(fake_env) == []


# --- combined target detection -----------------------------------------------------------


def test_mcp_client_targets_aggregates_every_installed_client(fake_env):
    _install_claude_regular(fake_env)
    _install_claude_store(fake_env)
    _install_antigravity(fake_env)
    targets = configure.mcp_client_targets(fake_env)
    assert any("Packages" not in str(t) and "claude_desktop" in str(t) for t in targets)  # regular
    assert any("Packages" in str(t) for t in targets)  # store
    assert any("mcp_config.json" in str(t) for t in targets)  # antigravity


def test_mcp_client_targets_empty_when_no_client(fake_env):
    assert configure.mcp_client_targets(fake_env) == []


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


def test_configure_writes_merged_config_to_all_clients(fake_env):
    # Claude regular + Claude Store + Antigravity all installed → all three get the entry.
    _install_claude_regular(fake_env)
    _install_claude_store(fake_env, family="AnthropicClaude_xyz")
    _install_antigravity(fake_env)

    written = configure.configure_mcp_clients(fake_env, entry={"command": "onenote.exe"})

    assert len(written) == 3  # regular + store + antigravity
    for path in written:
        data = json.loads(path.read_text(encoding="utf-8"))
        assert data["mcpServers"]["onenote"] == {"command": "onenote.exe"}


def test_configure_writes_nothing_when_no_client(fake_env):
    # SPEC §8: no supported client → write NOTHING (no guessed, invisible config file).
    written = configure.configure_mcp_clients(fake_env, entry={"command": "onenote.exe"})
    assert written == []
    # and nothing was created on disk under the would-be Claude path
    assert not (configure.Path(fake_env["APPDATA"]) / "Claude").exists()


def test_repair_onenote_typelib_is_guarded_on_host():
    # On Linux/macOS (no winreg) the repair must no-op cleanly: return a list, never raise.
    # The real registry behavior is VM-validated (Tier 2), not host-testable.
    result = configure.repair_onenote_typelib()
    assert isinstance(result, list)


def test_configure_does_not_write_log_level(fake_env):
    _install_antigravity(fake_env)
    written = configure.configure_mcp_clients(fake_env, entry={"command": "onenote.exe"})
    assert written
    for path in written:
        text = path.read_text(encoding="utf-8")
        assert "ONENOTE_MCP_LOG_LEVEL" not in text  # logging stays OFF by default (§7)


def test_configure_is_idempotent(fake_env):
    _install_claude_regular(fake_env)
    configure.configure_mcp_clients(fake_env, entry={"command": "onenote.exe"})
    configure.configure_mcp_clients(fake_env, entry={"command": "onenote.exe"})
    regular = next(p for p in configure.claude_config_paths(fake_env) if "Packages" not in str(p))
    servers = json.loads(regular.read_text(encoding="utf-8"))["mcpServers"]
    assert list(servers).count("onenote") == 1  # no duplication


def test_configure_merges_into_existing_antigravity_file(fake_env):
    _install_antigravity(fake_env)
    target = configure.antigravity_config_paths(fake_env)[0]
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps({"mcpServers": {"keep": {"command": "k"}}}), encoding="utf-8")

    configure.configure_mcp_clients(fake_env, entry={"command": "onenote.exe"})

    servers = json.loads(target.read_text(encoding="utf-8"))["mcpServers"]
    assert servers["keep"] == {"command": "k"}  # other connector preserved
    assert servers["onenote"] == {"command": "onenote.exe"}
