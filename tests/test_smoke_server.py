"""Phase 6 smoke test: the MCP catalog is complete and well-formed (host, Tier 1).

A cheap structural sanity sweep over the assembled server — every tool present and described,
the destructive ones flagged, the §4 mode enum and server instructions in place. Catches a
half-wired tool or a dropped description without needing a VM. (The guarded-import invariant
has its own smoke test, test_smoke_import.py.)
"""

from __future__ import annotations

import asyncio

import pytest

from onenote_com_mcp.server import mcp

EXPECTED_TOOLS = {
    "list_notebooks",
    "list_sections",
    "list_pages",
    "search_pages",
    "get_page",
    "get_page_images",
    "get_page_files_info",
    "get_page_files",
    "get_current_context",
    "create_section",
    "create_page",
    "update_page_content",
    "create_table",
    "modify_table",
    "insert_image",
    "copy_page",
    "copy_section",
    "restructure_section",
    "reorder_sections",
    "rename_node",
    "move_page",
    "delete_node",
    "delete_page_content",
    "delete_inline_content",
}
DESTRUCTIVE_TOOLS = {
    "delete_node",
    "delete_page_content",
    "delete_inline_content",
    "modify_table",
}


@pytest.fixture(scope="module")
def tools():
    return {t.name: t for t in asyncio.run(mcp.list_tools())}


def test_catalog_is_the_24_expected_tools(tools):
    assert set(tools) == EXPECTED_TOOLS
    assert len(tools) == 24


def test_every_tool_has_a_real_description(tools):
    thin = {name: t.description for name, t in tools.items() if not (t.description or "").strip()}
    assert not thin, f"tools missing a description: {sorted(thin)}"
    # the §4 pass made descriptions substantive, not one-liners
    assert all(len(t.description) >= 40 for t in tools.values())


def test_destructive_tools_are_flagged(tools):
    for name in DESTRUCTIVE_TOOLS:
        assert "DESTRUCTIVE" in tools[name].description, f"{name} must be marked DESTRUCTIVE"


def test_contrastive_borders_present(tools):
    # a spot-check that the confusable pairs name each other (the §4 borders)
    assert "move_page" in tools["copy_page"].description  # copy vs move
    assert "restructure_section" in tools["move_page"].description  # move vs reorder-within
    assert "get_page_files" in tools["get_page_images"].description  # image vs file reads
    assert "delete_page_content" in tools["delete_node"].description  # node vs in-page
    # the inline-vs-page-level delete border names both neighbours
    assert "delete_page_content" in tools["delete_inline_content"].description
    assert "delete_inline_content" in tools["delete_page_content"].description


def test_update_mode_is_a_per_value_enum(tools):
    mode = tools["update_page_content"].inputSchema["properties"]["mode"]
    assert mode.get("enum") == ["append", "insert_before", "insert_after", "replace"]


def test_server_instructions_present():
    assert mcp.instructions and "live" in mcp.instructions.lower()
    # the cross-tool rules SPEC §4 asked for
    assert "objectID" in mcp.instructions
    assert "copy_section" in mcp.instructions


def test_selftest_reports_failure_cleanly(monkeypatch, capsys):
    """--selftest is the frozen-install health check (it binds COM on Windows). When the
    backend can't be reached it must exit non-zero with a clear message, not a traceback."""
    import sys

    from onenote_com_mcp import server

    def _boom():
        raise RuntimeError("no backend here")

    monkeypatch.setattr(server, "get_backend", _boom)
    monkeypatch.setattr(sys, "argv", ["onenote-com-mcp", "--selftest"])
    with pytest.raises(SystemExit) as excinfo:
        server.main()
    assert excinfo.value.code == 1
    assert "SELFTEST FAIL" in capsys.readouterr().err
