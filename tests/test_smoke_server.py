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
    "insert_image",
    "copy_page",
    "copy_section",
    "restructure_section",
    "reorder_sections",
    "rename_node",
    "move_page",
    "delete_node",
    "delete_page_content",
}
DESTRUCTIVE_TOOLS = {"delete_node", "delete_page_content"}


@pytest.fixture(scope="module")
def tools():
    return {t.name: t for t in asyncio.run(mcp.list_tools())}


def test_catalog_is_the_22_expected_tools(tools):
    assert set(tools) == EXPECTED_TOOLS
    assert len(tools) == 22


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


def test_update_mode_is_a_per_value_enum(tools):
    mode = tools["update_page_content"].inputSchema["properties"]["mode"]
    assert mode.get("enum") == ["append", "insert_before", "insert_after", "replace"]


def test_server_instructions_present():
    assert mcp.instructions and "live" in mcp.instructions.lower()
    # the cross-tool rules SPEC §4 asked for
    assert "objectID" in mcp.instructions
    assert "copy_section" in mcp.instructions
