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
    "get_page_info",
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
    "copy_pages",
    "copy_page_subtree",
    "copy_section",
    "restructure_section",
    "reposition_page",
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


def test_catalog_is_the_28_expected_tools(tools):
    assert set(tools) == EXPECTED_TOOLS
    assert len(tools) == 28


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
    # get_page (full content) vs get_page_info (lightweight object inventory) name each other
    assert "get_page_info" in tools["get_page"].description
    assert "get_page" in tools["get_page_info"].description
    # file discovery (get_page_info) vs file-extraction precheck (get_page_files_info): the
    # precheck points back to the inventory for plain discovery
    assert "get_page_info" in tools["get_page_files_info"].description
    assert "delete_page_content" in tools["delete_node"].description  # node vs in-page
    # the inline-vs-page-level delete border names both neighbours
    assert "delete_page_content" in tools["delete_inline_content"].description
    assert "delete_inline_content" in tools["delete_page_content"].description
    # the one-page-vs-many-pages reorder border names each other
    assert "reposition_page" in tools["restructure_section"].description
    assert "restructure_section" in tools["reposition_page"].description
    # copy_page can place the copy directly (after_page_id) — the stuck-workflow fix
    assert "after_page_id" in tools["copy_page"].inputSchema["properties"]
    # the single-vs-many-vs-subtree copy border: copy_page names the multi-page tools, and the
    # multi-page tools name each other (so a "copy these pages / this page + subpages" ask routes
    # away from a scatter-prone string of copy_page calls)
    assert "copy_pages" in tools["copy_page"].description
    assert "copy_page_subtree" in tools["copy_page"].description
    assert "copy_page_subtree" in tools["copy_pages"].description
    assert "copy_pages" in tools["copy_page_subtree"].description
    assert "after_page_id" in tools["copy_pages"].inputSchema["properties"]
    assert "after_page_id" in tools["copy_page_subtree"].inputSchema["properties"]


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
