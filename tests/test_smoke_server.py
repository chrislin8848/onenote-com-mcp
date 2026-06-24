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
    "get_table",
    "get_object",
    "find_objects",
    "get_page_images",
    "get_page_files_info",
    "get_page_files",
    "get_current_context",
    "create_section",
    "create_page",
    "update_page_content",
    "find_and_replace",
    "batch_update",
    "create_table",
    "modify_table",
    "insert_svg_image",
    "insert_image_from_path",
    "apply_text_style",
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


def test_catalog_is_the_35_expected_tools(tools):
    assert set(tools) == EXPECTED_TOOLS
    assert len(tools) == 35
    # Picture insertion: vector SVG via insert_svg_image, and raster (PNG/JPEG/GIF) via
    # insert_image_from_path which reads a LOCAL FILE off disk — so the bytes never go through the
    # model. The old base64-param insert_image (model emits the bytes) stays REMOVED.
    assert "insert_image" not in tools
    assert "insert_file" not in tools
    assert "insert_svg_image" in tools
    assert "insert_image_from_path" in tools


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
    # get_page_info warns its preview is a truncated label, not content to proofread from
    assert "preview" in tools["get_page_info"].description.lower()
    assert "get_object" in tools["get_page_info"].description
    # get_table (single-table compact read) borders get_page both ways
    assert "get_table" in tools["get_page"].description
    assert "get_page" in tools["get_table"].description
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
    # enumerate-vs-search border: search_pages (full-text, misses non-matches) points at list_pages
    # for listing a section's pages / a page's subpages, and list_pages names subpages
    assert "list_pages" in tools["search_pages"].description
    assert "subpage" in tools["list_pages"].description.lower()


def test_update_mode_is_a_per_value_enum(tools):
    mode = tools["update_page_content"].inputSchema["properties"]["mode"]
    assert mode.get("enum") == ["append", "insert_before", "insert_after", "replace"]


def test_modify_table_operation_enum_covers_reorder_and_set_column(tools):
    op = tools["modify_table"].inputSchema["properties"]["operation"]
    assert set(op.get("enum")) == {
        "insert_columns",
        "insert_rows",
        "set_rows",
        "set_column",
        "reorder_columns",
        "reorder_rows",
        "delete_columns",
        "delete_rows",
    }
    props = tools["modify_table"].inputSchema["properties"]
    assert "order" in props and "values" in props  # the reorder / single-column params


def test_insert_svg_image_can_be_positioned_mid_page(tools):
    mode = tools["insert_svg_image"].inputSchema["properties"]["mode"]
    assert mode.get("enum") == ["append", "insert_before", "insert_after"]


def test_apply_text_style_borders_update_page_content(tools):
    # the bulk style patch names the per-paragraph editor (and vice versa) so a "restyle the whole
    # page" ask routes away from a string of update_page_content("replace") calls
    desc = tools["apply_text_style"].description
    assert "update_page_content" in desc
    assert "font" in desc.lower()
    # 1.1.4: table row/column granularity — a whole ROW is a scope objectID; a whole COLUMN is the
    # 0-indexed columns param (a column has no objectID), and both cover text AND cell_shading
    assert "row_object_ids" in desc
    assert "columns" in desc


def test_server_instructions_present():
    assert mcp.instructions and "live" in mcp.instructions.lower()
    # the cross-tool rules SPEC §4 asked for
    assert "objectID" in mcp.instructions
    assert "copy_section" in mcp.instructions
    # picture insertion is SVG-only — the instructions must name insert_svg_image and tell the
    # user to add photos/raster images by hand
    assert "insert_svg_image" in mcp.instructions
    assert "by hand" in mcp.instructions.lower()
    # 1.0.10: editing is in-place-by-default (don't rebuild a page to change it), confirmation is
    # proportionate to risk (recycle-bin deletes are reversible), and copy-then-modify is the safe
    # big-edit pattern.
    assert "in place" in mcp.instructions.lower()
    assert "recycle bin" in mcp.instructions.lower()
    assert "copy-then-modify" in mcp.instructions.lower()
    # 1.1.0: the bulk restyle tool is surfaced, and the App Ctrl+A shortcut is deliberately NOT
    # mentioned (users already know it — Chris 2026-06-14)
    assert "apply_text_style" in mcp.instructions
    # 1.1.4: the instructions tell how to target a whole table ROW vs a whole COLUMN
    assert "row_object_ids" in mcp.instructions
    assert "columns=" in mcp.instructions
    assert "ctrl+a" not in mcp.instructions.lower()
    # 1.1.1: enumerate pages/subpages with list_pages, not search_pages (full-text, silently misses)
    assert "list_pages" in mcp.instructions
    assert "search_pages" in mcp.instructions
    # 1.2.0: table editing — rearrange with reorder_columns/reorder_rows (not clear+retype), rewrite
    # one column with set_column, and read one big table compactly with get_table
    assert "reorder_columns" in mcp.instructions
    assert "set_column" in mcp.instructions
    assert "get_table" in mcp.instructions
    # pacing: heavy copy/write calls go in small batches (the server also serializes them via a
    # process-wide lock), because OneNote's COM is single-threaded
    assert "single-threaded" in mcp.instructions.lower()
    assert "batch" in mcp.instructions.lower()
    # Group B (instruction-only discoverability): append targeting (target_object_id), clearing a
    # table body while keeping a header (set_rows None=keep), create-then-restyle for a default
    # font, and CJK byte-faithful storage + the \uXXXX-escape tip for uncertain glyphs.
    assert "target_object_id" in mcp.instructions
    assert "keeping a header" in mcp.instructions
    assert "hand-repeat" in mcp.instructions
    assert "byte-for-byte" in mcp.instructions
    assert "\\uXXXX" in mcp.instructions
    # 1.3.0: precise-editing tools surfaced (typo fixes without re-typing, locate-by-text, single
    # object read, atomic multi-edit, opt-in id echo) + raster insert from a LOCAL FILE path
    assert "find_and_replace" in mcp.instructions
    assert "find_objects" in mcp.instructions
    assert "get_object" in mcp.instructions
    assert "batch_update" in mcp.instructions
    assert "return_ids" in mcp.instructions
    assert "insert_image_from_path" in mcp.instructions
    # 1.3.1: get_page_info's preview is a truncated label, not content to proofread from
    assert "preview" in mcp.instructions.lower()


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
