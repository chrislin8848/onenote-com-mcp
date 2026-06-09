"""Smoke tests: the package imports cleanly on Linux, and the guarded-import invariant holds.

This is the load-bearing Phase 0 test (SPEC §2.1): importing the COM backend module must NOT
pull in ``win32com`` / ``pywintypes`` / ``pythoncom``. If it ever does, the whole host dev loop
breaks. We assert those modules are absent from ``sys.modules`` after import.
"""

from __future__ import annotations

import sys


def test_package_imports():
    import onenote_com_mcp  # noqa: F401
    import onenote_com_mcp.enums  # noqa: F401
    import onenote_com_mcp.errors  # noqa: F401
    import onenote_com_mcp.server  # noqa: F401
    from onenote_com_mcp.backend import FixtureBackend, OneNoteBackend, get_backend  # noqa: F401


def test_win32_backend_module_imports_without_pywin32():
    # Importing the module must be safe on Linux...
    import onenote_com_mcp.backend.win32com_backend as w  # noqa: F401

    # ...and must not have imported any pywin32 component at module load time.
    for mod in ("win32com", "win32com.client", "pywintypes", "pythoncom"):
        assert mod not in sys.modules, f"guarded-import invariant violated: {mod} was imported"


def test_full_tool_catalog_registered():
    # The MCP surface should expose all 17 SPEC §4 tools even while bodies are stubbed.
    import asyncio

    from onenote_com_mcp.server import mcp

    tools = asyncio.run(mcp.list_tools())
    names = {t.name for t in tools}
    expected = {
        "list_notebooks",
        "list_sections",
        "list_pages",
        "search_pages",
        "get_page",
        "get_page_images",
        "create_notebook",
        "create_section",
        "create_page",
        "update_page_content",
        "create_table",
        "insert_image",
        "copy_page",
        "copy_section",
        "copy_notebook",
        "delete_node",
        "delete_page_content",
    }
    assert expected <= names
