"""Smoke tests: the package imports cleanly on Linux, and the guarded-import invariant holds.

This is the load-bearing Phase 0 test (SPEC §2.1): importing the COM backend module must NOT
pull in ``win32com`` / ``pywintypes`` / ``pythoncom``. If it ever does, the whole host dev loop
breaks. We assert those modules are absent from ``sys.modules`` after import.
"""

from __future__ import annotations

import subprocess
import sys


def test_package_imports():
    import onenote_com_mcp  # noqa: F401
    import onenote_com_mcp.enums  # noqa: F401
    import onenote_com_mcp.errors  # noqa: F401
    import onenote_com_mcp.server  # noqa: F401
    from onenote_com_mcp.backend import FixtureBackend, OneNoteBackend, get_backend  # noqa: F401


def test_win32_backend_module_imports_without_pywin32():
    # Checked in a clean subprocess: an in-process sys.modules assertion false-positives on
    # Windows, where the mcp SDK itself (mcp.os.win32.utilities) legitimately imports
    # pywintypes once another test has imported onenote_com_mcp.server. Only what OUR backend
    # module pulls in at import time is the invariant. (pywin32's site-packages .pth bootstrap
    # modules are allowed — they don't load pywintypes.)
    code = (
        "import sys; "
        "import onenote_com_mcp.backend.win32com_backend; "
        "bad = [m for m in ('win32com', 'win32com.client', 'pywintypes', 'pythoncom') "
        "if m in sys.modules]; "
        "sys.exit('guarded-import invariant violated: ' + repr(bad) if bad else 0)"
    )
    res = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert res.returncode == 0, res.stderr.strip() or res.stdout.strip()


def test_full_tool_catalog_registered():
    # The MCP surface should expose all 22 SPEC §4 tools even while bodies are stubbed.
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
        "get_current_context",
        "create_notebook",
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
    assert expected <= names
