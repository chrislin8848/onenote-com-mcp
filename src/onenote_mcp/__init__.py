"""COM-only OneNote MCP server.

All OneNote access goes through ``onenote_mcp.backend.OneNoteBackend`` (pywin32 COM on
Windows; fixture replay on Linux). There is intentionally **no** Microsoft Graph, ``msal``,
HTTP, or Azure/token code anywhere in this package (SPEC §9 red line).
"""

__version__ = "0.0.1"
