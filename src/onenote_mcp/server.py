"""FastMCP server — the full OneNote tool catalog (SPEC §4).

This module wires every tool's *name, signature, and description* (the MCP contract the LLM
sees) up front. Bodies raise ``NotImplementedError`` tagged with the phase that fills them in,
so the surface is reviewable now and lights up phase by phase. The shared write core and the
copy core live in ``onenote_mcp.service`` — these tools stay thin facades (SPEC §4).

stdio transport: nothing but MCP protocol may go to stdout. Logs go to stderr (SPEC §8).
"""

from __future__ import annotations

import sys

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("onenote")

# --- Read (Phase 2: wired to FixtureBackend on Linux) -----------------------


@mcp.tool()
def list_notebooks() -> str:
    """List all open OneNote notebooks (name + ID)."""
    raise NotImplementedError("Phase 2")


@mcp.tool()
def list_sections(notebook_id: str) -> str:
    """List sections in a notebook (name + ID)."""
    raise NotImplementedError("Phase 2")


@mcp.tool()
def list_pages(section_id: str) -> str:
    """List pages in a section, including each page's subpage level (pageLevel)."""
    raise NotImplementedError("Phase 2")


@mcp.tool()
def search_pages(query: str, scope_id: str = "") -> str:
    """Full-text search for pages. Scope to a notebook/section ID (recommended)."""
    raise NotImplementedError("Phase 2")


@mcp.tool()
def get_page(page_id: str) -> str:
    """Read a page preserving rich text formatting, structure, tables, and object IDs."""
    raise NotImplementedError("Phase 2")


@mcp.tool()
def get_page_images(page_id: str) -> str:
    """Return a page's images (binary) as viewable image content."""
    raise NotImplementedError("Phase 2")


# --- Create (Phase 4) -------------------------------------------------------


@mcp.tool()
def create_notebook(name: str, path: str) -> str:
    """Create a notebook. ``path`` must be a synced (e.g. OneDrive) location, else the
    notebook is local-only. Prefer create_section inside an existing synced notebook."""
    raise NotImplementedError("Phase 4")


@mcp.tool()
def create_section(notebook_id: str, name: str) -> str:
    """Create a section in an existing notebook (inherits its sync)."""
    raise NotImplementedError("Phase 4")


@mcp.tool()
def create_page(section_id: str, title: str, content: str = "", page_level: int = 1) -> str:
    """Create a page in a section. ``page_level`` (1/2/3) sets subpage indent."""
    raise NotImplementedError("Phase 4")


# --- Modify (Phase 4: shared write core) ------------------------------------


@mcp.tool()
def update_page_content(page_id: str, content: str, mode: str = "append") -> str:
    """Modify page content. mode = append | insert | replace. Preserves untouched
    formatting (in-place tree edit). Concurrency-guarded; will not force-overwrite."""
    raise NotImplementedError("Phase 4")


@mcp.tool()
def create_table(page_id: str, rows: list[list[str]]) -> str:
    """Add a table (rows = list of rows of cell text) to a page."""
    raise NotImplementedError("Phase 4")


@mcp.tool()
def insert_image(page_id: str, image_base64: str, media_type: str) -> str:
    """Insert an image (base64) into a page."""
    raise NotImplementedError("Phase 4")


# --- Copy (Phase 5: raw-XML faithful transfer) ------------------------------


@mcp.tool()
def copy_page(page_id: str, target_section_id: str) -> str:
    """Faithfully copy a page (formatting, tables, inline images, pageLevel) to a section."""
    raise NotImplementedError("Phase 5")


@mcp.tool()
def copy_section(section_id: str, target_notebook_id: str) -> str:
    """Faithfully copy a whole section into a notebook."""
    raise NotImplementedError("Phase 5")


@mcp.tool()
def copy_notebook(notebook_id: str, name: str, path: str) -> str:
    """Faithfully copy a whole notebook (subject to create_notebook sync constraints)."""
    raise NotImplementedError("Phase 5")


# --- Delete (Phase 6: destructive — conservative) ---------------------------


@mcp.tool()
def delete_node(object_id: str) -> str:
    """DESTRUCTIVE. Delete a hierarchy node (notebook / section / page) to the recycle bin."""
    raise NotImplementedError("Phase 6")


@mcp.tool()
def delete_page_content(page_id: str, object_id: str) -> str:
    """DESTRUCTIVE. Delete one page content object (image / table / outline) by ID."""
    raise NotImplementedError("Phase 6")


def main() -> None:
    """Console entry point. stdio transport; logs must stay on stderr."""
    print("onenote-mcp starting (stdio)", file=sys.stderr)
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
