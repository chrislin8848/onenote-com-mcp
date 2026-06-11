"""FastMCP server — the full OneNote tool catalog (SPEC §4).

This module wires every tool's *name, signature, and description* (the MCP contract the LLM
sees) up front. Bodies raise ``NotImplementedError`` tagged with the phase that fills them in,
so the surface is reviewable now and lights up phase by phase. The shared write core and the
copy core live in ``onenote_com_mcp.service`` — these tools stay thin facades (SPEC §4).

stdio transport: nothing but MCP protocol may go to stdout. Logs go to stderr (SPEC §8).
"""

from __future__ import annotations

import base64
import json
import sys

from mcp.server.fastmcp import FastMCP, Image

from onenote_com_mcp.backend import get_backend
from onenote_com_mcp.service import copy, create, hierarchy_edit, page_edit, read

mcp = FastMCP("onenote")


def _json(data: object) -> str:
    # ensure_ascii=False keeps CJK note content readable in the tool result
    return json.dumps(data, ensure_ascii=False, indent=2)


# --- Read (Phase 2: wired to FixtureBackend on Linux) -----------------------


@mcp.tool()
def list_notebooks() -> str:
    """List all open OneNote notebooks (name + ID)."""
    return _json(read.list_notebooks(get_backend()))


@mcp.tool()
def list_sections(notebook_id: str) -> str:
    """List sections in a notebook (name + ID), preserving section-group nesting
    (one:SectionGroup containers appear as nested groups, not flattened)."""
    return _json(read.list_sections(get_backend(), notebook_id))


@mcp.tool()
def list_pages(section_id: str) -> str:
    """List pages in a section, including each page's subpage level (pageLevel)."""
    return _json(read.list_pages(get_backend(), section_id))


@mcp.tool()
def search_pages(query: str, scope_id: str = "") -> str:
    """Full-text search for pages. Scope to a notebook/section ID (recommended)."""
    return _json(read.search_pages(get_backend(), query, scope_id))


@mcp.tool()
def get_page(page_id: str) -> str:
    """Read a page preserving rich text formatting, structure, tables, and object IDs."""
    return _json(read.get_page(get_backend(), page_id))


# structured_output=False: the return is image content, not a JSON schema — FastMCP can't
# build a pydantic output schema for Image, and we don't want one here.
@mcp.tool(structured_output=False)
def get_page_images(page_id: str) -> list[Image]:
    """Return a page's images (binary) as viewable image content. Per-image metadata
    (object IDs, dimensions, OCR text) is on get_page; this returns the pixels so they can
    be recognized visually."""
    return [
        Image(data=base64.b64decode(img["data_base64"]), format=img["media_type"].split("/")[-1])
        for img in read.get_page_images(get_backend(), page_id)
    ]


@mcp.tool()
def get_current_context() -> str:
    """Where is the user right now? Returns the active OneNote window's current notebook /
    section group / section / page (IDs + names), so the user can say "this page" or
    "the current section". Granularity stops at the page — in-page cursor position and
    selected text are not available. Errors clearly if OneNote has no open window.
    Before acting on this context, report it back ("you're currently on page X") so the
    user can confirm they haven't switched pages since."""
    return _json(read.get_current_context(get_backend()))


# --- Create (service/create.py) ----------------------------------------------
# NOTE: there is deliberately no create_notebook tool. VM ground truth (2026-06-11): this
# M365 OneNote build refuses COM notebook creation — OpenHierarchy(cftNotebook) returns
# hrFileDoesNotExist for local paths AND OneDrive https parents alike. Notebooks are created
# in the OneNote UI; create_section covers everything below them.


@mcp.tool()
def create_section(parent_id: str, name: str) -> str:
    """Create a section under ``parent_id`` — an existing notebook OR a section group —
    inheriting its sync."""
    return _json({"section_id": create.create_section(get_backend(), parent_id, name)})


@mcp.tool()
def create_page(
    section_id: str, title: str, content: str | list[dict] = "", page_level: int = 1
) -> str:
    """Create a page in a section. ``page_level`` (1/2/3) sets subpage indent. ``content``
    optionally adds initial paragraphs — same shapes as update_page_content (plain text with
    newlines, or styled paragraph dicts)."""
    return _json(
        {"page_id": create.create_page(get_backend(), section_id, title, content, page_level)}
    )


# --- Modify (shared write core — service/page_edit.py) -----------------------


@mcp.tool()
def update_page_content(
    page_id: str,
    content: str | list[dict],
    mode: str = "append",
    target_object_id: str = "",
    force: bool = False,
) -> str:
    """Edit page content surgically — untouched paragraphs keep their formatting verbatim.

    mode: "append" (add paragraphs at the end of an outline; target_object_id optionally
    names an outline objectID, default = the page's last outline), "insert_before" /
    "insert_after" (target_object_id = a paragraph objectID from get_page; new paragraphs
    become its siblings), or "replace" (swap that paragraph's text, keeping its paragraph
    style unless the new content overrides it). Table cell text is edited by targeting the
    paragraph inside the cell with "replace".

    content: plain text (newlines split paragraphs) OR a list of paragraph dicts —
    {"text": "...", "style": {...}} or {"runs": [{"text": "...", "style": {...}}, ...]},
    each optionally with "quick_style_index" / "alignment". Style keys are the CSS-like keys
    get_page returns: font-weight, font-style, text-decoration, color, background (highlight),
    font-family, font-size.

    Concurrency-guarded: fails instead of clobbering if the page changed since it was read.
    Set force=True only after explicit user confirmation."""
    page_edit.edit_page_content(
        get_backend(), page_id, content, mode, target_object_id=target_object_id, force=force
    )
    return f"updated {page_id}"


@mcp.tool()
def create_table(
    page_id: str,
    rows: list[list[str | dict]],
    borders_visible: bool = True,
    has_header_row: bool = False,
    target_object_id: str = "",
    force: bool = False,
) -> str:
    """Add a table to a page, or append rows to an existing table.

    rows: cells are plain strings or dicts {"text" | "runs", "style", "shading_color",
    "alignment"} (short rows are padded). target_object_id: empty → new table at the end of
    the page's last outline; an outline objectID → new table in that outline; an existing
    table's objectID (from get_page) → append the rows to that table (row width must fit its
    columns). Concurrency-guarded; force=True only after explicit user confirmation."""
    page_edit.add_table(
        get_backend(),
        page_id,
        rows,
        borders_visible=borders_visible,
        has_header_row=has_header_row,
        target_object_id=target_object_id,
        force=force,
    )
    return f"table added to {page_id}"


@mcp.tool()
def insert_image(
    page_id: str,
    image_base64: str,
    media_type: str,
    width: float | None = None,
    height: float | None = None,
    target_object_id: str = "",
    force: bool = False,
) -> str:
    """Insert an image (base64 + media type, e.g. "image/png") into a page, appended to an
    outline (target_object_id = outline objectID, default the page's last outline).
    width/height are points; omit to let OneNote size it. Concurrency-guarded."""
    page_edit.insert_image(
        get_backend(),
        page_id,
        image_base64,
        media_type,
        width=width,
        height=height,
        target_object_id=target_object_id,
        force=force,
    )
    return f"image inserted into {page_id}"


# --- Copy (Phase 5: raw-XML faithful transfer) ------------------------------


@mcp.tool()
def copy_page(page_id: str, target_section_id: str) -> str:
    """Faithfully copy a page (formatting, tables, inline images, pageLevel) to a section.
    Returns the new page's ID."""
    return _json({"page_id": copy.transfer_page(get_backend(), page_id, target_section_id)})


@mcp.tool()
def copy_section(section_id: str, target_parent_id: str) -> str:
    """Faithfully copy a whole section (pages in order, subpage levels kept) into a notebook
    OR section group. The copy keeps the source name, de-collided with " (2)" if taken.
    Returns the new section's ID."""
    return _json({"section_id": copy.transfer_section(get_backend(), section_id, target_parent_id)})


# NOTE: there is deliberately no copy_notebook tool — same ground truth as create_notebook
# (COM cannot create notebooks on this build). Whole-notebook cloning is done by copy_section
# into an existing notebook / section group, section by section.


# --- Restructure (Phase 4: whole-batch UpdateHierarchy — SPEC §5 discipline) -
# Structural changes are propose-then-confirm: suggest a clone backup (copy_section)
# first, and present the target order for user confirmation before applying.


@mcp.tool()
def restructure_section(section_id: str, ordered_pages: list[dict]) -> str:
    """STRUCTURAL. Reorder ALL pages of a section in one batch and adjust subpage levels.
    ordered_pages = the section's complete page list in target order, each entry
    {"page_id": str, "page_level": 1|2|3}. Back up first (copy_section) and confirm the
    target order with the user before applying."""
    hierarchy_edit.restructure_section(get_backend(), section_id, ordered_pages)
    return f"section {section_id} restructured"


@mcp.tool()
def reorder_sections(notebook_id: str, ordered_section_ids: list[str]) -> str:
    """STRUCTURAL. Reorder a notebook's (or section group's) children in one batch.
    ordered_section_ids = the COMPLETE child list in target order, including BOTH sections
    and section groups exactly as list_sections shows them at that level (the hidden recycle
    bin is handled automatically). Notebook-level ordering itself is not supported. Back up
    first (copy_section) and confirm with the user before applying."""
    hierarchy_edit.reorder_sections(get_backend(), notebook_id, ordered_section_ids)
    return f"sections of {notebook_id} reordered"


@mcp.tool()
def rename_node(parent_id: str, object_id: str, new_name: str) -> str:
    """STRUCTURAL. Rename a page, section, or section group. parent_id = the containing
    section/notebook ID. (A page rename edits its title — the hierarchy name follows it.)
    Confirm with the user before applying."""
    hierarchy_edit.rename_node(get_backend(), parent_id, object_id, new_name)
    return f"{object_id} renamed to {new_name}"


@mcp.tool()
def move_page(notebook_id: str, page_id: str, target_section_id: str) -> str:
    """STRUCTURAL. Move a page to another section within the same notebook (it lands at the
    end of the target section and its subpage level resets to 1). The moved page gets a NEW
    page ID — returned here; use it for any follow-up calls. Confirm with the user before
    applying."""
    new_id = hierarchy_edit.move_page(get_backend(), notebook_id, page_id, target_section_id)
    return _json({"new_page_id": new_id, "section_id": target_section_id})


# --- Delete (Phase 6: destructive — conservative) ---------------------------


@mcp.tool()
def delete_node(object_id: str) -> str:
    """DESTRUCTIVE. Delete a hierarchy node (notebook / section group / section / page) to
    the recycle bin."""
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
