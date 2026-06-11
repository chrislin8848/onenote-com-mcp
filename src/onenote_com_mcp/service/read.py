"""Read-tool orchestration (SPEC §4 read tools, Phase 2).

Each function takes a backend and returns JSON-serializable Python data — the MCP tools in
``server.py`` are thin facades that ``json.dumps`` (or, for images, wrap as MCP image content).
All parsing lives in ``xmllayer``; this layer only orchestrates the backend call, projects the
fields the LLM needs, and resolves IDs to names. Pure Python — testable on Linux via
``FixtureBackend``.

Lossless representation (SPEC §5): ``get_page`` emits runs with their effective resolved style,
structured tables (never flattened to one string), and every content object's ``objectID`` (so
``delete_page_content`` can target it later).
"""

from __future__ import annotations

import base64
from typing import Any

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import HierarchyScope, PageInfo
from onenote_com_mcp.xmllayer.models import Image, Paragraph, Table
from onenote_com_mcp.xmllayer.parse import parse_hierarchy, parse_page

# --- hierarchy listings --------------------------------------------------------------


def _notebook_summary(node: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": node["id"],
        "name": node["name"],
        "nickname": node.get("nickname"),
        "color": node.get("color"),
        "last_modified_time": node["last_modified_time"],
        "is_currently_viewed": node["is_currently_viewed"],
    }


def list_notebooks(backend: OneNoteBackend) -> list[dict[str, Any]]:
    """All open notebooks (name + ID). ``GetHierarchy("", hsNotebooks)``."""
    xml = backend.get_hierarchy("", HierarchyScope.hsNotebooks)
    return [_notebook_summary(n) for n in parse_hierarchy(xml) if n["type"] == "notebook"]


def list_sections(backend: OneNoteBackend, notebook_id: str) -> list[dict[str, Any]]:
    """A notebook's direct children — a MIXED section + section-group list, nesting preserved
    (SPEC v0611: no flattening), recycle bin filtered. ``GetHierarchy(notebook, hsSections)``.
    """
    xml = backend.get_hierarchy(notebook_id, HierarchyScope.hsSections)
    nodes = parse_hierarchy(xml)
    # the scoped call roots at the notebook; return its children (the section/group list)
    return nodes[0]["children"] if nodes and nodes[0]["type"] == "notebook" else nodes


def list_pages(backend: OneNoteBackend, section_id: str) -> list[dict[str, Any]]:
    """A section's pages, each with its subpage level. ``GetHierarchy(section, hsPages)``.

    Tolerant of the wrapper the scoped call returns (bare ``one:Section`` or a notebook
    wrapper): collects every page node under the requested section.
    """
    xml = backend.get_hierarchy(section_id, HierarchyScope.hsPages)
    return _collect_pages(parse_hierarchy(xml))


def search_pages(backend: OneNoteBackend, query: str, scope_id: str = "") -> list[dict[str, Any]]:
    """Full-text page search, returned as a flat page list. ``FindPages(scope, query)``."""
    xml = backend.find_pages(scope_id, query)
    return _collect_pages(parse_hierarchy(xml))


def _collect_pages(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Flatten a hierarchy node list to just its page entries, document order preserved."""
    pages: list[dict[str, Any]] = []
    for node in nodes:
        if node["type"] == "page":
            pages.append(node)
        else:
            pages.extend(_collect_pages(node.get("children", [])))
    return pages


# --- page content --------------------------------------------------------------------


def _run_dict(run: Any) -> dict[str, Any]:
    # effective style = QuickStyleDef baseline ← OE style ← inline span (resolved in parse.py)
    return {"text": run.text, "style": run.style}


def _oe_object_id(img: Image) -> str | None:
    """The deletable object's ID for an image.

    Ground truth (real dump): a ``one:Image`` carries NO ``objectID`` of its own — the
    enclosing ``one:OE`` holds it. ``delete_page_content`` (Phase 6) must target that OE.
    """
    parent = img.node.getparent()
    return parent.get("objectID") if parent is not None else img.object_id


def _image_dict(img: Image) -> dict[str, Any]:
    return {
        "type": "image",
        "object_id": _oe_object_id(img),
        "callback_id": img.callback_id,
        "width": img.width,
        "height": img.height,
        "ocr_text": img.ocr_text,
    }


def _table_dict(table: Table) -> dict[str, Any]:
    return {
        "type": "table",
        "object_id": table.object_id,
        "borders_visible": table.borders_visible,
        "has_header_row": table.has_header_row,
        "columns": table.columns,
        "rows": [
            [
                {
                    "object_id": cell.object_id,
                    "shading_color": cell.shading_color,
                    "text": cell.text,
                    "paragraphs": [_paragraph_dict(p) for p in cell.paragraphs],
                }
                for cell in row
            ]
            for row in table.rows
        ],
    }


def _paragraph_dict(para: Paragraph) -> dict[str, Any]:
    """One ``one:OE`` → a block. An OE holds a table, an image, or text runs (+ nested OEs)."""
    if para.table is not None:
        return _table_dict(para.table)
    if para.image is not None:
        return _image_dict(para.image)
    block: dict[str, Any] = {
        "type": "paragraph",
        "object_id": para.object_id,
        "alignment": para.alignment,
        "quick_style_index": para.quick_style_index,
        "text": para.text,
        "runs": [_run_dict(r) for r in para.runs],
    }
    if para.children:
        block["children"] = [_paragraph_dict(c) for c in para.children]
    return block


def get_page(backend: OneNoteBackend, page_id: str) -> dict[str, Any]:
    """Read a page as a lossless runs+style model with structured tables and object IDs."""
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    return {
        "id": page.id,
        "name": page.name,
        "page_level": page.page_level,
        "last_modified_time": page.last_modified_time,
        "lang": page.lang,
        "quick_styles": {
            str(i): {
                "name": qs.name,
                "font": qs.font,
                "font_size": qs.font_size,
                "font_color": qs.font_color,
                "highlight_color": qs.highlight_color,
            }
            for i, qs in page.quick_styles.items()
        },
        "title": (
            {"text": page.title.text, "runs": [_run_dict(r) for r in page.title.runs]}
            if page.title is not None
            else None
        ),
        "outlines": [
            {
                "object_id": outline.object_id,
                "blocks": [_paragraph_dict(p) for p in outline.paragraphs],
            }
            for outline in page.outlines
        ],
    }


_IMAGE_MAGIC = (
    (b"\x89PNG\r\n\x1a\n", "image/png"),
    (b"\xff\xd8\xff", "image/jpeg"),
    (b"GIF87a", "image/gif"),
    (b"GIF89a", "image/gif"),
    (b"RIFF", "image/webp"),  # RIFF....WEBP; good enough for the magic-number prefix
)


def _sniff_media_type(raw: bytes) -> str:
    for magic, media_type in _IMAGE_MAGIC:
        if raw.startswith(magic):
            return media_type
    return "application/octet-stream"


def get_page_images(backend: OneNoteBackend, page_id: str) -> list[dict[str, Any]]:
    """A page's images as base64 binary + metadata.

    The page XML carries the image structure (object/callback IDs, dimensions, OCR); the bytes
    come from ``GetBinaryPageContent(callbackID)`` (SPEC §4 — no inline ``one:Data`` on a basic
    dump). ``get_page`` exposes the per-image metadata; this returns the pixels for the MCP
    image-content facade so Claude can recognize them visually.
    """
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    out: list[dict[str, Any]] = []
    for img in page.images:
        if not img.callback_id:
            continue
        data_b64 = backend.get_binary_page_content(page_id, img.callback_id)
        media_type = _sniff_media_type(base64.b64decode(data_b64))
        out.append(
            {
                "object_id": _oe_object_id(img),
                "callback_id": img.callback_id,
                "media_type": media_type,
                "data_base64": data_b64,
                "width": img.width,
                "height": img.height,
                "ocr_text": img.ocr_text,
            }
        )
    return out


# --- current context -----------------------------------------------------------------


def _id_name_map(nodes: list[dict[str, Any]]) -> dict[str, str]:
    out: dict[str, str] = {}
    for node in nodes:
        if node.get("id") and node.get("name") is not None:
            out[node["id"]] = node["name"]
        out.update(_id_name_map(node.get("children", [])))
    return out


def get_current_context(backend: OneNoteBackend) -> dict[str, Any]:
    """The active window's notebook / section group / section / page (IDs + names).

    Raises ``NoCurrentWindowError`` when no window is open (SPEC §5 — never guess). Granularity
    stops at the page; in-page cursor/selection has no COM API. Resolve names with one scoped
    ``GetHierarchy(notebook, hsPages)`` (covers groups, sections, and pages of the notebook).
    """
    ids = backend.get_current_window_ids()  # raises NoCurrentWindowError if no window
    names: dict[str, str] = {}
    if ids.notebook_id:
        # one scoped call: the notebook-rooted tree names the notebook, its groups,
        # sections, and pages — everything the four Current*Ids can point at.
        xml = backend.get_hierarchy(ids.notebook_id, HierarchyScope.hsPages)
        names = _id_name_map(parse_hierarchy(xml))

    def resolve(node_id: str | None) -> dict[str, Any] | None:
        if not node_id:
            return None
        return {"id": node_id, "name": names.get(node_id)}

    return {
        "notebook": resolve(ids.notebook_id),
        "section_group": resolve(ids.section_group_id),
        "section": resolve(ids.section_id),
        "page": resolve(ids.page_id),
    }
