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
from onenote_com_mcp.errors import NodeNotFoundError, OneNoteComError
from onenote_com_mcp.xmllayer.models import Image, InsertedFile, Paragraph, Table
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
    d: dict[str, Any] = {"text": run.text, "style": run.style}
    if run.link:  # hyperlink href — surfaced so it round-trips through a replace edit
        d["link"] = run.link
    return d


def _oe_object_id(img: Image) -> str | None:
    """The deletable object's ID for an image.

    Ground truth (real dump): an INLINE ``one:Image`` carries NO ``objectID`` of its own —
    the enclosing ``one:OE`` holds it; ``delete_page_content`` (Phase 6) must target that OE.
    A PAGE-LEVEL image (printout render, direct ``one:Page`` child) has its own objectID and
    its parent (the page root) has none — fall through to the image's own.
    """
    parent = img.node.getparent()
    return (parent.get("objectID") if parent is not None else None) or img.object_id


def _image_dict(img: Image) -> dict[str, Any]:
    return {
        "type": "image",
        "object_id": _oe_object_id(img),
        "callback_id": img.callback_id,
        "width": img.width,
        "height": img.height,
        "ocr_text": img.ocr_text,
        "is_printout": img.is_printout,
    }


def _table_dict(
    table: Table, row_idx: list[int] | None = None, col_idx: list[int] | None = None
) -> dict[str, Any]:
    """The full structured table model. ``row_idx`` / ``col_idx`` (get_table's window) restrict it
    to those rows / columns; omitted = every row / column (get_page and get_object always pass
    none)."""
    rows = range(len(table.rows)) if row_idx is None else row_idx
    return {
        "type": "table",
        "object_id": table.object_id,
        "borders_visible": table.borders_visible,
        "has_header_row": table.has_header_row,
        "columns": table.columns if col_idx is None else [table.columns[j] for j in col_idx],
        # parallel to ``rows`` by index — pass one as apply_text_style scope_object_id to restyle a
        # whole row (a column has no objectID; use apply_text_style's ``columns`` index instead).
        "row_object_ids": [table.row_object_ids[i] for i in rows],
        "rows": [
            [
                {
                    "object_id": cell.object_id,
                    "shading_color": cell.shading_color,
                    "text": cell.text,
                    "paragraphs": [_paragraph_dict(p) for p in cell.paragraphs],
                }
                for cell in _pick(table.rows[i], col_idx)
            ]
            for i in rows
        ],
    }


def _pick(row: list[Any], col_idx: list[int] | None) -> list[Any]:
    """The cells of ``row`` at ``col_idx`` (all when None; a ragged row's missing cell skipped)."""
    return row if col_idx is None else [row[j] for j in col_idx if j < len(row)]


def _paragraph_dict(para: Paragraph) -> dict[str, Any]:
    """One ``one:OE`` → a block. An OE holds a table, an image, a file, or text runs."""
    if para.table is not None:
        return _table_dict(para.table)
    if para.image is not None:
        return _image_dict(para.image)
    if para.inserted_file is not None:
        # presence + identity only — metadata/sizes via get_page_files_info, content (text/
        # image/PDF) via get_page_files
        return {
            "type": "file",
            "object_id": para.object_id,
            "kind": para.inserted_file.kind,
            "preferred_name": para.inserted_file.preferred_name,
        }
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


def _text_grid(table: Table) -> list[list[str]]:
    return [[c.text for c in row] for row in table.rows]


def _text_block(para: Paragraph) -> Any:
    """``text_only`` projection of one ``one:OE`` — the words, nothing else (no runs, styles or
    objectIDs): a plain paragraph is just its text string; a table is ``{"table": [[cell text]]}``;
    an image is ``{"image": its OCR text or ""}``; an attachment is ``{"file": its name}``; nested
    paragraphs ride under ``"children"``."""
    if para.table is not None:
        block: dict[str, Any] = {"table": _text_grid(para.table)}
    elif para.image is not None:
        block = {"image": para.image.ocr_text or ""}
    elif para.inserted_file is not None:
        block = {"file": para.inserted_file.preferred_name}
    elif not para.children:
        return para.text
    else:
        block = {"text": para.text}
    if para.children:
        block["children"] = [_text_block(c) for c in para.children]
    return block


def get_page(backend: OneNoteBackend, page_id: str, *, text_only: bool = False) -> dict[str, Any]:
    """Read a page as a lossless runs+style model with structured tables and object IDs — or, with
    ``text_only``, as just its words: each outline a list of blocks (see ``_text_block``), the title
    a string, page-level printout images as their OCR text and page-level attachments as names.
    Typically 10-100x smaller; the right read for reviewing / summarizing / syncing content."""
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    if text_only:
        out: dict[str, Any] = {
            "id": page.id,
            "name": page.name,
            "page_level": page.page_level,
            "last_modified_time": page.last_modified_time,
            "title": page.title.text if page.title is not None else None,
            "outlines": [[_text_block(p) for p in o.paragraphs] for o in page.outlines],
        }
        if page.page_images:
            out["page_level_images"] = [img.ocr_text or "" for img in page.page_images]
        if page.page_files:
            out["page_level_files"] = [f.preferred_name for f in page.page_files]
        return out
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
        # Page-LEVEL objects: direct one:Page children that live OUTSIDE any outline — printout
        # render images and the page-level InsertedFile variant. Previously OMITTED here (only
        # `outlines` was emitted), so a printout page's images were INVISIBLE to a reader of
        # get_page and could not be found/deleted; surfaced now so the page read is complete.
        # These are deleted with delete_page_content (they ARE page-level objects).
        "page_level_images": [_image_dict(img) for img in page.page_images],
        "page_level_files": [_pagelevel_file_dict(f) for f in page.page_files],
    }


def _is_displayed(backend: OneNoteBackend, page_id: str) -> bool:
    """Is this page the one on screen in OneNote? Never raises (no window → False)."""
    try:
        return backend.get_current_window_ids().page_id == page_id
    except Exception:  # noqa: BLE001
        return False


# From ~this size a single-cell edit visibly stalls: OneNote re-processes every cell of the content
# box (VM 2026-10-05, 4,134 cells: ~25s per write, 134-141s if the page is on screen).
_BIG_TABLE_CELLS = 1000


def get_table(
    backend: OneNoteBackend,
    page_id: str,
    table_object_id: str,
    *,
    text_only: bool = False,
    start_row: int = 0,
    max_rows: int | None = None,
    columns: list[int] | None = None,
) -> dict[str, Any]:
    """Read ONE table WITHOUT the rest of the page, found anywhere on it (incl. nested in a cell).

    Two shapes. ``text_only`` → each cell's plain text as a 2-D string array — no runs, styles,
    shading or objectIDs: the compact read for a table's DATA. A real 106×39 班表 is megabytes in
    the full model (every cell, even an empty one, carries cell + paragraph objectIDs and a resolved
    style per run) but tens of KB as text. Default → the full structured model get_page emits for a
    table (runs + effective style + shading + row/cell/paragraph objectIDs).

    Either shape can be windowed: ``start_row`` + ``max_rows`` (0-indexed rows) and ``columns``
    (0-indexed column numbers, in the order given). ``total_rows`` / ``total_columns`` always report
    the whole table's size and ``start_row`` / ``returned_rows`` / ``column_indices`` say which
    slice came back, so a
    caller can page through a big table. Raises if no such table is on the page or the window is
    invalid."""
    if start_row < 0:
        raise ValueError(f"start_row must be >= 0, got {start_row}")
    if max_rows is not None and max_rows < 1:
        raise ValueError(f"max_rows must be >= 1, got {max_rows}")
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    table = next((t for t in page.tables if t.object_id == table_object_id), None)
    if table is None:
        raise NodeNotFoundError(
            f"no table with objectID {table_object_id!r} on this page — table object IDs come "
            "from get_page / get_page_info"
        )
    n_rows, n_cols = len(table.rows), len(table.columns)
    if columns is not None:
        bad = [j for j in columns if not 0 <= j < n_cols]
        if bad:
            raise ValueError(f"column index(es) {bad} out of range — the table has {n_cols}")
    end = n_rows if max_rows is None else min(n_rows, start_row + max_rows)
    row_idx = list(range(min(start_row, n_rows), end))
    col_idx = list(range(n_cols)) if columns is None else list(columns)
    window = {
        "total_rows": n_rows,
        "total_columns": n_cols,
        "start_row": row_idx[0] if row_idx else start_row,
        "returned_rows": len(row_idx),
        "column_indices": col_idx,
    }
    if text_only:
        body: dict[str, Any] = {
            "type": "table",
            "object_id": table.object_id,
            "has_header_row": table.has_header_row,
            **window,
            "rows": [[c.text for c in _pick(table.rows[i], col_idx)] for i in row_idx],
        }
    else:
        body = {**_table_dict(table, row_idx, col_idx), **window}
    if n_rows * n_cols >= _BIG_TABLE_CELLS:
        body["write_cost"] = (
            f"large table ({n_rows * n_cols} cells): OneNote rewrites the WHOLE table on every "
            "edit, so each write takes ~25s or more regardless of how many cells change — batch "
            "all cell changes into ONE batch_update / modify_table set_rows call; if a write times "
            "out it most likely still completed, so re-read (text_only) before retrying"
        )
        if _is_displayed(backend, page.id):
            body["write_cost"] += (
                ". This page is OPEN in OneNote right now, which makes writes ~5x slower (minutes) "
                "— before editing, ask the user to switch OneNote to another page"
            )
    return {"page_id": page.id, "last_modified_time": page.last_modified_time, "table": body}


def get_object(
    backend: OneNoteBackend, page_id: str, object_id: str, *, text_only: bool = False
) -> dict[str, Any]:
    """Read ONE object's full content — a paragraph's text + runs + style, a table, an image, or
    an attachment — by its objectID, WITHOUT the rest of the page.

    The targeted companion to get_page: to inspect or fix one paragraph you need only this, not the
    whole-page payload (which duplicates each run under both ``text`` and ``runs`` and carries the
    page-wide style table). Get the object_id from get_page_info or find_objects. Finds the object
    anywhere on the page (inline, nested in a table cell, or page-level). ``text_only`` returns
    just its words (same projection as get_page's text_only). Raises if no object on the page has
    that id."""
    if not object_id:
        raise ValueError("object_id is empty")
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    found = _locate_object(page, object_id)
    if found is None:
        raise NodeNotFoundError(
            f"no object with objectID {object_id!r} on this page — object IDs come from "
            "get_page / get_page_info / find_objects"
        )
    kind, model = found
    if text_only:
        obj: Any = {
            "table": lambda: {"table": _text_grid(model)},
            "image": lambda: {"image": model.ocr_text or ""},
            "file": lambda: {"file": model.preferred_name},
            "paragraph": lambda: _text_block(model),
        }[kind]()
    else:
        obj = {
            "table": _table_dict,
            "image": _image_dict,
            "file": _pagelevel_file_dict,
            "paragraph": _paragraph_dict,
        }[kind](model)
    return {"page_id": page.id, "last_modified_time": page.last_modified_time, "object": obj}


def _locate_object(page: Any, object_id: str) -> tuple[str, Any] | None:
    """Find an object by id → (kind, model); the caller projects it (full or text_only)."""
    for table in page.tables:  # incl. tables nested in a cell
        if table.object_id == object_id:
            return "table", table
    for img in page.images:  # inline AND page-level (printout renders)
        if _oe_object_id(img) == object_id:
            return "image", img
    for f in page.page_files:  # page-level attachments / printout carriers
        if f.object_id == object_id:
            return "file", f
    for p in page.paragraphs:  # text paragraphs + inline attachments (recursive: incl. cells)
        if p.object_id == object_id:
            return "paragraph", p
    return None


def find_objects(backend: OneNoteBackend, page_id: str, query: str) -> dict[str, Any]:
    """Locate the objects on a page whose TEXT contains a substring — return their objectIDs.

    The within-page counterpart to search_pages (which returns whole PAGES, never locations on a
    page). Use it to find exactly which paragraph(s) to edit — e.g. the one holding a typo — without
    pulling the whole page or eyeballing get_page_info's truncated previews: it matches the FULL
    paragraph text, not a 40-char preview. Searches body and table-cell paragraphs and returns each
    match's object_id (ready for get_object / update_page_content / find_and_replace) plus a short
    preview. Case-sensitive substring match (so 開鑿 ≠ 開逑)."""
    if not query:
        raise ValueError("query is empty")
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    matches: list[dict[str, Any]] = []
    for p in page.paragraphs:
        if p.table is not None or p.image is not None or p.inserted_file is not None:
            continue  # only text-bearing paragraphs have a body to search
        text = p.text or ""
        if query in text:
            matches.append(
                {"object_id": p.object_id, "type": "paragraph", "preview": _preview(text)}
            )
    return {"page_id": page.id, "name": page.name, "query": query, "matches": matches}


def _pagelevel_file_dict(f: InsertedFile) -> dict[str, Any]:
    """A page-level one:InsertedFile (own objectID; printout carrier / direct page attachment)."""
    return {
        "type": "file",
        "object_id": f.object_id,
        "kind": f.kind,
        "preferred_name": f.preferred_name,
    }


def _preview(text: str | None, limit: int = 40) -> str:
    """A short, single-line text preview so the model can tell paragraphs apart by content
    (an object_id alone is meaningless to a human/LLM) without get_page's full run model."""
    t = " ".join((text or "").split())
    return t if len(t) <= limit else t[:limit] + "…"


def _inventory_item(obj_type: str, object_id: str | None, page_level: bool, **meta: Any) -> dict:
    return {
        "type": obj_type,
        "object_id": object_id,
        # which delete tool removes it: page-level objects → delete_page_content; everything
        # inside an outline (paragraph / table / inline image / inline attachment) →
        # delete_inline_content. (Surfaced so the model picks the right tool without guessing.)
        "delete_with": "delete_page_content" if page_level else "delete_inline_content",
        "page_level": page_level,
        **meta,
    }


# A table with at least this many cells is listed in get_page_info as ONE summary entry: its plain
# text cell paragraphs are NOT enumerated (a 106×39 班表 made the inventory ~637K chars, 4,134
# near-useless entries). Images / attachments / nested tables inside its cells still are.
_COLLAPSE_TABLE_CELLS = 200


def _collapsed_cell_paragraph_ids(page: Any, include_cells: bool) -> dict[str, set[str]]:
    """table objectID → objectIDs of the plain-text paragraphs inside its cells (incl. nested
    tables' cells), for every table big enough to collapse. Empty when include_cells."""
    if include_cells:
        return {}
    out: dict[str, set[str]] = {}
    for table in page.tables:
        if len(table.rows) * len(table.columns) < _COLLAPSE_TABLE_CELLS:
            continue
        out[table.object_id] = set(_plain_cell_paragraph_ids(table))
    return out


def _plain_cell_paragraph_ids(table: Table):
    """objectIDs of the plain-text paragraphs in a table's cells — recursing into nested
    paragraphs and nested tables' cells (their images/attachments/tables are NOT yielded)."""

    def walk(paras):
        for p in paras:
            if p.table is not None:
                yield from _plain_cell_paragraph_ids(p.table)
            elif p.image is None and p.inserted_file is None and p.object_id:
                yield p.object_id
            yield from walk(p.children)

    for row in table.rows:
        for cell in row:
            yield from walk(cell.paragraphs)


def _page_object_inventory(page: Any, include_cells: bool = False) -> list[dict[str, Any]]:
    """A FLAT, document-order list of every object on the page — the lightweight half of the
    two-step objectID rule. Walks outlines recursively (into nested children AND table cells)
    so nothing is buried, then appends the page-level objects (printout renders, page-level
    attachments) that get_page's structured view used to hide. Each entry carries the object's
    targetable object_id, which delete tool removes it, and light type metadata — NO full text
    runs, NO style table, NO pixels. A table appears as one entry (rows×cols); its cells'
    contents follow as their own entries (they are distinct, separately-editable objects) — EXCEPT
    in a big table (>= _COLLAPSE_TABLE_CELLS cells, unless include_cells): its plain-text cell
    paragraphs are summarized on the table entry (``cell_paragraphs_not_listed``) instead of
    listed; images / attachments / nested tables in its cells are still listed."""
    collapsed = _collapsed_cell_paragraph_ids(page, include_cells)
    hidden = set().union(*collapsed.values()) if collapsed else set()
    items: list[dict[str, Any]] = []
    for p in page.paragraphs:  # recursive: outline paragraphs, children, and table-cell paragraphs
        if p.table is not None:
            entry = _inventory_item(
                "table",
                p.table.object_id,
                False,
                rows=len(p.table.rows),
                columns=len(p.table.columns),
            )
            if p.table.object_id in collapsed:
                entry["cell_paragraphs_not_listed"] = len(collapsed[p.table.object_id])
                entry["cells_note"] = (
                    "big table: its cell paragraphs are not listed here (images/attachments in "
                    "cells still are). Read cell text with get_table(text_only=True); find one "
                    "cell's paragraph objectID with find_objects, or call get_page_info with "
                    "include_cells=True"
                )
            items.append(entry)
        elif p.image is not None:
            img = p.image
            items.append(
                _inventory_item(
                    "image",
                    _oe_object_id(img),
                    False,
                    width=img.width,
                    height=img.height,
                    has_ocr=bool(img.ocr_text),
                    is_printout=img.is_printout,
                )
            )
        elif p.inserted_file is not None:
            items.append(
                _inventory_item(
                    "file",
                    p.object_id,
                    False,
                    kind=p.inserted_file.kind,
                    preferred_name=p.inserted_file.preferred_name,
                )
            )
        elif p.object_id not in hidden:
            items.append(_inventory_item("paragraph", p.object_id, False, preview=_preview(p.text)))
    for img in page.page_images:  # page-level printout renders — own objectID, delete_page_content
        items.append(
            _inventory_item(
                "image",
                _oe_object_id(img),
                True,
                width=img.width,
                height=img.height,
                has_ocr=bool(img.ocr_text),
                is_printout=img.is_printout,
            )
        )
    for f in page.page_files:  # page-level attachments / printout carriers
        items.append(
            _inventory_item("file", f.object_id, True, kind=f.kind, preferred_name=f.preferred_name)
        )
    return items


def get_page_info(
    backend: OneNoteBackend, page_id: str, *, include_cells: bool = False
) -> dict[str, Any]:
    """A lightweight inventory of every object on a page (IDs + metadata, NOT full content).

    The cheap companion to get_page: same parse, but projects only the flat object list — so
    "what's on this page / what is its objectID / what can I delete" costs no full-text/style
    payload. Crucially EXHAUSTIVE for images / attachments / tables: nested (table-cell) and
    page-level (printout) objects are all listed, which is exactly what get_page's structured
    view could bury. A big table's plain-text cell paragraphs are summarized, not listed, unless
    ``include_cells``."""
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    return {
        "id": page.id,
        "name": page.name,
        "page_level": page.page_level,
        "last_modified_time": page.last_modified_time,
        "objects": _page_object_inventory(page, include_cells),
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

    Some images cannot be served by COM — OCR-processed images return 0x8004200F from
    ``GetBinaryPageContent`` (VM ground truth; the same wall the copy path hits). Their pixels
    are simply unavailable, so they are SKIPPED here rather than crashing the whole read; the
    image still exists in ``get_page`` (object_id, dimensions, OCR text), which is the authority
    on what a page contains, so nothing is hidden — only the unviewable pixels are dropped.
    """
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    out: list[dict[str, Any]] = []
    for img in page.images:  # inline AND page-level (printout renders) — both have callbacks
        if not img.callback_id:
            continue
        try:
            data_b64 = backend.get_binary_page_content(page_id, img.callback_id)
        except OneNoteComError:
            continue  # un-fetchable (e.g. OCR'd image) — see get_page for its metadata
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
                "is_printout": img.is_printout,
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
