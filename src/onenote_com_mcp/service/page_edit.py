"""The single page-content write path (SPEC §4 convergence point).

EVERY content write — update_page_content / create_table / insert_image — funnels through
``apply_page_edit``, so there is exactly ONE ``UpdatePageContent`` call site and ONE place the
concurrency guard is applied. The facades below own only content shaping: their ``mutate``
callbacks do surgical in-place edits on the live GetPageContent tree (SPEC §5 — never rebuild
from a slimmed model, or untouched paragraphs lose their formatting).

Payload strategy (DECIDED 2026-06-11 — BOTH validated on the VM, images byte-identical
through an edit either way): ``UpdatePageContent`` merges page-level objects, touching only
the ones present in the submitted XML. The default ``changed_objects`` strategy prunes the
payload down to the page shell + definitional children (QuickStyleDef etc.) + only the
page-level objects the mutation actually changed — untouched outlines/images are never
re-sent, so the merge cannot disturb them, and no piBinaryData read is needed.
``whole_page`` (read piBinaryData, send everything) stays as the validated fallback. In BOTH
strategies, every ``one:Image`` left in the payload gets its binary inlined (``one:Data``
fetched via GetBinaryPageContent, ``CallbackID`` removed): CallbackID is a read-side
construct, and an Image submitted without Data risks losing its pixels.
"""

from __future__ import annotations

import datetime as _dt
from collections.abc import Callable
from typing import Any, Literal

from lxml import etree

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import PageInfo
from onenote_com_mcp.errors import NodeNotFoundError, OneNoteComError
from onenote_com_mcp.xmllayer.build import (
    _cell_runs,
    make_image,
    make_table,
    make_table_row,
    make_text_oe,
)
from onenote_com_mcp.xmllayer.namespaces import local_name, qn
from onenote_com_mcp.xmllayer.spans import build_spans

# A mutation edits the parsed page tree IN PLACE (SPEC §5 — never rebuild from a slimmed model,
# or untouched paragraphs lose their formatting).
Mutator = Callable[[etree._Element], None]

PayloadStrategy = Literal["changed_objects", "whole_page"]

# DECIDED 2026-06-11 (Phase-4 VM round-trips): both strategies preserve untouched content;
# changed_objects ships smaller payloads and needs no piBinaryData read, so it is the default.
DEFAULT_PAYLOAD_STRATEGY: PayloadStrategy = "changed_objects"

# strip_cdata=False keeps one:T CDATA sections verbatim through the read → write round-trip.
_PARSER = etree.XMLParser(strip_cdata=False)

# Page-level content objects UpdatePageContent merges by. Direct page children NOT in this set
# (QuickStyleDef, TagDef, PageSettings, ...) are definitions/settings and always ride along.
# InsertedFile/XPSFile joined in Phase 5b (ground truth 2026-06-12: the page-level InsertedFile
# variant and printout XPSFile carriers are direct page children) — pruning unchanged ones keeps
# edits from re-submitting read-side constructs (XPSFile carries a CallbackID) untouched.
_CONTENT_TAGS = frozenset(
    {"Title", "Outline", "Image", "InkDrawing", "MediaFile", "InsertedFile", "XPSFile"}
)

_MODES = ("append", "insert_before", "insert_after", "replace")

# A 1x1 fully-transparent PNG. Used as the Data for an image whose real binary OneNote won't
# serve via COM (see inline_image_binaries): an Image needs Data OR CallbackID to be valid XML,
# so a placeholder keeps the payload accepted and the layout box intact instead of failing the
# whole copy. Reported via file_notes — never a silent substitution.
_PLACEHOLDER_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAXpeqz8AAAAASUVORK5CYII="
)


def parse_onenote_datetime(value: str | None) -> _dt.datetime | None:
    """Parse a OneNote ``lastModifiedTime`` (ISO-8601, possibly ``...Z``) → datetime."""
    if not value:
        return None
    try:
        return _dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def apply_page_edit(
    backend: OneNoteBackend,
    page_id: str,
    mutate: Mutator,
    *,
    strategy: PayloadStrategy | None = None,
    force: bool = False,
) -> None:
    """Read the page, mutate its XML tree in place, write it back in ONE guarded call.

    This is the only function in the codebase that calls ``backend.update_page_content`` for an
    edit. The concurrency guard (``dateExpectedLastModified``) is taken from the page we just
    read, so a write is refused if the page changed underneath us; ``force`` defaults False.
    """
    strategy = strategy or DEFAULT_PAYLOAD_STRATEGY
    read_info = PageInfo.piBinaryData if strategy == "whole_page" else PageInfo.piBasic
    xml = backend.get_page_content(page_id, read_info)
    tree = etree.fromstring(xml.encode("utf-8"), parser=_PARSER)
    expected = parse_onenote_datetime(tree.get("lastModifiedTime"))
    before = {child: etree.tostring(child, with_tail=False) for child in tree}
    mutate(tree)  # in place; untouched paragraphs keep their quickStyleIndex/spans verbatim
    if strategy == "changed_objects":
        _prune_unchanged_content(tree, before)
    inline_image_binaries(backend, page_id, tree)
    etree.cleanup_namespaces(tree)  # grafted fragments carry redundant xmlns:one declarations
    payload = etree.tostring(tree, xml_declaration=True, encoding="UTF-8").decode("utf-8")
    backend.update_page_content(payload, expected_last_modified=expected, force=force)


def _prune_unchanged_content(tree: etree._Element, before: dict[etree._Element, bytes]) -> None:
    """Drop page-level content objects the mutation did not touch (merge ignores absentees)."""
    for child in list(tree):
        if local_name(child.tag) not in _CONTENT_TAGS:
            continue
        snapshot = before.get(child)
        if snapshot is not None and etree.tostring(child, with_tail=False) == snapshot:
            tree.remove(child)


_CONTENT_LEAF_TAGS = frozenset({"T", "Image", "Table", "InsertedFile", "InkDrawing", "MediaFile"})
_PRUNABLE_CONTAINERS = frozenset({"OE", "OEChildren", "Outline"})


def remove_content_element(el: etree._Element) -> None:
    """Remove a content element (an un-copyable one:Image / one:InsertedFile) AND prune any
    now-empty one:OE / one:OEChildren / one:Outline ancestors.

    Dropping a one:Image alone empties its OE/Outline and makes UpdatePageContent fail with
    hrInvalidXML (VM ground truth) — and a left-behind empty/placeholder marker could later be
    misread as real content or re-copied. Pruning the emptied containers keeps the payload valid
    and leaves no dead marker. Stops at the first ancestor that still holds real content or is not
    a prunable container (e.g. the page root).

    A table CELL is the exception: it MUST keep a ``one:OEChildren`` with at least one ``one:OE``
    (an empty ``<one:Cell/>`` is rejected with hrInvalidXML — VM ground truth 2026-06-13, a
    fully-un-synced 祕魯18天 page with an image-only cell). So when pruning reaches a cell's
    OEChildren, the OEChildren is REPLENISHED with a minimal empty paragraph instead of removed."""
    parent = el.getparent()
    if parent is None:
        return
    parent.remove(el)
    node = parent
    while node is not None and local_name(node.tag) in _PRUNABLE_CONTAINERS:
        if any(local_name(d.tag) in _CONTENT_LEAF_TAGS for d in node.iter()):
            break  # still holds real content — keep this container
        up = node.getparent()
        if up is None:
            break
        if local_name(up.tag) == "Cell" and local_name(node.tag) == "OEChildren":
            node.append(make_text_oe([]))  # keep the cell valid; do NOT empty it
            break
        up.remove(node)
        node = up


def inline_image_binaries(
    backend: OneNoteBackend, page_id: str, tree: etree._Element, remove_unfetchable: bool = False
) -> int:
    """Ensure every one:Image in the payload carries inline one:Data, never a CallbackID.

    Public: the copy path (service/copy.py) needs the same guarantee — even a piBinaryData
    read serves CallbackID without inline Data (VM ground truth), so any tree heading into
    UpdatePageContent must have its pixels fetched via GetBinaryPageContent first.

    GRACEFUL DEGRADATION: an image's binary may not be fetchable — GetBinaryPageContent returns
    0x8004200F (hrBinaryObjectDoesNotExist). VM-confirmed root cause (2026-06-13): the image is
    not yet downloaded to THIS machine (OneDrive files-on-demand hydrates image binaries lazily,
    per page); a fully-synced machine serves them fine. (The earlier "OCR-processed images" theory
    was wrong — the OCR'd photos were simply the large, last-to-hydrate ones.) On an un-fetchable
    image, ``remove_unfetchable`` selects the fallback:
      * False (default, EDIT path): keep a valid (blank) box — a 1x1 transparent placeholder. The
        edit path must NOT delete an image that lives in the cloud but isn't downloaded here.
      * True (COPY path): REMOVE the image and prune the emptied OE/Outline — a copy must not carry
        a dead placeholder (it can't self-heal and could later be misread / re-copied).
    Either way the loss is counted and the caller reports it (SPEC §5: explicit, never silent); the
    user's fix is always to fully sync the source then copy again. Returns the COUNT of images that
    could not carry real pixels (0 = all carried)."""
    dropped = 0
    for image in list(tree.iter(qn("Image"))):
        callback = image.find(qn("CallbackID"))
        if image.find(qn("Data")) is None:
            callback_id = (
                callback.get("callbackID") if callback is not None else image.get("callbackID")
            )
            if not callback_id:
                continue  # nothing to fetch — leave untouched rather than invent data
            try:
                binary = backend.get_binary_page_content(page_id, callback_id)
            except OneNoteComError:
                dropped += 1
                if remove_unfetchable:
                    remove_content_element(image)
                    continue  # image (and any emptied container) gone — nothing more to do
                binary = _PLACEHOLDER_PNG_B64  # edit path: keep a valid (blank) Image box
            data = etree.Element(qn("Data"))
            data.text = binary
            ocr = image.find(qn("OCRData"))
            anchor = callback if callback is not None else ocr
            if anchor is not None:
                anchor.addprevious(data)
            else:
                image.append(data)
        if callback is not None:
            image.remove(callback)
        image.attrib.pop("callbackID", None)
    return dropped


# --- in-place mutation helpers ------------------------------------------------------


def _find_content_object(tree: etree._Element, object_id: str) -> etree._Element:
    el = next((e for e in tree.iter() if e.get("objectID") == object_id), None)
    if el is None:
        raise NodeNotFoundError(
            f"no object with objectID {object_id!r} on this page — object IDs come from get_page"
        )
    return el


def _resolve_outline(tree: etree._Element, target_object_id: str) -> etree._Element:
    """Outline to receive appended content: the given one, else the page's last (or a new one)."""
    if target_object_id:
        el = _find_content_object(tree, target_object_id)
        if local_name(el.tag) != "Outline":
            raise ValueError(
                f"target {target_object_id!r} is a one:{local_name(el.tag)}, not a one:Outline "
                "— new content is appended to an outline; to position relative to a paragraph "
                "use update_page_content with insert_before/insert_after/replace"
            )
        return el
    outlines = tree.findall(qn("Outline"))
    if outlines:
        return outlines[-1]
    outline = etree.SubElement(tree, qn("Outline"))
    etree.SubElement(outline, qn("OEChildren"))
    return outline


def _outline_children(outline: etree._Element) -> etree._Element:
    children = outline.find(qn("OEChildren"))
    if children is None:
        children = etree.SubElement(outline, qn("OEChildren"))
    return children


def _require_oe(tree: etree._Element, target_object_id: str, mode: str) -> etree._Element:
    if not target_object_id:
        raise ValueError(
            f"mode {mode!r} requires target_object_id — a paragraph objectID from get_page"
        )
    el = _find_content_object(tree, target_object_id)
    if local_name(el.tag) != "OE":
        raise ValueError(
            f"target {target_object_id!r} is a one:{local_name(el.tag)}; "
            f"{mode} needs a paragraph (one:OE)"
        )
    return el


def _normalize_paragraphs(content: str | list[Any]) -> list[dict[str, Any]]:
    """Tool ``content`` → paragraph specs ``{"runs", "quick_style_index", "alignment"}``.

    A plain string splits on newlines into unstyled paragraphs. A list holds str items or
    dicts with ``"runs"`` ([{text, style}, ...]) or ``"text"`` (+ optional ``"style"``),
    plus optional ``"quick_style_index"`` / ``"alignment"``.
    """
    items: list[Any] = content.split("\n") if isinstance(content, str) else content
    if not isinstance(items, list) or not items:
        raise ValueError("content must be a non-empty string or list of paragraphs")
    paragraphs: list[dict[str, Any]] = []
    for item in items:
        if isinstance(item, str):
            paragraphs.append({"runs": [item], "quick_style_index": None, "alignment": None})
        elif isinstance(item, dict):
            runs = item.get("runs") or [
                {"text": item.get("text", ""), "style": item.get("style"), "link": item.get("link")}
            ]
            paragraphs.append(
                {
                    "runs": runs,
                    "quick_style_index": item.get("quick_style_index"),
                    "alignment": item.get("alignment"),
                }
            )
        else:
            raise ValueError(f"each paragraph must be str or dict, got {type(item).__name__}")
    return paragraphs


def _replace_oe_text(oe: etree._Element, paragraph: dict[str, Any]) -> None:
    """Swap the OE's text runs in place. objectID, author attrs, nested children all stay;
    paragraph style (quickStyleIndex/alignment/style attr) stays unless the spec overrides."""
    new_t = etree.Element(qn("T"))
    new_t.text = etree.CDATA(build_spans(paragraph["runs"]))
    old_ts = oe.findall(qn("T"))
    if old_ts:
        old_ts[0].addprevious(new_t)
        for t in old_ts:
            oe.remove(t)
    else:
        nested = oe.find(qn("OEChildren"))
        if nested is not None:
            nested.addprevious(new_t)
        else:
            oe.insert(0, new_t)
    if paragraph["quick_style_index"] is not None:
        oe.set("quickStyleIndex", str(paragraph["quick_style_index"]))
    if paragraph["alignment"]:
        oe.set("alignment", paragraph["alignment"])


def _table_columns_el(table: etree._Element) -> etree._Element:
    columns = table.find(qn("Columns"))
    if columns is None:
        raise ValueError("the target table has no one:Columns element")
    return columns


def _renumber_columns(columns: etree._Element) -> None:
    for i, col in enumerate(columns.findall(qn("Column"))):
        col.set("index", str(i))


def _empty_cell() -> etree._Element:
    """A minimal valid one:Cell (OEChildren > OE > empty T) — an empty <one:Cell/> is rejected."""
    cell = etree.Element(qn("Cell"))
    etree.SubElement(cell, qn("OEChildren")).append(make_text_oe([]))
    return cell


def _insert_table_rows(table: etree._Element, rows: list[list[Any]], at_index: int | None) -> None:
    """Insert rows at a 0-based position; ``at_index=None`` appends at the end."""
    n_cols = len(_table_columns_el(table).findall(qn("Column")))
    widest = max(len(r) for r in rows)
    if widest > n_cols:
        raise ValueError(
            f"a row has {widest} cells but the table has {n_cols} columns — "
            "add columns first (modify_table operation='add_columns')"
        )
    existing = table.findall(qn("Row"))
    n_rows = len(existing)
    if at_index is None:
        at_index = n_rows
    if not (0 <= at_index <= n_rows):
        raise ValueError(f"at_index {at_index} is out of range 0..{n_rows}")
    new_rows = [make_table_row(row, n_cols) for row in rows]
    if at_index < n_rows:
        for nr in new_rows:
            existing[at_index].addprevious(nr)
    else:
        for nr in new_rows:
            table.append(nr)


def _add_table_columns(
    table: etree._Element, at_index: int | None, count: int, width: float | None
) -> None:
    """Insert ``count`` empty columns at a 0-based position (None = append at the end). Adds an
    empty cell to every row at the same position so the table stays rectangular."""
    if count < 1:
        raise ValueError("count must be >= 1")
    columns = _table_columns_el(table)
    cols = columns.findall(qn("Column"))
    n_cols = len(cols)
    if at_index is None:
        at_index = n_cols
    if not (0 <= at_index <= n_cols):
        raise ValueError(f"at_index {at_index} is out of range 0..{n_cols}")
    if width is None:
        last = cols[-1].get("width") if cols else None
        width = float(last) if last else 120.0
    for i in range(count):
        col = etree.Element(qn("Column"))
        col.set("width", str(float(width)))
        columns.insert(at_index + i, col)
    _renumber_columns(columns)
    for row in table.findall(qn("Row")):
        for i in range(count):
            row.insert(at_index + i, _empty_cell())


def _delete_table_rows(table: etree._Element, indices: list[int]) -> None:
    rows = table.findall(qn("Row"))
    n = len(rows)
    bad = sorted({i for i in indices if not (0 <= i < n)})
    if bad:
        raise ValueError(f"row index(es) {bad} out of range 0..{n - 1}")
    targets = sorted(set(indices))
    if len(targets) >= n:
        raise ValueError(
            "that would delete every row — delete the whole table with delete_page_content instead"
        )
    for i in reversed(targets):
        table.remove(rows[i])


def _delete_table_columns(table: etree._Element, indices: list[int]) -> None:
    columns = _table_columns_el(table)
    cols = columns.findall(qn("Column"))
    n = len(cols)
    bad = sorted({i for i in indices if not (0 <= i < n)})
    if bad:
        raise ValueError(f"column index(es) {bad} out of range 0..{n - 1}")
    targets = sorted(set(indices))
    if len(targets) >= n:
        raise ValueError(
            "that would delete every column — delete the whole table with delete_page_content "
            "instead"
        )
    for i in reversed(targets):
        columns.remove(cols[i])
        for row in table.findall(qn("Row")):
            cells = row.findall(qn("Cell"))
            if i < len(cells):
                row.remove(cells[i])
    _renumber_columns(columns)


def _set_cell_content(
    cell: etree._Element, runs: list[Any], shading_color: str | None, alignment: str | None
) -> None:
    """Replace a cell's text content in place: rewrite the cell's FIRST one:OE runs (keeping its
    objectID), drop any extra paragraph OEs, and set shadingColor only when given. The enclosing
    one:Cell keeps its objectID — this is a content edit, never a structural one."""
    children = cell.find(qn("OEChildren"))
    if children is None:
        children = etree.SubElement(cell, qn("OEChildren"))
    oes = children.findall(qn("OE"))
    if oes:
        _replace_oe_text(oes[0], {"runs": runs, "quick_style_index": None, "alignment": alignment})
        for extra in oes[1:]:
            children.remove(extra)
    else:
        children.append(make_text_oe(runs, alignment=alignment))
    if shading_color:
        cell.set("shadingColor", shading_color)


def _set_table_rows(table: etree._Element, rows: list[list[Any]], at_index: int | None) -> None:
    """Replace the CONTENT of existing rows from ``at_index`` (None = row 0), one input row per
    existing row. Fixed-shape: rows/columns are never added or removed — every one:Cell keeps its
    objectID. A short input row leaves the trailing columns untouched, and a ``None`` cell leaves
    THAT cell unchanged (so ``[None, "", ""]`` keeps column 0 and clears the rest). Out-of-range
    writes are refused (point at insert_rows / add_columns) so a content edit never silently grows
    it."""
    existing = table.findall(qn("Row"))
    n_rows = len(existing)
    n_cols = len(_table_columns_el(table).findall(qn("Column")))
    if at_index is None:
        at_index = 0
    if not (0 <= at_index < n_rows):
        raise ValueError(f"at_index {at_index} is out of range 0..{n_rows - 1}")
    end = at_index + len(rows)
    if end > n_rows:
        raise ValueError(
            f"set_rows would write rows {at_index}..{end - 1} but the table has only {n_rows} "
            "rows — add rows first (modify_table operation='insert_rows')"
        )
    widest = max((len(r) for r in rows), default=0)
    if widest > n_cols:
        raise ValueError(
            f"a row has {widest} cells but the table has {n_cols} columns — "
            "add columns first (modify_table operation='add_columns')"
        )
    for j, row_input in enumerate(rows):
        cells = existing[at_index + j].findall(qn("Cell"))
        for k, cell_input in enumerate(row_input):
            if cell_input is None:
                continue  # None = leave this cell exactly as it is (keep its content + identity)
            runs, shading_color, alignment = _cell_runs(cell_input)
            _set_cell_content(cells[k], runs, shading_color, alignment)


# --- Composable mutators + the facades the MCP write tools delegate to ---------------
# Each facade builds a Mutator and hands it to the single core above. ``content_mutator``
# and ``set_title`` are public so other services (create_page) can compose them into ONE
# guarded write instead of opening a second UpdatePageContent path.


def set_title(tree: etree._Element, title: str) -> None:
    """Set the page title IN PLACE: replace the existing Title OE's runs (its objectID and
    style stay), or create ``one:Title`` ahead of the outlines on a fresh page."""
    title_el = tree.find(qn("Title"))
    if title_el is None:
        title_el = etree.Element(qn("Title"))
        first_outline = tree.find(qn("Outline"))
        if first_outline is not None:
            first_outline.addprevious(title_el)
        else:
            tree.append(title_el)
    oe = title_el.find(qn("OE"))
    if oe is None:
        title_el.append(make_text_oe([title]))
    else:
        _replace_oe_text(oe, {"runs": [title], "quick_style_index": None, "alignment": None})


def content_mutator(
    content: str | list[Any], mode: str = "append", target_object_id: str = ""
) -> Mutator:
    """Validate the content/mode contract EAGERLY and return the in-place Mutator.

    append: add paragraphs to an outline (``target_object_id`` = outline objectID, default
    the page's last outline, created if the page has none). insert_before / insert_after:
    splice paragraphs as siblings of a target paragraph (one:OE objectID). replace: swap the
    target paragraph's text, keeping its identity and paragraph style."""
    if mode not in _MODES:
        raise ValueError(f"mode must be one of {_MODES}, got {mode!r}")
    paragraphs = _normalize_paragraphs(content)

    def mutate(tree: etree._Element) -> None:
        if mode == "append":
            children = _outline_children(_resolve_outline(tree, target_object_id))
            for p in paragraphs:
                children.append(make_text_oe(p["runs"], p["quick_style_index"], p["alignment"]))
            return
        target = _require_oe(tree, target_object_id, mode)
        if mode == "replace":
            _replace_oe_text(target, paragraphs[0])
            extra, anchor = paragraphs[1:], target
        elif mode == "insert_after":
            extra, anchor = paragraphs, target
        else:  # insert_before
            for p in paragraphs:
                target.addprevious(make_text_oe(p["runs"], p["quick_style_index"], p["alignment"]))
            return
        for p in extra:
            oe = make_text_oe(p["runs"], p["quick_style_index"], p["alignment"])
            anchor.addnext(oe)
            anchor = oe

    return mutate


def edit_page_content(
    backend: OneNoteBackend,
    page_id: str,
    content: str | list[Any],
    mode: str = "append",
    *,
    target_object_id: str = "",
    force: bool = False,
) -> None:
    """See :func:`content_mutator` for the mode/content contract."""
    apply_page_edit(backend, page_id, content_mutator(content, mode, target_object_id), force=force)


def add_table(
    backend: OneNoteBackend,
    page_id: str,
    rows: list[list[Any]],
    *,
    borders_visible: bool = True,
    has_header_row: bool = False,
    col_widths: list[float] | None = None,
    target_object_id: str = "",
    force: bool = False,
) -> None:
    """Create a NEW one:Table (wrapped in its own one:OE). No target → the page's last outline
    (created if none); an outline objectID → that outline. To change an EXISTING table's shape
    (add/insert rows, add columns, delete rows/columns) use ``modify_table`` instead."""
    if not rows:
        raise ValueError("rows is empty")

    def mutate(tree: etree._Element) -> None:
        if target_object_id:
            target = _find_content_object(tree, target_object_id)
            kind = local_name(target.tag)
            if kind == "Table":
                raise ValueError(
                    f"create_table target {target_object_id!r} is an existing table — "
                    "create_table only makes NEW tables; use modify_table to add rows/columns"
                )
            if kind != "Outline":
                raise ValueError(
                    f"create_table target {target_object_id!r} is a one:{kind} — it must be "
                    "an outline objectID (or omitted to use the page's last outline)"
                )
            outline = target
        else:
            outline = _resolve_outline(tree, "")
        oe = etree.Element(qn("OE"))
        oe.append(make_table(rows, borders_visible, has_header_row, col_widths))
        _outline_children(outline).append(oe)

    apply_page_edit(backend, page_id, mutate, force=force)


_TABLE_OPS = ("add_columns", "insert_rows", "delete_columns", "delete_rows", "set_rows")
TableOp = Literal["add_columns", "insert_rows", "delete_columns", "delete_rows", "set_rows"]


def modify_table(
    backend: OneNoteBackend,
    page_id: str,
    table_object_id: str,
    operation: TableOp,
    *,
    rows: list[list[Any]] | None = None,
    indices: list[int] | None = None,
    at_index: int | None = None,
    count: int = 1,
    width: float | None = None,
    force: bool = False,
) -> None:
    """Change an EXISTING table's shape in place (its objectID + every cell's identity are kept).

    operation:
      * ``insert_rows``  — insert ``rows`` (cell content, same shape as create_table) at the
        0-based ``at_index``; omit ``at_index`` to append at the end.
      * ``add_columns``  — insert ``count`` empty columns at ``at_index`` (omit = append at the
        end); ``width`` defaults to the last column's. Every row gets an empty cell so the table
        stays rectangular. Fill the new cells afterwards with update_page_content (replace).
      * ``set_rows``     — REPLACE the content of existing rows with ``rows`` (cell content, same
        shape as create_table), starting at ``at_index`` (omit = row 0), one input row per existing
        row. Fixed-shape: no row/column is added or removed and every cell keeps its objectID; a
        short input row leaves trailing columns untouched, and a ``None`` cell leaves THAT cell
        unchanged (``[None, "", ""]`` keeps column 0, clears the rest). Give one row + ``at_index``
        to replace a single row. Writing past the last row / wider than the table is refused (grow
        it first with insert_rows / add_columns).
      * ``delete_rows``    — remove the rows at ``indices`` (0-based). DESTRUCTIVE.
      * ``delete_columns`` — remove the columns at ``indices`` (0-based) and the matching cell in
        every row. DESTRUCTIVE.
    Row/column indices match get_page's table layout. Deleting every row/column is refused
    (delete the whole table with delete_page_content instead)."""
    if operation not in _TABLE_OPS:
        raise ValueError(f"operation must be one of {_TABLE_OPS}, got {operation!r}")
    if not table_object_id:
        raise ValueError("table_object_id is required (a one:Table objectID from get_page)")
    if operation in ("insert_rows", "set_rows") and not rows:
        raise ValueError(f"{operation} requires non-empty rows")
    if operation in ("delete_rows", "delete_columns") and not indices:
        raise ValueError(f"{operation} requires non-empty indices")

    def mutate(tree: etree._Element) -> None:
        table = _find_content_object(tree, table_object_id)
        if local_name(table.tag) != "Table":
            raise ValueError(
                f"target {table_object_id!r} is a one:{local_name(table.tag)}, not a table — "
                "pass a table objectID from get_page"
            )
        if operation == "insert_rows":
            _insert_table_rows(table, rows, at_index)
        elif operation == "add_columns":
            _add_table_columns(table, at_index, count, width)
        elif operation == "set_rows":
            _set_table_rows(table, rows, at_index)
        elif operation == "delete_rows":
            _delete_table_rows(table, indices)
        else:  # delete_columns
            _delete_table_columns(table, indices)

    apply_page_edit(backend, page_id, mutate, force=force)


# Inline objects delete_inline_content can remove (an inline table is the one:Table; an inline
# image/attachment/paragraph is the enclosing one:OE — get_page reports the OE's objectID for
# those, the Table's own objectID for a table).
_INLINE_DELETABLE = frozenset({"Table", "OE", "Image", "InsertedFile", "InkDrawing", "MediaFile"})


def delete_inline_content(
    backend: OneNoteBackend, page_id: str, object_id: str, *, force: bool = False
) -> None:
    """Delete ONE inline object from inside an outline — a table, an inline image/attachment, or a
    paragraph — by objectID, via the edit seam (NOT DeletePageContent, which COM refuses for inline
    OEs). The deleted element's emptied OE/OEChildren/Outline ancestors are pruned (a table cell is
    kept valid, never emptied); sibling paragraphs in the same outline are untouched — so "delete
    the table, keep the paragraphs" just works. The complement of delete_page_content (page-level
    objects): a page-level objectID is refused here with a pointer to that tool."""
    if not object_id:
        raise ValueError("object_id is required (an inline objectID from get_page)")

    def mutate(tree: etree._Element) -> None:
        target = _find_content_object(tree, object_id)
        parent = target.getparent()
        name = local_name(target.tag)
        if parent is not None and local_name(parent.tag) == "Page":
            raise ValueError(
                f"{object_id!r} is a page-level one:{name} — delete_inline_content removes content "
                "INSIDE an outline; use delete_page_content for a whole outline, a page-level "
                "image, or a page-level attachment"
            )
        if name not in _INLINE_DELETABLE:
            raise ValueError(
                f"{object_id!r} is a one:{name}; delete_inline_content removes an inline table, "
                "image, attachment, or paragraph — pass a table's objectID, or the paragraph/OE "
                "objectID get_page reports (an inline image/attachment uses its enclosing OE's ID)"
            )
        remove_content_element(target)

    apply_page_edit(backend, page_id, mutate, force=force)


def insert_image(
    backend: OneNoteBackend,
    page_id: str,
    image_base64: str,
    media_type: str,
    *,
    width: float | None = None,
    height: float | None = None,
    target_object_id: str = "",
    force: bool = False,
) -> None:
    """Append a new image (inline one:Data) to an outline, wrapped in its own one:OE —
    the OE is what carries the deletable objectID (a one:Image has none)."""

    def mutate(tree: etree._Element) -> None:
        outline = _resolve_outline(tree, target_object_id)
        oe = etree.Element(qn("OE"))
        oe.append(make_image(image_base64, media_type, width, height))
        _outline_children(outline).append(oe)

    apply_page_edit(backend, page_id, mutate, force=force)
