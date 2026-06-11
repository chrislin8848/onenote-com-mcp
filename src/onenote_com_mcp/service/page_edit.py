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
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.xmllayer.build import make_image, make_table, make_table_row, make_text_oe
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
_CONTENT_TAGS = frozenset({"Title", "Outline", "Image", "InkDrawing", "MediaFile"})

_MODES = ("append", "insert_before", "insert_after", "replace")


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


def inline_image_binaries(backend: OneNoteBackend, page_id: str, tree: etree._Element) -> None:
    """Ensure every one:Image in the payload carries inline one:Data, never a CallbackID.

    Public: the copy path (service/copy.py) needs the same guarantee — even a piBinaryData
    read serves CallbackID without inline Data (VM ground truth), so any tree heading into
    UpdatePageContent must have its pixels fetched via GetBinaryPageContent first."""
    for image in tree.iter(qn("Image")):
        callback = image.find(qn("CallbackID"))
        if image.find(qn("Data")) is None:
            callback_id = (
                callback.get("callbackID") if callback is not None else image.get("callbackID")
            )
            if not callback_id:
                continue  # nothing to fetch — leave untouched rather than invent data
            data = etree.Element(qn("Data"))
            data.text = backend.get_binary_page_content(page_id, callback_id)
            ocr = image.find(qn("OCRData"))
            anchor = callback if callback is not None else ocr
            if anchor is not None:
                anchor.addprevious(data)
            else:
                image.append(data)
        if callback is not None:
            image.remove(callback)
        image.attrib.pop("callbackID", None)


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
            runs = item.get("runs") or [{"text": item.get("text", ""), "style": item.get("style")}]
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


def _append_table_rows(table: etree._Element, rows: list[list[Any]]) -> None:
    n_cols = len(table.findall(f"{qn('Columns')}/{qn('Column')}"))
    widest = max(len(r) for r in rows)
    if widest > n_cols:
        raise ValueError(
            f"a row has {widest} cells but the table has {n_cols} columns — "
            "adding columns to an existing table is not supported"
        )
    for row in rows:
        table.append(make_table_row(row, n_cols))


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
    """No target / an outline target → append a NEW one:Table (wrapped in its own one:OE).
    Target = an existing one:Table objectID → append ``rows`` to that table in place."""
    if not rows:
        raise ValueError("rows is empty")

    def mutate(tree: etree._Element) -> None:
        outline = None
        if target_object_id:
            target = _find_content_object(tree, target_object_id)
            kind = local_name(target.tag)
            if kind == "Table":
                _append_table_rows(target, rows)
                return
            if kind != "Outline":
                raise ValueError(
                    f"create_table target {target_object_id!r} is a one:{kind} — it must be "
                    "an outline (new table) or an existing table (append rows)"
                )
            outline = target
        else:
            outline = _resolve_outline(tree, "")
        oe = etree.Element(qn("OE"))
        oe.append(make_table(rows, borders_visible, has_header_row, col_widths))
        _outline_children(outline).append(oe)

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
