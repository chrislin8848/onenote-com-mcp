"""Faithful copy = raw-XML whole-page transfer (SPEC §5 — a SEPARATE path from editing).

``transfer_page`` pulls the source page's RAW XML and transplants it: image binaries inlined
(``GetBinaryPageContent`` — even a piBinaryData read serves only CallbackID, VM ground truth),
the QuickStyleDef table carried along verbatim, all object identity stripped so OneNote mints
fresh IDs in the target, ``pageLevel`` preserved through the hierarchy seam. It deliberately
does NOT route through the structured / ``get_page`` parse representation — that representation
is for human-readable editing and would lose fidelity (QuickStyleDef, exact spans). The "copy
then modify" workflow is: copy faithfully here, THEN edit the copy via ``page_edit``.

HARD STRUCTURAL RULE: this module must not import the parse layer. tests/test_copy_path.py
enforces both that rule and that the source read uses ``piBinaryData``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from lxml import etree

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import HierarchyScope, NewPageStyle, PageInfo
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service.create import create_section
from onenote_com_mcp.service.hierarchy_edit import apply_hierarchy_restructure
from onenote_com_mcp.service.page_edit import (
    inline_image_binaries,
    parse_onenote_datetime,
    remove_content_element,
)
from onenote_com_mcp.xmllayer.namespaces import local_name, qn

# strip_cdata=False keeps one:T CDATA sections verbatim — byte-level span fidelity.
_PARSER = etree.XMLParser(strip_cdata=False)

# Read-side identity/state that must NOT ride into a transplant: object identity is minted by
# OneNote in the target; stamps and view state belong to the source. (Author attributes stay —
# OneNote tolerates them on input, Phase-4 evidence.)
_STRIP_ATTRS = ("objectID", "lastModifiedTime", "creationTime", "selected", "isCurrentlyViewed")

# Printout bookkeeping on render images — dangling once their one:XPSFile carrier is stripped
# (see _rewrite_inserted_files); the render survives as a plain inlined image. Stripped from EVERY
# image, carrier or not: OneNote can dissolve a printout's link on its own (VM 2026-10-05 — after an
# Office update the PDF lost its Printout child and the XPSFile vanished, leaving an ORPHAN render
# still marked isPrintOut), and a copy never carries a printout structure.
_PRINTOUT_IMAGE_ATTRS = ("xpsFileIndex", "isPrintOut", "originalPageNumber")


@dataclass
class PageCopyResult:
    page_id: str  # the copy's ID in the target section
    name: str | None = None  # the copy's title (== source's) — so callers report by NAME, not ID
    page_level: int | None = None  # 1/2/3, preserved from the source
    # one line per piece of content that could NOT be transferred faithfully (SPEC §5: report
    # explicitly, never skip silently) or whose structure was changed (printout flattening)
    file_notes: list[str] = field(default_factory=list)
    # counts of content the copy could not carry because it is not yet downloaded to THIS
    # machine (OneDrive files-on-demand) — fed into the user-facing sync_warning below
    missing_images: int = 0
    missing_files: int = 0
    missing_objects: int = 0


@dataclass
class PagesCopyResult:
    page_ids: list[str]  # the copies' IDs in the target section, in copy order (for placement)
    # the copies as {page_id, name, page_level}, in copy order — so the caller can report results
    # by NAME (avoids dumping raw ID lists to the user); parallel to page_ids
    pages: list[dict] = field(default_factory=list)
    file_notes: list[str] = field(default_factory=list)
    missing_images: int = 0
    missing_files: int = 0
    missing_objects: int = 0


@dataclass
class SectionCopyResult:
    section_id: str
    file_notes: list[str] = field(default_factory=list)  # aggregated, prefixed per page
    missing_images: int = 0
    missing_files: int = 0
    missing_objects: int = 0


def sync_warning(missing_images: int, missing_files: int, missing_objects: int) -> str | None:
    """A user-facing warning when a copy hit content not yet downloaded to this machine.

    A failed image fetch (0x8004200F) or an unavailable attachment cache during copy means
    OneDrive has not hydrated that content locally (files-on-demand). Unlike the live original
    (which OneNote fills in on demand), the copied placeholders/empty references do NOT
    self-heal. The remedy is to fully sync the source first — so we surface this, never silently
    ship blanks (SPEC §5). Returns None when nothing was missing."""
    if not (missing_images or missing_files or missing_objects):
        return None
    parts = []
    if missing_images:
        parts.append(f"{missing_images} image(s)")
    if missing_files:
        parts.append(f"{missing_files} file(s)")
    if missing_objects:
        parts.append(f"{missing_objects} embedded object(s)")
    return (
        "Source is NOT fully synced: " + " + ".join(parts) + " could not be copied and were "
        "OMITTED from the copy — their content is not yet downloaded to this machine and would not "
        "self-heal if copied as placeholders (unlike the live original). Open and fully sync the "
        "source section in OneNote, then run the copy again for a faithful copy."
    )


def transfer_page(backend: OneNoteBackend, page_id: str, target_section_id: str) -> PageCopyResult:
    """Faithfully copy one page into a section, returning the new page ID + file notes."""
    # The direct path: RAW source XML — no structured parse. piBinaryData asks for inline
    # binaries; what it doesn't inline (ground truth: usually nothing) is fetched below.
    raw_xml = backend.get_page_content(page_id, PageInfo.piBinaryData)
    return _transplant_raw_page(backend, raw_xml, page_id, target_section_id)


def _rewrite_inserted_files(
    backend: OneNoteBackend, tree: etree._Element
) -> tuple[int, int, list[str]]:
    """Re-point every one:InsertedFile at a staged copy of its cache (SPEC §5 clone rule).

    The source's ``pathCache`` is a dead reference in a clone (OneNote owns it) — it never
    rides along. The cache bytes are copied aside via ``stage_cache_copy`` and ``pathSource``
    re-pointed at the copy so OneNote re-imports it. Printouts are flattened: their page-level
    ``one:XPSFile`` carriers hold read-side CallbackIDs that cannot ride into a write, so the
    carrier and the ``one:Printout`` child are stripped — the rendered pages survive as plain
    inlined images and the source file as a normal attachment. Flattening is treated as expected
    normalization, NOT a fidelity loss (the visible pages are preserved), so it is intentionally
    NOT reported (user decision 2026-06-12).

    Returns ``(missing_files, missing_objects, notes)``: counts of attachments (file icons /
    printouts) and embedded objects (e.g. Excel, marked by a ``one:Previews`` child) whose
    content could not be carried because the cache is unavailable locally — VM-confirmed cause
    is an under-synced source (OneDrive files-on-demand), NOT OCR. ``(0, 0, [])`` = all carried.
    """
    notes: list[str] = []
    missing_files = 0
    missing_objects = 0
    for xps in tree.findall(qn("XPSFile")):
        tree.remove(xps)
    for img in tree.iter(qn("Image")):
        for attr in _PRINTOUT_IMAGE_ATTRS:
            img.attrib.pop(attr, None)
    for f in list(tree.iter(qn("InsertedFile"))):
        name = f.get("preferredName") or "attachment"
        is_embedded = f.find(qn("Previews")) is not None  # embedded object vs plain file icon
        printout = f.find(qn("Printout"))
        if printout is not None:
            # Flatten (strip the printout structure); the rendered pages survive as images, so
            # this is normalization, not a loss — deliberately no note (see docstring).
            f.remove(printout)
        path_cache = f.attrib.pop("pathCache", None)
        staged = backend.stage_cache_copy(path_cache, name) if path_cache else None
        if staged:
            f.set("pathSource", staged)
            continue
        # cache not available locally → the content is not synced to this machine. A dead
        # reference can't self-heal and could later be misread / re-copied, so REMOVE it (pruning
        # any emptied OE) and report — the user syncs the source, then re-copies for fidelity.
        if is_embedded:
            missing_objects += 1
            kind = "embedded object"
        else:
            missing_files += 1
            kind = "file"
        remove_content_element(f)
        notes.append(
            f"{name}: {kind} content not downloaded to this machine (source not fully synced) "
            "— removed from the copy; sync the source, then copy again for a faithful copy"
        )
    return missing_files, missing_objects, notes


def _transplant_raw_page(
    backend: OneNoteBackend, raw_xml: str, source_page_id: str, target_section_id: str
) -> PageCopyResult:
    tree = etree.fromstring(raw_xml.encode("utf-8"), parser=_PARSER)
    source_level = tree.get("pageLevel")
    source_name = tree.get("name")  # the page title (== the copy's name); for name-not-ID reporting

    # 1. pixels: every one:Image gets inline one:Data (callbacks resolve against the SOURCE).
    #    Images whose binary isn't downloaded locally are REMOVED (+ emptied OE pruned) and counted
    #    — a copy must not carry a dead placeholder (it can't self-heal, could be misread later).
    missing_images = inline_image_binaries(backend, source_page_id, tree, remove_unfetchable=True)
    # 1b. attachments: stage cache copies, re-point pathSource, flatten printouts (Phase 5b)
    missing_files, missing_objects, file_notes = _rewrite_inserted_files(backend, tree)
    if missing_images:
        file_notes.insert(
            0,
            f"{missing_images} image(s) not downloaded to this machine (source not fully synced) "
            "— removed from the copy; sync the source, then copy again for a faithful copy",
        )
    # 2. reset identity/state — QuickStyleDef/TagDef tables, spans, tables, author attrs all
    #    stay verbatim; only IDs/stamps/view-state go
    for el in tree.iter():
        for attr in _STRIP_ATTRS:
            el.attrib.pop(attr, None)
    tree.attrib.pop("dateTime", None)
    tree.attrib.pop("pageLevel", None)  # hierarchy-level, set via the hierarchy seam below

    # 3. transplant: blank target page, then ONE UpdatePageContent with the full body.
    #    The expected stamp is the BLANK page's (the source stamp was stripped above and
    #    must not leak into the guard).
    new_page_id = backend.create_new_page(target_section_id, NewPageStyle.npsBlankPageNoTitle)
    blank_tree = etree.fromstring(
        backend.get_page_content(new_page_id, PageInfo.piBasic).encode("utf-8")
    )
    tree.set("ID", new_page_id)
    etree.cleanup_namespaces(tree)
    payload = etree.tostring(tree, xml_declaration=True, encoding="UTF-8").decode("utf-8")
    backend.update_page_content(
        payload, expected_last_modified=parse_onenote_datetime(blank_tree.get("lastModifiedTime"))
    )

    # 4. pageLevel rides the hierarchy, not the page XML (SPEC §5: whole-batch discipline)
    if source_level and source_level != "1":

        def set_level(htree: etree._Element) -> None:
            page = next((el for el in htree.iter(qn("Page")) if el.get("ID") == new_page_id), None)
            if page is None:
                raise NodeNotFoundError(
                    f"copied page {new_page_id!r} is not in the target section's hierarchy"
                )
            page.set("pageLevel", source_level)

        apply_hierarchy_restructure(backend, target_section_id, HierarchyScope.hsPages, set_level)

    return PageCopyResult(
        page_id=new_page_id,
        name=source_name,
        page_level=int(source_level) if source_level else 1,
        file_notes=file_notes,
        missing_images=missing_images,
        missing_files=missing_files,
        missing_objects=missing_objects,
    )


def _find_node(tree: etree._Element, node_id: str) -> etree._Element:
    el = next((e for e in tree.iter() if e.get("ID") == node_id), None)
    if el is None:
        raise NodeNotFoundError(f"no node with ID {node_id!r} in this hierarchy scope")
    return el


def _unique_child_name(backend: OneNoteBackend, parent_id: str, name: str) -> str:
    """De-collide a section/group name among the target parent's direct children.

    ``OpenHierarchy`` OPENS an existing same-named node instead of creating one — a copy
    would silently merge into it. Existing name → "name (2)", "name (3)", …"""
    tree = etree.fromstring(
        backend.get_hierarchy(parent_id, HierarchyScope.hsSections).encode("utf-8")
    )
    parent = _find_node(tree, parent_id)
    taken = {
        child.get("name")
        for child in parent
        if local_name(child.tag) in ("Section", "SectionGroup")
    }
    if name not in taken:
        return name
    n = 2
    while f"{name} ({n})" in taken:
        n += 1
    return f"{name} ({n})"


def transfer_section(
    backend: OneNoteBackend, section_id: str, target_parent_id: str
) -> SectionCopyResult:
    """Faithfully copy a whole section into a notebook OR section group.

    The new section takes the source's name (de-collided); pages are copied in document
    order via ``transfer_page``, each keeping its ``pageLevel`` (subpage nesting survives).
    Returns the new section ID plus per-page attachment notes (prefixed with the page name)."""
    tree = etree.fromstring(
        backend.get_hierarchy(section_id, HierarchyScope.hsPages).encode("utf-8")
    )
    source = _find_node(tree, section_id)
    name = _unique_child_name(backend, target_parent_id, source.get("name") or "Section")
    new_section_id = create_section(backend, target_parent_id, name)
    file_notes: list[str] = []
    missing_images = missing_files = missing_objects = 0
    for page in source.findall(qn("Page")):
        result = transfer_page(backend, page.get("ID"), new_section_id)
        page_name = page.get("name") or page.get("ID")
        file_notes.extend(f"page '{page_name}': {note}" for note in result.file_notes)
        missing_images += result.missing_images
        missing_files += result.missing_files
        missing_objects += result.missing_objects
    return SectionCopyResult(
        section_id=new_section_id,
        file_notes=file_notes,
        missing_images=missing_images,
        missing_files=missing_files,
        missing_objects=missing_objects,
    )


def copy_pages(
    backend: OneNoteBackend, page_ids: list[str], target_section_id: str
) -> PagesCopyResult:
    """Faithfully copy SEVERAL pages into a section, IN ORDER, returning their new IDs + notes.

    Each page is cloned via ``transfer_page`` (so every copy keeps its own ``pageLevel`` — a
    subpage stays a subpage). The copies land at the section END in copy order; gathering them
    into a contiguous block at a chosen position is the caller's job (the server facade chains
    ``hierarchy_edit.reposition_pages`` once). This is the engine under copy_page_subtree and the
    copy_pages tool — the fix for "copy these pages as a block somewhere", which a string of
    single copy_page calls placed unpredictably (each landed below its own source)."""
    if not page_ids:
        raise ValueError("page_ids is empty")
    if len(page_ids) != len(set(page_ids)):
        raise ValueError("page_ids contains duplicate IDs")
    new_ids: list[str] = []
    pages: list[dict] = []
    file_notes: list[str] = []
    missing_images = missing_files = missing_objects = 0
    for pid in page_ids:
        result = transfer_page(backend, pid, target_section_id)
        new_ids.append(result.page_id)
        pages.append(
            {"page_id": result.page_id, "name": result.name, "page_level": result.page_level}
        )
        file_notes.extend(result.file_notes)
        missing_images += result.missing_images
        missing_files += result.missing_files
        missing_objects += result.missing_objects
    return PagesCopyResult(
        page_ids=new_ids,
        pages=pages,
        file_notes=file_notes,
        missing_images=missing_images,
        missing_files=missing_files,
        missing_objects=missing_objects,
    )


def subtree_page_ids(backend: OneNoteBackend, section_id: str, page_id: str) -> list[str]:
    """The page plus its subpages, in document order: ``page_id`` followed by the consecutive
    pages whose ``pageLevel`` is DEEPER than it, stopping at the next page at the same-or-shallower
    level. A page with no subpages (or one that is itself a subpage) returns just ``[page_id]``.

    This is OneNote's positional subpage model — a page "owns" the more-indented pages that
    follow it until the indentation steps back out. copy_page_subtree uses it to turn "copy this
    page and its subpages" into the explicit ordered list that ``copy_pages`` clones."""
    tree = etree.fromstring(
        backend.get_hierarchy(section_id, HierarchyScope.hsPages).encode("utf-8")
    )
    section = _find_node(tree, section_id)
    pages = section.findall(qn("Page"))
    ids = [p.get("ID") for p in pages]
    levels = [int(p.get("pageLevel") or "1") for p in pages]
    if page_id not in ids:
        raise NodeNotFoundError(f"page {page_id!r} is not directly in section {section_id!r}")
    start = ids.index(page_id)
    base = levels[start]
    out = [ids[start]]
    for j in range(start + 1, len(pages)):
        if levels[j] > base:
            out.append(ids[j])
        else:
            break
    return out


# NOTE: there is deliberately no transfer_notebook. VM ground truth (2026-06-11): this M365
# OneNote build refuses COM notebook creation — OpenHierarchy(cftNotebook) returns
# hrFileDoesNotExist (0x80042006) for local folder paths AND OneDrive https parents alike.
# Whole-notebook cloning = transfer_section per section into an existing notebook/group.
