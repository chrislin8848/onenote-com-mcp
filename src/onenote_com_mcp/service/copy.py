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
from onenote_com_mcp.service.page_edit import inline_image_binaries, parse_onenote_datetime
from onenote_com_mcp.xmllayer.namespaces import local_name, qn

# strip_cdata=False keeps one:T CDATA sections verbatim — byte-level span fidelity.
_PARSER = etree.XMLParser(strip_cdata=False)

# Read-side identity/state that must NOT ride into a transplant: object identity is minted by
# OneNote in the target; stamps and view state belong to the source. (Author attributes stay —
# OneNote tolerates them on input, Phase-4 evidence.)
_STRIP_ATTRS = ("objectID", "lastModifiedTime", "creationTime", "selected", "isCurrentlyViewed")

# Printout bookkeeping on render images — dangling once their one:XPSFile carrier is stripped
# (see _rewrite_inserted_files); the render survives as a plain inlined image.
_PRINTOUT_IMAGE_ATTRS = ("xpsFileIndex", "isPrintOut", "originalPageNumber")


@dataclass
class PageCopyResult:
    page_id: str  # the copy's ID in the target section
    # one line per attachment that could NOT be transferred faithfully (SPEC §5: report
    # explicitly, never skip silently) or whose structure was changed (printout flattening)
    file_notes: list[str] = field(default_factory=list)


@dataclass
class SectionCopyResult:
    section_id: str
    file_notes: list[str] = field(default_factory=list)  # aggregated, prefixed per page


def transfer_page(backend: OneNoteBackend, page_id: str, target_section_id: str) -> PageCopyResult:
    """Faithfully copy one page into a section, returning the new page ID + file notes."""
    # The direct path: RAW source XML — no structured parse. piBinaryData asks for inline
    # binaries; what it doesn't inline (ground truth: usually nothing) is fetched below.
    raw_xml = backend.get_page_content(page_id, PageInfo.piBinaryData)
    return _transplant_raw_page(backend, raw_xml, page_id, target_section_id)


def _rewrite_inserted_files(backend: OneNoteBackend, tree: etree._Element) -> list[str]:
    """Re-point every one:InsertedFile at a staged copy of its cache (SPEC §5 clone rule).

    The source's ``pathCache`` is a dead reference in a clone (OneNote owns it) — it never
    rides along. The cache bytes are copied aside via ``stage_cache_copy`` and ``pathSource``
    re-pointed at the copy so OneNote re-imports it. Printouts are flattened: their page-level
    ``one:XPSFile`` carriers hold read-side CallbackIDs that cannot ride into a write, so the
    carrier and the ``one:Printout`` child are stripped — the rendered pages survive as plain
    inlined images and the source file as a normal attachment. Flattening is treated as expected
    normalization, NOT a fidelity loss (the visible pages are preserved), so it is intentionally
    NOT reported (user decision 2026-06-12). Returns one note per attachment whose CONTENT could
    not be transferred (unavailable cache); an empty list = nothing was dropped.
    """
    notes: list[str] = []
    xps_files = tree.findall(qn("XPSFile"))
    if xps_files:
        for xps in xps_files:
            tree.remove(xps)
        for img in tree.iter(qn("Image")):
            for attr in _PRINTOUT_IMAGE_ATTRS:
                img.attrib.pop(attr, None)
    for f in tree.iter(qn("InsertedFile")):
        name = f.get("preferredName") or "attachment"
        printout = f.find(qn("Printout"))
        if printout is not None:
            # Flatten (strip the printout structure); the rendered pages survive as images, so
            # this is normalization, not a loss — deliberately no note (see docstring).
            f.remove(printout)
        path_cache = f.attrib.pop("pathCache", None)
        staged = backend.stage_cache_copy(path_cache, name) if path_cache else None
        if staged:
            f.set("pathSource", staged)
        elif f.get("pathSource"):
            notes.append(
                f"{name}: source cache unavailable — kept the original pathSource "
                f"({f.get('pathSource')}); OneNote can only re-import it if that path "
                "still exists on this machine"
            )
        else:
            notes.append(
                f"{name}: source cache unavailable and no pathSource — the file content "
                "could not be transferred (only the attachment entry was copied)"
            )
    return notes


def _transplant_raw_page(
    backend: OneNoteBackend, raw_xml: str, source_page_id: str, target_section_id: str
) -> PageCopyResult:
    tree = etree.fromstring(raw_xml.encode("utf-8"), parser=_PARSER)
    source_level = tree.get("pageLevel")

    # 1. pixels: every one:Image gets inline one:Data (callbacks resolve against the SOURCE)
    inline_image_binaries(backend, source_page_id, tree)
    # 1b. attachments: stage cache copies, re-point pathSource, flatten printouts (Phase 5b)
    file_notes = _rewrite_inserted_files(backend, tree)
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

    return PageCopyResult(page_id=new_page_id, file_notes=file_notes)


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
    for page in source.findall(qn("Page")):
        result = transfer_page(backend, page.get("ID"), new_section_id)
        page_name = page.get("name") or page.get("ID")
        file_notes.extend(f"page '{page_name}': {note}" for note in result.file_notes)
    return SectionCopyResult(section_id=new_section_id, file_notes=file_notes)


# NOTE: there is deliberately no transfer_notebook. VM ground truth (2026-06-11): this M365
# OneNote build refuses COM notebook creation — OpenHierarchy(cftNotebook) returns
# hrFileDoesNotExist (0x80042006) for local folder paths AND OneDrive https parents alike.
# Whole-notebook cloning = transfer_section per section into an existing notebook/group.
