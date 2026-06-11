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

from lxml import etree

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import HierarchyScope, NewPageStyle, PageInfo
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service.hierarchy_edit import apply_hierarchy_restructure
from onenote_com_mcp.service.page_edit import inline_image_binaries, parse_onenote_datetime
from onenote_com_mcp.xmllayer.namespaces import qn

# strip_cdata=False keeps one:T CDATA sections verbatim — byte-level span fidelity.
_PARSER = etree.XMLParser(strip_cdata=False)

# Read-side identity/state that must NOT ride into a transplant: object identity is minted by
# OneNote in the target; stamps and view state belong to the source. (Author attributes stay —
# OneNote tolerates them on input, Phase-4 evidence.)
_STRIP_ATTRS = ("objectID", "lastModifiedTime", "creationTime", "selected", "isCurrentlyViewed")


def transfer_page(backend: OneNoteBackend, page_id: str, target_section_id: str) -> str:
    """Faithfully copy one page into a section, returning the new page ID."""
    # The direct path: RAW source XML — no structured parse. piBinaryData asks for inline
    # binaries; what it doesn't inline (ground truth: usually nothing) is fetched below.
    raw_xml = backend.get_page_content(page_id, PageInfo.piBinaryData)
    return _transplant_raw_page(backend, raw_xml, page_id, target_section_id)


def _transplant_raw_page(
    backend: OneNoteBackend, raw_xml: str, source_page_id: str, target_section_id: str
) -> str:
    tree = etree.fromstring(raw_xml.encode("utf-8"), parser=_PARSER)
    source_level = tree.get("pageLevel")

    # 1. pixels: every one:Image gets inline one:Data (callbacks resolve against the SOURCE)
    inline_image_binaries(backend, source_page_id, tree)
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

    return new_page_id


def transfer_section(backend: OneNoteBackend, section_id: str, target_notebook_id: str) -> str:
    raise NotImplementedError("Phase 5 Stage 2: create target section, then transfer_page each")


def transfer_notebook(backend: OneNoteBackend, notebook_id: str, name: str, path: str) -> str:
    raise NotImplementedError(
        "Phase 5 Stage 2: create target notebook (sync-path constraints), recreate each source "
        "section group via OpenHierarchy(cftFolder) so sections land INSIDE their groups "
        "(never flattened — SPEC §5), SKIP recycle-bin groups, then transfer_section each"
    )
