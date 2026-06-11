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
from onenote_com_mcp.enums import CreateFileType, HierarchyScope, NewPageStyle, PageInfo
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service.create import create_notebook, create_section
from onenote_com_mcp.service.hierarchy_edit import apply_hierarchy_restructure
from onenote_com_mcp.service.page_edit import inline_image_binaries, parse_onenote_datetime
from onenote_com_mcp.xmllayer.namespaces import local_name, qn

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


def transfer_section(backend: OneNoteBackend, section_id: str, target_parent_id: str) -> str:
    """Faithfully copy a whole section into a notebook OR section group.

    The new section takes the source's name (de-collided); pages are copied in document
    order via ``transfer_page``, each keeping its ``pageLevel`` (subpage nesting survives).
    Returns the new section ID."""
    tree = etree.fromstring(
        backend.get_hierarchy(section_id, HierarchyScope.hsPages).encode("utf-8")
    )
    source = _find_node(tree, section_id)
    name = _unique_child_name(backend, target_parent_id, source.get("name") or "Section")
    new_section_id = create_section(backend, target_parent_id, name)
    for page in source.findall(qn("Page")):
        transfer_page(backend, page.get("ID"), new_section_id)
    return new_section_id


def transfer_notebook(backend: OneNoteBackend, notebook_id: str, name: str, path: str) -> str:
    """Faithfully copy a whole notebook (subject to create_notebook's sync-path constraints).

    Section groups are recreated via ``OpenHierarchy(cftFolder)`` so sections land INSIDE
    their groups, never flattened (SPEC §5). Recycle-bin groups are SKIPPED — the
    user-approved recycle-bin policy: never clone another notebook's wastebasket."""
    new_notebook_id = create_notebook(backend, name, path)
    tree = etree.fromstring(
        backend.get_hierarchy(notebook_id, HierarchyScope.hsSections).encode("utf-8")
    )
    _transfer_children(backend, _find_node(tree, notebook_id), new_notebook_id)
    return new_notebook_id


def _transfer_children(
    backend: OneNoteBackend, container: etree._Element, target_parent_id: str
) -> None:
    """Recreate a container's mixed Section + SectionGroup children in document order."""
    for child in container:
        kind = local_name(child.tag)
        if kind == "Section":
            transfer_section(backend, child.get("ID"), target_parent_id)
        elif kind == "SectionGroup":
            if child.get("isRecycleBin") == "true":
                continue  # policy: the recycle bin never rides along on a clone
            group_name = _unique_child_name(backend, target_parent_id, child.get("name") or "Group")
            group_id = backend.open_hierarchy(
                group_name, target_parent_id, CreateFileType.cftFolder
            )
            _transfer_children(backend, child, group_id)
