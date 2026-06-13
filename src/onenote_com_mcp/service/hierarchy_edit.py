"""The single hierarchy-restructure path (SPEC §5 hierarchy-restructure discipline).

``UpdateHierarchy`` has no position-index attribute: order IS the child-element order of the
submitted XML, and Microsoft documents that submitting a *partial* child list makes OneNote
"infer" placement of the omitted siblings — unpredictably. The discipline therefore is:
**always submit the COMPLETE child list of the affected scope, in target order, in one batch.**

This module structurally enforces that, mirroring ``page_edit``'s write-core convergence:
``apply_hierarchy_restructure`` is the ONLY ``update_hierarchy`` call site for structural
changes — restructure_section / reorder_sections / rename_node / move_page are facades over it.
It reads the full hierarchy of the scope, lets the mutator rearrange/rename/re-nest IN PLACE,
verifies **node-ID conservation** (nothing dropped, nothing invented — the structural encoding
of "complete list"), then submits the whole tree once.

The facades add input-contract checks ON TOP of the seam's conservation backstop: reorders
demand the complete child-ID list up front (better errors than a post-mutation conservation
failure) and respect OneNote's positional schema (sections before groups), recycle-bin nodes —
which the list_* tools deliberately hide — are pinned in place automatically, and ``move_page``
is VM-VALIDATED (2026-06-11; formerly experimental): it reads at notebook scope so the move
conserves the node-ID set and both sections' complete page lists ride in the one batch, and the
moved page comes back under a NEW ID.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable

from lxml import etree

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import HierarchyScope
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service.names import checked_name
from onenote_com_mcp.service.page_edit import apply_page_edit, set_title
from onenote_com_mcp.xmllayer.namespaces import local_name, qn

# A restructure mutation rearranges the hierarchy tree IN PLACE (reorder children, set
# pageLevel, change name attributes, re-parent a page). It must not add or drop nodes.
Restructurer = Callable[[etree._Element], None]


def _node_id_set(tree: etree._Element) -> frozenset[str]:
    """All OneNote object IDs present in the tree (the complete-list invariant currency)."""
    return frozenset(el.get("ID") for el in tree.iter() if el.get("ID"))


def apply_hierarchy_restructure(
    backend: OneNoteBackend,
    scope_id: str,
    scope: HierarchyScope,
    mutate: Restructurer,
) -> None:
    """Read the scope's FULL hierarchy, mutate in place, submit it whole in ONE guarded call.

    The only function in the codebase that calls ``backend.update_hierarchy`` for structural
    changes. Raises ``ValueError`` if the mutation adds/drops nodes — partial child lists are
    exactly what the SPEC forbids (OneNote would "infer" the rest, unpredictably).
    """
    xml = backend.get_hierarchy(scope_id, scope)
    tree = etree.fromstring(xml.encode("utf-8"))
    before = _node_id_set(tree)
    mutate(tree)  # in place; order of child elements = target order
    after = _node_id_set(tree)
    if after != before:
        missing, invented = before - after, after - before
        raise ValueError(
            "hierarchy restructure must submit the complete child list — "
            f"dropped: {sorted(missing)} invented: {sorted(invented)}"
        )
    backend.update_hierarchy(
        etree.tostring(tree, xml_declaration=True, encoding="UTF-8").decode("utf-8")
    )


# --- in-place mutation helpers --------------------------------------------------------


def _find_node(tree: etree._Element, node_id: str) -> etree._Element:
    el = next((e for e in tree.iter() if e.get("ID") == node_id), None)
    if el is None:
        raise NodeNotFoundError(f"no node with ID {node_id!r} in this hierarchy scope")
    return el


def _is_recycle_bin(el: etree._Element) -> bool:
    return el.get("isRecycleBin") == "true"


def _reorder_children(
    parent: etree._Element, kinds: frozenset[str], ordered_ids: Iterable[str], what: str
) -> None:
    """Reorder ``parent``'s direct children of the given kinds into ``ordered_ids``.

    ``ordered_ids`` must be the COMPLETE list of those children (SPEC §5) — except
    recycle-bin nodes, which the list_* tools hide and which therefore keep their positions
    automatically. Everything still rides in the submitted batch; only the order changes.
    """

    def reorderable(el: etree._Element) -> bool:
        return local_name(el.tag) in kinds and not _is_recycle_bin(el)

    by_id = {c.get("ID"): c for c in parent if reorderable(c)}
    want = list(ordered_ids)
    if len(want) != len(set(want)):
        raise ValueError(f"{what}: ordered IDs contain duplicates")
    if set(want) != set(by_id):
        missing = sorted(set(by_id) - set(want))
        unknown = sorted(set(want) - set(by_id))
        raise ValueError(
            f"{what} needs the COMPLETE child list in target order (SPEC §5) — "
            f"missing: {missing} unknown: {unknown}"
        )
    replacements = iter(by_id[node_id] for node_id in want)
    sequence = [next(replacements) if reorderable(c) else c for c in list(parent)]
    for child in sequence:
        parent.append(child)  # appending an existing child MOVES it; final order = sequence


def _require_sections_before_groups(container: etree._Element, what: str) -> None:
    """OneNote's hierarchy schema is POSITIONAL: at any level every ``one:Section`` must come
    before all ``one:SectionGroup`` siblings (the UI renders groups last). Submitting an
    interleaved order is rejected with hrInvalidXML — VM ground truth 2026-06-11. Fail with
    a contract error before any COM write instead."""
    seen_group = False
    for child in container:
        kind = local_name(child.tag)
        if kind == "SectionGroup":
            seen_group = True
        elif kind == "Section" and seen_group:
            raise ValueError(
                f"{what}: OneNote requires every section to come before all section groups "
                "at the same level — sections can only be reordered among sections, and "
                "groups among groups"
            )


# --- Facades the MCP structure tools delegate to ---------------------------------------


def restructure_section(
    backend: OneNoteBackend, section_id: str, ordered_pages: list[dict]
) -> None:
    """Whole-batch reorder of a section's pages + pageLevel adjustments.

    ``ordered_pages`` must be the section's COMPLETE page list in target order, each entry
    ``{"page_id": str, "page_level": 1|2|3}`` (page_level omitted = keep the current one)."""
    entries: list[tuple[str, int | None]] = []
    for entry in ordered_pages:
        if not isinstance(entry, dict) or not entry.get("page_id"):
            raise ValueError(
                'each ordered_pages entry is {"page_id": str, "page_level": 1|2|3 (optional)}'
            )
        level = entry.get("page_level")
        if level is not None and level not in (1, 2, 3):
            raise ValueError(f"page_level must be 1, 2 or 3, got {level!r}")
        entries.append((entry["page_id"], level))

    def mutate(tree: etree._Element) -> None:
        section = _find_node(tree, section_id)
        _reorder_children(
            section, frozenset({"Page"}), [pid for pid, _ in entries], "restructure_section"
        )
        for pid, level in entries:
            if level is not None:
                _find_node(section, pid).set("pageLevel", str(level))

    apply_hierarchy_restructure(backend, section_id, HierarchyScope.hsPages, mutate)


def reposition_page(
    backend: OneNoteBackend,
    section_id: str,
    page_id: str,
    after_page_id: str = "",
    page_level: int | None = None,
) -> None:
    """Move ONE page to a new position WITHIN its section — right after ``after_page_id`` (empty →
    the top of the section) — and optionally set its ``pageLevel``.

    The lightweight alternative to restructure_section for the common "put this page here" case:
    the caller supplies only IDs, and the seam's whole-tree read + node-ID conservation provide
    the SPEC §5 completeness (nothing is dropped — the full page list still rides in the one batch;
    only this page's position changes). Same-section only — moving a page to a DIFFERENT section is
    move_page. lxml ``addnext``/``addprevious`` MOVE the existing element, so no node is added."""
    if page_level is not None and page_level not in (1, 2, 3):
        raise ValueError(f"page_level must be 1, 2 or 3, got {page_level!r}")
    if after_page_id and after_page_id == page_id:
        raise ValueError("after_page_id cannot be the page being moved")

    def mutate(tree: etree._Element) -> None:
        section = _find_node(tree, section_id)
        page = _find_node(tree, page_id)
        if local_name(page.tag) != "Page":
            raise ValueError(f"{page_id!r} is a one:{local_name(page.tag)}, not a page")
        if page.getparent() is not section:
            raise ValueError(
                f"page {page_id!r} is not directly in section {section_id!r} — "
                "to move a page to another section use move_page"
            )
        if after_page_id:
            anchor = _find_node(tree, after_page_id)
            if local_name(anchor.tag) != "Page" or anchor.getparent() is not section:
                raise ValueError(
                    f"after_page_id {after_page_id!r} must be a page in section {section_id!r}"
                )
            anchor.addnext(page)  # MOVE page to immediately after the anchor
        else:
            first = section.find(qn("Page"))
            if first is not None and first is not page:
                first.addprevious(page)  # MOVE page to the top of the section
        if page_level is not None:
            page.set("pageLevel", str(page_level))

    apply_hierarchy_restructure(backend, section_id, HierarchyScope.hsPages, mutate)


def reorder_sections(
    backend: OneNoteBackend, notebook_id: str, ordered_section_ids: list[str]
) -> None:
    """Whole-batch reorder of a notebook's (or section group's) direct children.

    A notebook's direct children are a MIXED list of ``one:Section`` and ``one:SectionGroup``
    (SPEC §5): ``ordered_section_ids`` must contain BOTH kinds — every ID ``list_sections``
    shows at that level. The recycle-bin group (hidden from listings) is pinned in place
    automatically but still rides in the batch. Notebook-level ordering is out of scope."""

    def mutate(tree: etree._Element) -> None:
        container = _find_node(tree, notebook_id)
        _reorder_children(
            container,
            frozenset({"Section", "SectionGroup"}),
            ordered_section_ids,
            "reorder_sections",
        )
        _require_sections_before_groups(container, "reorder_sections")

    apply_hierarchy_restructure(backend, notebook_id, HierarchyScope.hsSections, mutate)


_RENAMABLE = frozenset({"Page", "Section", "SectionGroup"})


def rename_node(backend: OneNoteBackend, parent_id: str, object_id: str, new_name: str) -> None:
    """Rename a page / section / section group.

    Pages: ``UpdateHierarchy`` silently IGNORES a ``one:Page`` name attribute — the
    hierarchy name is derived from the page title (VM ground truth 2026-06-11) — so a page
    rename IS a title edit, routed through the page-content seam (``apply_page_edit``).
    Sections/groups: routed through the full-list hierarchy core (same-order full
    submission is inference-proof and keeps the single call site)."""
    scope_tree = etree.fromstring(
        backend.get_hierarchy(parent_id, HierarchyScope.hsPages).encode("utf-8")
    )
    kind = local_name(_find_node(scope_tree, object_id).tag)
    if kind == "Page":
        name = new_name.strip()
        if not name:
            raise ValueError("page name is empty")
        apply_page_edit(backend, object_id, lambda tree: set_title(tree, name))
        return
    if kind not in _RENAMABLE:
        raise ValueError(f"cannot rename a one:{kind} — only pages, sections, and section groups")
    label = "section" if kind == "Section" else "section group"
    name = checked_name(new_name, label)  # section names are .one filenames

    def mutate(tree: etree._Element) -> None:
        _find_node(tree, object_id).set("name", name)

    apply_hierarchy_restructure(backend, parent_id, HierarchyScope.hsPages, mutate)


def move_page(
    backend: OneNoteBackend, notebook_id: str, page_id: str, target_section_id: str
) -> str:
    """Move a page to another section in the same notebook; returns the page's NEW ID.

    VM ground truth (2026-06-11): page IDs embed the owning section, so re-parenting
    assigns the page a NEW ID — callers must use the returned ID from then on. Reads at
    notebook scope so the move conserves the node-ID set and BOTH sections' complete page
    lists ride in the one batch. The moved page's ``pageLevel`` resets to 1 — a subpage
    moved alone must not dangle under a parent that stayed behind."""

    def mutate(tree: etree._Element) -> None:
        page = _find_node(tree, page_id)
        if local_name(page.tag) != "Page":
            raise ValueError(f"{page_id!r} is a one:{local_name(page.tag)}, not a page")
        try:
            target = _find_node(tree, target_section_id)
        except NodeNotFoundError:
            raise NodeNotFoundError(
                f"target section {target_section_id!r} is not in this notebook — "
                "cross-notebook moves are unsupported; use copy_page + delete_node"
            ) from None
        if local_name(target.tag) != "Section":
            raise ValueError(
                f"target {target_section_id!r} is a one:{local_name(target.tag)}, not a section"
            )
        if _is_recycle_bin(target) or any(_is_recycle_bin(a) for a in target.iterancestors()):
            raise ValueError(
                "target section is in the recycle bin — use delete_node to delete a page"
            )
        target.append(page)  # lxml MOVES the element: source loses it, target gains it
        page.set("pageLevel", "1")

    apply_hierarchy_restructure(backend, notebook_id, HierarchyScope.hsPages, mutate)
    # the page landed LAST under the target section, under a NEW ID — re-read to learn it
    after = etree.fromstring(
        backend.get_hierarchy(target_section_id, HierarchyScope.hsPages).encode("utf-8")
    )
    pages = _find_node(after, target_section_id).findall(qn("Page"))
    if not pages:
        raise NodeNotFoundError(
            f"move_page: target section {target_section_id!r} shows no pages after the move"
        )
    return pages[-1].get("ID")
