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

Implemented now (Linux-testable): the read → mutate → conserve → single-write orchestration.
Deferred to **Phase 4** (needs real VM fixtures + round-trips): the actual reorder/pageLevel/
rename/move mutations. ``move_page`` additionally stays EXPERIMENTAL until cross-section moves
are validated on the VM (SPEC §4/§9); within-notebook moves keep ID conservation at notebook
scope, which is why its facade reads the notebook, not the section.
"""

from __future__ import annotations

from collections.abc import Callable

from lxml import etree

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import HierarchyScope

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


# --- Facades the MCP structure tools delegate to (mutations are Phase 4) -----------
# Same pattern as page_edit: delegation is real and observable now; the NotImplementedError
# lives inside the lazy mutator.


def restructure_section(
    backend: OneNoteBackend, section_id: str, ordered_pages: list[dict]
) -> None:
    """Whole-batch reorder of a section's pages + pageLevel adjustments."""

    def mutate(tree: etree._Element) -> None:
        raise NotImplementedError(
            "Phase 4: reorder the section's complete page list to match ordered_pages "
            "(each entry {page_id, page_level}), setting pageLevel per page"
        )

    apply_hierarchy_restructure(backend, section_id, HierarchyScope.hsPages, mutate)


def reorder_sections(
    backend: OneNoteBackend, notebook_id: str, ordered_section_ids: list[str]
) -> None:
    """Whole-batch reorder of a notebook's sections.

    A notebook's direct children are a MIXED list of ``one:Section`` and ``one:SectionGroup``
    (SPEC §5) — the submitted batch must contain both kinds; the core's ID conservation makes
    dropping the groups impossible. (Notebook-level ordering itself is unvalidated and out of
    scope.)"""

    def mutate(tree: etree._Element) -> None:
        raise NotImplementedError(
            "Phase 4: reorder the notebook's complete child list (one:Section AND "
            "one:SectionGroup, both kinds kept) to ordered_section_ids"
        )

    apply_hierarchy_restructure(backend, notebook_id, HierarchyScope.hsSections, mutate)


def rename_node(backend: OneNoteBackend, parent_id: str, object_id: str, new_name: str) -> None:
    """Rename a page/section. Routed through the full-list core anyway: same-order full
    submission is inference-proof and keeps the single call site."""

    def mutate(tree: etree._Element) -> None:
        raise NotImplementedError(
            "Phase 4: set the name attribute on the node with ID == object_id, order unchanged"
        )

    apply_hierarchy_restructure(backend, parent_id, HierarchyScope.hsPages, mutate)


def move_page(
    backend: OneNoteBackend, notebook_id: str, page_id: str, target_section_id: str
) -> None:
    """EXPERIMENTAL (SPEC §4/§9): cross-section reliability unvalidated until VM round-trips.
    Reads at notebook scope so a within-notebook move conserves the node-ID set."""

    def mutate(tree: etree._Element) -> None:
        raise NotImplementedError(
            "Phase 4 (VM-gated): re-parent the one:Page under the target one:Section, "
            "submitting both sections' complete page lists"
        )

    apply_hierarchy_restructure(backend, notebook_id, HierarchyScope.hsPages, mutate)
