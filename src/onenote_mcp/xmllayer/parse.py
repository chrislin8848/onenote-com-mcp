"""Parse OneNote XML → structured data (Phase 1).

PLACEHOLDER. Implemented in Phase 1 against **real** fixtures dumped from the VM
(see docs/onenote-xml-schema.md — fixtures are not synthetic in this project).

Hard requirements when implemented (SPEC §5 format preservation):
  * Resolve effective run style = QuickStyleDef baseline (via the OE ``quickStyleIndex``)
    overlaid with inline ``<span style>`` overrides inside the ``one:T`` CDATA.
  * Read highlight from either ``background:`` or ``mso-highlight:``.
  * Tables → structured rows (list[list[cell]]), never flattened to a string.
  * Keep references to the raw lxml nodes; never collapse content to plain text.
"""

from __future__ import annotations

from typing import Any


def parse_hierarchy(xml: str) -> list[dict[str, Any]]:
    """GetHierarchy/FindPages XML → list of notebook/section/page dicts (with IDs)."""
    raise NotImplementedError("Phase 1: implement against real VM fixtures")


def parse_page(xml: str) -> Any:
    """GetPageContent XML → structured page (title, outlines, runs+style, tables, images)."""
    raise NotImplementedError("Phase 1: implement against real VM fixtures")
