"""Faithful copy = raw-XML whole-page transfer (SPEC §5 — a SEPARATE path from editing).

``transfer_page`` pulls the source page's RAW XML with image binary inlined
(``piBinaryData``) and transplants it. It deliberately does NOT route through the structured /
``get_page`` parse representation — that representation is for human-readable editing and would
lose fidelity (QuickStyleDef, exact spans). The "copy then modify" workflow is: copy faithfully
here, THEN edit the copy via ``page_edit`` (the editing path).

HARD STRUCTURAL RULE: this module must not import the parse layer. tests/test_copy_path.py
enforces both that rule and that the read uses ``piBinaryData``.
"""

from __future__ import annotations

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import PageInfo


def transfer_page(backend: OneNoteBackend, page_id: str, target_section_id: str) -> str:
    """Faithfully copy one page into a section, returning the new page ID."""
    # The direct path: RAW source XML with image binary inlined — no structured parse.
    raw_xml = backend.get_page_content(page_id, PageInfo.piBinaryData)
    return _transplant_raw_page(backend, raw_xml, target_section_id)


def _transplant_raw_page(backend: OneNoteBackend, raw_xml: str, target_section_id: str) -> str:
    # Phase 5 (needs fixtures + VM round-trip): operate on raw_xml directly —
    #   1. carry the source page's QuickStyleDef table (else quickStyleIndex dangles),
    #   2. reset/strip object IDs so OneNote mints fresh ones in the target,
    #   3. preserve pageLevel,
    #   4. CreateNewPage in target_section_id, then one UpdatePageContent with the body.
    raise NotImplementedError(
        "Phase 5: reset object IDs, carry QuickStyleDef, preserve pageLevel, "
        "create target page and write the transplanted raw body"
    )


def transfer_section(backend: OneNoteBackend, section_id: str, target_notebook_id: str) -> str:
    raise NotImplementedError("Phase 5: create target section, then transfer_page each page")


def transfer_notebook(backend: OneNoteBackend, notebook_id: str, name: str, path: str) -> str:
    raise NotImplementedError(
        "Phase 5: create target notebook (sync-path constraints), then transfer_section each"
    )
