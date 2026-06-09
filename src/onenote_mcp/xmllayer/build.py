"""Build OneNote XML payloads (Phase 1 / Phase 4).

PLACEHOLDER. Implemented in Phase 1 (new-content builders) and exercised by the Phase 4
write tools. See docs/onenote-xml-schema.md.

Hard requirements when implemented (SPEC §5):
  * New ``one:T`` runs carry ``<span style>``; paragraphs set ``quickStyleIndex``/OE style.
  * Highlight writes BOTH ``background:`` and ``mso-highlight:``.
  * Tables emitted as full ``one:Table/one:Row/one:Cell`` structure (cells are OE/T).
  * Editing existing pages mutates the real GetPageContent tree IN PLACE — never rebuild
    from a slimmed model (that path lives in service/, not here).
"""

from __future__ import annotations

from typing import Any


def build_table_xml(rows: list[list[Any]], **opts: Any) -> str:
    """Structured rows → a ``one:Table`` XML fragment."""
    raise NotImplementedError("Phase 1/4: implement against real VM fixtures")


def build_image_xml(data_b64: str, media_type: str, **opts: Any) -> str:
    """base64 image → a ``one:Image`` fragment with inline ``one:Data``."""
    raise NotImplementedError("Phase 1/4: implement against real VM fixtures")


def build_text_oe_xml(runs: list[Any], **opts: Any) -> str:
    """Styled runs → a ``one:OE`` with ``one:T`` spans."""
    raise NotImplementedError("Phase 1/4: implement against real VM fixtures")
