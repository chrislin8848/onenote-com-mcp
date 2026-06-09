"""The single page-content write path (SPEC §4 convergence point).

EVERY content write — update_page_content / create_table / insert_image — funnels through
``apply_page_edit``, so there is exactly ONE ``UpdatePageContent`` call site and ONE place the
concurrency guard is applied. The MCP tools and the facades below are thin: the content shaping
(the ``mutate`` callback) is theirs; the read → guard → write orchestration is here.

Implemented now (testable on Linux): the orchestration + concurrency guard. Deferred to
**Phase 4** (needs real fixtures + VM round-trip): the actual in-place tree mutations that build
``one:OE``/``one:T``/``one:Table``/``one:Image`` XML — and the open question of whether to send
the whole page or only the changed page-level object (if whole-page, read ``piBinaryData`` so
images aren't stripped by the merge). That choice is confirmed on the VM.
"""

from __future__ import annotations

import datetime as _dt
from collections.abc import Callable

from lxml import etree

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import PageInfo

# A mutation edits the parsed page tree IN PLACE (SPEC §5 — never rebuild from a slimmed model,
# or untouched paragraphs lose their formatting).
Mutator = Callable[[etree._Element], None]


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
    read_info: PageInfo = PageInfo.piBasic,
    force: bool = False,
) -> None:
    """Read the page, mutate its XML tree in place, write it back in ONE guarded call.

    This is the only function in the codebase that calls ``backend.update_page_content`` for an
    edit. The concurrency guard (``dateExpectedLastModified``) is taken from the page we just
    read, so a write is refused if the page changed underneath us; ``force`` defaults False.
    """
    xml = backend.get_page_content(page_id, read_info)
    tree = etree.fromstring(xml.encode("utf-8"))
    expected = parse_onenote_datetime(tree.get("lastModifiedTime"))
    mutate(tree)  # in place; untouched paragraphs keep their quickStyleIndex/spans verbatim
    payload = etree.tostring(tree, xml_declaration=True, encoding="UTF-8").decode("utf-8")
    backend.update_page_content(payload, expected_last_modified=expected, force=force)


# --- Facades the MCP write tools delegate to (content-building is Phase 4) ---------
# Each builds a Mutator and hands it to the single core above. The NotImplementedError lives
# inside the (lazy) mutate so the delegation itself is real and observable now.


def edit_page_content(
    backend: OneNoteBackend,
    page_id: str,
    content: str,
    mode: str = "append",
    *,
    force: bool = False,
) -> None:
    def mutate(tree: etree._Element) -> None:
        raise NotImplementedError(
            "Phase 4: build one:OE/one:T (with span styles) from content and apply by mode "
            "(append/insert/replace) in place"
        )

    apply_page_edit(backend, page_id, mutate, force=force)


def add_table(
    backend: OneNoteBackend, page_id: str, rows: list[list[str]], *, force: bool = False
) -> None:
    def mutate(tree: etree._Element) -> None:
        raise NotImplementedError("Phase 4: build one:Table and append it to an outline")

    apply_page_edit(backend, page_id, mutate, force=force)


def insert_image(
    backend: OneNoteBackend,
    page_id: str,
    image_base64: str,
    media_type: str,
    *,
    force: bool = False,
) -> None:
    def mutate(tree: etree._Element) -> None:
        raise NotImplementedError("Phase 4: build one:Image with inline base64 one:Data")

    apply_page_edit(backend, page_id, mutate, force=force)
