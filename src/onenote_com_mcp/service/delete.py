"""Delete tools (SPEC §5/§6, Phase 6). DESTRUCTIVE — propose-then-confirm at the tool surface.

Two deletes that must not be confused (a §4 contrastive pair):

- ``delete_node`` — a whole HIERARCHY node (notebook / section group / section / page) via
  ``DeleteHierarchy``. Defaults to the recycle bin (``permanent=False``) — the undo net
  (recycle-bin policy, CLAUDE.md). NOT for in-page objects.
- ``delete_page_content`` — ONE PAGE-LEVEL content object (a whole outline, a page-level image,
  a page-level attachment) via ``DeletePageContent``. **Page-level only.** VM ground truth
  (Phase 4/5b): ``DeletePageContent`` ACCEPTS page-level ``one:Outline`` / ``one:Image`` /
  ``one:InsertedFile``, but REFUSES an inline ``one:OE`` — a paragraph OE (0x8004200E) and an
  attachment-bearing OE alike. So inline content (a paragraph, or a table/image/attachment
  *inside* an outline) is removed by editing its outline through ``update_page_content``, not
  here. This module validates the target is page-level BEFORE the COM call, turning that COM
  refusal into a clear, actionable error.

Concurrency: ``delete_page_content`` carries the stamp from the page it reads to validate the
target (matching ``apply_page_edit``'s read-then-act guard); ``force`` defaults False.
``delete_node`` lets the backend resolve the node's current stamp (UpdateHierarchy/DeleteHierarchy
have no optimistic-concurrency guard anyway — the recycle bin is the safety net).
"""

from __future__ import annotations

from lxml import etree

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import PageInfo
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service.page_edit import parse_onenote_datetime
from onenote_com_mcp.xmllayer.namespaces import local_name

# Direct one:Page children that DeletePageContent accepts (VM-verified page-level objects).
# Title is a page child too but is structurally required, never a delete target.
_PAGE_LEVEL_TAGS = frozenset({"Outline", "Image", "InsertedFile", "InkDrawing", "MediaFile"})


def delete_node(backend: OneNoteBackend, object_id: str, *, permanent: bool = False) -> None:
    """Delete a hierarchy node (notebook / section group / section / page).

    Defaults to the recycle bin (``permanent=False``). The backend resolves the node's current
    stamp for the DeleteHierarchy call.
    """
    backend.delete_hierarchy(object_id, permanent=permanent)


def _locate(tree: etree._Element, object_id: str) -> etree._Element | None:
    return next((el for el in tree.iter() if el.get("objectID") == object_id), None)


def _inline_guidance(target: etree._Element) -> str:
    """A specific hint for why an in-page object is not deletable here, by what it is."""
    name = local_name(target.tag)
    if name == "Table":
        what = "a table"
    elif name == "OE":
        # an OE wrapping an image/file/table, or a plain text paragraph
        child = next((local_name(c.tag) for c in target if local_name(c.tag) != "T"), None)
        what = {
            "Image": "an inline image",
            "InsertedFile": "an inline attachment",
            "Table": "a table",
        }.get(child, "a paragraph")
    else:
        what = f"an inline {name}"
    return (
        f"{what} (inline content inside an outline). delete_page_content removes PAGE-LEVEL "
        "objects only — a whole outline, a page-level image, or a page-level attachment. To "
        "remove inline content, edit its outline with update_page_content (replace its text, "
        "or delete the whole containing outline if that is what you want)."
    )


def delete_page_content(
    backend: OneNoteBackend, page_id: str, object_id: str, *, force: bool = False
) -> None:
    """Delete ONE page-level content object (outline / page-level image / page-level attachment).

    Validates the target is a page-level object first: a clear error beats COM's opaque
    0x8004200E refusal for inline OEs (VM ground truth). Raises ``NodeNotFoundError`` if the ID
    is not on the page, ``ValueError`` if it names inline (non-page-level) content.
    """
    xml = backend.get_page_content(page_id, PageInfo.piBasic)
    tree = etree.fromstring(xml.encode("utf-8"))
    target = _locate(tree, object_id)
    if target is None:
        raise NodeNotFoundError(
            f"object {object_id!r} is not on page {page_id!r}. Re-read the page with get_page / "
            "get_page_images / get_page_files_info to get a current objectID."
        )
    parent = target.getparent()
    page_level = (
        parent is not None
        and local_name(parent.tag) == "Page"
        and local_name(target.tag) in _PAGE_LEVEL_TAGS
    )
    if not page_level:
        raise ValueError(f"cannot delete {object_id!r}: it is {_inline_guidance(target)}")
    expected = parse_onenote_datetime(tree.get("lastModifiedTime"))
    backend.delete_page_content(page_id, object_id, expected_last_modified=expected, force=force)
