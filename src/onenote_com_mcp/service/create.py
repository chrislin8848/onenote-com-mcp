"""Create tools — service layer (SPEC §4: create_notebook / create_section / create_page).

Notebooks and sections are ``OpenHierarchy`` calls (SPEC §5): a section is ``name.one``
relative to its parent notebook OR section group (it inherits their sync); a notebook needs
an explicit folder path that syncs (e.g. OneDrive) or it ends up local-only — empty path
falls back to OneNote's configured default notebook folder.

``create_page`` composes the existing seams instead of opening new write paths: title +
initial content are ONE ``apply_page_edit`` call (single guarded UpdatePageContent), and a
``page_level`` other than 1 goes through ``apply_hierarchy_restructure`` — the section's
complete page list, order unchanged, only the new page's ``pageLevel`` set (SPEC §5
whole-batch discipline).
"""

from __future__ import annotations

from typing import Any

from lxml import etree

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import CreateFileType, HierarchyScope, SpecialLocation
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service.hierarchy_edit import apply_hierarchy_restructure
from onenote_com_mcp.service.page_edit import apply_page_edit, content_mutator, set_title
from onenote_com_mcp.xmllayer.namespaces import qn

# OneNote rejects these in notebook/section names (section = a .one filename; notebook = a
# folder that also becomes part of the sync URL).
_INVALID_NAME_CHARS = set('\\/:*?"<>|&#%~')


def _checked_name(name: str, kind: str) -> str:
    name = name.strip()
    if not name:
        raise ValueError(f"{kind} name is empty")
    bad = _INVALID_NAME_CHARS & set(name)
    if bad:
        raise ValueError(
            f"{kind} name contains characters OneNote forbids: {' '.join(sorted(bad))}"
        )
    return name


def create_notebook(backend: OneNoteBackend, name: str, path: str = "") -> str:
    """Create a notebook folder at ``path``; empty path = OneNote's default notebook folder
    (synced setups point it at OneDrive). Returns the new notebook ID."""
    name = _checked_name(name, "notebook")
    if not path:
        path = backend.get_special_location(SpecialLocation.slDefaultNotebookFolder)
    full_path = path.rstrip("\\/") + "\\" + name
    return backend.open_hierarchy(full_path, "", CreateFileType.cftNotebook)


def create_section(backend: OneNoteBackend, parent_id: str, name: str) -> str:
    """Create ``name.one`` under a notebook OR section group; returns the new section ID."""
    name = _checked_name(name, "section")
    return backend.open_hierarchy(f"{name}.one", parent_id, CreateFileType.cftSection)


def create_page(
    backend: OneNoteBackend,
    section_id: str,
    title: str = "",
    content: str | list[Any] = "",
    page_level: int = 1,
) -> str:
    """CreateNewPage, then title + initial content in ONE guarded write; returns the page ID."""
    if page_level not in (1, 2, 3):
        raise ValueError(f"page_level must be 1, 2 or 3, got {page_level!r}")
    page_id = backend.create_new_page(section_id)

    if title or content:
        # validate the content contract BEFORE the write round-trip
        append = content_mutator(content, "append") if content else None

        def mutate(tree: etree._Element) -> None:
            if title:
                set_title(tree, title)
            if append:
                append(tree)

        apply_page_edit(backend, page_id, mutate)

    if page_level != 1:

        def set_level(tree: etree._Element) -> None:
            page = next((el for el in tree.iter(qn("Page")) if el.get("ID") == page_id), None)
            if page is None:
                raise NodeNotFoundError(
                    f"new page {page_id!r} is not in the section's hierarchy — "
                    "cannot set its pageLevel"
                )
            page.set("pageLevel", str(page_level))

        apply_hierarchy_restructure(backend, section_id, HierarchyScope.hsPages, set_level)

    return page_id
