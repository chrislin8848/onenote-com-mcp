"""``OneNoteBackend`` — the COM abstraction boundary (SPEC §2.1, §3).

Every method maps 1:1 to a OneNote Application COM method or property (see
docs/com-api-reference.md).
Higher layers (xml, service, MCP) depend only on this interface, so they are fully testable
on Linux via ``FixtureBackend``. The real ``Win32ComBackend`` lives behind a guarded import.

XML in/out is passed as raw strings; parsing/building belongs to ``onenote_com_mcp.xmllayer``,
not here. This keeps the backend a thin, faithful COM mirror.
"""

from __future__ import annotations

import datetime as _dt
from abc import ABC, abstractmethod
from dataclasses import dataclass

from onenote_com_mcp.enums import (
    CreateFileType,
    HierarchyScope,
    NewPageStyle,
    PageInfo,
    SpecialLocation,
)


@dataclass(frozen=True)
class CurrentWindowIds:
    """The four ``Current*Id`` properties of ``Application.Windows.CurrentWindow``.

    Granularity stops at the page (SPEC §5): there is no COM API for in-page cursor position
    or selected text. Any field may be None depending on what the active window shows
    (e.g. section_group_id is None when the current section sits directly in the notebook).
    """

    notebook_id: str | None
    section_group_id: str | None
    section_id: str | None
    page_id: str | None


class OneNoteBackend(ABC):
    """Abstract OneNote COM surface. Implementations must not leak pywin32 types."""

    # --- Notebook structure -------------------------------------------------

    @abstractmethod
    def get_hierarchy(self, start_node_id: str, scope: HierarchyScope) -> str:
        """Return GetHierarchy XML from ``start_node_id`` down to ``scope``."""

    @abstractmethod
    def update_hierarchy(self, changes_xml: str) -> None:
        """Apply UpdateHierarchy changes (rename, reorder, set pageLevel, ...)."""

    @abstractmethod
    def open_hierarchy(
        self,
        path: str,
        relative_to_object_id: str,
        create_file_type: CreateFileType = CreateFileType.cftNone,
    ) -> str:
        """Open-or-create a hierarchy node; return its object ID."""

    @abstractmethod
    def create_new_page(
        self, section_id: str, style: NewPageStyle = NewPageStyle.npsDefault
    ) -> str:
        """Create a blank page at the end of a section; return the new page ID."""

    @abstractmethod
    def delete_hierarchy(
        self,
        object_id: str,
        expected_last_modified: _dt.datetime | None = None,
        permanent: bool = False,
    ) -> None:
        """Delete a section group / section / page (to recycle bin unless ``permanent``)."""

    @abstractmethod
    def get_hierarchy_parent(self, object_id: str) -> str:
        """Return the parent object ID of a hierarchy node."""

    @abstractmethod
    def get_special_location(self, location: SpecialLocation) -> str:
        """Return a special folder path (backup / unfiled notes / default notebook)."""

    # --- Page content -------------------------------------------------------

    @abstractmethod
    def get_page_content(self, page_id: str, page_info: PageInfo = PageInfo.piBasic) -> str:
        """Return GetPageContent XML (binary inlined only when ``page_info`` requests it)."""

    @abstractmethod
    def update_page_content(
        self,
        changes_xml: str,
        expected_last_modified: _dt.datetime | None = None,
        force: bool = False,
    ) -> None:
        """Merge page-level objects from ``changes_xml`` into the page.

        Guarded by ``expected_last_modified`` (concurrency). ``force`` defaults False and
        should only be set on explicit user opt-in. ``None`` means the implementation
        resolves the page's CURRENT stamp itself (OneNote's COM marshalling cannot express
        "skip the check" — VM ground truth 2026-06-11); callers that read first should pass
        the stamp from that read.
        """

    @abstractmethod
    def get_binary_page_content(self, page_id: str, callback_id: str) -> str:
        """Return a binary page object (image/ink) as a base64 string."""

    @abstractmethod
    def delete_page_content(
        self,
        page_id: str,
        object_id: str,
        expected_last_modified: _dt.datetime | None = None,
        force: bool = False,
    ) -> None:
        """Delete one page content object (outline / image / table) by ID."""

    # --- Attachment cache files (Phase 5b) ------------------------------------
    # InsertedFile content is NOT served by GetBinaryPageContent (that is the image callback
    # path) — it lives in the pathCache file on the OneNote machine's disk. The cache may be
    # missing (unsynced / purged): both methods return None then, never raise for that.
    # Living behind the backend keeps the service layer Linux-testable (FixtureBackend replays).

    @abstractmethod
    def stat_cache_file(self, path: str) -> int | None:
        """Size in bytes of an InsertedFile ``pathCache`` file, or None if unavailable."""

    @abstractmethod
    def read_cache_file(self, path: str) -> bytes | None:
        """Raw bytes of an InsertedFile ``pathCache`` file, or None if unavailable."""

    @abstractmethod
    def stage_cache_copy(self, path: str, preferred_name: str) -> str | None:
        """Copy a ``pathCache`` file to a staging location for re-import; return its path.

        Used by the copy path (SPEC §5): a clone must not carry the source's ``pathCache``
        (a dead reference owned by OneNote) — instead the cache bytes are copied aside and
        the clone's ``pathSource`` points at the copy so OneNote re-imports it. Staged files
        are NOT auto-deleted (OneNote's re-import timing is asynchronous/unknown). Returns
        None when the source cache is unavailable — the caller must report that attachment
        explicitly, never skip it silently.
        """

    # --- Navigation ---------------------------------------------------------

    @abstractmethod
    def get_current_window_ids(self) -> CurrentWindowIds:
        """Return the active window's Current*Id quadruple (``Windows.CurrentWindow``).

        Raises ``NoCurrentWindowError`` when OneNote has no open window — report that
        clearly instead of guessing the user's location (SPEC §5).
        """

    @abstractmethod
    def find_pages(self, start_node_id: str, query: str, include_unindexed: bool = False) -> str:
        """Return FindPages hierarchy XML for pages matching ``query`` under the scope."""

    @abstractmethod
    def get_hyperlink_to_object(self, hierarchy_id: str, object_id: str = "") -> str:
        """Return a onenote:// hyperlink to a node or in-page object."""
