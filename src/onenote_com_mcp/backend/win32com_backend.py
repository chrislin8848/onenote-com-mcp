"""``Win32ComBackend`` — live OneNote via in-process pywin32 COM (Windows only).

GUARDED IMPORT INVARIANT (SPEC §2.1): this module must ``import`` cleanly on Linux. Therefore
**no** ``win32com`` / ``pywintypes`` / ``pythoncom`` import appears at module top level — they
are imported lazily inside methods. ``tests/test_smoke_import.py`` enforces this.

⚠ PHASE 3 (VM): the exact pywin32 out-parameter marshalling for OneNote is confirmed on the
VM. This module assumes early binding (``gencache.EnsureDispatch``) where ``[out] BSTR`` params
are returned as the call's return value. If the generated typelib differs, adjust the call
sites here — the rest of the codebase depends only on the string in/out contract, not on this.

Until then, every method body below is written to the documented signature but is
**unvalidated against a real OneNote**.
"""

from __future__ import annotations

import datetime as _dt
import time

from onenote_com_mcp.backend.base import CurrentWindowIds, OneNoteBackend
from onenote_com_mcp.enums import (
    CreateFileType,
    HierarchyScope,
    NewPageStyle,
    PageInfo,
    SpecialLocation,
    XMLSchema,
)
from onenote_com_mcp.errors import (
    BackendUnavailableError,
    NoCurrentWindowError,
    OneNoteComError,
    OneNoteError,
    is_retryable_hresult,
)

_MAX_RETRIES = 6
_BASE_DELAY_S = 0.25


def _hresult_of(exc: Exception) -> int | None:
    """Pull the HRESULT out of a pywintypes.com_error, if present."""
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0]
    return getattr(exc, "hresult", None)


class Win32ComBackend(OneNoteBackend):
    def __init__(self) -> None:
        self._app = None  # lazily created COM object

    # --- connection ---------------------------------------------------------

    @property
    def app(self):
        """Lazily connect to the running OneNote application via COM."""
        if self._app is None:
            try:
                # Imported here, never at module top, to keep Linux import clean.
                import win32com.client  # noqa: PLC0415

                # EnsureDispatch = early binding (typed). For a frozen build the gen cache
                # must be bundled (SPEC §8); see PyInstaller spec when packaging.
                self._app = win32com.client.gencache.EnsureDispatch("OneNote.Application")
            except Exception as exc:  # noqa: BLE001
                raise BackendUnavailableError(
                    "Could not connect to OneNote via COM. Ensure the OneNote desktop app "
                    "is installed and running in this interactive session."
                ) from exc
        return self._app

    def _call(self, name: str, fn):
        """Invoke a COM call with retry/backoff on 'OneNote is busy' HRESULTs."""
        last: Exception | None = None
        for attempt in range(_MAX_RETRIES):
            try:
                return fn()
            except OneNoteError:
                raise  # our own typed errors (e.g. NoCurrentWindowError) pass through
            except Exception as exc:  # noqa: BLE001  (pywintypes.com_error is dynamic)
                hr = _hresult_of(exc)
                if is_retryable_hresult(hr):
                    last = exc
                    time.sleep(_BASE_DELAY_S * (2**attempt))
                    continue
                raise OneNoteComError(f"{name} failed", hresult=hr) from exc
        raise OneNoteComError(
            f"{name} kept returning busy after retries", hresult=_hresult_of(last)
        )

    @staticmethod
    def _com_date(value: _dt.datetime | None):
        """Convert a datetime to a COM DATE; ``None`` → 0 (skip the concurrency check)."""
        if value is None:
            return 0
        import pythoncom  # noqa: PLC0415
        import pywintypes  # noqa: PLC0415

        _ = pythoncom  # imported to ensure COM types registered
        return pywintypes.Time(value)

    _SCHEMA = int(XMLSchema.xs2013)

    # --- Notebook structure -------------------------------------------------

    def get_hierarchy(self, start_node_id: str, scope: HierarchyScope) -> str:
        return self._call(
            "GetHierarchy",
            lambda: self.app.GetHierarchy(start_node_id, int(scope), self._SCHEMA),
        )

    def update_hierarchy(self, changes_xml: str) -> None:
        self._call("UpdateHierarchy", lambda: self.app.UpdateHierarchy(changes_xml, self._SCHEMA))

    def open_hierarchy(
        self,
        path: str,
        relative_to_object_id: str,
        create_file_type: CreateFileType = CreateFileType.cftNone,
    ) -> str:
        return self._call(
            "OpenHierarchy",
            lambda: self.app.OpenHierarchy(path, relative_to_object_id, int(create_file_type)),
        )

    def create_new_page(
        self, section_id: str, style: NewPageStyle = NewPageStyle.npsDefault
    ) -> str:
        return self._call("CreateNewPage", lambda: self.app.CreateNewPage(section_id, int(style)))

    def delete_hierarchy(
        self,
        object_id: str,
        expected_last_modified: _dt.datetime | None = None,
        permanent: bool = False,
    ) -> None:
        self._call(
            "DeleteHierarchy",
            lambda: self.app.DeleteHierarchy(
                object_id, self._com_date(expected_last_modified), permanent
            ),
        )

    def get_hierarchy_parent(self, object_id: str) -> str:
        return self._call("GetHierarchyParent", lambda: self.app.GetHierarchyParent(object_id))

    def get_special_location(self, location: SpecialLocation) -> str:
        return self._call("GetSpecialLocation", lambda: self.app.GetSpecialLocation(int(location)))

    # --- Page content -------------------------------------------------------

    def get_page_content(self, page_id: str, page_info: PageInfo = PageInfo.piBasic) -> str:
        return self._call(
            "GetPageContent",
            lambda: self.app.GetPageContent(page_id, int(page_info), self._SCHEMA),
        )

    def update_page_content(
        self,
        changes_xml: str,
        expected_last_modified: _dt.datetime | None = None,
        force: bool = False,
    ) -> None:
        self._call(
            "UpdatePageContent",
            lambda: self.app.UpdatePageContent(
                changes_xml, self._com_date(expected_last_modified), self._SCHEMA, force
            ),
        )

    def get_binary_page_content(self, page_id: str, callback_id: str) -> str:
        return self._call(
            "GetBinaryPageContent",
            lambda: self.app.GetBinaryPageContent(page_id, callback_id),
        )

    def delete_page_content(
        self,
        page_id: str,
        object_id: str,
        expected_last_modified: _dt.datetime | None = None,
        force: bool = False,
    ) -> None:
        self._call(
            "DeletePageContent",
            lambda: self.app.DeletePageContent(
                page_id, object_id, self._com_date(expected_last_modified), force
            ),
        )

    # --- Navigation ---------------------------------------------------------

    def find_pages(self, start_node_id: str, query: str, include_unindexed: bool = False) -> str:
        return self._call(
            "FindPages",
            lambda: self.app.FindPages(
                start_node_id, query, include_unindexed, False, self._SCHEMA
            ),
        )

    def get_hyperlink_to_object(self, hierarchy_id: str, object_id: str = "") -> str:
        return self._call(
            "GetHyperlinkToObject",
            lambda: self.app.GetHyperlinkToObject(hierarchy_id, object_id),
        )

    def get_current_window_ids(self) -> CurrentWindowIds:
        # ⚠ PHASE 3 (VM): property (not method) marshalling — Windows/CurrentWindow access
        # under early binding must be confirmed on the VM alongside the [out]-param question.
        def read() -> CurrentWindowIds:
            windows = self.app.Windows
            current = windows.CurrentWindow if windows.Count > 0 else None
            if current is None:
                raise NoCurrentWindowError(
                    "OneNote has no open window; cannot read the current viewing context."
                )
            return CurrentWindowIds(
                notebook_id=current.CurrentNotebookId or None,
                section_group_id=current.CurrentSectionGroupId or None,
                section_id=current.CurrentSectionId or None,
                page_id=current.CurrentPageId or None,
            )

        return self._call("Windows.CurrentWindow", read)
