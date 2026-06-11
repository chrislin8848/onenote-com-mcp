"""``Win32ComBackend`` — live OneNote via in-process pywin32 COM (Windows only).

GUARDED IMPORT INVARIANT (SPEC §2.1): this module must ``import`` cleanly on Linux. Therefore
**no** ``win32com`` / ``pywintypes`` / ``pythoncom`` import appears at module top level — they
are imported lazily inside methods. ``tests/test_smoke_import.py`` enforces this.

CONNECTION (confirmed on the VM, 2026-06-11): OneNote must be early-bound via an explicit
``gencache.EnsureModule`` + coclass instantiation (see ``app`` below) — ``EnsureDispatch`` and
plain ``Dispatch`` both fail. Under that binding ``[out] BSTR`` params are returned as the
call's return value, which is the string in/out contract the rest of the codebase depends on.
``get_hierarchy`` round-trips real notebook XML over both the SPICE console and SSH (OneNote is
an out-of-process COM server in the autologon session, so DCOM activation reaches it either way).
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
        """Lazily connect to the running OneNote application via COM.

        OneNote cannot be late-bound: ``Dispatch("OneNote.Application")`` yields a dynamic
        object whose ``GetIDsOfNames`` can't resolve ``GetHierarchy``, and
        ``gencache.EnsureDispatch`` fails with "can not automate the makepy process". The
        working recipe (confirmed on the VM, 2026-06-11) is to locate the OneNote type library,
        build its early-bound makepy module explicitly, then instantiate the coclass — after
        which ``[out] BSTR`` params come back as the call's return value. For a frozen build the
        gen cache must be bundled (SPEC §8); see the PyInstaller spec when packaging.
        """
        if self._app is None:
            try:
                # Imported here, never at module top, to keep Linux import clean.
                from win32com.client import gencache, selecttlb  # noqa: PLC0415

                # Pick the OneNote 15.x (Office 2013+/M365 desktop) type library. A stale
                # "OneNote 12.0" / version 1.0 registration with no backing file makes COM
                # calls fail later with TYPE_E_LIBNOTREGISTERED (-2147319779) — prefer the
                # highest version, which is the live one.
                tlbs = [t for t in selecttlb.EnumTlbs() if "onenote" in t.desc.lower()]
                if not tlbs:
                    raise RuntimeError("No OneNote type library registered.")
                tlb = max(tlbs, key=lambda t: (int(t.major, 16), int(t.minor, 16)))
                mod = gencache.EnsureModule(tlb.clsid, 0, int(tlb.major, 16), int(tlb.minor, 16))
                self._app = mod.Application()
            except Exception as exc:  # noqa: BLE001
                raise BackendUnavailableError(
                    "Could not connect to OneNote via COM. Ensure the OneNote desktop app "
                    "(M365, not the UWP 'OneNote for Windows 10') is installed and running in "
                    "this interactive session. If a call fails with 'program library not "
                    "registered', delete the stale HKCR\\TypeLib\\{0EA692EE-...}\\1.0 subkey."
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
        # CONFIRMED on the VM (2026-06-11, Tier 2): under the early-bound makepy module the
        # Windows / CurrentWindow PROPERTY chain and the Current*Id properties marshal fine —
        # no special handling needed beyond what methods get.
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
