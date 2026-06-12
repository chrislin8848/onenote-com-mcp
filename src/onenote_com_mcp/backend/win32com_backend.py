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
import re
import time
from pathlib import Path

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
    ConcurrencyError,
    NoCurrentWindowError,
    OneNoteComError,
    OneNoteError,
    is_concurrency_hresult,
    is_retryable_hresult,
)

_MAX_RETRIES = 6
_BASE_DELAY_S = 0.25

_STAMP_RE = re.compile(r'lastModifiedTime="([^"]+)"')
_PAGE_ID_RE = re.compile(r'\bID="([^"]+)"')


def _hresult_of(exc: Exception) -> int | None:
    """Pull the meaningful HRESULT out of a pywintypes.com_error, if present.

    ``com_error.args`` is ``(hresult, message, excepinfo, argerror)``. When OneNote raises a
    server-side exception, the outer hresult is just DISP_E_EXCEPTION — the actual OneNote
    error code (hrLastModifiedDateDidNotMatch, hrInvalidXML, …) is the excepinfo's ``scode``
    (its last element). VM ground truth 2026-06-11.
    """
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and args[2]:
        scode = args[2][-1]
        if isinstance(scode, int) and scode != 0:
            return scode
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
                if is_concurrency_hresult(hr):
                    raise ConcurrencyError(
                        f"{name} refused: the target changed since it was read "
                        "(dateExpectedLastModified mismatch). Re-read and retry; force only "
                        "with explicit user approval."
                    ) from exc
                raise OneNoteComError(f"{name} failed", hresult=hr) from exc
        raise OneNoteComError(
            f"{name} kept returning busy after retries", hresult=_hresult_of(last)
        )

    @staticmethod
    def _date_kwargs(value: _dt.datetime | None) -> dict:
        """``dateExpectedLastModified`` kwargs for a makepy call.

        VM ground truth (2026-06-11): the early-bound VT_DATE param accepts ONLY a PyTime —
        a plain int, the makepy default tuple, and even omitting the parameter all fail with
        ``must be a pywintypes time object``; AND pythoncom cannot marshal pre-1970 stamps
        (mktime → OSError), so the documented "DATE 0 == skip the check" is UNREACHABLE from
        Python. Callers therefore must always supply a real stamp — the write/delete methods
        below resolve the node's CURRENT stamp when given None (which is also what SPEC §5
        wants: every write/delete carries the lastModifiedTime from a read). VT_DATE is
        timezone-less local time and ``pywintypes.Time`` does NOT localize aware datetimes,
        so OneNote's UTC ``...Z`` stamps are converted to local naive before the comparison.
        """
        import pywintypes  # noqa: PLC0415

        if value is None:
            raise OneNoteComError(
                "dateExpectedLastModified is required (COM cannot express 'skip the check') "
                "— resolve the node's current lastModifiedTime first"
            )
        if value.tzinfo is not None:
            value = value.astimezone().replace(tzinfo=None)
        return {"dateExpectedLastModified": pywintypes.Time(value)}

    @staticmethod
    def _xml_stamp(xml: str) -> _dt.datetime | None:
        """First ``lastModifiedTime`` attribute in an XML string → aware datetime.

        Deliberately a regex, not a parse-layer call: the backend stays a thin COM mirror
        (SPEC §3) and only needs this one attribute to satisfy VT_DATE marshalling."""
        match = _STAMP_RE.search(xml)
        if not match:
            return None
        try:
            return _dt.datetime.fromisoformat(match.group(1).replace("Z", "+00:00"))
        except ValueError:
            return None

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
        # cftNotebook is not used: this M365 build refuses COM notebook creation outright
        # (hrFileDoesNotExist for local paths AND OneDrive https parents — VM 2026-06-11).
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
        if expected_last_modified is None:
            # SPEC §5: deletes carry the node's stamp. Resolve the CURRENT one — "skip the
            # check" is not marshallable (see _date_kwargs).
            expected_last_modified = self._xml_stamp(
                self.get_hierarchy(object_id, HierarchyScope.hsSelf)
            )
        self._call(
            "DeleteHierarchy",
            lambda: self.app.DeleteHierarchy(
                object_id,
                deletePermanently=permanent,
                **self._date_kwargs(expected_last_modified),
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
        if force:
            # VM ground truth (2026-06-11): COM's force flag does NOT bypass the date check —
            # it only overrides unsaved-UI-edit protection. Our contract's force means
            # "overwrite even though it changed", so resolve and carry the CURRENT stamp.
            expected_last_modified = None
        elif expected_last_modified is None:
            # the payload normally carries the stamp of the read it was built from
            expected_last_modified = self._xml_stamp(changes_xml)
        if expected_last_modified is None and (match := _PAGE_ID_RE.search(changes_xml)):
            expected_last_modified = self._xml_stamp(
                self.get_page_content(match.group(1), PageInfo.piBasic)
            )
        self._call(
            "UpdatePageContent",
            lambda: self.app.UpdatePageContent(
                changes_xml,
                xsSchema=self._SCHEMA,
                force=force,
                **self._date_kwargs(expected_last_modified),
            ),
        )

    def get_binary_page_content(self, page_id: str, callback_id: str) -> str:
        return self._call(
            "GetBinaryPageContent",
            lambda: self.app.GetBinaryPageContent(page_id, callback_id),
        )

    # --- Attachment cache files (Phase 5b) — plain disk IO, no COM ----------
    # pathCache points at this machine's %LOCALAPPDATA%\Temp (the server runs next to
    # OneNote). Missing/purged cache → None, never an exception (SPEC §5: graceful
    # "cache unavailable").

    def stat_cache_file(self, path: str) -> int | None:
        try:
            return Path(path).stat().st_size
        except OSError:
            return None

    def read_cache_file(self, path: str) -> bytes | None:
        try:
            return Path(path).read_bytes()
        except OSError:
            return None

    def stage_cache_copy(self, path: str, preferred_name: str) -> str | None:
        # Staged copies live under %TEMP%\OneNoteMCP\staging\<uuid>\<name> and are NOT
        # auto-deleted: OneNote re-imports pathSource asynchronously and there is no signal
        # for "import finished" (Stage-3 VM question). The uuid dir keeps the real filename
        # (extension drives re-import) while avoiding collisions.
        import shutil  # noqa: PLC0415
        import tempfile  # noqa: PLC0415
        import uuid  # noqa: PLC0415

        src = Path(path)
        if not src.is_file():
            return None
        safe_name = re.sub(r'[\\/:*?"<>|]', "_", preferred_name) or "attachment"
        dest = Path(tempfile.gettempdir()) / "OneNoteMCP" / "staging" / uuid.uuid4().hex[:12]
        try:
            dest.mkdir(parents=True, exist_ok=True)
            target = dest / safe_name
            shutil.copyfile(src, target)
            return str(target)
        except OSError:
            return None

    def delete_page_content(
        self,
        page_id: str,
        object_id: str,
        expected_last_modified: _dt.datetime | None = None,
        force: bool = False,
    ) -> None:
        if expected_last_modified is None:
            expected_last_modified = self._xml_stamp(
                self.get_page_content(page_id, PageInfo.piBasic)
            )
        self._call(
            "DeletePageContent",
            lambda: self.app.DeletePageContent(
                page_id,
                object_id,
                force=force,
                **self._date_kwargs(expected_last_modified),
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
