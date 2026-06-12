"""``FixtureBackend`` — replays recorded OneNote XML from disk (Linux, no COM).

Read methods serve XML from a fixtures directory using a filename convention (below).
Write methods record their payloads in ``self.calls`` and return deterministic fake IDs, so
service-layer logic can be exercised end-to-end on Linux without a live OneNote.

Fixtures are produced by ``scripts/dump_fixtures.py`` on the Windows VM (Phase 3); until then
the directory may be empty and read methods raise a clear "missing fixture" error.

Filename convention (under ``fixtures_dir``):
  hierarchy_<scope>.xml                 e.g. hierarchy_hsNotebooks.xml
  hierarchy_<scope>__<sanitized id>.xml scoped variant (preferred if present)
  page_<sanitized id>.xml               GetPageContent (piBasic)
  page_<sanitized id>__binary.xml       GetPageContent (piBinaryData)
  binary_<sanitized callback>.b64       GetBinaryPageContent
  cachefile_<sanitized basename>.bin    InsertedFile pathCache bytes (raw), keyed by the
                                        Windows path's basename ({GUID}.bin — unique); a
                                        missing file replays "cache unavailable" (None)
  find__<sanitized query>.xml           FindPages
  current_window.json                   Windows.CurrentWindow Current*Id quadruple (JSON
                                        object; literal "null" = no open window)
"""

from __future__ import annotations

import datetime as _dt
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from onenote_com_mcp.backend.base import CurrentWindowIds, OneNoteBackend
from onenote_com_mcp.enums import (
    CreateFileType,
    HierarchyScope,
    NewPageStyle,
    PageInfo,
    SpecialLocation,
)
from onenote_com_mcp.errors import NoCurrentWindowError, NodeNotFoundError


def _sanitize(token: str) -> str:
    """Make a OneNote ID/query safe for use in a filename."""
    return re.sub(r"[^A-Za-z0-9]+", "_", token).strip("_") or "root"


@dataclass
class RecordedCall:
    method: str
    kwargs: dict[str, Any]


class FixtureBackend(OneNoteBackend):
    def __init__(self, fixtures_dir: str | Path) -> None:
        self.fixtures_dir = Path(fixtures_dir)
        self.calls: list[RecordedCall] = []
        self._fake_id_seq = 0

    # --- helpers ------------------------------------------------------------

    def _read(self, *candidates: str) -> str:
        for name in candidates:
            path = self.fixtures_dir / name
            if path.exists():
                return path.read_text(encoding="utf-8")
        raise NodeNotFoundError(
            f"No fixture found. Tried {list(candidates)} under {self.fixtures_dir}. "
            "Dump real fixtures with scripts/dump_fixtures.py on the VM (Phase 3)."
        )

    def _record(self, method: str, **kwargs: Any) -> None:
        self.calls.append(RecordedCall(method, kwargs))

    def _next_fake_id(self, kind: str) -> str:
        self._fake_id_seq += 1
        return f"{{FIXTURE-{kind}-{self._fake_id_seq}}}{{1}}{{B0}}"

    # --- reads --------------------------------------------------------------

    def get_hierarchy(self, start_node_id: str, scope: HierarchyScope) -> str:
        scoped = f"hierarchy_{scope.name}__{_sanitize(start_node_id)}.xml"
        return self._read(scoped, f"hierarchy_{scope.name}.xml")

    def get_page_content(self, page_id: str, page_info: PageInfo = PageInfo.piBasic) -> str:
        sid = _sanitize(page_id)
        if page_info in (PageInfo.piBinaryData, PageInfo.piBinaryDataSelection, PageInfo.piAll):
            return self._read(f"page_{sid}__binary.xml", f"page_{sid}.xml")
        return self._read(f"page_{sid}.xml")

    def get_binary_page_content(self, page_id: str, callback_id: str) -> str:
        return self._read(f"binary_{_sanitize(callback_id)}.b64")

    def find_pages(self, start_node_id: str, query: str, include_unindexed: bool = False) -> str:
        return self._read(f"find__{_sanitize(query)}.xml")

    def _cache_fixture(self, path: str) -> Path:
        # key by the Windows path's basename — pathCache is always ...\Temp\{GUID}.bin
        # (ground truth), so the GUID basename is unique and host-OS-agnostic
        basename = re.split(r"[\\/]", path)[-1]
        return self.fixtures_dir / f"cachefile_{_sanitize(basename)}.bin"

    def stat_cache_file(self, path: str) -> int | None:
        fixture = self._cache_fixture(path)
        return fixture.stat().st_size if fixture.exists() else None

    def read_cache_file(self, path: str) -> bytes | None:
        fixture = self._cache_fixture(path)
        return fixture.read_bytes() if fixture.exists() else None

    def get_hierarchy_parent(self, object_id: str) -> str:
        return self._read(f"parent_{_sanitize(object_id)}.txt").strip()

    def get_special_location(self, location: SpecialLocation) -> str:
        return self._read(f"special_{location.name}.txt").strip()

    def get_hyperlink_to_object(self, hierarchy_id: str, object_id: str = "") -> str:
        return self._read(f"hyperlink_{_sanitize(hierarchy_id)}_{_sanitize(object_id)}.txt").strip()

    def get_current_window_ids(self) -> CurrentWindowIds:
        data = json.loads(self._read("current_window.json"))
        if data is None:  # file contains literal "null" → replay the no-window case
            raise NoCurrentWindowError("OneNote has no open window (fixture replay)")
        return CurrentWindowIds(
            notebook_id=data.get("notebook_id") or None,
            section_group_id=data.get("section_group_id") or None,
            section_id=data.get("section_id") or None,
            page_id=data.get("page_id") or None,
        )

    # --- writes (recorded, not executed) ------------------------------------

    def update_hierarchy(self, changes_xml: str) -> None:
        self._record("update_hierarchy", changes_xml=changes_xml)

    def open_hierarchy(
        self,
        path: str,
        relative_to_object_id: str,
        create_file_type: CreateFileType = CreateFileType.cftNone,
    ) -> str:
        self._record(
            "open_hierarchy",
            path=path,
            relative_to_object_id=relative_to_object_id,
            create_file_type=create_file_type,
        )
        return self._next_fake_id(create_file_type.name)

    def create_new_page(
        self, section_id: str, style: NewPageStyle = NewPageStyle.npsDefault
    ) -> str:
        self._record("create_new_page", section_id=section_id, style=style)
        return self._next_fake_id("page")

    def update_page_content(
        self,
        changes_xml: str,
        expected_last_modified: _dt.datetime | None = None,
        force: bool = False,
    ) -> None:
        self._record(
            "update_page_content",
            changes_xml=changes_xml,
            expected_last_modified=expected_last_modified,
            force=force,
        )

    def delete_hierarchy(
        self,
        object_id: str,
        expected_last_modified: _dt.datetime | None = None,
        permanent: bool = False,
    ) -> None:
        self._record(
            "delete_hierarchy",
            object_id=object_id,
            expected_last_modified=expected_last_modified,
            permanent=permanent,
        )

    def delete_page_content(
        self,
        page_id: str,
        object_id: str,
        expected_last_modified: _dt.datetime | None = None,
        force: bool = False,
    ) -> None:
        self._record(
            "delete_page_content",
            page_id=page_id,
            object_id=object_id,
            expected_last_modified=expected_last_modified,
            force=force,
        )
