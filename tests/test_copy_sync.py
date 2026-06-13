"""Tier-1: under-synced content during copy is COUNTED and surfaced (SPEC §5), never silent.

VM-confirmed (2026-06-13) root cause of a copy losing images/files: the source isn't fully
downloaded on this machine (OneDrive files-on-demand), so GetBinaryPageContent → 0x8004200F and
the attachment cache file is absent. The copy must report how many images / files / embedded
objects it could not carry, categorized, so the user can fully sync and re-copy.
"""

from __future__ import annotations

from lxml import etree

from onenote_com_mcp.errors import OneNoteComError
from onenote_com_mcp.service import copy
from onenote_com_mcp.service.page_edit import inline_image_binaries

_ONE = "http://schemas.microsoft.com/office/onenote/2013/onenote"


class _UnsyncedBackend:
    """Simulates a machine where the source section is NOT fully downloaded."""

    def get_binary_page_content(self, page_id: str, callback_id: str) -> str:
        raise OneNoteComError("binary not local", hresult=0x8004200F)

    def stage_cache_copy(self, path: str, preferred_name: str) -> str | None:
        return None  # cache file not downloaded → unavailable


def test_inline_image_binaries_counts_unfetchable_and_keeps_payload_valid():
    tree = etree.fromstring(
        f'<one:Page xmlns:one="{_ONE}">'
        f'<one:Image><one:CallbackID callbackID="{{CB1}}"/></one:Image>'
        f'<one:Image><one:CallbackID callbackID="{{CB2}}"/></one:Image>'
        "</one:Page>".encode()
    )

    dropped = inline_image_binaries(_UnsyncedBackend(), "{PID}", tree)

    assert dropped == 2
    # both images still carry a (placeholder) one:Data so UpdatePageContent stays valid XML,
    # and the read-side CallbackID is gone
    assert len(tree.findall(f".//{{{_ONE}}}Data")) == 2
    assert tree.find(f".//{{{_ONE}}}CallbackID") is None


def test_rewrite_inserted_files_categorizes_file_vs_embedded_object():
    tree = etree.fromstring(
        f'<one:Page xmlns:one="{_ONE}">'
        f'<one:InsertedFile preferredName="report.pdf" pathCache="C:\\Temp\\a.bin" '
        f'pathSource="C:\\orig\\report.pdf"/>'
        f'<one:InsertedFile preferredName="sheet.xlsx" pathCache="C:\\Temp\\b.bin">'
        f"<one:Previews/></one:InsertedFile>"
        "</one:Page>".encode()
    )

    missing_files, missing_objects, notes = copy._rewrite_inserted_files(_UnsyncedBackend(), tree)

    assert missing_files == 1  # the .pdf attachment icon
    assert missing_objects == 1  # the embedded .xlsx (has one:Previews)
    assert len(notes) == 2
    assert any("report.pdf" in n and "file" in n for n in notes)
    assert any("sheet.xlsx" in n and "embedded object" in n for n in notes)


def test_sync_warning_none_when_nothing_missing():
    assert copy.sync_warning(0, 0, 0) is None


def test_sync_warning_lists_each_category_and_warns_to_sync():
    warning = copy.sync_warning(3, 1, 2)

    assert "3 image(s)" in warning
    assert "1 file(s)" in warning
    assert "2 embedded object(s)" in warning
    assert "NOT fully synced" in warning
    assert "will NOT auto-download" in warning
