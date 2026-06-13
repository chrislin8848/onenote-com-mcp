"""Tier-1: under-synced content during copy is REMOVED (not left as a dead marker) and reported.

VM-confirmed (2026-06-13): a copy losing images/files is caused by the source not being fully
downloaded on this machine (OneDrive files-on-demand) — GetBinaryPageContent → 0x8004200F and the
attachment cache is absent. A copied placeholder/broken reference can't self-heal and could later
be misread or re-copied, so the COPY path removes the un-copyable element (pruning the emptied
OE/Outline) and reports a categorized sync_warning. The EDIT path keeps a placeholder instead — it
must not delete an image that lives in the cloud but isn't downloaded locally.
"""

from __future__ import annotations

from lxml import etree

from onenote_com_mcp.errors import OneNoteComError
from onenote_com_mcp.service import copy
from onenote_com_mcp.service.page_edit import inline_image_binaries

_ONE = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _tag(name: str) -> str:
    return f"{{{_ONE}}}{name}"


class _UnsyncedBackend:
    """Simulates a machine where the source section is NOT fully downloaded."""

    def get_binary_page_content(self, page_id: str, callback_id: str) -> str:
        raise OneNoteComError("binary not local", hresult=0x8004200F)

    def stage_cache_copy(self, path: str, preferred_name: str) -> str | None:
        return None  # cache file not downloaded → unavailable


def test_inline_image_binaries_default_keeps_placeholder_for_edit_path():
    tree = etree.fromstring(
        f'<one:Page xmlns:one="{_ONE}">'
        f'<one:Image><one:CallbackID callbackID="{{CB1}}"/></one:Image>'
        f'<one:Image><one:CallbackID callbackID="{{CB2}}"/></one:Image>'
        "</one:Page>".encode()
    )

    dropped = inline_image_binaries(_UnsyncedBackend(), "{PID}", tree)  # default: edit path

    assert dropped == 2
    # default keeps a valid (blank) box so an edit never deletes a cloud-only image
    assert len(tree.findall(f".//{_tag('Data')}")) == 2
    assert tree.find(f".//{_tag('CallbackID')}") is None


def test_inline_image_binaries_removes_and_prunes_in_copy_mode():
    tree = etree.fromstring(
        f'<one:Page xmlns:one="{_ONE}"><one:Outline><one:OEChildren>'
        f'<one:OE><one:Image><one:CallbackID callbackID="{{CB}}"/></one:Image></one:OE>'
        f"<one:OE><one:T>keep me</one:T></one:OE>"
        "</one:OEChildren></one:Outline></one:Page>".encode()
    )

    dropped = inline_image_binaries(_UnsyncedBackend(), "{PID}", tree, remove_unfetchable=True)

    assert dropped == 1
    assert tree.find(f".//{_tag('Image')}") is None  # the un-fetchable image is gone
    # its now-empty OE was pruned; the sibling text OE (real content) is untouched
    assert len(tree.findall(f".//{_tag('OE')}")) == 1
    assert tree.find(f".//{_tag('T')}").text == "keep me"


def test_rewrite_inserted_files_removes_and_categorizes():
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
    assert tree.find(f".//{_tag('InsertedFile')}") is None  # both removed, no dead references left
    assert len(notes) == 2
    assert any("report.pdf" in n and "file" in n and "removed from the copy" in n for n in notes)
    assert any("sheet.xlsx" in n and "embedded object" in n for n in notes)


def test_sync_warning_none_when_nothing_missing():
    assert copy.sync_warning(0, 0, 0) is None


def test_sync_warning_lists_each_category_and_warns_to_sync():
    warning = copy.sync_warning(3, 1, 2)

    assert "3 image(s)" in warning
    assert "1 file(s)" in warning
    assert "2 embedded object(s)" in warning
    assert "NOT fully synced" in warning
    assert "OMITTED" in warning
    assert "self-heal" in warning
