"""Tier 2 — live-COM attachment/embedded-object validation (SPEC v0612-2 §5, Phase 5b
Stage 3). Windows only.

Validates against the REAL attachment pages (附件與嵌入物件-1/-2, the dump sources):
the two read tools round-trip live, the copy seam's rewritten payload (staged pathSource,
no pathCache, no XPSFile) is ACCEPTED by UpdatePageContent, editing a printout-bearing page
through changed_objects works, and the Phase-6 delete questions get probed.

PROBES vs TESTS: re-import timing, embedded clone behavior, and inline-attachment-OE deletion
are open VM questions — those report their observed outcome via ``pytest.skip(...)`` (visible
in the -ra summary remote_test.sh collects) instead of failing the gate.

Assertions stay content-light on purpose (status/lengths/names, never file text): tier2.log
travels back to the host, and live attachment content must not leak into it (PII policy).

Self-cleaning: throwaway sections carry the "P5B暫存" prefix, swept at module setup and
deleted in teardown.
"""

from __future__ import annotations

import base64
import time
import uuid

import pytest

from onenote_com_mcp.errors import OneNoteError
from onenote_com_mcp.service import copy, files, page_edit, read

pytestmark = pytest.mark.windows

TEST_NOTEBOOK = "MCP Test"
TEST_SECTION = "Phase 0 測試用"
PAGE_1 = "附件與嵌入物件-1"
PAGE_2 = "附件與嵌入物件-2"
TEMP_PREFIX = "P5B暫存"

# ground truth from the dumps — what the live pages must keep reporting
PAGE_1_KINDS = {
    "濁水溪發電之旅.txt": ("inline", "attachment_icon"),
    "丘山行問卷_中英對照.pdf": ("inline", "attachment_icon"),
    "A4文宣-25.7.8月分享會.pdf": ("inline", "printout"),
    "丘山行Word頁籤(中文) .docx": ("page_level", "attachment_icon"),
    "2026 客人問卷_NEW.xlsx": ("page_level", "attachment_icon"),
}


def _find(nodes, name):
    return next((n for n in nodes if n.get("name") == name), None)


@pytest.fixture(scope="module")
def backend():
    from onenote_com_mcp.backend.win32com_backend import Win32ComBackend

    return Win32ComBackend()


@pytest.fixture(scope="module")
def notebook_id(backend):
    nb = _find(read.list_notebooks(backend), TEST_NOTEBOOK)
    if nb is None:
        pytest.skip(f"{TEST_NOTEBOOK!r} notebook not present on this VM")
    for node in read.list_sections(backend, nb["id"]):
        if node["name"].startswith(TEMP_PREFIX):
            backend.delete_hierarchy(node["id"], permanent=True)
    return nb["id"]


@pytest.fixture(scope="module")
def attachment_pages(backend, notebook_id):
    sec = _find(read.list_sections(backend, notebook_id), TEST_SECTION)
    if sec is None:
        pytest.skip(f"{TEST_SECTION!r} section not found")
    pages = {p["name"]: p["id"] for p in read.list_pages(backend, sec["id"])}
    missing = [n for n in (PAGE_1, PAGE_2) if n not in pages]
    if missing:
        pytest.skip(f"attachment fixture pages missing on the VM: {missing}")
    return pages


@pytest.fixture
def temp_section(backend, notebook_id):
    from onenote_com_mcp.service import create

    section_id = create.create_section(
        backend, notebook_id, f"{TEMP_PREFIX}-{uuid.uuid4().hex[:8]}"
    )
    yield section_id
    backend.delete_hierarchy(section_id, permanent=True)


def _png_dimensions(raw: bytes) -> tuple[int, int]:
    assert raw[:8] == b"\x89PNG\r\n\x1a\n", "expected a PNG"
    return int.from_bytes(raw[16:20], "big"), int.from_bytes(raw[20:24], "big")


def _poll(fn, timeout_s: float = 30.0, interval_s: float = 2.0):
    """Poll fn() until it returns truthy or the timeout elapses; returns the last value."""
    deadline = time.monotonic() + timeout_s
    value = fn()
    while not value and time.monotonic() < deadline:
        time.sleep(interval_s)
        value = fn()
    return value


# --- read tools, live --------------------------------------------------------------------


def test_files_info_live_inventory(backend, attachment_pages):
    info = files.get_page_files_info(backend, attachment_pages[PAGE_1])
    assert {e["preferred_name"]: (e["placement"], e["kind"]) for e in info} == PAGE_1_KINDS
    assert all(e["object_id"] for e in info), "every entry must carry a deletable object_id"
    # locally-inserted files: their caches must be on this machine's disk
    unavailable = [e["preferred_name"] for e in info if not e["cache_available"]]
    assert not unavailable, f"live caches unexpectedly missing: {unavailable}"
    assert all(e["size_bytes"] > 0 for e in info)


def test_files_info_live_embedded(backend, attachment_pages):
    info = files.get_page_files_info(backend, attachment_pages[PAGE_2])
    embedded = next(e for e in info if e["kind"] == "embedded_preview")
    assert embedded["preview_pages"] == ["工作表1"]
    assert embedded["path_source"] is None, "embedded objects carry no pathSource (ground truth)"
    assert embedded["cache_available"] is True


def test_files_content_live(backend, attachment_pages):
    by_name = {
        e["preferred_name"]: e for e in files.get_page_files(backend, attachment_pages[PAGE_1])
    }
    txt = by_name["濁水溪發電之旅.txt"]
    assert txt["status"] == "ok" and len(txt["text"]) > 0
    pdf = by_name["丘山行問卷_中英對照.pdf"]
    assert pdf["status"] == "ok" and pdf["pdf_pages"] >= 1
    assert by_name["丘山行Word頁籤(中文) .docx"]["status"] == "unsupported"

    jpg = next(
        e
        for e in files.get_page_files(backend, attachment_pages[PAGE_2])
        if e["preferred_name"] == "捷斯山屋.jpg"
    )
    assert jpg["status"] == "ok" and jpg["media_type"] == "image/jpeg"
    assert base64.b64decode(jpg["data_base64"]).startswith(b"\xff\xd8\xff"), "real JPEG bytes"


def test_get_page_images_returns_printout_render_live(backend, attachment_pages):
    images = read.get_page_images(backend, attachment_pages[PAGE_1])
    renders = [i for i in images if i["is_printout"]]
    assert len(renders) == 1, "the page-level printout render must be visible (Phase 5b fix)"
    assert renders[0]["object_id"]
    assert len(base64.b64decode(renders[0]["data_base64"])) > 0


# --- copy seam: the rewritten payload must be ACCEPTED live ------------------------------


def test_copy_attachment_page_live(backend, attachment_pages, temp_section):
    result = copy.transfer_page(backend, attachment_pages[PAGE_1], temp_section)

    # printout flattening is expected normalization, NOT reported (user decision 2026-06-12);
    # all caches are live here, so the copy must be note-free.
    assert result.file_notes == [], result.file_notes

    copy_info = files.get_page_files_info(backend, result.page_id)
    assert sorted(e["preferred_name"] for e in copy_info) == sorted(PAGE_1_KINDS), (
        "every attachment entry must survive the clone"
    )
    # the flattened render survives as a plain image. GROUND TRUTH (first Stage-3 run,
    # 2026-06-12): OneNote RE-ENCODES the large render PNG on transplant — byte-identity does
    # NOT hold for printout renders (unlike normal image inserts, which round-trip
    # byte-identical). The fidelity claim is therefore semantic: same pixel dimensions.
    src_render = next(
        i for i in read.get_page_images(backend, attachment_pages[PAGE_1]) if i["is_printout"]
    )
    copy_images = read.get_page_images(backend, result.page_id)
    assert len(copy_images) == 1, "exactly the render — no other image on this page"
    assert _png_dimensions(base64.b64decode(copy_images[0]["data_base64"])) == _png_dimensions(
        base64.b64decode(src_render["data_base64"])
    ), "render pixel dimensions must survive the clone"
    assert all(not i["is_printout"] for i in copy_images), "flattened render is a plain image"


def test_copy_embedded_page_live(backend, attachment_pages, temp_section):
    result = copy.transfer_page(backend, attachment_pages[PAGE_2], temp_section)
    assert result.file_notes == [], "both caches live — the copy must be note-free"
    copy_info = files.get_page_files_info(backend, result.page_id)
    assert sorted(e["preferred_name"] for e in copy_info) == [
        "捷斯山屋.jpg",
        "附件與嵌入物件-2 - 工作表.xlsx",
    ]


def test_copy_reimport_probe(backend, attachment_pages, temp_section):
    """OPEN VM QUESTION (SPEC §5): does OneNote re-import staged pathSource files — rebuild
    its own cache — and what happens to a cloned embedded object? Reports, never fails."""
    page1 = copy.transfer_page(backend, attachment_pages[PAGE_1], temp_section)
    page2 = copy.transfer_page(backend, attachment_pages[PAGE_2], temp_section)

    def caches_ready():
        info = files.get_page_files_info(backend, page1.page_id)
        return info if all(e["cache_available"] for e in info) else None

    ready = _poll(caches_ready, timeout_s=45.0)
    info1 = ready or files.get_page_files_info(backend, page1.page_id)
    cache_state = {e["preferred_name"]: e["cache_available"] for e in info1}

    info2 = files.get_page_files_info(backend, page2.page_id)
    embedded = next(
        (e for e in info2 if e["preferred_name"] == "附件與嵌入物件-2 - 工作表.xlsx"), None
    )
    pytest.skip(
        "RE-IMPORT PROBE — "
        f"copy caches rebuilt within 45s: {cache_state}; "
        f"cloned embedded object reports kind={embedded['kind'] if embedded else 'MISSING'}, "
        f"preview_pages={embedded['preview_pages'] if embedded else '-'}"
    )


# --- edit seam: changed_objects on a printout-bearing page -------------------------------


def test_edit_printout_page_changed_objects_live(backend, attachment_pages):
    """The pruned payload (no InsertedFile/XPSFile/render — Phase 5b _CONTENT_TAGS) must be
    accepted by UpdatePageContent and must not disturb the attachments. Self-limiting wear:
    first run appends a marker paragraph, later runs replace it in place."""
    page_id = attachment_pages[PAGE_1]
    marker_prefix = "P5B-edit-marker"
    new_text = f"{marker_prefix} {uuid.uuid4().hex[:8]}"
    before = files.get_page_files_info(backend, page_id)

    existing = next(
        (
            b
            for o in read.get_page(backend, page_id)["outlines"]
            for b in o["blocks"]
            if b["type"] == "paragraph" and b["text"].startswith(marker_prefix)
        ),
        None,
    )
    if existing is None:
        page_edit.edit_page_content(backend, page_id, new_text, "append")
    else:
        page_edit.edit_page_content(
            backend, page_id, new_text, "replace", target_object_id=existing["object_id"]
        )

    after_page = read.get_page(backend, page_id)
    assert any(
        b["type"] == "paragraph" and b["text"] == new_text
        for o in after_page["outlines"]
        for b in o["blocks"]
    ), "the edit must stick"
    after = files.get_page_files_info(backend, page_id)
    assert {e["preferred_name"]: e["kind"] for e in after} == {
        e["preferred_name"]: e["kind"] for e in before
    }, "attachments must be untouched by a pruned edit"
    assert [i["is_printout"] for i in read.get_page_images(backend, page_id)] == [True], (
        "the XPSFile carrier + render must survive (they were pruned, so merge ignores them)"
    )


# --- Phase-6 delete probes (on throwaway copies, never the fixture pages) ----------------


def test_delete_page_level_inserted_file_probe(backend, attachment_pages, temp_section):
    """Ground-truth expectation: a page-level InsertedFile (own objectID) is a page-level
    object, so DeletePageContent should accept it — but it is UNVERIFIED. Probe + report."""
    page = copy.transfer_page(backend, attachment_pages[PAGE_1], temp_section)
    docx = next(
        e
        for e in files.get_page_files_info(backend, page.page_id)
        if e["placement"] == "page_level" and e["extension"] == "docx"
    )
    try:
        backend.delete_page_content(page.page_id, docx["object_id"])
    except OneNoteError as exc:
        pytest.skip(f"DELETE PROBE page-level InsertedFile: REFUSED — {exc}")
    remaining = {e["preferred_name"] for e in files.get_page_files_info(backend, page.page_id)}
    pytest.skip(
        "DELETE PROBE page-level InsertedFile: ACCEPTED — "
        f"docx gone: {docx['preferred_name'] not in remaining}; remaining={sorted(remaining)}"
    )


def test_delete_inline_attachment_oe_probe(backend, attachment_pages, temp_section):
    """OPEN VM QUESTION: DeletePageContent refuses paragraph OEs (0x8004200E) — is an
    attachment-bearing OE refused too (⇒ inline delete = outline rewrite in Phase 6)?"""
    page = copy.transfer_page(backend, attachment_pages[PAGE_1], temp_section)
    txt = next(
        e
        for e in files.get_page_files_info(backend, page.page_id)
        if e["placement"] == "inline" and e["extension"] == "txt"
    )
    try:
        backend.delete_page_content(page.page_id, txt["object_id"])
    except OneNoteError as exc:
        pytest.skip(f"DELETE PROBE inline attachment OE: REFUSED — {exc}")
    remaining = {e["preferred_name"] for e in files.get_page_files_info(backend, page.page_id)}
    pytest.skip(
        "DELETE PROBE inline attachment OE: ACCEPTED — "
        f"txt gone: {txt['preferred_name'] not in remaining}; remaining={sorted(remaining)}"
    )
