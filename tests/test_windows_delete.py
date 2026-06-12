"""Tier 2 — live-COM delete round-trips (SPEC §6, Phase 6 Stage 4). Windows only.

delete_node (DeleteHierarchy, recycle-bin default) and delete_page_content (page-level objects
only). The interesting live checks: a recycled node disappears from the list tools; a page-level
Outline / Image / InsertedFile is actually removed by DeletePageContent; and the service's
page-level pre-check turns the VM-verified inline-OE refusal into a clean ValueError WITHOUT
reaching COM.

Everything operates on THROWAWAY nodes ("P6暫存" prefix) and copies of the fixture pages — never
the fixture pages themselves. Self-cleaning: the prefix is swept at module setup and sections are
deleted in teardown (best-effort, since some tests delete their own section).
"""

from __future__ import annotations

import sys
import uuid

import pytest

from onenote_com_mcp.service import copy, create, delete, files, read

pytestmark = pytest.mark.windows

TEST_NOTEBOOK = "MCP Test"
TEST_SECTION = "Phase 0 測試用"
PAGE_1 = "附件與嵌入物件-1"
TEMP_PREFIX = "P6暫存"


def _find(nodes, name):
    return next((n for n in nodes if n.get("name") == name), None)


def _temp_name(tag: str) -> str:
    return f"{TEMP_PREFIX}-{tag}-{uuid.uuid4().hex[:8]}"


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
def attachment_page(backend, notebook_id):
    sec = _find(read.list_sections(backend, notebook_id), TEST_SECTION)
    if sec is None:
        pytest.skip(f"{TEST_SECTION!r} section not found")
    page = _find(read.list_pages(backend, sec["id"]), PAGE_1)
    if page is None:
        pytest.skip(f"{PAGE_1!r} attachment page not found")
    return page["id"]


@pytest.fixture
def temp_section(backend, notebook_id):
    section_id = create.create_section(backend, notebook_id, _temp_name("節"))
    yield section_id
    try:  # best-effort: a test may have deleted it already
        backend.delete_hierarchy(section_id, permanent=True)
    except Exception as exc:  # noqa: BLE001
        print(f"temp_section teardown: {exc}", file=sys.stderr)


# --- delete_node (hierarchy) -------------------------------------------------------------


def test_delete_node_section_to_recycle_bin(backend, notebook_id):
    section_id = create.create_section(backend, notebook_id, _temp_name("刪節"))
    assert any(s["id"] == section_id for s in read.list_sections(backend, notebook_id))

    delete.delete_node(backend, section_id)  # recycle-bin default

    # list_sections filters the recycle bin, so a recycled section is gone from the listing
    assert not any(s["id"] == section_id for s in read.list_sections(backend, notebook_id))


def test_delete_node_page(backend, temp_section):
    page_id = create.create_page(backend, temp_section, _temp_name("刪頁"))
    assert any(p["id"] == page_id for p in read.list_pages(backend, temp_section))

    delete.delete_node(backend, page_id)

    assert not any(p["id"] == page_id for p in read.list_pages(backend, temp_section))


def test_delete_node_permanent(backend, notebook_id):
    section_id = create.create_section(backend, notebook_id, _temp_name("永久刪"))
    delete.delete_node(backend, section_id, permanent=True)
    assert not any(s["id"] == section_id for s in read.list_sections(backend, notebook_id))


# --- delete_page_content (page-level objects) --------------------------------------------


def test_delete_page_level_outline(backend, temp_section):
    page_id = create.create_page(backend, temp_section, _temp_name("大綱"), content="一\n二\n三")
    outlines = read.get_page(backend, page_id)["outlines"]
    assert outlines, "the new page must have a content outline to delete"

    delete.delete_page_content(backend, page_id, outlines[0]["object_id"])

    assert read.get_page(backend, page_id)["outlines"] == [], "the outline must be gone"


def test_delete_page_level_attachment(backend, attachment_page, temp_section):
    page_id = copy.transfer_page(backend, attachment_page, temp_section).page_id
    before = files.get_page_files_info(backend, page_id)
    docx = next(e for e in before if e["placement"] == "page_level" and e["extension"] == "docx")

    delete.delete_page_content(backend, page_id, docx["object_id"])

    after = {e["object_id"] for e in files.get_page_files_info(backend, page_id)}
    assert docx["object_id"] not in after
    assert len(after) == len(before) - 1, "only the targeted attachment is removed"


def test_delete_page_level_image(backend, attachment_page, temp_section):
    page_id = copy.transfer_page(backend, attachment_page, temp_section).page_id
    images = read.get_page_images(backend, page_id)
    assert images, "the copied attachment page carries the (flattened) render image"

    delete.delete_page_content(backend, page_id, images[0]["object_id"])

    assert not read.get_page_images(backend, page_id), "the page-level image must be gone"


def test_inline_oe_refused_before_com(backend, attachment_page, temp_section):
    """The service's page-level pre-check rejects an inline OE with a clear ValueError — the
    VM-verified DeletePageContent refusal (0x8004200E) never has to surface."""
    page_id = copy.transfer_page(backend, attachment_page, temp_section).page_id
    paragraph = next(
        b
        for o in read.get_page(backend, page_id)["outlines"]
        for b in o["blocks"]
        if b["type"] == "paragraph"
    )
    with pytest.raises(ValueError, match="inline content|a paragraph"):
        delete.delete_page_content(backend, page_id, paragraph["object_id"])
    # and the object is still there — nothing was deleted
    still = read.get_page(backend, page_id)["outlines"]
    assert any(b["object_id"] == paragraph["object_id"] for o in still for b in o["blocks"])
