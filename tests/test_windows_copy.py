"""Tier 2 — live-COM COPY fidelity (SPEC §5/§6 — Phase 5 Stage 3). Windows only.

The B≡A check: clone the real dump-source pages and assert the copy's SEMANTIC fingerprint
equals the source's — title, paragraph runs with effective styles (QuickStyleDef indices may
renumber on a copy; the user-visible style must not change), table structure incl. shading,
image dimensions — plus byte-identical image pixels. Then the tree levels: copy_section
preserves page order + subpage levels, de-collides its name live, and lands inside section
groups. (No copy_notebook: COM notebook creation is refused on this build.)

Self-cleaning: throwaway sections/notebooks carry the "P5暫存"/"P5克隆" prefixes, get swept at
module setup and deleted in teardown. Notebook deletion on a live OneNote can be refused —
that cleanup is best-effort (leftovers are swept next run).
"""

from __future__ import annotations

import base64
import sys
import uuid

import pytest

from onenote_com_mcp.service import copy, read

pytestmark = pytest.mark.windows

TEST_NOTEBOOK = "MCP Test"
TEST_SECTION = "Phase 0 測試用"
TEMP_PREFIX = "P5暫存"
CLONE_NB_PREFIX = "P5克隆"


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
        if node["name"].startswith((TEMP_PREFIX, TEST_SECTION + " (")):
            backend.delete_hierarchy(node["id"], permanent=True)
    for other in read.list_notebooks(backend):
        if other["name"].startswith((CLONE_NB_PREFIX, "P5probe")):
            try:
                backend.delete_hierarchy(other["id"], permanent=True)
            except Exception as exc:  # noqa: BLE001 — best-effort leftover sweep
                print(f"leftover notebook sweep failed: {exc}", file=sys.stderr)
    return nb["id"]


def _is_empty_outline(blocks) -> bool:
    return all(b["type"] == "paragraph" and not b["text"] and not b.get("children") for b in blocks)


def _sweep_empty_outlines(backend, page_id):
    """Phase-4 test wear: past edits can leave outlines holding only empty paragraphs. An
    outline IS a page-level object, so DeletePageContent is the right tool here (unlike
    paragraph OEs, which it refuses). Best-effort with a retry — the fingerprint ignores
    empty outlines anyway (they render as nothing and OneNote drops them on a transplant)."""
    for _ in range(2):
        doomed = [
            o["object_id"]
            for o in read.get_page(backend, page_id)["outlines"]
            if _is_empty_outline(o["blocks"])
        ]
        if not doomed:
            return
        for object_id in doomed:
            backend.delete_page_content(page_id, object_id)
    leftover = [
        o["object_id"]
        for o in read.get_page(backend, page_id)["outlines"]
        if _is_empty_outline(o["blocks"])
    ]
    if leftover:
        print(f"empty-outline sweep did not stick on {page_id}: {leftover}", file=sys.stderr)


@pytest.fixture(scope="module")
def fixture_pages(backend, notebook_id):
    sec = _find(read.list_sections(backend, notebook_id), TEST_SECTION)
    if sec is None:
        pytest.skip(f"{TEST_SECTION!r} section not found")
    pages = {p["name"]: p["id"] for p in read.list_pages(backend, sec["id"])}
    for page_id in pages.values():
        _sweep_empty_outlines(backend, page_id)
    return pages


@pytest.fixture
def temp_section(backend, notebook_id):
    from onenote_com_mcp.service import create

    section_id = create.create_section(backend, notebook_id, _temp_name("節"))
    yield section_id
    backend.delete_hierarchy(section_id, permanent=True)


# --- the semantic fingerprint (B≡A currency) --------------------------------------------


def _norm_runs(block) -> list[tuple[str, tuple]]:
    runs: list[tuple[str, tuple]] = []
    for r in block["runs"]:
        key = tuple(sorted(r["style"].items()))
        if runs and runs[-1][1] == key:
            runs[-1] = (runs[-1][0] + r["text"], key)
        else:
            runs.append((r["text"], key))
    return runs


def _block_fingerprint(block):
    if block["type"] == "paragraph":
        fp = ("p", _norm_runs(block), block["alignment"])
        if block.get("children"):
            return (*fp, [_block_fingerprint(c) for c in block["children"]])
        return fp
    if block["type"] == "table":
        return (
            "table",
            block["borders_visible"],
            block["has_header_row"],
            [[(cell["text"], cell["shading_color"]) for cell in row] for row in block["rows"]],
        )
    return ("image", block["width"], block["height"])


def _fingerprint(backend, page_id):
    page = read.get_page(backend, page_id)
    return {
        "title": page["title"]["text"] if page["title"] else None,
        # empty outlines render as nothing and OneNote drops them on a transplant — they are
        # not part of the semantic B≡A claim
        "outlines": [
            [_block_fingerprint(b) for b in o["blocks"]]
            for o in page["outlines"]
            if not _is_empty_outline(o["blocks"])
        ],
    }


def _image_bytes(backend, page_id):
    return [base64.b64decode(i["data_base64"]) for i in read.get_page_images(backend, page_id)]


# --- copy_page: B≡A ----------------------------------------------------------------------


@pytest.mark.parametrize("page_name", ["混合樣式頁", "表格頁", "圖片頁"])
def test_copy_page_b_equals_a(backend, fixture_pages, temp_section, page_name):
    source_id = fixture_pages.get(page_name)
    if source_id is None:
        pytest.skip(f"{page_name} not present")

    new_id = copy.transfer_page(backend, source_id, temp_section).page_id

    assert new_id != source_id
    assert [p["id"] for p in read.list_pages(backend, temp_section)] == [new_id]
    assert _fingerprint(backend, new_id) == _fingerprint(backend, source_id), (
        f"the {page_name} copy must be semantically identical to the source"
    )
    if page_name == "圖片頁":
        assert _image_bytes(backend, new_id) == _image_bytes(backend, source_id), (
            "image pixels must survive the clone byte-identically"
        )


def test_copy_page_preserves_subpage_level(backend, fixture_pages, temp_section):
    parent_id = fixture_pages.get("單節多頁")
    sub_id = fixture_pages.get("測試頁面7")
    if parent_id is None or sub_id is None:
        pytest.skip("subpage fixtures not present")
    copy.transfer_page(backend, parent_id, temp_section)  # a level-1 anchor first
    copy.transfer_page(backend, sub_id, temp_section)  # source is a level-2 subpage
    landed = read.list_pages(backend, temp_section)
    assert [p["page_level"] for p in landed] == [1, 2], "source pageLevel survives the clone"


# --- copy_section ------------------------------------------------------------------------


def test_copy_section_preserves_order_levels_and_decollides_name(backend, notebook_id):
    source = _find(read.list_sections(backend, notebook_id), TEST_SECTION)
    if source is None:
        pytest.skip(f"{TEST_SECTION!r} section not found")
    src_pages = read.list_pages(backend, source["id"])

    new_section_id = copy.transfer_section(backend, source["id"], notebook_id).section_id
    try:
        cloned = next(
            n for n in read.list_sections(backend, notebook_id) if n["id"] == new_section_id
        )
        assert cloned["name"] == f"{TEST_SECTION} (2)", (
            "same-parent copy must de-collide, not merge into the source"
        )
        copied_pages = read.list_pages(backend, new_section_id)
        assert [p["name"] for p in copied_pages] == [p["name"] for p in src_pages]
        assert [p["page_level"] for p in copied_pages] == [p["page_level"] for p in src_pages], (
            "the 1/2/3 subpage hierarchy survives whole"
        )
        # content spot-check on one cloned page
        mixed_src = next(p["id"] for p in src_pages if p["name"] == "混合樣式頁")
        mixed_copy = next(p["id"] for p in copied_pages if p["name"] == "混合樣式頁")
        assert _fingerprint(backend, mixed_copy) == _fingerprint(backend, mixed_src)
    finally:
        backend.delete_hierarchy(new_section_id, permanent=True)


# --- copy_section into a section group ----------------------------------------------------
# (There is deliberately no copy_notebook test or tool: OpenHierarchy(cftNotebook) refuses
# both local paths and OneDrive https parents with hrFileDoesNotExist on this M365 build —
# VM ground truth 2026-06-11. Whole-notebook cloning = copy_section per section.)


def test_copy_section_into_section_group_live(backend, notebook_id):
    group = _find(read.list_sections(backend, notebook_id), "節群組 測試用")
    if group is None or _find(group["children"], "第2節") is None:
        pytest.skip("節群組 測試用/第2節 not present")
    source = _find(group["children"], "第2節")

    new_id = copy.transfer_section(backend, source["id"], group["id"]).section_id
    try:
        regroup = _find(read.list_sections(backend, notebook_id), "節群組 測試用")
        cloned = next(c for c in regroup["children"] if c["id"] == new_id)
        assert cloned["name"] == "第2節 (2)", "lands INSIDE the group, name de-collided"
        assert [p["name"] for p in read.list_pages(backend, new_id)] == [
            p["name"] for p in read.list_pages(backend, source["id"])
        ]
    finally:
        backend.delete_hierarchy(new_id, permanent=True)
