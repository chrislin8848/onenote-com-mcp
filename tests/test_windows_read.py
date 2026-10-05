"""Tier 2 — live-COM READ integration (SPEC §2.4, Phase 3). Windows + OneNote only.

Auto-skipped off Windows by conftest; on the VM these run in the autologon interactive
session via scripts/remote_test.sh. They drive the SAME stack production uses
(``Win32ComBackend`` → ``xmllayer`` → ``service.read``), so they prove three things the
host-side fixtures can't:

  1. the early-bound COM connection + ``[out] BSTR`` contract actually round-trips live XML;
  2. ``Windows.CurrentWindow`` PROPERTY marshalling works under early binding (the open
     question flagged in win32com_backend.get_current_window_ids);
  3. the parser handles freshly-served XML, not just the committed dump.

Tests are NAME-driven (discover the "MCP Test" notebook → its sections → its pages by name),
not ID-hardcoded, so they survive a re-created notebook. The whole module skips cleanly if
the "MCP Test" notebook isn't present, so a bare VM reports "not set up" rather than errors.

Phase 3 is READ-only integration; create→update→delete round-trips are Phase 4's Tier 2.
"""

from __future__ import annotations

import base64

import pytest

from onenote_com_mcp.errors import NoCurrentWindowError
from onenote_com_mcp.service import read

pytestmark = pytest.mark.windows

TEST_NOTEBOOK = "MCP Test"
TEST_SECTION = "Phase 0 測試用"


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
        pytest.skip(f"{TEST_NOTEBOOK!r} notebook not present on this VM — nothing to read")
    return nb["id"]


@pytest.fixture(scope="module")
def section_id(backend, notebook_id):
    sec = _find(read.list_sections(backend, notebook_id), TEST_SECTION)
    if sec is None:
        pytest.skip(f"{TEST_SECTION!r} section not found in {TEST_NOTEBOOK!r}")
    return sec["id"]


@pytest.fixture(scope="module")
def pages_by_name(backend, section_id):
    return {p["name"]: p["id"] for p in read.list_pages(backend, section_id)}


# --- connection + hierarchy read -----------------------------------------------------


def test_connect_and_list_notebooks(backend):
    notebooks = read.list_notebooks(backend)
    assert isinstance(notebooks, list)
    assert all(n["id"] and n["name"] for n in notebooks)


def test_list_sections_mixed_nesting_live(backend, notebook_id):
    nodes = read.list_sections(backend, notebook_id)
    # the test notebook has both a plain section and a section group (SPEC v0611)
    assert any(n["type"] == "section" for n in nodes)
    groups = [n for n in nodes if n["type"] == "section_group"]
    assert groups, "expected the '節群組 測試用' section group to survive a live read"
    assert all("children" in g for g in groups)


def test_list_pages_have_levels(backend, section_id):
    pages = read.list_pages(backend, section_id)
    assert pages, "expected pages in the test section"
    assert all(isinstance(p["page_level"], int) for p in pages)


# --- page content: format preservation on a LIVE read -------------------------------


def test_get_page_highlight_survives_live_read(backend, pages_by_name):
    pid = pages_by_name.get("混合樣式頁")
    if pid is None:
        pytest.skip("混合樣式頁 not present")
    page = read.get_page(backend, pid)
    runs = [
        r
        for o in page["outlines"]
        for b in o["blocks"]
        if b["type"] == "paragraph"
        for r in b["runs"]
    ]
    # the dual-attribute highlight (background:) must come back from live COM, not just the dump
    assert any(r["style"].get("background") for r in runs), "highlight lost on live read"
    assert any(r["style"].get("font-weight") == "bold" for r in runs)


def test_get_page_table_is_structured_live(backend, pages_by_name):
    pid = pages_by_name.get("表格頁")
    if pid is None:
        pytest.skip("表格頁 not present")
    page = read.get_page(backend, pid)
    tables = [b for o in page["outlines"] for b in o["blocks"] if b["type"] == "table"]
    assert tables, "expected a table block"
    table = tables[0]
    assert table["object_id"]
    assert len(table["rows"]) >= 2
    assert all(isinstance(row, list) for row in table["rows"])  # structured, never one string


def test_get_table_reads_one_table_live(backend, pages_by_name):
    pid = pages_by_name.get("表格頁")
    if pid is None:
        pytest.skip("表格頁 not present")
    table_id = next(
        b["object_id"]
        for o in read.get_page(backend, pid)["outlines"]
        for b in o["blocks"]
        if b["type"] == "table"
    )
    out = read.get_table(backend, pid, table_id)
    assert out["page_id"] == pid
    assert out["table"]["object_id"] == table_id
    assert len(out["table"]["rows"]) >= 2
    assert all(isinstance(row, list) for row in out["table"]["rows"])
    # text_only: a plain string grid matching the full model's cell texts; windowing slices it
    grid = [[c["text"] for c in row] for row in out["table"]["rows"]]
    text = read.get_table(backend, pid, table_id, text_only=True)["table"]
    assert text["rows"] == grid
    one = read.get_table(
        backend, pid, table_id, text_only=True, start_row=1, max_rows=1, columns=[0]
    )
    assert one["table"]["rows"] == [[grid[1][0]]]


def test_get_page_images_returns_binary_live(backend, pages_by_name):
    pid = pages_by_name.get("圖片頁")
    if pid is None:
        pytest.skip("圖片頁 not present")
    images = read.get_page_images(backend, pid)
    assert images, "expected at least one image"
    raw = base64.b64decode(images[0]["data_base64"])
    assert len(raw) > 0
    assert images[0]["media_type"].startswith("image/")
    assert images[0]["object_id"]  # the enclosing OE id (images carry no objectID of their own)


# --- current context: the property-marshalling check (the Phase 3 open question) -----


def test_get_current_context_property_marshalling(backend):
    """The point is that Windows.CurrentWindow's Current*Id PROPERTIES marshal under early
    binding without a COM error — a real window's exact location can't be asserted here."""
    try:
        ctx = read.get_current_context(backend)
    except NoCurrentWindowError:
        return  # acceptable: no window open. The property path still didn't crash.
    assert set(ctx) == {"notebook", "section_group", "section", "page"}
    # whatever's open, an id present should resolve to a name (or be None)
    for slot in ("notebook", "section", "page"):
        if ctx[slot] is not None:
            assert "id" in ctx[slot] and "name" in ctx[slot]


# --- 1.3.0 targeted reads: get_object (one object) + find_objects (locate by text) ----


def test_get_object_reads_one_paragraph_live(backend, pages_by_name):
    pid = pages_by_name.get("混合樣式頁")
    if pid is None:
        pytest.skip("混合樣式頁 not present")
    para = next(
        b
        for o in read.get_page(backend, pid)["outlines"]
        for b in o["blocks"]
        if b["type"] == "paragraph"
    )
    out = read.get_object(backend, pid, para["object_id"])
    assert out["page_id"] == pid
    obj = out["object"]
    assert obj["type"] == "paragraph"
    assert obj["object_id"] == para["object_id"]
    # the single-object read returns the SAME full text + per-run style get_page gives
    assert obj["text"] == para["text"]
    assert obj["runs"] == para["runs"]


def test_find_objects_locates_paragraph_by_text_live(backend, pages_by_name):
    pid = pages_by_name.get("混合樣式頁")
    if pid is None:
        pytest.skip("混合樣式頁 not present")
    paras = [
        b
        for o in read.get_page(backend, pid)["outlines"]
        for b in o["blocks"]
        if b["type"] == "paragraph" and (b.get("text") or "").strip()
    ]
    if not paras:
        pytest.skip("no text paragraphs on 混合樣式頁")
    target = paras[0]
    needle = (target["text"] or "")[:6]  # a real leading slice — full-text match, not a preview
    out = read.find_objects(backend, pid, needle)
    assert any(m["object_id"] == target["object_id"] for m in out["matches"])


def test_get_page_text_only_live(backend, pages_by_name):
    pid = pages_by_name.get("表格頁")
    if pid is None:
        pytest.skip("表格頁 not present")
    full = read.get_page(backend, pid)
    text = read.get_page(backend, pid, text_only=True)
    blocks = [b for o in text["outlines"] for b in o]
    grids = [b["table"] for b in blocks if isinstance(b, dict) and "table" in b]
    full_tables = [b for o in full["outlines"] for b in o["blocks"] if b["type"] == "table"]
    assert grids == [[[c["text"] for c in row] for row in t["rows"]] for t in full_tables]
    assert "object_id" not in str(text)
