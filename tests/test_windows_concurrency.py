"""Tier 2 — does UpdateHierarchy carry optimistic-concurrency protection? (fact-finding).

The type library and Microsoft's docs already settle the SIGNATURE axis: ``UpdateHierarchy``
takes only ``(bstrChangesXmlIn, xsSchema)`` — no ``dateExpectedLastModified``, no ``force`` —
whereas ``UpdatePageContent`` / ``DeleteHierarchy`` / ``DeletePageContent`` all carry the date
guard. These tests close the BEHAVIOR axis on real COM, with no inference:

  1. last-write-wins: read a section's page order, let a SECOND COM call reorder it (standing
     in for a concurrent human/add-in edit), then submit the STALE first read. If the stale
     write silently applies and clobbers the middle change → no concurrency protection.
  2. body is not a token: submit a reorder whose XML body carries an obviously-stale
     ``lastModifiedTime`` on the section and every page. If it still applies → UpdateHierarchy
     does not read concurrency information from the XML body either (it has neither a param
     nor a body-based check).

Raw backend on purpose: ``apply_hierarchy_restructure`` would add ID-conservation, which is
ours, not OneNote's — the question is what the COM call itself does. Self-cleaning: one
throwaway "P6並發" section per test, permanently deleted in teardown; leftovers swept at setup.
"""

from __future__ import annotations

import uuid

import pytest
from lxml import etree

from onenote_com_mcp.enums import HierarchyScope
from onenote_com_mcp.errors import ConcurrencyError, OneNoteError
from onenote_com_mcp.service import create, read
from onenote_com_mcp.xmllayer.namespaces import qn

pytestmark = pytest.mark.windows

TEST_NOTEBOOK = "MCP Test"
TEMP_PREFIX = "P6並發"
_STALE_TIME = "2000-01-01T00:00:00.000Z"


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
        if node["type"] == "section" and node["name"].startswith(TEMP_PREFIX):
            backend.delete_hierarchy(node["id"], permanent=True)
    return nb["id"]


@pytest.fixture
def three_page_section(backend, notebook_id):
    """A throwaway section with three identifiable pages, in a known order."""
    name = f"{TEMP_PREFIX}-{uuid.uuid4().hex[:8]}"
    section_id = create.create_section(backend, notebook_id, name)
    for title in ("頁甲", "頁乙", "頁丙"):
        create.create_page(backend, section_id, title)
    yield section_id
    backend.delete_hierarchy(section_id, permanent=True)


def _page_order(backend, section_id) -> list[str]:
    """Current page names in document order, straight from a fresh GetHierarchy."""
    raw = backend.get_hierarchy(section_id, HierarchyScope.hsPages)
    tree = etree.fromstring(raw.encode("utf-8"))
    section = next(e for e in tree.iter() if e.get("ID") == section_id)
    return [p.get("name") for p in section.findall(qn("Page"))]


def _reorder_xml(raw_xml: str, section_id: str, name_order: list[str]) -> str:
    """Return the section XML with its Page children rearranged into ``name_order``."""
    tree = etree.fromstring(raw_xml.encode("utf-8"))
    section = next(e for e in tree.iter() if e.get("ID") == section_id)
    pages = {p.get("name"): p for p in section.findall(qn("Page"))}
    for name in name_order:
        section.append(pages[name])  # appending an existing child MOVES it → final order
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8").decode("utf-8")


def test_update_hierarchy_is_silent_last_write_wins(backend, three_page_section):
    section_id = three_page_section
    assert _page_order(backend, section_id) == ["頁甲", "頁乙", "頁丙"]

    # READ A — the stale snapshot a client would hold (carries the pages' real stamps).
    stale_xml = backend.get_hierarchy(section_id, HierarchyScope.hsPages)

    # MIDDLE CHANGE — a second COM call reverses the order, standing in for a concurrent edit.
    backend.update_hierarchy(_reorder_xml(stale_xml, section_id, ["頁丙", "頁乙", "頁甲"]))
    assert _page_order(backend, section_id) == ["頁丙", "頁乙", "頁甲"], "middle change took effect"

    # STALE WRITE — resubmit READ A verbatim (original order + original stamps). If OneNote
    # guarded concurrency this would raise; we assert it does NOT, and that it clobbers.
    try:
        backend.update_hierarchy(stale_xml)
    except (ConcurrencyError, OneNoteError) as exc:  # pragma: no cover — would refute the thesis
        pytest.fail(f"UpdateHierarchy rejected a stale write — it DOES guard concurrency: {exc!r}")

    assert _page_order(backend, section_id) == ["頁甲", "頁乙", "頁丙"], (
        "the stale write silently overwrote the middle change → last-write-wins, NO protection"
    )


def test_update_hierarchy_ignores_body_last_modified_time(backend, three_page_section):
    section_id = three_page_section

    raw_xml = backend.get_hierarchy(section_id, HierarchyScope.hsPages)
    tree = etree.fromstring(raw_xml.encode("utf-8"))
    section = next(e for e in tree.iter() if e.get("ID") == section_id)
    # backdate every lastModifiedTime in the body to an obviously-stale value
    for el in tree.iter():
        if el.get("lastModifiedTime"):
            el.set("lastModifiedTime", _STALE_TIME)
    pages = {p.get("name"): p for p in section.findall(qn("Page"))}
    for name in ("頁丙", "頁甲", "頁乙"):
        section.append(pages[name])
    mangled = etree.tostring(tree, xml_declaration=True, encoding="UTF-8").decode("utf-8")

    try:
        backend.update_hierarchy(mangled)
    except (ConcurrencyError, OneNoteError) as exc:  # pragma: no cover — would refute the thesis
        pytest.fail(f"a stale body lastModifiedTime was rejected — body IS a token: {exc!r}")

    assert _page_order(backend, section_id) == ["頁丙", "頁甲", "頁乙"], (
        "reorder applied despite a stale body stamp → UpdateHierarchy ignores the body's "
        "lastModifiedTime; it has no concurrency token at all"
    )
