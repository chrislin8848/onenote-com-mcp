"""parse_hierarchy against the real VM hierarchy dumps (MCP Test notebook, 2026-06-11)."""

from __future__ import annotations

from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend
from onenote_com_mcp.enums import HierarchyScope
from onenote_com_mcp.xmllayer.namespaces import NSMAP, qn
from onenote_com_mcp.xmllayer.parse import parse_hierarchy

NOTEBOOK_ID = "{C94E632E-9829-45FF-914E-5E4031B2439D}{1}{B0}"
SECTION_ID = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{B0}"


def _backend(fixtures_dir):
    return FixtureBackend(fixtures_dir)


def test_notebook_scope_single_node_no_children(fixtures_dir):
    xml = _backend(fixtures_dir).get_hierarchy(NOTEBOOK_ID, HierarchyScope.hsNotebooks)
    nodes = parse_hierarchy(xml)
    assert len(nodes) == 1
    nb = nodes[0]
    assert nb["type"] == "notebook"
    assert nb["id"] == NOTEBOOK_ID
    assert nb["name"] == "MCP Test"
    assert nb["nickname"] == "MCP Test"
    assert nb["color"] == "#EE9597"
    assert nb["is_currently_viewed"] is True
    assert nb["last_modified_time"] == "2026-06-11T03:34:14.000Z"
    assert nb["children"] == []


def test_sections_scope_mixed_children_preserve_order(fixtures_dir):
    """A notebook's direct children are a MIXED Section + SectionGroup list — keep nesting

    and document order, no flattening (SPEC v0611)."""
    xml = _backend(fixtures_dir).get_hierarchy(NOTEBOOK_ID, HierarchyScope.hsSections)
    (nb,) = parse_hierarchy(xml)
    kinds = [(c["type"], c["name"]) for c in nb["children"]]
    # recycle bin group is skipped by default; order otherwise preserved
    assert kinds == [
        ("section", "Phase 0 測試用"),
        ("section_group", "節群組 測試用"),
    ]
    group = nb["children"][1]
    assert [s["name"] for s in group["children"]] == ["第1節", "第2節"]
    assert group["id"] == "{5724F5C7-9310-4BF6-9BE2-108011605C2E}{1}{B0}"
    sec = nb["children"][0]
    assert sec["id"] == SECTION_ID
    assert sec["color"] == "#8AA8E4"


def test_recycle_bin_included_on_request(fixtures_dir):
    xml = _backend(fixtures_dir).get_hierarchy(NOTEBOOK_ID, HierarchyScope.hsSections)
    (nb,) = parse_hierarchy(xml, include_recycle_bin=True)
    names = [c["name"] for c in nb["children"]]
    assert names == ["Phase 0 測試用", "OneNote_RecycleBin", "節群組 測試用"]
    bin_group = nb["children"][1]
    assert bin_group["is_recycle_bin"] is True
    assert [s["name"] for s in bin_group["children"]] == ["刪除的頁面", "新的節 1"]


def test_pages_scope_page_levels_and_ids(fixtures_dir):
    xml = _backend(fixtures_dir).get_hierarchy(NOTEBOOK_ID, HierarchyScope.hsPages)
    (nb,) = parse_hierarchy(xml)
    sec = nb["children"][0]
    assert sec["name"] == "Phase 0 測試用"
    pages = sec["children"]
    assert [p["type"] for p in pages] == ["page"] * 8
    assert [p["name"] for p in pages] == [
        "混合樣式頁",
        "表格頁",
        "圖片頁",
        "單節多頁",
        "測試頁面7",
        "測試頁面6",
        "測試頁面5",
        "測試頁面4",
    ]
    # pageLevel is the subpage indent — must survive parsing (copy needs it, SPEC §4/§5)
    assert [p["page_level"] for p in pages] == [1, 1, 1, 1, 2, 3, 3, 2]
    assert pages[0]["id"] == (
        "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E19540013362017467321520163829129860902849621}"
    )
    # pages have both timestamps
    assert pages[0]["date_time"] == "2026-06-10T17:05:01.000Z"
    assert pages[0]["last_modified_time"]
    # section-group sections carry their pages too
    group = nb["children"][1]
    assert [len(s["children"]) for s in group["children"]] == [2, 1]


def test_find_pages_result_parses_with_same_function(fixtures_dir):
    xml = _backend(fixtures_dir).find_pages(NOTEBOOK_ID, "root")
    (nb,) = parse_hierarchy(xml)
    group = nb["children"][0]
    assert group["type"] == "section_group"
    sec = group["children"][0]
    assert [p["name"] for p in sec["children"]] == ["測試頁面1", "測試頁面2"]


def test_container_root_notebooks(fixtures_dir):
    """GetHierarchy from the root returns <one:Notebooks> — handle the container root."""
    xml = _backend(fixtures_dir).get_hierarchy(NOTEBOOK_ID, HierarchyScope.hsNotebooks)
    nb_el = etree.fromstring(xml.encode("utf-8"))
    container = etree.Element(qn("Notebooks"), nsmap=NSMAP)
    container.append(nb_el)
    nodes = parse_hierarchy(etree.tostring(container, encoding="unicode"))
    assert len(nodes) == 1
    assert nodes[0]["type"] == "notebook"
    assert nodes[0]["name"] == "MCP Test"
