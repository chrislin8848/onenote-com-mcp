"""Phase 2 read tools, wired through the service layer onto FixtureBackend (Linux green).

Most assertions run against the REAL VM dumps in tests/fixtures/ (the same call shapes the
tools make: list_sections → GetHierarchy(notebook, hsSections); get_current_context →
window IDs + scoped GetHierarchy; get_page / get_page_images → GetPageContent /
GetBinaryPageContent; search_pages → FindPages).

Two tools query at a scope the notebook-scoped dump didn't capture — list_notebooks
(GetHierarchy("", hsNotebooks)) and list_pages (GetHierarchy(section, hsPages)). Their parse
correctness is already covered by Phase 1's real-fixture tests; here they get tiny inline-XML
control-flow tests (per the project's "inline XML for control-flow seams only" rule) that
verify the service wiring and field projection.
"""

from __future__ import annotations

import base64

import pytest

from onenote_com_mcp.backend.fixture import FixtureBackend
from onenote_com_mcp.errors import NoCurrentWindowError
from onenote_com_mcp.service import read

NOTEBOOK_ID = "{C94E632E-9829-45FF-914E-5E4031B2439D}{1}{B0}"
SECTION_ID = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{B0}"  # "Phase 0 測試用"
GROUP_ID = "{5724F5C7-9310-4BF6-9BE2-108011605C2E}{1}{B0}"  # "節群組 測試用"
MIXED_PAGE_ID = (
    "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E19540013362017467321520163829129860902849621}"
)
TABLE_PAGE_ID = (
    "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E19500773287729139935320149797721816501902621}"
)
IMAGE_PAGE_ID = (
    "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E1953306013858222940101982353039053288030011}"
)
CURRENT_PAGE_ID = (
    "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E1948435450880818651941999633141658297826701}"
)
ONE = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _be(fixtures_dir) -> FixtureBackend:
    return FixtureBackend(fixtures_dir)


# --- list_sections (real notebook-scoped dump) ---------------------------------------


def test_list_sections_preserves_mixed_nesting(fixtures_dir):
    nodes = read.list_sections(_be(fixtures_dir), NOTEBOOK_ID)
    # mixed Section + SectionGroup list, document order, recycle bin filtered (SPEC v0611)
    assert [(n["type"], n["name"]) for n in nodes] == [
        ("section", "Phase 0 測試用"),
        ("section_group", "節群組 測試用"),
    ]
    group = nodes[1]
    assert group["id"] == GROUP_ID
    assert [s["name"] for s in group["children"]] == ["第1節", "第2節"]
    assert nodes[0]["id"] == SECTION_ID
    assert nodes[0]["color"] == "#8AA8E4"


# --- search_pages (real FindPages dump) ----------------------------------------------


def test_search_pages_returns_flat_page_list(fixtures_dir):
    # a fully-CJK query sanitizes to "root" → find__root.xml (FixtureBackend keys on query)
    results = read.search_pages(_be(fixtures_dir), "測試查詢", scope_id=NOTEBOOK_ID)
    assert [p["name"] for p in results] == ["測試頁面1", "測試頁面2"]
    assert all(p["type"] == "page" for p in results)
    assert all(p["id"].startswith("{42020881-") for p in results)


def test_search_pages_records_scope(fixtures_dir):
    be = _be(fixtures_dir)
    read.search_pages(be, "測試查詢", scope_id=NOTEBOOK_ID)
    # FixtureBackend doesn't record reads, but the call must not raise and must scope by query
    assert read.search_pages(be, "root") is not None


# --- get_page (real content pages — lossless runs+style, structured tables) ----------


def test_get_page_shell_and_quick_styles(fixtures_dir):
    page = read.get_page(_be(fixtures_dir), MIXED_PAGE_ID)
    assert page["id"] == MIXED_PAGE_ID
    assert page["name"] == "混合樣式頁"
    assert page["page_level"] == 1
    assert page["last_modified_time"]
    # QuickStyleDef table carried for lossless reconstruction (SPEC §5)
    assert page["quick_styles"]["0"]["name"] == "PageTitle"
    assert page["quick_styles"]["1"]["font_size"] == "11.0"
    assert page["title"]["text"] == "混合樣式頁"


def test_get_page_runs_carry_effective_style(fixtures_dir):
    page = read.get_page(_be(fixtures_dir), MIXED_PAGE_ID)
    blocks = page["outlines"][0]["blocks"]
    # first paragraph: the highlight run with the dual-attribute background
    para = blocks[0]
    assert para["type"] == "paragraph"
    assert para["object_id"]
    hl = para["runs"][1]
    assert hl["text"] == "螢光標示文字"
    assert hl["style"]["background"] == "yellow"
    assert hl["style"]["font-family"] == "Microsoft JhengHei"
    # decoration runs resolve through the three-layer merge
    assert blocks[5]["runs"][0]["style"]["font-weight"] == "bold"
    assert blocks[7]["runs"][0]["style"]["text-decoration"] == "underline"


def test_get_page_table_is_structured_with_object_ids(fixtures_dir):
    page = read.get_page(_be(fixtures_dir), TABLE_PAGE_ID)
    tables = [b for b in page["outlines"][0]["blocks"] if b["type"] == "table"]
    assert len(tables) == 1
    table = tables[0]
    assert table["object_id"]  # needed for delete_page_content (SPEC §5)
    assert table["has_header_row"] is True
    assert len(table["columns"]) == 2
    assert len(table["rows"]) == 10
    # row objectIDs are exposed (parallel to rows) so a whole ROW can be an apply_text_style scope
    assert len(table["row_object_ids"]) == 10
    assert all(table["row_object_ids"]), "every row carries its own objectID"
    # rows are structured lists of cells, never one string
    assert table["rows"][0][0]["text"] == "DAY 1"
    assert table["rows"][0][0]["shading_color"] == "#FFFFCC"
    assert table["rows"][2][0]["text"] == "07:00"
    assert table["rows"][9][1]["text"] == "範例飯店"


def test_get_table_returns_one_table_compactly(fixtures_dir):
    # the table-only read: find the page's table by its objectID and project just it (the compact
    # companion to get_page for a big table-heavy page)
    page = read.get_page(_be(fixtures_dir), TABLE_PAGE_ID)
    table_id = next(b["object_id"] for b in page["outlines"][0]["blocks"] if b["type"] == "table")

    out = read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, table_id)
    assert out["page_id"] == TABLE_PAGE_ID
    assert out["last_modified_time"]
    table = out["table"]
    assert table["object_id"] == table_id
    # same structured shape get_page emits for a table — columns, rows, row/cell objectIDs
    assert len(table["columns"]) == 2
    assert len(table["rows"]) == 10
    assert len(table["row_object_ids"]) == 10
    assert table["rows"][0][0]["text"] == "DAY 1"


def test_get_table_raises_for_unknown_table_id(fixtures_dir):
    from onenote_com_mcp.errors import NodeNotFoundError

    with pytest.raises(NodeNotFoundError, match="no table"):
        read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, "{NOPE}{1}{B0}")


def test_get_page_image_block_has_callback_and_ocr(fixtures_dir):
    page = read.get_page(_be(fixtures_dir), IMAGE_PAGE_ID)
    images = [b for b in page["outlines"][0]["blocks"] if b["type"] == "image"]
    assert len(images) == 1
    img = images[0]
    assert img["callback_id"] == "{4F0825D2-04C7-4D62-A3A3-B517E0CA0F87}{14}{B0}"
    assert img["object_id"]
    assert "富士山" in img["ocr_text"]
    # surrounding paragraph still present
    texts = [b.get("text") for b in page["outlines"][0]["blocks"] if b["type"] == "paragraph"]
    assert "這裡有圖片" in texts


# --- get_page_images (real binary fixture; structured for the MCP Image facade) -------


def test_get_page_images_returns_binary_and_metadata(fixtures_dir):
    images = read.get_page_images(_be(fixtures_dir), IMAGE_PAGE_ID)
    assert len(images) == 1
    img = images[0]
    assert img["callback_id"] == "{4F0825D2-04C7-4D62-A3A3-B517E0CA0F87}{14}{B0}"
    assert img["media_type"] == "image/png"  # sniffed from the PNG magic number
    raw = base64.b64decode(img["data_base64"])
    assert raw[:8] == b"\x89PNG\r\n\x1a\n"
    assert img["object_id"]


def test_get_page_images_empty_when_no_images(fixtures_dir):
    assert read.get_page_images(_be(fixtures_dir), MIXED_PAGE_ID) == []


def test_get_page_surfaces_hyperlinks_on_runs(fixtures_dir):
    # the real 表格頁 fixture has cells with <a href="..."> links; get_page must expose the href
    # on the run (so a replace edit can round-trip it instead of silently dropping the link)
    page = read.get_page(_be(fixtures_dir), TABLE_PAGE_ID)
    links = {
        run.get("link")
        for outline in page["outlines"]
        for block in outline["blocks"]
        if block["type"] == "table"
        for row in block["rows"]
        for cell in row
        for para in cell["paragraphs"]
        for run in para["runs"]
        if run.get("link")
    }
    assert "https://example.com/loc-a" in links
    assert "https://example.com/loc-b" in links


def test_get_page_images_skips_unfetchable_images(fixtures_dir):
    """OCR-processed images return 0x8004200F from GetBinaryPageContent (VM ground truth).
    get_page_images must SKIP them, not crash — get_page remains the authority on what's there.
    """
    from onenote_com_mcp.errors import OneNoteComError

    class _NoBinaryBackend(FixtureBackend):
        def get_binary_page_content(self, page_id, callback_id):
            raise OneNoteComError("GetBinaryPageContent failed", hresult=0x8004200F)

    be = _NoBinaryBackend(fixtures_dir)
    # the image page has one image; with its binary un-fetchable the result is empty, not an error
    assert read.get_page_images(be, IMAGE_PAGE_ID) == []


# --- get_current_context (real window IDs + scoped GetHierarchy for names) -------------


def test_get_current_context_resolves_names(fixtures_dir):
    ctx = read.get_current_context(_be(fixtures_dir))
    assert ctx["notebook"] == {"id": NOTEBOOK_ID, "name": "MCP Test"}
    assert ctx["section_group"] is None  # current section sits directly in the notebook
    assert ctx["section"] == {"id": SECTION_ID, "name": "Phase 0 測試用"}
    assert ctx["page"] == {"id": CURRENT_PAGE_ID, "name": "測試頁面4"}


def test_get_current_context_no_window_raises(tmp_path):
    (tmp_path / "current_window.json").write_text("null", encoding="utf-8")
    with pytest.raises(NoCurrentWindowError):
        read.get_current_context(FixtureBackend(tmp_path))


# --- list_notebooks / list_pages: service wiring (tiny inline XML, control-flow only) -


def test_list_notebooks_projects_summary_fields(tmp_path):
    # GetHierarchy("", hsNotebooks) → the list-all-notebooks container shape
    (tmp_path / "hierarchy_hsNotebooks.xml").write_text(
        f'<one:Notebooks xmlns:one="{ONE}">'
        f'<one:Notebook name="Work" ID="{{NB1}}{{1}}{{B0}}" color="#FF0000" '
        f'lastModifiedTime="2026-06-11T00:00:00.000Z"/>'
        f'<one:Notebook name="Personal" ID="{{NB2}}{{1}}{{B0}}"/>'
        f"</one:Notebooks>",
        encoding="utf-8",
    )
    nbs = read.list_notebooks(FixtureBackend(tmp_path))
    assert [n["name"] for n in nbs] == ["Work", "Personal"]
    assert nbs[0]["id"] == "{NB1}{1}{B0}"
    assert nbs[0]["color"] == "#FF0000"
    assert all("children" not in n for n in nbs)  # notebooks scope doesn't descend


def test_list_pages_includes_page_level(tmp_path):
    # GetHierarchy(section, hsPages) → that section's pages, with subpage levels
    sec = "{SEC}{1}{B0}"
    (tmp_path / "hierarchy_hsPages__SEC_1_B0.xml").write_text(
        f'<one:Section xmlns:one="{ONE}" name="S" ID="{sec}">'
        f'<one:Page ID="{{P1}}{{1}}{{B0}}" name="Top" pageLevel="1" '
        f'dateTime="2026-06-01T00:00:00.000Z" lastModifiedTime="2026-06-02T00:00:00.000Z"/>'
        f'<one:Page ID="{{P2}}{{1}}{{B0}}" name="Sub" pageLevel="2"/>'
        f"</one:Section>",
        encoding="utf-8",
    )
    pages = read.list_pages(FixtureBackend(tmp_path), sec)
    assert [(p["name"], p["page_level"]) for p in pages] == [("Top", 1), ("Sub", 2)]
    assert pages[0]["id"] == "{P1}{1}{B0}"
    assert pages[0]["date_time"] == "2026-06-01T00:00:00.000Z"


# --- page-level objects (printout) + get_page_info inventory --------------------------
# Real printout fixture "附件與嵌入物件-1": one page-level render Image + two page-level
# InsertedFiles (a .docx + an .xlsx) + a one:XPSFile, all direct one:Page children OUTSIDE
# any outline. These were INVISIBLE in get_page before the fix (only `outlines` was emitted),
# so the printout images could not be found/deleted.
PRINTOUT_PAGE_ID = (
    "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E19113778280437107490620108457794036705249631}"
)
PAGE_LEVEL_IMAGE_ID = "{FF3818B8-3EE5-0E18-33E2-DFECB54FC950}{77}{B0}"


def test_get_page_surfaces_page_level_images_and_files(fixtures_dir):
    page = read.get_page(_be(fixtures_dir), PRINTOUT_PAGE_ID)
    # the previously-omitted page-level objects are now in the output
    imgs = page["page_level_images"]
    assert [i["object_id"] for i in imgs] == [PAGE_LEVEL_IMAGE_ID]
    names = sorted(f["preferred_name"] for f in page["page_level_files"])
    assert len(page["page_level_files"]) == 2
    assert any(n.endswith(".docx") for n in names) and any(n.endswith(".xlsx") for n in names)


def test_get_page_info_lists_the_page_level_image_with_delete_tool(fixtures_dir):
    info = read.get_page_info(_be(fixtures_dir), PRINTOUT_PAGE_ID)
    images = [o for o in info["objects"] if o["type"] == "image"]
    # the page-level render image is present, flagged page-level, routed to delete_page_content
    pl = [o for o in images if o["page_level"]]
    assert any(o["object_id"] == PAGE_LEVEL_IMAGE_ID for o in pl)
    for o in pl:
        assert o["delete_with"] == "delete_page_content"


def test_get_page_info_is_flat_exhaustive_and_lightweight(fixtures_dir):
    info = read.get_page_info(_be(fixtures_dir), PRINTOUT_PAGE_ID)
    objs = info["objects"]
    types = {o["type"] for o in objs}
    assert {"image", "file"} <= types  # at least the page-level media surfaced
    # inline objects route to delete_inline_content; page-level to delete_page_content
    assert {o["delete_with"] for o in objs} <= {"delete_inline_content", "delete_page_content"}
    assert any(o["page_level"] for o in objs) and any(not o["page_level"] for o in objs)
    # paragraphs carry a short preview, NOT full runs; no heavy fields leak in
    for o in objs:
        assert "runs" not in o and "quick_styles" not in o
        if o["type"] == "paragraph":
            assert "preview" in o and len(o["preview"]) <= 41  # 40 + ellipsis


def test_get_page_info_matches_inline_image_page(fixtures_dir):
    # a page whose images are INLINE (inside the outline) → listed, delete_inline_content
    info = read.get_page_info(_be(fixtures_dir), IMAGE_PAGE_ID)
    images = [o for o in info["objects"] if o["type"] == "image"]
    assert images, "the image page must surface its inline image in the inventory"
    assert all(not o["page_level"] for o in images)
    assert all(o["delete_with"] == "delete_inline_content" for o in images)
    assert all(o["object_id"] for o in images)
