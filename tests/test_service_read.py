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
import json

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


def _table_id(fixtures_dir):
    page = read.get_page(_be(fixtures_dir), TABLE_PAGE_ID)
    return next(b["object_id"] for b in page["outlines"][0]["blocks"] if b["type"] == "table")


def test_get_table_text_only_is_a_plain_string_grid(fixtures_dir):
    # text_only: just each cell's text — no runs / style / objectIDs (the compact data read)
    tid = _table_id(fixtures_dir)
    full = read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, tid)["table"]
    out = read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, tid, text_only=True)
    table = out["table"]
    assert table["object_id"] == tid
    assert table["total_rows"] == 10 and table["total_columns"] == 2
    assert table["rows"] == [[c["text"] for c in row] for row in full["rows"]]
    assert all(isinstance(x, str) for row in table["rows"] for x in row)
    assert "row_object_ids" not in table and "paragraphs" not in json.dumps(table)
    compact, verbose = (len(json.dumps(t, ensure_ascii=False)) for t in (table, full))
    assert compact < verbose / 5


def test_get_table_window_rows_and_columns(fixtures_dir):
    tid = _table_id(fixtures_dir)
    grid = read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, tid, text_only=True)["table"]["rows"]
    out = read.get_table(
        _be(fixtures_dir), TABLE_PAGE_ID, tid, text_only=True, start_row=2, max_rows=3, columns=[1]
    )["table"]
    assert out["rows"] == [[grid[r][1]] for r in (2, 3, 4)]
    assert out["start_row"] == 2 and out["returned_rows"] == 3
    assert out["column_indices"] == [1] and out["total_rows"] == 10
    # the full model windows the same way, keeping row ids / column widths parallel to the slice
    full = read.get_table(
        _be(fixtures_dir), TABLE_PAGE_ID, tid, start_row=9, max_rows=5, columns=[1, 0]
    )["table"]
    assert len(full["rows"]) == 1 and len(full["row_object_ids"]) == 1
    assert [c["text"] for c in full["rows"][0]] == [grid[9][1], grid[9][0]]
    assert len(full["columns"]) == 2 and full["total_columns"] == 2


def test_get_table_flags_write_cost_only_on_a_big_table(fixtures_dir, monkeypatch):
    tid = _table_id(fixtures_dir)
    small = read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, tid, text_only=True)["table"]
    assert "write_cost" not in small  # 10×2 — edits are quick
    monkeypatch.setattr(read, "_BIG_TABLE_CELLS", 20)
    big = read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, tid, text_only=True)["table"]
    assert "batch" in big["write_cost"] and "20 cells" in big["write_cost"]
    assert "OPEN in OneNote" not in big["write_cost"]  # the fixture's current page is another one


def test_get_table_write_cost_says_when_the_page_is_on_screen(fixtures_dir, monkeypatch):
    from onenote_com_mcp.backend.base import CurrentWindowIds

    be = _be(fixtures_dir)
    be.get_current_window_ids = lambda: CurrentWindowIds(None, None, None, TABLE_PAGE_ID)
    monkeypatch.setattr(read, "_BIG_TABLE_CELLS", 20)
    out = read.get_table(be, TABLE_PAGE_ID, _table_id(fixtures_dir), text_only=True)
    assert "switch OneNote to another page" in out["table"]["write_cost"]


def test_get_table_rejects_bad_window(fixtures_dir):
    tid = _table_id(fixtures_dir)
    with pytest.raises(ValueError, match="out of range"):
        read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, tid, columns=[2])
    with pytest.raises(ValueError, match="start_row"):
        read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, tid, start_row=-1)
    with pytest.raises(ValueError, match="max_rows"):
        read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, tid, max_rows=0)


def test_get_table_raises_for_unknown_table_id(fixtures_dir):
    from onenote_com_mcp.errors import NodeNotFoundError

    with pytest.raises(NodeNotFoundError, match="no table"):
        read.get_table(_be(fixtures_dir), TABLE_PAGE_ID, "{NOPE}{1}{B0}")


def test_get_object_returns_one_paragraph_full(fixtures_dir):
    # the targeted single-object read: same full text + runs + style get_page gives for that
    # paragraph, but without the rest of the page
    page = read.get_page(_be(fixtures_dir), MIXED_PAGE_ID)
    para = page["outlines"][0]["blocks"][0]

    out = read.get_object(_be(fixtures_dir), MIXED_PAGE_ID, para["object_id"])
    assert out["page_id"] == MIXED_PAGE_ID
    assert out["last_modified_time"]
    obj = out["object"]
    assert obj["type"] == "paragraph"
    assert obj["object_id"] == para["object_id"]
    assert obj["text"] == para["text"]
    assert obj["runs"] == para["runs"]  # full per-run style, identical to get_page's projection


def test_get_object_returns_a_table(fixtures_dir):
    page = read.get_page(_be(fixtures_dir), TABLE_PAGE_ID)
    table_id = next(b["object_id"] for b in page["outlines"][0]["blocks"] if b["type"] == "table")

    obj = read.get_object(_be(fixtures_dir), TABLE_PAGE_ID, table_id)["object"]
    assert obj["type"] == "table"
    assert obj["object_id"] == table_id
    assert len(obj["rows"]) == 10


def _texts(blocks):
    """Flatten a full get_page outline's blocks to the text_only shape, for comparison."""
    out = []
    for b in blocks:
        if b["type"] == "table":
            item = {"table": [[c["text"] for c in row] for row in b["rows"]]}
        elif b["type"] == "image":
            item = {"image": b["ocr_text"] or ""}
        elif b["type"] == "file":
            item = {"file": b["preferred_name"]}
        elif b.get("children"):
            item = {"text": b["text"]}
        else:
            item = b["text"]
        if b.get("children"):
            item["children"] = _texts(b["children"])
        out.append(item)
    return out


@pytest.mark.parametrize("page_id", ["MIXED_PAGE_ID", "TABLE_PAGE_ID", "IMAGE_PAGE_ID"])
def test_get_page_text_only_is_the_words_of_the_full_read(fixtures_dir, page_id):
    pid = globals()[page_id]
    full = read.get_page(_be(fixtures_dir), pid)
    text = read.get_page(_be(fixtures_dir), pid, text_only=True)
    assert text["id"] == pid and text["last_modified_time"] == full["last_modified_time"]
    assert text["title"] == (full["title"]["text"] if full["title"] else None)
    assert text["outlines"] == [_texts(o["blocks"]) for o in full["outlines"]]
    blob = json.dumps(text, ensure_ascii=False)
    assert "object_id" not in blob and "runs" not in blob and "font" not in blob
    assert len(blob) < len(json.dumps(full, ensure_ascii=False)) / 3


def test_get_page_text_only_shows_tables_and_images_compactly(fixtures_dir):
    table = read.get_page(_be(fixtures_dir), TABLE_PAGE_ID, text_only=True)
    grid = next(b["table"] for o in table["outlines"] for b in o if isinstance(b, dict))
    assert grid[0][0] == "DAY 1" and grid[9][1] == "範例飯店"
    images = read.get_page(_be(fixtures_dir), IMAGE_PAGE_ID, text_only=True)
    assert any(isinstance(b, dict) and "image" in b for o in images["outlines"] for b in o)


def test_get_page_text_only_lists_page_level_objects(fixtures_dir):
    out = read.get_page(_be(fixtures_dir), PRINTOUT_PAGE_ID, text_only=True)
    full = read.get_page(_be(fixtures_dir), PRINTOUT_PAGE_ID)
    assert len(out["page_level_images"]) == len(full["page_level_images"]) > 0
    assert out["page_level_files"] == [f["preferred_name"] for f in full["page_level_files"]]


def test_get_object_text_only(fixtures_dir):
    page = read.get_page(_be(fixtures_dir), MIXED_PAGE_ID)
    para = page["outlines"][0]["blocks"][0]
    out = read.get_object(_be(fixtures_dir), MIXED_PAGE_ID, para["object_id"], text_only=True)
    assert out["object"] == para["text"]

    tpage = read.get_page(_be(fixtures_dir), TABLE_PAGE_ID)
    tbl = next(b for b in tpage["outlines"][0]["blocks"] if b["type"] == "table")
    obj = read.get_object(_be(fixtures_dir), TABLE_PAGE_ID, tbl["object_id"], text_only=True)
    assert obj["object"] == {"table": [[c["text"] for c in row] for row in tbl["rows"]]}


def test_get_object_unknown_id_raises(fixtures_dir):
    from onenote_com_mcp.errors import NodeNotFoundError

    with pytest.raises(NodeNotFoundError, match="no object"):
        read.get_object(_be(fixtures_dir), MIXED_PAGE_ID, "{NOPE}{1}{B0}")
    with pytest.raises(ValueError, match="object_id is empty"):
        read.get_object(_be(fixtures_dir), MIXED_PAGE_ID, "")


def test_find_objects_locates_paragraph_by_text(fixtures_dir):
    out = read.find_objects(_be(fixtures_dir), MIXED_PAGE_ID, "螢光標示文字")
    assert out["page_id"] == MIXED_PAGE_ID
    matches = out["matches"]
    assert matches and all(m["type"] == "paragraph" for m in matches)
    assert any("螢光" in m["preview"] for m in matches)
    # the returned id is a real, editable paragraph objectID (matches get_page's projection)
    page = read.get_page(_be(fixtures_dir), MIXED_PAGE_ID)
    para_ids = {b["object_id"] for b in page["outlines"][0]["blocks"] if b["type"] == "paragraph"}
    assert any(m["object_id"] in para_ids for m in matches)


def test_find_objects_matches_full_text_not_truncated_preview(fixtures_dir):
    # find_objects must search the FULL paragraph text, not the 40-char preview get_page_info shows
    page = read.get_page(_be(fixtures_dir), MIXED_PAGE_ID)
    long = next(
        (
            b
            for b in page["outlines"][0]["blocks"]
            if b["type"] == "paragraph" and len(b["text"]) > 45
        ),
        None,
    )
    if long is None:
        pytest.skip("no paragraph longer than the preview cap in this fixture")
    needle = long["text"][41:46]  # a slice past where get_page_info's preview would have cut off
    matches = read.find_objects(_be(fixtures_dir), MIXED_PAGE_ID, needle)["matches"]
    assert any(m["object_id"] == long["object_id"] for m in matches)


def test_find_objects_no_match_and_empty_query(fixtures_dir):
    assert read.find_objects(_be(fixtures_dir), MIXED_PAGE_ID, "絕不存在的字串XYZ")["matches"] == []
    with pytest.raises(ValueError, match="query is empty"):
        read.find_objects(_be(fixtures_dir), MIXED_PAGE_ID, "")


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


# --- get_page_info: big-table cell paragraphs are summarized, not listed --------------------


def _cell_paragraph_ids(fixtures_dir):
    full = read.get_page(_be(fixtures_dir), TABLE_PAGE_ID)
    table = next(b for b in full["outlines"][0]["blocks"] if b["type"] == "table")
    return table["object_id"], {
        p["object_id"] for row in table["rows"] for c in row for p in c["paragraphs"]
    }


def test_get_page_info_small_table_still_lists_every_cell(fixtures_dir):
    # 10×2 is under the threshold: unchanged behaviour, every cell paragraph is listed
    _, cell_ids = _cell_paragraph_ids(fixtures_dir)
    listed = {
        o["object_id"] for o in read.get_page_info(_be(fixtures_dir), TABLE_PAGE_ID)["objects"]
    }
    assert cell_ids <= listed


def test_get_page_info_big_table_summarizes_its_cells(fixtures_dir, monkeypatch):
    monkeypatch.setattr(read, "_COLLAPSE_TABLE_CELLS", 10)
    tid, cell_ids = _cell_paragraph_ids(fixtures_dir)
    objs = read.get_page_info(_be(fixtures_dir), TABLE_PAGE_ID)["objects"]
    listed = {o["object_id"] for o in objs}
    assert not (cell_ids & listed), "no cell paragraph is enumerated"
    entry = next(o for o in objs if o["object_id"] == tid)
    assert entry["cell_paragraphs_not_listed"] == len(cell_ids)
    assert "find_objects" in entry["cells_note"] and "include_cells" in entry["cells_note"]

    expanded = read.get_page_info(_be(fixtures_dir), TABLE_PAGE_ID, include_cells=True)["objects"]
    assert cell_ids <= {o["object_id"] for o in expanded}
    assert "cell_paragraphs_not_listed" not in next(o for o in expanded if o["object_id"] == tid)


def test_get_page_info_big_table_still_lists_images_in_its_cells(tmp_path, monkeypatch):
    # "find/delete every image" must stay complete: an image inside a collapsed table's cell is
    # still an inventory entry
    from onenote_com_mcp.backend.fixture import _sanitize

    monkeypatch.setattr(read, "_COLLAPSE_TABLE_CELLS", 2)
    pid = "{P}{1}{B0}"

    def cell(oid, inner):
        return (
            f'<one:Cell objectID="{{C{oid}}}"><one:OEChildren>{inner}</one:OEChildren></one:Cell>'
        )

    text = '<one:OE objectID="{{T{0}}}"><one:T><![CDATA[x{0}]]></one:T></one:OE>'
    img = (
        '<one:OE objectID="{IMG-OE}"><one:Image>'
        '<one:CallbackID callbackID="{CB}"/></one:Image></one:OE>'
    )
    xml = (
        '<?xml version="1.0"?><one:Page '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{pid}" lastModifiedTime="2026-10-05T00:00:00.000Z"><one:Outline objectID="{{O}}">'
        '<one:OEChildren><one:OE objectID="{TOE}"><one:Table objectID="{TBL}"><one:Columns>'
        '<one:Column index="0" width="50"/><one:Column index="1" width="50"/></one:Columns>'
        f'<one:Row objectID="{{R1}}">{cell(1, text.format(1))}{cell(2, img)}</one:Row>'
        f'<one:Row objectID="{{R2}}">{cell(3, text.format(3))}{cell(4, text.format(4))}</one:Row>'
        "</one:Table></one:OE></one:OEChildren></one:Outline></one:Page>"
    )
    (tmp_path / f"page_{_sanitize(pid)}.xml").write_text(xml, encoding="utf-8")
    objs = read.get_page_info(FixtureBackend(tmp_path), pid)["objects"]
    assert [o["type"] for o in objs] == ["table", "image"]
    assert objs[0]["cell_paragraphs_not_listed"] == 3
    assert objs[1]["object_id"] == "{IMG-OE}"
