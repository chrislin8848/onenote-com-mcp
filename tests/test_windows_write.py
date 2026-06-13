"""Tier 2 — live-COM WRITE integration (SPEC §2.4, §6 — Phase 4 Stage 4). Windows only.

What only real OneNote can prove, in test order:

  1. create→get→update round-trips through the production seams (create / page_edit / read);
  2. the format-preservation regression (SPEC §6): edit ONE paragraph of the real mixed-format
     page, assert every untouched paragraph's text + effective per-run style is unchanged;
  3. BOTH payload strategies (changed_objects default, whole_page) leave an image intact when
     its own outline is edited — this run DECIDES the provisional DEFAULT_PAYLOAD_STRATEGY;
  4. the dateExpectedLastModified concurrency guard actually trips (→ ConcurrencyError) and
     force=True actually overrides it;
  5. whole-batch hierarchy mutators round-trip (restructure / rename / reorder-with-recycle-
     bin-pinned) and the EXPERIMENTAL move_page gate (SPEC §4/§9) passes or stays experimental.

Self-cleaning: throwaway sections are named "P4暫存…" and deleted permanently in teardown;
leftovers from a crashed previous run are swept at module setup. The two edits that touch the
dump-source pages (混合樣式頁 / 圖片頁) are reverted in finally blocks through the edit seam
(outline rewrite — DeletePageContent refuses paragraph-level OEs).
"""

from __future__ import annotations

import base64
import datetime as dt
import uuid

import pytest
from lxml import etree

from onenote_com_mcp.enums import PageInfo
from onenote_com_mcp.errors import ConcurrencyError
from onenote_com_mcp.service import create, hierarchy_edit, page_edit, read
from onenote_com_mcp.xmllayer.namespaces import qn

pytestmark = pytest.mark.windows

TEST_NOTEBOOK = "MCP Test"
TEST_SECTION = "Phase 0 測試用"
TEMP_PREFIX = "P4暫存"
PNG_1PX = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk"
    "+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)


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
        pytest.skip(f"{TEST_NOTEBOOK!r} notebook not present on this VM — nothing to write to")
    # sweep leftovers from a crashed previous run, so reruns stay deterministic
    for node in read.list_sections(backend, nb["id"]):
        if node["type"] == "section" and node["name"].startswith(TEMP_PREFIX):
            backend.delete_hierarchy(node["id"], permanent=True)
    return nb["id"]


@pytest.fixture
def temp_section(backend, notebook_id):
    section_id = create.create_section(backend, notebook_id, _temp_name("節"))
    yield section_id
    backend.delete_hierarchy(section_id, permanent=True)


@pytest.fixture(scope="module")
def fixture_pages(backend, notebook_id):
    sec = _find(read.list_sections(backend, notebook_id), TEST_SECTION)
    if sec is None:
        pytest.skip(f"{TEST_SECTION!r} section not found")
    pages = {p["name"]: p["id"] for p in read.list_pages(backend, sec["id"])}
    # sweep leftover markers a crashed previous run left on the dump-source pages
    for name in ("混合樣式頁", "圖片頁"):
        if pages.get(name):
            _remove_marker_paragraphs(backend, pages[name], "P4")
    return pages


def _paragraphs(backend, page_id):
    return [
        b
        for o in read.get_page(backend, page_id)["outlines"]
        for b in o["blocks"]
        if b["type"] == "paragraph"
    ]


def _remove_marker_paragraphs(backend, page_id, prefix):
    """Remove test paragraphs surgically through the edit seam. VM ground truth (2026-06-11):
    DeletePageContent REFUSES paragraph-level OEs (hr=0x8004200E) — it is for page-level
    objects. UpdatePageContent's merge replaces a submitted outline's content wholesale, so
    dropping the OE from its outline IS the paragraph delete."""
    doomed = {
        b["object_id"]
        for b in _paragraphs(backend, page_id)
        if (b["text"] or "").startswith(prefix)
    }
    if not doomed:
        return

    def mutate(tree):
        for oe in list(tree.iter(qn("OE"))):
            if oe.get("objectID") in doomed:
                oe.getparent().remove(oe)

    page_edit.apply_page_edit(backend, page_id, mutate)


# --- 1. create → get → update round-trips ---------------------------------------------


def test_create_page_title_and_content_roundtrip(backend, temp_section):
    page_id = create.create_page(backend, temp_section, "建立測試頁", "第一段\n第二段")
    page = read.get_page(backend, page_id)
    assert page["title"]["text"] == "建立測試頁"
    assert [p["text"] for p in _paragraphs(backend, page_id)] == ["第一段", "第二段"]
    assert [p["id"] for p in read.list_pages(backend, temp_section)] == [page_id]


def test_create_page_level_sets_subpage_indent(backend, temp_section):
    parent = create.create_page(backend, temp_section, "父頁")
    child = create.create_page(backend, temp_section, "子頁", page_level=2)
    levels = {p["id"]: p["page_level"] for p in read.list_pages(backend, temp_section)}
    assert levels[parent] == 1
    assert levels[child] == 2


def test_update_modes_roundtrip_with_styles(backend, temp_section):
    page_id = create.create_page(backend, temp_section, "編輯模式頁", "甲\n乙\n丙")
    mid_id = _paragraphs(backend, page_id)[1]["object_id"]

    page_edit.edit_page_content(backend, page_id, "乙後", "insert_after", target_object_id=mid_id)
    first_id = _paragraphs(backend, page_id)[0]["object_id"]
    page_edit.edit_page_content(
        backend, page_id, "甲前", "insert_before", target_object_id=first_id
    )
    page_edit.edit_page_content(
        backend,
        page_id,
        [{"runs": [{"text": "改寫", "style": {"font-weight": "bold", "background": "yellow"}}]}],
        "replace",
        target_object_id=mid_id,
    )

    final = _paragraphs(backend, page_id)
    assert [p["text"] for p in final] == ["甲前", "甲", "改寫", "乙後", "丙"]
    style = final[2]["runs"][0]["style"]
    assert style.get("font-weight") == "bold"
    assert style.get("background"), "highlight must survive the live round-trip"


# --- 2. format-preservation regression (SPEC §6) on the REAL mixed-format page ----------


def _normalized_runs(backend, page_id):
    """Per paragraph: adjacent runs with identical effective style merged (OneNote may re-chunk
    spans on redraw — semantic identity is the SPEC §5 guarantee, not byte identity)."""
    result = []
    for block in _paragraphs(backend, page_id):
        runs: list[tuple[str, tuple]] = []
        for r in block["runs"]:
            key = tuple(sorted(r["style"].items()))
            if runs and runs[-1][1] == key:
                runs[-1] = (runs[-1][0] + r["text"], key)
            else:
                runs.append((r["text"], key))
        result.append(runs)
    return result


def test_format_preservation_regression(backend, fixture_pages):
    page_id = fixture_pages.get("混合樣式頁")
    if page_id is None:
        pytest.skip("混合樣式頁 not present")
    marker = f"P4回歸-{uuid.uuid4().hex[:6]}"
    before = _normalized_runs(backend, page_id)

    page_edit.edit_page_content(backend, page_id, marker)
    try:
        after = _normalized_runs(backend, page_id)
        # OneNote splits fresh text into per-script font runs — match the concatenation
        kept = [runs for runs in after if "".join(t for t, _ in runs) != marker]
        assert len(after) == len(before) + 1, "exactly one paragraph was added"
        assert kept == before, (
            "untouched paragraphs' text/effective styles must be byte-for-byte identical"
        )
    finally:
        _remove_marker_paragraphs(backend, page_id, marker)
    assert _normalized_runs(backend, page_id) == before, "page restored to its pristine state"


# --- 3. payload strategies: an edit must never cost an image its pixels -----------------


def _image_bytes(backend, page_id):
    return [base64.b64decode(i["data_base64"]) for i in read.get_page_images(backend, page_id)]


@pytest.mark.parametrize("strategy", ["changed_objects", "whole_page"])
def test_edit_in_image_outline_keeps_pixels(backend, fixture_pages, strategy):
    page_id = fixture_pages.get("圖片頁")
    if page_id is None:
        pytest.skip("圖片頁 not present")
    marker = f"P4圖測-{strategy}-{uuid.uuid4().hex[:6]}"
    before = _image_bytes(backend, page_id)
    assert before, "the fixture image page must have an image"
    page = read.get_page(backend, page_id)
    image_outline = next(
        o["object_id"] for o in page["outlines"] if any(b["type"] == "image" for b in o["blocks"])
    )

    mutate = page_edit.content_mutator(marker, "append", image_outline)
    page_edit.apply_page_edit(backend, page_id, mutate, strategy=strategy)
    try:
        after = _image_bytes(backend, page_id)
        assert after == before, f"strategy {strategy!r} must not disturb the image binary"
    finally:
        _remove_marker_paragraphs(backend, page_id, marker)


# --- 4. concurrency guard ----------------------------------------------------------------


def test_concurrency_guard_trips_then_force_overrides(backend, temp_section):
    """VM ground truth (2026-06-11): GetPageContent's lastModifiedTime does NOT refresh right
    after a programmatic UpdatePageContent (polled unchanged for 10s), so the guard is
    exercised with an explicitly wrong stamp rather than an intervening edit."""
    page_id = create.create_page(backend, temp_section, "並發頁", "原始內容")
    xml = backend.get_page_content(page_id, PageInfo.piBasic)
    actual = page_edit.parse_onenote_datetime(
        etree.fromstring(xml.encode()).get("lastModifiedTime")
    )
    assert actual is not None
    wrong = actual - dt.timedelta(hours=1)

    with pytest.raises(ConcurrencyError):
        backend.update_page_content(xml, expected_last_modified=wrong)
    # the matching stamp passes…
    backend.update_page_content(xml, expected_last_modified=actual)
    # …and explicit opt-in force bypasses the guard entirely (SPEC §5)
    backend.update_page_content(xml, expected_last_modified=wrong, force=True)


# --- 5. tables ---------------------------------------------------------------------------


def test_table_create_append_rows_and_cell_edit(backend, temp_section):
    page_id = create.create_page(backend, temp_section, "表格測試頁")
    page_edit.add_table(backend, page_id, [["品名", "數量"], ["蘋果", "3"]], has_header_row=True)

    def table():
        return next(
            b
            for o in read.get_page(backend, page_id)["outlines"]
            for b in o["blocks"]
            if b["type"] == "table"
        )

    t1 = table()
    assert t1["has_header_row"] is True
    assert [[c["text"] for c in row] for row in t1["rows"]] == [["品名", "數量"], ["蘋果", "3"]]

    # append a row via modify_table (replaces the old create_table append-rows path)
    page_edit.modify_table(backend, page_id, t1["object_id"], "insert_rows", rows=[["香蕉", "5"]])
    t2 = table()
    assert [[c["text"] for c in row] for row in t2["rows"]] == [
        ["品名", "數量"],
        ["蘋果", "3"],
        ["香蕉", "5"],
    ]

    cell_oe = t2["rows"][1][1]["paragraphs"][0]["object_id"]
    page_edit.edit_page_content(backend, page_id, "30", "replace", target_object_id=cell_oe)
    t3 = table()
    assert [[c["text"] for c in row] for row in t3["rows"]] == [
        ["品名", "數量"],
        ["蘋果", "30"],
        ["香蕉", "5"],
    ], "one cell edited; every other cell untouched"


def test_modify_table_shape_roundtrips(backend, temp_section):
    page_id = create.create_page(backend, temp_section, "表格形狀頁")
    page_edit.add_table(backend, page_id, [["a", "b"], ["c", "d"], ["e", "f"]])

    def table():
        return next(
            b
            for o in read.get_page(backend, page_id)["outlines"]
            for b in o["blocks"]
            if b["type"] == "table"
        )

    tid = table()["object_id"]

    # insert a row at position 1
    page_edit.modify_table(backend, page_id, tid, "insert_rows", rows=[["x", "y"]], at_index=1)
    assert [[c["text"] for c in r] for r in table()["rows"]] == [
        ["a", "b"],
        ["x", "y"],
        ["c", "d"],
        ["e", "f"],
    ]

    # add a column at the end → every row gains an (empty) cell, table stays rectangular
    page_edit.modify_table(backend, page_id, tid, "add_columns")
    t = table()
    assert all(len(r) == 3 for r in t["rows"]), "every row has 3 cells after add_columns"
    assert [c["text"] for c in t["rows"][0]] == ["a", "b", ""], "new cell is empty"

    # delete the middle column (index 1) → drops that column from every row
    page_edit.modify_table(backend, page_id, tid, "delete_columns", indices=[1])
    t = table()
    assert all(len(r) == 2 for r in t["rows"])
    assert [c["text"] for c in t["rows"][0]] == ["a", ""]

    # delete the first row
    page_edit.modify_table(backend, page_id, tid, "delete_rows", indices=[0])
    assert [[c["text"] for c in r] for r in table()["rows"]] == [
        ["x", ""],
        ["c", ""],
        ["e", ""],
    ]


# --- 6. images ---------------------------------------------------------------------------


def test_insert_image_roundtrip(backend, temp_section):
    page_id = create.create_page(backend, temp_section, "圖片插入頁")
    page_edit.insert_image(backend, page_id, PNG_1PX, "image/png", width=24.0, height=24.0)
    images = read.get_page_images(backend, page_id)
    assert len(images) == 1
    raw = base64.b64decode(images[0]["data_base64"])
    assert raw.startswith(b"\x89PNG"), "inserted image must come back as a PNG"
    assert raw == base64.b64decode(PNG_1PX), "byte-identical round-trip (loosen if re-encoded)"


# --- 7. hierarchy mutators ---------------------------------------------------------------


def test_restructure_section_roundtrip(backend, temp_section):
    a = create.create_page(backend, temp_section, "頁A")
    b = create.create_page(backend, temp_section, "頁B")
    c = create.create_page(backend, temp_section, "頁C")
    hierarchy_edit.restructure_section(
        backend,
        temp_section,
        [
            {"page_id": c, "page_level": 1},
            {"page_id": a, "page_level": 2},
            {"page_id": b, "page_level": 2},
        ],
    )
    pages = read.list_pages(backend, temp_section)
    assert [p["id"] for p in pages] == [c, a, b]
    assert [p["page_level"] for p in pages] == [1, 2, 2]


def test_rename_page_and_section_roundtrip(backend, notebook_id, temp_section):
    page_id = create.create_page(backend, temp_section, "改名前")
    hierarchy_edit.rename_node(backend, temp_section, page_id, "改名後")
    assert read.list_pages(backend, temp_section)[0]["name"] == "改名後"

    new_section_name = _temp_name("已改名節")
    hierarchy_edit.rename_node(backend, notebook_id, temp_section, new_section_name)
    renamed = _find(read.list_sections(backend, notebook_id), new_section_name)
    assert renamed is not None and renamed["id"] == temp_section


def test_reorder_sections_live_inside_group(backend, notebook_id):
    """OneNote's schema is positional (sections before groups — VM ground truth), so the live
    reorder happens where same-kind siblings exist: inside the 節群組 測試用 group. The
    submitted batch still rides through the same whole-batch core."""

    def group():
        return _find(read.list_sections(backend, notebook_id), "節群組 測試用")

    g = group()
    if g is None or len(g["children"]) < 2:
        pytest.skip("need the 節群組 測試用 group with two sections")
    original = [c["id"] for c in g["children"]]
    swapped = list(reversed(original))

    hierarchy_edit.reorder_sections(backend, g["id"], swapped)
    try:
        assert [c["id"] for c in group()["children"]] == swapped, "live reorder must hold"
    finally:
        hierarchy_edit.reorder_sections(backend, g["id"], original)
    assert [c["id"] for c in group()["children"]] == original


def test_reorder_sections_rejects_interleaved_kinds_live(backend, notebook_id):
    nodes = read.list_sections(backend, notebook_id)
    sections = [n["id"] for n in nodes if n["type"] == "section"]
    groups = [n["id"] for n in nodes if n["type"] == "section_group"]
    if not sections or not groups:
        pytest.skip("need both kinds at the notebook level")
    # groups ahead of sections would be hrInvalidXML live — refused before any write
    with pytest.raises(ValueError, match="section groups"):
        hierarchy_edit.reorder_sections(backend, notebook_id, groups + sections)


def test_move_page_experimental_gate(backend, notebook_id):
    """SPEC §4/§9: move_page is EXPERIMENTAL until this very test passes on real OneNote."""
    source = create.create_section(backend, notebook_id, _temp_name("搬移源"))
    target = create.create_section(backend, notebook_id, _temp_name("搬移標"))
    try:
        keeper = create.create_page(backend, source, "留守頁")
        mover = create.create_page(backend, source, "被搬頁", page_level=2)

        new_id = hierarchy_edit.move_page(backend, notebook_id, mover, target)

        assert [p["id"] for p in read.list_pages(backend, source)] == [keeper]
        landed = read.list_pages(backend, target)
        assert [p["id"] for p in landed] == [new_id], "moved page is the target's only page"
        assert new_id != mover, "VM ground truth: re-parenting assigns a NEW page ID"
        assert landed[0]["page_level"] == 1, "subpage level resets to 1 on move"
        # the moved page's CONTENT survived the hierarchy move (read via the NEW id)
        assert read.get_page(backend, new_id)["title"]["text"] == "被搬頁"
    finally:
        backend.delete_hierarchy(target, permanent=True)
        backend.delete_hierarchy(source, permanent=True)
