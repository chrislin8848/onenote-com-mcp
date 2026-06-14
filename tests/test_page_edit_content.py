"""Tier-1 tests for the Phase-4 content mutators, against REAL VM fixture pages.

Format preservation is asserted the only way it can be on Linux (SPEC §5): every byte of an
untouched paragraph/object must survive the read → mutate → payload round-trip verbatim —
CDATA sections included. The payload-strategy seam (changed_objects vs whole_page) and the
image-binary inlining rule are exercised here too; the VM round-trips (Stage 4) decide which
strategy ships.
"""

from __future__ import annotations

import pytest
from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service import page_edit
from onenote_com_mcp.xmllayer.namespaces import qn

_PARSER = etree.XMLParser(strip_cdata=False)


def _page_root(fixtures_dir, page_name: str) -> etree._Element:
    for path in sorted(fixtures_dir.glob("page_*.xml")):
        if "__binary" in path.name:
            continue
        root = etree.fromstring(path.read_bytes(), parser=_PARSER)
        if root.get("name") == page_name:
            return root
    raise AssertionError(f"no fixture page named {page_name!r}")


def _sent_payload(be: FixtureBackend) -> tuple[dict, etree._Element]:
    writes = [c for c in be.calls if c.method == "update_page_content"]
    assert len(writes) == 1, "every edit must be exactly ONE UpdatePageContent"
    return writes[0].kwargs, etree.fromstring(
        writes[0].kwargs["changes_xml"].encode("utf-8"), parser=_PARSER
    )


def _oes(outline: etree._Element) -> list[etree._Element]:
    return outline.findall(f"{qn('OEChildren')}/{qn('OE')}")


def _cdata(oe: etree._Element) -> str:
    t = oe.find(qn("T"))
    return t.text or ""


@pytest.fixture
def be(fixtures_dir) -> FixtureBackend:
    return FixtureBackend(fixtures_dir)


@pytest.fixture
def mixed(fixtures_dir) -> etree._Element:
    return _page_root(fixtures_dir, "混合樣式頁")


@pytest.fixture
def table_page(fixtures_dir) -> etree._Element:
    return _page_root(fixtures_dir, "表格頁")


@pytest.fixture
def image_page(fixtures_dir) -> etree._Element:
    return _page_root(fixtures_dir, "圖片頁")


# --- append + payload pruning (changed_objects default) -----------------------------


def test_append_sends_only_touched_outline_and_preserves_untouched_oes(be, mixed):
    page_id = mixed.get("ID")
    original_outline = mixed.findall(qn("Outline"))[-1]
    original_oes = [etree.tostring(oe, with_tail=False) for oe in _oes(original_outline)]

    page_edit.edit_page_content(be, page_id, "新增段落")

    kwargs, sent = _sent_payload(be)
    assert kwargs["force"] is False
    assert kwargs["expected_last_modified"] is not None, "concurrency guard must be carried"
    # pruned payload: the touched outline only — Title is not re-sent, definitions ride along
    assert sent.find(qn("Title")) is None
    assert len(sent.findall(qn("QuickStyleDef"))) == len(mixed.findall(qn("QuickStyleDef")))
    outlines = sent.findall(qn("Outline"))
    assert len(outlines) == 1
    sent_oes = _oes(outlines[0])
    assert len(sent_oes) == len(original_oes) + 1
    # untouched paragraphs survive BYTE-IDENTICAL, CDATA sections included
    for before, after in zip(original_oes, sent_oes, strict=False):
        assert etree.tostring(after, with_tail=False) == before
    assert "新增段落" in _cdata(sent_oes[-1])
    assert "<![CDATA[" in kwargs["changes_xml"]


def test_untouched_styled_spans_survive_verbatim(be, mixed):
    page_edit.edit_page_content(be, mixed.get("ID"), "x")
    kwargs, _ = _sent_payload(be)
    # the fixture's own word-wrapped span markup, byte-for-byte
    assert ">粗粗粗</span>" in kwargs["changes_xml"]
    assert "background:yel" in kwargs["changes_xml"]


def test_append_styled_runs_write_dual_highlight(be, mixed):
    content = [{"runs": [{"text": "螢光", "style": {"background": "yellow"}}]}]
    page_edit.edit_page_content(be, mixed.get("ID"), content)
    _, sent = _sent_payload(be)
    new_oe = _oes(sent.findall(qn("Outline"))[-1])[-1]
    assert "background:yellow" in _cdata(new_oe)
    assert "mso-highlight:yellow" in _cdata(new_oe)


def test_append_multiline_string_becomes_multiple_paragraphs(be, mixed):
    n_before = len(_oes(mixed.findall(qn("Outline"))[-1]))
    page_edit.edit_page_content(be, mixed.get("ID"), "第一段\n第二段")
    _, sent = _sent_payload(be)
    sent_oes = _oes(sent.findall(qn("Outline"))[-1])
    assert len(sent_oes) == n_before + 2
    assert "第一段" in _cdata(sent_oes[-2]) and "第二段" in _cdata(sent_oes[-1])


def test_append_creates_outline_on_empty_page(tmp_path):
    from onenote_com_mcp.backend.fixture import _sanitize

    page_id = "{P}{1}{B0}"
    (tmp_path / f"page_{_sanitize(page_id)}.xml").write_text(
        '<?xml version="1.0"?><one:Page '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{page_id}" lastModifiedTime="2026-06-10T17:39:30.000Z"/>',
        encoding="utf-8",
    )
    be = FixtureBackend(tmp_path)
    page_edit.edit_page_content(be, page_id, "hi")
    _, sent = _sent_payload(be)
    outline = sent.find(qn("Outline"))
    assert outline is not None, "a page with no outline gets a fresh one"
    assert "hi" in _cdata(_oes(outline)[0])


# --- insert / replace ----------------------------------------------------------------


def test_insert_after_and_before_position_relative_to_target(be, mixed):
    page_id = mixed.get("ID")
    target_oe = _oes(mixed.findall(qn("Outline"))[-1])[1]
    target_id = target_oe.get("objectID")

    page_edit.edit_page_content(be, page_id, "插在後", "insert_after", target_object_id=target_id)
    _, sent = _sent_payload(be)
    oes = _oes(sent.findall(qn("Outline"))[-1])
    idx = next(i for i, oe in enumerate(oes) if oe.get("objectID") == target_id)
    assert "插在後" in _cdata(oes[idx + 1])

    be2 = FixtureBackend(be.fixtures_dir)
    page_edit.edit_page_content(be2, page_id, "插在前", "insert_before", target_object_id=target_id)
    _, sent2 = _sent_payload(be2)
    oes2 = _oes(sent2.findall(qn("Outline"))[-1])
    idx2 = next(i for i, oe in enumerate(oes2) if oe.get("objectID") == target_id)
    assert "插在前" in _cdata(oes2[idx2 - 1])


def test_replace_swaps_text_but_keeps_identity_and_paragraph_style(be, mixed):
    page_id = mixed.get("ID")
    target_oe = next(  # a styled paragraph (has quickStyleIndex AND an OE style attr)
        oe
        for oe in _oes(mixed.findall(qn("Outline"))[-1])
        if oe.get("quickStyleIndex") and oe.get("style")
    )
    target_id = target_oe.get("objectID")
    page_edit.edit_page_content(be, page_id, "改寫後的字", "replace", target_object_id=target_id)
    _, sent = _sent_payload(be)
    new_oe = next(oe for oe in sent.iter(qn("OE")) if oe.get("objectID") == target_id)
    assert "改寫後的字" in _cdata(new_oe)
    assert len(new_oe.findall(qn("T"))) == 1, "old one:T runs must be gone"
    # paragraph-level formatting is untouched
    assert new_oe.get("quickStyleIndex") == target_oe.get("quickStyleIndex")
    assert new_oe.get("style") == target_oe.get("style")
    assert new_oe.get("author") == target_oe.get("author")


def test_replace_with_extra_paragraphs_inserts_them_after(be, mixed):
    target = _oes(mixed.findall(qn("Outline"))[-1])[0]
    page_edit.edit_page_content(
        be, mixed.get("ID"), "甲\n乙", "replace", target_object_id=target.get("objectID")
    )
    _, sent = _sent_payload(be)
    oes = _oes(sent.findall(qn("Outline"))[-1])
    idx = next(i for i, oe in enumerate(oes) if oe.get("objectID") == target.get("objectID"))
    assert "甲" in _cdata(oes[idx]) and "乙" in _cdata(oes[idx + 1])
    assert oes[idx + 1].get("objectID") is None, "the extra paragraph is NEW (no invented IDs)"


def test_replace_edits_table_cell_text_without_touching_other_cells(be, table_page):
    page_id = table_page.get("ID")
    table = next(table_page.iter(qn("Table")))
    cells = table.findall(f"{qn('Row')}/{qn('Cell')}")
    target_oe = cells[0].find(f"{qn('OEChildren')}/{qn('OE')}")
    other_cell_before = etree.tostring(cells[1], with_tail=False)

    page_edit.edit_page_content(
        be, page_id, "改儲存格", "replace", target_object_id=target_oe.get("objectID")
    )
    _, sent = _sent_payload(be)
    sent_table = next(sent.iter(qn("Table")))
    sent_cells = sent_table.findall(f"{qn('Row')}/{qn('Cell')}")
    assert "改儲存格" in _cdata(sent_cells[0].find(f"{qn('OEChildren')}/{qn('OE')}"))
    assert etree.tostring(sent_cells[1], with_tail=False) == other_cell_before


def test_replace_can_write_a_hyperlink(be, mixed):
    page_id = mixed.get("ID")
    target_oe = next(o for o in mixed.iter(qn("OE")) if o.get("objectID"))
    page_edit.edit_page_content(
        be,
        page_id,
        [{"runs": [{"text": "OneNote", "link": "https://example.com/a?x=1&y=2"}]}],
        "replace",
        target_object_id=target_oe.get("objectID"),
    )
    _, sent = _sent_payload(be)
    edited = next(o for o in sent.iter(qn("OE")) if o.get("objectID") == target_oe.get("objectID"))
    cdata = _cdata(edited)
    assert '<a href="https://example.com/a?x=1&amp;y=2">' in cdata
    assert "OneNote</a>" in cdata


# --- tables ---------------------------------------------------------------------------


def test_add_table_appends_new_table_in_its_own_oe(be, mixed):
    page_edit.add_table(be, mixed.get("ID"), [["a", "b"], ["c"]], has_header_row=True)
    _, sent = _sent_payload(be)
    last_oe = _oes(sent.findall(qn("Outline"))[-1])[-1]
    table = last_oe.find(qn("Table"))
    assert table is not None
    assert table.get("hasHeaderRow") == "true"
    assert len(table.findall(f"{qn('Columns')}/{qn('Column')}")) == 2
    rows = table.findall(qn("Row"))
    assert len(rows) == 2
    assert len(rows[1].findall(qn("Cell"))) == 2, "short rows are padded"


def test_create_table_rejects_an_existing_table_target(be, table_page):
    # shape edits on an existing table go through modify_table now, not create_table
    table = next(table_page.iter(qn("Table")))
    with pytest.raises(ValueError, match="modify_table"):
        page_edit.add_table(
            be, table_page.get("ID"), [["新左", "新右"]], target_object_id=table.get("objectID")
        )
    assert not [c for c in be.calls if c.method == "update_page_content"], "no write on failure"


# --- modify_table: structural edits on an existing table ------------------------------


def test_modify_table_insert_rows_appends_when_at_index_omitted(be, table_page):
    table = next(table_page.iter(qn("Table")))
    rows_before = table.findall(qn("Row"))
    first_row_before = etree.tostring(rows_before[0], with_tail=False)

    page_edit.modify_table(
        be, table_page.get("ID"), table.get("objectID"), "insert_rows", rows=[["新左", "新右"]]
    )
    _, sent = _sent_payload(be)
    sent_rows = next(sent.iter(qn("Table"))).findall(qn("Row"))
    assert len(sent_rows) == len(rows_before) + 1
    assert etree.tostring(sent_rows[0], with_tail=False) == first_row_before
    new_cell = sent_rows[-1].findall(qn("Cell"))[0]
    assert "新左" in _cdata(new_cell.find(f"{qn('OEChildren')}/{qn('OE')}"))


def test_modify_table_insert_rows_at_position(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n_before = len(table.findall(qn("Row")))

    page_edit.modify_table(
        be,
        table_page.get("ID"),
        table.get("objectID"),
        "insert_rows",
        rows=[["X", "Y"]],
        at_index=1,
    )
    _, sent = _sent_payload(be)
    sent_rows = next(sent.iter(qn("Table"))).findall(qn("Row"))
    assert len(sent_rows) == n_before + 1
    assert "X" in _cdata(sent_rows[1].findall(qn("Cell"))[0].find(f"{qn('OEChildren')}/{qn('OE')}"))


def test_modify_table_insert_rows_rejects_rows_wider_than_table(be, table_page):
    table = next(table_page.iter(qn("Table")))
    with pytest.raises(ValueError, match="columns"):
        page_edit.modify_table(
            be, table_page.get("ID"), table.get("objectID"), "insert_rows", rows=[["a", "b", "c"]]
        )
    assert not [c for c in be.calls if c.method == "update_page_content"]


def test_modify_table_insert_columns_appends_column_to_every_row(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n_cols = len(table.findall(f"{qn('Columns')}/{qn('Column')}"))
    n_rows = len(table.findall(qn("Row")))

    page_edit.modify_table(be, table_page.get("ID"), table.get("objectID"), "insert_columns")
    _, sent = _sent_payload(be)
    sent_table = next(sent.iter(qn("Table")))
    cols = sent_table.findall(f"{qn('Columns')}/{qn('Column')}")
    assert len(cols) == n_cols + 1
    assert [c.get("index") for c in cols] == [str(i) for i in range(n_cols + 1)], "re-indexed"
    for row in sent_table.findall(qn("Row")):
        assert len(row.findall(qn("Cell"))) == n_cols + 1  # every row stays rectangular
    # the new cell is a valid, empty cell (OEChildren > OE), never an empty <Cell/>
    last_cell = sent_table.findall(qn("Row"))[0].findall(qn("Cell"))[-1]
    assert last_cell.find(f"{qn('OEChildren')}/{qn('OE')}") is not None
    assert n_rows == len(sent_table.findall(qn("Row")))  # insert_columns doesn't change row count


def test_modify_table_insert_columns_at_position(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n_cols = len(table.findall(f"{qn('Columns')}/{qn('Column')}"))
    page_edit.modify_table(
        be, table_page.get("ID"), table.get("objectID"), "insert_columns", at_index=0, count=2
    )
    _, sent = _sent_payload(be)
    cols = next(sent.iter(qn("Table"))).findall(f"{qn('Columns')}/{qn('Column')}")
    assert len(cols) == n_cols + 2


def test_modify_table_delete_rows(be, table_page):
    table = next(table_page.iter(qn("Table")))
    rows_before = table.findall(qn("Row"))
    kept = etree.tostring(rows_before[-1], with_tail=False)

    page_edit.modify_table(
        be, table_page.get("ID"), table.get("objectID"), "delete_rows", indices=[0]
    )
    _, sent = _sent_payload(be)
    sent_rows = next(sent.iter(qn("Table"))).findall(qn("Row"))
    assert len(sent_rows) == len(rows_before) - 1
    assert etree.tostring(sent_rows[-1], with_tail=False) == kept, "untouched row is byte-identical"


def test_modify_table_delete_columns_drops_column_and_each_rows_cell(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n_cols = len(table.findall(f"{qn('Columns')}/{qn('Column')}"))

    page_edit.modify_table(
        be, table_page.get("ID"), table.get("objectID"), "delete_columns", indices=[0]
    )
    _, sent = _sent_payload(be)
    sent_table = next(sent.iter(qn("Table")))
    cols = sent_table.findall(f"{qn('Columns')}/{qn('Column')}")
    assert len(cols) == n_cols - 1
    assert [c.get("index") for c in cols] == [str(i) for i in range(n_cols - 1)], "re-indexed"
    for row in sent_table.findall(qn("Row")):
        assert len(row.findall(qn("Cell"))) == n_cols - 1


def test_modify_table_refuses_deleting_every_row(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n = len(table.findall(qn("Row")))
    with pytest.raises(ValueError, match="every row"):
        page_edit.modify_table(
            be, table_page.get("ID"), table.get("objectID"), "delete_rows", indices=list(range(n))
        )
    assert not [c for c in be.calls if c.method == "update_page_content"]


# --- modify_table set_rows: bulk CONTENT replace, fixed shape, cell identity kept -----


def test_modify_table_set_rows_replaces_content_and_keeps_cell_ids(be, table_page):
    table = next(table_page.iter(qn("Table")))
    cells_before = table.findall(qn("Row"))[0].findall(qn("Cell"))
    n_cols = len(cells_before)
    ids_before = [c.get("objectID") for c in cells_before]
    n_rows_before = len(table.findall(qn("Row")))

    page_edit.modify_table(
        be,
        table_page.get("ID"),
        table.get("objectID"),
        "set_rows",
        rows=[["甲" + str(i) for i in range(n_cols)]],
    )
    _, sent = _sent_payload(be)
    sent_table = next(sent.iter(qn("Table")))
    sent_cells = sent_table.findall(qn("Row"))[0].findall(qn("Cell"))
    # content replaced...
    assert "甲0" in _cdata(sent_cells[0].find(f"{qn('OEChildren')}/{qn('OE')}"))
    # ...shape unchanged, cell identities preserved
    assert len(sent_table.findall(qn("Row"))) == n_rows_before
    assert [c.get("objectID") for c in sent_cells] == ids_before


def test_modify_table_set_rows_single_row_at_index_leaves_other_rows(be, table_page):
    table = next(table_page.iter(qn("Table")))
    rows_before = table.findall(qn("Row"))
    n_cols = len(rows_before[0].findall(qn("Cell")))
    first_row_before = etree.tostring(rows_before[0], with_tail=False)

    page_edit.modify_table(
        be,
        table_page.get("ID"),
        table.get("objectID"),
        "set_rows",
        rows=[["乙"] * n_cols],
        at_index=1,
    )
    _, sent = _sent_payload(be)
    sent_rows = next(sent.iter(qn("Table"))).findall(qn("Row"))
    assert etree.tostring(sent_rows[0], with_tail=False) == first_row_before, "row 0 untouched"
    changed = sent_rows[1].findall(qn("Cell"))[0].find(f"{qn('OEChildren')}/{qn('OE')}")
    assert "乙" in _cdata(changed)


def test_modify_table_set_rows_short_row_leaves_trailing_columns(be, table_page):
    table = next(table_page.iter(qn("Table")))
    last_cell_before = etree.tostring(
        table.findall(qn("Row"))[0].findall(qn("Cell"))[-1], with_tail=False
    )

    page_edit.modify_table(
        be, table_page.get("ID"), table.get("objectID"), "set_rows", rows=[["只改第一格"]]
    )
    _, sent = _sent_payload(be)
    sent_cells = next(sent.iter(qn("Table"))).findall(qn("Row"))[0].findall(qn("Cell"))
    assert "只改第一格" in _cdata(sent_cells[0].find(f"{qn('OEChildren')}/{qn('OE')}"))
    assert etree.tostring(sent_cells[-1], with_tail=False) == last_cell_before, "trailing cell kept"


def test_modify_table_set_rows_rejects_writing_past_last_row(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n_rows = len(table.findall(qn("Row")))
    n_cols = len(table.findall(qn("Row"))[0].findall(qn("Cell")))
    with pytest.raises(ValueError, match="insert_rows"):
        page_edit.modify_table(
            be,
            table_page.get("ID"),
            table.get("objectID"),
            "set_rows",
            rows=[["x"] * n_cols] * (n_rows + 1),  # one row too many
        )
    assert not [c for c in be.calls if c.method == "update_page_content"]


def test_modify_table_set_rows_rejects_rows_wider_than_table(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n_cols = len(table.findall(qn("Row"))[0].findall(qn("Cell")))
    with pytest.raises(ValueError, match="columns"):
        page_edit.modify_table(
            be, table_page.get("ID"), table.get("objectID"), "set_rows", rows=[["x"] * (n_cols + 1)]
        )
    assert not [c for c in be.calls if c.method == "update_page_content"]


def test_modify_table_set_rows_none_cell_leaves_that_cell_unchanged(be, table_page):
    # None = "keep this cell": the "keep first column, clear the rest" pattern is [None, "", ...]
    table = next(table_page.iter(qn("Table")))
    cell0_before = etree.tostring(
        table.findall(qn("Row"))[0].findall(qn("Cell"))[0], with_tail=False
    )
    page_edit.modify_table(
        be, table_page.get("ID"), table.get("objectID"), "set_rows", rows=[[None, "改第二格"]]
    )
    _, sent = _sent_payload(be)
    sent_cells = next(sent.iter(qn("Table"))).findall(qn("Row"))[0].findall(qn("Cell"))
    assert etree.tostring(sent_cells[0], with_tail=False) == cell0_before, "None kept col 0"
    changed = sent_cells[1].find(f"{qn('OEChildren')}/{qn('OE')}")
    assert "改第二格" in _cdata(changed)


# --- modify_table reorder_columns / reorder_rows: rearrange without retyping ----------


def _oe_text(cell: etree._Element) -> str:
    return _cdata(cell.find(f"{qn('OEChildren')}/{qn('OE')}"))


def test_modify_table_reorder_columns_swaps_cells_keeping_ids(be, table_page):
    table = next(table_page.iter(qn("Table")))
    assert len(table.findall(f"{qn('Columns')}/{qn('Column')}")) == 2
    cells_before = table.findall(qn("Row"))[0].findall(qn("Cell"))
    ids_before = [c.get("objectID") for c in cells_before]
    texts_before = [_oe_text(c) for c in cells_before]

    page_edit.modify_table(
        be, table_page.get("ID"), table.get("objectID"), "reorder_columns", order=[1, 0]
    )
    _, sent = _sent_payload(be)
    sent_table = next(sent.iter(qn("Table")))
    cols = sent_table.findall(f"{qn('Columns')}/{qn('Column')}")
    assert [c.get("index") for c in cols] == ["0", "1"], "columns re-indexed after the move"
    sent_cells = sent_table.findall(qn("Row"))[0].findall(qn("Cell"))
    # columns swapped: cell order reversed, each cell keeps its own objectID (no retyping)
    assert [c.get("objectID") for c in sent_cells] == ids_before[::-1]
    assert [_oe_text(c) for c in sent_cells] == texts_before[::-1]
    # every row stays rectangular and re-aligned
    for row in sent_table.findall(qn("Row")):
        assert len(row.findall(qn("Cell"))) == 2


def test_modify_table_reorder_rows_reorders_keeping_row_ids(be, table_page):
    table = next(table_page.iter(qn("Table")))
    rows_before = table.findall(qn("Row"))
    n = len(rows_before)
    ids_before = [r.get("objectID") for r in rows_before]
    order = [1, 0] + list(range(2, n))  # swap the first two rows, keep the rest

    page_edit.modify_table(
        be, table_page.get("ID"), table.get("objectID"), "reorder_rows", order=order
    )
    _, sent = _sent_payload(be)
    sent_table = next(sent.iter(qn("Table")))
    assert sent_table[0].tag == qn("Columns"), "one:Columns stays the first child"
    sent_rows = sent_table.findall(qn("Row"))
    assert [r.get("objectID") for r in sent_rows] == [ids_before[i] for i in order]


def test_modify_table_reorder_rejects_partial_permutation(be, table_page):
    table = next(table_page.iter(qn("Table")))
    with pytest.raises(ValueError, match="permutation"):  # must list EVERY column index once
        page_edit.modify_table(
            be, table_page.get("ID"), table.get("objectID"), "reorder_columns", order=[0]
        )
    assert not [c for c in be.calls if c.method == "update_page_content"]


def test_modify_table_reorder_requires_order(be, table_page):
    table = next(table_page.iter(qn("Table")))
    with pytest.raises(ValueError, match="order"):
        page_edit.modify_table(be, table_page.get("ID"), table.get("objectID"), "reorder_rows")
    assert not [c for c in be.calls if c.method == "update_page_content"]


# --- modify_table set_column: rewrite ONE column compactly ----------------------------


def test_modify_table_set_column_rewrites_one_column_leaving_others(be, table_page):
    table = next(table_page.iter(qn("Table")))
    rows = table.findall(qn("Row"))
    n_rows = len(rows)
    col1_before = [etree.tostring(r.findall(qn("Cell"))[1], with_tail=False) for r in rows]
    ids0_before = [r.findall(qn("Cell"))[0].get("objectID") for r in rows]

    page_edit.modify_table(
        be,
        table_page.get("ID"),
        table.get("objectID"),
        "set_column",
        at_index=0,
        values=[f"第{i}" for i in range(n_rows)],
    )
    _, sent = _sent_payload(be)
    sent_rows = next(sent.iter(qn("Table"))).findall(qn("Row"))
    # column 0 rewritten, cell identities kept...
    assert "第0" in _oe_text(sent_rows[0].findall(qn("Cell"))[0])
    assert "第9" in _oe_text(sent_rows[-1].findall(qn("Cell"))[0])
    assert [r.findall(qn("Cell"))[0].get("objectID") for r in sent_rows] == ids0_before
    # ...column 1 byte-identical
    for r, before in zip(sent_rows, col1_before, strict=True):
        assert etree.tostring(r.findall(qn("Cell"))[1], with_tail=False) == before


def test_modify_table_set_column_short_list_and_none_leave_cells(be, table_page):
    table = next(table_page.iter(qn("Table")))
    rows = table.findall(qn("Row"))
    last0_before = etree.tostring(rows[-1].findall(qn("Cell"))[0], with_tail=False)
    row0c0_before = etree.tostring(rows[0].findall(qn("Cell"))[0], with_tail=False)

    page_edit.modify_table(
        be,
        table_page.get("ID"),
        table.get("objectID"),
        "set_column",
        at_index=0,
        values=[None, "乙"],
    )
    _, sent = _sent_payload(be)
    sent_rows = next(sent.iter(qn("Table"))).findall(qn("Row"))
    assert etree.tostring(sent_rows[0].findall(qn("Cell"))[0], with_tail=False) == row0c0_before
    assert "乙" in _oe_text(sent_rows[1].findall(qn("Cell"))[0])
    # rows past the short list are untouched
    assert etree.tostring(sent_rows[-1].findall(qn("Cell"))[0], with_tail=False) == last0_before


def test_modify_table_set_column_rejects_out_of_range_and_missing_values(be, table_page):
    table = next(table_page.iter(qn("Table")))
    with pytest.raises(ValueError, match="out of range"):
        page_edit.modify_table(
            be, table_page.get("ID"), table.get("objectID"), "set_column", at_index=9, values=["x"]
        )
    with pytest.raises(ValueError, match="values"):
        page_edit.modify_table(
            be, table_page.get("ID"), table.get("objectID"), "set_column", at_index=0
        )
    assert not [c for c in be.calls if c.method == "update_page_content"]


# --- modify_table insert_columns with values: add a column WITH content in one step ------


def test_modify_table_insert_columns_with_values_fills_new_column(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n_rows = len(table.findall(qn("Row")))
    n_cols = len(table.findall(f"{qn('Columns')}/{qn('Column')}"))

    page_edit.modify_table(
        be,
        table_page.get("ID"),
        table.get("objectID"),
        "insert_columns",
        values=[f"備註{i}" for i in range(n_rows)],
    )
    _, sent = _sent_payload(be)
    sent_table = next(sent.iter(qn("Table")))
    assert len(sent_table.findall(f"{qn('Columns')}/{qn('Column')}")) == n_cols + 1
    sent_rows = sent_table.findall(qn("Row"))
    assert "備註0" in _oe_text(sent_rows[0].findall(qn("Cell"))[-1])  # the new last column
    assert "備註9" in _oe_text(sent_rows[-1].findall(qn("Cell"))[-1])


def test_modify_table_insert_columns_values_rejects_multiple_columns(be, table_page):
    table = next(table_page.iter(qn("Table")))
    with pytest.raises(ValueError, match="SINGLE column"):
        page_edit.modify_table(
            be, table_page.get("ID"), table.get("objectID"), "insert_columns", count=2, values=["a"]
        )
    assert not [c for c in be.calls if c.method == "update_page_content"]


# --- nested tables (a table inside a cell — the 業務塔斯作業 PAYMENT layout) --------------

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _nested_table_tree():
    def cell(cid, pid, text):
        return (
            f'<one:Cell objectID="{cid}"><one:OEChildren>'
            f'<one:OE objectID="{pid}"><one:T><![CDATA[{text}]]></one:T></one:OE>'
            "</one:OEChildren></one:Cell>"
        )

    inner = (
        '<one:Table objectID="T-INNER">'
        '<one:Columns><one:Column index="0" width="50"/><one:Column index="1" width="150"/>'
        "</one:Columns>"
        f'<one:Row objectID="IR1">{cell("IC1", "IP1", "全額")}'
        f"{cell('IC2', 'IP2', '91900')}</one:Row>"
        f'<one:Row objectID="IR2">{cell("IC3", "IP3", "已付")}'
        f"{cell('IC4', 'IP4', '23000')}</one:Row>"
        "</one:Table>"
    )
    xml = (
        f'<one:Page xmlns:one="{_ONE_NS}" ID="P"><one:Outline objectID="OUT"><one:OEChildren>'
        '<one:OE objectID="OE-OUTER"><one:Table objectID="T-OUTER">'
        '<one:Columns><one:Column index="0" width="100"/><one:Column index="1" width="300"/>'
        "</one:Columns>"
        f'<one:Row objectID="R1">{cell("C1", "P1", "PAYMENT")}'
        f'<one:Cell objectID="C2"><one:OEChildren><one:OE objectID="P2">{inner}'
        "</one:OE></one:OEChildren></one:Cell></one:Row>"
        "</one:Table></one:OE></one:OEChildren></one:Outline></one:Page>"
    )
    return etree.fromstring(xml.encode("utf-8"), parser=_PARSER)


def test_parse_surfaces_nested_table_with_its_own_object_id():
    from onenote_com_mcp.xmllayer.parse import parse_page

    page = parse_page(etree.tostring(_nested_table_tree()).decode("utf-8"))
    outer = page.outlines[0].paragraphs[0].table
    assert outer.object_id == "T-OUTER"
    inner = outer.rows[0][1].paragraphs[0].table  # the PAYMENT value cell holds a table
    assert inner is not None and inner.object_id == "T-INNER"
    assert [c.text for c in inner.rows[0]] == ["全額", "91900"]


def test_modify_table_on_inner_table_does_not_touch_outer():
    tree = _nested_table_tree()
    inner = page_edit._find_content_object(tree, "T-INNER")
    outer = page_edit._find_content_object(tree, "T-OUTER")

    page_edit._insert_table_columns(inner, None, 1, None)  # add a column to the INNER table

    assert len(inner.findall(f"{qn('Columns')}/{qn('Column')}")) == 3
    for row in inner.findall(qn("Row")):
        assert len(row.findall(qn("Cell"))) == 3
    # the OUTER table is untouched: still 2 columns, its row still has 2 cells
    assert len(outer.findall(f"{qn('Columns')}/{qn('Column')}")) == 2
    assert len(outer.findall(qn("Row"))[0].findall(qn("Cell"))) == 2


def test_modify_table_delete_row_on_inner_table_scoped():
    tree = _nested_table_tree()
    inner = page_edit._find_content_object(tree, "T-INNER")
    outer = page_edit._find_content_object(tree, "T-OUTER")

    page_edit._delete_table_rows(inner, [0])  # drop the inner table's first row

    assert len(inner.findall(qn("Row"))) == 1
    assert len(outer.findall(qn("Row"))) == 1, "outer table's single row is untouched"


def test_modify_table_rejects_non_table_target(be, mixed):
    # a paragraph OE objectID is not a table
    oe = next(o for o in mixed.iter(qn("OE")) if o.get("objectID"))
    with pytest.raises(ValueError, match="not a table"):
        page_edit.modify_table(be, mixed.get("ID"), oe.get("objectID"), "delete_rows", indices=[0])


# --- delete_inline_content: remove a table / paragraph from INSIDE an outline ---------


def _inject(tmp_path, page_id: str, body: str) -> FixtureBackend:
    """Write a one:Page fixture (body = the page's child XML) and return a FixtureBackend on it."""
    from onenote_com_mcp.backend.fixture import _sanitize

    (tmp_path / f"page_{_sanitize(page_id)}.xml").write_text(
        '<?xml version="1.0"?>'
        f'<one:Page xmlns:one="{_ONE_NS}" ID="{page_id}" '
        f'lastModifiedTime="2026-06-13T00:00:00.000Z">{body}</one:Page>',
        encoding="utf-8",
    )
    return FixtureBackend(tmp_path)


def test_delete_inline_content_removes_table_but_keeps_sibling_paragraphs(tmp_path):
    page_id = "{P}{1}{D0}"
    body = (
        '<one:Outline objectID="OUT"><one:OEChildren>'
        '<one:OE objectID="PARA-A"><one:T><![CDATA[保留我]]></one:T></one:OE>'
        '<one:OE objectID="OE-TBL"><one:Table objectID="TBL">'
        '<one:Columns><one:Column index="0" width="100"/></one:Columns>'
        '<one:Row><one:Cell objectID="C1"><one:OEChildren>'
        '<one:OE objectID="CP1"><one:T><![CDATA[格子]]></one:T></one:OE>'
        "</one:OEChildren></one:Cell></one:Row>"
        "</one:Table></one:OE>"
        "</one:OEChildren></one:Outline>"
    )
    be = _inject(tmp_path, page_id, body)
    page_edit.delete_inline_content(be, page_id, "TBL")
    _, sent = _sent_payload(be)
    assert sent.find(f".//{qn('Table')}") is None, "the table is gone"
    survivors = [o.get("objectID") for o in sent.iter(qn("OE")) if o.get("objectID")]
    assert "PARA-A" in survivors, "the sibling paragraph is kept"
    assert "OE-TBL" not in survivors, "the table's now-empty wrapping OE is pruned"


def test_delete_inline_content_removes_a_paragraph(tmp_path):
    page_id = "{P}{1}{D1}"
    body = (
        '<one:Outline objectID="OUT"><one:OEChildren>'
        '<one:OE objectID="KEEP"><one:T><![CDATA[留]]></one:T></one:OE>'
        '<one:OE objectID="DROP"><one:T><![CDATA[刪]]></one:T></one:OE>'
        "</one:OEChildren></one:Outline>"
    )
    be = _inject(tmp_path, page_id, body)
    page_edit.delete_inline_content(be, page_id, "DROP")
    _, sent = _sent_payload(be)
    survivors = [o.get("objectID") for o in sent.iter(qn("OE")) if o.get("objectID")]
    assert survivors == ["KEEP"]


def test_delete_inline_content_keeps_table_cell_valid(tmp_path):
    """A nested table lives in a cell. Deleting it must leave the cell valid (OEChildren > OE),
    never an empty <one:Cell/> (rejected by COM with hrInvalidXML)."""
    page_id = "{P}{1}{D2}"
    inner = (
        '<one:Table objectID="T-INNER">'
        '<one:Columns><one:Column index="0" width="50"/></one:Columns>'
        '<one:Row><one:Cell objectID="IC1"><one:OEChildren>'
        '<one:OE objectID="IP1"><one:T><![CDATA[x]]></one:T></one:OE>'
        "</one:OEChildren></one:Cell></one:Row></one:Table>"
    )
    body = (
        '<one:Outline objectID="OUT"><one:OEChildren><one:OE objectID="OE-OUTER">'
        '<one:Table objectID="T-OUTER">'
        '<one:Columns><one:Column index="0" width="100"/></one:Columns>'
        f'<one:Row><one:Cell objectID="C2"><one:OEChildren><one:OE objectID="P2">{inner}'
        "</one:OE></one:OEChildren></one:Cell></one:Row>"
        "</one:Table></one:OE></one:OEChildren></one:Outline>"
    )
    be = _inject(tmp_path, page_id, body)
    page_edit.delete_inline_content(be, page_id, "T-INNER")
    _, sent = _sent_payload(be)
    inner_tables = [t for t in sent.iter(qn("Table")) if t.get("objectID") == "T-INNER"]
    assert not inner_tables, "the inner table is gone"
    cell = next(c for c in sent.iter(qn("Cell")) if c.get("objectID") == "C2")
    assert cell.find(f"{qn('OEChildren')}/{qn('OE')}") is not None, "cell stays valid, not empty"


def test_delete_inline_content_rejects_page_level_object(be, mixed):
    # the page-level outline is delete_page_content's job, not this tool's
    outline = mixed.find(qn("Outline"))
    with pytest.raises(ValueError, match="delete_page_content"):
        page_edit.delete_inline_content(be, mixed.get("ID"), outline.get("objectID"))
    assert not [c for c in be.calls if c.method == "update_page_content"], "no write on refusal"


# --- images ---------------------------------------------------------------------------


def test_insert_svg_image_appends_oe_wrapped_rasterized_png(be, mixed):
    import base64

    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20">'
        '<rect width="20" height="20" fill="#3366cc"/></svg>'
    )
    page_edit.insert_svg_image(be, mixed.get("ID"), svg, width=100.0, height=50.0)
    _, sent = _sent_payload(be)
    last_oe = _oes(sent.findall(qn("Outline"))[-1])[-1]
    image = last_oe.find(qn("Image"))
    assert image is not None, "the image rides in its own one:OE (the deletable objectID)"
    raw = base64.b64decode(image.find(qn("Data")).text)
    assert raw.startswith(b"\x89PNG"), "the SVG was rasterized to a PNG and inlined as one:Data"
    assert image.find(qn("CallbackID")) is None
    assert image.find(qn("Size")).get("width") == "100.0"


def test_insert_svg_image_rejects_embedded_raster_before_any_write(be, mixed):
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg">'
        '<image href="data:image/png;base64,iVBORw0KGgo="/></svg>'
    )
    with pytest.raises(ValueError, match="vector-only"):
        page_edit.insert_svg_image(be, mixed.get("ID"), svg)
    assert not [c for c in be.calls if c.method == "update_page_content"], "rejected before write"


_SVG_DOT = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20">'
    '<rect width="20" height="20" fill="#3366cc"/></svg>'
)


def test_insert_svg_image_positions_relative_to_a_paragraph(be, mixed):
    page_id = mixed.get("ID")
    outline = mixed.findall(qn("Outline"))[-1]
    target_id = _oes(outline)[1].get("objectID")

    page_edit.insert_svg_image(
        be, page_id, _SVG_DOT, mode="insert_after", target_object_id=target_id
    )
    _, sent = _sent_payload(be)
    oes = _oes(sent.findall(qn("Outline"))[-1])
    idx = next(i for i, oe in enumerate(oes) if oe.get("objectID") == target_id)
    assert oes[idx + 1].find(qn("Image")) is not None, (
        "image landed right AFTER the target paragraph"
    )

    be2 = FixtureBackend(be.fixtures_dir)
    page_edit.insert_svg_image(
        be2, page_id, _SVG_DOT, mode="insert_before", target_object_id=target_id
    )
    _, sent2 = _sent_payload(be2)
    oes2 = _oes(sent2.findall(qn("Outline"))[-1])
    idx2 = next(i for i, oe in enumerate(oes2) if oe.get("objectID") == target_id)
    assert oes2[idx2 - 1].find(qn("Image")) is not None, "image landed right BEFORE the target"


def test_insert_svg_image_mode_validation(be, mixed):
    page_id = mixed.get("ID")
    with pytest.raises(ValueError, match="mode must be one of"):
        page_edit.insert_svg_image(be, page_id, _SVG_DOT, mode="replace")
    # insert_before/after need a paragraph anchor
    with pytest.raises(ValueError, match="target_object_id"):
        page_edit.insert_svg_image(be, page_id, _SVG_DOT, mode="insert_after")
    assert not [c for c in be.calls if c.method == "update_page_content"]


def test_editing_outline_with_existing_image_inlines_its_binary(be, image_page, fixtures_dir):
    """The touched outline contains an untouched image: its pixels must be inlined (CallbackID
    is read-side only) so the merge cannot strip them."""
    page_id = image_page.get("ID")
    callback_id = next(image_page.iter(qn("CallbackID"))).get("callbackID")
    from onenote_com_mcp.backend.fixture import _sanitize

    expected_b64 = (fixtures_dir / f"binary_{_sanitize(callback_id)}.b64").read_text("utf-8")
    image_outline = next(  # the outline holding the image (NOT the page's last outline)
        o for o in image_page.findall(qn("Outline")) if o.find(f".//{qn('Image')}") is not None
    )

    page_edit.edit_page_content(
        be, page_id, "在圖片頁加字", target_object_id=image_outline.get("objectID")
    )
    _, sent = _sent_payload(be)
    assert len(sent.findall(qn("Outline"))) == 1, "the OTHER outline stays pruned"
    image = next(sent.iter(qn("Image")))
    assert image.find(qn("CallbackID")) is None
    assert image.find(qn("Data")).text == expected_b64


# --- payload strategy seam -------------------------------------------------------------


def test_whole_page_strategy_sends_everything_and_reads_binary(be, image_page):
    def mutate(tree):
        tree.set("marker", "edited")

    page_edit.apply_page_edit(be, image_page.get("ID"), mutate, strategy="whole_page")
    _, sent = _sent_payload(be)
    assert sent.find(qn("Title")) is not None, "whole_page keeps untouched objects"
    assert sent.findall(qn("Outline")), "whole_page keeps untouched outlines"
    image = next(sent.iter(qn("Image")))
    assert image.find(qn("Data")) is not None, "image binary must be inline"
    assert image.find(qn("CallbackID")) is None


# --- Phase 5b: attachments must not ride into edit payloads ------------------------------


@pytest.fixture
def attachment_page(fixtures_dir) -> etree._Element:
    return _page_root(fixtures_dir, "附件與嵌入物件-1")


def test_edit_prunes_untouched_inserted_files_and_xps_carriers(be, attachment_page):
    """changed_objects payloads must drop unchanged page-level InsertedFiles AND the printout's
    one:XPSFile carrier (a read-side construct holding a CallbackID) AND the unchanged
    page-level render Image — only the touched outline goes back."""
    page_id = attachment_page.get("ID")
    assert attachment_page.findall(qn("InsertedFile")), "fixture must have page-level files"
    assert attachment_page.findall(qn("XPSFile")), "fixture must have an XPSFile carrier"

    page_edit.edit_page_content(be, page_id, "附註一行")

    kwargs, sent = _sent_payload(be)
    assert sent.findall(qn("InsertedFile")) == []  # page-level ones pruned
    assert sent.findall(qn("XPSFile")) == []
    assert sent.findall(qn("Image")) == []  # the unchanged printout render too
    outlines = sent.findall(qn("Outline"))
    assert len(outlines) == 1
    # nothing image-shaped left to inline → no one:Data in the payload at all
    assert sent.findall(f".//{qn('Data')}") == []
    # INLINE attachments live inside untouched outlines, which are pruned whole; the one
    # touched outline carries no InsertedFile — so none appear above. (Inline protection
    # comes from outline pruning, ground truth 2026-06-12.)
    assert "附註一行" in kwargs["changes_xml"]


# --- contract errors --------------------------------------------------------------------


def test_input_contract_errors(be, mixed):
    page_id = mixed.get("ID")
    oe_id = _oes(mixed.findall(qn("Outline"))[-1])[0].get("objectID")
    with pytest.raises(ValueError, match="mode"):
        page_edit.edit_page_content(be, page_id, "x", "prepend")
    with pytest.raises(ValueError, match="target_object_id"):
        page_edit.edit_page_content(be, page_id, "x", "insert_after")
    with pytest.raises(NodeNotFoundError):
        page_edit.edit_page_content(be, page_id, "x", "replace", target_object_id="{NOPE}{1}{B0}")
    with pytest.raises(ValueError, match="not a one:Outline"):
        page_edit.edit_page_content(be, page_id, "x", "append", target_object_id=oe_id)
    with pytest.raises(ValueError, match="rows is empty"):
        page_edit.add_table(be, page_id, [])
    with pytest.raises(ValueError, match="non-empty"):
        page_edit.edit_page_content(be, page_id, [])
    assert not [c for c in be.calls if c.method == "update_page_content"]


# --- apply_text_style (batch font / size / color patch) -------------------------------------


def _all_runs(page):
    """Every Run on a parsed page — recursing nested OEs and into table cells."""
    out = []

    def walk(paras):
        for p in paras:
            out.extend(p.runs)
            if p.table:
                for row in p.table.rows:
                    for cell in row:
                        walk(cell.paragraphs)
            walk(p.children)

    walk(page.paragraphs)
    return out


def test_apply_text_style_whole_page_sets_every_run_and_quickstyledef(be, mixed):
    from onenote_com_mcp.xmllayer.parse import _parse_quick_styles, parse_page

    page_id = mixed.get("ID")
    summary = page_edit.apply_text_style(be, page_id, font_family="微軟正黑體")
    kwargs, sent = _sent_payload(be)

    # EVERY visible run is now 微軟正黑體 (span wins the cascade) — the real guarantee
    sent_page = parse_page(kwargs["changes_xml"])
    fonts = {r.style.get("font-family") for r in _all_runs(sent_page)}
    assert fonts == {"微軟正黑體"}, fonts
    # whole-page also rewrote the QuickStyleDef baseline (was Calibri)
    qs = _parse_quick_styles(sent)
    assert qs and all(d.font == "微軟正黑體" for d in qs.values())
    assert summary["quick_styles_updated"] == len(qs)
    assert summary["runs_changed"] >= 1 and summary["scope"] == "page"
    # the page title is left alone → unchanged → pruned from the payload
    assert sent.find(qn("Title")) is None


def test_apply_text_style_preserves_bold_color_highlight_and_link(be):
    # the 混合樣式頁 carries bold / a custom color / dual-highlight / (table page) hyperlinks
    page_id = _page_root(be.fixtures_dir, "混合樣式頁").get("ID")
    page_edit.apply_text_style(be, page_id, font_family="微軟正黑體")
    xml = _sent_payload(be)[0]["changes_xml"]
    assert "font-weight:bold" in xml  # bold survived
    assert "color:#FA0000" in xml  # the custom text colour survived
    assert "background:yellow" in xml and "mso-highlight:yellow" in xml  # dual highlight survived
    assert "font-family:微軟正黑體" in xml  # the new font is in


def test_apply_text_style_sub_scope_is_byte_isolated_and_skips_quickstyledef(be, mixed):
    from onenote_com_mcp.xmllayer.parse import _parse_quick_styles

    page_id = mixed.get("ID")
    outline = mixed.findall(qn("Outline"))[-1]
    target_id = _oes(outline)[0].get("objectID")
    before_by_id = {oe.get("objectID"): etree.tostring(oe, with_tail=False) for oe in _oes(outline)}

    summary = page_edit.apply_text_style(
        be, page_id, font_family="標楷體", scope_object_id=target_id
    )
    _, sent = _sent_payload(be)

    for oe in _oes(sent.findall(qn("Outline"))[-1]):
        oid = oe.get("objectID")
        if oid == target_id:
            assert "標楷體" in etree.tostring(oe, with_tail=False, encoding="unicode")
        else:  # every sibling paragraph is byte-identical
            assert etree.tostring(oe, with_tail=False) == before_by_id[oid]
    # a sub-scope must NOT touch the page-global QuickStyleDef
    assert all(d.font == "Calibri" for d in _parse_quick_styles(sent).values())
    assert summary["quick_styles_updated"] == 0


def test_apply_text_style_size_and_color_only_keep_existing_font(be, mixed):
    from onenote_com_mcp.xmllayer.parse import parse_page

    page_id = mixed.get("ID")
    page_edit.apply_text_style(be, page_id, size=18, color="#0000FF")
    kwargs, _ = _sent_payload(be)
    assert "font-size:18.0pt" in kwargs["changes_xml"]
    sent_page = parse_page(kwargs["changes_xml"])
    runs = _all_runs(sent_page)
    assert all(r.style.get("font-size") == "18.0pt" for r in runs)
    assert all(r.style.get("color") == "#0000FF" for r in runs)
    # font-family was NOT requested → the page's original mix of fonts is preserved
    assert {r.style.get("font-family") for r in runs} != {None}


def test_apply_text_style_requires_a_property_and_validates_scope(be, mixed):
    page_id = mixed.get("ID")
    with pytest.raises(ValueError, match="at least one"):
        page_edit.apply_text_style(be, page_id)
    with pytest.raises(NodeNotFoundError):
        page_edit.apply_text_style(be, page_id, font_family="X", scope_object_id="{NOPE}{1}{B0}")
    assert not [c for c in be.calls if c.method == "update_page_content"]


def test_apply_text_style_bold_italic_whole_page(be, mixed):
    from onenote_com_mcp.xmllayer.parse import _parse_quick_styles, parse_page

    page_id = mixed.get("ID")
    page_edit.apply_text_style(be, page_id, bold=True, italic=True)
    kwargs, sent = _sent_payload(be)
    runs = _all_runs(parse_page(kwargs["changes_xml"]))
    assert runs and all(r.style.get("font-weight") == "bold" for r in runs)
    assert all(r.style.get("font-style") == "italic" for r in runs)
    # whole-page also flips the QuickStyleDef baseline bold/italic
    qs = _parse_quick_styles(sent)
    assert qs and all(d.bold and d.italic for d in qs.values())


def test_apply_text_style_bold_false_turns_emphasis_off(be, mixed):
    from onenote_com_mcp.xmllayer.parse import parse_page

    page_id = mixed.get("ID")  # the page has a bold run (粗粗粗)
    page_edit.apply_text_style(be, page_id, bold=False)
    runs = _all_runs(parse_page(_sent_payload(be)[0]["changes_xml"]))
    assert all(r.style.get("font-weight") in (None, "normal") for r in runs)
    formerly_bold = [r for r in runs if "粗" in r.text]
    assert formerly_bold and all(r.style.get("font-weight") == "normal" for r in formerly_bold)


def test_apply_text_style_underline_merges_with_existing_strikethrough(be, mixed):
    # the page has a line-through run (刪刪刪); adding underline must keep BOTH decorations
    page_id = mixed.get("ID")
    page_edit.apply_text_style(be, page_id, underline=True)
    xml = _sent_payload(be)[0]["changes_xml"]
    assert "text-decoration:line-through underline" in xml  # both, merged (sorted), not clobbered


def test_apply_text_style_strikethrough_only_does_not_touch_font(be, mixed):
    from onenote_com_mcp.xmllayer.parse import parse_page

    page_id = mixed.get("ID")
    page_edit.apply_text_style(be, page_id, strikethrough=True)
    runs = _all_runs(parse_page(_sent_payload(be)[0]["changes_xml"]))
    assert runs and all("line-through" in (r.style.get("text-decoration") or "") for r in runs)
    # font-family was not requested → the original mix is preserved
    assert {r.style.get("font-family") for r in runs} != {None}


def test_apply_text_style_highlight_sets_dual_background(be, mixed):
    from onenote_com_mcp.xmllayer.parse import _parse_quick_styles, parse_page

    page_id = mixed.get("ID")
    page_edit.apply_text_style(be, page_id, highlight="yellow")
    kwargs, sent = _sent_payload(be)
    xml = kwargs["changes_xml"]
    assert "background:yellow" in xml and "mso-highlight:yellow" in xml  # OneNote's dual highlight
    runs = _all_runs(parse_page(xml))
    assert runs and all(r.style.get("background") == "yellow" for r in runs)
    # highlight is span-only — the QuickStyleDef baseline is NOT touched (OneNote rejects a
    # highlightColor set to a CSS color name; VM-confirmed). font baseline still rewrites, though.
    assert all(d.highlight_color is None for d in _parse_quick_styles(sent).values())


def test_apply_text_style_highlight_none_clears_existing(be, mixed):
    from onenote_com_mcp.xmllayer.parse import parse_page

    page_id = mixed.get("ID")  # the page has a highlighted run (螢光標示文字, background:yellow)
    page_edit.apply_text_style(be, page_id, highlight="none")
    runs = _all_runs(parse_page(_sent_payload(be)[0]["changes_xml"]))
    assert runs and all(not r.style.get("background") for r in runs), "all highlights removed"


def _one_run_page(tmp_path, page_id, cdata):
    from onenote_com_mcp.backend.fixture import _sanitize

    (tmp_path / f"page_{_sanitize(page_id)}.xml").write_text(
        '<?xml version="1.0"?><one:Page '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{page_id}" lastModifiedTime="2026-06-10T17:39:30.000Z">'
        '<one:Outline><one:OEChildren><one:OE objectID="{OE}{1}{B0}"><one:T>'
        f"<![CDATA[{cdata}]]>"
        "</one:T></one:OE></one:OEChildren></one:Outline></one:Page>",
        encoding="utf-8",
    )


def test_apply_text_style_highlight_none_clears_a_highlight_only_run(tmp_path):
    # REGRESSION: a run whose ONLY span style is the highlight (the table-cell case Chris hit).
    # Popping background left span_style EMPTY, and build_spans' `span_style or style` fallback
    # resurrected the highlight from run.style — so "most" highlights survived. run.style is now
    # kept in sync.
    from onenote_com_mcp.backend.fixture import FixtureBackend

    page_id = "{P}{1}{B0}"
    _one_run_page(tmp_path, page_id, "<span style='background:yellow'>整格黃</span>")
    be = FixtureBackend(tmp_path)
    page_edit.apply_text_style(be, page_id, highlight="none")
    xml = _sent_payload(be)[0]["changes_xml"]
    assert "background" not in xml and "mso-highlight" not in xml, "highlight-only run cleared"


def test_apply_text_style_cell_shading_sets_and_clears(be, table_page, tmp_path):
    from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize

    # SET: every cell on the real table page gets the shadingColor
    page_id = table_page.get("ID")
    summary = page_edit.apply_text_style(be, page_id, cell_shading="yellow")
    _, sent = _sent_payload(be)
    cells = sent.findall(f".//{qn('Cell')}")
    # cell_shading is normalized to hex (OneNote's shadingColor attribute rejects a color NAME)
    assert cells and all(c.get("shadingColor") == "#FFFF00" for c in cells)
    assert summary["cells_changed"] == len(cells)

    # CLEAR: a cell that already has shadingColor → the attribute is REMOVED (only clean clear)
    shaded_id = "{SP}{1}{B0}"
    (tmp_path / f"page_{_sanitize(shaded_id)}.xml").write_text(
        '<?xml version="1.0"?><one:Page '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{shaded_id}" lastModifiedTime="2026-06-10T17:39:30.000Z">'
        "<one:Outline><one:OEChildren>"
        '<one:OE><one:Table><one:Columns><one:Column index="0" width="100"/></one:Columns>'
        '<one:Row><one:Cell objectID="{C}{1}{B0}" shadingColor="yellow"><one:OEChildren>'
        "<one:OE><one:T><![CDATA[x]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row>"
        "</one:Table></one:OE></one:OEChildren></one:Outline></one:Page>",
        encoding="utf-8",
    )
    be2 = FixtureBackend(tmp_path)
    page_edit.apply_text_style(be2, shaded_id, cell_shading="none")
    xml = _sent_payload(be2)[0]["changes_xml"]
    assert "shadingColor" not in xml, "cell shading attribute removed"


def test_apply_text_style_cell_shading_only_does_not_rewrite_text(be, table_page):
    # a cell-shading-only call must leave the table's text/runs untouched (no T rewrite)
    page_id = table_page.get("ID")
    summary = page_edit.apply_text_style(be, page_id, cell_shading="yellow")
    assert summary["runs_changed"] == 0 and summary["text_blocks_changed"] == 0
    assert summary["cells_changed"] >= 1


def _cell_text(cell: etree._Element) -> str:
    return "".join(t.text or "" for t in cell.iter(qn("T")))


def test_apply_text_style_columns_restricts_text_and_shading_to_that_column(be, table_page):
    # columns=[0] → only the FIRST cell of every row is touched (BOTH text restyle and shading);
    # column 1 is left alone. A column has no objectID, so this is the only way to address it.
    page_id = table_page.get("ID")
    summary = page_edit.apply_text_style(
        be, page_id, color="#123456", cell_shading="cyan", columns=[0]
    )
    _, sent = _sent_payload(be)
    rows = sent.find(f".//{qn('Table')}").findall(qn("Row"))
    assert len(rows) >= 2
    col0 = [r.findall(qn("Cell"))[0] for r in rows]
    col1 = [r.findall(qn("Cell"))[1] for r in rows if len(r.findall(qn("Cell"))) > 1]
    # shading: every column-0 cell set (cyan→hex); no column-1 cell turned cyan
    assert all(c.get("shadingColor") == "#00FFFF" for c in col0)
    assert all(c.get("shadingColor") != "#00FFFF" for c in col1)
    assert summary["cells_changed"] == len(col0)
    # text restyle reached column 0 but not column 1
    assert any("#123456" in _cell_text(c) for c in col0)
    assert all("#123456" not in _cell_text(c) for c in col1)


def test_apply_text_style_row_scope_restyles_only_that_row(be, table_page):
    # a whole ROW = its one:Row objectID as scope (now exposed via get_page's row_object_ids);
    # the text loop (scope.iter(T)) and shading loop (scope.iter(Cell)) both fall inside that row.
    page_id = table_page.get("ID")
    rows = table_page.find(f".//{qn('Table')}").findall(qn("Row"))
    row_id = rows[1].get("objectID")
    assert row_id  # fixture rows carry objectIDs
    page_edit.apply_text_style(
        be, page_id, color="#123456", cell_shading="lime", scope_object_id=row_id
    )
    _, sent = _sent_payload(be)
    for r in sent.find(f".//{qn('Table')}").findall(qn("Row")):
        cells = r.findall(qn("Cell"))
        if r.get("objectID") == row_id:
            assert all(c.get("shadingColor") == "#00FF00" for c in cells)  # lime→hex
            assert any("#123456" in _cell_text(c) for c in cells)
        else:
            assert all(c.get("shadingColor") != "#00FF00" for c in cells)
            assert all("#123456" not in _cell_text(c) for c in cells)


def test_apply_text_style_columns_rejects_negative(be, table_page):
    with pytest.raises(ValueError, match="0-indexed"):
        page_edit.apply_text_style(be, table_page.get("ID"), cell_shading="red", columns=[-1])


def test_apply_text_style_cell_shading_rejects_unknown_color_name(be, table_page):
    # OneNote's shadingColor needs hex; an unknown color name fails fast (before COM)
    with pytest.raises(ValueError, match="hex"):
        page_edit.apply_text_style(be, table_page.get("ID"), cell_shading="notacolor")
