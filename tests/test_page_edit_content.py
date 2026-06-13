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


def test_modify_table_add_columns_appends_column_to_every_row(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n_cols = len(table.findall(f"{qn('Columns')}/{qn('Column')}"))
    n_rows = len(table.findall(qn("Row")))

    page_edit.modify_table(be, table_page.get("ID"), table.get("objectID"), "add_columns")
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
    assert n_rows == len(sent_table.findall(qn("Row")))  # add_columns doesn't change row count


def test_modify_table_add_columns_at_position(be, table_page):
    table = next(table_page.iter(qn("Table")))
    n_cols = len(table.findall(f"{qn('Columns')}/{qn('Column')}"))
    page_edit.modify_table(
        be, table_page.get("ID"), table.get("objectID"), "add_columns", at_index=0, count=2
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


def test_modify_table_rejects_non_table_target(be, mixed):
    # a paragraph OE objectID is not a table
    oe = next(o for o in mixed.iter(qn("OE")) if o.get("objectID"))
    with pytest.raises(ValueError, match="not a table"):
        page_edit.modify_table(be, mixed.get("ID"), oe.get("objectID"), "delete_rows", indices=[0])


# --- images ---------------------------------------------------------------------------


def test_insert_image_appends_oe_wrapped_image_with_inline_data(be, mixed):
    page_edit.insert_image(be, mixed.get("ID"), "QUJD", "image/png", width=100.0, height=50.0)
    _, sent = _sent_payload(be)
    last_oe = _oes(sent.findall(qn("Outline"))[-1])[-1]
    image = last_oe.find(qn("Image"))
    assert image is not None, "the image rides in its own one:OE (the deletable objectID)"
    assert image.find(qn("Data")).text == "QUJD"
    assert image.find(qn("CallbackID")) is None
    assert image.find(qn("Size")).get("width") == "100.0"


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
