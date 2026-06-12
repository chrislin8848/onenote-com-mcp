"""Tier-1 tests for the Phase-5 Stage-1 copy core (transfer_page), on REAL fixtures.

Fidelity is asserted at the byte level wherever the source bytes should survive verbatim:
CDATA spans, QuickStyleDef tables, table structure. Identity must NOT survive: objectIDs,
stamps, view state are stripped so OneNote mints fresh ones; the page ID becomes the blank
target page's; pageLevel travels through the whole-batch hierarchy seam.
"""

from __future__ import annotations

import pytest
from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize
from onenote_com_mcp.enums import NewPageStyle
from onenote_com_mcp.service import copy
from onenote_com_mcp.xmllayer.namespaces import qn

_PARSER = etree.XMLParser(strip_cdata=False)
_NEW_PAGE_ID = "{FIXTURE-page-1}{1}{B0}"  # FixtureBackend's first create_new_page result
_BLANK_STAMP = "2026-06-11T09:00:00.000Z"
_TARGET_SECTION = "{TARGET}{1}{B0}"


@pytest.fixture
def be(fixtures_dir, tmp_path) -> FixtureBackend:
    """Real page + binary fixtures, plus the blank page the transplant will be written onto."""
    for src in list(fixtures_dir.glob("page_*.xml")) + list(fixtures_dir.glob("binary_*.b64")):
        (tmp_path / src.name).write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
    (tmp_path / f"page_{_sanitize(_NEW_PAGE_ID)}.xml").write_text(
        '<?xml version="1.0"?><one:Page '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{_NEW_PAGE_ID}" lastModifiedTime="{_BLANK_STAMP}"/>',
        encoding="utf-8",
    )
    return FixtureBackend(tmp_path)


def _page_by_name(fixtures_dir, page_name: str) -> tuple[str, etree._Element]:
    """(page_id, binary-variant root) for a named real fixture page."""
    for path in sorted(fixtures_dir.glob("page_*__binary.xml")):
        root = etree.fromstring(path.read_bytes(), parser=_PARSER)
        if root.get("name") == page_name:
            return root.get("ID"), root
    raise AssertionError(f"no binary fixture page named {page_name!r}")


def _sent_payload(be: FixtureBackend) -> tuple[dict, etree._Element]:
    writes = [c for c in be.calls if c.method == "update_page_content"]
    assert len(writes) == 1, "the transplant must be exactly ONE UpdatePageContent"
    return writes[0].kwargs, etree.fromstring(
        writes[0].kwargs["changes_xml"].encode("utf-8"), parser=_PARSER
    )


def test_transfer_mixed_page_is_byte_faithful_with_fresh_identity(be, fixtures_dir):
    source_id, source = _page_by_name(fixtures_dir, "混合樣式頁")

    result = copy.transfer_page(be, source_id, _TARGET_SECTION)

    assert result.page_id == _NEW_PAGE_ID
    assert result.file_notes == []  # no attachments on this page — fully faithful
    # source read is raw piBinaryData; the blank page is created title-less in the target
    first_read = next(c for c in be.calls if c.method == "create_new_page")
    assert first_read.kwargs == {
        "section_id": _TARGET_SECTION,
        "style": NewPageStyle.npsBlankPageNoTitle,
    }
    kwargs, sent = _sent_payload(be)
    # identity is FRESH: root ID is the blank page's, no objectID/stamps anywhere
    assert sent.get("ID") == _NEW_PAGE_ID
    assert not [el for el in sent.iter() if el.get("objectID")], "objectIDs must be reset"
    assert not [el for el in sent.iter() if el.get("lastModifiedTime")]
    assert sent.get("pageLevel") is None, "pageLevel rides the hierarchy, not the page XML"
    # the concurrency guard carries the BLANK page's stamp, not the source's
    assert kwargs["expected_last_modified"].isoformat() == "2026-06-11T09:00:00+00:00"
    # fidelity: QuickStyleDef table carried whole; CDATA spans byte-identical
    assert len(sent.findall(qn("QuickStyleDef"))) == len(source.findall(qn("QuickStyleDef")))
    assert ">粗粗粗</span>" in kwargs["changes_xml"]
    assert "<![CDATA[" in kwargs["changes_xml"]
    title = sent.find(f"{qn('Title')}/{qn('OE')}/{qn('T')}")
    assert title is not None and "混合樣式頁" in (title.text or "")


def test_transfer_image_page_inlines_binary_from_source_callbacks(be, fixtures_dir):
    source_id, source = _page_by_name(fixtures_dir, "圖片頁")
    callback_id = next(source.iter(qn("CallbackID"))).get("callbackID")
    expected_b64 = (fixtures_dir / f"binary_{_sanitize(callback_id)}.b64").read_text("utf-8")

    copy.transfer_page(be, source_id, _TARGET_SECTION)

    _, sent = _sent_payload(be)
    image = next(sent.iter(qn("Image")))
    assert image.find(qn("Data")).text == expected_b64, "pixels inlined into the transplant"
    assert image.find(qn("CallbackID")) is None, "read-side callback must not ride along"


def test_transfer_table_page_preserves_structure_and_shading(be, fixtures_dir):
    source_id, source = _page_by_name(fixtures_dir, "表格頁")
    copy.transfer_page(be, source_id, _TARGET_SECTION)
    _, sent = _sent_payload(be)
    src_table = next(source.iter(qn("Table")))
    sent_table = next(sent.iter(qn("Table")))
    assert len(sent_table.findall(qn("Row"))) == len(src_table.findall(qn("Row")))
    assert len(sent_table.findall(f"{qn('Columns')}/{qn('Column')}")) == len(
        src_table.findall(f"{qn('Columns')}/{qn('Column')}")
    )
    src_shaded = [c.get("shadingColor") for c in src_table.iter(qn("Cell"))]
    sent_shaded = [c.get("shadingColor") for c in sent_table.iter(qn("Cell"))]
    assert sent_shaded == src_shaded, "cell shading preserved verbatim"


def test_transfer_preserves_page_level_via_whole_batch(be, fixtures_dir, tmp_path):
    # synthesize a level-2 source from the real mixed page; target hierarchy contains the
    # about-to-be-created page (deterministic fake ID), plus a sibling that must ride along
    source_id, source = _page_by_name(fixtures_dir, "混合樣式頁")
    level2 = etree.fromstring(etree.tostring(source), parser=_PARSER)
    level2.set("pageLevel", "2")
    sub_id = "{SUB}{1}{B0}"
    level2.set("ID", sub_id)
    (tmp_path / f"page_{_sanitize(sub_id)}__binary.xml").write_bytes(etree.tostring(level2))
    (tmp_path / f"hierarchy_hsPages__{_sanitize(_TARGET_SECTION)}.xml").write_text(
        '<?xml version="1.0"?><one:Section '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{_TARGET_SECTION}" name="目標節">'
        '<one:Page ID="{EXISTING}{1}{B0}" name="既有頁" pageLevel="1"/>'
        f'<one:Page ID="{_NEW_PAGE_ID}" name="新頁" pageLevel="1"/>'
        "</one:Section>",
        encoding="utf-8",
    )

    copy.transfer_page(be, sub_id, _TARGET_SECTION)

    (batch,) = [c for c in be.calls if c.method == "update_hierarchy"]
    sent = etree.fromstring(batch.kwargs["changes_xml"].encode("utf-8"))
    pages = sent.findall(qn("Page"))
    assert [p.get("ID") for p in pages] == ["{EXISTING}{1}{B0}", _NEW_PAGE_ID], (
        "complete page list, order unchanged"
    )
    assert [p.get("pageLevel") for p in pages] == ["1", "2"], "source pageLevel preserved"


def test_transfer_level1_page_skips_the_hierarchy_batch(be, fixtures_dir):
    source_id, _ = _page_by_name(fixtures_dir, "混合樣式頁")  # pageLevel=1
    copy.transfer_page(be, source_id, _TARGET_SECTION)
    assert not [c for c in be.calls if c.method == "update_hierarchy"]
