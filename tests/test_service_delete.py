"""Phase 6 Stage 1: delete tools on FixtureBackend (Linux green).

delete_node is a thin DeleteHierarchy facade (recycle-bin default). delete_page_content is
the interesting one: it validates the target is a PAGE-LEVEL object before the COM call, so
the VM-verified refusal of inline OEs (paragraph OEs 0x8004200E, attachment OEs likewise)
becomes a clear error instead of an opaque COM failure. Asserted against the real attachment
page (page-level Outline/Image/InsertedFile + inline paragraph/attachment OEs) and table page.
"""

from __future__ import annotations

import pytest
from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service import delete
from onenote_com_mcp.xmllayer.namespaces import qn

_PAGE_PREFIX = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}"
PAGE_1_ID = f"{_PAGE_PREFIX}{{E19113778280437107490620108457794036705249631}}"
TABLE_PAGE_ID = f"{_PAGE_PREFIX}{{E19500773287729139935320149797721816501902621}}"

FF = "{FF3818B8-3EE5-0E18-33E2-DFECB54FC950}"
C4 = "{C4BBAA5E-D7C6-0E3C-33BC-D6C40B7752D5}"

# page-1 (附件與嵌入物件-1) object IDs, from the committed dump
OUTLINE_ID = f"{C4}{{43}}{{B0}}"  # page-level one:Outline
DOCX_FILE_ID = f"{FF}{{16}}{{B0}}"  # page-level one:InsertedFile (own objectID)
RENDER_IMAGE_ID = f"{FF}{{77}}{{B0}}"  # page-level one:Image (printout render)
PARAGRAPH_OE_ID = f"{C4}{{54}}{{B0}}"  # inline paragraph OE (holds one:T)
ATTACHMENT_OE_ID = f"{C4}{{53}}{{B0}}"  # inline OE wrapping the txt one:InsertedFile


@pytest.fixture
def be(fixtures_dir) -> FixtureBackend:
    return FixtureBackend(fixtures_dir)


def _deletes(be: FixtureBackend, method: str):
    return [c for c in be.calls if c.method == method]


# --- delete_node -------------------------------------------------------------------------


def test_delete_node_recycle_bin_default(be):
    node_id = "{SOME-SECTION}{1}{B0}"
    delete.delete_node(be, node_id)
    [call] = _deletes(be, "delete_hierarchy")
    assert call.kwargs["object_id"] == node_id
    assert call.kwargs["permanent"] is False  # recycle bin = the undo net (default)


def test_delete_node_permanent(be):
    delete.delete_node(be, "{P}{1}{B0}", permanent=True)
    [call] = _deletes(be, "delete_hierarchy")
    assert call.kwargs["permanent"] is True


# --- delete_page_content: page-level targets accepted ------------------------------------


@pytest.mark.parametrize("object_id", [OUTLINE_ID, DOCX_FILE_ID, RENDER_IMAGE_ID])
def test_delete_page_level_object(be, object_id):
    delete.delete_page_content(be, PAGE_1_ID, object_id)
    [call] = _deletes(be, "delete_page_content")
    assert call.kwargs["page_id"] == PAGE_1_ID
    assert call.kwargs["object_id"] == object_id
    assert call.kwargs["force"] is False
    # concurrency guard: the stamp from the page we read to validate (read-then-act)
    assert call.kwargs["expected_last_modified"] is not None


def test_delete_page_content_force_passthrough(be):
    delete.delete_page_content(be, PAGE_1_ID, OUTLINE_ID, force=True)
    [call] = _deletes(be, "delete_page_content")
    assert call.kwargs["force"] is True


# --- delete_page_content: inline targets refused with guidance ---------------------------


def test_paragraph_oe_refused_with_guidance(be):
    with pytest.raises(ValueError, match="a paragraph"):
        delete.delete_page_content(be, PAGE_1_ID, PARAGRAPH_OE_ID)
    assert not _deletes(be, "delete_page_content"), "must not reach COM"


def test_inline_attachment_oe_refused_with_guidance(be):
    with pytest.raises(ValueError, match="an inline attachment"):
        delete.delete_page_content(be, PAGE_1_ID, ATTACHMENT_OE_ID)
    assert not _deletes(be, "delete_page_content")


def test_table_refused_with_guidance(be, fixtures_dir):
    root = etree.fromstring((fixtures_dir / _table_fixture_name(fixtures_dir)).read_bytes())
    table_id = next(t.get("objectID") for t in root.iter(qn("Table")) if t.get("objectID"))
    with pytest.raises(ValueError, match="a table"):
        delete.delete_page_content(be, TABLE_PAGE_ID, table_id)
    assert not _deletes(be, "delete_page_content")


def test_unknown_object_id_raises_not_found(be):
    with pytest.raises(NodeNotFoundError, match="not on page"):
        delete.delete_page_content(be, PAGE_1_ID, "{NOPE}{9}{B0}")
    assert not _deletes(be, "delete_page_content")


def _table_fixture_name(fixtures_dir) -> str:
    from onenote_com_mcp.backend.fixture import _sanitize

    return f"page_{_sanitize(TABLE_PAGE_ID)}.xml"
