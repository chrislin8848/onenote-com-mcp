"""Tier-1 tests for the create service (Phase 4 Stage 2).

FixtureBackend replays reads and records writes, so the assertions here are about the COM
calls the service emits: OpenHierarchy shapes, ONE guarded UpdatePageContent for title +
content, and the whole-batch UpdateHierarchy for pageLevel (ID conservation included).
FixtureBackend's fake IDs are deterministic ("{FIXTURE-<kind>-<n>}{1}{B0}"), which lets the
replay fixtures contain the page the service is about to "create".
"""

from __future__ import annotations

import pytest
from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize
from onenote_com_mcp.enums import CreateFileType
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service import create
from onenote_com_mcp.xmllayer.namespaces import qn

_PARSER = etree.XMLParser(strip_cdata=False)

_SECTION_ID = "{SEC}{1}{B0}"
_FAKE_PAGE_ID = "{FIXTURE-page-1}{1}{B0}"  # FixtureBackend's first create_new_page result


def _calls(be: FixtureBackend, method: str) -> list:
    return [c for c in be.calls if c.method == method]


def _write_blank_page_fixture(tmp_path, body: str = "") -> None:
    (tmp_path / f"page_{_sanitize(_FAKE_PAGE_ID)}.xml").write_text(
        '<?xml version="1.0"?><one:Page '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{_FAKE_PAGE_ID}" lastModifiedTime="2026-06-11T08:00:00.000Z">{body}</one:Page>',
        encoding="utf-8",
    )


# --- create_section / create_notebook --------------------------------------------------


def test_create_section_opens_dot_one_under_parent(tmp_path):
    be = FixtureBackend(tmp_path)
    section_id = create.create_section(be, "{NB}{1}{B0}", "六月行程")
    (call,) = _calls(be, "open_hierarchy")
    assert call.kwargs["path"] == "六月行程.one"
    assert call.kwargs["relative_to_object_id"] == "{NB}{1}{B0}"
    assert call.kwargs["create_file_type"] is CreateFileType.cftSection
    assert section_id.startswith("{FIXTURE-cftSection")


def test_create_notebook_joins_parent_path_and_name(tmp_path):
    be = FixtureBackend(tmp_path)
    create.create_notebook(be, "新筆記本", "C:\\Users\\dev\\OneDrive\\Notebooks\\")
    (call,) = _calls(be, "open_hierarchy")
    assert call.kwargs["path"] == "C:\\Users\\dev\\OneDrive\\Notebooks\\新筆記本"
    assert call.kwargs["relative_to_object_id"] == ""
    assert call.kwargs["create_file_type"] is CreateFileType.cftNotebook


def test_create_notebook_empty_path_uses_default_notebook_folder(tmp_path):
    (tmp_path / "special_slDefaultNotebookFolder.txt").write_text(
        "C:\\Users\\dev\\OneDrive\\文件\\OneNote 筆記本", encoding="utf-8"
    )
    be = FixtureBackend(tmp_path)
    create.create_notebook(be, "新筆記本")
    (call,) = _calls(be, "open_hierarchy")
    assert call.kwargs["path"] == "C:\\Users\\dev\\OneDrive\\文件\\OneNote 筆記本\\新筆記本"


@pytest.mark.parametrize("bad", ["行程/六月", "a:b", "x?y", "  ", "tag#1"])
def test_invalid_names_rejected_before_any_com_call(tmp_path, bad):
    be = FixtureBackend(tmp_path)
    with pytest.raises(ValueError):
        create.create_section(be, "{NB}{1}{B0}", bad)
    with pytest.raises(ValueError):
        create.create_notebook(be, bad, "C:\\x")
    assert not be.calls


# --- create_page ------------------------------------------------------------------------


def test_create_page_sets_title_and_content_in_one_guarded_write(tmp_path):
    _write_blank_page_fixture(tmp_path)
    be = FixtureBackend(tmp_path)

    page_id = create.create_page(be, _SECTION_ID, "會議記錄", "第一行\n第二行")

    assert page_id == _FAKE_PAGE_ID
    (created,) = _calls(be, "create_new_page")
    assert created.kwargs["section_id"] == _SECTION_ID
    (write,) = _calls(be, "update_page_content")  # title + content = ONE UpdatePageContent
    assert write.kwargs["force"] is False
    assert write.kwargs["expected_last_modified"] is not None
    sent = etree.fromstring(write.kwargs["changes_xml"].encode("utf-8"), parser=_PARSER)
    title_oe = sent.find(f"{qn('Title')}/{qn('OE')}")
    assert "會議記錄" in (title_oe.find(qn("T")).text or "")
    oes = sent.findall(f"{qn('Outline')}/{qn('OEChildren')}/{qn('OE')}")
    assert ["第一行" in (oe.find(qn("T")).text or "") for oe in oes] == [True, False]
    assert "第二行" in (oes[1].find(qn("T")).text or "")
    # Title must precede the outline (fresh-page element order)
    children = [etree.QName(el).localname for el in sent]
    assert children.index("Title") < children.index("Outline")


def test_create_page_replaces_existing_title_oe_keeping_identity(tmp_path):
    _write_blank_page_fixture(
        tmp_path,
        '<one:Title><one:OE objectID="{T}{1}{B0}" quickStyleIndex="0">'
        "<one:T><![CDATA[無標題]]></one:T></one:OE></one:Title>",
    )
    be = FixtureBackend(tmp_path)
    create.create_page(be, _SECTION_ID, "正式標題")
    (write,) = _calls(be, "update_page_content")
    sent = etree.fromstring(write.kwargs["changes_xml"].encode("utf-8"), parser=_PARSER)
    title_oe = sent.find(f"{qn('Title')}/{qn('OE')}")
    assert title_oe.get("objectID") == "{T}{1}{B0}", "existing Title OE identity must be kept"
    assert title_oe.get("quickStyleIndex") == "0"
    assert "正式標題" in (title_oe.find(qn("T")).text or "")
    assert "無標題" not in write.kwargs["changes_xml"]


def test_create_page_without_title_or_content_writes_nothing(tmp_path):
    be = FixtureBackend(tmp_path)
    page_id = create.create_page(be, _SECTION_ID)
    assert page_id == _FAKE_PAGE_ID
    assert not _calls(be, "update_page_content")
    assert not _calls(be, "update_hierarchy")


def test_create_page_page_level_submits_whole_batch(tmp_path):
    hierarchy = (
        '<?xml version="1.0"?><one:Section '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{_SECTION_ID}" name="節">'
        '<one:Page ID="{P1}{1}{B0}" name="頁一" pageLevel="1"/>'
        '<one:Page ID="{P2}{1}{B0}" name="子頁" pageLevel="2"/>'
        f'<one:Page ID="{_FAKE_PAGE_ID}" name="新頁" pageLevel="1"/>'
        "</one:Section>"
    )
    (tmp_path / f"hierarchy_hsPages__{_sanitize(_SECTION_ID)}.xml").write_text(
        hierarchy, encoding="utf-8"
    )
    be = FixtureBackend(tmp_path)

    create.create_page(be, _SECTION_ID, page_level=2)

    (batch,) = _calls(be, "update_hierarchy")  # ONE whole-batch UpdateHierarchy
    sent = etree.fromstring(batch.kwargs["changes_xml"].encode("utf-8"))
    pages = sent.findall(qn("Page"))
    assert [p.get("ID") for p in pages] == ["{P1}{1}{B0}", "{P2}{1}{B0}", _FAKE_PAGE_ID], (
        "the COMPLETE page list rides along, order unchanged"
    )
    assert [p.get("pageLevel") for p in pages] == ["1", "2", "2"]


def test_create_page_page_level_fails_clearly_when_page_missing_from_hierarchy(tmp_path):
    (tmp_path / f"hierarchy_hsPages__{_sanitize(_SECTION_ID)}.xml").write_text(
        '<?xml version="1.0"?><one:Section '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{_SECTION_ID}" name="節"><one:Page ID="{{P1}}{{1}}{{B0}}" pageLevel="1"/>'
        "</one:Section>",
        encoding="utf-8",
    )
    be = FixtureBackend(tmp_path)
    with pytest.raises(NodeNotFoundError, match="pageLevel"):
        create.create_page(be, _SECTION_ID, page_level=2)
    assert not _calls(be, "update_hierarchy"), "no partial-list write on failure"


def test_create_page_invalid_level_rejected_before_creation(tmp_path):
    be = FixtureBackend(tmp_path)
    with pytest.raises(ValueError, match="page_level"):
        create.create_page(be, _SECTION_ID, "t", page_level=4)
    assert not be.calls


def test_create_notebook_url_path_joins_with_slash(tmp_path):
    # synced locations are URLs — the notebook folder must join with "/" (VM ground truth:
    # this M365 build refuses local-path notebooks, so URL parents are the normal case)
    be = FixtureBackend(tmp_path)
    create.create_notebook(be, "新本", "https://d.docs.live.net/abc/Documents/")
    (call,) = [c for c in be.calls if c.method == "open_hierarchy"]
    assert call.kwargs["path"] == "https://d.docs.live.net/abc/Documents/新本"
