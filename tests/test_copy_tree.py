"""Tier-1 tests for Phase-5 Stage-2: transfer_section (incl. section-group targets).

The section copy walks REAL page fixtures (content fidelity is test_copy_page's job — here
the contract is orchestration: order, pageLevel batching, name de-collision, group targets).
"""

from __future__ import annotations

import pytest
from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize
from onenote_com_mcp.enums import CreateFileType
from onenote_com_mcp.service import copy

_ONE = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_SRC_SECTION = "{SRCSEC}{1}{B0}"
_TARGET_NB = "{TARGETNB}{1}{B0}"

# the three REAL dumped pages (content fixtures + binary exist for these)
_MIXED = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E19540013362017467321520163829129860902849621}"
_TABLE = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E19500773287729139935320149797721816501902621}"
_IMAGE = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E1953306013858222940101982353039053288030011}"
_TABLE2 = "{TBL2}{1}{B0}"  # level-2 variant of the table page, written by the fixture

_NEW_SECTION = "{FIXTURE-cftSection-1}{1}{B0}"  # deterministic fake-ID sequence
_NEW_PAGES = [f"{{FIXTURE-page-{n}}}{{1}}{{B0}}" for n in (2, 3, 4)]


def _write(tmp_path, name: str, body: str) -> None:
    (tmp_path / name).write_text(f'<?xml version="1.0"?>{body}', encoding="utf-8")


def _blank_page(tmp_path, page_id: str) -> None:
    _write(
        tmp_path,
        f"page_{_sanitize(page_id)}.xml",
        f'<one:Page xmlns:one="{_ONE}" ID="{page_id}" '
        'lastModifiedTime="2026-06-11T09:00:00.000Z"/>',
    )


@pytest.fixture
def section_be(fixtures_dir, tmp_path) -> FixtureBackend:
    for src in list(fixtures_dir.glob("page_*.xml")) + list(fixtures_dir.glob("binary_*.b64")):
        (tmp_path / src.name).write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
    for page_id in _NEW_PAGES:
        _blank_page(tmp_path, page_id)
    # a level-2 variant of the REAL table page (content XML carries pageLevel — live, the
    # hierarchy and page XML agree; the replay fixtures must too)
    table2 = etree.fromstring(
        (fixtures_dir / f"page_{_sanitize(_TABLE)}__binary.xml").read_bytes(),
        etree.XMLParser(strip_cdata=False),
    )
    table2.set("ID", _TABLE2)
    table2.set("pageLevel", "2")
    (tmp_path / f"page_{_sanitize(_TABLE2)}__binary.xml").write_bytes(etree.tostring(table2))
    # source section: the three real pages, 表格頁 as a level-2 subpage
    _write(
        tmp_path,
        f"hierarchy_hsPages__{_sanitize(_SRC_SECTION)}.xml",
        f'<one:Section xmlns:one="{_ONE}" ID="{_SRC_SECTION}" name="藥品紀錄">'
        f'<one:Page ID="{_MIXED}" name="混合樣式頁" pageLevel="1"/>'
        f'<one:Page ID="{_TABLE2}" name="表格頁" pageLevel="2"/>'
        f'<one:Page ID="{_IMAGE}" name="圖片頁" pageLevel="1"/>'
        "</one:Section>",
    )
    # target notebook already HAS a section named 藥品紀錄 → the copy must de-collide
    _write(
        tmp_path,
        f"hierarchy_hsSections__{_sanitize(_TARGET_NB)}.xml",
        f'<one:Notebook xmlns:one="{_ONE}" ID="{_TARGET_NB}" name="目標本">'
        f'<one:Section ID="{{EXIST}}{{1}}{{B0}}" name="藥品紀錄"/>'
        "</one:Notebook>",
    )
    # the level-2 batch reads the NEW section's hierarchy (replay: both copied-so-far pages)
    _write(
        tmp_path,
        f"hierarchy_hsPages__{_sanitize(_NEW_SECTION)}.xml",
        f'<one:Section xmlns:one="{_ONE}" ID="{_NEW_SECTION}" name="藥品紀錄 (2)">'
        f'<one:Page ID="{_NEW_PAGES[0]}" name="混合樣式頁" pageLevel="1"/>'
        f'<one:Page ID="{_NEW_PAGES[1]}" name="表格頁" pageLevel="1"/>'
        "</one:Section>",
    )
    return FixtureBackend(tmp_path)


def test_transfer_section_orchestration(section_be):
    be = section_be
    result = copy.transfer_section(be, _SRC_SECTION, _TARGET_NB)
    assert result.section_id == _NEW_SECTION
    assert result.file_notes == []

    # the section is created under the target with a DE-COLLIDED name
    (sec_call,) = [c for c in be.calls if c.method == "open_hierarchy"]
    assert sec_call.kwargs["path"] == "藥品紀錄 (2).one"
    assert sec_call.kwargs["relative_to_object_id"] == _TARGET_NB
    assert sec_call.kwargs["create_file_type"] is CreateFileType.cftSection

    # pages copied in document order into the new section
    creates = [c for c in be.calls if c.method == "create_new_page"]
    assert [c.kwargs["section_id"] for c in creates] == [_NEW_SECTION] * 3
    writes = [c for c in be.calls if c.method == "update_page_content"]
    sent_ids = [etree.fromstring(w.kwargs["changes_xml"].encode("utf-8")).get("ID") for w in writes]
    assert sent_ids == _NEW_PAGES

    # 表格頁 was a level-2 subpage → ONE whole-batch UpdateHierarchy sets it on the copy
    (batch,) = [c for c in be.calls if c.method == "update_hierarchy"]
    sent = etree.fromstring(batch.kwargs["changes_xml"].encode("utf-8"))
    levels = {p.get("ID"): p.get("pageLevel") for p in sent.iter(f"{{{_ONE}}}Page")}
    assert levels[_NEW_PAGES[1]] == "2", "subpage nesting survives the copy"


def test_transfer_section_name_collision_escalates(tmp_path):
    _write(
        tmp_path,
        f"hierarchy_hsPages__{_sanitize(_SRC_SECTION)}.xml",
        f'<one:Section xmlns:one="{_ONE}" ID="{_SRC_SECTION}" name="甲節"/>',
    )
    _write(
        tmp_path,
        f"hierarchy_hsSections__{_sanitize(_TARGET_NB)}.xml",
        f'<one:Notebook xmlns:one="{_ONE}" ID="{_TARGET_NB}" name="目標本">'
        '<one:Section ID="{E1}{1}{B0}" name="甲節"/>'
        '<one:SectionGroup ID="{E2}{1}{B0}" name="甲節 (2)"/>'
        "</one:Notebook>",
    )
    be = FixtureBackend(tmp_path)
    copy.transfer_section(be, _SRC_SECTION, _TARGET_NB)
    (sec_call,) = [c for c in be.calls if c.method == "open_hierarchy"]
    assert sec_call.kwargs["path"] == "甲節 (3).one", "both kinds count as taken names"


# --- section copy INTO a section group ----------------------------------------------------
# (There is deliberately no transfer_notebook: COM notebook creation is refused by this M365
# build — VM ground truth 2026-06-11. Whole-notebook cloning = transfer_section per section.)


def test_transfer_section_into_a_section_group(tmp_path):
    group_id = "{G}{1}{B0}"
    _write(
        tmp_path,
        f"hierarchy_hsPages__{_sanitize(_SRC_SECTION)}.xml",
        f'<one:Section xmlns:one="{_ONE}" ID="{_SRC_SECTION}" name="丙節"/>',
    )
    _write(
        tmp_path,
        f"hierarchy_hsSections__{_sanitize(group_id)}.xml",
        f'<one:SectionGroup xmlns:one="{_ONE}" ID="{group_id}" name="乙群">'
        '<one:Section ID="{S2}{1}{B0}" name="丙節"/>'
        "</one:SectionGroup>",
    )
    be = FixtureBackend(tmp_path)
    copy.transfer_section(be, _SRC_SECTION, group_id)
    (sec_call,) = [c for c in be.calls if c.method == "open_hierarchy"]
    assert sec_call.kwargs["relative_to_object_id"] == group_id, "section lands IN the group"
    assert sec_call.kwargs["path"] == "丙節 (2).one", "de-collided against the group's children"
