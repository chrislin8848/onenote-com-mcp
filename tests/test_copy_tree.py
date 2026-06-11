"""Tier-1 tests for Phase-5 Stage-2: transfer_section / transfer_notebook.

The section copy walks REAL page fixtures (content fidelity is test_copy_page's job — here
the contract is orchestration: order, pageLevel batching, name de-collision). The notebook
copy is structure-only sources (empty sections): groups recreated via cftFolder in document
order, sections landing INSIDE their groups, the recycle-bin group skipped (user-approved
policy).
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
    new_id = copy.transfer_section(be, _SRC_SECTION, _TARGET_NB)
    assert new_id == _NEW_SECTION

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


# --- transfer_notebook -------------------------------------------------------------------

_SRC_NB = "{SRCNB}{1}{B0}"
_S1, _S2 = "{S1}{1}{B0}", "{S2}{1}{B0}"
_NEW_NB = "{FIXTURE-cftNotebook-1}{1}{B0}"
_NEW_GROUP = "{FIXTURE-cftFolder-3}{1}{B0}"  # nb=1, S1-copy=2, folder=3, S2-copy=4


@pytest.fixture
def notebook_be(tmp_path) -> FixtureBackend:
    _write(
        tmp_path,
        f"hierarchy_hsSections__{_sanitize(_SRC_NB)}.xml",
        f'<one:Notebook xmlns:one="{_ONE}" ID="{_SRC_NB}" name="來源本">'
        f'<one:Section ID="{_S1}" name="甲節"/>'
        '<one:SectionGroup ID="{RB}{1}{B0}" name="OneNote_RecycleBin" isRecycleBin="true">'
        '<one:Section ID="{RBS}{1}{B0}" name="刪除的頁面"/>'
        "</one:SectionGroup>"
        '<one:SectionGroup ID="{G}{1}{B0}" name="乙群">'
        f'<one:Section ID="{_S2}" name="丙節"/>'
        "</one:SectionGroup>"
        "</one:Notebook>",
    )
    # structure-only sources: both sections are empty (page fidelity is covered elsewhere)
    for sid, sname in ((_S1, "甲節"), (_S2, "丙節")):
        _write(
            tmp_path,
            f"hierarchy_hsPages__{_sanitize(sid)}.xml",
            f'<one:Section xmlns:one="{_ONE}" ID="{sid}" name="{sname}"/>',
        )
    # unique-name lookups against the freshly created targets (empty replay containers)
    _write(
        tmp_path,
        f"hierarchy_hsSections__{_sanitize(_NEW_NB)}.xml",
        f'<one:Notebook xmlns:one="{_ONE}" ID="{_NEW_NB}" name="克隆本"/>',
    )
    _write(
        tmp_path,
        f"hierarchy_hsSections__{_sanitize(_NEW_GROUP)}.xml",
        f'<one:SectionGroup xmlns:one="{_ONE}" ID="{_NEW_GROUP}" name="乙群"/>',
    )
    return FixtureBackend(tmp_path)


def test_transfer_notebook_recreates_groups_and_skips_recycle_bin(notebook_be):
    be = notebook_be
    new_id = copy.transfer_notebook(be, _SRC_NB, "克隆本", "C:\\Users\\dev\\OneDrive\\NB")
    assert new_id == _NEW_NB

    opens = [c.kwargs for c in be.calls if c.method == "open_hierarchy"]
    assert [(o["path"], o["relative_to_object_id"], o["create_file_type"]) for o in opens] == [
        ("C:\\Users\\dev\\OneDrive\\NB\\克隆本", "", CreateFileType.cftNotebook),
        ("甲節.one", _NEW_NB, CreateFileType.cftSection),
        ("乙群", _NEW_NB, CreateFileType.cftFolder),
        ("丙節.one", _NEW_GROUP, CreateFileType.cftSection),
    ], (
        "document order; the section lands INSIDE its recreated group; "
        "the recycle-bin group (and its inner section) is never cloned"
    )
