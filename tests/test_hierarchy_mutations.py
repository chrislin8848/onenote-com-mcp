"""Tier-1 tests for the Phase-4 Stage-3 hierarchy mutators, against REAL VM fixtures.

The real notebook fixture is ideal ground truth here: its direct children are a MIXED
Section + recycle-bin SectionGroup + SectionGroup list, so completeness, both-kinds, and
recycle-bin pinning are all exercised on real data. The seam's node-ID conservation is
guard-tested in test_hierarchy_core.py; these tests cover the mutators' own contracts.
"""

from __future__ import annotations

import pytest
from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service import hierarchy_edit
from onenote_com_mcp.xmllayer.namespaces import qn

NOTEBOOK_ID = "{C94E632E-9829-45FF-914E-5E4031B2439D}{1}{B0}"
SECTION_ID = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{B0}"  # Phase 0 測試用 (8 pages)
GROUP_ID = "{5724F5C7-9310-4BF6-9BE2-108011605C2E}{1}{B0}"  # 節群組 測試用
RECYCLE_GROUP_ID = "{85A5F7F6-7D27-4C1B-9E95-E0F823864630}{1}{B0}"  # OneNote_RecycleBin
RECYCLE_SECTION_ID = "{17E428A8-3C4B-4138-B9F4-220E154F704E}{1}{B0}"  # 刪除的頁面
SEC1_ID = "{42020881-06D7-426D-B9FA-E742EB5C150D}{1}{B0}"  # 第1節 (2 pages)
SEC2_ID = "{C565F5C9-1AA3-480D-B0F8-4A5086FE359A}{1}{B0}"  # 第2節 (1 page)


@pytest.fixture
def be(fixtures_dir) -> FixtureBackend:
    return FixtureBackend(fixtures_dir)


@pytest.fixture
def notebook_pages(fixtures_dir) -> etree._Element:
    path = fixtures_dir / f"hierarchy_hsPages__{_sanitize(NOTEBOOK_ID)}.xml"
    return etree.fromstring(path.read_bytes())


@pytest.fixture
def section_be(fixtures_dir, tmp_path) -> FixtureBackend:
    """Backend whose section-scoped hsPages read replays the real notebook dump (the section
    is found inside it, exactly as a scoped GetHierarchy would return it as root)."""
    src = fixtures_dir / f"hierarchy_hsPages__{_sanitize(NOTEBOOK_ID)}.xml"
    dst = tmp_path / f"hierarchy_hsPages__{_sanitize(SECTION_ID)}.xml"
    dst.write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
    return FixtureBackend(tmp_path)


def _sent(be: FixtureBackend) -> etree._Element:
    writes = [c for c in be.calls if c.method == "update_hierarchy"]
    assert len(writes) == 1, "a structural change must be exactly ONE UpdateHierarchy"
    return etree.fromstring(writes[0].kwargs["changes_xml"].encode("utf-8"))


def _no_write(be: FixtureBackend) -> bool:
    return not [c for c in be.calls if c.method == "update_hierarchy"]


def _section_pages(tree: etree._Element, section_id: str) -> list[etree._Element]:
    section = next(el for el in tree.iter(qn("Section")) if el.get("ID") == section_id)
    return section.findall(qn("Page"))


# --- restructure_section ---------------------------------------------------------------


def test_restructure_section_reorders_whole_batch_and_sets_levels(section_be, notebook_pages):
    pages = _section_pages(notebook_pages, SECTION_ID)
    ids = [p.get("ID") for p in pages]
    # move the last page to the front, demote it to a level-2 subpage of the new last page
    target = [{"page_id": ids[-1], "page_level": 2}] + [{"page_id": i} for i in ids[:-1]]

    hierarchy_edit.restructure_section(section_be, SECTION_ID, target)

    sent_pages = _section_pages(_sent(section_be), SECTION_ID)
    assert [p.get("ID") for p in sent_pages] == [ids[-1]] + ids[:-1]
    assert sent_pages[0].get("pageLevel") == "2", "requested page_level applied"
    # entries without page_level keep their original levels (this section has 1/2/3 mix)
    original_levels = {p.get("ID"): p.get("pageLevel") for p in pages}
    for p in sent_pages[1:]:
        assert p.get("pageLevel") == original_levels[p.get("ID")]


def test_restructure_section_other_sections_ride_along_untouched(section_be, notebook_pages):
    pages = _section_pages(notebook_pages, SECTION_ID)
    ids = [p.get("ID") for p in pages]
    hierarchy_edit.restructure_section(
        section_be, SECTION_ID, [{"page_id": i} for i in reversed(ids)]
    )
    sent = _sent(section_be)
    # the whole scope is submitted: untouched sections keep their page lists verbatim
    for sid in (SEC1_ID, SEC2_ID):
        before = [p.get("ID") for p in _section_pages(notebook_pages, sid)]
        after = [p.get("ID") for p in _section_pages(sent, sid)]
        assert after == before


def test_restructure_section_partial_list_rejected(section_be, notebook_pages):
    ids = [p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID)]
    with pytest.raises(ValueError, match="COMPLETE"):
        hierarchy_edit.restructure_section(
            section_be, SECTION_ID, [{"page_id": i} for i in ids[:-1]]
        )
    assert _no_write(section_be), "no partial-list write may reach OneNote"


def test_restructure_section_unknown_and_duplicate_ids_rejected(section_be, notebook_pages):
    ids = [p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID)]
    with pytest.raises(ValueError, match="unknown"):
        hierarchy_edit.restructure_section(
            section_be, SECTION_ID, [{"page_id": i} for i in [*ids[:-1], "{BOGUS}{1}{B0}"]]
        )
    with pytest.raises(ValueError, match="duplicates"):
        hierarchy_edit.restructure_section(
            section_be, SECTION_ID, [{"page_id": i} for i in [*ids, ids[0]]]
        )
    assert _no_write(section_be)


def test_restructure_section_bad_entry_rejected_before_any_read(section_be):
    with pytest.raises(ValueError, match="page_level"):
        hierarchy_edit.restructure_section(
            section_be, SECTION_ID, [{"page_id": "{P}{1}{B0}", "page_level": 4}]
        )
    with pytest.raises(ValueError, match="page_id"):
        hierarchy_edit.restructure_section(section_be, SECTION_ID, [{"page_level": 1}])
    assert not section_be.calls


# --- reposition_page (single-page move; the seam supplies §5 completeness) --------------


def test_reposition_page_moves_one_page_after_anchor(section_be, notebook_pages):
    ids = [p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID)]
    # move the LAST page to right after the FIRST — caller gives only the two IDs
    hierarchy_edit.reposition_page(section_be, SECTION_ID, ids[-1], after_page_id=ids[0])
    sent_ids = [p.get("ID") for p in _section_pages(_sent(section_be), SECTION_ID)]
    assert sent_ids == [ids[0], ids[-1], *ids[1:-1]]
    assert set(sent_ids) == set(ids), "node-ID conserved — nothing dropped or invented"


def test_reposition_page_to_top_when_after_is_empty(section_be, notebook_pages):
    ids = [p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID)]
    hierarchy_edit.reposition_page(section_be, SECTION_ID, ids[2])  # empty after → section top
    sent_ids = [p.get("ID") for p in _section_pages(_sent(section_be), SECTION_ID)]
    assert sent_ids == [ids[2], *[i for i in ids if i != ids[2]]]


def test_reposition_page_sets_optional_page_level(section_be, notebook_pages):
    ids = [p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID)]
    hierarchy_edit.reposition_page(
        section_be, SECTION_ID, ids[-1], after_page_id=ids[0], page_level=2
    )
    moved = next(p for p in _section_pages(_sent(section_be), SECTION_ID) if p.get("ID") == ids[-1])
    assert moved.get("pageLevel") == "2"


def test_reposition_page_other_sections_ride_along_untouched(section_be, notebook_pages):
    ids = [p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID)]
    hierarchy_edit.reposition_page(section_be, SECTION_ID, ids[-1], after_page_id=ids[0])
    sent = _sent(section_be)
    for sid in (SEC1_ID, SEC2_ID):
        before = [p.get("ID") for p in _section_pages(notebook_pages, sid)]
        after = [p.get("ID") for p in _section_pages(sent, sid)]
        assert after == before, "whole batch submitted — untouched sections verbatim"


def test_reposition_page_rejects_after_self(section_be, notebook_pages):
    ids = [p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID)]
    with pytest.raises(ValueError, match="cannot be the page being moved"):
        hierarchy_edit.reposition_page(section_be, SECTION_ID, ids[0], after_page_id=ids[0])
    assert _no_write(section_be)


def test_reposition_page_rejects_bad_level_before_any_read(section_be):
    with pytest.raises(ValueError, match="page_level"):
        hierarchy_edit.reposition_page(section_be, SECTION_ID, "{P}{1}{B0}", page_level=4)
    assert not section_be.calls, "bad page_level rejected before any COM read"


def test_reposition_page_rejects_unknown_page(section_be):
    with pytest.raises(NodeNotFoundError):
        hierarchy_edit.reposition_page(section_be, SECTION_ID, "{NOPE}{1}{B0}")
    assert _no_write(section_be)


# --- reorder_sections ------------------------------------------------------------------


def test_reorder_sections_rejects_groups_before_sections(be):
    # VM ground truth (2026-06-11, hrInvalidXML): OneNote's hierarchy schema is positional —
    # every one:Section precedes all one:SectionGroup siblings. Interleaving is refused
    # BEFORE any COM write.
    with pytest.raises(ValueError, match="section groups"):
        hierarchy_edit.reorder_sections(be, NOTEBOOK_ID, [GROUP_ID, SECTION_ID])
    assert _no_write(be)


def test_reorder_sections_keeps_recycle_bin_pinned_in_batch(be):
    # real notebook children: [Section, RecycleBin group, SectionGroup] — a kind-respecting
    # full submission goes through; the hidden recycle bin keeps its slot AND rides along
    hierarchy_edit.reorder_sections(be, NOTEBOOK_ID, [SECTION_ID, GROUP_ID])
    sent = _sent(be)
    children = [(etree.QName(c).localname, c.get("ID")) for c in sent]
    assert children == [
        ("Section", SECTION_ID),
        ("SectionGroup", RECYCLE_GROUP_ID),
        ("SectionGroup", GROUP_ID),
    ]
    group = sent[-1]
    assert [s.get("ID") for s in group.findall(qn("Section"))] == [SEC1_ID, SEC2_ID], (
        "nesting inside the groups is untouched"
    )


def test_reorder_sections_must_include_section_groups_too(be):
    with pytest.raises(ValueError, match="COMPLETE"):
        hierarchy_edit.reorder_sections(be, NOTEBOOK_ID, [SECTION_ID])
    assert _no_write(be)


def test_reorder_sections_works_inside_a_section_group(fixtures_dir, tmp_path):
    # scope = the section group: its complete child list is just its two sections
    src = fixtures_dir / f"hierarchy_hsSections__{_sanitize(NOTEBOOK_ID)}.xml"
    dst = tmp_path / f"hierarchy_hsSections__{_sanitize(GROUP_ID)}.xml"
    dst.write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
    be = FixtureBackend(tmp_path)

    hierarchy_edit.reorder_sections(be, GROUP_ID, [SEC2_ID, SEC1_ID])
    sent = _sent(be)
    group = next(el for el in sent.iter(qn("SectionGroup")) if el.get("ID") == GROUP_ID)
    assert [s.get("ID") for s in group.findall(qn("Section"))] == [SEC2_ID, SEC1_ID]


# --- rename_node -----------------------------------------------------------------------


def test_rename_page_routes_through_title_edit(be, notebook_pages):
    # VM ground truth (2026-06-11): UpdateHierarchy silently IGNORES a one:Page name attr —
    # the hierarchy name follows the TITLE, so a page rename is a title edit through the
    # page-content seam (one guarded UpdatePageContent, zero UpdateHierarchy).
    page_id = next(
        p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID) if p.get("name") == "圖片頁"
    )
    hierarchy_edit.rename_node(be, NOTEBOOK_ID, page_id, "圖片頁改名")
    assert _no_write(be), "a page rename must not touch UpdateHierarchy"
    (write,) = [c for c in be.calls if c.method == "update_page_content"]
    assert "圖片頁改名" in write.kwargs["changes_xml"]
    assert write.kwargs["expected_last_modified"] is not None

    with pytest.raises(ValueError, match="page name"):
        hierarchy_edit.rename_node(be, NOTEBOOK_ID, page_id, "   ")


def test_rename_section_group_and_section_validate_filename_chars(be):
    hierarchy_edit.rename_node(be, NOTEBOOK_ID, GROUP_ID, "改名群組")
    sent = _sent(be)
    assert next(el for el in sent.iter() if el.get("ID") == GROUP_ID).get("name") == "改名群組"

    be2 = FixtureBackend(be.fixtures_dir)
    with pytest.raises(ValueError, match="forbids"):
        hierarchy_edit.rename_node(be2, NOTEBOOK_ID, SEC1_ID, "六月/行程")
    assert _no_write(be2)


def test_rename_notebook_rejected(be):
    with pytest.raises(ValueError, match="one:Notebook"):
        hierarchy_edit.rename_node(be, NOTEBOOK_ID, NOTEBOOK_ID, "新名")
    assert _no_write(be)


# --- move_page (EXPERIMENTAL) ----------------------------------------------------------


@pytest.fixture
def move_be(fixtures_dir, tmp_path) -> FixtureBackend:
    """move_page reads the notebook AND re-reads the target section (for the page's NEW ID —
    VM ground truth: re-parenting re-IDs the page). Serve the real dump under both names."""
    src = (fixtures_dir / f"hierarchy_hsPages__{_sanitize(NOTEBOOK_ID)}.xml").read_text("utf-8")
    for node_id in (NOTEBOOK_ID, SEC1_ID):
        (tmp_path / f"hierarchy_hsPages__{_sanitize(node_id)}.xml").write_text(src, "utf-8")
    return FixtureBackend(tmp_path)


def test_move_page_reparents_resets_level_and_keeps_both_lists(move_be, notebook_pages):
    be = move_be
    # 測試頁面7 is a level-2 subpage in the big section; move it to 第1節
    moved_id = next(
        p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID) if p.get("pageLevel") == "2"
    )
    source_before = [p.get("ID") for p in _section_pages(notebook_pages, SECTION_ID)]
    target_before = [p.get("ID") for p in _section_pages(notebook_pages, SEC1_ID)]

    new_id = hierarchy_edit.move_page(be, NOTEBOOK_ID, moved_id, SEC1_ID)
    # replay semantics: the post-move re-read serves the pre-move fixture, so the "new ID"
    # resolves to the fixture target section's last page; live, OneNote re-IDs the moved page
    assert new_id == target_before[-1]

    sent = _sent(be)
    source_after = [p.get("ID") for p in _section_pages(sent, SECTION_ID)]
    target_after = _section_pages(sent, SEC1_ID)
    assert moved_id not in source_after
    assert source_after == [i for i in source_before if i != moved_id], "source order kept"
    assert [p.get("ID") for p in target_after] == [*target_before, moved_id], "appended last"
    assert target_after[-1].get("pageLevel") == "1", "subpage level resets on move"


def test_move_page_unknown_target_hints_cross_notebook(be, notebook_pages):
    page_id = _section_pages(notebook_pages, SEC2_ID)[0].get("ID")
    with pytest.raises(NodeNotFoundError, match="cross-notebook"):
        hierarchy_edit.move_page(be, NOTEBOOK_ID, page_id, "{ELSEWHERE}{1}{B0}")
    assert _no_write(be)


def test_move_page_target_must_be_a_section(be, notebook_pages):
    page_id = _section_pages(notebook_pages, SEC2_ID)[0].get("ID")
    with pytest.raises(ValueError, match="not a section"):
        hierarchy_edit.move_page(be, NOTEBOOK_ID, page_id, GROUP_ID)
    with pytest.raises(ValueError, match="not a page"):
        hierarchy_edit.move_page(be, NOTEBOOK_ID, SEC1_ID, SEC2_ID)
    assert _no_write(be)


def test_move_page_into_recycle_bin_rejected(be, notebook_pages):
    page_id = _section_pages(notebook_pages, SEC2_ID)[0].get("ID")
    with pytest.raises(ValueError, match="recycle bin"):
        hierarchy_edit.move_page(be, NOTEBOOK_ID, page_id, RECYCLE_SECTION_ID)
    assert _no_write(be)
