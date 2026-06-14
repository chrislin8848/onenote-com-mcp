"""Tier-1 tests for the multi-page copy block: copy_pages / copy_page_subtree (the fix for
"copy ●ITIN and its subpages below ●Local", which a string of copy_page calls scattered).

Three layers, mirroring the existing split:
  * ``hierarchy_edit.reposition_pages`` — the block-placement core, against the REAL notebook
    fixture ("Phase 0 測試用" has a L1 page + a L2/L3 subpage tree — ideal ground truth).
  * ``copy.subtree_page_ids`` — OneNote's positional subpage model, same fixture.
  * ``copy.copy_pages`` engine + the two server facades — orchestration, patched (no backend),
    like test_copy_page_placement.py.
"""

from __future__ import annotations

import json
from types import SimpleNamespace

import pytest
from lxml import etree

from onenote_com_mcp import server
from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize
from onenote_com_mcp.errors import NodeNotFoundError
from onenote_com_mcp.service import copy, hierarchy_edit
from onenote_com_mcp.xmllayer.namespaces import qn

NOTEBOOK_ID = "{C94E632E-9829-45FF-914E-5E4031B2439D}{1}{B0}"
SECTION_ID = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{B0}"  # Phase 0 測試用
# pages of that section, in document order, with their levels (shared section prefix factored out)
_PFX = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{"
P_MIXED = _PFX + "E19540013362017467321520163829129860902849621}"  # L1 混合樣式頁
P_TABLE = _PFX + "E19500773287729139935320149797721816501902621}"  # L1 表格頁
P_IMAGE = _PFX + "E1953306013858222940101982353039053288030011}"  # L1 圖片頁
P_MULTI = _PFX + "E1951925919026217488301965742466186134432541}"  # L1 單節多頁
P_ROLE = _PFX + "E1947901022035288224521936062629174653003591}"  # L2 測試頁面7
P_DUTY = _PFX + "E19540249790908611533520153475103951208513811}"  # L3 測試頁面6
P_COMM = _PFX + "E19525711178706526991120182911230823215973441}"  # L3 測試頁面5
P_TIME = _PFX + "E1948435450880818651941999633141658297826701}"  # L2 測試頁面4

ORIGINAL_ORDER = [P_MIXED, P_TABLE, P_IMAGE, P_MULTI, P_ROLE, P_DUTY, P_COMM, P_TIME]


@pytest.fixture
def section_be(fixtures_dir, tmp_path) -> FixtureBackend:
    """Backend whose section-scoped hsPages read replays the real notebook dump (the section is
    found inside it, exactly as a scoped GetHierarchy would return it as the root)."""
    src = fixtures_dir / f"hierarchy_hsPages__{_sanitize(NOTEBOOK_ID)}.xml"
    dst = tmp_path / f"hierarchy_hsPages__{_sanitize(SECTION_ID)}.xml"
    dst.write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
    return FixtureBackend(tmp_path)


def _sent_section_order(be: FixtureBackend) -> list[str]:
    writes = [c for c in be.calls if c.method == "update_hierarchy"]
    assert len(writes) == 1, "a block reposition must be exactly ONE UpdateHierarchy"
    tree = etree.fromstring(writes[0].kwargs["changes_xml"].encode("utf-8"))
    section = next(el for el in tree.iter(qn("Section")) if el.get("ID") == SECTION_ID)
    return [p.get("ID") for p in section.findall(qn("Page"))]


# --- reposition_pages: the block-placement core ---------------------------------------


def test_block_placed_after_an_anchor_keeps_block_order(section_be):
    # move [圖片頁, 混合樣式頁] (note: NOT their original relative order) to right after 表格頁
    hierarchy_edit.reposition_pages(
        section_be, SECTION_ID, [P_IMAGE, P_MIXED], after_page_id=P_TABLE
    )
    assert _sent_section_order(section_be) == [
        P_TABLE,
        P_IMAGE,
        P_MIXED,
        P_MULTI,
        P_ROLE,
        P_DUTY,
        P_COMM,
        P_TIME,
    ]


def test_block_with_no_anchor_goes_to_the_top(section_be):
    hierarchy_edit.reposition_pages(section_be, SECTION_ID, [P_IMAGE, P_MIXED])
    order = _sent_section_order(section_be)
    assert order[:2] == [P_IMAGE, P_MIXED]
    assert set(order) == set(ORIGINAL_ORDER)  # nothing dropped/invented


def test_block_preserves_the_full_page_set(section_be):
    hierarchy_edit.reposition_pages(section_be, SECTION_ID, [P_MULTI], after_page_id=P_MIXED)
    order = _sent_section_order(section_be)
    assert set(order) == set(ORIGINAL_ORDER)
    assert order.index(P_MULTI) == order.index(P_MIXED) + 1


def test_anchor_inside_the_moved_block_is_rejected(section_be):
    with pytest.raises(ValueError, match="cannot be one of the pages being moved"):
        hierarchy_edit.reposition_pages(
            section_be, SECTION_ID, [P_IMAGE, P_MIXED], after_page_id=P_IMAGE
        )


def test_absent_anchor_raises_node_not_found(section_be):
    with pytest.raises(NodeNotFoundError):
        hierarchy_edit.reposition_pages(
            section_be, SECTION_ID, [P_IMAGE], after_page_id="{NOPE}{1}{E0}"
        )


def test_empty_page_ids_is_rejected(section_be):
    with pytest.raises(ValueError, match="empty"):
        hierarchy_edit.reposition_pages(section_be, SECTION_ID, [])


# --- subtree_page_ids: OneNote's positional subpage model ------------------------------


def test_subtree_of_l1_page_includes_all_following_deeper_pages(section_be):
    # 單節多頁 (L1) owns the L2/L3 run that follows it, to the end of the section
    assert copy.subtree_page_ids(section_be, SECTION_ID, P_MULTI) == [
        P_MULTI,
        P_ROLE,
        P_DUTY,
        P_COMM,
        P_TIME,
    ]


def test_subtree_stops_at_the_next_same_level_page(section_be):
    # 混合樣式頁 (L1) is followed immediately by another L1 → it has no subpages
    assert copy.subtree_page_ids(section_be, SECTION_ID, P_MIXED) == [P_MIXED]


def test_subtree_of_a_subpage_takes_only_its_deeper_children(section_be):
    # 測試頁面7 (L2) owns its two L3 children but stops at the next L2 (測試頁面4)
    assert copy.subtree_page_ids(section_be, SECTION_ID, P_ROLE) == [P_ROLE, P_DUTY, P_COMM]


def test_subtree_of_a_missing_page_raises(section_be):
    with pytest.raises(NodeNotFoundError):
        copy.subtree_page_ids(section_be, SECTION_ID, "{NOPE}{1}{E0}")


# --- copy_pages engine: clone each in order, aggregate notes ---------------------------


def test_copy_pages_clones_each_in_order_and_aggregates(monkeypatch):
    seen = []

    def fake_transfer_page(backend, page_id, target_section_id):
        seen.append((page_id, target_section_id))
        return SimpleNamespace(
            page_id=f"COPY-{page_id}",
            name=f"name-{page_id}",
            page_level=2,
            file_notes=[f"note-{page_id}"],
            missing_images=1,
            missing_files=0,
            missing_objects=2,
        )

    monkeypatch.setattr(copy, "transfer_page", fake_transfer_page)
    result = copy.copy_pages(object(), ["A", "B", "C"], "SEC")
    assert seen == [("A", "SEC"), ("B", "SEC"), ("C", "SEC")]
    assert result.page_ids == ["COPY-A", "COPY-B", "COPY-C"]
    # pages carries name + level so callers report by name, not raw ids
    assert result.pages == [
        {"page_id": "COPY-A", "name": "name-A", "page_level": 2},
        {"page_id": "COPY-B", "name": "name-B", "page_level": 2},
        {"page_id": "COPY-C", "name": "name-C", "page_level": 2},
    ]
    assert result.file_notes == ["note-A", "note-B", "note-C"]
    assert (result.missing_images, result.missing_objects) == (3, 6)


def test_copy_pages_rejects_duplicates():
    with pytest.raises(ValueError, match="duplicate"):
        copy.copy_pages(object(), ["A", "A"], "SEC")


# --- server facades: orchestration (patched, no backend) -------------------------------


@pytest.fixture
def facade_calls(monkeypatch):
    recorded: dict = {}

    def fake_copy_pages(backend, page_ids, target_section_id):
        recorded["copy_pages"] = {"ids": list(page_ids), "target": target_section_id}
        return SimpleNamespace(
            page_ids=[f"COPY-{p}" for p in page_ids],
            pages=[
                {"page_id": f"COPY-{p}", "name": f"name-{p}", "page_level": 1} for p in page_ids
            ],
            missing_images=0,
            missing_files=0,
            missing_objects=0,
            file_notes=[],
        )

    def fake_subtree(backend, section_id, page_id):
        recorded["subtree"] = {"section": section_id, "page": page_id}
        return [page_id, f"{page_id}-sub1", f"{page_id}-sub2"]

    def fake_reposition_pages(backend, section_id, page_ids, *, after_page_id=""):
        recorded["reposition"] = {
            "section": section_id,
            "ids": list(page_ids),
            "after": after_page_id,
        }
        if after_page_id == "ABSENT":  # stand-in for "anchor not in the target section"
            raise NodeNotFoundError("anchor not in this section")

    monkeypatch.setattr(server, "get_backend", lambda: object())
    monkeypatch.setattr(server.copy, "copy_pages", fake_copy_pages)
    monkeypatch.setattr(server.copy, "subtree_page_ids", fake_subtree)
    monkeypatch.setattr(server.copy, "sync_warning", lambda *a, **k: None)
    monkeypatch.setattr(server.hierarchy_edit, "reposition_pages", fake_reposition_pages)
    return recorded


def test_copy_pages_facade_no_anchor_leaves_block_at_end(facade_calls):
    out = json.loads(server.copy_pages(["A", "B"], "SEC"))
    # facade reports pages (id + name + level), not a bare id list
    assert [p["page_id"] for p in out["pages"]] == ["COPY-A", "COPY-B"]
    assert [p["name"] for p in out["pages"]] == ["name-A", "name-B"]
    assert "reposition" not in facade_calls  # no placement step when no anchor given


def test_copy_pages_facade_places_block_after_explicit_anchor(facade_calls):
    server.copy_pages(["A", "B"], "SEC", after_page_id="LOCAL")
    assert facade_calls["reposition"] == {
        "section": "SEC",
        "ids": ["COPY-A", "COPY-B"],
        "after": "LOCAL",
    }


def test_copy_pages_facade_explicit_missing_anchor_raises(facade_calls):
    with pytest.raises(NodeNotFoundError):
        server.copy_pages(["A"], "SEC", after_page_id="ABSENT")


def test_subtree_facade_defaults_below_the_source_subtree(facade_calls):
    # same-section copy, no anchor → block lands right after the source subtree's LAST page
    out = json.loads(server.copy_page_subtree("SEC", "ITIN"))
    assert facade_calls["subtree"] == {"section": "SEC", "page": "ITIN"}
    assert facade_calls["copy_pages"]["target"] == "SEC"
    assert facade_calls["reposition"]["after"] == "ITIN-sub2"  # last page of the source subtree
    assert [p["page_id"] for p in out["pages"]] == [
        "COPY-ITIN",
        "COPY-ITIN-sub1",
        "COPY-ITIN-sub2",
    ]


def test_subtree_facade_explicit_anchor_wins(facade_calls):
    # the headline case: "copy ●ITIN and its subpages below ●Local"
    server.copy_page_subtree("SEC", "ITIN", after_page_id="LOCAL")
    assert facade_calls["reposition"]["after"] == "LOCAL"


def test_subtree_facade_cross_section_lands_at_end(facade_calls):
    # different target section → no same-section "below source" default, no placement step
    server.copy_page_subtree("SEC", "ITIN", target_section_id="OTHER")
    assert facade_calls["copy_pages"]["target"] == "OTHER"
    assert "reposition" not in facade_calls


def test_subtree_facade_explicit_missing_anchor_raises(facade_calls):
    with pytest.raises(NodeNotFoundError):
        server.copy_page_subtree("SEC", "ITIN", after_page_id="ABSENT")
