"""Guard tests for the hierarchy-restructure discipline (SPEC §5).

Structure, not OneNote semantics: every structural change goes through ONE update_hierarchy
call site, submits the COMPLETE child list (node-ID conservation — no drops, no inventions),
and all four facades route through the single core. The inline hierarchy XML is a control-flow
input for the orchestration seam, NOT a parse-layer fixture (those come from real VM dumps).
"""

from __future__ import annotations

import pytest

from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize
from onenote_com_mcp.enums import HierarchyScope
from onenote_com_mcp.service import hierarchy_edit

_SEC_ID = "{SEC}{1}{B0}"
_HIERARCHY = (
    '<?xml version="1.0"?>'
    '<one:Section xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
    'ID="{SEC}{1}{B0}" name="S">'
    '<one:Page ID="{P1}{1}{B0}" name="a" pageLevel="1"/>'
    '<one:Page ID="{P2}{1}{B0}" name="b" pageLevel="2"/>'
    '<one:Page ID="{P3}{1}{B0}" name="c" pageLevel="1"/>'
    "</one:Section>"
)


def _backend_with_hierarchy(tmp_path) -> FixtureBackend:
    name = f"hierarchy_hsPages__{_sanitize(_SEC_ID)}.xml"
    (tmp_path / name).write_text(_HIERARCHY, encoding="utf-8")
    return FixtureBackend(tmp_path)


def test_restructure_makes_exactly_one_whole_batch_write(tmp_path):
    be = _backend_with_hierarchy(tmp_path)

    def mutate(tree):
        tree.append(tree[0])  # move first page to the end — reorder, same node set

    hierarchy_edit.apply_hierarchy_restructure(be, _SEC_ID, HierarchyScope.hsPages, mutate)

    writes = [c for c in be.calls if c.method == "update_hierarchy"]
    assert len(writes) == 1, "a structural change must be exactly ONE UpdateHierarchy"
    xml = writes[0].kwargs["changes_xml"]
    for pid in ("{P1}{1}{B0}", "{P2}{1}{B0}", "{P3}{1}{B0}"):
        assert pid in xml, "the COMPLETE child list must be submitted, not a partial one"
    assert xml.index("{P2}") < xml.index("{P3}") < xml.index("{P1}"), "target order preserved"


def test_dropping_a_node_is_refused(tmp_path):
    be = _backend_with_hierarchy(tmp_path)

    def drops_one(tree):
        tree.remove(tree[1])  # partial list — exactly what OneNote would mis-infer

    with pytest.raises(ValueError, match="complete child list"):
        hierarchy_edit.apply_hierarchy_restructure(be, _SEC_ID, HierarchyScope.hsPages, drops_one)
    assert not [c for c in be.calls if c.method == "update_hierarchy"], "must not write"


def test_inventing_a_node_is_refused(tmp_path):
    be = _backend_with_hierarchy(tmp_path)

    def invents_one(tree):
        import copy as _copy

        clone = _copy.deepcopy(tree[0])
        clone.set("ID", "{P9}{1}{B0}")
        tree.append(clone)

    with pytest.raises(ValueError, match="complete child list"):
        hierarchy_edit.apply_hierarchy_restructure(be, _SEC_ID, HierarchyScope.hsPages, invents_one)


_NB_ID = "{NB}{1}{B0}"
_MIXED_HIERARCHY = (
    '<?xml version="1.0"?>'
    '<one:Notebook xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
    'ID="{NB}{1}{B0}" name="N">'
    '<one:Section ID="{S1}{1}{B0}" name="a"/>'
    '<one:SectionGroup ID="{SG}{1}{B0}" name="g">'
    '<one:Section ID="{S2}{1}{B0}" name="b"/>'
    "</one:SectionGroup>"
    '<one:Section ID="{S3}{1}{B0}" name="c"/>'
    "</one:Notebook>"
)


def test_mixed_section_and_group_children_must_both_survive(tmp_path):
    # SPEC §5: a notebook's direct children are a MIXED Section + SectionGroup list; the
    # whole batch must contain both kinds. Flattening (dropping the group) is refused.
    name = f"hierarchy_hsSections__{_sanitize(_NB_ID)}.xml"
    (tmp_path / name).write_text(_MIXED_HIERARCHY, encoding="utf-8")
    be = FixtureBackend(tmp_path)

    def flattens_the_group(tree):
        group = tree[1]
        tree.remove(group)  # drops the group AND its nested section

    with pytest.raises(ValueError, match="complete child list"):
        hierarchy_edit.apply_hierarchy_restructure(
            be, _NB_ID, HierarchyScope.hsSections, flattens_the_group
        )

    # A pure reorder that keeps both kinds goes through, group nesting intact.
    def reorders_keeping_both(tree):
        tree.append(tree[0])  # section "a" to the end; group stays with its child

    hierarchy_edit.apply_hierarchy_restructure(
        be, _NB_ID, HierarchyScope.hsSections, reorders_keeping_both
    )
    xml = next(c for c in be.calls if c.method == "update_hierarchy").kwargs["changes_xml"]
    assert "SectionGroup" in xml and "{S2}" in xml, "group + nested section submitted"
    assert xml.index("{SG}") < xml.index("{S1}"), "reorder reflected"


@pytest.mark.parametrize(
    "invoke",
    [
        lambda be: hierarchy_edit.restructure_section(be, _SEC_ID, [{"page_id": "x"}]),
        lambda be: hierarchy_edit.reorder_sections(be, _SEC_ID, ["a", "b"]),
        lambda be: hierarchy_edit.rename_node(be, _SEC_ID, "{P1}{1}{B0}", "new"),
        lambda be: hierarchy_edit.move_page(be, _SEC_ID, "{P1}{1}{B0}", "{SEC2}{1}{B0}"),
    ],
    ids=["restructure_section", "reorder_sections", "rename_node", "move_page"],
)
def test_all_structure_facades_delegate_to_single_core(tmp_path, monkeypatch, invoke):
    be = _backend_with_hierarchy(tmp_path)
    seen: dict[str, str] = {}

    def spy(backend, scope_id, scope, mutate):
        seen["scope_id"] = scope_id

    monkeypatch.setattr(hierarchy_edit, "apply_hierarchy_restructure", spy)
    invoke(be)
    assert seen.get("scope_id") == _SEC_ID, (
        "every structure facade must route through apply_hierarchy_restructure"
    )
