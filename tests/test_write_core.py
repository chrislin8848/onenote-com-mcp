"""Guard tests for the write-core convergence invariant (SPEC §4).

These assert STRUCTURE, not OneNote XML semantics: exactly one UpdatePageContent per edit, the
concurrency guard is carried, and all write facades funnel through the single core. The inline
page XML below is a minimal *control-flow input* for the orchestration seam — NOT a parse-layer
fixture (the parse/build layer is TDD'd separately against real VM dumps).
"""

from __future__ import annotations

import datetime as dt

import pytest

from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize
from onenote_com_mcp.service import page_edit

_PAGE_ID = "{P}{1}{B0}"
_MINIMAL_PAGE = (
    '<?xml version="1.0"?>'
    '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
    'ID="{P}{1}{B0}" lastModifiedTime="2026-06-10T17:39:30.000Z">'
    "<one:Outline><one:OEChildren/></one:Outline></one:Page>"
)


def _backend_with_page(tmp_path) -> FixtureBackend:
    (tmp_path / f"page_{_sanitize(_PAGE_ID)}.xml").write_text(_MINIMAL_PAGE, encoding="utf-8")
    return FixtureBackend(tmp_path)


def test_apply_page_edit_makes_exactly_one_guarded_write(tmp_path):
    be = _backend_with_page(tmp_path)

    def mutate(tree):
        tree.set("name", "edited")  # any in-place change

    page_edit.apply_page_edit(be, _PAGE_ID, mutate)

    writes = [c for c in be.calls if c.method == "update_page_content"]
    assert len(writes) == 1, "a content edit must be exactly ONE UpdatePageContent"
    assert writes[0].kwargs["force"] is False, "must never force by default"
    assert writes[0].kwargs["expected_last_modified"] == dt.datetime(
        2026, 6, 10, 17, 39, 30, tzinfo=dt.UTC
    ), "concurrency guard must carry the page's lastModifiedTime"
    assert "edited" in writes[0].kwargs["changes_xml"], "the in-place mutation must be sent"


def test_parse_onenote_datetime_handles_z_suffix():
    assert page_edit.parse_onenote_datetime("2026-06-10T17:39:30.000Z") == dt.datetime(
        2026, 6, 10, 17, 39, 30, tzinfo=dt.UTC
    )
    assert page_edit.parse_onenote_datetime(None) is None
    assert page_edit.parse_onenote_datetime("not-a-date") is None


@pytest.mark.parametrize(
    "invoke",
    [
        lambda be: page_edit.edit_page_content(be, _PAGE_ID, "hi"),
        lambda be: page_edit.add_table(be, _PAGE_ID, [["a", "b"]]),
        lambda be: page_edit.insert_image(be, _PAGE_ID, "BASE64", "image/png"),
    ],
    ids=["update_page_content", "create_table", "insert_image"],
)
def test_all_write_facades_delegate_to_single_core(tmp_path, monkeypatch, invoke):
    be = _backend_with_page(tmp_path)
    seen: dict[str, str] = {}

    def spy(backend, page_id, mutate, **kwargs):  # replaces the real core
        seen["page_id"] = page_id

    monkeypatch.setattr(page_edit, "apply_page_edit", spy)
    invoke(be)
    assert seen.get("page_id") == _PAGE_ID, "every write facade must route through apply_page_edit"
