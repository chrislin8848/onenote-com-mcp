"""Tier-1 tests for copy_page's and create_page's DEFAULT placement (the facade orchestration
in server.py), exercising the logic without a backend by patching the composed service calls.

copy_page = copy.transfer_page (lands at the section END) + hierarchy_edit.reposition_page:
  * no after_page_id → placed right BELOW the SOURCE page (the common "copy this page" case);
  * an explicit after_page_id overrides it;
  * a cross-section copy (anchor not in the target section) falls back to the end, but only for
    the IMPLICIT default — an explicitly named missing anchor is a real error.

create_page = (read current window) + create.create_page (lands at END) + reposition_page:
  * no after_page_id → placed right BELOW the page the user is currently on (get_current_context);
  * no open window, or the current page is in another section → append at the END;
  * an explicit after_page_id overrides; an explicit missing anchor is a real error.
"""

from __future__ import annotations

import json
from types import SimpleNamespace

import pytest

from onenote_com_mcp import server
from onenote_com_mcp.backend.base import CurrentWindowIds
from onenote_com_mcp.errors import NoCurrentWindowError, NodeNotFoundError


@pytest.fixture
def calls(monkeypatch):
    recorded: dict = {}

    def fake_transfer_page(backend, page_id, target_section_id):
        recorded["transfer"] = (page_id, target_section_id)
        return SimpleNamespace(
            page_id="COPY-ID",
            missing_images=[],
            missing_files=[],
            missing_objects=[],
            file_notes=[],
        )

    def fake_reposition_page(backend, section_id, page_id, *, after_page_id="", page_level=None):
        recorded["reposition"] = {"section": section_id, "page": page_id, "after": after_page_id}
        if after_page_id == "ABSENT":  # stand-in for "anchor not in the target section"
            raise NodeNotFoundError("anchor not in this section")

    monkeypatch.setattr(server, "get_backend", lambda: object())
    monkeypatch.setattr(server.copy, "transfer_page", fake_transfer_page)
    monkeypatch.setattr(server.copy, "sync_warning", lambda *a, **k: None)
    monkeypatch.setattr(server.hierarchy_edit, "reposition_page", fake_reposition_page)
    return recorded


def test_default_places_copy_below_the_source(calls):
    out = json.loads(server.copy_page("SRC", "SEC"))
    assert out["page_id"] == "COPY-ID"
    # default anchor = the source page → the copy is moved to right after SRC
    assert calls["reposition"] == {"section": "SEC", "page": "COPY-ID", "after": "SRC"}


def test_explicit_after_page_id_overrides_the_default(calls):
    server.copy_page("SRC", "SEC", after_page_id="OTHER")
    assert calls["reposition"]["after"] == "OTHER"


def test_cross_section_default_falls_back_to_section_end(calls):
    # source not in the target section → the IMPLICIT reposition is swallowed; the copy stays last
    out = json.loads(server.copy_page("ABSENT", "SEC"))
    assert out["page_id"] == "COPY-ID", "copy still succeeds; it simply stays at the section end"


def test_explicitly_named_missing_anchor_is_a_real_error(calls):
    with pytest.raises(NodeNotFoundError):
        server.copy_page("SRC", "SEC", after_page_id="ABSENT")


# --- create_page: default below the current page -------------------------------------


@pytest.fixture
def create_calls(monkeypatch):
    recorded: dict = {}

    class FakeBackend:
        def __init__(self, current_page):
            self._current = current_page

        def get_current_window_ids(self):
            if self._current == "NO_WINDOW":
                raise NoCurrentWindowError("no window open")
            return CurrentWindowIds(
                notebook_id=None, section_group_id=None, section_id=None, page_id=self._current
            )

    def fake_create_page(backend, section_id, title, content, page_level):
        recorded["create"] = (section_id, title, page_level)
        return "NEW-ID"

    def fake_reposition_page(backend, section_id, page_id, *, after_page_id="", page_level=None):
        recorded["reposition"] = {"section": section_id, "page": page_id, "after": after_page_id}
        if after_page_id == "OTHER_SECTION":  # stand-in for "anchor not in this section"
            raise NodeNotFoundError("anchor not in this section")

    monkeypatch.setattr(server.create, "create_page", fake_create_page)
    monkeypatch.setattr(server.hierarchy_edit, "reposition_page", fake_reposition_page)

    def configure(current_page):
        monkeypatch.setattr(server, "get_backend", lambda: FakeBackend(current_page))
        return recorded

    return configure


def test_create_page_default_places_below_current_page(create_calls):
    rec = create_calls(current_page="CUR")
    out = json.loads(server.create_page("SEC", "標題"))
    assert out["page_id"] == "NEW-ID"
    assert rec["reposition"] == {"section": "SEC", "page": "NEW-ID", "after": "CUR"}


def test_create_page_explicit_after_page_id_overrides_current(create_calls):
    rec = create_calls(current_page="CUR")
    server.create_page("SEC", "標題", after_page_id="PINNED")
    assert rec["reposition"]["after"] == "PINNED"


def test_create_page_no_open_window_appends_at_end(create_calls):
    rec = create_calls(current_page="NO_WINDOW")
    out = json.loads(server.create_page("SEC", "標題"))
    assert out["page_id"] == "NEW-ID"
    assert "reposition" not in rec, "no current page → no reposition, stays at the end"


def test_create_page_current_page_in_other_section_appends_at_end(create_calls):
    create_calls(current_page="OTHER_SECTION")  # reposition raises NodeNotFoundError (implicit)
    out = json.loads(server.create_page("SEC", "標題"))
    assert out["page_id"] == "NEW-ID", "implicit anchor not in this section → stays at the end"


def test_create_page_explicit_missing_anchor_raises(create_calls):
    create_calls(current_page="CUR")
    with pytest.raises(NodeNotFoundError):
        server.create_page("SEC", "標題", after_page_id="OTHER_SECTION")
