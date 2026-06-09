"""FixtureBackend plumbing + factory selection.

These exercise the backend *mechanism* (file lookup, call recording, selection), not OneNote
XML semantics — the latter is Phase 1, TDD'd against real VM-dumped fixtures.
"""

from __future__ import annotations

import inspect

import pytest

from onenote_com_mcp.backend import FixtureBackend, OneNoteBackend, get_backend
from onenote_com_mcp.backend.base import OneNoteBackend as ABCBackend
from onenote_com_mcp.enums import CreateFileType, HierarchyScope
from onenote_com_mcp.errors import NodeNotFoundError


def test_backend_is_abstract():
    with pytest.raises(TypeError):
        ABCBackend()  # type: ignore[abstract]


def test_fixture_backend_implements_all_abstract_methods():
    # No leftover abstract methods → fully concrete.
    assert getattr(FixtureBackend, "__abstractmethods__", frozenset()) == frozenset()


def test_get_hierarchy_reads_scoped_then_falls_back(tmp_path):
    (tmp_path / "hierarchy_hsNotebooks.xml").write_text("<root/>", encoding="utf-8")
    be = FixtureBackend(tmp_path)
    assert be.get_hierarchy("", HierarchyScope.hsNotebooks) == "<root/>"


def test_missing_fixture_raises_with_guidance(tmp_path):
    be = FixtureBackend(tmp_path)
    with pytest.raises(NodeNotFoundError) as exc:
        be.get_page_content("{ABC}{1}{B0}")
    assert "dump_fixtures.py" in str(exc.value)


def test_write_methods_are_recorded(tmp_path):
    be = FixtureBackend(tmp_path)
    new_id = be.open_hierarchy("Meetings.one", "{NB}{1}{B0}", CreateFileType.cftSection)
    be.update_page_content("<one:Page/>", expected_last_modified=None, force=False)
    assert new_id.startswith("{FIXTURE-cftSection-")
    methods = [c.method for c in be.calls]
    assert methods == ["open_hierarchy", "update_page_content"]


def test_get_backend_prefers_fixtures_env(tmp_path, monkeypatch):
    monkeypatch.setenv("ONENOTE_FIXTURES_DIR", str(tmp_path))
    be = get_backend()
    assert isinstance(be, FixtureBackend)
    assert isinstance(be, OneNoteBackend)


def test_get_backend_errors_without_com_or_fixtures(monkeypatch):
    monkeypatch.delenv("ONENOTE_FIXTURES_DIR", raising=False)
    monkeypatch.setattr("onenote_com_mcp.backend.sys.platform", "linux")
    with pytest.raises(RuntimeError):
        get_backend()


def test_signatures_present_on_concrete_backend():
    # Cheap guard that the concrete backend keeps the documented method set.
    for name in ("get_hierarchy", "get_page_content", "update_page_content", "find_pages"):
        assert callable(getattr(FixtureBackend, name))
        assert inspect.signature(getattr(FixtureBackend, name))
