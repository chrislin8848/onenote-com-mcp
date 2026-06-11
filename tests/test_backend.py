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


def test_current_window_ids_from_fixture(tmp_path):
    (tmp_path / "current_window.json").write_text(
        '{"notebook_id": "{NB}{1}{B0}", "section_group_id": null, '
        '"section_id": "{SEC}{1}{B0}", "page_id": "{P}{1}{B0}"}',
        encoding="utf-8",
    )
    ids = FixtureBackend(tmp_path).get_current_window_ids()
    assert ids.notebook_id == "{NB}{1}{B0}"
    assert ids.section_group_id is None  # current section sits directly in the notebook
    assert ids.page_id == "{P}{1}{B0}"


def test_current_window_no_window_is_a_clear_error(tmp_path):
    from onenote_com_mcp.errors import NoCurrentWindowError

    (tmp_path / "current_window.json").write_text("null", encoding="utf-8")
    with pytest.raises(NoCurrentWindowError):
        FixtureBackend(tmp_path).get_current_window_ids()


def test_signatures_present_on_concrete_backend():
    # Cheap guard that the concrete backend keeps the documented method set.
    for name in ("get_hierarchy", "get_page_content", "update_page_content", "find_pages"):
        assert callable(getattr(FixtureBackend, name))
        assert inspect.signature(getattr(FixtureBackend, name))


# --- Win32ComBackend._call error mapping (no COM needed: the connection is lazy) --------


def _win32_backend():
    from onenote_com_mcp.backend.win32com_backend import Win32ComBackend

    return Win32ComBackend()  # __init__ touches no COM; only .app would


def test_call_maps_concurrency_hresult_to_concurrency_error():
    from onenote_com_mcp.errors import ConcurrencyError

    be = _win32_backend()

    def stale_write():
        raise Exception(0x80042010 - 0x1_0000_0000, "hrLastModifiedDateDidNotMatch")

    with pytest.raises(ConcurrencyError, match="changed since it was read"):
        be._call("UpdatePageContent", stale_write)


def test_call_wraps_other_com_errors():
    from onenote_com_mcp.errors import OneNoteComError

    be = _win32_backend()

    def boom():
        raise Exception(-2147213312, "some other failure")

    with pytest.raises(OneNoteComError):
        be._call("GetPageContent", boom)


def test_call_unwraps_excepinfo_scode():
    # a real OneNote refusal arrives as DISP_E_EXCEPTION with the actual code (here
    # hrLastModifiedDateDidNotMatch) buried in excepinfo's scode — VM ground truth
    from onenote_com_mcp.errors import ConcurrencyError

    be = _win32_backend()

    def stale_write():
        raise Exception(
            -2147352567,  # DISP_E_EXCEPTION
            "exception occurred",
            (0, None, None, None, 0, 0x80042010 - 0x1_0000_0000),
            None,
        )

    with pytest.raises(ConcurrencyError):
        be._call("UpdatePageContent", stale_write)
