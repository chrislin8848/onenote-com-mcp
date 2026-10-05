"""The displayed-page guard in apply_page_edit (the single write core).

VM ground truth 2026-10-05: a single-cell edit of a 4,134-cell table took 134-141s while its page
was DISPLAYED in OneNote vs 23-26s while another page was — OneNote redraws every cell. So a BIG
write (>= _BIG_WRITE_OES paragraphs in the payload) to the page currently on screen is refused with
PageDisplayedError BEFORE anything is written, unless allow_displayed. The guard never blocks a
small write, a write to another page, or a write when the window can't be read.
"""

from __future__ import annotations

import json

import pytest
from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend
from onenote_com_mcp.errors import PageDisplayedError
from onenote_com_mcp.service import page_edit, read

_PARSER = etree.XMLParser(strip_cdata=False)
TABLE_PAGE = "表格頁"


def _page_id(fixtures_dir, name: str) -> str:
    for path in sorted(fixtures_dir.glob("page_*.xml")):
        if "__binary" in path.name:
            continue
        root = etree.fromstring(path.read_bytes(), parser=_PARSER)
        if root.get("name") == name:
            return root.get("ID")
    raise AssertionError(f"no fixture page named {name!r}")


def _backend(fixtures_dir, tmp_path, current_page: str | None) -> FixtureBackend:
    for src in fixtures_dir.iterdir():
        if src.is_file():
            (tmp_path / src.name).write_bytes(src.read_bytes())
    window = None if current_page is None else {"page_id": current_page}
    (tmp_path / "current_window.json").write_text(json.dumps(window), encoding="utf-8")
    return FixtureBackend(tmp_path)


def _writes(be: FixtureBackend) -> int:
    return sum(1 for c in be.calls if c.method == "update_page_content")


def _edit_one_cell(be, pid, **kwargs):
    page = read.get_page(be, pid)
    blocks = [b for o in page["outlines"] for b in o["blocks"]]
    table = next(b for b in blocks if b["type"] == "table")
    oe = table["rows"][1][1]["paragraphs"][0]["object_id"]
    page_edit.edit_page_content(be, pid, "改", "replace", target_object_id=oe, **kwargs)


@pytest.fixture
def big(monkeypatch):
    # the 表格頁 fixture's table is 10×2; make it count as "big" for the guard
    monkeypatch.setattr(page_edit, "_BIG_WRITE_OES", 5)


def test_big_write_to_the_displayed_page_is_refused_before_writing(fixtures_dir, tmp_path, big):
    pid = _page_id(fixtures_dir, TABLE_PAGE)
    be = _backend(fixtures_dir, tmp_path, current_page=pid)
    with pytest.raises(PageDisplayedError, match="switch OneNote to any OTHER page"):
        _edit_one_cell(be, pid)
    assert _writes(be) == 0


def test_allow_displayed_overrides(fixtures_dir, tmp_path, big):
    pid = _page_id(fixtures_dir, TABLE_PAGE)
    be = _backend(fixtures_dir, tmp_path, current_page=pid)
    _edit_one_cell(be, pid, allow_displayed=True)
    assert _writes(be) == 1


def test_big_write_proceeds_when_another_page_is_shown(fixtures_dir, tmp_path, big):
    pid = _page_id(fixtures_dir, TABLE_PAGE)
    be = _backend(fixtures_dir, tmp_path, current_page="{SOME-OTHER-PAGE}{1}{E1}")
    _edit_one_cell(be, pid)
    assert _writes(be) == 1


def test_big_write_proceeds_when_no_window_can_be_read(fixtures_dir, tmp_path, big):
    pid = _page_id(fixtures_dir, TABLE_PAGE)
    be = _backend(fixtures_dir, tmp_path, current_page=None)  # replays NoCurrentWindowError
    _edit_one_cell(be, pid)
    assert _writes(be) == 1


def test_small_write_never_checks_the_window(fixtures_dir, tmp_path):
    # default threshold (1,000): a 10×2 table is small — written even while displayed, and the
    # current window is not even read
    pid = _page_id(fixtures_dir, TABLE_PAGE)
    be = _backend(fixtures_dir, tmp_path, current_page=pid)
    window_reads = []
    real = be.get_current_window_ids
    be.get_current_window_ids = lambda: window_reads.append(1) or real()
    _edit_one_cell(be, pid)
    assert _writes(be) == 1
    assert window_reads == []


def test_facades_pass_allow_displayed_through(fixtures_dir, tmp_path, big):
    pid = _page_id(fixtures_dir, TABLE_PAGE)
    be = _backend(fixtures_dir, tmp_path, current_page=pid)
    table = next(
        b for o in read.get_page(be, pid)["outlines"] for b in o["blocks"] if b["type"] == "table"
    )
    with pytest.raises(PageDisplayedError):
        page_edit.find_and_replace(be, pid, "DAY", "D", object_id=table["object_id"])
    page_edit.find_and_replace(
        be, pid, "DAY", "D", object_id=table["object_id"], allow_displayed=True
    )
    page_edit.modify_table(
        be, pid, table["object_id"], "set_column", at_index=1, values=["x"], allow_displayed=True
    )
    page_edit.batch_update(
        be,
        pid,
        [{"op": "find_replace", "find": "D", "replace": "DAY", "object_id": table["object_id"]}],
        allow_displayed=True,
    )
    assert _writes(be) == 3
