"""Guard tests for the copy direct-path invariant (SPEC §5).

Copy must transfer RAW page XML (binary inlined) and must NOT route through the structured
parse / get_page representation. Two guards: a behavioral one (the read uses piBinaryData) and a
structural one (the copy module does not depend on the parse layer).
"""

from __future__ import annotations

import inspect

from onenote_com_mcp.backend.fixture import FixtureBackend
from onenote_com_mcp.enums import PageInfo
from onenote_com_mcp.service import copy


def test_copy_module_does_not_import_structured_parser():
    src = inspect.getsource(copy)
    assert "parse_page" not in src, "copy must not call the structured parser"
    assert "import parse" not in src, "copy must not import the parse layer"
    assert "xmllayer.parse" not in src


def test_transfer_page_reads_raw_binary_not_structured(tmp_path, monkeypatch):
    be = FixtureBackend(tmp_path)
    reads: list[PageInfo] = []

    def fake_get(page_id, page_info=PageInfo.piBasic):
        reads.append(page_info)
        return (
            '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
            'ID="{P}{1}{B0}" lastModifiedTime="2026-06-11T08:00:00.000Z"/>'
        )

    monkeypatch.setattr(be, "get_page_content", fake_get)
    copy.transfer_page(be, "{P}{1}{B0}", "{SEC}{1}{B0}")

    assert reads[0] == PageInfo.piBinaryData, (
        "the SOURCE read must be raw XML with binary inlined, not piBasic / structured parse"
    )
    assert [c.method for c in be.calls if c.method == "update_page_content"] == [
        "update_page_content"
    ], "the transplant is ONE UpdatePageContent"
