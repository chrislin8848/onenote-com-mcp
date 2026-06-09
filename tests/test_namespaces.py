"""Namespace helpers."""

from __future__ import annotations

from onenote_mcp.xmllayer.namespaces import ONE_NS, local_name, qn


def test_one_ns_is_2013():
    assert ONE_NS == "http://schemas.microsoft.com/office/onenote/2013/onenote"


def test_qn_roundtrip():
    assert qn("Outline") == f"{{{ONE_NS}}}Outline"
    assert local_name(qn("Page")) == "Page"
    assert local_name("Page") == "Page"
