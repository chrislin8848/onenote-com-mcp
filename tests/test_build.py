"""Builders for NEW content (Phase 1 side of Phase 4 write tools).

Round-trip property: whatever ``build_*`` emits must parse back through our own parse
layer with the same text/styles (semantic self-consistency — SPEC §5). Editing existing
pages is NOT done with these (that's in-place tree mutation in service/page_edit.py).
"""

from __future__ import annotations

import base64

from lxml import etree

from onenote_com_mcp.xmllayer.build import build_image_xml, build_table_xml, build_text_oe_xml
from onenote_com_mcp.xmllayer.namespaces import ONE_NS, qn
from onenote_com_mcp.xmllayer.parse import parse_image, parse_oe, parse_table

PNG_1PX = base64.b64encode(
    base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGBgAAAABQAB"
        "h6FO1AAAAABJRU5ErkJggg=="
    )
).decode()


def _fromstring(xml: str) -> etree._Element:
    el = etree.fromstring(xml.encode("utf-8"))
    assert el.tag.startswith(f"{{{ONE_NS}}}")  # fragment is namespaced, parseable standalone
    return el


# --- build_text_oe_xml ----------------------------------------------------------------


def test_text_oe_roundtrip():
    xml = build_text_oe_xml(
        [
            {"text": "plain "},
            {"text": "bold", "style": {"font-weight": "bold"}},
            {"text": " marked", "style": {"background": "yellow"}},
        ],
        quick_style_index=1,
        alignment="left",
    )
    el = _fromstring(xml)
    assert el.tag == qn("OE")
    assert el.get("quickStyleIndex") == "1"
    para = parse_oe(el, {})
    assert para.text == "plain bold marked"
    assert para.runs[1].span_style == {"font-weight": "bold"}
    # dual-attribute highlight written (SPEC §5)
    assert para.runs[2].span_style["background"] == "yellow"
    assert para.runs[2].span_style["mso-highlight"] == "yellow"


def test_text_oe_uses_cdata():
    xml = build_text_oe_xml([{"text": "hi", "style": {"font-weight": "bold"}}])
    assert "<![CDATA[" in xml


def test_text_oe_accepts_plain_string_runs():
    para = parse_oe(_fromstring(build_text_oe_xml(["just text"])), {})
    assert para.text == "just text"
    assert para.runs[0].span_style == {}


# --- build_table_xml ------------------------------------------------------------------


def test_table_roundtrip():
    xml = build_table_xml(
        [
            ["DAY", {"text": "route", "style": {"font-weight": "bold"}}],
            [{"text": "07:00", "shading_color": "#FFFFCC"}, "集合"],
        ],
        borders_visible=True,
        has_header_row=True,
        col_widths=[80.0, 240.0],
    )
    table = parse_table(_fromstring(xml), {})
    assert table.borders_visible is True
    assert table.has_header_row is True
    assert table.columns == [80.0, 240.0]
    assert [[c.text for c in row] for row in table.rows] == [["DAY", "route"], ["07:00", "集合"]]
    assert table.rows[0][1].paragraphs[0].runs[0].span_style == {"font-weight": "bold"}
    assert table.rows[1][0].shading_color == "#FFFFCC"


def test_table_default_columns_from_widest_row():
    xml = build_table_xml([["a", "b", "c"]])
    table = parse_table(_fromstring(xml), {})
    assert len(table.columns) == 3
    assert len(table.rows[0]) == 3


# --- build_image_xml --------------------------------------------------------------------


def test_image_roundtrip_inline_data():
    xml = build_image_xml(PNG_1PX, "image/png", width=120.0, height=80.0)
    img = parse_image(_fromstring(xml))
    assert img.format == "png"
    assert img.data_b64 is not None
    assert base64.b64decode(img.data_b64)[:8] == b"\x89PNG\r\n\x1a\n"
    assert img.width == 120.0
    assert img.height == 80.0
    assert img.callback_id is None
