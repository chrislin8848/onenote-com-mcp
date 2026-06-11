"""parse_page against the three real content pages (混合樣式頁 / 表格頁 / 圖片頁).

The assertions here ARE the format-preservation contract (SPEC §5): effective run style =
QuickStyleDef baseline (via the OE ``quickStyleIndex``) overlaid with the OE ``style``
attribute, overlaid with the inline span style. Highlight reads from either ``background:``
or ``mso-highlight:``.
"""

from __future__ import annotations

import base64

from onenote_com_mcp.backend.fixture import FixtureBackend
from onenote_com_mcp.xmllayer.namespaces import qn
from onenote_com_mcp.xmllayer.parse import parse_page

MIXED_ID = (
    "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E19540013362017467321520163829129860902849621}"
)
TABLE_ID = (
    "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E19500773287729139935320149797721816501902621}"
)
IMAGE_ID = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E1953306013858222940101982353039053288030011}"


def _page(fixtures_dir, page_id):
    return parse_page(FixtureBackend(fixtures_dir).get_page_content(page_id))


# --- page shell ----------------------------------------------------------------------


def test_page_attrs_and_quick_style_defs(fixtures_dir):
    page = _page(fixtures_dir, MIXED_ID)
    assert page.id == MIXED_ID
    assert page.name == "混合樣式頁"
    assert page.page_level == 1
    assert page.last_modified_time == "2026-06-11T03:42:19.000Z"
    # QuickStyleDef table parsed: index → baseline
    assert set(page.quick_styles) == {0, 1}
    assert page.quick_styles[0].name == "PageTitle"
    assert page.quick_styles[0].font == "Calibri"
    assert page.quick_styles[0].font_size == "20.0"
    assert page.quick_styles[1].name == "p"
    assert page.quick_styles[1].font_size == "11.0"
    # "automatic" colors are normalized to None
    assert page.quick_styles[1].font_color is None
    assert page.quick_styles[1].highlight_color is None


def test_title_effective_style(fixtures_dir):
    page = _page(fixtures_dir, MIXED_ID)
    assert page.title is not None
    assert page.title.text == "混合樣式頁"
    (run,) = page.title.runs
    # QuickStyleDef 0 (Calibri 20.0) overlaid by the OE style attr
    assert run.style["font-family"] == "Microsoft JhengHei"
    assert run.style["font-size"] == "20.0pt"


def test_raw_nodes_are_kept_and_live(fixtures_dir):
    """Paragraphs keep references to the REAL lxml nodes of the page tree (SPEC §5:

    the edit model is the tree itself — never collapse to plain text)."""
    page = _page(fixtures_dir, MIXED_ID)
    para = page.outlines[0].paragraphs[0]
    assert para.node.tag == qn("OE")
    # identity, not a copy: the node is part of page.node's tree
    assert any(para.node is el for el in page.node.iter(qn("OE")))


# --- 混合樣式頁: the three-layer style merge ------------------------------------------


def test_highlight_run_dual_attribute(fixtures_dir):
    page = _page(fixtures_dir, MIXED_ID)
    para = page.outlines[0].paragraphs[0]
    assert para.text == (
        "務必含螢光標示文字（highlight 雙屬性 background: + mso-highlight: "
        "是已知坑，沒有它 parser 這段就沒 ground truth）"
    )
    assert len(para.runs) == 11
    hl = para.runs[1]
    assert hl.text == "螢光標示文字"
    # span layer wins for family+highlight, OE style wins for size, QSD gives the rest
    assert hl.style["font-family"] == "Microsoft JhengHei"
    assert hl.style["font-size"] == "12.0pt"
    assert hl.style["background"] == "yellow"
    # the un-highlighted neighbour run has no background
    assert "background" not in para.runs[0].style


def test_plain_paragraph_inherits_oe_and_qsd(fixtures_dir):
    page = _page(fixtures_dir, MIXED_ID)
    para = page.outlines[0].paragraphs[1]  # bare "+ " — no spans at all
    (run,) = para.runs
    assert run.text == "+ "
    assert run.span_style == {}
    assert run.style["font-family"] == "Microsoft JhengHei"
    assert run.style["font-size"] == "12.0pt"


def test_color_font_size_and_decoration_runs(fixtures_dir):
    page = _page(fixtures_dir, MIXED_ID)
    paras = page.outlines[0].paragraphs
    # 不同字色: span color over OE/QSD
    assert paras[2].runs[0].text == "不同字色"
    assert paras[2].runs[0].style["color"] == "#FA0000"
    assert paras[2].runs[1].text == "/"
    assert "color" not in paras[2].runs[1].style
    # 字型: unquoted CJK font name in the span
    assert paras[3].runs[0].text == "字型"
    assert paras[3].runs[0].style["font-family"] == "新細明體"
    # 字級: span font-size overrides QSD 11.0 (OE style has no size here)
    assert paras[4].runs[0].text == "字級"
    assert paras[4].runs[0].style["font-size"] == "16.0pt"
    # bold / italic / underline / strikethrough paragraphs
    assert paras[5].runs[0].style["font-weight"] == "bold"
    assert paras[6].runs[0].style["font-style"] == "italic"
    assert paras[7].runs[0].style["text-decoration"] == "underline"
    assert paras[8].runs[0].style["text-decoration"] == "line-through"
    assert len(paras) == 9


# --- 表格頁: structured rows, never flattened -----------------------------------------


def test_table_structure(fixtures_dir):
    page = _page(fixtures_dir, TABLE_ID)
    (table,) = page.tables
    assert table.borders_visible is True
    assert table.has_header_row is True
    assert len(table.columns) == 2
    assert table.columns[0] == 66.56893920898437
    assert len(table.rows) == 10
    assert all(len(row) == 2 for row in table.rows)


def test_table_header_cell_styles(fixtures_dir):
    page = _page(fixtures_dir, TABLE_ID)
    table = page.tables[0]
    cell = table.rows[0][0]
    assert cell.shading_color == "#FFFFCC"
    assert cell.text == "DAY 1"
    para = cell.paragraphs[0]
    assert para.alignment == "center"
    # both runs bold (span), 新細明體 10.5pt from the cell OE style
    for run in para.runs:
        assert run.style["font-weight"] == "bold"
        assert run.style["font-family"] == "新細明體"
        assert run.style["font-size"] == "10.5pt"
    assert table.rows[0][1].shading_color == "#FFCC99"
    assert table.rows[0][1].text == "地點一 – 地點二 – 地點三 – 地點四"


def test_table_empty_cell_and_nbsp(fixtures_dir):
    page = _page(fixtures_dir, TABLE_ID)
    table = page.tables[0]
    assert table.rows[1][0].text == ""  # <one:T/> — empty, not missing
    # &nbsp; entity inside CDATA decodes to U+00A0
    assert "FL999\xa0" in table.rows[3][1].text


def test_table_bold_italic_colored_cell(fixtures_dir):
    page = _page(fixtures_dir, TABLE_ID)
    run = page.tables[0].rows[3][0].paragraphs[0].runs[0]
    assert run.text == "13:10"
    assert run.style["font-weight"] == "bold"
    assert run.style["font-style"] == "italic"
    assert run.style["color"] == "#FA0000"  # from the OE style attr, not the span


def test_table_cell_text_rows_never_one_string(fixtures_dir):
    page = _page(fixtures_dir, TABLE_ID)
    rows_text = [[c.text for c in row] for row in page.tables[0].rows]
    assert rows_text[2] == ["07:00", "機場集合"]
    assert rows_text[9] == ["HOTEL", "範例飯店"]


# --- 圖片頁: callback id (no inline data in the basic dump) ---------------------------


def test_image_callback_id_and_ocr(fixtures_dir):
    page = _page(fixtures_dir, IMAGE_ID)
    (img,) = page.images
    assert img.callback_id == "{4F0825D2-04C7-4D62-A3A3-B517E0CA0F87}{14}{B0}"
    assert img.data_b64 is None  # ground truth: no inline one:Data in this dump
    assert img.width == 371.5449523925781
    assert img.height == 275.2499389648437
    assert "富士山" in img.ocr_text
    # surrounding text still parses
    assert page.outlines[0].paragraphs[0].text == "這裡有圖片"


def test_image_binary_fixture_resolves_via_callback(fixtures_dir):
    """The parsed callbackID keys straight into GetBinaryPageContent (read path wiring)."""
    page = _page(fixtures_dir, IMAGE_ID)
    b64 = FixtureBackend(fixtures_dir).get_binary_page_content(IMAGE_ID, page.images[0].callback_id)
    raw = base64.b64decode(b64)
    assert raw[:8] == b"\x89PNG\r\n\x1a\n"
