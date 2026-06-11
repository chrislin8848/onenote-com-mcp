"""Build OneNote XML fragments for NEW content. Pure lxml — no COM.

The ``make_*`` functions return lxml elements the service layer grafts into a live
GetPageContent tree; the ``build_*_xml`` wrappers serialize them standalone (namespaced).
Editing EXISTING content never goes through here — that is in-place tree mutation in
``service/page_edit.py`` (SPEC §5); rebuilding a page from a model drops formatting.

Highlight duality (SPEC §5) is enforced one level down in
:func:`onenote_com_mcp.xmllayer.spans.build_style_attr`: every ``background:`` is
emitted with a matching ``mso-highlight:``.
"""

from __future__ import annotations

from typing import Any

from lxml import etree

from onenote_com_mcp.xmllayer.namespaces import NSMAP, qn
from onenote_com_mcp.xmllayer.spans import build_spans


def _tostring(el: etree._Element) -> str:
    return etree.tostring(el, encoding="unicode")


def make_text_oe(
    runs: list[Any],
    quick_style_index: int | None = None,
    alignment: str | None = None,
    *,
    root: bool = False,
) -> etree._Element:
    """Runs → a ``one:OE`` element holding one ``one:T`` (CDATA spans)."""
    oe = etree.Element(qn("OE"), nsmap=NSMAP if root else None)
    if alignment:
        oe.set("alignment", alignment)
    if quick_style_index is not None:
        oe.set("quickStyleIndex", str(quick_style_index))
    t = etree.SubElement(oe, qn("T"))
    # build_spans escapes ">" in run text, so the CDATA terminator "]]>" cannot occur
    t.text = etree.CDATA(build_spans(runs))
    return oe


def build_text_oe_xml(
    runs: list[Any],
    quick_style_index: int | None = None,
    alignment: str | None = None,
) -> str:
    """Runs (str | dict ``{"text", "style"}`` | :class:`~..spans.Run`) → ``one:OE`` XML."""
    return _tostring(make_text_oe(runs, quick_style_index, alignment, root=True))


def _cell_runs(cell: Any) -> tuple[list[Any], str | None, str | None]:
    """Normalize one table-cell input → (runs, shading_color, alignment)."""
    if isinstance(cell, dict):
        runs = cell.get("runs") or [{"text": cell.get("text", ""), "style": cell.get("style")}]
        return runs, cell.get("shading_color"), cell.get("alignment")
    if isinstance(cell, list):
        return cell, None, None
    return [cell], None, None


def make_table_row(cells: list[Any], n_cols: int) -> etree._Element:
    """One row of cell inputs → a ``one:Row`` element, padded to ``n_cols`` cells."""
    row_el = etree.Element(qn("Row"))
    for cell in list(cells) + [""] * (n_cols - len(cells)):
        runs, shading_color, alignment = _cell_runs(cell)
        cell_el = etree.SubElement(row_el, qn("Cell"))
        if shading_color:
            cell_el.set("shadingColor", shading_color)
        children = etree.SubElement(cell_el, qn("OEChildren"))
        children.append(make_text_oe(runs, alignment=alignment))
    return row_el


def make_table(
    rows: list[list[Any]],
    borders_visible: bool = True,
    has_header_row: bool = False,
    col_widths: list[float] | None = None,
) -> etree._Element:
    """Structured rows → a full ``one:Table`` element (Columns + Row/Cell/OEChildren/OE)."""
    n_cols = max((len(r) for r in rows), default=0)
    widths = col_widths if col_widths is not None else [120.0] * n_cols
    table = etree.Element(qn("Table"), nsmap=NSMAP)
    table.set("bordersVisible", "true" if borders_visible else "false")
    table.set("hasHeaderRow", "true" if has_header_row else "false")
    columns = etree.SubElement(table, qn("Columns"))
    for i, width in enumerate(widths):
        col = etree.SubElement(columns, qn("Column"))
        col.set("index", str(i))
        col.set("width", str(float(width)))
    for row in rows:
        table.append(make_table_row(row, n_cols))
    return table


def build_table_xml(
    rows: list[list[Any]],
    borders_visible: bool = True,
    has_header_row: bool = False,
    col_widths: list[float] | None = None,
) -> str:
    """Structured rows → full ``one:Table`` XML.

    Cells accept str, run-list, or dict ``{"text"|"runs", "style", "shading_color",
    "alignment"}``. Short rows are padded to the widest row.
    """
    return _tostring(make_table(rows, borders_visible, has_header_row, col_widths))


def make_image(
    data_b64: str,
    media_type: str = "image/png",
    width: float | None = None,
    height: float | None = None,
) -> etree._Element:
    """base64 bytes → a ``one:Image`` element with inline ``one:Data``."""
    image = etree.Element(qn("Image"), nsmap=NSMAP)
    image.set("format", media_type.rsplit("/", 1)[-1])
    if width is not None or height is not None:
        size = etree.SubElement(image, qn("Size"))
        if width is not None:
            size.set("width", str(float(width)))
        if height is not None:
            size.set("height", str(float(height)))
        size.set("isSetByUser", "true")
    data = etree.SubElement(image, qn("Data"))
    data.text = data_b64
    return image


def build_image_xml(
    data_b64: str,
    media_type: str = "image/png",
    width: float | None = None,
    height: float | None = None,
) -> str:
    """base64 bytes → ``one:Image`` XML with inline ``one:Data`` (insert path, SPEC §4)."""
    return _tostring(make_image(data_b64, media_type, width, height))
