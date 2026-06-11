"""Parse OneNote XML → structured data. Pure lxml over strings — no COM.

TDD'd against REAL fixtures dumped from the VM (tests/fixtures/, 2026-06-11); where the
ground truth disagreed with the Microsoft-docs sketch, the fixture won. Notable findings
baked in here:

* ``one:Image`` carries its callback as a ``one:CallbackID`` CHILD element (the docs
  sketch shows a ``callbackID`` attribute — we read both).
* Inline ``one:Data`` may be absent even on a ``piBinaryData`` dump; images resolve via
  ``GetBinaryPageContent(callback_id)``.
* Effective run style is a three-layer overlay (SPEC §5): QuickStyleDef baseline (via the
  OE ``quickStyleIndex``) ← OE ``style`` attribute ← inline span style; highlight readable
  from either ``background:`` or ``mso-highlight:``.
"""

from __future__ import annotations

from typing import Any

from lxml import etree

from onenote_com_mcp.xmllayer.models import (
    Cell,
    Image,
    Outline,
    Page,
    Paragraph,
    QuickStyleDef,
    Table,
)
from onenote_com_mcp.xmllayer.namespaces import local_name, qn
from onenote_com_mcp.xmllayer.spans import parse_spans, parse_style_attr

# ---------------------------------------------------------------------------- hierarchy

_HIERARCHY_TYPES = {
    "Notebook": "notebook",
    "SectionGroup": "section_group",
    "Section": "section",
    "Page": "page",
}
# container roots GetHierarchy/FindPages may return when not scoped to a single node
_CONTAINER_ROOTS = {"Notebooks", "Sections", "SectionGroups", "Pages"}


def _bool(value: str | None) -> bool:
    return value == "true"


def _hierarchy_node(el: etree._Element, include_recycle_bin: bool) -> dict[str, Any] | None:
    kind = _HIERARCHY_TYPES.get(local_name(el.tag))
    if kind is None:
        return None
    if not include_recycle_bin and (
        _bool(el.get("isRecycleBin")) or _bool(el.get("isInRecycleBin"))
    ):
        return None
    node: dict[str, Any] = {
        "type": kind,
        "id": el.get("ID"),
        "name": el.get("name"),
        "last_modified_time": el.get("lastModifiedTime"),
        "is_currently_viewed": _bool(el.get("isCurrentlyViewed")),
    }
    if kind == "page":
        node["date_time"] = el.get("dateTime")
        node["page_level"] = int(el.get("pageLevel", "1"))
    else:
        node["path"] = el.get("path")
        node["children"] = [
            child for sub in el if (child := _hierarchy_node(sub, include_recycle_bin)) is not None
        ]
    if kind == "notebook":
        node["nickname"] = el.get("nickname")
    if kind in ("notebook", "section"):
        node["color"] = el.get("color")
    if kind == "section_group":
        node["is_recycle_bin"] = _bool(el.get("isRecycleBin"))
    if kind == "section":
        node["is_locked"] = _bool(el.get("locked"))
    if include_recycle_bin and kind in ("section", "page"):
        node["is_in_recycle_bin"] = _bool(el.get("isInRecycleBin"))
    return node


def parse_hierarchy(xml: str, include_recycle_bin: bool = False) -> list[dict[str, Any]]:
    """GetHierarchy/FindPages XML → list of nested node dicts, document order preserved.

    The root may be a single scoped node (``one:Notebook`` …) or a container
    (``one:Notebooks`` …). A notebook/group's children stay a MIXED section +
    section-group list (SPEC v0611: no flattening). Recycle-bin nodes are skipped
    unless ``include_recycle_bin``.
    """
    root = etree.fromstring(xml.encode("utf-8"))
    elements = list(root) if local_name(root.tag) in _CONTAINER_ROOTS else [root]
    return [
        node for el in elements if (node := _hierarchy_node(el, include_recycle_bin)) is not None
    ]


# --------------------------------------------------------------------------------- page


def _effective_style(
    quick_styles: dict[int, QuickStyleDef],
    quick_style_index: int | None,
    oe_style: dict[str, str],
    span_style: dict[str, str],
) -> dict[str, str]:
    base = quick_styles.get(quick_style_index)
    merged = {**(base.as_css() if base else {}), **oe_style, **span_style}
    highlight = merged.pop("mso-highlight", None)
    if highlight and "background" not in merged:
        merged["background"] = highlight
    merged.pop("text-align", None)  # paragraph-level; exposed as Paragraph.alignment
    return merged


def parse_image(el: etree._Element) -> Image:
    """``one:Image`` element → :class:`Image`. Callback may be a child OR an attribute."""
    callback_el = el.find(qn("CallbackID"))
    size_el = el.find(qn("Size"))
    data_el = el.find(qn("Data"))
    ocr_el = el.find(f"{qn('OCRData')}/{qn('OCRText')}")
    return Image(
        node=el,
        object_id=el.get("objectID"),
        callback_id=(
            callback_el.get("callbackID") if callback_el is not None else el.get("callbackID")
        ),
        format=el.get("format"),
        width=float(size_el.get("width")) if size_el is not None else None,
        height=float(size_el.get("height")) if size_el is not None else None,
        data_b64=(data_el.text or None) if data_el is not None else None,
        ocr_text=ocr_el.text if ocr_el is not None else None,
    )


def parse_table(el: etree._Element, quick_styles: dict[int, QuickStyleDef]) -> Table:
    """``one:Table`` element → :class:`Table` with structured rows — never one string."""
    columns = [float(col.get("width")) for col in el.findall(f"{qn('Columns')}/{qn('Column')}")]
    rows: list[list[Cell]] = []
    for row_el in el.findall(qn("Row")):
        cells: list[Cell] = []
        for cell_el in row_el.findall(qn("Cell")):
            cells.append(
                Cell(
                    node=cell_el,
                    object_id=cell_el.get("objectID"),
                    shading_color=cell_el.get("shadingColor"),
                    paragraphs=[
                        parse_oe(oe, quick_styles)
                        for oe in cell_el.findall(f"{qn('OEChildren')}/{qn('OE')}")
                    ],
                )
            )
        rows.append(cells)
    return Table(
        node=el,
        object_id=el.get("objectID"),
        borders_visible=_bool(el.get("bordersVisible")),
        has_header_row=_bool(el.get("hasHeaderRow")),
        columns=columns,
        rows=rows,
    )


def parse_oe(el: etree._Element, quick_styles: dict[int, QuickStyleDef]) -> Paragraph:
    """``one:OE`` element → :class:`Paragraph` with effective per-run styles resolved."""
    qsi_attr = el.get("quickStyleIndex")
    quick_style_index = int(qsi_attr) if qsi_attr is not None else None
    oe_style = parse_style_attr(el.get("style"))
    runs = [run for t in el.findall(qn("T")) for run in parse_spans(t.text)]
    for run in runs:
        run.style = _effective_style(quick_styles, quick_style_index, oe_style, run.span_style)
    table_el = el.find(qn("Table"))
    image_el = el.find(qn("Image"))
    return Paragraph(
        node=el,
        object_id=el.get("objectID"),
        alignment=el.get("alignment"),
        quick_style_index=quick_style_index,
        style=oe_style,
        runs=runs,
        table=parse_table(table_el, quick_styles) if table_el is not None else None,
        image=parse_image(image_el) if image_el is not None else None,
        children=[
            parse_oe(child, quick_styles) for child in el.findall(f"{qn('OEChildren')}/{qn('OE')}")
        ],
    )


def _parse_quick_styles(root: etree._Element) -> dict[int, QuickStyleDef]:
    defs: dict[int, QuickStyleDef] = {}
    for el in root.findall(qn("QuickStyleDef")):
        index = int(el.get("index"))
        defs[index] = QuickStyleDef(
            node=el,
            index=index,
            name=el.get("name", ""),
            font=el.get("font"),
            font_size=el.get("fontSize"),
            font_color=_color_or_none(el.get("fontColor")),
            highlight_color=_color_or_none(el.get("highlightColor")),
            bold=_bool(el.get("bold")),
            italic=_bool(el.get("italic")),
        )
    return defs


def _color_or_none(value: str | None) -> str | None:
    return None if value in (None, "automatic", "none") else value


def parse_page(xml: str) -> Page:
    """GetPageContent XML → :class:`Page`. ``page.node`` is the live tree — edits mutate it."""
    root = etree.fromstring(xml.encode("utf-8"))
    quick_styles = _parse_quick_styles(root)
    title_oe = root.find(f"{qn('Title')}/{qn('OE')}")
    page_level = root.get("pageLevel")
    return Page(
        node=root,
        id=root.get("ID"),
        name=root.get("name"),
        date_time=root.get("dateTime"),
        last_modified_time=root.get("lastModifiedTime"),
        page_level=int(page_level) if page_level is not None else None,
        lang=root.get("lang"),
        quick_styles=quick_styles,
        title=parse_oe(title_oe, quick_styles) if title_oe is not None else None,
        outlines=[
            Outline(
                node=outline_el,
                object_id=outline_el.get("objectID"),
                paragraphs=[
                    parse_oe(oe, quick_styles)
                    for oe in outline_el.findall(f"{qn('OEChildren')}/{qn('OE')}")
                ],
            )
            for outline_el in root.findall(qn("Outline"))
        ],
    )
