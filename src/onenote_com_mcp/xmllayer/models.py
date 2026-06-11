"""Structured views over a parsed ``one:Page`` tree.

Every model keeps a reference to its REAL lxml node (``.node``) — the edit data model is
the tree itself (SPEC §5), these classes are read-side conveniences layered on top, never
a replacement the page gets rebuilt from.

Timestamps stay ISO strings as OneNote emits them; the service layer converts when it
needs a ``datetime`` (e.g. ``dateExpectedLastModified``).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from onenote_com_mcp.xmllayer.spans import Run

if TYPE_CHECKING:
    from lxml import etree

__all__ = ["Cell", "Image", "Outline", "Page", "Paragraph", "QuickStyleDef", "Run", "Table"]


@dataclass
class QuickStyleDef:
    """One ``one:QuickStyleDef`` row — the paragraph-style baseline table (SPEC §5 layer 1)."""

    node: etree._Element
    index: int
    name: str
    font: str | None = None
    font_size: str | None = None  # points, as OneNote writes it: "11.0"
    font_color: str | None = None  # None = "automatic"
    highlight_color: str | None = None
    bold: bool = False
    italic: bool = False

    def as_css(self) -> dict[str, str]:
        """Baseline expressed in the same vocabulary as span styles, for overlaying."""
        css: dict[str, str] = {}
        if self.font:
            css["font-family"] = self.font
        if self.font_size:
            css["font-size"] = f"{self.font_size}pt"
        if self.font_color:
            css["color"] = self.font_color
        if self.highlight_color:
            css["background"] = self.highlight_color
        if self.bold:
            css["font-weight"] = "bold"
        if self.italic:
            css["font-style"] = "italic"
        return css


@dataclass
class Image:
    node: etree._Element
    object_id: str | None = None
    callback_id: str | None = None  # ground truth: a one:CallbackID CHILD, not an attribute
    format: str | None = None
    width: float | None = None
    height: float | None = None
    data_b64: str | None = None  # inline one:Data — absent unless dumped with piBinaryData
    ocr_text: str | None = None


@dataclass
class Cell:
    node: etree._Element
    object_id: str | None = None
    shading_color: str | None = None
    paragraphs: list[Paragraph] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)


@dataclass
class Table:
    node: etree._Element
    object_id: str | None = None
    borders_visible: bool = True
    has_header_row: bool = False
    columns: list[float] = field(default_factory=list)  # widths, in column order
    rows: list[list[Cell]] = field(default_factory=list)


@dataclass
class Paragraph:
    """One ``one:OE``. Content union: text runs, a table, or an image (+ nested children)."""

    node: etree._Element
    object_id: str | None = None
    alignment: str | None = None
    quick_style_index: int | None = None
    style: dict[str, str] = field(default_factory=dict)  # parsed OE style attr (layer 2)
    runs: list[Run] = field(default_factory=list)
    table: Table | None = None
    image: Image | None = None
    children: list[Paragraph] = field(default_factory=list)  # nested one:OEChildren

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class Outline:
    node: etree._Element
    object_id: str | None = None
    paragraphs: list[Paragraph] = field(default_factory=list)


@dataclass
class Page:
    node: etree._Element  # the live root — mutate THIS for edits, then serialize whole
    id: str
    name: str | None = None
    date_time: str | None = None
    last_modified_time: str | None = None
    page_level: int | None = None
    lang: str | None = None
    quick_styles: dict[int, QuickStyleDef] = field(default_factory=dict)
    title: Paragraph | None = None
    outlines: list[Outline] = field(default_factory=list)

    def _walk(self, paragraphs: list[Paragraph]):
        for p in paragraphs:
            yield p
            yield from self._walk(p.children)
            if p.table:
                for row in p.table.rows:
                    for cell in row:
                        yield from self._walk(cell.paragraphs)

    @property
    def paragraphs(self) -> list[Paragraph]:
        """All paragraphs in document order, recursing into children and table cells."""
        return [p for o in self.outlines for p in self._walk(o.paragraphs)]

    @property
    def tables(self) -> list[Table]:
        return [p.table for p in self.paragraphs if p.table is not None]

    @property
    def images(self) -> list[Image]:
        return [p.image for p in self.paragraphs if p.image is not None]
