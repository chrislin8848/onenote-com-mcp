"""Rich-text runs inside ``one:T`` CDATA — parse and build.

The CDATA payload is Office-flavoured HTML: a flat sequence of ``<span style=...>`` runs
and bare text. Ground truth (real VM dumps, tests/fixtures/) shows OneNote word-wraps the
markup with newlines *anywhere* — between attributes and inside attribute values — uses
single-quoted ``style``, unquoted ``lang=zh-TW``, and ``&nbsp;`` entities. Python's
``html.parser`` handles all of that; a regex would not.

Pure string-level module: no lxml, no COM.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from html import escape
from html.parser import HTMLParser


@dataclass
class Run:
    """One uniformly-styled stretch of text.

    ``span_style`` is the raw inline span style (empty for bare text). ``style`` is the
    effective style after the QuickStyleDef/OE overlays — filled in by ``parse.py``;
    at the spans level both start out equal.
    """

    text: str
    span_style: dict[str, str] = field(default_factory=dict)
    style: dict[str, str] = field(default_factory=dict)
    lang: str | None = None


def parse_style_attr(style: str | None) -> dict[str, str]:
    """``"font-family:\\n'X';font-size:12.0pt"`` → ``{"font-family": "X", ...}``.

    Collapses OneNote's word-wrap whitespace, strips quotes around font names.
    """
    if not style:
        return {}
    out: dict[str, str] = {}
    for decl in " ".join(style.split()).split(";"):
        if ":" not in decl:
            continue
        key, value = decl.split(":", 1)
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "'\"":
            value = value[1:-1]
        out[key.strip().lower()] = value
    return out


class _SpanParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.runs: list[Run] = []
        self._stack: list[tuple[dict[str, str], str | None]] = []

    def _current(self) -> tuple[dict[str, str], str | None]:
        return self._stack[-1] if self._stack else ({}, None)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "span":
            d = dict(attrs)
            style, lang = self._current()
            merged = {**style, **parse_style_attr(d.get("style"))}
            self._stack.append((merged, d.get("lang") or lang))
        elif tag == "br":
            self.handle_data("\n")

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "br":
            self.handle_data("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag == "span" and self._stack:
            self._stack.pop()

    def handle_data(self, data: str) -> None:
        if not data:
            return
        style, lang = self._current()
        if self.runs and self.runs[-1].span_style == style and self.runs[-1].lang == lang:
            self.runs[-1].text += data  # merge adjacent chunks under the same span
        else:
            self.runs.append(Run(text=data, span_style=style, style=dict(style), lang=lang))


def parse_spans(cdata: str | None) -> list[Run]:
    """``one:T`` CDATA → list of :class:`Run` in document order. ``None``/empty → ``[]``."""
    if not cdata:
        return []
    parser = _SpanParser()
    parser.feed(cdata)
    parser.close()
    return parser.runs


_MULTIWORD_QUOTE_KEYS = {"font-family"}


def build_style_attr(style: dict[str, str]) -> str:
    """Style dict → ``style`` attribute value.

    Highlight is dual-attribute (SPEC §5): any ``background:`` gets a matching
    ``mso-highlight:`` (and vice versa) so every OneNote build renders it.
    """
    style = dict(style)
    highlight = style.get("background") or style.get("mso-highlight")
    if highlight:
        style["background"] = highlight
        style["mso-highlight"] = highlight
    parts = []
    for key, value in style.items():
        if key in _MULTIWORD_QUOTE_KEYS and " " in value:
            value = f"'{value}'"
        parts.append(f"{key}:{value}")
    return ";".join(parts)


def _escape_text(text: str) -> str:
    # also closes the door on "]]>" leaking into the CDATA section (">" is escaped)
    return escape(text, quote=False).replace(">", "&gt;")


def build_spans(runs: list) -> str:
    """Runs (Run | dict | str) → CDATA payload. Inverse of :func:`parse_spans`."""
    out: list[str] = []
    for run in runs:
        if isinstance(run, str):
            text, style = run, {}
        elif isinstance(run, Run):
            text, style = run.text, run.span_style or run.style
        else:
            text, style = run["text"], run.get("style") or {}
        text = _escape_text(text)
        if style:
            out.append(f"<span style='{build_style_attr(style)}'>{text}</span>")
        else:
            out.append(text)
    return "".join(out)
