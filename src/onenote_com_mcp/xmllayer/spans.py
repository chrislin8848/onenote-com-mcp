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
    link: str | None = None  # hyperlink href if this run sits inside an <a href="...">…</a>


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
        # each frame: (style, lang, link). <a href> wraps <span>s (VM ground truth), so an
        # inner span inherits the enclosing link; the link rides until the </a>.
        self._stack: list[tuple[dict[str, str], str | None, str | None]] = []

    def _current(self) -> tuple[dict[str, str], str | None, str | None]:
        return self._stack[-1] if self._stack else ({}, None, None)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        style, lang, link = self._current()
        if tag == "span":
            d = dict(attrs)
            merged = {**style, **parse_style_attr(d.get("style"))}
            self._stack.append((merged, d.get("lang") or lang, link))
        elif tag == "a":
            # inherit style/lang, set the hyperlink for everything until </a>
            self._stack.append((style, lang, dict(attrs).get("href") or link))
        elif tag == "br":
            self.handle_data("\n")

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "br":
            self.handle_data("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in ("span", "a") and self._stack:
            self._stack.pop()

    def handle_data(self, data: str) -> None:
        if not data:
            return
        style, lang, link = self._current()
        last = self.runs[-1] if self.runs else None
        if last and last.span_style == style and last.lang == lang and last.link == link:
            last.text += data  # merge adjacent chunks under the same span + link
        else:
            self.runs.append(
                Run(text=data, span_style=style, style=dict(style), lang=lang, link=link)
            )


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
            # DOUBLE quotes — build_spans wraps the whole style in SINGLE quotes
            # (style='...'), so a single-quoted value here would collide and truncate the
            # attribute. OneNote itself alternates: style='font-family:"Microsoft JhengHei"'.
            value = f'"{value}"'
        parts.append(f"{key}:{value}")
    return ";".join(parts)


def _escape_text(text: str) -> str:
    # also closes the door on "]]>" leaking into the CDATA section (">" is escaped)
    return escape(text, quote=False).replace(">", "&gt;")


def _escape_attr(value: str) -> str:
    # href attribute value: escape & < > " ' so it is safe inside double quotes and the CDATA
    # (HTMLParser unescapes these on the way back in). escape(quote=True) covers all five.
    return escape(value, quote=True)


def build_spans(runs: list) -> str:
    """Runs (Run | dict | str) → CDATA payload. Inverse of :func:`parse_spans`.

    A run carrying a ``link`` is wrapped in ``<a href="…">…</a>`` (VM ground truth: OneNote
    stores hyperlinks as an ``<a>`` around the styled span, inside the one:T CDATA)."""
    out: list[str] = []
    for run in runs:
        link: str | None = None
        lang: str | None = None
        if isinstance(run, str):
            text, style = run, {}
        elif isinstance(run, Run):
            text, style, link, lang = run.text, run.span_style or run.style, run.link, run.lang
        else:
            text = run["text"]
            style = run.get("style") or {}
            link = run.get("link")
            lang = run.get("lang")
        inner = _escape_text(text)
        if style or lang:
            # lang rides on the span (VM ground truth); a run with lang but no style still gets a
            # span to carry it — otherwise the spell-check language is lost on every rebuild.
            attrs = []
            if style:
                attrs.append(f"style='{build_style_attr(style)}'")
            if lang:
                attrs.append(f'lang="{lang}"')
            inner = f"<span {' '.join(attrs)}>{inner}</span>"
        if link:
            inner = f'<a href="{_escape_attr(link)}">{inner}</a>'
        out.append(inner)
    return "".join(out)
