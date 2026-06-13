"""CDATA span parsing/building — the rich-text core of the XML layer.

All input snippets here are taken verbatim from real VM fixtures (the ``one:T`` CDATA of
the "混合樣式頁"/"表格頁" pages), including OneNote's word-wrap newlines *inside* attribute
values, unquoted ``lang=`` attributes, and ``&nbsp;`` entities.
"""

from __future__ import annotations

from onenote_com_mcp.xmllayer.spans import (
    build_spans,
    build_style_attr,
    parse_spans,
    parse_style_attr,
)

# --- parse_style_attr ---------------------------------------------------------------


def test_style_attr_basic():
    assert parse_style_attr("font-weight:bold;color:#FF0000") == {
        "font-weight": "bold",
        "color": "#FF0000",
    }


def test_style_attr_newlines_inside_value():
    # real shape: OneNote word-wraps inside the attribute value
    assert parse_style_attr('font-family:\n\n"Microsoft JhengHei";background:yellow') == {
        "font-family": "Microsoft JhengHei",
        "background": "yellow",
    }
    assert parse_style_attr("font-weight:\n\nbold;font-style:italic") == {
        "font-weight": "bold",
        "font-style": "italic",
    }


def test_style_attr_quote_styles_and_unquoted_cjk_font():
    assert parse_style_attr("font-family:'Microsoft JhengHei';font-size:12.0pt") == {
        "font-family": "Microsoft JhengHei",
        "font-size": "12.0pt",
    }
    assert parse_style_attr("font-family:新細明體;font-size:10.5pt;text-align:center") == {
        "font-family": "新細明體",
        "font-size": "10.5pt",
        "text-align": "center",
    }


def test_style_attr_empty_and_trailing_semicolon():
    assert parse_style_attr("") == {}
    assert parse_style_attr(None) == {}
    assert parse_style_attr("color:#FA0000;") == {"color": "#FA0000"}


# --- parse_spans --------------------------------------------------------------------


def test_parse_plain_text_is_single_unstyled_run():
    runs = parse_spans("+ ")
    assert [(r.text, r.span_style) for r in runs] == [("+ ", {})]


def test_parse_empty_and_none():
    assert parse_spans("") == []
    assert parse_spans(None) == []


def test_parse_real_highlight_run():
    # verbatim from 混合樣式頁 OE[0] — newline between "span" and "style", newline inside value
    cdata = (
        "<span\n\nstyle='font-family:\"Microsoft JhengHei\"'>務必含</span>"
        '<span style=\'font-family:\n\n"Microsoft JhengHei";background:yellow;'
        "mso-highlight:yellow'>螢光標示文字</span>"
    )
    runs = parse_spans(cdata)
    assert [r.text for r in runs] == ["務必含", "螢光標示文字"]
    assert runs[0].span_style == {"font-family": "Microsoft JhengHei"}
    assert runs[1].span_style == {
        "font-family": "Microsoft JhengHei",
        "background": "yellow",
        "mso-highlight": "yellow",
    }


def test_parse_mixed_spans_and_bare_text():
    runs = parse_spans("<span\n\nstyle='color:#FA0000'>不同字色</span>/")
    assert [(r.text, r.span_style) for r in runs] == [
        ("不同字色", {"color": "#FA0000"}),
        ("/", {}),
    ]


def test_parse_unquoted_lang_attr():
    # verbatim from 表格頁 header cell
    runs = parse_spans(
        "<span\n\nstyle='font-weight:bold' lang=zh-TW>DAY </span>"
        "<span style='font-weight:bold'\n\nlang=en-US>1</span>"
    )
    assert [r.text for r in runs] == ["DAY ", "1"]
    assert all(r.span_style == {"font-weight": "bold"} for r in runs)
    assert [r.lang for r in runs] == ["zh-TW", "en-US"]


def test_parse_nbsp_entity():
    runs = parse_spans("FL999&nbsp; 09:00")
    assert runs[0].text == "FL999\xa0 09:00"


def test_parse_br_becomes_newline():
    runs = parse_spans("line1<br>line2")
    assert "".join(r.text for r in runs) == "line1\nline2"


# --- build side ---------------------------------------------------------------------


def test_build_style_attr_quotes_multiword_font():
    s = build_style_attr({"font-family": "Microsoft JhengHei", "font-size": "12.0pt"})
    assert s == "font-family:'Microsoft JhengHei';font-size:12.0pt"


def test_build_style_attr_highlight_writes_both_attributes():
    # SPEC §5: builder must write BOTH background: and mso-highlight:
    s = build_style_attr({"background": "yellow"})
    assert "background:yellow" in s
    assert "mso-highlight:yellow" in s


def test_build_spans_roundtrip_through_parser():
    original = [
        {"text": "plain "},
        {"text": "bold red", "style": {"font-weight": "bold", "color": "#FF0000"}},
        {"text": "marked", "style": {"background": "yellow"}},
    ]
    cdata = build_spans(original)
    runs = parse_spans(cdata)
    assert [r.text for r in runs] == ["plain ", "bold red", "marked"]
    assert runs[0].span_style == {}
    assert runs[1].span_style == {"font-weight": "bold", "color": "#FF0000"}
    # dual-attribute highlight round-trips
    assert runs[2].span_style["background"] == "yellow"
    assert runs[2].span_style["mso-highlight"] == "yellow"


def test_build_spans_escapes_markup_in_text():
    cdata = build_spans([{"text": "a <b> & c"}])
    assert "<b>" not in cdata
    runs = parse_spans(cdata)
    assert runs[0].text == "a <b> & c"


def test_build_spans_never_emits_cdata_terminator():
    cdata = build_spans([{"text": "x ]]> y"}])
    assert "]]>" not in cdata
    assert "".join(r.text for r in parse_spans(cdata)) == "x ]]> y"


# --- hyperlinks (<a href> inside one:T CDATA) — VM ground truth from the 表格頁 fixture ----


def test_parse_captures_hyperlink_from_a_tag():
    # the exact shape OneNote stores (real fixture): <a href="..."><span ...>text</span></a>
    cdata = '<a href="https://example.com/loc-a"><span lang=zh-TW>範例溫泉</span></a>'
    runs = parse_spans(cdata)
    assert len(runs) == 1
    assert runs[0].text == "範例溫泉"
    assert runs[0].link == "https://example.com/loc-a"
    assert runs[0].lang == "zh-TW"


def test_parse_link_only_on_the_linked_run():
    cdata = '<a href="https://example.com">linked</a><span lang=en-US> tail</span>'
    runs = parse_spans(cdata)
    assert [(r.text, r.link) for r in runs] == [("linked", "https://example.com"), (" tail", None)]


def test_build_wraps_linked_run_in_a_tag():
    cdata = build_spans([{"text": "OneNote", "link": "https://example.com/x?a=1&b=2"}])
    assert cdata.startswith('<a href="https://example.com/x?a=1&amp;b=2">')
    assert cdata.endswith("</a>")


def test_hyperlink_round_trips_text_style_and_href():
    original = [
        {"text": "plain "},
        {"text": "site", "style": {"font-weight": "bold"}, "link": "https://example.com/p?x=1&y=2"},
    ]
    runs = parse_spans(build_spans(original))
    assert [r.text for r in runs] == ["plain ", "site"]
    assert runs[0].link is None
    assert runs[1].link == "https://example.com/p?x=1&y=2"
    assert runs[1].span_style["font-weight"] == "bold"
