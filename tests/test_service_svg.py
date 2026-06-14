"""Tier-1 unit tests for service/svg.py — the SVG→PNG rasterizer behind insert_svg_image.

Pure computation (no COM), so it runs on the Linux host. We assert the PNG is valid and that the
vector-only contract holds (an SVG embedding a raster image is rejected), NOT font fidelity —
real-glyph CJK rendering is a VM/freeze check (resvg uses 微軟正黑體 there; Linux falls back).
"""

from __future__ import annotations

import pytest

from onenote_com_mcp.service.svg import rasterize_svg

_PNG_MAGIC = b"\x89PNG\r\n\x1a\n"
_OK_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16">'
    '<circle cx="8" cy="8" r="6" fill="#3366cc"/></svg>'
)


def test_rasterizes_valid_svg_to_png():
    png = rasterize_svg(_OK_SVG)
    assert png.startswith(_PNG_MAGIC)
    assert len(png) > 100


def test_rejects_empty_input():
    with pytest.raises(ValueError, match="empty"):
        rasterize_svg("   ")


@pytest.mark.parametrize(
    "uri",
    [
        'href="data:image/png;base64,iVBOR"',
        'xlink:href="data:image/jpeg;base64,/9j/4A"',
        "data:image/webp;base64,UklGR",
    ],
)
def test_rejects_embedded_raster_data_uri(uri):
    svg = f'<svg xmlns="http://www.w3.org/2000/svg"><image {uri}/></svg>'
    with pytest.raises(ValueError, match="vector-only"):
        rasterize_svg(svg)


def test_allows_text_and_paths_without_raster():
    # CJK text must not be mistaken for an embedded raster; it just rasterizes (font falls back
    # on Linux). The point is it does NOT raise.
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="40">'
        '<text x="4" y="28" font-family="sans-serif" font-size="20">下龍灣</text></svg>'
    )
    assert rasterize_svg(svg).startswith(_PNG_MAGIC)
