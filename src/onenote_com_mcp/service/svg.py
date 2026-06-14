"""SVG -> PNG rasterization for insert_svg_image (the ONLY picture-insert path).

OneNote's page schema takes raster binary only (``one:Data``), not SVG, so a generated vector
graphic is rendered to PNG here — in-process, via ``resvg_py`` (a PyO3 binding to the resvg Rust
library) — before it is wrapped in a ``one:Image``. Pure computation, no COM: it runs on Linux
(Tier-1) and inside the frozen Windows exe alike (PyInstaller freeze validated 2026-06-14).

Vector-only by contract: an SVG that embeds a RASTER image (an ``<image>`` with a ``data:`` URI of
a PNG/JPG/…) is REJECTED — that would smuggle a photo back through the model as base64, the exact
slowness insert_svg_image exists to avoid. Photos are inserted by hand in OneNote.
"""

from __future__ import annotations

import re

# resvg maps the generic "sans-serif" family to an arbitrary face on Windows (a handwriting font
# in the freeze spike); pin a CJK-clean default so Chinese text renders in 微軟正黑體 even when the
# SVG only says font-family="sans-serif" or omits it. Absent on Linux → resvg falls back gracefully
# (Tier-1 only asserts the PNG is valid, not the face — real-font fidelity is a VM check).
_DEFAULT_FONT = "Microsoft JhengHei"

# A data: URI whose media type is a RASTER image = an embedded photo/PNG/JPG. (data:image/svg+xml
# is vector and intentionally NOT matched.)
_RASTER_DATA_URI = re.compile(
    r"data:image/(?:png|jpe?g|gif|bmp|webp|tiff?|x-icon|vnd)", re.IGNORECASE
)

_PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def rasterize_svg(svg: str) -> bytes:
    """Render SVG markup to PNG bytes.

    Raises ValueError on empty input, an embedded raster image, or output that is not a PNG.
    """
    if not svg or not svg.strip():
        raise ValueError("svg is empty — pass complete <svg>…</svg> markup")
    if _RASTER_DATA_URI.search(svg):
        raise ValueError(
            "the SVG embeds a raster image (a data: URI for a PNG/JPG/…). insert_svg_image is "
            "vector-only; a photo or existing raster image must be inserted BY HAND in OneNote."
        )

    import resvg_py  # in-process PyO3 rasterizer (cross-platform wheel; imported lazily)

    png = bytes(
        resvg_py.svg_to_bytes(
            svg_string=svg,
            font_family=_DEFAULT_FONT,
            sans_serif_family=_DEFAULT_FONT,
        )
    )
    if not png.startswith(_PNG_MAGIC):
        raise ValueError("rasterization did not produce a PNG (malformed SVG?)")
    return png
