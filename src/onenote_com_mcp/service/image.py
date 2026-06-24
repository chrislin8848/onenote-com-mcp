"""Load a local raster image file for ``insert_image_from_path``.

Reads the bytes of an image FILE ON DISK (on the machine running the server) and validates it is a
real raster (PNG / JPEG / GIF) within a size cap, returning base64 + media type. The point is that
the bytes never travel through the model: the removed base64 ``insert_image`` was unusable because
the MODEL had to emit ~tens-of-thousands of base64 characters per image. Here a client that can put
a file on disk (e.g. an agentic IDE that generated a chart, downloaded an image, or was handed a
photo path) passes just the PATH; the server reads the pixels. Pure file IO + validation, no COM.

Vector graphics still go through ``insert_svg_image`` (compact markup, rasterized server-side).
"""

from __future__ import annotations

import base64
from pathlib import Path

# Same ceiling as attachment extraction (service/files.py): a sane bound, not a hard OneNote limit.
_MAX_BYTES = 20 * 1024 * 1024

# Leading magic bytes → media type. PNG/JPEG/GIF only: these are what OneNote's one:Image reliably
# accepts and what make_image's format attribute (media_type.rsplit("/")) maps cleanly. WebP/BMP/…
# are deliberately not accepted (uncertain OneNote support) — reject with a clear error.
_MAGIC: tuple[tuple[bytes, str], ...] = (
    (b"\x89PNG\r\n\x1a\n", "image/png"),
    (b"\xff\xd8\xff", "image/jpeg"),
    (b"GIF87a", "image/gif"),
    (b"GIF89a", "image/gif"),
)


def load_local_image(path: str) -> tuple[str, str]:
    """Return ``(base64_data, media_type)`` for a local raster image file.

    Raises ``ValueError`` if the path is empty, not a readable file, over the size cap, or its bytes
    are not a PNG/JPEG/GIF (an SVG, PDF, text, or other type is rejected — use ``insert_svg_image``
    for vector graphics). The path is resolved on the machine running the server.
    """
    if not path or not path.strip():
        raise ValueError("path is empty")
    p = Path(path)
    if not p.is_file():
        raise ValueError(
            f"no readable file at {path!r} — the path must be on the machine running the server"
        )
    size = p.stat().st_size
    if size > _MAX_BYTES:
        raise ValueError(f"image is {size} bytes, over the {_MAX_BYTES}-byte limit")
    raw = p.read_bytes()
    media_type = next((mt for magic, mt in _MAGIC if raw.startswith(magic)), None)
    if media_type is None:
        raise ValueError(
            "file is not a PNG, JPEG, or GIF raster image — insert_image_from_path is raster-only; "
            "for vector graphics use insert_svg_image"
        )
    return base64.b64encode(raw).decode("ascii"), media_type
