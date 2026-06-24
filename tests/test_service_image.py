"""Tier-1 tests for insert_image_from_path (#8): local-file raster insert.

The bytes are read off disk by the server (never through the model), so these run on Linux with a
tiny real PNG written to a temp file. The COM round-trip (the image actually landing in OneNote) is
a Tier-2/VM check; here we assert the load/validation logic and that the insert routes a one:Image
with inline one:Data through the single write core.
"""

from __future__ import annotations

import base64

import pytest
from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize
from onenote_com_mcp.service import image, page_edit
from onenote_com_mcp.xmllayer.namespaces import qn

# a real 1x1 transparent PNG
_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAXpeqz8AAAAASUVORK5CYII="
)
_PAGE_ID = "{P}{1}{B0}"
_MINIMAL_PAGE = (
    '<?xml version="1.0"?>'
    '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
    f'ID="{_PAGE_ID}" lastModifiedTime="2026-06-10T17:39:30.000Z">'
    "<one:Outline><one:OEChildren/></one:Outline></one:Page>"
)
_PARSER = etree.XMLParser(strip_cdata=False)


def _png_file(tmp_path) -> str:
    p = tmp_path / "pic.png"
    p.write_bytes(_PNG)
    return str(p)


def _backend_with_page(tmp_path) -> FixtureBackend:
    (tmp_path / f"page_{_sanitize(_PAGE_ID)}.xml").write_text(_MINIMAL_PAGE, encoding="utf-8")
    return FixtureBackend(tmp_path)


def _sent(be: FixtureBackend) -> etree._Element:
    writes = [c for c in be.calls if c.method == "update_page_content"]
    assert len(writes) == 1, "an image insert must be exactly ONE UpdatePageContent"
    return etree.fromstring(writes[0].kwargs["changes_xml"].encode("utf-8"), parser=_PARSER)


# --- load_local_image (validation) --------------------------------------------------


def test_load_local_image_reads_png(tmp_path):
    data_b64, media_type = image.load_local_image(_png_file(tmp_path))
    assert media_type == "image/png"
    assert base64.b64decode(data_b64) == _PNG


def test_load_local_image_sniffs_jpeg_and_gif(tmp_path):
    jpeg = tmp_path / "j.jpg"
    jpeg.write_bytes(b"\xff\xd8\xff\xe0" + b"\x00" * 32)
    assert image.load_local_image(str(jpeg))[1] == "image/jpeg"
    gif = tmp_path / "g.gif"
    gif.write_bytes(b"GIF89a" + b"\x00" * 32)
    assert image.load_local_image(str(gif))[1] == "image/gif"


def test_load_local_image_rejects_missing_and_empty(tmp_path):
    with pytest.raises(ValueError, match="no readable file"):
        image.load_local_image(str(tmp_path / "nope.png"))
    with pytest.raises(ValueError, match="path is empty"):
        image.load_local_image("")


def test_load_local_image_rejects_non_raster(tmp_path):
    # an SVG (text) is not a raster — must point the user at insert_svg_image
    svg = tmp_path / "v.svg"
    svg.write_text("<svg xmlns='http://www.w3.org/2000/svg'/>", encoding="utf-8")
    with pytest.raises(ValueError, match="raster"):
        image.load_local_image(str(svg))


def test_load_local_image_rejects_oversized(tmp_path, monkeypatch):
    monkeypatch.setattr(image, "_MAX_BYTES", 4)
    with pytest.raises(ValueError, match="over the"):
        image.load_local_image(_png_file(tmp_path))


# --- insert_image_from_path (routes through the write core) -------------------------


def test_insert_image_from_path_appends_inline_image(tmp_path):
    be = _backend_with_page(tmp_path)
    page_edit.insert_image_from_path(be, _PAGE_ID, _png_file(tmp_path))

    sent = _sent(be)
    images = list(sent.iter(qn("Image")))
    assert len(images) == 1
    data = images[0].find(qn("Data"))
    assert data is not None and base64.b64decode(data.text) == _PNG
    # wrapped in its own OE under the outline (the OE carries the deletable objectID)
    assert images[0].getparent().tag == qn("OE")
    assert images[0].find(qn("CallbackID")) is None  # inline Data, no read-side callback


def test_insert_image_from_path_rejects_missing_file_before_writing(tmp_path):
    be = _backend_with_page(tmp_path)
    with pytest.raises(ValueError, match="no readable file"):
        page_edit.insert_image_from_path(be, _PAGE_ID, str(tmp_path / "ghost.png"))
    assert not [c for c in be.calls if c.method == "update_page_content"]


def test_insert_image_from_path_rejects_bad_mode(tmp_path):
    be = _backend_with_page(tmp_path)
    with pytest.raises(ValueError, match="mode must be one of"):
        page_edit.insert_image_from_path(be, _PAGE_ID, _png_file(tmp_path), mode="replace")
