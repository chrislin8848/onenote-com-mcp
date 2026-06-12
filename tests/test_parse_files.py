"""Phase 5b parse layer: ``one:InsertedFile`` (both placements, kind rule) and page-level
images, asserted against the REAL attachment-page dumps (附件與嵌入物件-1/-2, 2026-06-12).

Ground truth being locked in (docs/onenote-xml-schema.md):
* TWO placements — inline (inside an Outline OE, no own objectID) and page-level (direct
  ``one:Page`` child with Position/Size and its OWN objectID).
* kind: no children = attachment_icon; ``one:Printout`` child = printout; ``one:Previews``
  child = embedded_preview (and embedded objects carry NO pathSource).
* A printout's rendered pages are PAGE-LEVEL ``one:Image``s (``isPrintOut="true"``) —
  they must surface through ``Page.images`` (they were invisible before Phase 5b).
"""

from __future__ import annotations

import pytest

from onenote_com_mcp.xmllayer.parse import parse_page

_PREFIX = "page_65FA3E6E_E6E0_4610_B605_EE3D348FAD13_1_"
PAGE_1 = f"{_PREFIX}E19113778280437107490620108457794036705249631.xml"
PAGE_2 = f"{_PREFIX}E186755685473036950911912026369089049789941.xml"

FF = "{FF3818B8-3EE5-0E18-33E2-DFECB54FC950}"
C4 = "{C4BBAA5E-D7C6-0E3C-33BC-D6C40B7752D5}"
OC = "{0C5E1260-AF8C-0F5A-0B39-3EC5C6124BE2}"


@pytest.fixture
def page1(fixtures_dir):
    return parse_page((fixtures_dir / PAGE_1).read_text(encoding="utf-8"))


@pytest.fixture
def page2(fixtures_dir):
    return parse_page((fixtures_dir / PAGE_2).read_text(encoding="utf-8"))


def test_page1_inventory_inline_first_document_order(page1):
    files = page1.inserted_files
    assert [(f.placement, f.kind, f.preferred_name) for f in files] == [
        ("inline", "attachment_icon", "濁水溪發電之旅.txt"),
        ("inline", "attachment_icon", "丘山行問卷_中英對照.pdf"),
        ("inline", "printout", "A4文宣-25.7.8月分享會.pdf"),
        ("page_level", "attachment_icon", "丘山行Word頁籤(中文) .docx"),
        ("page_level", "attachment_icon", "2026 客人問卷_NEW.xlsx"),
    ]


def test_inline_inserted_file_has_no_own_object_id_enclosing_oe_carries_it(page1):
    txt = page1.inserted_files[0]
    assert txt.object_id is None  # like one:Image — the OE is the deletable unit
    assert txt.node.getparent().get("objectID") == f"{C4}{{53}}{{B0}}"
    assert txt.path_cache and txt.path_cache.endswith(".bin")
    assert txt.path_source == "C:\\Users\\chris\\Desktop\\濁水溪發電之旅.txt"


def test_page_level_inserted_file_carries_its_own_object_id(page1):
    docx, xlsx = page1.page_files
    assert docx.object_id == f"{FF}{{16}}{{B0}}"
    assert xlsx.object_id == f"{FF}{{60}}{{B0}}"
    assert docx.last_modified_time  # page-level variant carries its own stamp


def test_printout_kind_and_xps_index(page1):
    printout = page1.inserted_files[2]
    assert printout.kind == "printout"
    assert printout.xps_file_index == 0  # joins the page-level one:XPSFile carrier


def test_printout_render_is_a_page_level_image_and_now_visible(page1):
    # the rendered page is a DIRECT one:Page child — Page.images must include it
    assert [img.object_id for img in page1.page_images] == [f"{FF}{{77}}{{B0}}"]
    render = page1.page_images[0]
    assert render.is_printout is True
    assert render.callback_id == f"{FF}{{77}}{{B0}}"
    assert render in page1.images


def test_page2_embedded_preview_kind(page2):
    files = page2.inserted_files
    assert [(f.kind, f.preferred_name) for f in files] == [
        ("attachment_icon", "捷斯山屋.jpg"),
        ("embedded_preview", "附件與嵌入物件-2 - 工作表.xlsx"),
    ]
    embedded = files[1]
    assert embedded.path_source is None  # ground truth: embedded objects have NO pathSource
    assert embedded.path_cache
    assert embedded.preview_pages == ["工作表1"]
    assert embedded.node.getparent().get("objectID") == f"{OC}{{82}}{{B0}}"


def test_inline_images_unaffected_by_page_level_collection(page2):
    # page 2's two images are inline (inside OEs); page_images must stay empty there
    assert page2.page_images == []
    assert len(page2.images) == 2
    assert all(not img.is_printout for img in page2.images)
