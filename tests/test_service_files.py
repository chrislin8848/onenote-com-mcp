"""Phase 5b read tools: get_page_files_info / get_page_files on FixtureBackend.

Page XML = real VM dumps (附件與嵌入物件-1/-2); cache-file bytes = the hand-authored NEUTRAL
fixtures from scripts/make_cache_fixtures.py (the read path is a plain disk read, so synthetic
bytes are faithful — and the PII policy forbids committing real user files). The docx cache is
deliberately absent: it replays "cache unavailable".

Also covers the Phase 5b get_page_images fix: page-level printout renders (previously
invisible) now come back, with their OWN objectID.
"""

from __future__ import annotations

import base64
import shutil

import pytest

from onenote_com_mcp.backend.fixture import FixtureBackend
from onenote_com_mcp.service import files, read

PAGE_1_ID = (
    "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E19113778280437107490620108457794036705249631}"
)
PAGE_2_ID = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E186755685473036950911912026369089049789941}"

FF = "{FF3818B8-3EE5-0E18-33E2-DFECB54FC950}"
C4 = "{C4BBAA5E-D7C6-0E3C-33BC-D6C40B7752D5}"
OC = "{0C5E1260-AF8C-0F5A-0B39-3EC5C6124BE2}"

TXT_OE_ID = f"{C4}{{53}}{{B0}}"
PDF_OE_ID = f"{FF}{{44}}{{B0}}"
PRINTOUT_OE_ID = f"{FF}{{74}}{{B0}}"
DOCX_ID = f"{FF}{{16}}{{B0}}"  # page-level, own objectID; cache deliberately missing
XLSX_ID = f"{FF}{{60}}{{B0}}"  # page-level, own objectID
JPG_OE_ID = f"{OC}{{62}}{{B0}}"
EMBEDDED_OE_ID = f"{OC}{{82}}{{B0}}"
RENDER_IMAGE_ID = f"{FF}{{77}}{{B0}}"


@pytest.fixture
def backend(fixtures_dir) -> FixtureBackend:
    return FixtureBackend(fixtures_dir)


# --- get_page_files_info ---------------------------------------------------------------


def test_info_lists_every_inserted_file_with_deletable_ids(backend):
    info = files.get_page_files_info(backend, PAGE_1_ID)
    assert [(e["object_id"], e["kind"], e["placement"]) for e in info] == [
        (TXT_OE_ID, "attachment_icon", "inline"),
        (PDF_OE_ID, "attachment_icon", "inline"),
        (PRINTOUT_OE_ID, "printout", "inline"),
        (DOCX_ID, "attachment_icon", "page_level"),
        (XLSX_ID, "attachment_icon", "page_level"),
    ]


def test_info_metadata_fields(backend, fixtures_dir):
    by_id = {e["object_id"]: e for e in files.get_page_files_info(backend, PAGE_1_ID)}
    txt = by_id[TXT_OE_ID]
    assert (txt["extension"], txt["media_class"]) == ("txt", "text")
    expected_size = (
        (fixtures_dir / "cachefile_F5154E7A_7DFD_40B1_9FB8_8C73704C7B7D_bin.bin").stat().st_size
    )
    assert txt["size_bytes"] == expected_size
    assert txt["cache_available"] is True
    assert txt["path_source"].endswith("濁水溪發電之旅.txt")
    xlsx = by_id[XLSX_ID]
    assert (xlsx["extension"], xlsx["media_class"]) == ("xlsx", "unsupported")


def test_info_reports_cache_unavailable_not_a_crash(backend):
    by_id = {e["object_id"]: e for e in files.get_page_files_info(backend, PAGE_1_ID)}
    docx = by_id[DOCX_ID]  # its cache fixture is deliberately absent
    assert docx["cache_available"] is False
    assert docx["size_bytes"] is None


def test_info_embedded_preview_pages(backend):
    by_id = {e["object_id"]: e for e in files.get_page_files_info(backend, PAGE_2_ID)}
    embedded = by_id[EMBEDDED_OE_ID]
    assert embedded["kind"] == "embedded_preview"
    assert embedded["preview_pages"] == ["工作表1"]
    assert embedded["path_source"] is None


# --- get_page_files --------------------------------------------------------------------


def test_files_text_decodes_utf8(backend):
    [entry] = files.get_page_files(backend, PAGE_1_ID, object_id=TXT_OE_ID)
    assert entry["status"] == "ok"
    assert entry["media_class"] == "text"
    assert "濁水溪發電之旅" in entry["text"]
    assert entry["truncated"] is False
    assert entry["encoding"] == "utf-8-sig"


def test_files_pdf_extracts_text_server_side(backend):
    [entry] = files.get_page_files(backend, PAGE_1_ID, object_id=PDF_OE_ID)
    assert entry["status"] == "ok"
    assert entry["media_class"] == "pdf"
    assert "Sample questionnaire PDF fixture" in entry["text"]
    assert entry["pdf_pages"] == 1


def test_files_printout_pdf_content_also_extractable(backend):
    # a printout's InsertedFile still carries a pathCache — content comes from there
    [entry] = files.get_page_files(backend, PAGE_1_ID, object_id=PRINTOUT_OE_ID)
    assert entry["kind"] == "printout"
    assert entry["status"] == "ok"
    assert "Sample flyer PDF fixture" in entry["text"]


def test_files_image_returns_base64_bytes(backend, fixtures_dir):
    [entry] = files.get_page_files(backend, PAGE_2_ID, object_id=JPG_OE_ID)
    assert entry["status"] == "ok"
    assert entry["media_type"] == "image/jpeg"
    raw = (fixtures_dir / "cachefile_DBC4674A_857C_476E_8A5F_5BD55E5C5D3A_bin.bin").read_bytes()
    assert base64.b64decode(entry["data_base64"]) == raw


def test_files_office_types_are_explicitly_unsupported(backend):
    [entry] = files.get_page_files(backend, PAGE_2_ID, object_id=EMBEDDED_OE_ID)
    assert entry["status"] == "unsupported"
    assert entry["media_class"] == "unsupported"
    assert "text" in entry["note"]  # points at what IS supported
    assert "data_base64" not in entry and "text" not in entry


def test_files_cache_unavailable_for_supported_type(fixtures_dir, tmp_path):
    # clone the fixtures minus the txt cache: a SUPPORTED class with a purged cache
    clone = tmp_path / "fixtures"
    shutil.copytree(fixtures_dir, clone)
    (clone / "cachefile_F5154E7A_7DFD_40B1_9FB8_8C73704C7B7D_bin.bin").unlink()
    [entry] = files.get_page_files(FixtureBackend(clone), PAGE_1_ID, object_id=TXT_OE_ID)
    assert entry["status"] == "cache_unavailable"
    assert "text" not in entry


def test_files_truncates_at_max_chars(backend):
    [entry] = files.get_page_files(backend, PAGE_1_ID, object_id=TXT_OE_ID, max_chars=10)
    assert entry["truncated"] is True
    assert len(entry["text"]) == 10


def test_files_all_files_when_no_object_id(backend):
    entries = files.get_page_files(backend, PAGE_1_ID)
    assert len(entries) == 5
    # docx/xlsx report unsupported (class check comes BEFORE the cache check — even a missing
    # cache doesn't change what we would refuse to parse)
    assert [e["status"] for e in entries] == ["ok", "ok", "ok", "unsupported", "unsupported"]


# --- get_page / get_page_images integration (Phase 5b additions) -----------------------


def test_get_page_shows_inline_attachments_as_file_blocks(backend):
    page = read.get_page(backend, PAGE_1_ID)
    blocks = [b for o in page["outlines"] for b in o["blocks"] if b["type"] == "file"]
    assert [(b["object_id"], b["kind"]) for b in blocks] == [
        (TXT_OE_ID, "attachment_icon"),
        (PDF_OE_ID, "attachment_icon"),
        (PRINTOUT_OE_ID, "printout"),
    ]
    assert blocks[0]["preferred_name"] == "濁水溪發電之旅.txt"


def test_get_page_images_includes_page_level_printout_render(backend):
    images = read.get_page_images(backend, PAGE_1_ID)
    assert [(i["object_id"], i["is_printout"]) for i in images] == [(RENDER_IMAGE_ID, True)]
    assert images[0]["data_base64"]  # pixels came from the render's callback binary fixture
