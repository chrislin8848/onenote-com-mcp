"""Phase 5b Stage 2: attachment fidelity through the copy seam (transfer_page).

SPEC §5 clone rule: a clone must NOT carry the source's ``pathCache`` (a dead reference
OneNote owns) — the cache bytes are staged aside and the clone's ``pathSource`` re-pointed
at the copy so OneNote re-imports it. Printouts are flattened (their page-level one:XPSFile
carriers hold read-side CallbackIDs that cannot ride into a write): renders survive as plain
inlined images, the source file as a normal attachment — this is expected normalization, not a
loss, so it is NOT reported (user decision 2026-06-12). Anything whose CONTENT could not be
transferred (an unavailable cache) lands in ``file_notes`` — explicitly, never silently (SPEC §5).

Source pages are the real attachment dumps (附件與嵌入物件-1/-2); cache bytes are the neutral
hand-authored fixtures (the docx cache is deliberately absent = unavailable).
"""

from __future__ import annotations

import pytest
from lxml import etree

from onenote_com_mcp.backend.fixture import FixtureBackend, _sanitize
from onenote_com_mcp.service import copy
from onenote_com_mcp.xmllayer.namespaces import qn

_PARSER = etree.XMLParser(strip_cdata=False)
_NEW_PAGE_ID = "{FIXTURE-page-1}{1}{B0}"
_BLANK_STAMP = "2026-06-12T09:00:00.000Z"
_TARGET_SECTION = "{TARGET}{1}{B0}"

_PAGE_PREFIX = "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}"
PAGE_1_ID = f"{_PAGE_PREFIX}{{E19113778280437107490620108457794036705249631}}"
PAGE_2_ID = f"{_PAGE_PREFIX}{{E186755685473036950911912026369089049789941}}"


@pytest.fixture
def be(fixtures_dir, tmp_path) -> FixtureBackend:
    for src in fixtures_dir.glob("page_*.xml"):
        (tmp_path / src.name).write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
    for src in list(fixtures_dir.glob("binary_*.b64")) + list(fixtures_dir.glob("cachefile_*")):
        (tmp_path / src.name).write_bytes(src.read_bytes())
    (tmp_path / f"page_{_sanitize(_NEW_PAGE_ID)}.xml").write_text(
        '<?xml version="1.0"?><one:Page '
        'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" '
        f'ID="{_NEW_PAGE_ID}" lastModifiedTime="{_BLANK_STAMP}"/>',
        encoding="utf-8",
    )
    return FixtureBackend(tmp_path)


def _sent_payload(be: FixtureBackend) -> etree._Element:
    writes = [c for c in be.calls if c.method == "update_page_content"]
    assert len(writes) == 1
    return etree.fromstring(writes[0].kwargs["changes_xml"].encode("utf-8"), parser=_PARSER)


def test_clone_restages_caches_and_never_carries_path_cache(be):
    result = copy.transfer_page(be, PAGE_1_ID, _TARGET_SECTION)
    sent = _sent_payload(be)

    files = list(sent.iter(qn("InsertedFile")))
    # 5 source attachments; the docx (cache deliberately unavailable) is REMOVED → 4 carried
    assert len(files) == 4
    assert all(f.get("pathCache") is None for f in files), "pathCache must NEVER ride"
    assert not any(f.get("preferredName") == "丘山行Word頁籤(中文) .docx" for f in files)

    by_name = {f.get("preferredName"): f for f in files}
    # cache available → pathSource re-pointed at the STAGED copy (FixtureBackend fake path)
    for name in ("濁水溪發電之旅.txt", "丘山行問卷_中英對照.pdf", "2026 客人問卷_NEW.xlsx"):
        assert by_name[name].get("pathSource").startswith("C:\\FixtureStaging\\"), name
        assert by_name[name].get("pathSource").endswith(name)
    # one stage call per cache-bearing attachment (incl. the docx attempt that found nothing)
    stage_calls = [c for c in be.calls if c.method == "stage_cache_copy"]
    assert len(stage_calls) == 5
    assert result.page_id == _NEW_PAGE_ID


def test_clone_strips_printout_marks_from_an_orphan_render(be, tmp_path):
    # OneNote can dissolve a printout's link on its own (VM 2026-10-05: the PDF lost its Printout
    # child, the XPSFile vanished, the render image kept isPrintOut). With NO carrier left, the
    # copy must still deliver the render as a PLAIN image.
    path = tmp_path / f"page_{_sanitize(PAGE_1_ID)}.xml"
    root = etree.fromstring(path.read_bytes(), parser=_PARSER)
    for xps in root.findall(qn("XPSFile")):
        root.remove(xps)
    for printout in list(root.iter(qn("Printout"))):
        printout.getparent().remove(printout)
    path.write_bytes(etree.tostring(root, xml_declaration=True, encoding="UTF-8"))

    copy.transfer_page(be, PAGE_1_ID, _TARGET_SECTION)

    renders = _sent_payload(be).findall(qn("Image"))
    assert len(renders) == 1
    for attr in ("xpsFileIndex", "isPrintOut", "originalPageNumber"):
        assert renders[0].get(attr) is None


def test_clone_flattens_printouts_silently(be):
    result = copy.transfer_page(be, PAGE_1_ID, _TARGET_SECTION)
    sent = _sent_payload(be)

    # the XPSFile carrier (read-side CallbackID construct) must not ride into the write
    assert sent.findall(qn("XPSFile")) == []
    by_name = {f.get("preferredName"): f for f in sent.iter(qn("InsertedFile"))}
    printout = by_name["A4文宣-25.7.8月分享會.pdf"]
    assert printout.find(qn("Printout")) is None  # its xpsFileIndex would dangle
    assert printout.get("pathSource").startswith("C:\\FixtureStaging\\")

    # the render survives as a PLAIN image: inline data, printout bookkeeping stripped
    renders = [
        img
        for img in sent.findall(qn("Image"))  # page-level images
    ]
    assert len(renders) == 1
    render = renders[0]
    assert render.find(qn("Data")) is not None and render.find(qn("CallbackID")) is None
    for attr in ("xpsFileIndex", "isPrintOut", "originalPageNumber"):
        assert render.get(attr) is None

    # Flattening is expected normalization (the render pages survive as images), NOT a fidelity
    # loss — it is deliberately NOT reported (user decision 2026-06-12). file_notes carries only
    # CONTENT-transfer losses (see test_clone_reports_unavailable_cache_explicitly).
    assert not any("printout" in note for note in result.file_notes)


def test_clone_reports_unavailable_cache_explicitly(be):
    result = copy.transfer_page(be, PAGE_1_ID, _TARGET_SECTION)
    sent = _sent_payload(be)

    # the docx cache fixture is deliberately absent → the attachment cannot be carried; a dead
    # reference would not self-heal, so it is REMOVED from the copy and the loss is reported.
    names = {f.get("preferredName") for f in sent.iter(qn("InsertedFile"))}
    assert "丘山行Word頁籤(中文) .docx" not in names
    assert result.missing_files == 1  # categorized as a file (no one:Previews → not embedded)
    [note] = [n for n in result.file_notes if "docx" in n]
    assert "removed from the copy" in note


def test_clone_embedded_object_gets_staged_source_and_keeps_previews(be):
    result = copy.transfer_page(be, PAGE_2_ID, _TARGET_SECTION)
    sent = _sent_payload(be)

    embedded = next(
        f
        for f in sent.iter(qn("InsertedFile"))
        if f.get("preferredName") == "附件與嵌入物件-2 - 工作表.xlsx"
    )
    # embedded objects have NO pathSource on read (ground truth) — the clone GAINS one,
    # pointing at the staged cache copy, so OneNote can re-import the workbook
    assert embedded.get("pathSource").startswith("C:\\FixtureStaging\\")
    assert embedded.get("pathCache") is None
    assert embedded.find(qn("Previews")) is not None  # preview structure rides verbatim

    jpg = next(f for f in sent.iter(qn("InsertedFile")) if f.get("preferredName") == "捷斯山屋.jpg")
    assert jpg.get("pathSource").endswith("捷斯山屋.jpg")
    assert result.file_notes == []  # both caches available — fully faithful copy
