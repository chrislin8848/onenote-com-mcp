"""Tier 2 — real working-data smoke (read + copy), Windows only.

Chris's actual sections 測試章節1 / 測試章節2 are big, image-heavy, attachment-heavy real
notebooks — exactly the kind of load that exposed the OCR-image copy bug. This is a
NON-DESTRUCTIVE pass over them:
  * every page is read through all four read tools (get_page, get_page_images,
    get_page_files_info, get_page_files) — assert they never crash and report light stats;
  * the whole section is copied into the MCP Test notebook (a throwaway, de-collided name),
    then the copy's page count / names / levels are checked against the source.

The source notebooks are NEVER modified. The throwaway copy is deleted in teardown.

PII policy: this test logs only COUNTS and STATUS strings — never file text, never image
bytes, never page content. tier2.log is collected to the host, so nothing readable rides.
"""

from __future__ import annotations

import sys

import pytest

from onenote_com_mcp.service import copy, delete, files, hierarchy_edit, read

pytestmark = pytest.mark.windows

TEST_NOTEBOOK = "MCP Test"
# substrings — the real names carry date prefixes / suffixes; match loosely
TARGET_SECTIONS = ["測試章節1", "測試章節2"]
# throwaway copies this test makes (and must never mistake for a source)
COPY_PREFIX = "實測複製-"


def _find(nodes, name):
    return next((n for n in nodes if n.get("name") == name), None)


def _walk_sections(nodes):
    """Yield (section_node) for every section anywhere in a notebook's tree (recurse groups)."""
    for node in nodes:
        if node.get("type") == "section":
            yield node
        for child in node.get("children", []) or []:
            yield from _walk_sections([child])


def _find_real_sections(backend):
    """Locate the target real-data sections across ALL notebooks.

    The real sources now live in MCP Test, so MCP Test is INCLUDED — only this test's own
    throwaway copies (COPY_PREFIX) are skipped so a leftover copy can't pose as a source.
    """
    found = {}
    for nb in read.list_notebooks(backend):
        for sec in _walk_sections(read.list_sections(backend, nb["id"])):
            if sec["name"].startswith(COPY_PREFIX):
                continue
            for needle in TARGET_SECTIONS:
                if needle in sec["name"] and needle not in found:
                    found[needle] = {"id": sec["id"], "name": sec["name"], "notebook": nb["name"]}
    return found


@pytest.fixture(scope="module")
def backend():
    from onenote_com_mcp.backend.win32com_backend import Win32ComBackend

    return Win32ComBackend()


@pytest.fixture(scope="module")
def test_notebook_id(backend):
    nb = _find(read.list_notebooks(backend), TEST_NOTEBOOK)
    if nb is None:
        pytest.skip(f"{TEST_NOTEBOOK!r} notebook not present on this VM")
    # sweep any leftover throwaway copies from a previous real-data run
    for sec in _walk_sections(read.list_sections(backend, nb["id"])):
        if sec["name"].startswith(COPY_PREFIX):
            try:
                backend.delete_hierarchy(sec["id"], permanent=True)
            except Exception as exc:  # noqa: BLE001 — best-effort sweep
                print(f"leftover real-data copy sweep failed: {exc}", file=sys.stderr)
    return nb["id"]


@pytest.fixture(scope="module")
def real_sections(backend):
    found = _find_real_sections(backend)
    if not found:
        pytest.skip(f"none of {TARGET_SECTIONS} present on this VM")
    for needle, info in found.items():
        print(
            f"[realdata] {needle}: '{info['name']}' in notebook '{info['notebook']}' "
            f"(id={info['id']})",
            file=sys.stderr,
        )
    return found


@pytest.mark.parametrize("needle", TARGET_SECTIONS)
def test_read_every_page(backend, real_sections, needle):
    """All four read tools run over every page of the real section without crashing."""
    if needle not in real_sections:
        pytest.skip(f"{needle} not present")
    section_id = real_sections[needle]["id"]
    pages = read.list_pages(backend, section_id)
    assert pages, f"{needle} section has no pages"

    total_imgs = total_files = 0
    file_status_counts: dict[str, int] = {}
    for p in pages:
        pid = p["id"]
        page = read.get_page(backend, pid)  # structured read
        assert page["title"] is not None or page["outlines"] is not None
        imgs = read.get_page_images(backend, pid)
        info = files.get_page_files_info(backend, pid)
        contents = files.get_page_files(backend, pid)  # text/image/PDF extraction
        total_imgs += len(imgs)
        total_files += len(info)
        for c in contents:
            st = c.get("status", "?")
            file_status_counts[st] = file_status_counts.get(st, 0) + 1
        # every reported file must carry a deletable object_id (get_page_files_info contract)
        for f in info:
            assert "object_id" in f
    print(
        f"[realdata] {needle}: {len(pages)} pages, {total_imgs} images, "
        f"{total_files} files, file-status={file_status_counts}",
        file=sys.stderr,
    )


@pytest.mark.parametrize("needle", TARGET_SECTIONS)
def test_copy_section_then_delete(backend, real_sections, test_notebook_id, needle):
    """Copy the whole real section into MCP Test; verify structure; then DELETE it and
    confirm it is gone (exercises the delete tool on a real-data-sized section)."""
    if needle not in real_sections:
        pytest.skip(f"{needle} not present")
    src_id = real_sections[needle]["id"]
    src_pages = read.list_pages(backend, src_id)

    # transfer_section takes the SOURCE name (de-collided). Immediately rename the copy to the
    # COPY_PREFIX throwaway name so the module sweep can reclaim it even if this aborts mid-way.
    result = copy.transfer_section(backend, src_id, test_notebook_id)
    new_id = result.section_id
    deleted = False
    try:
        hierarchy_edit.rename_node(backend, test_notebook_id, new_id, f"{COPY_PREFIX}{needle}")
        copied = read.list_pages(backend, new_id)
        assert [p["name"] for p in copied] == [p["name"] for p in src_pages], (
            f"{needle}: copied page names/order must match the source"
        )
        assert [p["page_level"] for p in copied] == [p["page_level"] for p in src_pages], (
            f"{needle}: subpage levels must survive the copy"
        )
        print(
            f"[realdata] {needle}: copied {len(copied)} pages OK; "
            f"{len(result.file_notes)} file_notes",
            file=sys.stderr,
        )
        # file_notes are expected (OCR images → placeholders) but must be well-formed strings
        for note in result.file_notes:
            assert isinstance(note, str) and note

        # --- delete the copy through the real tool path, then confirm it's gone -----------
        delete.delete_node(backend, new_id, permanent=True)
        deleted = True
        remaining = {s["id"] for s in _walk_sections(read.list_sections(backend, test_notebook_id))}
        assert new_id not in remaining, f"{needle}: deleted copy must vanish from the section list"
        print(f"[realdata] {needle}: copy deleted OK (gone from MCP Test)", file=sys.stderr)
    finally:
        if not deleted:  # backstop only if the asserted delete above never ran
            try:
                backend.delete_hierarchy(new_id, permanent=True)
            except Exception as exc:  # noqa: BLE001 — best-effort cleanup
                print(f"backstop cleanup of {new_id} failed: {exc}", file=sys.stderr)
