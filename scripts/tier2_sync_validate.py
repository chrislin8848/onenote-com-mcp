"""Tier-2 live validation of the copy REMOVE-un-synced-content behavior (interactive COM session).

Targets ONE section (default "祕魯18天"; override via argv[1] substring) that is deliberately not
fully synced, to exercise BOTH the image path AND the still-open file/object (attachment) path of
the copy sync_warning. Confirms: (1) transfer_section COMPLETES (no hrInvalidXML), (2) sync_warning
reports the counts, (3) the copy has FEWER one:Image elements than the source (removal happened, not
placeholdering), (4) un-synced attachments are categorized into missing_files / missing_objects.

PII-safe: counts + warning + note labels only, never file/page content. Deletes its own throwaway
copy at the end (permanent) so MCP Test stays clean.
"""

# ruff: noqa: UP031  (throwaway VM probe — %-format is fine, not shipped)

import sys
import traceback

from onenote_com_mcp.backend.win32com_backend import Win32ComBackend
from onenote_com_mcp.enums import PageInfo
from onenote_com_mcp.errors import OneNoteComError
from onenote_com_mcp.service import copy, files, read
from onenote_com_mcp.xmllayer.parse import parse_page

TARGET = sys.argv[1] if len(sys.argv) > 1 else "祕魯18天"


def collect(nodes, pfx=""):
    out = []
    for n in nodes:
        if n.get("type") == "section":
            out.append((pfx + n["name"], n["id"]))
        elif n.get("type") == "section_group":
            out += collect(n.get("children", []), pfx + n["name"] + " / ")
    return out


def image_census(be, section_id):
    """(image-element count, fetch-failures) across the section's pages."""
    total = fail = 0
    for pg in read.list_pages(be, section_id):
        try:
            page = parse_page(be.get_page_content(pg["id"], PageInfo.piBasic))
        except Exception:  # noqa: BLE001
            continue
        for img in page.images:
            if not img.callback_id:
                continue
            total += 1
            try:
                be.get_binary_page_content(pg["id"], img.callback_id)
            except OneNoteComError:
                fail += 1
    return total, fail


def attachment_census(be, section_id):
    """(attachment count, cache-unavailable count) — the file/object un-synced signal."""
    total = unavailable = 0
    for pg in read.list_pages(be, section_id):
        try:
            info = files.get_page_files_info(be, pg["id"])
        except Exception:  # noqa: BLE001
            continue
        for f in info:
            total += 1
            # cache_available False ⇒ binary not hydrated on this machine (un-synced)
            if not f.get("cache_available", False):
                unavailable += 1
    return total, unavailable


def count_image_elements(be, section_id):
    total = 0
    for pg in read.list_pages(be, section_id):
        try:
            page = parse_page(be.get_page_content(pg["id"], PageInfo.piBasic))
        except Exception:  # noqa: BLE001
            continue
        total += len(page.images)
    return total


def main():
    print("python", sys.version.split()[0])
    print("target substring:", TARGET)
    be = Win32ComBackend()
    mcp = next(nb for nb in read.list_notebooks(be) if nb["name"] == "MCP Test")
    secs = collect(read.list_sections(be, mcp["id"]))
    match = next(((name, sid) for name, sid in secs if TARGET in name), None)
    if match is None:
        print("!!! no section matching %r in MCP Test. Sections present:" % TARGET)
        for name, _ in secs:
            print("   -", name)
        return
    name, sid = match
    pages = read.list_pages(be, sid)
    img_total, img_fail = image_census(be, sid)
    att_total, att_unavail = attachment_census(be, sid)
    src_imgs = count_image_elements(be, sid)
    print("\n>>> SECTION: %s" % name)
    print("    pages=%d  image_callbacks=%d  fetch_fail=%d" % (len(pages), img_total, img_fail))
    print("    attachments=%d  cache_unavailable=%d" % (att_total, att_unavail))
    print("    source one:Image elements=%d" % src_imgs)

    print("\n>>> COPYING into MCP Test ...")
    try:
        result = copy.transfer_section(be, sid, mcp["id"])
    except Exception:  # noqa: BLE001
        print("!!! transfer_section RAISED (hrInvalidXML?) — removal payload REJECTED:")
        traceback.print_exc()
        return

    print("\n=== COPY COMPLETED (no hrInvalidXML) ===")
    print("new section id :", result.section_id)
    print("missing_images :", result.missing_images)
    print("missing_files  :", result.missing_files)
    print("missing_objects:", result.missing_objects)
    print("sync_warning   ->")
    warn = copy.sync_warning(result.missing_images, result.missing_files, result.missing_objects)
    print("  ", warn)
    copy_imgs = count_image_elements(be, result.section_id)
    print(
        "\nimage-element removal check: source=%d copy=%d removed=%d (expected ~ missing_images=%d)"
        % (src_imgs, copy_imgs, src_imgs - copy_imgs, result.missing_images)
    )
    print("\nfile/object (attachment) notes:")
    found = False
    for note in result.file_notes:
        if ("file content" in note) or ("embedded object" in note) or ("not yet" in note):
            print("   -", note)
            found = True
    if not found:
        print("   (none flagged — no attachments were un-synced)")

    # clean up the throwaway copy so MCP Test stays tidy
    try:
        be.delete_hierarchy(result.section_id, permanent=True)
        gone = result.section_id not in {s for _, s in collect(read.list_sections(be, mcp["id"]))}
        print("\ncleanup: copy deleted, gone-from-list=%s" % gone)
    except Exception:  # noqa: BLE001
        print("\ncleanup: delete failed (leftover copy %s — delete by hand)" % result.section_id)


if __name__ == "__main__":
    main()
