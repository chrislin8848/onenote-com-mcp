"""dump_fixtures.py — capture real OneNote XML into tests/fixtures (runs on the VM).

Run inside the autologon interactive session on the Windows VM (COM needs the desktop session;
drive it via the ``schtasks /it`` task, not a bare SSH command — SSH lands in session 0 where
CoCreateInstance can't launch OneNote). It uses the same guarded ``Win32ComBackend`` the server
uses, so whatever it captures is exactly what the parser will see in production.

Goal: cover the formatting cases the parse/build layer must preserve —
  * a page with plain text + mixed inline span styles (bold/color/highlight/font),
  * a page containing a table,
  * a page containing an image (piBinaryData inlines the bytes; if the image is referenced by a
    ``callbackID`` instead, GetBinaryPageContent is pulled into binary_<callback>.b64),
so Phase 1 can be TDD'd on Linux against ground truth (not synthetic XML).

Scope the hierarchy to a start node (e.g. the test notebook's ID) with ``--scope`` so personal
notebooks are never captured. Filenames follow the convention in
``src/onenote_com_mcp/backend/fixture.py`` so FixtureBackend can replay them directly.

Usage:
  python scripts/dump_fixtures.py --scope "{NB-ID}{1}{B0}" --query "藥" \
      "{PAGE-ID-1}..." "{PAGE-ID-2}..." "{PAGE-ID-3}..."
  # or resolve pages by NAME (UTF-8 file, one page name per line — avoids cmd.exe
  # codepage mangling for non-ASCII names on the schtasks command line):
  python scripts/dump_fixtures.py --scope "{NB-ID}{1}{B0}" --names-file scripts/dump_pages.txt \
      --out test-results/dump
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from lxml import etree

from onenote_com_mcp.backend.fixture import _sanitize
from onenote_com_mcp.backend.win32com_backend import Win32ComBackend
from onenote_com_mcp.enums import HierarchyScope, PageInfo
from onenote_com_mcp.errors import NoCurrentWindowError
from onenote_com_mcp.xmllayer.namespaces import qn

DEFAULT_OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures"

_CALLBACK_RE = re.compile(r'callbackID="([^"]+)"')


def _resolve_page_ids_by_name(pages_xml: str, names: list[str]) -> list[str]:
    """Map page names → IDs from a scoped hsPages hierarchy dump. Hard-fails on a miss."""
    tree = etree.fromstring(pages_xml.encode("utf-8"))
    by_name: dict[str, str] = {}
    for page in tree.iter(qn("Page")):
        name, pid = page.get("name"), page.get("ID")
        if name and pid:
            by_name.setdefault(name, pid)
    missing = [n for n in names if n not in by_name]
    if missing:
        raise SystemExit(
            f"pages not found in scope: {missing}; available: {sorted(by_name)}"
        )
    return [by_name[n] for n in names]


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Dump real OneNote fixtures (run on the VM).")
    ap.add_argument("--scope", default="", help="start-node ID to scope the hierarchy to")
    ap.add_argument("--query", action="append", default=[], help="FindPages query (repeatable)")
    ap.add_argument(
        "--names-file",
        default="",
        help="UTF-8 file of page NAMES (one per line) to resolve to IDs within --scope",
    )
    ap.add_argument(
        "--out",
        default=str(DEFAULT_OUT),
        help="output dir (default tests/fixtures; point elsewhere to avoid clobbering "
        "already-sanitized fixtures, then move files over by hand after inspection)",
    )
    ap.add_argument("page_ids", nargs="*", help="page IDs to dump (piBasic + piBinaryData)")
    args = ap.parse_args(argv)

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    be = Win32ComBackend()
    scope = args.scope

    # Hierarchy: notebooks → sections → pages, scoped if --scope was given (preferred filename).
    pages_xml = ""
    for scp in (HierarchyScope.hsNotebooks, HierarchyScope.hsSections, HierarchyScope.hsPages):
        xml = be.get_hierarchy(scope, scp)
        if scp is HierarchyScope.hsPages:
            pages_xml = xml
        name = (
            f"hierarchy_{scp.name}__{_sanitize(scope)}.xml"
            if scope
            else f"hierarchy_{scp.name}.xml"
        )
        (out / name).write_text(xml, encoding="utf-8")
        print(f"wrote {name}")

    page_ids = list(args.page_ids)
    if args.names_file:
        names = [
            line.strip()
            for line in Path(args.names_file).read_text(encoding="utf-8").splitlines()
            if line.strip()
        ]
        resolved = _resolve_page_ids_by_name(pages_xml, names)
        for n, pid in zip(names, resolved, strict=True):
            print(f"resolved {n!r} -> {pid}")
        page_ids.extend(resolved)

    # Pages: piBasic + piBinaryData for each id; collect any image callbackIDs to fetch.
    callbacks: set[tuple[str, str]] = set()
    for pid in page_ids:
        sid = _sanitize(pid)
        (out / f"page_{sid}.xml").write_text(be.get_page_content(pid), encoding="utf-8")
        binary = be.get_page_content(pid, PageInfo.piBinaryData)
        (out / f"page_{sid}__binary.xml").write_text(binary, encoding="utf-8")
        for cb in _CALLBACK_RE.findall(binary):
            callbacks.add((pid, cb))
        print(f"wrote page_{sid}.xml (+ __binary)")

    # Image binaries referenced by callbackID (skipped when the bytes are inlined in __binary).
    for pid, cb in sorted(callbacks):
        (out / f"binary_{_sanitize(cb)}.b64").write_text(
            be.get_binary_page_content(pid, cb), encoding="utf-8"
        )
        print(f"wrote binary_{_sanitize(cb)}.b64")

    # FindPages samples (scoped to the same start node).
    for q in args.query:
        (out / f"find__{_sanitize(q)}.xml").write_text(be.find_pages(scope, q), encoding="utf-8")
        print(f"wrote find__{_sanitize(q)}.xml")

    # Current viewing context (get_current_context fixture). With no open window we record the
    # no-window case as literal null.
    try:
        ids = be.get_current_window_ids()
        payload = json.dumps(
            {
                "notebook_id": ids.notebook_id,
                "section_group_id": ids.section_group_id,
                "section_id": ids.section_id,
                "page_id": ids.page_id,
            },
            indent=2,
        )
    except NoCurrentWindowError:
        payload = "null"
    (out / "current_window.json").write_text(payload, encoding="utf-8")
    print("wrote current_window.json")

    print(f"\nFixtures in {out}. Scrub personal content before committing.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
