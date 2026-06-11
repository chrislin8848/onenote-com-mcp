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
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from onenote_com_mcp.backend.fixture import _sanitize
from onenote_com_mcp.backend.win32com_backend import Win32ComBackend
from onenote_com_mcp.enums import HierarchyScope, PageInfo
from onenote_com_mcp.errors import NoCurrentWindowError

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures"

_CALLBACK_RE = re.compile(r'callbackID="([^"]+)"')


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Dump real OneNote fixtures (run on the VM).")
    ap.add_argument("--scope", default="", help="start-node ID to scope the hierarchy to")
    ap.add_argument("--query", action="append", default=[], help="FindPages query (repeatable)")
    ap.add_argument("page_ids", nargs="*", help="page IDs to dump (piBasic + piBinaryData)")
    args = ap.parse_args(argv)

    OUT.mkdir(parents=True, exist_ok=True)
    be = Win32ComBackend()
    scope = args.scope

    # Hierarchy: notebooks → sections → pages, scoped if --scope was given (preferred filename).
    for scp in (HierarchyScope.hsNotebooks, HierarchyScope.hsSections, HierarchyScope.hsPages):
        xml = be.get_hierarchy(scope, scp)
        name = (
            f"hierarchy_{scp.name}__{_sanitize(scope)}.xml"
            if scope
            else f"hierarchy_{scp.name}.xml"
        )
        (OUT / name).write_text(xml, encoding="utf-8")
        print(f"wrote {name}")

    # Pages: piBasic + piBinaryData for each id; collect any image callbackIDs to fetch.
    callbacks: set[tuple[str, str]] = set()
    for pid in args.page_ids:
        sid = _sanitize(pid)
        (OUT / f"page_{sid}.xml").write_text(be.get_page_content(pid), encoding="utf-8")
        binary = be.get_page_content(pid, PageInfo.piBinaryData)
        (OUT / f"page_{sid}__binary.xml").write_text(binary, encoding="utf-8")
        for cb in _CALLBACK_RE.findall(binary):
            callbacks.add((pid, cb))
        print(f"wrote page_{sid}.xml (+ __binary)")

    # Image binaries referenced by callbackID (skipped when the bytes are inlined in __binary).
    for pid, cb in sorted(callbacks):
        (OUT / f"binary_{_sanitize(cb)}.b64").write_text(
            be.get_binary_page_content(pid, cb), encoding="utf-8"
        )
        print(f"wrote binary_{_sanitize(cb)}.b64")

    # FindPages samples (scoped to the same start node).
    for q in args.query:
        (OUT / f"find__{_sanitize(q)}.xml").write_text(be.find_pages(scope, q), encoding="utf-8")
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
    (OUT / "current_window.json").write_text(payload, encoding="utf-8")
    print("wrote current_window.json")

    print(f"\nFixtures in {OUT}. Scrub personal content before committing.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
