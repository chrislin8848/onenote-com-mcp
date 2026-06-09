"""dump_fixtures.py — capture real OneNote XML into tests/fixtures (runs on the VM).

STATUS: skeleton. Finalized in Phase 3. Run inside the autologon interactive session on the
Windows VM (COM needs the desktop session). It uses the same guarded Win32ComBackend the
server uses, so whatever it captures is exactly what the parser will see in production.

Goal: cover the formatting cases the parse/build layer must preserve —
  * a page with plain text + mixed inline span styles (bold/color/highlight/font),
  * a page containing a table,
  * a page containing an image (dump both piBasic and piBinaryData variants),
so Phase 1 can be TDD'd on Linux against ground truth (not synthetic XML).

Filenames follow the convention in src/onenote_mcp/backend/fixture.py so FixtureBackend can
replay them directly. Scrub personal content before committing (raw/ is gitignored).
"""

from __future__ import annotations

import sys
from pathlib import Path

from onenote_mcp.backend.fixture import _sanitize
from onenote_mcp.backend.win32com_backend import Win32ComBackend
from onenote_mcp.enums import HierarchyScope, PageInfo

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures"


def main(page_ids: list[str]) -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    be = Win32ComBackend()

    # Hierarchy: notebooks → sections → pages.
    for scope in (HierarchyScope.hsNotebooks, HierarchyScope.hsSections, HierarchyScope.hsPages):
        xml = be.get_hierarchy("", scope)
        (OUT / f"hierarchy_{scope.name}.xml").write_text(xml, encoding="utf-8")
        print(f"wrote hierarchy_{scope.name}.xml")

    # Pages: capture piBasic and a binary variant for each id given on the CLI.
    for pid in page_ids:
        sid = _sanitize(pid)
        (OUT / f"page_{sid}.xml").write_text(be.get_page_content(pid), encoding="utf-8")
        (OUT / f"page_{sid}__binary.xml").write_text(
            be.get_page_content(pid, PageInfo.piBinaryData), encoding="utf-8"
        )
        print(f"wrote page_{sid}.xml (+ __binary)")

    # TODO(Phase 3): also pull GetBinaryPageContent for each image callbackID found in the
    # binary page XML → binary_<callback>.b64, and a couple of FindPages results.
    print(f"\nFixtures in {OUT}. Scrub personal content before committing.")
    return 0


if __name__ == "__main__":
    # Usage: python scripts/dump_fixtures.py "{PAGE-ID-1}{1}{B0}" "{PAGE-ID-2}{1}{B0}" ...
    raise SystemExit(main(sys.argv[1:]))
