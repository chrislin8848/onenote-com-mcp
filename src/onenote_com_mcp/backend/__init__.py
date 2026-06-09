"""Backend selection.

``OneNoteBackend`` is the abstract COM surface. ``FixtureBackend`` is always importable
(Linux). ``Win32ComBackend`` is imported lazily so this package imports cleanly on Linux.
"""

from __future__ import annotations

import os
import sys

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.backend.fixture import FixtureBackend

__all__ = ["OneNoteBackend", "FixtureBackend", "get_backend"]


def get_backend() -> OneNoteBackend:
    """Return the backend appropriate for the current environment.

    * ``ONENOTE_FIXTURES_DIR`` set → ``FixtureBackend`` (offline replay; honored on any OS).
    * Windows → ``Win32ComBackend`` (live COM).
    * otherwise → error (no COM on this platform and no fixtures configured).
    """
    fixtures_dir = os.environ.get("ONENOTE_FIXTURES_DIR")
    if fixtures_dir:
        return FixtureBackend(fixtures_dir)

    if sys.platform == "win32":
        from onenote_com_mcp.backend.win32com_backend import Win32ComBackend  # noqa: PLC0415

        return Win32ComBackend()

    raise RuntimeError(
        "No OneNote backend available: not on Windows and ONENOTE_FIXTURES_DIR is unset. "
        "Set ONENOTE_FIXTURES_DIR to replay fixtures on Linux, or run on Windows for live COM."
    )
