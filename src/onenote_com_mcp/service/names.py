"""Notebook / section / section-group name validation (service layer).

Shared by create (new nodes) and hierarchy_edit (rename): a section name becomes a ``.one``
filename and notebook/group names become folder names in the sync URL, so OneNote rejects
these characters server-side — better to fail fast before any COM call.
"""

from __future__ import annotations

_INVALID_NAME_CHARS = set('\\/:*?"<>|&#%~')


def checked_name(name: str, kind: str) -> str:
    """Validate and normalize a node name; returns the stripped name or raises ValueError."""
    name = name.strip()
    if not name:
        raise ValueError(f"{kind} name is empty")
    bad = _INVALID_NAME_CHARS & set(name)
    if bad:
        raise ValueError(
            f"{kind} name contains characters OneNote forbids: {' '.join(sorted(bad))}"
        )
    return name
