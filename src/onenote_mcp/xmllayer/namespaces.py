"""OneNote XML namespace helpers. Pure — testable on Linux."""

from __future__ import annotations

# OneNote 2013 schema. We pin this everywhere (never the 2007 .../12/2004/ namespace).
ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

# Map for lxml find/xpath calls: tree.findall("one:Outline", NSMAP)
NSMAP = {"one": ONE_NS}


def qn(tag: str) -> str:
    """Clark-notation qualified name for a ``one:`` element, e.g. qn("Page").

    >>> qn("Outline")
    '{http://schemas.microsoft.com/office/onenote/2013/onenote}Outline'
    """
    return f"{{{ONE_NS}}}{tag}"


def local_name(tag: str) -> str:
    """Strip the ``{ns}`` prefix from a Clark-notation tag, returning the local name.

    >>> local_name(qn("Outline"))
    'Outline'
    >>> local_name("Outline")
    'Outline'
    """
    return tag.rsplit("}", 1)[-1]
