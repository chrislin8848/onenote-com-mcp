"""OneNote COM enumerations.

Values are verbatim from the Microsoft "Enumerations (OneNote developer reference)" page
(see docs/com-api-reference.md). Pure data — fully testable on Linux.
"""

from __future__ import annotations

from enum import IntEnum


class HierarchyScope(IntEnum):
    """Lowest descendant level returned by GetHierarchy."""

    hsSelf = 0
    hsChildren = 1
    hsNotebooks = 2
    hsSections = 3
    hsPages = 4


class PageInfo(IntEnum):
    """What GetPageContent includes. piBasic omits binary + selection markup."""

    piBasic = 0
    piBinaryData = 1
    piSelection = 2
    piBinaryDataSelection = 3
    piFileType = 4
    piBinaryDataFileType = 5
    piSelectionFileType = 6
    piAll = 7


class CreateFileType(IntEnum):
    """OpenHierarchy create-if-missing target type."""

    cftNone = 0
    cftNotebook = 1
    cftFolder = 2  # section group
    cftSection = 3


class NewPageStyle(IntEnum):
    """CreateNewPage page style."""

    npsDefault = 0
    npsBlankPageWithTitle = 1
    npsBlankPageNoTitle = 2


class XMLSchema(IntEnum):
    """OneNote XML schema version. We always pin xs2013 (never xsCurrent)."""

    xs2007 = 0
    xs2010 = 1
    xs2013 = 2
    xsCurrent = 2  # same value as xs2013 by definition in this OneNote version


class SpecialLocation(IntEnum):
    """GetSpecialLocation targets."""

    slBackupFolder = 0
    slUnfiledNotesSection = 1
    slDefaultNotebookFolder = 2
