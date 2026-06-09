"""Enum values must match the Microsoft reference exactly (docs/com-api-reference.md)."""

from __future__ import annotations

from onenote_mcp.enums import (
    CreateFileType,
    HierarchyScope,
    NewPageStyle,
    PageInfo,
    SpecialLocation,
    XMLSchema,
)


def test_hierarchy_scope():
    assert (HierarchyScope.hsSelf, HierarchyScope.hsChildren) == (0, 1)
    assert (HierarchyScope.hsNotebooks, HierarchyScope.hsSections, HierarchyScope.hsPages) == (
        2,
        3,
        4,
    )


def test_page_info():
    assert PageInfo.piBasic == 0
    assert PageInfo.piBinaryData == 1
    assert PageInfo.piAll == 7


def test_create_file_type():
    assert (CreateFileType.cftNone, CreateFileType.cftNotebook) == (0, 1)
    assert (CreateFileType.cftFolder, CreateFileType.cftSection) == (2, 3)


def test_new_page_style():
    assert NewPageStyle.npsDefault == 0


def test_xml_schema_2013():
    assert XMLSchema.xs2013 == 2
    # xsCurrent shares the 2013 value in this OneNote version.
    assert XMLSchema.xsCurrent == XMLSchema.xs2013


def test_special_location():
    assert SpecialLocation.slBackupFolder == 0
    assert SpecialLocation.slDefaultNotebookFolder == 2
