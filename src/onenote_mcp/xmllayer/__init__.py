"""OneNote XML layer — pure functions over OneNote 2013 ``one:`` XML.

This is the core of the project (SPEC §3). Everything here is a pure function operating on
strings / lxml trees, with **no COM dependency**, so it is fully TDD-able on Linux against
fixtures dumped from the VM. See docs/onenote-xml-schema.md.

- ``parse``  — GetHierarchy / GetPageContent XML → structured data (preserving style).
- ``build``  — structured input → CreateNewPage / UpdatePageContent payloads (with style).
"""
