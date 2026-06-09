# tests/fixtures

Real OneNote XML samples used by the Tier 1 (host) parse/build tests.

**These are produced on the Windows VM**, not hand-authored (decision 2026-06-10: the
parse/build layer is TDD'd against ground-truth XML, not synthetic fixtures). Generate them
with:

```
# on the VM, inside the autologon interactive session:
python scripts/dump_fixtures.py
```

`dump_fixtures.py` calls live COM (`GetHierarchy` + a handful of `GetPageContent`, covering
plain text, mixed inline styles, a table, and an image) and writes files here using the
filename convention documented in `src/onenote_com_mcp/backend/fixture.py`.

`FixtureBackend` then replays them on Linux so the host loop needs no COM.

Raw dumps may contain personal note content — scrub before committing. Unscrubbed dumps go
under `raw/` (gitignored).
