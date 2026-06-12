# tests/fixtures

Real OneNote XML samples used by the Tier 1 (host) parse/build tests.

**These are produced on the Windows VM**, not hand-authored (decision 2026-06-10: the
parse/build layer is TDD'd against ground-truth XML, not synthetic fixtures). Generate them
with:

```
# on the VM, inside the autologon interactive session:
python scripts/dump_fixtures.py
```

`dump_fixtures.py` must run via the `schtasks /it` task (COM needs the autologon interactive
session — a bare SSH command lands in session 0 where OneNote can't be activated). It calls live
COM (`GetHierarchy` + `GetPageContent` + `GetBinaryPageContent` + `FindPages`) and writes files
here using the filename convention documented in `src/onenote_com_mcp/backend/fixture.py`.

`FixtureBackend` then replays them on Linux so the host loop needs no COM.

## What's here (dumped 2026-06-11 from the "MCP Test" notebook, sanitized)

- `hierarchy_hs{Notebooks,Sections,Pages}__<MCP Test id>.xml` — scoped to the test notebook
  (no personal notebooks captured).
- `page_<id>.xml` / `page_<id>__binary.xml` — three content pages: mixed inline styles
  (incl. the highlight dual-attribute `background:` + `mso-highlight:`), a table (styled cells),
  and an image.
- `binary_<callback>.b64` — the image bytes (GetBinaryPageContent), base64 PNG.
- `page_…E191…/E186…` — the two Phase-5b attachment pages (附件與嵌入物件-1/-2, dumped
  2026-06-12): inline icons (txt/pdf/jpg), an inline printout (pdf + page-level XPSFile +
  render image), an inline embedded sheet (xlsx Previews), and two page-level icons
  (docx/xlsx) with their own objectIDs.
- `cachefile_<sanitized {GUID}.bin>.bin` — InsertedFile `pathCache` bytes. **NOT dumps:**
  hand-authored NEUTRAL files from `scripts/make_cache_fixtures.py` (the read path is a plain
  disk read, so synthetic bytes are faithful — and the PII policy forbids committing real user
  files). The docx cache is deliberately absent: replays "cache unavailable".
- `find__root.xml` — a FindPages sample (note: `_sanitize` maps non-ASCII queries to `root`).
- `current_window.json` — the Current*Id quadruple for `get_current_context`.

Sanitization done before commit: the table page was rebuilt with non-personal data, and the
OneDrive drive ID in `path=` attributes was replaced with zeros. Raw dumps may contain personal
note content — scrub before committing. Unscrubbed dumps go under `raw/` (gitignored).
