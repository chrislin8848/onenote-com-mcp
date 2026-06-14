# onenote-com-mcp

**Author:** Chris Lin

A **COM-only** Model Context Protocol server that gives Claude full CRUD over the live
**OneNote desktop** app on Windows — reading and writing notebooks, sections, pages, rich
text (formatting preserved), tables, and images.

No Azure app registration. No Microsoft Graph. No tokens. It drives the OneNote desktop COM
API in-process via pywin32, inside the signed-in user's own OneNote session, so it works for
any employee with synced OneNote and needs no admin setup.

> Why not Graph? Existing public OneNote MCP servers are almost all Graph-based and inherit a
> shared set of constraints: an Azure app + token, search/scale problems across large
> notebooks, and lossy HTML formatting. This project takes the COM route deliberately —
> full-fidelity round-trips, scoped queries, and zero cloud onboarding. See `SPEC` §1.2.

## Status — 2026-06-11

| Phase | What | State |
|---|---|---|
| **0a** | Host scaffold: backend interface, guarded import, FixtureBackend, server tool catalog, CI | ✅ done |
| **0b** | Windows 11 VM + autologon + secondary-session runner + `remote_test` + COM smoke | ✅ done |
| **1** | XML parse/build (lxml) + format preservation, TDD against **real VM fixtures** | ✅ done |
| **2** | Read tools incl. `get_current_context` on `FixtureBackend` (Linux green) | ✅ done |
| **3** | `Win32ComBackend` live + `dump_fixtures.py` + VM read-integration loop | ✅ done |
| 4 | Write tools + hierarchy restructure tools (whole-batch UpdateHierarchy; move_page VM-gated) + concurrency guard + format-preservation regression | ▶ **next** |
| 5 | Copy/transfer (raw-XML faithful copy) | ⛔ blocked |
| 6 | Delete + retry hardening + **PyInstaller/Inno installer** | ⛔ blocked |

Phase 1 (2026-06-11): `xmllayer/` parse + build implemented TDD-first against the real VM
dumps in `tests/fixtures/` — `parse_hierarchy` (mixed section/section-group nesting,
recycle-bin filtering), `parse_page` (three-layer effective style: `QuickStyleDef` →
OE `style` → inline span, highlight dual-attribute, structured tables, image callback IDs),
and `build_*` fragment builders that round-trip through the parser. Ground-truth findings
are recorded in [docs/onenote-xml-schema.md](docs/onenote-xml-schema.md).

Phase 2 (2026-06-11): the seven read tools (`list_notebooks` / `list_sections` / `list_pages`
/ `search_pages` / `get_page` / `get_page_images` / `get_current_context`) wired through
[`service/read.py`](src/onenote_com_mcp/service/read.py) onto `FixtureBackend`, Linux green.
`get_page` returns a lossless runs+style model with structured tables and every content
object's `objectID`; `get_page_images` returns MCP image content (binary via
`GetBinaryPageContent`, media type sniffed from the magic number); `get_current_context`
resolves the active window's four `Current*Id`s to names via one scoped `GetHierarchy`.
Ground-truth finding baked in: a `one:Image` has no `objectID` of its own — the deletable
object ID lives on the enclosing `one:OE` (matters for Phase 6 `delete_page_content`).

Phase 3 (2026-06-11): the live-COM read path is validated end-to-end on the VM. `Win32ComBackend`
+ `dump_fixtures.py` were already in place (commit 803ca55); this closed the loop with
[`tests/test_windows_read.py`](tests/test_windows_read.py) (Tier 2, `@pytest.mark.windows`) and a
tar-based [`scripts/remote_test.sh`](scripts/remote_test.sh) (ships code → triggers the `/it`
`onenote-tier2` task in the autologon session → polls → collects → mirrors the pytest exit code).
All 7 Tier-2 read round-trips pass on real OneNote — including the previously-open question that
`Windows.CurrentWindow`'s `Current*Id` properties marshal under early binding, and that the
highlight dual-attribute survives a live read (not just the committed dump).

## Develop (Linux host)

```bash
uv sync
uv run pytest                       # Tier 1; Windows COM tests auto-skip here
uv run ruff check . && uv run ruff format --check .
```

## Run (against fixtures, Linux)

```bash
ONENOTE_FIXTURES_DIR=tests/fixtures uv run python -m onenote_com_mcp
```

On Windows with OneNote installed, omit the env var to use live COM.

## Layout

```
src/onenote_com_mcp/    backend/ (OneNoteBackend + Fixture + Win32Com)  xmllayer/  service/  server.py
docs/               com-api-reference.md   onenote-xml-schema.md
scripts/            remote_test.sh   dump_fixtures.py        (VM bridge — Phase 0b/3)
tests/              Tier 1 unit tests + fixtures/
refs/               mhzarem reference clone (gitignored)
```

See `CLAUDE.md` for the architecture and the non-negotiable invariants.
