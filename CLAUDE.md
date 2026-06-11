# CLAUDE.md — OneNote MCP server

## The one red line (highest priority — SPEC §9)
**COM-only.** No Microsoft Graph, no `msal`, no `requests`/HTTP to `graph.microsoft.com`, no
Azure app/token, no OAuth. All OneNote access goes through `OneNoteBackend → pywin32 COM`.
If you find yourself writing `import msal`, an HTTP call, or token logic — stop, you're off
the path. Existing public OneNote MCP servers are nearly all Graph-based; we are deliberately not.

## What this is
A COM-only MCP server giving Claude full CRUD over the live OneNote **desktop** app on
Windows, preserving rich-text formatting, tables, and images. No backup-file parsing, no
PowerShell subprocess (those are the reference repo's approach, which we reject — SPEC §1.1).

## Architecture (SPEC §3) — dependencies point downward
```
MCP layer      src/onenote_com_mcp/server.py      FastMCP tool catalog (thin facades)
Service layer  src/onenote_com_mcp/service/       orchestration; shared write core + copy core
XML layer      src/onenote_com_mcp/xmllayer/      pure parse/build over one: XML  ← core, TDD
Backend layer  src/onenote_com_mcp/backend/       OneNoteBackend ABC
                 ├ fixture.py        FixtureBackend  (Linux, replays VM dumps)
                 └ win32com_backend  Win32ComBackend (Windows, guarded import)
```

## Two invariants that keep the host loop alive (SPEC §2.1)
1. **Guarded import.** Never `import win32com`/`pywintypes`/`pythoncom` at module top level.
   Import them inside methods. `tests/test_smoke_import.py` enforces this — keep it green.
2. **Pure XML layer.** parse/build are pure functions over strings/lxml. No COM. Fully
   testable on Linux against fixtures.

## Two code paths that must never cross (SPEC §5)
- **Editing** (`update_page_content`, `create_table`, `insert_image`): `GetPageContent →
  mutate the real lxml tree IN PLACE → UpdatePageContent`. Never rebuild from a slimmed model
  (drops formatting on untouched paragraphs). One shared core; tools are facades.
- **Copy** (`copy_*`): raw-XML whole-page transfer — inline image binary (`piBinaryData`),
  carry the `QuickStyleDef` table, reset object IDs, set `pageLevel`. Do NOT route through the
  structured/`get_page` representation. "Copy then modify" = copy faithfully, then edit the
  copy via the editing path.

**Enforced, not just documented:** the single edit path is `apply_page_edit()` in
`service/page_edit.py` — the ONLY `update_page_content` call site; `edit_page_content`/`add_table`/
`insert_image` delegate to it. Copy is `transfer_page()` in `service/copy.py` (reads
`piBinaryData`, must not import the parse layer). Guard tests `tests/test_write_core.py` +
`tests/test_copy_path.py` fail if either invariant breaks — keep them green; build the Phase 4/5
content logic *inside* these seams, never around them.

## Third discipline: hierarchy restructure = whole batch (SPEC §5)
`UpdateHierarchy` order = child-element order of the submitted XML; a *partial* child list makes
OneNote "infer" omitted siblings' placement unpredictably. So structural changes
(`restructure_section`/`reorder_sections`/`rename_node`/`move_page`) submit the scope's
**complete child list, target order, one batch** — enforced by `apply_hierarchy_restructure()`
in `service/hierarchy_edit.py`: the ONLY `update_hierarchy` call site for restructures, with
node-ID conservation (drop/invent ⇒ `ValueError`). Guard tests `tests/test_hierarchy_core.py`.
`move_page` is EXPERIMENTAL until VM-validated; notebook-level ordering is out of scope.
Structure tools are propose-then-confirm + suggest clone backup first.

## Section groups + current context (SPEC v0611)
- A notebook's direct children are a **mixed** `one:Section` + `one:SectionGroup` list: listings
  keep the nesting (no flattening, no extra tool), `create_section` parent may be a group,
  `copy_notebook` recreates groups via `OpenHierarchy(cftFolder)`, and whole-batch restructures
  must include both kinds (ID conservation already enforces this — see test_hierarchy_core.py).
- `get_current_context` (Phase 2) = `Windows.CurrentWindow` four Current*Ids
  (backend `get_current_window_ids()`) + scoped GetHierarchy for names. Three limits (SPEC §5):
  no open window ⇒ `NoCurrentWindowError` (never guess); granularity stops at the page (no
  cursor/selection API); report "you're on page X" back before acting on it.

## Format preservation = self-consistency, not byte-identity
OneNote re-normalizes spans / renumbers `QuickStyleDef` on redraw. Verify the user-visible
font/size/color of *untouched* paragraphs is unchanged — not raw-XML equality. Highlight is
dual-attribute: write BOTH `background:` and `mso-highlight:`. See docs/onenote-xml-schema.md.

## Concurrency + safety (SPEC §5, §7)
Always pass `dateExpectedLastModified` (the page's `lastModifiedTime` from the read) on
writes/deletes; default `force=False`. Surface conflicts as `ConcurrencyError`. Retry
`RPC_E_SERVERCALL_RETRYLATER` / `RPC_E_CALL_REJECTED` with backoff (`_call` in win32 backend).

## Dev commands (host / Linux)
```
uv sync                         # install deps + dev group
uv run pytest                   # Tier 1 (windows-marked tests auto-skip here)
uv run ruff check . && uv run ruff format --check .
ONENOTE_FIXTURES_DIR=tests/fixtures uv run python -m onenote_com_mcp   # run server on fixtures
```

## Testing tiers (SPEC §2, §6)
- **Tier 1 (host, every change):** pure logic + fixtures. Must stay green on Linux.
- **Tier 2 (VM, checkpoint):** `@pytest.mark.windows`, real COM round-trips. Auto-skipped off
  Windows. Driven by `scripts/remote_test.sh` once the VM exists.

## Status (2026-06-11)
Phases 0a, 0b, 1, and **2 done**. The VM is up (COM smoke passed), real fixtures are dumped into
`tests/fixtures/`, the XML layer is implemented TDD-first against them (`spans.py`, `models.py`,
`parse.py` three-layer effective style, `build.py` highlight dual-write), and the seven read
tools are wired through `service/read.py` onto `FixtureBackend` (Linux green). All three seams
remain guard-tested. Pushed to GitHub (`chrislin8848/onenote-com-mcp`, private).

Phase 2 shape: `service/read.py` does the orchestration (backend call → parse → project →
resolve IDs to names), pure-Python and testable on Linux; `server.py` tools are thin facades
(`json.dumps` for text, FastMCP `Image` content for `get_page_images` — `@mcp.tool(
structured_output=False)` since `Image` has no pydantic output schema). `get_page` is the
lossless runs+style model (effective style per run, structured tables, every content object's
`objectID`); `get_current_context` resolves the window's four `Current*Id`s via one scoped
`GetHierarchy(notebook, hsPages)`. Ground truth: a `one:Image` has NO `objectID` — the deletable
ID is on the enclosing `one:OE` (Phase 6 `delete_page_content` must target the OE).

**Next (Phase 3):** `Win32ComBackend` + `dump_fixtures.py` are already done (commit 803ca55);
remaining is the VM read-integration loop via `scripts/remote_test.sh` (`@pytest.mark.windows`
round-trips), plus the deferred `remote_test.sh` rsync→tar fix (guest has no rsync) and the
formal `onenote-tier2` schtasks task. Trust fixtures over the schema sketch.

## Grounding
- `docs/SPEC.md` — the spec itself (v0611: + get_current_context, section-group integration).
- `docs/com-api-reference.md` — COM signatures + enums (from Microsoft Learn).
- `docs/onenote-xml-schema.md` — `one:` page/hierarchy XML + format-preservation rules.
- `docs/vm-setup.md` — Phase 0b Windows VM build (autologon, desktop OneNote, COM smoke, Tier-2).
- `refs/mhzarem-onenote-mcp/` — reference clone (gitignored). Payload shapes only; its
  backup-parse + PowerShell model is rejected.
