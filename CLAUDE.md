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

## Status (2026-06-10)
Phase 0a (host scaffold) done; write-core convergence + copy direct-path seams in place and
guard-tested (29 Tier-1 tests green); all pushed to GitHub (`chrislin8848/onenote-com-mcp`,
private). **Blocked on the Windows VM** for fixtures → Phase 1 (parse/build TDD) and everything
COM. Phase 0b build guide: `docs/vm-setup.md`.

**Tomorrow's handoff (after the VM is up):** Chris brings (a) the COM smoke result — which pywin32
binding worked + how `GetHierarchy` returns its `[out]` value, (b) 2–3 real page IDs (mixed
styling / table / image). Then: finalize `Win32ComBackend` out-param handling, complete
`scripts/dump_fixtures.py`, dump real fixtures, start Phase 1 parse/build TDD. See README "Status".

## Grounding
- `docs/com-api-reference.md` — COM signatures + enums (from Microsoft Learn).
- `docs/onenote-xml-schema.md` — `one:` page/hierarchy XML + format-preservation rules.
- `docs/vm-setup.md` — Phase 0b Windows VM build (autologon, desktop OneNote, COM smoke, Tier-2).
- `refs/mhzarem-onenote-mcp/` — reference clone (gitignored). Payload shapes only; its
  backup-parse + PowerShell model is rejected.
