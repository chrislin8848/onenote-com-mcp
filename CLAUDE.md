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
Phases 0a, 0b, 1, 2, 3, and **4 done**. The VM is up, the XML layer is TDD'd against real dumps,
the seven read tools run through `service/read.py` on `FixtureBackend` (Linux green), AND the
live-COM read path is validated end-to-end on the VM: `tests/test_windows_read.py` (Tier 2,
`@pytest.mark.windows`) passes all 7 read round-trips via the tar-based `scripts/remote_test.sh`
loop (`onenote-tier2` `/it` task → poll → collect → mirror exit code). All three seams remain
guard-tested. Pushed to GitHub (`chrislin8848/onenote-com-mcp`, private).

Phase 2 shape: `service/read.py` does the orchestration (backend call → parse → project →
resolve IDs to names), pure-Python and testable on Linux; `server.py` tools are thin facades
(`json.dumps` for text, FastMCP `Image` content for `get_page_images` — `@mcp.tool(
structured_output=False)` since `Image` has no pydantic output schema). `get_page` is the
lossless runs+style model (effective style per run, structured tables, every content object's
`objectID`); `get_current_context` resolves the window's four `Current*Id`s via one scoped
`GetHierarchy(notebook, hsPages)`. Ground truth: a `one:Image` has NO `objectID` — the deletable
ID is on the enclosing `one:OE` (Phase 6 `delete_page_content` must target the OE).

Phase 3 details: `remote_test.sh` ships code via **tar-over-SSH** (guest has no rsync/git; uses
the guest's bsdtar; shell builtins wrapped in `cmd /c` to be cmd/PowerShell-agnostic), registers
the `/it` `onenote-tier2` task idempotently (`/f`), triggers, polls a `test-results\exit_code.txt`
sentinel, collects results + fixtures, mirrors the exit code. Guest runner is `run_tier2.bat`
(watch the `echo %RC%>file` → `0>` stdin-redirect footgun; use redirect-first). VM is `dev@
192.168.122.13` (DHCP — `virsh domifaddr --source agent win11-onenote`); default SSH shell is cmd.
**Fixture re-dump stays a deliberate MANUAL step** (dump_fixtures.py + hand-sanitize), NOT
auto-run in the Tier-2 loop — the PII lesson (a real tour roster once leaked into a dump) makes
auto-collecting raw dumps to host unsafe; run_tier2.bat runs pytest only.

**Phase 4 DONE (staged, user-confirmed per stage).** Stage 1 DONE (2026-06-11):
the surgical content mutators are built INSIDE `apply_page_edit()` — `update_page_content`
(modes append / insert_before / insert_after / replace; `target_object_id` anchors on
get_page objectIDs; content = plain text or styled runs), `create_table` (new table, or
append rows when the target is an existing `one:Table` objectID), `insert_image` (OE-wrapped,
inline `one:Data`). Mechanics locked in Stage 1: parse with `strip_cdata=False` (CDATA
survives the round-trip byte-identical); payload-strategy seam — `DEFAULT_PAYLOAD_STRATEGY
= "changed_objects"` prunes untouched page-level objects from the payload (merge can't
disturb them), `whole_page` (piBinaryData read) is the alternative — **PROVISIONAL until the
VM round-trips decide**; every `one:Image` left in a payload gets its binary inlined via
`GetBinaryPageContent` and `CallbackID` stripped (read-side construct). Tier-1 suite:
`tests/test_page_edit_content.py` (untouched OEs asserted byte-identical, real fixtures).
Stage 2 DONE (2026-06-11): `service/create.py` — `create_notebook` (OpenHierarchy cftNotebook;
empty path → `GetSpecialLocation(slDefaultNotebookFolder)`), `create_section` (`name.one`
relative to notebook OR section group), `create_page` (CreateNewPage → title + initial content
in ONE `apply_page_edit` write via the now-public composable `page_edit.content_mutator()` /
`set_title()`; `page_level`≠1 → whole-batch `apply_hierarchy_restructure` setting only the new
page's pageLevel, order unchanged). Name validation rejects `\\/:*?"<>|&#%~` before any COM
call. Tier-1: `tests/test_service_create.py` (FixtureBackend's deterministic fake IDs let
replay fixtures contain the about-to-be-created page).
Stage 3 DONE (2026-06-11): hierarchy mutators inside `apply_hierarchy_restructure()` —
`restructure_section` (complete page list in target order, optional per-page pageLevel),
`reorder_sections` (notebook OR section-group scope; mixed Section+SectionGroup list, BOTH
kinds required), `rename_node` (page/section/section group; section names filename-validated
via shared `service/names.py`), `move_page` (still EXPERIMENTAL: notebook-scope read, page
appended to target section, pageLevel resets to 1; rejects cross-notebook + recycle-bin
targets). Key behavior: recycle-bin nodes (hidden from list_*) are PINNED in place during
reorders but still ride in the submitted batch — completeness is demanded only for visible
children. Tier-1: `tests/test_hierarchy_mutations.py` on the real notebook fixture (real
recycle-bin group + section group).
Stage 4 DONE (2026-06-11) — **Phase 4 COMPLETE**: `tests/test_windows_write.py` (14 Tier-2
write tests) green on the VM alongside the 7 read tests (21 passed, 8 iterations). Validated
live: create→get→update round-trips, styled writes (dual highlight), the SPEC §6
format-preservation regression on the real mixed-format page, BOTH payload strategies
(**DECIDED: `changed_objects` stays the default**; `whole_page` = validated fallback),
concurrency guard (wrong stamp → `ConcurrencyError`; matching passes; force overrides),
table create/append-rows/cell-edit, PNG byte-identical insert round-trip, restructure,
rename, group-scope reorder, and **move_page (EXPERIMENTAL dropped — returns the page's NEW
ID)**. VM ground truths recorded in docs/com-api-reference.md ("VM-validated COM behaviors"):
VT_DATE only takes PyTime + DATE-0 unreachable (backend resolves current stamps; UTC→local
naive), real HRESULT lives in com_error excepinfo scode, COM `force` doesn't bypass the date
check (forced writes carry the current stamp), GetPageContent's lastModifiedTime is refresh-
lazy after programmatic writes, page rename = Title edit (UpdateHierarchy ignores page name),
hierarchy schema is positional (sections before groups, else hrInvalidXML — enforced in
`reorder_sections`), moved pages get NEW IDs, and `DeletePageContent` refuses paragraph OEs
(0x8004200E — paragraph delete = outline rewrite via the edit seam; **Phase 6 must scope
`delete_page_content` to page-level objects**).

**Recycle-bin policy (DECIDED 2026-06-11, user-approved):** invisible at the tool surface
(reads filter it, `include_recycle_bin` stays unexposed), pinned-in-place but fully submitted
in whole-batch UpdateHierarchy, never an operation target (move_page rejects it). Phase 5:
`copy_notebook` must SKIP recycle-bin groups when recreating section groups (+ test). Phase 6:
deletes stay default-to-recycle-bin (`permanent=False`) — the recycle bin is the undo net.

## Grounding
- `docs/SPEC.md` — the spec itself (v0611: + get_current_context, section-group integration).
- `docs/com-api-reference.md` — COM signatures + enums (from Microsoft Learn).
- `docs/onenote-xml-schema.md` — `one:` page/hierarchy XML + format-preservation rules.
- `docs/vm-setup.md` — Phase 0b Windows VM build (autologon, desktop OneNote, COM smoke, Tier-2).
- `refs/mhzarem-onenote-mcp/` — reference clone (gitignored). Payload shapes only; its
  backup-parse + PowerShell model is rejected.
