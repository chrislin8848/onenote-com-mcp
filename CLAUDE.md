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
git config core.hooksPath .githooks   # once per clone: enable the pre-push Tier-1 gate
```
The `.githooks/pre-push` hook runs the CI Tier-1 gate (ruff lint + format + pytest) before
every push so a red CI is caught locally; `git push --no-verify` bypasses it for a one-off.

## Testing tiers (SPEC §2, §6)
- **Tier 1 (host, every change):** pure logic + fixtures. Must stay green on Linux.
- **Tier 2 (VM, checkpoint):** `@pytest.mark.windows`, real COM round-trips. Auto-skipped off
  Windows. Driven by `scripts/remote_test.sh` once the VM exists.

## Status (2026-06-14)

**v1.1.2 — apply_text_style gains EMPHASIS + HIGHLIGHT (29-tool catalog unchanged; Tier-1 310 green;
Tier-2 VM-VALIDATED 59 passed).** Two real-Claude-Desktop findings: "set these 10 pages bold +
italic" and "highlight pages 1-5" — apply_text_style only had font/size/color, so Claude punted to
manual Ctrl+B/I. New params (same span-overlay seam, no new COM):
- **bold / italic / underline / strikethrough** — tri-state (true=on, false=off, omit=leave). Span
  attrs font-weight / font-style / text-decoration. underline + strikethrough SHARE text-decoration,
  so a flat overlay would clobber the other — `_apply_text_decoration()` merges per-run (keep
  existing tokens, add/remove only the requested one; empty→"none"). Whole-page also flips the
  QuickStyleDef bold/italic baseline.
- **highlight** — a color (name "yellow" or hex); writes OneNote's dual background+mso-highlight;
  `highlight="none"` removes it. **SPAN-ONLY** (NOT written to the QuickStyleDef baseline): VM
  ground truth 2026-06-14 — a QuickStyleDef `highlightColor` set to a CSS color NAME ("yellow") is
  rejected with hrInvalidXML (the whole UpdatePageContent fails); the span background already makes
  highlight visible everywhere, so the baseline is left alone (like underline/strikethrough). (font/
  size/color/bold/italic baselines still rewrite on whole-page.)
- §4: apply_text_style description + _SERVER_INSTRUCTIONS "Restyling text in bulk" para list the new
  attributes. Version 1.1.1→1.1.2. **Installer rebuilt as OneNoteMCP-Setup_1.1.2.exe.** Tier-2
  `test_windows_write.py` +2 (bold+italic+underline applied; highlight applied) — both PASS live.

**v1.1.1 — search_pages vs list_pages §4 fix (instruction-only, 29-tool catalog unchanged, Tier-1
304 green).** Driven by a real-Claude-Desktop bug: asked to restyle ●ITIN AND all its subpages,
Claude used **search_pages** (full-text FindPages) and found only ●ITIN + 2 subpages (the ones whose
content matched the query), silently missing 官網行程 / D1-0107…D6-0112 — then restyled just 3 of
~10 pages. Root cause = tool-choice, not COM: search_pages returns only pages that MATCH the query;
list_pages is the complete ordered page list with pageLevel. Fix (instruction text only, no
code/logic/COM change):
- `search_pages` description: now says it returns ONLY matching pages and silently MISSES
  non-matches — NOT a way to enumerate a section's pages or a page's subpages; use list_pages.
- `list_pages` description: the COMPLETE page list; a page's SUBPAGES are the consecutive following
  pages at a DEEPER pageLevel — use it (not search_pages) to act on "a page and all its subpages".
- `_SERVER_INSTRUCTIONS` gained a "Finding a section's pages and a page's subpages" paragraph.
- test_smoke_server: search_pages↔list_pages border + instructions asserts. Version 1.1.0→1.1.1.
  **Installer rebuilt as OneNoteMCP-Setup_1.1.1.exe.** (Instruction-only → no Tier-2 logic change.)

**v1.1.0 — `apply_text_style` (NEW tool #29) + `insert_svg_image` positioning. Tier-1 304 green;
Tier-2 VM-VALIDATED (57 passed, the 4 new write round-trips PASS live).** Two tool-surface features
(no new COM, both inside the existing apply_page_edit seam):
- **`apply_text_style`** (NEW tool, catalog 28→29): batch-patch font-family / size / color across
  EVERY text run in scope (default whole page; or an outline/table/paragraph objectID) in ONE
  read-mutate-write, preserving everything else (bold/italic/underline, untouched colors+sizes,
  highlight, hyperlinks, images, tables). Collapses "dozens of update_page_content(replace) + a
  get_page each" into one call. **Mechanism (empirically grounded, docs/PROPOSAL-apply-text-style.md):
  per-run SPAN OVERLAY is the workhorse — the span layer WINS the QuickStyleDef ← OE-style ← span
  cascade, so writing the font into each run's span makes it effective for all existing visible text
  in any scope.** Whole-page scope ALSO rewrites the page-global QuickStyleDef baseline (font/
  fontSize/fontColor) for empty paragraphs + future typing; a SUB-scope must NOT touch the shared
  QuickStyleDef (Tier-2 confirms siblings are not swept). ≥1 of font/size/color required; size in
  pt, color hex. `service/page_edit.apply_text_style`. Known limit (deferred): an empty paragraph
  whose OE @style pins an old font isn't reached by the span overlay (rare; visible text is always
  correct).
- **`insert_svg_image` gains `mode`** (append / insert_before / insert_after): the picture can now
  land MID-page (before/after a paragraph objectID), not only at an outline's end — closes a
  tool-surface gap (it had inherited the old raster insert's append-only shape). Mirrors
  update_page_content's positioning.
- **Two fidelity fixes in `build_spans` (benefit ALL edit paths):** (1) `lang` was silently dropped
  on every rebuild — real runs carry en-US/zh-TW; now preserved. (2) **Latent quote-collision bug:**
  build_spans wraps the style in SINGLE quotes but build_style_attr quoted a multiword font-family
  (e.g. "Microsoft JhengHei") ALSO in single quotes → `style='font-family:'Microsoft JhengHei';...'`
  collided, parse read `font-family:''` and DROPPED font-size/color. Fixed to DOUBLE quotes (OneNote
  itself alternates: `style='font-family:"Microsoft JhengHei"'`). Would have hit any edit rebuilding
  a spaced-font run.
- **§4 + instructions**: apply_text_style description borders update_page_content(replace) /
  modify_table(set_rows) (style-only patch vs content rewrite); `_SERVER_INSTRUCTIONS` gained a
  "Restyling text in bulk" para + an "image can be POSITIONED" note. **The OneNote App Ctrl+A
  shortcut is deliberately NOT mentioned (Chris: users already know it)** — smoke test asserts it's
  absent. Catalog 28→29.
- Tier-1: `test_spans.py` (lang round-trip, multiword-font regression), `test_page_edit_content.py`
  (apply_text_style whole-page/sub-scope/size-only/validation; insert_svg positioning), `test_smoke_
  server.py` (29 catalog, borders, mode enum, instructions). Tier-2: `test_windows_write.py` (+4:
  whole-page restyle keeps emphasis, size-only keeps font, sub-scope isolates siblings, insert_svg
  mid-page) — **all PASS live 2026-06-14 (57 passed)**. Version 1.0.10→1.1.0 (pyproject + __init__ +
  .iss + uv.lock). **Installer rebuilt as OneNoteMCP-Setup_1.1.0.exe.**

**v1.0.10 — `_SERVER_INSTRUCTIONS` behavior batch (instruction-only, 28-tool catalog unchanged,
Tier-1 293 green).** Driven by real-Claude-Desktop testing of 1.0.9 — three nudges folded into the
server instructions, NO code/tool/Tier-2 logic change (patch bump, NOT a feature; 1.1.0+ is reserved
for the apply_text_style feature):
- **(A) Don't read SVG back to verify.** After `insert_svg_image` succeeds, trust the result — do
  NOT routinely `get_page_images` to "verify" it (that re-introduces the base64-through-the-model
  cost in the READ direction — the exact slowness 1.0.9 removed on the write side); read back only
  on a reported rendering problem. Plus: get the SVG layout right in one pass (margins, no
  overlapping labels).
- **(B) In-place edit is the DEFAULT for modifying a page.** Don't rebuild a page from scratch
  (new page + re-emit content + recycle the old) just to change it — a new page gets a new ID
  (breaks links/subpage structure), loses history, and may force re-inserting images you can't
  recover. Build a NEW page only for a genuinely new/merged page. For large/risky edits the cheap
  safe pattern is copy-then-modify (copy_page/copy_section, then edit the COPY — original is the
  backup). Driven by a real trace where Claude waffled "delete outlines in place" → "build new page
  + recycle old" for a merge task; the pivot was defensible but the gap was the missing explicit
  "in-place is the default" line.
- **(C) PROPORTIONATE confirmation = irreversibility × scope, not blanket propose-then-confirm.**
  Replaced the flat confirm list. **No pre-confirm** (do + report) for reversible/lossless ops:
  in-place edits, `rename_node`, `reposition_page`, `reorder_sections`, `restructure_section`, and
  moving a SINGLE page/section to the recycle bin (report "recoverable"). **Confirm first** for
  irreversible/large-scope: `permanent=True`, `delete_page_content`/`delete_inline_content` (in-page
  content NOT recoverable), `modify_table` delete_rows/columns, `force` overwrites, deleting a whole
  section/section-group or several pages at once, notebook-wide restructures, and `move_page` (new
  ID + experimental → light confirm). **EXCEPTION:** destructive edits to a fresh COPY (just made
  via copy_* to modify it) need NO confirm — the original is the backup. Key safety fact: the
  recycle bin saves whole hierarchy nodes (pages/sections), but in-page object deletes are NOT
  recoverable.
- **Rejected:** a "decide-then-act, don't narrate internal re-planning" nudge — the final decision
  in the waffling trace was correct, so it's cosmetic / a Claude Desktop reasoning-summary display
  matter, not the server's to fight.
- Edits: `_SERVER_INSTRUCTIONS` in `server.py` (A appended to "Adding pictures"; B+C replaced the
  old destructive-ops paragraph). Version bumped 1.0.9→1.0.10 (pyproject + __init__ + .iss +
  uv.lock). `tests/test_smoke_server.py::test_server_instructions_present` +3 asserts (`in place` /
  `recycle bin` / `copy-then-modify`). **Installer rebuilt as OneNoteMCP-Setup_1.0.10.exe.**

**v1.0.9 — picture insert is now SVG-ONLY: `insert_svg_image` replaces raster `insert_image`
(28-tool catalog, Tier-1 293 green; freeze-gate VM-validated, full Tier-2 + installer PENDING).**
Driven by a real-Claude-Desktop pain: inserting an image "hung". Root cause (Chris + Claude
confirmed): the bottleneck is the MODEL emitting the image's base64 into the tool call (~39k chars
for a generated route map) — the OneNote MCP server never even received the request. base64-through-
the-model is structurally slow and unavoidable for raster bytes, so raster insert was DROPPED.
- **`insert_image` (raster, base64 param) REMOVED.** No `insert_image`/`insert_file` tool.
- **`insert_svg_image` (NEW tool, #16 slot)**: the model generates **SVG markup** (text, compact —
  no base64 token flood) and the server rasterizes it to PNG via **`resvg_py`** (new runtime dep;
  in-process PyO3 binding to the resvg Rust lib, tiny cross-platform wheel). `service/svg.py`
  `rasterize_svg()` (resvg with `sans_serif_family`/`font_family` pinned to **"Microsoft JhengHei"**
  so CJK renders cleanly even when the SVG says generic `sans-serif`; REJECTS an embedded raster
  `data:image/…` URI — vector-only, no smuggling a photo back through base64). Then the EXISTING
  seam: `page_edit.insert_svg_image` = rasterize → `make_image` (kept) → OE → `apply_page_edit`
  (the one write core). Photos/existing raster images = insert BY HAND in OneNote (instructed).
- **Copy path UNAFFECTED**: `transfer_page` / `inline_image_binaries` / `make_image` untouched —
  copying a page/section still carries existing images & attachments faithfully.
- **§4 + instructions**: `_SERVER_INSTRUCTIONS` "Adding pictures" para = SVG-only + use explicit
  CJK font + photos by hand; `update_page_content` border points adds-a-picture → `insert_svg_image`;
  facade description spells out SVG-only / CJK font / data-URI rejection. Catalog stays 28 (raster
  insert out, svg insert in).
- **Freeze gate VM-VALIDATED (2026-06-14)**: a throwaway spike (`packaging/spike_svg/`,
  `scripts/remote_spike_svg.sh`) PyInstaller-froze `resvg_py` and the frozen Windows exe rendered a
  Traditional-Chinese SVG to a real-glyph PNG (exit 0). Finding: generic `sans-serif` mapped to a
  handwriting font on Windows → fixed by pinning the default family (above). `onenote-mcp.spec` now
  `collect_all("resvg_py")`.
- Tier-1: `tests/test_service_svg.py` (rasterize valid SVG; reject empty; reject raster data-URI;
  CJK text not mis-rejected), `test_page_edit_content.py` (+SVG→OE-wrapped-PNG, +raster-reject-
  before-write), `test_write_core.py` (svg facade routes through apply_page_edit), smoke catalog
  swap. Tier-2 `test_windows_write.py` (+SVG round-trip, +raster-reject) — **runs on VM but NOT yet
  executed**. **PENDING: full freeze (onenote-mcp.spec) + Tier-2 live run + CJK eyeball with an
  explicit font + installer rebuild as OneNoteMCP-Setup_1.0.9.exe.** Version 1.0.9 (pyproject +
  __init__ + .iss + uv.lock).

**v1.0.8 — copy tools report by NAME + harder no-raw-ID instruction (28-tool catalog, Tier-1 286
green).** Driven by real-Claude-Desktop UX friction: the model kept dumping raw page-ID lists to
the user (e.g. narrating "I'll scan these 10 pages:" + 10 opaque IDs), because `copy_pages` /
`copy_page_subtree` returned BARE `page_ids` — the model had only IDs in hand to refer to the new
pages. Fix (no new tool, no new COM): the copy results now carry NAMES so the model narrates by
name.
- `transfer_page` reads the source page's title from the page-content XML root's `name` attribute
  (already parsed — NO extra COM read) and returns it + `page_level` on `PageCopyResult`.
- `copy_pages` aggregates a `pages: [{page_id, name, page_level}]` list (parallel to `page_ids`,
  which stays for placement); server `copy_pages` / `copy_page_subtree` now return `pages` (not a
  bare `page_ids` list), and `copy_page` returns `name` + `page_level` alongside `page_id`.
- `_SERVER_INSTRUCTIONS` hardened: "NEVER paste a raw ID — especially a LIST of IDs — into your
  reply; narrate by NAME; the copy tools return names alongside ids for exactly this." Copy tool
  descriptions updated to say "report by NAME, never as a raw id list".
- This is a NUDGE, not a guarantee (whether to print IDs is ultimately Claude Desktop's model
  behavior; the most effective lever is a user-level Claude Desktop instruction). Tier-1 286 green
  (test_copy_pages engine asserts the `pages` list; facade tests assert `out["pages"]`;
  test_copy_page_placement transfer-stub gains name/level). Installer rebuilt as
  OneNoteMCP-Setup_1.0.8.exe.
- Follow-up (SAME 1.0.8 version, SEPARATE commit since 1.0.8 was already pushed): `_SERVER_INSTRUCTIONS`
  gained a "keep narration BRIEF — don't walk the user through tool names / param shapes (set_rows
  arrays, 0-indexed columns) / objectIDs / step-by-step mechanics; say it in human terms" paragraph;
  and the installer now shows a **Traditional-Chinese post-install page** (`packaging/post_install_zh-TW.txt`,
  UTF-8 BOM, wired via `InfoAfterFile`) reminding the user to paste the recommended Claude Desktop
  custom instruction. RATIONALE: Claude Desktop user custom instructions are ACCOUNT-level cloud
  settings (not a local file) — an installer CANNOT auto-apply them (`claude_desktop_config.json`
  holds only `mcpServers`). So the strong lever (account custom instruction) stays a one-time MANUAL
  user step; the server `instructions` nudge + the install-time reminder are the most the
  server/installer can do. (InfoAfterFile shows on a normal install; `/VERYSILENT` skips it.) The
  rebuilt 1.0.8 installer therefore has a NEW sha (supersedes the earlier 1.0.8 artifact).

**v1.0.7 — page-level object visibility + lightweight object inventory (28-tool catalog). Tier-1
286 green; fix VM-VALIDATED live (2026-06-14).** Driven by a real-Claude-Desktop failure: asked to
delete the printout images on `行程Final` / `Final 郵輪行程`, Claude "saw" no images and skipped them
(and separately mis-claimed "GetPageContent failed / not synced" — that was transient, NOT
reproducible: those pages read fine, 77k–185k chars). Root cause, VM-confirmed by a diagnostic on
the real pages: the printout render images are **page-level** objects (direct `one:Page` children,
OUTSIDE any outline), and `get_page` emitted ONLY `outlines` — so `page.page_images` / `page.page_files`
were INVISIBLE in the page read. The model couldn't delete what it couldn't see (not laziness; a real
fidelity gap). Fixed inside the read layer (no new COM, no new invariant):
- **`get_page` now surfaces page-level objects**: new `page_level_images` / `page_level_files` keys
  (printout renders + page-level InsertedFile variant). Previously omitted; printout images are now
  visible and route to `delete_page_content`. `service/read.py`.
- **`get_page_info`** (NEW tool #28): a cheap, FLAT, EXHAUSTIVE object inventory — every object's
  `object_id` + `type` + `delete_with` (delete_page_content for page-level, delete_inline_content
  for in-outline) + `page_level` + light metadata (image w/h/OCR-flag, table rows×cols, file
  name/kind, a ~40-char paragraph preview). NO full runs / style table / pixels, NO disk I/O. Walks
  outlines recursively (into children AND table cells) then appends page-level objects, so images
  buried after a table, inside a cell, or page-level printout renders are all listed. This is the
  "lightweight half of the two-step objectID rule" and the right tool for "find/delete ALL images".
  `read.get_page_info` + `_page_object_inventory`. (Considered but DROPPED `find_images`: a fixed
  `get_page` + `get_page_info` cover the verified within-page bug; the cross-page "assumed later
  pages match earlier" miss is unproven-recurring model laziness, defer a multi-page tool until it
  recurs — YAGNI.)
- **§4 borders**: `get_page` (full CONTENT) ↔ `get_page_info` (lightweight INVENTORY) cross-named;
  `_SERVER_INSTRUCTIONS` two-step rule now leads with `get_page_info` and says "to find/delete all
  images use get_page_info, do NOT eyeball get_page's nested tree, and check EACH page's inventory".
  Also retitled `get_page_files_info` as the file-EXTRACTION pre-check (size_bytes + media_class,
  the get_page_files prerequisite) and pointed plain file DISCOVERY at `get_page_info` (cheaper, no
  disk read, all object types) — the two were superficially overlapping; `get_page_files_info` kept
  (its size/media_class need disk I/O that must stay OUT of the lightweight inventory).
- Tier-1: `tests/test_service_read.py` (+4: get_page page-level images/files on the real printout
  fixture `附件與嵌入物件-1`; get_page_info lists the page-level image with delete_page_content; flat
  /exhaustive/lightweight; inline-image page routes to delete_inline_content); `test_smoke_server.py`
  27→28 + get_page↔get_page_info + get_page_files_info→get_page_info borders. **VM live-validated
  (2026-06-14): `行程Final` now shows 5 page-level images, `Final 郵輪行程` 2 — all delete_page_content;
  pre-fix both showed 0.** Tier-2 windows test not yet added (read-only change; live-validated by the
  diagnostic instead).

**v1.0.6 — multi-page block copy (27-tool catalog). Tier-1 282 green; Tier-2 VM-VALIDATED
(2026-06-14: both new subtree round-trips PASS live — `copy_page_subtree`'s subtree detection
matches the live positional model, and a cloned subtree lands as ONE contiguous block, in source
order, with each page's pageLevel preserved).** Driven by a real-Claude-Desktop failure: "copy
●ITIN and its subpages below the ●Local page" came out SCATTERED — the model had no single-call way
to copy a page subtree, so it strung together `copy_page` calls, and `copy_page`'s default ("place
the copy right below its SOURCE page") left each copy glued beside its own original instead of as a
block below ●Local. Root cause = a missing first-class operation (we had only single-page `copy_page`
and whole-section `copy_section`); fixed with two new tools + one internal block-placement seam, all
inside the existing copy / hierarchy seams (no new COM, no new invariant):
- **`copy_pages`** (NEW tool #26): clone SEVERAL pages, in given order, as ONE contiguous block
  (each page keeps its pageLevel); `after_page_id` places the block after that page, else section
  end. `service/copy.py` (engine = `copy_pages`).
- **`copy_page_subtree`** (NEW tool #27): clone a page TOGETHER WITH its subpages (the consecutive
  following pages at a DEEPER pageLevel — OneNote's positional subpage model, computed by
  `copy.subtree_page_ids`) as a block. Same-section DEFAULT places the block right below the source
  subtree (like `copy_page`); `after_page_id` overrides; cross-section → section end. This is the
  one-call fix for "copy ●ITIN and its subpages below ●Local" (page_id=●ITIN, after_page_id=●Local).
- **`reposition_pages`** (NEW internal seam in `service/hierarchy_edit.py`, NOT an MCP tool): the
  multi-page sibling of `reposition_page` — gathers a list of pages into a contiguous block after an
  anchor (or section top) in ONE `UpdateHierarchy`, node-IDs conserved (SPEC §5 whole-batch). Both
  copy facades chain it for placement; copy_pages re-raises an explicit missing anchor, copy_page_
  subtree swallows an IMPLICIT (cross-section) one — same pattern as copy_page.
- §4: new "copy granularity four-way" contrastive border (copy_page single / copy_pages explicit
  list / copy_page_subtree page+subpages / copy_section whole section), with the explicit "do NOT
  loop copy_page — it scatters" rule in both the tool descriptions and `_SERVER_INSTRUCTIONS`. SPEC
  §4 tool table + borders updated. Tier-1: `tests/test_copy_pages.py` (+21: reposition_pages block
  placement, subtree detection incl. deep-level rule, copy_pages engine, both facades' orchestration);
  Tier-2: `tests/test_windows_copy.py` (+2 live round-trips, both PASS 2026-06-14).
- **Installer BUILT (2026-06-14): `dist/installer/OneNoteMCP-Setup_1.0.6.exe`** (~22.9 MB,
  sha256 70ea3a0e…5912e0) — freeze + iscc on the VM (`packaging/build.bat`), and the frozen exe's
  COM `--selftest` passed in the interactive session (bound OneNote, 3 notebooks, `bound via:
  vendored`). **Still pending: Chris's real-Claude-Desktop §4 acceptance** — install on the PC
  running Claude Desktop and confirm the natural-language "copy ●ITIN and its subpages below ●Local"
  actually routes to `copy_page_subtree` and lands as a contiguous block (the Tier-2 pass validated
  the COM mechanics; the model's tool CHOICE is the §4 acceptance, which only real Claude Desktop tests).

**v1.0.5 — page positioning + table-clear ergonomics (25-tool catalog). Tier-1 263 green; Tier-2
VM-VALIDATED (reposition_page / set_rows-None / copy_page-below-source all PASS live).** Driven by
a real-Claude-Desktop failure: "copy this page below the original, then clear the tables" got STUCK
because the only way to position a page was restructure_section's whole 46-page list — the model
reached for an external scratch file and hit Claude Desktop's read-only sandbox. Fixes, all inside
the existing hierarchy/edit seams:
- **`reposition_page`** (NEW tool #25): move ONE page to a new spot within its section — right
  after `after_page_id` (empty → section top), optional `page_level`. Caller gives only IDs; the
  seam reads the full page list and conserves node-IDs, so the SPEC §5 whole-batch discipline holds
  without the LLM enumerating 46 pages (same pattern as move_page). `service/hierarchy_edit.py`.
- **copy_page + create_page default placement**: copy_page gained `after_page_id` and now DEFAULTS
  to placing the copy right BELOW the source page; create_page gained `after_page_id` and DEFAULTS
  to placing the new page below the page the user is currently on (`get_current_window_ids`, read
  BEFORE creating). Both fall back to the section end when the anchor isn't in the target section
  (cross-section copy / current page elsewhere / no window) — an EXPLICIT missing anchor still
  raises. Composition lives in the server facades; orchestration is Tier-1 facade-unit-tested
  (`tests/test_copy_page_placement.py`, patched service calls) + the create_page facade is exercised
  live (explicit-anchor, deterministic) in Tier-2.
- **`set_rows` None=keep**: a `None` cell leaves THAT cell unchanged (`[None, "", …]` = keep column
  0, clear the body — the "keep first row/column" ask), so clearing a table body no longer needs to
  re-supply the kept column's values.
- §4: the reorder set (reposition_page / restructure_section / reorder_sections / move_page /
  rename_node) cross-named; copy_page/create_page document the default-below-anchor; server
  instructions say "single page → reposition_page, never an external scratch file". SPEC §4/§5 +
  tool table updated.

**v1.0.4 — inline delete + bulk table-content replace (24-tool catalog). Tier-1 246 green;
Tier-2 VM-VALIDATED (2026-06-13): both new live round-trips PASS** — `set_rows` overwrites a row
and the whole table in place (cell objectIDs preserved), and `delete_inline_content` drops an
inline table while the surrounding paragraphs survive. (The run also surfaced ONE pre-existing,
unrelated failure: `test_copy_section_…` did `next(... "混合樣式頁")` without a skip-guard and hit
StopIteration because that dump-source page is no longer in the VM's "Phase 0 測試用" section — its
read/write siblings already skip on that. Fixed the test to skip the content spot-check when the
sample page is absent; the de-collide/order/level assertions still run.) Closes two real gaps found
in hands-on use, both built INSIDE the edit seam (`apply_page_edit`) — no new COM, no new invariant:
- **`delete_inline_content`** (NEW tool, #24): delete ONE object from INSIDE an outline — a table,
  a paragraph, or an inline image/attachment — by objectID, via the edit seam (NOT
  DeletePageContent, which COM refuses for inline OEs). The exact gap reported: `delete_page_content`
  is page-level-only, so it could not drop a table while keeping the surrounding paragraphs. Reuses
  `remove_content_element` (prunes the emptied OE/OEChildren/Outline; a table cell is replenished,
  never emptied). Sibling paragraphs survive. A page-level objectID is refused here with a pointer
  to `delete_page_content` (the symmetric complement). Lives in `service/page_edit.py` (it composes
  a Mutator, so it must); `delete.py`'s inline-refusal guidance now points at it. §4 delete trio is
  now node / page-level / inline, cross-named both ways (guard-tested in test_smoke_server.py).
- **`modify_table` `set_rows`** (NEW operation): OVERWRITE the content of existing rows with a 2-D
  array from `at_index` (one row = "replace this row", every row = "refresh the whole table"),
  fixed-shape — no row/column added/removed, every cell keeps its objectID, a short input row leaves
  trailing columns untouched. Only rewrites EXISTING rows (add rows = `insert_rows`); out-of-range /
  too-wide is refused. Reuses `_cell_runs` + `_replace_oe_text`. Distinct from update_page_content
  "replace" (ONE cell) and create_table (NEW table) — borders added to all three descriptions.
- Tier-1: `test_page_edit_content.py` (+set_rows ×5, +delete_inline_content ×4 incl. keep-siblings,
  cell-stays-valid, page-level-refused); `test_smoke_server.py` 23→24 + both-way border assert;
  Tier-2 `test_windows_write.py` (+set_rows + delete-inline-keeps-paragraphs round-trips — both
  PASS live 2026-06-13).
- Text-paragraph edits confirmed WHOLE-OE replacement (a fresh one:T swapped in), never
  character-level — `_replace_oe_text`, untouched paragraphs stay byte-identical.

**v1.0.3 — table-shape editing + hyperlinks (23-tool catalog).** Two text/table-editing
capabilities added inside the existing seams + VM-validated:
- **`modify_table`** (NEW tool, #23): change an EXISTING table's shape in place — `insert_rows`
  (at a position, or append when at_index omitted), `add_columns` (empty columns at a position,
  every row kept rectangular), `delete_rows`, `delete_columns` (both DESTRUCTIVE). Row/column ops
  are symmetric. `create_table` is now create-ONLY (its old append-rows branch moved here). Nested
  tables (table-in-cell) are addressed by the inner table's own objectID — read + modify both
  recurse/scope correctly (Tier-1 `test_page_edit_content.py`); creating a NEW nested table is
  still unsupported. Built in `service/page_edit.py`, `remove_content_element` keeps a cell valid
  (never an empty `<one:Cell/>`).
- **Hyperlinks**: text runs carry an optional `link` (href). `xmllayer/spans.py` parses/builds
  `<a href="…">…</a>` inside one:T CDATA (VM ground truth: the 表格頁 fixture). `get_page` now
  reports each run's `link` — FIXING a real fidelity gap (a link's URL was dropped on read, so a
  replace edit silently stripped it). `update_page_content` content may set `link`.
- §4 contrastive borders re-balanced across create_table / modify_table / update_page_content /
  delete_page_content so a naive LLM disambiguates; modify_table's deletes added to the
  propose-then-confirm list in `_SERVER_INSTRUCTIONS`. Tier-1 237 + Tier-2 46 green.

**Post-1.0 (IN PROGRESS) — see [docs/HANDOFF-copy-sync.md](docs/HANDOFF-copy-sync.md):** the
long-standing "OCR images can't be copied" issue was RE-DIAGNOSED as OneDrive files-on-demand
UNDER-SYNC (NOT OCR, NOT a COM limitation — a fully-synced machine copies fine). `copy_page`/
`copy_section` now DETECT + REPORT un-synced images / files / embedded objects via a
`sync_warning` (Tier-1 green, `tests/test_copy_sync.py`). Tier-2 validation + doc hygiene (the
"Real-data resilience" paragraph below still says OCR images are "not retrievable via COM" — now
known to be a sync artifact, not OCR) are open TODOs in that handoff file.

Phases 0a, 0b, 1, 2, 3, 4, 5, **5b done**, and **Phase 6 COMPLETE through Stage 5** (deletes +
diagnostic log + §4 descriptions + Tier-2 deletes; PyInstaller freeze AND the Inno Setup
installer both VM-validated end-to-end — silent install → `--configure` registers Claude
Desktop → installed exe `--selftest` binds OneNote, exit 0; **version 1.0.0**). Tier-1
214 + Tier-2 47 green. Remaining: only Chris's real-Claude-Desktop §4 acceptance pass (by hand,
NOT CC sub-agents) — packaging is done.

**Real-data resilience (2026-06-13, VM-validated on Chris's actual 測試章節1/測試章節2 sections,
`tests/test_windows_realdata.py`):** when a section is NOT fully downloaded on this machine
(OneDrive files-on-demand), `GetBinaryPageContent` returns `0x8004200F` (`hrBinaryObjectDoesNotExist`)
for the not-yet-hydrated image binaries — a **sync artifact, NOT an OCR or COM limitation** (a
fully-synced machine serves them fine; the earlier "OCR-processed images un-extractable, full stop"
claim here was WRONG — re-diagnosed 2026-06-13, see [docs/HANDOFF-copy-sync.md](docs/HANDOFF-copy-sync.md)).
Both content paths handle the under-synced case: the **copy path** REMOVES the un-fetchable image /
attachment and prunes the emptied OE/Outline, then returns a categorized `sync_warning` (N images +
X files + Y objects OMITTED — sync the source + re-copy for fidelity; `service/copy.py`,
`tests/test_copy_sync.py`, Tier-2 validated: no hrInvalidXML, 148/156 images removed on an un-synced
section); the **read path** (`read.get_page_images`) SKIPS the un-fetchable image (`get_page` stays
the authority on what images a page has — object_id/dimensions/OCR text — so nothing is hidden, only
un-downloaded pixels are dropped). Real-data smoke = read all 4 tools over every page + copy section
to a throwaway + delete-and-confirm-gone, all green on both real sections.
The VM is up, the XML layer is TDD'd against real dumps,
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
Stage 2 DONE (2026-06-11): `service/create.py` — `create_notebook` (REMOVED post-Phase-5:
COM can't create notebooks on this build), `create_section` (`name.one`
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
in whole-batch UpdateHierarchy, never an operation target (move_page rejects it). (The Phase-5
copy_notebook recycle-skip clause became moot — copy_notebook was removed; copies of sections
never touch the bin.) Phase 6: deletes stay default-to-recycle-bin (`permanent=False`) — the
recycle bin is the undo net.

**Phase 5 DONE (staged, user-confirmed per stage).** Stage 1 DONE (2026-06-11): `transfer_page()` raw-XML clone
inside the copy seam — piBinaryData read, pixels inlined via the now-public
`page_edit.inline_image_binaries()` (binary dumps serve only CallbackID), identity/stamps/view
state stripped (QuickStyleDef/TagDef/spans/author attrs ride verbatim, `strip_cdata=False`),
transplanted onto a `npsBlankPageNoTitle` page in ONE UpdatePageContent guarded by the BLANK
page's stamp, source pageLevel preserved via the whole-batch hierarchy seam. Tier-1:
`tests/test_copy_page.py` (byte-level fidelity on real binary fixtures).
Stage 2 DONE (2026-06-11): `transfer_section` (pages in document order, per-page pageLevel;
target = notebook OR group) + `transfer_notebook` (groups recreated via
`OpenHierarchy(cftFolder)`, sections land INSIDE groups, **recycle-bin groups skipped** per
policy). Key trap handled: `OpenHierarchy` OPENS an existing same-named node instead of
creating — copies de-collide names with " (2)", " (3)", … against the target's direct children
(both kinds). Tier-1: `tests/test_copy_tree.py`.
Stage 3 DONE (2026-06-11) — Tier-2 全綠 (27 passed): B≡A semantic fingerprints + byte-identical
image pixels on the three dump-source pages; section copy preserves order + 1/2/3 levels and
de-collides live; section copy lands inside groups. **`copy_notebook` AND `create_notebook`
REMOVED (user-approved): `OpenHierarchy(cftNotebook)` refuses BOTH local paths and OneDrive
https parents with hrFileDoesNotExist on this M365 build — COM cannot create notebooks.
Notebooks are made in the OneNote UI; whole-notebook cloning = copy_section per section.** Further
ground truths (docs/com-api-reference.md Phase-5 section): OpenHierarchy OPENS same-named
nodes (→ copies de-collide with " (n)"); DeletePageContent DOES work on one:Outline
(page-level) — so Phase 6 delete_page_content = page-level objects only; empty-paragraph-only
outlines render as nothing and are dropped on transplant (B≡A ignores them). VM-ops lesson:
NEVER probe COM writes over plain SSH — a hung OpenHierarchy blocked OneNote's single-threaded
COM for everything; writes only via the interactive Tier-2 task.

## Phase 5b DONE (2026-06-12 — SPEC v0612-2 §5): attachments / embedded objects (`one:InsertedFile`)
New feature slice, deliberately BEFORE Phase 6 (the §4 description pass needs the final 22-tool
catalog; adding tools after that pass would force a redo).

**Stage 1 DONE (2026-06-12, Tier-1 173 green):** the read slice. `xmllayer` parses
InsertedFile (both placements; kind = attachment_icon/printout/embedded_preview per the
children rule) and now collects PAGE-LEVEL objects: `Page.page_files` + `Page.page_images` —
fixing a real gap: printout render images (direct page children) were invisible to
`get_page_images`, and `_oe_object_id` returned None for them (page-level objects carry their
OWN objectID; inline ones use the enclosing OE's). Backend grew `stat_cache_file`/
`read_cache_file` (pathCache disk reads; None = cache unavailable, never raise; Win32 = plain
file IO, FixtureBackend replays `cachefile_<sanitized {GUID}.bin>.bin` — those fixtures are
HAND-AUTHORED neutral bytes from `scripts/make_cache_fixtures.py`, not dumps; docx cache
deliberately absent = unavailable replay). `service/files.py`: `get_page_files_info` (metadata
any type, deletable object_id, media_class) + `get_page_files` (text decode utf-8-sig→cp950→
lossy; image base64; PDF text via **pypdf** [new runtime dep]; office types explicit
"unsupported"; 20MB cap + max_chars truncate; per-file status, never crash). `get_page` shows
inline attachments as `type:"file"` blocks; `_CONTENT_TAGS` += InsertedFile/XPSFile (unchanged
ones pruned from edit payloads — guard-tested in test_page_edit_content). Server: 22-tool
catalog complete (`get_page_files` emits mixed JSON + MCP image content). Tier-2 validation of
all of this = Stage 3.

**Stage 2 DONE (2026-06-12, Tier-1 177 green):** copy-path attachment fidelity, inside the
copy seam. New backend method `stage_cache_copy(path, preferred_name)` (copy the cache file
to `%TEMP%\OneNoteMCP\staging\<uuid>\<name>`, NOT auto-deleted — re-import timing unknown;
FixtureBackend returns deterministic `C:\FixtureStaging\<n>\<name>` on its OWN counter so
node-ID replay fixtures don't shift). `transfer_page` now: pops `pathCache` from EVERY
InsertedFile (never rides — dead reference), re-points `pathSource` at the staged copy
(embedded objects GAIN a pathSource; Previews ride verbatim), and FLATTENS printouts —
page-level `one:XPSFile` carriers (read-side CallbackID constructs) are stripped, the
`one:Printout` child dropped (its xpsFileIndex would dangle), printout bookkeeping attrs
(`xpsFileIndex`/`isPrintOut`/`originalPageNumber`) removed from the render images, which
survive as plain inlined images. Returns `PageCopyResult(page_id, file_notes)` /
`SectionCopyResult` — every fidelity loss (unavailable cache, flattening) reported
explicitly, never silent (SPEC §5); copy_page/copy_section emit `file_notes`. **NEW GROUND
TRUTH: pathCache GUIDs are PER-READ ephemera** (basic vs binary dump of the same page carry
different Temp GUIDs) — never persist one across reads; cache fixtures exist under both
dumps' GUIDs. Re-import mechanics + embedded clone behavior remain VM-gated → Stage 3.

**Stage 3 DONE (2026-06-12) — Tier-2 全綠 (36 passed + 3 probes, `tests/test_windows_files.py`):**
both read tools round-trip live (kinds/sizes/caches; txt+PDF+jpg content; printout render
visible via get_page_images); the rewritten copy payload is ACCEPTED live, and the VM-gated
questions are now VM-ANSWERED (docs/com-api-reference.md "Phase 5b" section): **staged-
pathSource re-import WORKS** (all 5 caches rebuilt on the copy within 45s), **embedded objects
clone as embedded** (Previews honored, kind stays embedded_preview), **printout render PNGs
are re-encoded on transplant** (dimensions identical, bytes not — semantic fidelity only,
unlike normal images), `changed_objects` edits on a printout page merge cleanly. **Phase-6
delete probes: page-level InsertedFile = DeletePageContent ACCEPTED; inline attachment-bearing
OE = REFUSED** (same class as paragraph OEs) ⇒ inline-attachment delete must be an outline
rewrite through the edit seam. Probes report via pytest.skip in -ra (never gate the run);
Tier-2 assertions stay content-light (no live file text in tier2.log — PII policy).

Original scope (all delivered):
- **Two new read tools** — `get_page_files_info` (metadata for ANY InsertedFile: preferredName,
  type, size from the cache file, `objectID`, `kind` attachment-icon/embedded-preview/printout
  via `Previews`/`Printout` children — discrimination法待 VM dump; never parses content) and
  `get_page_files` (content extraction LIMITED to: text-class decode, image → MCP image content,
  PDF → server-side text via a pure-Python lib [default candidate: pypdf — PyInstaller-safe];
  docx/xlsx/pptx etc. → metadata + explicit "unsupported"; size cap + truncate).
- **Binary path differs from images:** content = read the `pathCache` file on disk, NOT
  `GetBinaryPageContent`. Cache may be missing (unsynced/purged) → graceful "cache unavailable",
  never crash. File reads live BEHIND the backend (new backend method; FixtureBackend replays).
- **Copy fidelity extension (transfer_page):** old `pathCache` = dead reference; clone = copy
  the cache file aside, rewrite the element to `pathSource` → the copy, drop `pathCache`, let
  OneNote re-import. Re-import mechanics + embedded-spreadsheet clone behavior are VM-gated.
  Source cache unavailable ⇒ report that attachment explicitly in the result, never skip silently.
- **Ground truth IS IN (dumps 2026-06-12, docs/onenote-xml-schema.md InsertedFile section;
  fixtures committed):** InsertedFile has TWO placements — inline in `Outline>OEChildren>OE`
  (no own objectID; enclosing OE has it) AND page-level (direct Page child, Position+Size,
  OWN objectID). Kind rule: no children=icon / Printout=printout / Previews=embedded (embedded
  has NO pathSource). pathCache verified live (`Temp\{GUID}.bin`); printout = InsertedFile +
  page-level one:XPSFile (callback → ORIGINAL source bytes) + page-level one:Image renders;
  deleting the InsertedFile leaves XPSFile/Image orphans, deleting the render GC's the XPSFile.
  ⇒ page-level variant must join `_CONTENT_TAGS`; inline delete = open VM question
  (DeletePageContent refuses paragraph OEs — outline rewrite may be needed).
- **No `insert_file` tool** (deliberate, in the limits list). `Printout` renders as page images
  → same burned-in-pixels limits as images.
- Fixtures: dumped via `scripts/remote_dump.sh` (onenote-dump /it task → test-results/dump,
  never straight into tests/fixtures; inspect for PII before moving — the first dump caught
  real flight tickets and was rejected; page 1 is being rebuilt with neutral files).

## Phase 6 (IN PROGRESS — SPEC v0612) — Stages 1–4 DONE, Stage 5 (packaging) NEXT
**Stage 1 DONE (commit 8890557, Tier-1 187):** `service/delete.py`. `delete_node` = thin
DeleteHierarchy facade, recycle-bin default (`permanent=False`). `delete_page_content` validates
the target is a PAGE-LEVEL object (direct `one:Page` child: Outline/Image/InsertedFile/
InkDrawing/MediaFile) BEFORE the COM call, turning the inline-OE refusal into a clear error
naming what the target is + pointing at update_page_content. Carries the stamp from the
validation read.
**Stage 2 DONE (commit 9434e05, Tier-1 196):** `logging_config.py` (§7). `ONENOTE_MCP_LOG_LEVEL`
(default ERROR=detailed off; DEBUG=on) + `ONENOTE_MCP_LOG_FILE` (default `%LOCALAPPDATA%\OneNoteMCP
\logs\`, omitted→stderr-only). Rotating file + stderr, NEVER stdout. `log_tool_call` composed into
`logged_tool` on all 22 tools: DEBUG logs redacted params+result, ERROR logs failures+hresult even
when off. Redaction bounds every field (no base64/note content); fully guarded (never throws,
bad path→stderr). `configure_logging()` in main().
**Stage 3 DONE (commit 55b7019):** §4 cross-set description pass — contrastive borders on all
confusable sets, DESTRUCTIVE+propose-confirm contracts in text, `mode` as a Literal enum,
server-level `_SERVER_INSTRUCTIONS` (live-OneNote caution, two-step objectID rule, propose-confirm,
date-rewrite benchmark). **restructure/reorder verb-plural mismatch: KEPT (user decision
2026-06-12 — the singular/plural already signals one-section-of-pages vs many-sections; the
borders handle the confusion).** §4 acceptance on real Claude Desktop is still Chris's to run.
**Stage 4 DONE (host smoke commit 6c64da7 + Tier-2 commit pending; Tier-2 43 passed):**
`tests/test_smoke_server.py` (catalog completeness, host) + `tests/test_windows_delete.py` (7
Tier-2 delete round-trips, ALL GREEN on the VM): delete_node section/page/permanent (recycled →
gone from list_*), delete_page_content on a page-level outline/attachment/image actually removes
it, and an inline paragraph OE is refused with ValueError before COM. **VM-confirmed delete
semantics** (already in docs/com-api-reference.md Phase-5b section): page-level Outline/Image/
InsertedFile ACCEPTED; inline OEs REFUSED.
**Stage 5 DONE — freeze + installer both VM-VALIDATED end-to-end (2026-06-12):**
- **Host authoring DONE (commits 0495522 + e052acf):** `configure.py` (`--configure`: detect
  regular + Store Claude config by glob, merge-not-clobber, no log-level var) + `--selftest`
  (bind COM, list notebooks — the install health check) + `packaging/` (PyInstaller spec [onedir,
  console, makepy bundled], `rthook_win32com_gen_py.py` [writable gen_py], `onenote-mcp.iss`
  [Inno, per-user, runs --configure post-install], `build.bat`, README) + `scripts/remote_build.sh`
  + `run_selftest.bat`. New `packaging` dep group (pyinstaller). Tier-1 212.
- **Freeze + COM smoke VALIDATED ON THE VM (2026-06-12):** `OneNoteMCP.exe --selftest` (frozen
  onedir, interactive session) connected to OneNote and listed 3 notebooks (exit 0). The
  gen_py-regeneration recipe works (see docs/com-api-reference.md "Phase 6 Stage 5"): EnsureModule
  + bundled makepy + writable `win32com.__gen_path__`. **SPEC §8's late-bound-Dispatch note is
  moot** (late binding never worked for OneNote; this is the right fix, now proven).
- **Installer BUILT + END-TO-END VALIDATED ON THE VM (2026-06-12):** Inno Setup 6.7.3 installed
  via `winget install --id JRSoftware.InnoSetup --source winget` (per-user, NOT on PATH, at
  `%LocalAppData%\Programs\Inno Setup 6\ISCC.exe`; avoid the msstore source — its agreement
  prompt hangs non-interactive SSH). `packaging/build.bat` now resolves iscc from PATH→that
  location; the `.iss` anchors Source/OutputDir with `{#SourcePath}..\` (freeze writes repo-root
  `dist\`, but iscc resolves relatives against the script's `packaging\` dir). Output filename
  carries the version: `dist/installer/OneNoteMCP-Setup_<AppVersion>.exe` (`OneNoteMCP-Setup_
  1.0.0.exe` as of the 2026-06-13 real-data-resilience rebuild, ~22.8 MB). Validated full chain:
  silent install (`/VERYSILENT`, per-user, no admin) → exe lands at
  `%LocalAppData%\Programs\OneNoteMCP\` → post-install `--configure` registers the `onenote`
  entry in `claude_desktop_config.json` pointing at the INSTALLED exe → the installed exe's
  `--selftest` bound OneNote, exit 0, via the interactive task. Version is 1.0.0 (pyproject +
  __init__ + .iss + uv.lock).
- **Still pending (Chris, NOT a packaging blocker):** the §4 real-Claude-Desktop acceptance pass
  — install the .exe (or the source server) and exercise the 22-tool catalog by hand in real
  Claude Desktop (SPEC §4 — NOT CC sub-agents).

Original Phase-6 scope as planned (all but packaging delivered):
Deletes (`delete_node` hierarchy incl. section group + `delete_page_content` PAGE-LEVEL objects
only — VM-verified semantics: outline/page-level Image/page-level InsertedFile are ACCEPTED;
paragraph OEs AND inline attachment-bearing OEs are REFUSED, so inline content removal = outline
rewrite through the edit seam; recycle-bin default), error/retry hardening, smoke tests, then
packaging (PyInstaller → Inno/NSIS `OneNoteMCP-Setup.exe`). Two new SPEC v0612 deliverables:
- **Diagnostic log (§7), default OFF.** Env switch `ONENOTE_MCP_LOG_LEVEL` (default `ERROR`;
  `DEBUG` on) + `ONENOTE_MCP_LOG_FILE`; rotating file (`%LOCALAPPDATA%\OneNoteMCP\logs\`) and/or
  stderr, **never stdout** (JSON-RPC). One record per tool call: name, params (large/base64
  fields truncated), result + COM error code. Logging must never throw/interrupt; write-fail →
  silent-degrade to stderr. `--configure` does NOT write the var. Mostly host-doable + testable.
- **Tool-description enhancement (§4), one cross-set pass.** Contrastive/negative borders on
  confusable pairs (delete_node vs delete_page_content; update_page_content vs create_table vs
  insert_image; restructure_section vs reorder_sections vs move_page vs rename_node; get_page vs
  get_page_images vs get_page_files_info [metadata, any type] vs get_page_files [content, only
  text/image/PDF — info is the prerequisite of files]), DESTRUCTIVE + propose-confirm contracts
  in description text, append/insert/replace as a per-value enum, server-level `instructions`,
  read tools return `objectID`s (incl. get_page_files_info → delete_page_content).
  **Acceptance is real Claude Desktop (Chris), NOT CC sub-agents** (SPEC §4 — sub-agents pollute
  the naive-Claude test). Open decision: unify the `restructure_section`/`reorder_sections`
  verb/plural mismatch IF the API isn't externally frozen.

## Grounding
- `docs/SPEC.md` — the spec itself (v0612-2: + attachments/embedded objects §5 → Phase 5b;
  create_notebook/copy_notebook removed; tool-description enhancement §4 + diagnostic log §7
  → Phase 6).
- `docs/com-api-reference.md` — COM signatures + enums (from Microsoft Learn).
- `docs/onenote-xml-schema.md` — `one:` page/hierarchy XML + format-preservation rules.
- `docs/vm-setup.md` — Phase 0b Windows VM build (autologon, desktop OneNote, COM smoke, Tier-2).
- `refs/mhzarem-onenote-mcp/` — reference clone (gitignored). Payload shapes only; its
  backup-parse + PowerShell model is rejected.
