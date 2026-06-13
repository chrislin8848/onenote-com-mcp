# Handoff — copy under-sync detection & reporting (post-v1.0.0, 2026-06-13)

> For the next session/account continuing this work. (The diagnosis below lives only here +
> in git — it is NOT in any per-account memory.)

## TL;DR — the "OCR images can't be copied" problem was MISDIAGNOSED

It is **not** an OCR limitation and **not** a COM limitation. The real cause is **OneDrive
files-on-demand**: image/attachment binaries that haven't been downloaded to *this* machine yet
make `GetBinaryPageContent` return `0x8004200F` (`hrBinaryObjectDoesNotExist`) and leave the
attachment cache file absent. A **fully-synced** machine copies everything fine.

Proof (VM, 2026-06-13) — same `測試章節2` section, three sync states, one gradient:

| machine / state | images copied OK |
|---|---|
| fully-synced PC (Chris's) | ALL (he confirmed the D1 page's images present) |
| VM, section "downloaded" (partial hydration) | 42 / 156 |
| VM, freshly-added, not-yet-synced (`(同步測試)`) | 0 / 156 — all placeholders |

The OCR correlation in the old notes was coincidental (OCR'd photos are the large, last-to-
hydrate ones). A copied placeholder does **not** self-heal (dead bytes / broken reference); only
the **live original** self-heals via OneNote. So the remedy is always: **fully sync the source,
then copy.**

## What this commit changed

Copy path now DETECTS + REPORTS under-synced content (SPEC §5: explicit, never silent) instead of
silently shipping blanks:

- `service/page_edit.py::inline_image_binaries` → returns the **count** of images that fell back to
  the 1×1 placeholder (was a notes list). Docstring corrected (sync, not OCR; placeholders don't
  self-heal). `apply_page_edit` still ignores this return — see TODO #2.
- `service/copy.py`:
  - `_rewrite_inserted_files` → returns `(missing_files, missing_objects, notes)`, categorizing
    each `one:InsertedFile` by kind (`one:Previews` child ⇒ **embedded object**; else ⇒ **file**).
  - `PageCopyResult` / `SectionCopyResult` gained `missing_images` / `missing_files` /
    `missing_objects` counts.
  - `sync_warning(images, files, objects)` builds the user-facing warning (returns `None` when
    nothing is missing).
  - `transfer_section` aggregates the counts across pages.
- `server.py` `copy_page` / `copy_section` return `sync_warning` in the JSON; descriptions instruct
  Claude to ALWAYS surface a non-null `sync_warning` and advise the user to fully sync + re-copy.
- Tests: new `tests/test_copy_sync.py` (count / categorize / warning); `tests/test_copy_files.py`
  wording updated. **Tier-1: 218 passed, ruff clean.**

## TODO (continuing session)

1. **Tier-2 validate `sync_warning` live.** Copy an under-synced section on the VM; confirm the
   warning lists images + files + embedded objects, including attachments (the D1 `GFK…pdf`).
   COM must run in the autologon interactive session (schtasks `/it`, like `scripts/run_tier2.bat`),
   NEVER plain SSH (a hung COM call blocks OneNote's single-threaded server). VM =
   `dev@192.168.122.13` (DHCP — `virsh domifaddr --source agent win11-onenote`); repo at
   `C:\onenote-mcp` with a synced `.venv`.
2. **Edit path warns too?** `service/page_edit.py::apply_page_edit` ignores `inline_image_binaries`'
   return, so editing a page with un-synced images silently placeholders them. Decide whether the
   edit tools should surface the same warning.
3. **Optional product step (Chris's idea):** a *pre-copy* sync check — detect un-synced content and
   block/prompt (or trigger + await OneDrive hydration) BEFORE copying, instead of only warning
   after. The warning shipped here is the minimum; this is the nicer UX.
4. **Doc hygiene.** Correct the stale "OCR-processed images un-extractable via COM" framing where it
   still appears (this CLAUDE.md "Real-data resilience" status paragraph; `docs/com-api-reference.md`;
   `docs/onenote-xml-schema.md` if any) to the sync explanation above.
5. **Parked — do NOT pursue unless requirements change:** disk-parse via pyOneNote (would fork an
   MS-ONESTORE parser — a fragile patch-treadmill; investigated and rejected 2026-06-13) and COM
   `Publish`-render. Both are UNNECESSARY now that the cause is sync.

## VM state left by this session

- pyOneNote uninstalled; all probe artifacts + the `onenote-probesync` / `onenote-copytest`
  scheduled tasks and `scripts/probe_sync*` / `scripts/copy_test*` removed.
- **MCP Test sections were KEPT** (Chris's request) for further testing: the downloaded `測試章節2`,
  the `(同步測試)` section, and the test copies. **Nothing in OneNote was deleted.**
