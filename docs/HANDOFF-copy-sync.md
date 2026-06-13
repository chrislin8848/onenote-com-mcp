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

- `service/page_edit.py`: `inline_image_binaries` returns the **count** of un-fetchable images and
  takes `remove_unfetchable`. COPY path (True) → **REMOVES** the image and prunes the emptied
  OE/Outline via the new `remove_content_element()` helper (a copy must not carry a dead placeholder
  that can't self-heal and could later be misread / re-copied; pruning the empty container is what
  avoids the hrInvalidXML that made the original "just drop it" fail). EDIT path (False, default) →
  keeps a 1×1 placeholder (must not delete a cloud-only image). Docstring corrected (sync, not OCR).
- `service/copy.py`:
  - `_rewrite_inserted_files` → returns `(missing_files, missing_objects, notes)`, categorizing each
    `one:InsertedFile` by kind (`one:Previews` child ⇒ **embedded object**; else ⇒ **file**), and
    **REMOVES** the un-synced ones (prune) instead of leaving a broken reference.
  - `PageCopyResult` / `SectionCopyResult` gained `missing_images` / `missing_files` /
    `missing_objects` counts; `transfer_section` aggregates them.
  - `sync_warning(images, files, objects)` builds the user-facing warning (None when nothing missing;
    "… OMITTED … would not self-heal … sync the source, then copy again").
- `server.py` `copy_page` / `copy_section` return `sync_warning` in the JSON; descriptions instruct
  Claude to ALWAYS surface a non-null `sync_warning` and advise the user to fully sync + re-copy.
- Tests: new `tests/test_copy_sync.py`; `tests/test_copy_files.py` updated. **Tier-1: 219 passed,
  ruff clean. Tier-2 validated** (un-synced 156-image section: copy completed with NO hrInvalidXML,
  148 images REMOVED, correct sync_warning). **Released as v1.0.1.**

## TODO (continuing session)

1. **Tier-2 validate `sync_warning` live.** ✅ **Image path DONE (2026-06-13):** copied the
   un-synced `測試章節2(同步測試)` section live — `transfer_section` returned `missing_images=148`,
   `missing_files=0`, `missing_objects=0`, the correct `sync_warning` ("148 image(s) … not yet
   downloaded …"), correct per-page notes, NO regression; and a now-synced attachment (the D1 PDF)
   was correctly NOT flagged. ⏳ **File/object path still OPEN live:** the PDF synced before the run
   so the un-synced-attachment branch wasn't exercised (it IS covered at Tier-1 —
   `test_copy_sync.py`, `test_copy_files.py`). To finish: copy a section while an attachment is
   still un-downloaded (or clear OneNote's local cache to force it). Re-run helper left on the VM at
   `C:\onenote-mcp\scripts\tier2_sync_validate.py` — register an interactive schtasks task (`/it`,
   like `scripts/run_tier2.bat`) to run it. COM must run in the autologon interactive session,
   NEVER plain SSH (a hung COM call blocks OneNote's single-threaded server). VM =
   `dev@192.168.122.13` (DHCP — `virsh domifaddr --source agent win11-onenote`); repo at
   `C:\onenote-mcp` with a synced `.venv`.
2. **Edit-path warning** — ❌ **decided NO (Chris, 2026-06-13):** only warn at copy time; the edit
   path stays as-is (keeps the placeholder; must not delete a cloud-only image).
3. **Pre-copy sync check** — ❌ **decided NO (Chris):** the copy-time `sync_warning` is sufficient.
4. **Doc hygiene** — ✅ **DONE:** the CLAUDE.md "Real-data resilience" paragraph now states the sync
   diagnosis (not OCR). `com-api-reference.md` / `onenote-xml-schema.md` had no stale OCR wording.
5. **Parked — do NOT pursue unless requirements change:** disk-parse via pyOneNote (would fork an
   MS-ONESTORE parser — a fragile patch-treadmill; investigated and rejected 2026-06-13) and COM
   `Publish`-render. Both are UNNECESSARY now that the cause is sync.
6. **Remaining (Chris):** file/object Tier-2 live (he is testing on his own fully/partially-synced
   PC) and the real-Claude-Desktop install test of the 1.0.1 build.

## VM state left by this session

- pyOneNote uninstalled; all probe artifacts + the `onenote-probesync` / `onenote-copytest`
  scheduled tasks and `scripts/probe_sync*` / `scripts/copy_test*` removed.
- **MCP Test sections were KEPT** (Chris's request) for further testing: the downloaded `測試章節2`,
  the `(同步測試)` section, and the test copies. **Nothing in OneNote was deleted.**
