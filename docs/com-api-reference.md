# OneNote Desktop COM API — distilled reference

Source: Microsoft Learn "Application interface (OneNote)" and "Enumerations (OneNote
developer reference)". Captured 2026-06-10 for grounding. This is the **authoritative
contract** the `OneNoteBackend` interface mirrors 1:1.

> Red line (SPEC §9): everything below is reached **only** through pywin32 COM via
> `Win32ComBackend`. No Microsoft Graph, no `msal`, no HTTP to `graph.microsoft.com`,
> no Azure app/token. COM runs inside the signed-in user's own OneNote session.

## Namespace

OneNote 2013 schema namespace (the one we target):

```
http://schemas.microsoft.com/office/onenote/2013/onenote
```

Always request/emit `xs2013` explicitly (never `xsCurrent` — it breaks across OneNote
versions). Note: older docs examples show the `.../12/2004/onenote` 2007 namespace —
**do not** use that one.

## Methods (verbatim signatures → our backend method)

COM `[out]` params are returned by pywin32, not passed in. Exact out-param marshalling
**must be confirmed on the VM in Phase 3** (early vs late binding differs).

### Notebook structure

| COM | Signature | Backend method |
|---|---|---|
| `GetHierarchy` | `(BSTR bstrStartNodeID, HierarchyScope hsScope, [out]BSTR* xml, [in,def xs2013]XMLSchema)` | `get_hierarchy(start_node_id, scope) -> str` |
| `UpdateHierarchy` | `(BSTR bstrChangesXmlIn, [in,def xsCurrent]XMLSchema)` | `update_hierarchy(changes_xml) -> None` |
| `OpenHierarchy` | `(BSTR bstrPath, BSTR bstrRelativeToObjectID, [out]BSTR* objectID, [in,def cftNone]CreateFileType)` | `open_hierarchy(path, relative_to_id, create_file_type) -> str` |
| `DeleteHierarchy` | `(BSTR bstrObjectID, [in,def 0]DATE dateExpectedLastModified, [in,def false]VARIANT_BOOL deletePermanently)` | `delete_hierarchy(object_id, expected_last_modified=None, permanent=False) -> None` |
| `CreateNewPage` | `(BSTR bstrSectionID, [out]BSTR* pageID, [in,def npsDefault]NewPageStyle)` | `create_new_page(section_id, style=npsDefault) -> str` |
| `GetHierarchyParent` | `(BSTR bstrObjectID, [out]BSTR* parentID)` | `get_hierarchy_parent(object_id) -> str` |
| `GetSpecialLocation` | `(SpecialLocation slToGet, [out]BSTR* path)` | `get_special_location(location) -> str` |
| `CloseNotebook` | `(BSTR bstrNotebookID, [in,def false]VARIANT_BOOL force)` | (not in MVP) |

### Page content

| COM | Signature | Backend method |
|---|---|---|
| `GetPageContent` | `(BSTR bstrPageID, [out]BSTR* xml, [in,def piBasic]PageInfo, [in,def xsCurrent]XMLSchema)` | `get_page_content(page_id, page_info=piBasic) -> str` |
| `UpdatePageContent` | `(BSTR bstrPageChangesXmlIn, [in,def 0]DATE dateExpectedLastModified, [in,def xsCurrent]XMLSchema, [in,def false]VARIANT_BOOL force)` | `update_page_content(changes_xml, expected_last_modified=None, force=False) -> None` |
| `GetBinaryPageContent` | `(BSTR bstrPageID, BSTR bstrCallbackID, [out]BSTR* b64)` | `get_binary_page_content(page_id, callback_id) -> str` (base64) |
| `DeletePageContent` | `(BSTR bstrPageID, BSTR bstrObjectID, [in,def 0]DATE dateExpectedLastModified, [in,def false]VARIANT_BOOL force)` | `delete_page_content(page_id, object_id, expected_last_modified=None, force=False) -> None` |

### Navigation

| COM | Signature | Backend method |
|---|---|---|
| `FindPages` | `(BSTR bstrStartNodeID, BSTR bstrSearchString, [out]BSTR* xml, [in,def false]VARIANT_BOOL fIncludeUnindexedPages, [in,def false]VARIANT_BOOL fDisplay, [in,def]XMLSchema)` | `find_pages(start_node_id, query, include_unindexed=False) -> str` |
| `GetHyperlinkToObject` | `(BSTR bstrHierarchyID, BSTR bstrPageContentObjectID, [out]BSTR* hyperlink)` | `get_hyperlink_to_object(hierarchy_id, object_id="") -> str` |

### Windows interface (current viewing context)

`Application.Windows` is a *property* (window collection); `Windows.CurrentWindow` is the
active window, whose `CurrentNotebookId` / `CurrentSectionGroupId` / `CurrentSectionId` /
`CurrentPageId` give the user's current location → backend
`get_current_window_ids() -> CurrentWindowIds`. Limits (SPEC §5): with no open window there
is nothing to read — raise `NoCurrentWindowError`, never guess; multi-window resolves to the
active one; **granularity stops at the page** — no COM API exposes in-page cursor position or
selected text. Property (vs method) marshalling under early binding is a Phase 3 VM check.

## Semantics that drive the design

- **`UpdatePageContent` is a page-level-object merge, not a whole-page replace.** It only
  touches page-level objects (`Outline`, `Image`, `Ink`...) whose IDs you include; objects
  you omit are left untouched. It *entirely replaces* a page-level object whose ID matches
  — so you must include that object's **full** content, not a diff. This is exactly why the
  SPEC mandates the `GetPageContent → mutate the real XML tree in place → UpdatePageContent`
  core, and why you must NOT rebuild from a slimmed/plain-text model.
- **You cannot delete a page-level object by omitting it** from `UpdatePageContent` (merge
  semantics). Use `DeletePageContent(pageID, objectID)`.
- **`DeleteHierarchy` is for hierarchy nodes only** (section group / section / page) — not
  page content objects.
- **Concurrency guard:** pass `dateExpectedLastModified` (the page's `lastModifiedTime` from
  the prior `GetPageContent`). If the page changed since, the call fails instead of clobbering
  the user's edits. Default `force=False`. Surface conflicts up; never silently `force`.
- **`GetPageContent` omits binary by default** (`piBasic`). Binary objects (`Image`, `Ink`)
  carry a `callbackID`; fetch real bytes with `GetBinaryPageContent(pageID, callbackID)`.
- **`CreateNewPage` adds a blank last page**; set position/level via `UpdateHierarchy`
  (`pageLevel` for subpages). The docs note `UpdateHierarchy` gives more control and can make
  subpages directly.
- **`UpdateHierarchy` ordering semantics (SPEC §5 discipline):** there is NO position-index
  attribute — order is the child-element order of the submitted XML. Microsoft documents that a
  *partial* child list makes OneNote "infer" placement of omitted siblings, unpredictably.
  Restructures must therefore submit the scope's **complete child list, in target order, in one
  batch** (`apply_hierarchy_restructure` in `service/hierarchy_edit.py` enforces this). A
  notebook's direct children are a **mixed** `one:Section` + `one:SectionGroup` list — the batch
  must contain both kinds. Validated surface: page order within a section, section order within
  a notebook; top-level notebook ordering is unvalidated and out of scope; cross-section page
  moves are experimental until VM-validated.
- **`OpenHierarchy` is create-or-open** for hierarchy nodes: section uses `cftSection` (path
  ends `.one`, `relativeTo` = notebook ID); notebook uses `cftNotebook` (needs a real
  filesystem/OneDrive path so it syncs — a bare name makes a local-only notebook).
- **`FindPages` needs Windows Search** and is scoped by `bstrStartNodeID`. Always scope to a
  notebook/section — never search from root across everything.

## Enumerations (member = value)

```
HierarchyScope:  hsSelf=0  hsChildren=1  hsNotebooks=2  hsSections=3  hsPages=4
PageInfo:        piBasic=0  piBinaryData=1  piSelection=2  piBinaryDataSelection=3
                 piFileType=4  piBinaryDataFileType=5  piSelectionFileType=6  piAll=7
CreateFileType:  cftNone=0  cftNotebook=1  cftFolder=2  cftSection=3
NewPageStyle:    npsDefault=0  npsBlankPageWithTitle=1  npsBlankPageNoTitle=2
XMLSchema:       xs2007=0  xs2010=1  xs2013=2  xsCurrent=2
SpecialLocation: slBackupFolder=0  slUnfiledNotesSection=1  slDefaultNotebookFolder=2
```

## Error handling

- OneNote is busy/syncing → `RPC_E_SERVERCALL_RETRYLATER` (`0x8001010A`, signed `-2147417846`)
  and sometimes `RPC_E_CALL_REJECTED` (`0x80010001`). Retry with backoff.
- pywin32 raises `pywintypes.com_error`; the HRESULT is in `excepinfo`/`hresult`.

## Lessons taken from `mhzarem/onenote-mcp` (reference only)

We reuse the **payload shapes** and tool-naming intuition, NOT its execution model:

- `one:T` accepts **inline HTML only** (`b`, `i`, `span`, `br`). Block-level tags (`p`,
  `div`, `h1`, `table`, `ul`...) make `UpdatePageContent` **silently fail**. Our builder emits
  proper `one:Table/one:OE` structure instead of dumping HTML, sidestepping this.
- `]]>` inside text breaks the CDATA wrapper — must be escaped.
- mhzarem reads via **backup-file parsing (pyOneNote)** and writes via **PowerShell
  subprocess**. SPEC §1.1 forbids both: we use live COM through in-process pywin32 only.

## VM-validated COM behaviors (Phase 4 Tier-2, 2026-06-11)

Discovered against real OneNote (M365 desktop, early-bound makepy module) — these override
any doc sketch above:

- **`VT_DATE` params accept ONLY a PyTime instance.** A plain int, the makepy default tuple,
  and omitting the parameter all raise `TypeError: must be a pywintypes time object`; and
  pythoncom cannot marshal pre-1970 stamps (mktime → `OSError`), so the documented
  "DATE 0 ⇒ skip the check" is **unreachable from Python**. The backend therefore always
  carries a real stamp and resolves the node's CURRENT `lastModifiedTime` when the caller has
  none. `pywintypes.Time` does NOT localize tz-aware datetimes — OneNote's UTC `…Z` stamps
  must be converted to local naive before the comparison.
- **The real error code hides in `com_error`'s excepinfo.** A server-side refusal surfaces as
  DISP_E_EXCEPTION; the OneNote HRESULT (e.g. `hrLastModifiedDateDidNotMatch` 0x80042010,
  `hrInvalidXML` 0x80042001) is the excepinfo tuple's `scode` (last element).
- **`force=true` does NOT bypass `dateExpectedLastModified`** — it only overrides
  unsaved-UI-edit protection. A forced write must carry the page's current stamp.
- **`GetPageContent`'s `lastModifiedTime` does not refresh promptly after a programmatic
  `UpdatePageContent`** (observed unchanged ≥10 s). Served and compared stamps stay consistent
  with each other, so read-then-write flows are unaffected.
- **`UpdateHierarchy` ignores a `one:Page` `name` attribute** — the hierarchy page name derives
  from the title. A page rename is a Title edit via `UpdatePageContent`.
- **The hierarchy schema is positional: all `one:Section` children precede all
  `one:SectionGroup` siblings** at the same level (the UI renders groups last); an interleaved
  child order is rejected with `hrInvalidXML`.
- **A page moved to another section gets a NEW page ID** (page IDs embed the owning section
  GUID). `move_page` re-reads the target section and returns the new ID.
- **`DeletePageContent` refuses paragraph-level `one:OE` targets** (hr 0x8004200E) — it is for
  page-level objects (Outline/Image/…). Deleting a paragraph = submitting its outline without
  it via `UpdatePageContent` (merge replaces a submitted object's content wholesale). Phase 6
  must scope `delete_page_content` accordingly.
- **An empty `<one:Cell/>` is rejected with `hrInvalidXML` (0x80042001)** — a table cell MUST
  keep `one:OEChildren > one:OE` (VM ground truth 2026-06-13, 祕魯18天: removing the only image
  in a cell and pruning its OEChildren left an empty cell → the whole `modify_table`/copy write
  failed). The fix replenishes a cleared cell with a minimal empty paragraph. (`modify_table`
  add/delete column/row edits ride the same `UpdatePageContent` core; no new COM signature.)

## VM-validated: UpdatePageContent has NO sub-outline merge; big tables are slow (2026-10-05)

Probed (`scripts/probe_bigtable.py`) after a real report: a single-cell `find_and_replace` on a
106×39 table page (4,134 cells) timed out at the client's 60s.

- **An `one:Outline` in the payload REPLACES that outline's whole content — there is no merge
  by objectID below the outline.** Live: an outline carrying only p2 (of p1/p2/p3) → p1 and p3
  DELETED; a table carrying only one full `one:Row` → every other row DELETED; a `one:Row`
  carrying only one of its `one:Cell`s → REJECTED. So the `changed_objects` granularity (whole
  page-level objects) is the floor: editing ONE cell re-submits the whole outline = the whole
  table. A "sparse" payload is NOT a valid optimization — it destroys content.
- **UpdatePageContent cost scales with the table's cell count, not the payload bytes.** Creating
  a 4,134-cell table (414K-char payload) took 39s in one call (~10ms/cell). Every write that
  touches such a table's outline pays this again.
- **Editing an EXISTING big outline is ~2x slower than creating it:** a single-cell
  find_and_replace on a copy of the real 106×39 page = 41–46s in UpdatePageContent (48–50s
  end-to-end with the read), while transplanting the whole page onto a blank page took 19s.
  Close to — and on a slower PC over — Claude Desktop's 60s client timeout. The write is NOT
  cancelled by a client timeout; it completes in OneNote.
- **Stripping per-object author/timestamp attributes does NOT help:** payload 3.6M → 1.37M
  chars, write time unchanged (40.7–45.3s vs 41.4–41.9s). Cost is per object, not per byte.
  (OneNote kept the untouched cells' author/creationTime when they were omitted.)
- **Reads are cheap by comparison:** GetPageContent of the real page = 4.0M chars in ~2.5s;
  `get_table(text_only=True)` = 21.9K chars vs 1.40M (full, compact JSON) / ~2.9M (old,
  indented).
- **THE dominant factor: is the page DISPLAYED in OneNote?** Same single-cell edit on the same
  copy, alternating: page on screen = **141s / 134s**; another page on screen = **23s / 26s**
  (OneNote redraws every cell of the rewritten box). The earlier 41–54s figures were taken with
  an unknown window state. The user's 60s timeout was almost certainly "editing the 班表 while
  looking at it". → `apply_page_edit` refuses a big write (>= 1,000 OEs in the payload) to the
  page `Windows.CurrentWindow.CurrentPageId` reports, with `PageDisplayedError`, unless
  `allow_displayed=True`; `get_table`'s `write_cost` note says so up front.
- **Rebuild-as-new-page is not worth it:** read → edit in memory → transplant onto a blank page
  → delete the old one = 17.8s total (2.8 + 14.2 + 0.8) with the old page hidden — only ~6s
  faster than an in-place edit with the page hidden, at the cost of a new page ID, new objectIDs
  for every cell, lost page history and a full-page re-sync. Rejected.
- **Mitigation is behavioural, not a payload trick:** batch every change to a big table into ONE
  write (batch_update / set_rows / set_column), and treat a timed-out write as probably-landed
  (re-read before retrying). `get_table` emits a `write_cost` note at >= 1,000 cells. A user can
  also split a huge table into several separate content boxes (outlines) — the outline, not the
  table, is the rewrite unit.

## VM-validated COM behaviors (Phase 5 Tier-2, 2026-06-11)

- **`OpenHierarchy(cftNotebook)` cannot create notebooks on this M365 build** — it returns
  `hrFileDoesNotExist` (0x80042006) for BOTH local folder paths (even with the parent folder
  present) and OneDrive `https://` parents. BOTH notebook-creation tools were REMOVED
  (user-approved): `copy_notebook` and `create_notebook`. Notebooks are created in the
  OneNote UI; whole-notebook cloning = `copy_section` per section into an existing
  notebook/group.
- **`OpenHierarchy` OPENS an existing same-named section/group instead of creating one** — a
  copy without name de-collision would silently merge into the existing node. Copies append
  " (2)", " (3)", … against the target parent's direct children of both kinds.
- **`DeletePageContent` works on `one:Outline` objects** (page-level) — confirmed live while
  sweeping empty outlines; together with the Phase-4 finding (paragraph OEs refused), Phase 6's
  `delete_page_content` contract is: page-level objects only.
- **Outlines holding only empty paragraphs render as nothing and are dropped by OneNote when
  transplanted** — semantic fidelity (B≡A) is asserted ignoring them.
- B≡A validated live: mixed-format/table/image pages clone with identical semantic
  fingerprints and byte-identical image pixels; section copies preserve page order and
  1/2/3 subpage levels; section copies land inside section groups.

## VM-validated COM behaviors (Phase 5b Tier-2, 2026-06-12)

- **Staged-pathSource re-import WORKS.** A clone whose `one:InsertedFile` carries
  `pathSource` → a staged copy of the cache file (and NO `pathCache`) is accepted by
  `UpdatePageContent`, and OneNote re-imports the file: all five attachments of the test
  page had `pathCache` rebuilt on the copy within 45 s of the write.
- **Embedded objects survive the clone as embedded objects** — the copied xlsx still reports
  `kind=embedded_preview` with its `one:Previews`/`one:Preview page="工作表1"` intact (the
  Previews structure rides verbatim and OneNote honors it).
- **Printout render PNGs are RE-ENCODED on transplant** — unlike normal image inserts (which
  round-trip byte-identical, Phase 5), the large render PNG comes back with identical pixel
  dimensions but different bytes. Image fidelity for printout renders is therefore semantic
  (dimensions/visual), not byte-level — consistent with the SPEC §6 boundary.
- **`DeletePageContent` ACCEPTS a page-level `one:InsertedFile`** (its own objectID) — the
  attachment disappears, siblings untouched. Phase 6 `delete_page_content` can target them
  directly.
- **`DeletePageContent` REFUSES an inline attachment-bearing `one:OE`** — same refusal class
  as paragraph OEs (Phase 4, 0x8004200E). Inline-attachment deletion in Phase 6 must be an
  outline rewrite through the edit seam, not a DeletePageContent call.
- **Editing a printout-bearing page through `changed_objects` works live** — the pruned
  payload (no InsertedFile/XPSFile/render, Phase 5b `_CONTENT_TAGS`) is accepted and the
  merge leaves attachments, XPSFile carrier, and render untouched.
- (Host-side, same date) **`pathCache` GUIDs are per-read ephemera** — basic vs piBinaryData
  dumps of the same page carry different Temp GUIDs for the same attachments; never persist
  a pathCache across reads.

## VM-validated COM behaviors (Phase 6 Stage 5 — PyInstaller freeze, 2026-06-12)

- **The frozen exe binds OneNote via COM — the gen_py regeneration approach WORKS.** A
  PyInstaller onedir build of the server, run as `OneNoteMCP.exe --selftest` in the autologon
  interactive session, connected to OneNote and listed 3 notebooks (exit 0). This validates the
  freeze recipe: keep `gencache.EnsureModule` (OneNote can't be late-bound — Phase 0b), bundle
  the makepy machinery (spec `hiddenimports`: win32com.client.makepy/genpy/build/gencache/
  selecttlb), and redirect `win32com.__gen_path__` to a writable temp dir via a runtime hook so
  EnsureModule can generate the typelib module on first call. **SPEC §8's "use late-bound
  Dispatch for freeze compatibility" does NOT apply here** — late-bound Dispatch never worked for
  OneNote; the writable-gen_py path is the right fix and is now proven.
- The freeze + `uv sync --group packaging` are COM-free and run over plain SSH (session 0); only
  the `--selftest` COM smoke needs the interactive task (`onenote-selftest`, like the Tier-2
  loop). Inno Setup (iscc.exe) is NOT on the VM yet — the installer build is gated on installing
  it; the freeze + COM smoke are the validated parts.

## UpdateHierarchy has NO optimistic-concurrency protection (fact-finding, 2026-06-12)

Triple-confirmed — type-library signature, Microsoft docs, AND live VM behavior all agree:

- **Signature (makepy, OneNote 15.0 TypeLib `{0EA692EE-…}` v1.1):**
  `UpdateHierarchy(bstrChangesXmlIn, xsSchema=2)` — only two params. It is the ONLY mutating
  method on `IApplication` WITHOUT a `dateExpectedLastModified`; `UpdatePageContent`,
  `DeleteHierarchy`, and `DeletePageContent` all carry it (the first/last also carry `force`).
- **Docs:** the `UpdateHierarchy` page lists only those two params and says nothing about
  concurrency; the "proceeds only if the value matches … prevents accidentally overwriting"
  language appears ONLY on the three methods that have the date param. The `UpdateHierarchy`
  example even strips `lastModifiedTime` from the submitted XML (it is a decorative/output
  attribute, not a write-path token).
- **Live behavior (`tests/test_windows_concurrency.py`, Tier-2 green):** a stale read
  resubmitted after a second COM call reordered the section is applied SILENTLY and clobbers
  the middle change (last-write-wins, no error); a reorder whose body carries a stale
  `lastModifiedTime` (backdated to 2000) still applies — so UpdateHierarchy reads no
  concurrency token from the body either.
- **Consequence (Phase 6 hardening):** hierarchy/structural writes cannot be optimistically
  guarded at the COM layer. Compensate in the service layer — minimize the read→write window
  (already: read-full → mutate → submit in one call), optionally re-read after the write and
  report any diff vs. intended order/names/pageLevel, and keep propose-confirm + suggest a
  clone backup before structural ops. Page-CONTENT writes keep their real `UpdatePageContent`
  date guard (→ `ConcurrencyError`) — the gap is structural only.
