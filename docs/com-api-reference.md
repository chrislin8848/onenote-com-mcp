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
