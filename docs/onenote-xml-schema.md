# OneNote 2013 page/hierarchy XML — schema notes for the parse/build layer

Grounding for Phase 1 (`onenote_com_mcp.xmllayer`). The **format-preservation requirement
(SPEC §5)** is the single hardest constraint, so this doc is written around it.

> ⚠ **Fixtures are not synthetic in this project.** Per the decision on 2026-06-10, the
> parse/build layer is TDD'd against **real** fixtures dumped from OneNote on the Windows VM
> (`scripts/dump_fixtures.py`), not hand-authored XML. The shapes below are from Microsoft
> docs + the `mhzarem` reference and are accurate enough to *design* the API, but every
> detail (attribute spelling, `QuickStyleDef` shape, highlight encoding) is **confirmed
> against a real dump before the parser is trusted.** Treat this doc as the design target,
> the dumped fixtures as ground truth.

## Ground-truth findings (2026-06-11, real VM dumps — Phase 1)

Where the real fixtures disagreed with the sketches below, the fixtures won. The parser
(`xmllayer/parse.py`, `xmllayer/spans.py`) bakes these in:

- **`one:T` CDATA is word-wrapped by OneNote** — newlines appear between attributes AND
  *inside* attribute values (`style='font-family:\n\n"Microsoft JhengHei"'`). Parse the
  CDATA with a real HTML parser (`html.parser`), never a regex; collapse whitespace when
  splitting style declarations.
- **Attribute quoting varies**: `style` is single-quoted, font names may be double-quoted,
  single-quoted, or bare CJK (`font-family:新細明體`); `lang=zh-TW` is unquoted. HTML
  entities (`&nbsp;`) appear inside the CDATA and must decode (→ U+00A0).
- **`one:Image` carries `one:CallbackID` as a CHILD element** (`<one:CallbackID
  callbackID="{…}"/>`), not a `callbackID` attribute as sketched below — read both. No
  `format` attribute was present on the dumped image.
- **`one:Image` has NO `objectID` of its own** (confirmed Phase 2): the deletable page-content
  object is the enclosing `one:OE`, which holds the `objectID`. So `delete_page_content`
  (Phase 6) for an image must target the wrapping OE's ID, and the read tools surface that OE
  ID as the image block's `object_id`. Tables and cells *do* carry their own `objectID`.
- **Inline `one:Data` may be absent even on a `piBinaryData` dump** (observed on the 圖片頁
  dump); image bytes reliably come from `GetBinaryPageContent(callback_id)` instead. The
  copy path (Phase 5) must verify on the VM what `piBinaryData` actually inlines.
- **The OE `style` attribute is a third style layer**: effective run style =
  `QuickStyleDef` baseline (via `quickStyleIndex`) ← OE `style` attr ← inline span style.
  The sketch below only shows two layers. OE `style` may itself contain `color`,
  `text-align`, etc., and is also word-wrapped.
- **Table cells** carry `shadingColor` on `one:Cell`; alignment lives on the cell's inner
  OE (`alignment` + `text-align` in its `style`). An empty cell is `<one:T/>` (present but
  empty). `one:Table` has `hasHeaderRow`; `one:Column` has `isLocked`.
- The mixed-style page's `QuickStyleDef`s carried only
  `name/fontColor/highlightColor/font/fontSize/spaceBefore/spaceAfter` (no `bold`/`italic`
  attrs observed — the parser still accepts them).

Namespace prefix `one:` = `http://schemas.microsoft.com/office/onenote/2013/onenote`.

## Hierarchy XML (`GetHierarchy` / `FindPages` output)

```xml
<one:Notebooks xmlns:one="...">
  <one:Notebook name="..." nickname="..." ID="{GUID}{1}{B0}" path="..." lastModifiedTime="...">
    <one:SectionGroup name="..." ID="..." isRecycleBin="true|false">
      <one:Section name="..." ID="..." path="...one" lastModifiedTime="..." color="..." locked="true|false">
        <one:Page ID="..." name="..." dateTime="..." lastModifiedTime="..."
                  pageLevel="1|2|3" isCurrentlyViewed="true|false"/>
      </one:Section>
    </one:SectionGroup>
  </one:Notebook>
</one:Notebooks>
```

- Every node carries `ID`. `pageLevel` (1/2/3) is the subpage indent — **must be preserved
  on copy** (SPEC §4/§5).
- Skip `isInRecycleBin`/`isRecycleBin` nodes in listings.
- Scope listings with `start_node_id` — never enumerate from root.

## Page content XML (`GetPageContent` output)

```xml
<one:Page xmlns:one="..." ID="{GUID}{n}{B0}" name="Title"
          dateTime="..." lastModifiedTime="2026-06-10T...Z" pageLevel="1" lang="en-US">
  <one:QuickStyleDef index="0" name="p" fontColor="automatic" highlightColor="automatic"
                     font="Calibri" fontSize="11.0" bold="false" italic="false" .../>
  <one:QuickStyleDef index="1" name="h1" .../>
  <one:PageSettings RTL="false" color="automatic"> ... </one:PageSettings>
  <one:Title>
    <one:OE quickStyleIndex="1"><one:T><![CDATA[Title]]></one:T></one:OE>
  </one:Title>
  <one:Outline>
    <one:Position x="36.0" y="84.0"/>
    <one:Size width="..." height="..."/>
    <one:OEChildren>
      <one:OE quickStyleIndex="0" alignment="left">
        <one:T><![CDATA[plain run ]]></one:T>
        <one:T><![CDATA[<span style='font-weight:bold;color:#FF0000'>styled run</span>]]></one:T>
      </one:OE>
      <one:OE><one:Table .../></one:OE>
    </one:OEChildren>
  </one:Outline>
  <one:Image format="png" callbackID="{GUID}{n}{B0}">
    <one:Position .../><one:Size .../>
    <one:Data><!-- base64, ONLY when piBinaryData requested; else absent + callbackID --></one:Data>
  </one:Image>
</one:Page>
```

### The two-level style model (the crux of format preservation)

Formatting lives in **two** places and BOTH must round-trip:

1. **Paragraph/run baseline → `QuickStyleDef` + `quickStyleIndex`.** Each `one:OE` (and
   `one:T` in title) references a `quickStyleIndex`; the matching top-level
   `one:QuickStyleDef index="N"` defines the baseline `font`, `fontSize`, `fontColor`,
   `highlightColor`, `bold`, `italic`, etc. for that paragraph. If you drop the
   `QuickStyleDef` table or renumber indices, paragraph styling silently changes.
2. **Inline character overrides → `<span style=...>` inside the `one:T` CDATA.** The text of
   a `one:T` is HTML; bold/italic/underline/strike, `color`, highlight, `font-family`,
   `font-size` appear as inline `<span style>` runs. A `one:T` may hold multiple spans +
   bare text.

**Parser** must surface BOTH layers (resolve effective style per run = QuickStyleDef baseline
overlaid with inline span). Represent runs as `(text, resolved_style)` PLUS keep the raw
node — never collapse to plain text.

**Builder** for *new* content emits `one:T` with `<span style>` and sets the `OE`
`quickStyleIndex`. For *edits*, see "editing model" below.

### Highlight is dual-attribute — a known gotcha (SPEC §5)

OneNote writes highlight as **both** `background:yellow` (standard CSS) **and**
`mso-highlight:yellow` (Office hint) on the span. Parser reads either; **builder must write
both**, or some OneNote builds render inconsistently. Color may be any CSS color
(`yellow`, `#FFFF00`, ...).

### Tables

```xml
<one:Table bordersVisible="true">
  <one:Columns><one:Column index="0" width=".."/>...</one:Columns>
  <one:Row>
    <one:Cell><one:OEChildren><one:OE><one:T>...</one:T></one:OE></one:OEChildren></one:Cell>
  </one:Row>
</one:Table>
```

Cells contain `OE`/`T` — same style rules apply. Parse to structured rows (list of list of
cell-content), **never** flatten to one string. Build full `one:Table` XML for create/modify.

### Images

- `one:Image` has `format` + `callbackID`; binary `one:Data` (base64) is present only with
  `piBinaryData`. Read: `get_page_images` → `GetBinaryPageContent(callbackID)` → return as MCP
  image content. Insert: emit `one:Image` with inline base64 `one:Data`.

### Attachments / embedded objects (`one:InsertedFile`) — DESIGN TARGET, pending VM dump

⚠ Phase 5b: everything here must be confirmed against a real dump of an attachment page /
embedded spreadsheet / file printout before the parser is trusted (SPEC v0612-2 §5).

- Sketch: `<one:InsertedFile pathSource="..." pathCache="..." preferredName="..."
  objectID="...">` with optional `Previews` / `Printout` children (these drive the `kind`
  discrimination: attachment icon vs embedded preview vs printout — exact rule TBD on dump).
- **Binary location differs from images:** content is the on-disk cache file at `pathCache`
  (OneNote-managed), read directly — NOT `GetBinaryPageContent`. Cache may be absent
  (unsynced/purged): degrade gracefully, report "cache unavailable".
- Copy: old `pathCache` is a dead reference in the clone — copy the cache file aside, point
  `pathSource` at the copy, drop `pathCache`, let OneNote re-import (mechanics VM-gated).
- `Printout` renders as page images — burned-in pixels, same limits as images.
- No write path (no `insert_file`) this round, by SPEC decision.

## Editing model (in-place tree mutation — SPEC §5, hard rule)

The edit data model **is the lxml tree itself**, not a DTO. Flow:

```
xml = GetPageContent(page_id)            # real, full XML (incl. QuickStyleDef table)
tree = lxml.etree.fromstring(xml)
# locate target node by ID/xpath; mutate IN PLACE (append/insert/replace that subtree only)
UpdatePageContent(etree.tostring(tree), expected_last_modified=page.lastModifiedTime)
```

Because `UpdatePageContent` replaces a page-level object whole-for-whole by ID, we send back
the **entire** edited tree (or the full edited page-level object), preserving every untouched
paragraph's `quickStyleIndex` and inline spans verbatim. Rebuilding from a model would drop
formatting on untouched paragraphs — forbidden.

## Copy model (raw-XML transfer — SPEC §5, separate path from editing)

`copy_*` does NOT go through the editing/structured model. It:

1. `GetPageContent` with **binary inlined** (`piBinaryData`) so images travel as `one:Data`.
2. Carries the source page's `QuickStyleDef` table along (else `quickStyleIndex` dangles).
3. **Strips/resets object `ID`s** so OneNote mints fresh ones in the target.
4. Sets target `pageLevel` from source (preserve subpage hierarchy).
5. Writes into the target section (new page via `CreateNewPage`, then `UpdatePageContent`
   with the transplanted body).

"Copy then modify" (e.g. date shifts) reuses `get_page` + `update_page_content` (in-place
edit) on the *copy* — it is not a separate tool.

## Self-consistency, not byte-identity (SPEC §5 honesty note)

OneNote re-normalizes spans / re-numbers `QuickStyleDef` on its own redraw, so a round-trip
won't be byte-identical. Preservation is verified at the **semantic level**: the user-visible
font/size/color of untouched paragraphs is unchanged. Regression tests assert *semantic
equality of untouched paragraphs*, not raw-XML equality.
