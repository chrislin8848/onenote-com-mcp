# PROPOSAL — `apply_text_style` (batch font / size / color patch) → v1.1.0

Status: DESIGN (confirmed fork: whole-page also rewrites QuickStyleDef). Grounded empirically
against `tests/fixtures/` (the 混合樣式頁 dump) on 2026-06-14.

## Goal
One tool that changes the **font-family / size / color** of all text in a scope (default = the
whole page; or one outline/table/OE subtree) in ONE read-mutate-write, preserving everything else
(bold / italic / underline / strike / the colors & sizes you did NOT ask to change / highlight /
hyperlinks / images / tables / untouched paragraphs). Collapses the current "dozens of
`update_page_content(replace)` calls, each re-supplying every run, each preceded by a `get_page`"
into "one call per page, one read". This is the canonical *format-preservation = self-consistency,
not byte-identity* operation.

Primary driver: "change the font of ●ITIN and its subpages to 微軟正黑體" was very slow.

## Empirical ground truth (混合樣式頁, via the real parse layer)
The effective run style is a three-layer cascade (`parse._effective_style`), lowest → highest:

```
QuickStyleDef baseline   ←   OE @style attribute   ←   inline <span style>      (span wins)
```

Observed in the real fixture — **all three layers are actively in use**:
- `QuickStyleDef`: `[0] PageTitle Calibri 20.0`, `[1] p Calibri 11.0` — the page baseline is Calibri.
- `OE @style`: MANY paragraphs carry `font-family:Microsoft JhengHei;font-size:12.0pt` at the OE
  level (layer 2). Bare-text runs (`span={}`) get their font from here.
- `<span>`: per-run `font-family` (Microsoft JhengHei / Calibri / 新細明體), plus `color:#FA0000`,
  `font-weight:bold`, `font-style:italic`, `text-decoration:underline|line-through`, and dual
  `background`+`mso-highlight` for highlight.

**Two consequences that shape the design:**
1. **The span layer wins**, so writing `font-family:X` into *every run's span* makes X effective
   for all existing visible text — in ANY scope. This is the workhorse mechanism.
2. **Changing only QuickStyleDef is nearly invisible on real pages** (OE @style + spans shadow the
   baseline). So QuickStyleDef rewrite is *not* what makes existing text change — it only fixes the
   baseline for future-typed text / empty paragraphs, and it is **page-global** (can't be scoped).

## Mechanism (inside the existing `apply_page_edit` seam — one read-mutate-write)

```python
def mutate(tree):
    scope = _find_content_object(tree, scope_object_id) if scope_object_id else tree
    overlay = {}                                  # only the keys the caller asked for
    if font_family: overlay["font-family"] = font_family
    if size is not None: overlay["font-size"] = f"{float(size)}pt"
    if color: overlay["color"] = color
    for t in scope.iter(qn("T")):
        if scope is tree and _under_title(t): continue      # whole-page skips the page title
        runs = parse_spans(t.text)
        for r in runs:
            r.span_style = {**r.span_style, **overlay}        # MERGE — keep bold/color/link/...
        if runs:
            t.text = etree.CDATA(build_spans(runs))
    if scope is tree:                                         # WHOLE-PAGE ONLY (baseline)
        for qd in tree.findall(qn("QuickStyleDef")):
            if font_family: qd.set("font", font_family)
            if size is not None: qd.set("fontSize", f"{float(size)}")
            if color: qd.set("fontColor", color)
```

- **Preservation is automatic**: the overlay only writes the requested keys into each run's
  `span_style`; bold (`font-weight`), italic, underline/strike (`text-decoration`), the highlight
  pair, the link (`Run.link` → `<a>`), and any color/size you did NOT change all ride through. The
  `changed_objects` payload strategy prunes scope-external outlines → they are byte-identical
  (never re-sent).
- **Scope** uses the existing `_find_content_object`; `iter(qn("T"))` descends into table cells and
  nested OE children automatically. A scope with no text → 0 runs changed (not an error).
- **Whole-page** also rewrites every `QuickStyleDef` (font/fontSize/fontColor). QuickStyleDef is a
  definition child (not in `_CONTENT_TAGS`) so it always ships — editing it is safe. For a SUB-scope
  we MUST NOT touch QuickStyleDef (it is shared page-wide).

## Required companion fix — `build_spans` must preserve `lang`
Ground truth: real runs carry `lang` (en-US / zh-TW), but `build_spans` currently **drops it**
(parser reads `lang`, builder never writes it). The existing replace/set_rows edit paths already
silently strip `lang`; `apply_text_style` would amplify that to every touched run. Fix `build_spans`
to emit `lang` (and emit a `<span>` when a run has `lang` even without style). The existing span
tests pass dict runs with no `lang`, so this does not break them. This is a small fidelity win for
ALL edit paths.

## Known limit (documented; VM question, NOT chased in v1)
An **empty** paragraph whose `OE @style` pins an old `font-family` is not reached by the span
overlay (no runs) and the OE-style font shadows the rewritten QuickStyleDef baseline — so it could
still type in the old font. Rare; deferred. (Fast-follow if it bites: also rewrite OE-@style font
in whole-page mode.) Existing VISIBLE text is always correct because the span layer wins.

## Tool surface
`apply_text_style(page_id, *, font_family=None, size=None, color=None, scope_object_id="",
force=False)`
- At least one of `font_family` / `size` / `color` is required (else `ValueError`).
- `size` is points (number); `color` is a hex string (e.g. `#FA0000`).
- `scope_object_id` omitted = whole page (all outlines + tables, page title excluded); an
  outline/table/OE objectID = just that subtree.
- Returns `{scope, runs_changed, quick_styles_updated, paragraphs_seen}` so the model narrates by
  effect ("changed the whole page to 微軟正黑體 — 47 runs"), not mechanics.
- §4 borders: vs `update_page_content(replace)` (ONE paragraph, re-supplies its runs) and
  `modify_table(set_rows)` (whole cells) — `apply_text_style` is a STYLE-ONLY patch across MANY
  runs, content untouched. (Do NOT mention the OneNote App Ctrl+A shortcut — users already know it.)

## Staging (project discipline: user confirms between stages)
- **Stage 1 (host, Tier-1):** `build_spans` lang fix (+span tests); `apply_text_style` core in
  `page_edit.py` + `service` facade. Tier-1 on the real 混合樣式頁: scope-external byte-identical;
  in-scope font changed; bold/italic/underline/color/highlight/link/lang preserved; size-only and
  color-only patches; sub-scope (one outline / one table) leaves siblings byte-identical;
  QuickStyleDef rewritten ONLY on whole-page; return counts.
- **Stage 2 (host):** server tool (#29 — catalog 28→29) wrapped in `logged_tool`; `_SERVER_INSTRUCTIONS`
  (App-Ctrl+A-first + contrastive border); `test_smoke_server` catalog + border asserts.
- **Stage 3 (VM, Tier-2):** format-preservation regression — whole-page font change keeps
  bold/color/links live; sub-scope leaves siblings byte-identical; CJK font renders; **VM question:
  does rewriting QuickStyleDef actually re-render baseline paragraphs?**
- **Release:** version 1.0.10 → **1.1.0** (pyproject + __init__ + .iss + uv.lock); one installer
  rebuild (OneNoteMCP-Setup_1.1.0.exe).

## Deferred (out of v1)
`replace_font(from:[...], to:...)` conditional = this tool + a source-font filter; OE-@style font
rewrite for empty paragraphs; cross-page batching is just N calls (one per page); cross-tool
transaction (skip — `apply_page_edit` is already one read-mutate-write per page).
