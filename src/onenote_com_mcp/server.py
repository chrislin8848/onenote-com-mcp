"""FastMCP server — the full 30-tool OneNote catalog (SPEC §4).

Every tool is a thin facade over ``onenote_com_mcp.service`` (the shared write core, the copy
core, the hierarchy core); the orchestration lives there, not here. Descriptions carry the §4
contract the LLM reads: contrastive borders on confusable pairs (get_page vs get_page_info vs
get_table vs get_page_images vs get_page_files_info vs get_page_files; update_page_content vs
create_table vs modify_table;
restructure_section vs reposition_page vs reorder_sections vs move_page vs rename_node;
delete_node vs delete_page_content vs delete_inline_content; copy vs move), DESTRUCTIVE +
propose-then-confirm contracts in text, and
``mode`` as a per-value enum. Cross-tool rules that belong to no single tool live in
``_SERVER_INSTRUCTIONS`` (MCP ``initialize`` instructions). ``logged_tool`` adds the §7 per-call
diagnostic log at the decorator seam.

Description quality is ACCEPTANCE-TESTED on real Claude Desktop (SPEC §4 — a naive Claude that
sees only these descriptions must pick the right tool; CC sub-agents are explicitly not a valid
proxy). Pending: unifying the restructure_section/reorder_sections verb-plural mismatch (Chris's
call — only if the API is not externally frozen).

stdio transport: nothing but MCP protocol may go to stdout. Logs go to stderr (SPEC §8).
"""

from __future__ import annotations

import base64
import functools
import json
import sys
import threading
from typing import Literal

from mcp.server.fastmcp import FastMCP, Image

from onenote_com_mcp.backend import get_backend
from onenote_com_mcp.errors import NoCurrentWindowError, NodeNotFoundError
from onenote_com_mcp.logging_config import configure_logging, log_tool_call
from onenote_com_mcp.service import copy, create, delete, files, hierarchy_edit, page_edit, read

# Server-level guidance (SPEC §4): cross-tool rules that belong to no single tool. Claude
# Desktop's uptake of `instructions` is to be confirmed empirically (Stage-3 acceptance).
_SERVER_INSTRUCTIONS = """\
This server edits the LIVE OneNote desktop app: every create/update/delete/copy/restructure \
takes effect immediately in the real notebooks — there is no staging copy and no undo beyond \
OneNote's own recycle bin. Work conservatively.

Two-step rule for in-page objects: to edit or delete something INSIDE a page (a paragraph, \
table, image, or attachment) you first need its objectID — get it from get_page_info (a cheap, \
FLAT, EXHAUSTIVE inventory of every object's id + type + which delete tool removes it; the \
preferred first step), get_page (the full text/style content), get_page_images (image pixels), \
or get_page_files_info (attachment metadata); to find the object that CONTAINS a given string \
(e.g. the paragraph with a typo) use find_objects. To find/delete ALL images on a page, use \
get_page_info — it lists images nested in table cells AND page-level printout renders, which \
get_page's nested tree can bury; do NOT eyeball get_page to hunt for images, and when sweeping \
several pages check EACH page's inventory rather than assuming later pages match earlier ones. \
Picking the right tool but omitting the objectID it needs is as wrong as picking the wrong tool. \
get_page_info's `preview` is a TRUNCATED label (trailing "…" = more follows), NOT content — never \
proofread or judge a paragraph as clean from it; read the full text with get_object / get_page, or \
search with find_objects, to verify text correctness.

objectIDs and node IDs (page/section/notebook/section-group IDs) are INTERNAL plumbing — \
use them to chain calls, but do NOT surface them to the user. They are long, opaque, and \
meaningless to a human reading the conversation. NEVER paste a raw ID — and ESPECIALLY never \
a LIST of IDs — into your reply as a way to refer to pages/objects; that is exactly the wrong \
way to report progress. Refer to things by their NAME instead ("copied ●ITIN and its 9 \
subpages: 官網行程, D1-0107, …", "the third paragraph", "the first table"). Results that carry \
IDs also carry names for this reason — e.g. copy_page / copy_pages / copy_page_subtree return \
each new page's name and level alongside its id; narrate from the names. Only show a raw ID \
when the user explicitly asks for it, when names alone are genuinely ambiguous (two pages share \
a title and the user must disambiguate), or when the user will paste it back into a tool call.

Keep user-facing narration BRIEF and in plain language. Do NOT walk the user through your \
implementation steps — tool names, parameter shapes (e.g. set_rows arrays, 0-indexed columns), \
objectIDs, or other internal mechanics are noise to them. Say what you are doing in human terms \
("updating the dates on D1–D6", "removing the printout images from these two pages") and report \
the outcome by name; keep the plumbing inside the tool calls, not in your prose.

Finding a section's pages and a page's subpages: list_pages returns ALL pages in a section, in \
order, with each page's level — use it to ENUMERATE. search_pages is FULL-TEXT search: it returns \
only pages whose title/content matches the query, so it MUST NOT be used to list a section's pages \
or to gather a page's subpages — it silently misses every page that doesn't match. A page's \
subpages are the consecutive pages that follow it at a DEEPER level (up to the next same-or- \
shallower page). To act on "this page and all its subpages" (e.g. restyle every subpage, copy a \
subtree), call list_pages, take that page plus the deeper-level run beneath it, and operate on \
EACH — never lean on search_pages to discover them.

Adding pictures, two ways. For a VECTOR graphic (a diagram, map, chart, simple banner) use \
insert_svg_image — you generate SVG markup and the server renders it to an image. For a RASTER \
image that exists as a FILE ON DISK — a photo the user pointed you at, or \
(in a client that can write files or run code, e.g. an agentic IDE) an image you produced on \
disk such as a code-rendered chart or a downloaded picture — use insert_image_from_path with \
the file path; the server reads the bytes off disk, so they never pass through you. If a raster \
image exists only in your context with NO file on disk you cannot insert it — tell the user to \
add it BY HAND in the OneNote app (drag-and-drop, or Insert ▸ Picture/File). Never try to insert \
a picture by emitting base64 or by smuggling a raster <image data:…> inside the SVG (rejected, \
and slow). For CJK text in the SVG, use an explicit \
Windows font-family such as "Microsoft JhengHei", not the generic "sans-serif". After \
insert_svg_image succeeds, trust the result — do NOT routinely read the image back with \
get_page_images to "verify" it (pulling the whole rasterized PNG back as base64 is slow); read it \
back only if the user reports a rendering problem. Get the SVG layout right in one pass: leave \
margins and keep labels from overlapping nodes or markers. The picture can be POSITIONED, like \
text: both insert tools take mode insert_before / insert_after with a paragraph objectID so it \
lands MID-page, not only at the end (default mode append). Copying a page or section still carries \
its existing images and attachments along — fully supported.

Restyling text in bulk: to change FONT / SIZE / COLOR / HIGHLIGHT or toggle BOLD / ITALIC / \
UNDERLINE / STRIKETHROUGH across many paragraphs at once — "make this whole page 微軟正黑體", \
"every heading 16pt", "the body blue", "make these pages bold + italic", "highlight pages 1-5 \
yellow" — use apply_text_style, NOT a string of update_page_content("replace") calls. It patches \
styling across every run in scope (default the whole page; or one outline/table/paragraph/cell \
objectID) in ONE pass and preserves everything else (the emphasis/colors/sizes you did not touch, \
links, images, tables); color/highlight are colors (highlight="none" removes it) and bold/italic/ \
underline/strikethrough are tri-state (true=on, false=off, omit=leave). TABLE row/column: a whole \
ROW = pass that row's id (get_page's table.row_object_ids) as scope_object_id; a whole COLUMN has \
no id, so pass columns=[j] (0-indexed) with the table as scope — both work for text AND for \
cell_shading. A yellow background in a TABLE may be the TEXT highlight (behind the words) OR the \
whole-CELL shading — different \
attributes: highlight clears the text marker, cell_shading="none" clears the cell background. If \
"remove the yellow" doesn't fully work, clear BOTH (highlight="none", cell_shading="none"). \
update_page_content("replace") is for rewriting ONE paragraph's text; apply_text_style changes \
style only, never the words. To restyle a page AND its subpages, enumerate with list_pages, then \
call it per page. When CREATING content (create_page or update_page_content), do NOT hand-repeat \
the same font/size/color block on every paragraph: write the text first, then run apply_text_style \
once on the WHOLE page to set the font — it rewrites the page's baseline style so every paragraph \
(and future typing) inherits it.

Editing tables: pick the narrowest operation instead of rebuilding the table. To READ just one \
table (e.g. a long Guest List) use get_table, not the whole-page get_page. To REARRANGE columns or \
rows — move a column to the front, drop one and shift the rest — use modify_table \
reorder_columns / reorder_rows with the complete target order (e.g. order=[2,0,1]); do NOT \
clear and re-type cells \
with set_rows just to reorder. To rewrite ONE column's text use modify_table set_column (a flat \
list, one value per row) rather than a full set_rows grid of mostly-unchanged cells; to add a \
column WITH content in one step use insert_columns with values; to change a whole column's STYLE \
or COLOR (not its text) use apply_text_style(columns=[j]); to edit ONE cell's text use \
update_page_content("replace") on that cell's paragraph objectID. To CLEAR a table's body while \
keeping a header row or column, do not re-supply the kept text: set_rows treats a None cell as \
"leave it unchanged" — e.g. [None, "", ""] keeps column 0 and clears the rest.

Appending content: update_page_content mode=append adds to the page's LAST outline by default. If \
a page has several outlines and you must add to a SPECIFIC one, pass that outline's objectID as \
target_object_id, or the content may land in an unexpected outline. Append once per addition: if a \
call seems to time out, do NOT blindly re-append — it may already have succeeded; re-read with \
get_page_info to check before retrying, or you will paste the same content twice.

Text fidelity: the server stores text byte-for-byte — it does no Unicode normalization and no font \
substitution, so the characters you send are written exactly. Homoglyph and simplified/traditional \
slips happen at GENERATION time, not in storage; if you are not fully confident a rare or easily- \
confused CJK character will come out right, pin the exact codepoint by writing it as its JSON \
\\uXXXX escape rather than the glyph.

Targeted text edits without re-typing: to fix a typo or change a word, do NOT re-supply the whole \
paragraph. find_and_replace(find, replace) swaps text in place across the page (or within one \
object) keeping each run's style — the cheapest fix, and it cannot corrupt text you did not touch. \
To find WHICH paragraph to edit, find_objects(query) returns the matching objectIDs on a page \
(search_pages only finds whole PAGES, not locations); get_object(object_id) reads ONE paragraph's \
full text + style without the whole-page payload. When you have SEVERAL edits to one page, \
batch_update applies them in ONE atomic write (all-or-nothing) instead of a burst of \
update_page_content calls. After an edit, pass return_ids=True (on update_page_content or \
batch_update) to get the affected or newly created objectIDs back instead of re-reading the page.

Editing a page = edit it IN PLACE (update_page_content, modify_table, delete_inline_content, \
insert_svg_image); this is the normal, expected, safe-enough path for ordinary changes. Do NOT \
rebuild a page from scratch (create a new page, re-emit the content, recycle the old) just to \
change it: a new page gets a NEW ID — breaking links and subpage structure — loses the page's \
history, and may force you to re-insert images you cannot recover. Build a NEW page only when the \
user genuinely wants a new or merged page (e.g. integrating two pages into one). For large or \
risky edits the cheap, safe pattern is copy-then-modify: copy_page/copy_section the page, then \
edit the COPY freely — the original is your backup.

Match how much you confirm to the RISK; do NOT gate everything. Just DO it and report afterwards \
(no pre-confirm) for reversible or lossless operations: editing in place (update_page_content, \
find_and_replace, batch_update, modify_table insert_rows/insert_columns/set_rows), rename_node, \
reposition_page, reorder_sections, \
restructure_section (these only rename, reorder, or re-level — no data is lost), and moving a \
SINGLE page or section to the recycle bin (do it, then report that it is recoverable). Propose and \
get explicit go-ahead FIRST for irreversible or large-scope operations: permanent deletes \
(permanent=True), removing in-page content (delete_page_content, delete_inline_content — the old \
content is NOT recoverable), modify_table's delete_rows/delete_columns, force overwrites, deleting \
a whole section/section-group or several pages at once, notebook-wide restructures (suggest a \
copy_section clone backup for these), and move_page (it gives the page a new ID and is \
experimental). EXCEPTION: when you have just made a faithful copy \
(copy_page/copy_pages/copy_page_subtree/copy_section) specifically in order to modify it, \
destructive edits to THAT copy need NO confirmation — the original is the backup and editing the \
copy is the whole point; confirm destructive ops only on original/source content.

Positioning a page: copy_page and create_page already place the new page right BELOW its natural \
anchor by default — copy_page below the SOURCE page, create_page below the page the user is \
currently on — so "copy this page" / "add a page here" need no position argument. Pass their \
after_page_id only to place the page after a DIFFERENT page. To move an EXISTING page use \
reposition_page. For a SINGLE page do NOT reach for restructure_section's whole-list reorder, and \
NEVER write an external scratch file to organize the order: reorder lists in context.

Copying SEVERAL pages to a position: do NOT call copy_page repeatedly — each copy lands below its \
own source, scattering them. To copy a page together with its subpages (e.g. "copy this page and \
its subpages below page X") use copy_page_subtree(page_id, after_page_id=X); for an arbitrary set \
of pages use copy_pages(page_ids, after_page_id=X). Both place the copies as ONE contiguous block, \
in order, in a single placement step.

Pace heavy calls: OneNote's COM is single-threaded — it runs ONE operation at a time. Send copy_* \
and other heavy writes (a large update_page_content, insert_svg_image) in SMALL BATCHES — a few at \
a time (e.g. 3-5) — and wait for each batch to return before sending more; do NOT fire a long \
burst of them in a single turn. Bursting is not faster (the server serializes the calls) and only \
risks timeouts; light reads you may issue freely.

Benchmark workflow for "copy these pages and change the dates" (faithful copy, then edit the \
copy): (1) the user manually creates a synced notebook B in the OneNote UI (COM cannot create \
notebooks); (2) copy_section clones each source section into B — a perfect, mechanical copy; \
(3) search_pages (scoped to B) finds the pages with dates; (4) get_page reads the clean text; \
(5) update_page_content (replace) changes only the date paragraphs, preserving all other \
formatting. Verify B≡A before editing, then verify the dates."""

mcp = FastMCP("onenote", instructions=_SERVER_INSTRUCTIONS)


# OneNote's COM server is single-threaded (STA): it runs ONE call at a time. Concurrent calls do
# NOT parallelize — the loser is rejected with RPC_E_SERVERCALL_RETRYLATER and bounces off the
# backend's finite busy-retry budget (~16s; see win32com_backend._call), so a burst of tool calls
# (e.g. several copy_page in one assistant turn) can exhaust that budget and FAIL even for small
# pages. FastMCP runs our sync tools in a worker-thread pool, so such a burst really does arrive in
# parallel. This process-wide lock serializes every tool at the COM boundary: contenders wait in an
# orderly queue (cheap, no retry budget burned) and each then runs against a free server instead of
# fighting for it. Uncontended — the normal case — acquiring it is ~free. Safe because tools call
# the service layer, never another @logged_tool function, so it can't re-enter and self-deadlock.
_COM_LOCK = threading.Lock()


def _serialize_com(func):
    """Run the tool body under the process-wide COM lock (see _COM_LOCK)."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        with _COM_LOCK:
            return func(*args, **kwargs)

    return wrapper


def logged_tool(*args, **kwargs):
    """``@mcp.tool`` + the §7 per-call diagnostic log + COM serialization, in one decorator.

    Composes so FastMCP sees the wrapped function (signature/annotations preserved by
    functools.wraps, so the generated tool schema is unchanged): logging is OUTSIDE the lock (a
    call is recorded the moment it arrives, before it queues) and serialization is INNERMOST,
    around the actual COM work. Tools are facades over ``onenote_com_mcp.service``."""

    def decorate(func):
        return mcp.tool(*args, **kwargs)(log_tool_call(_serialize_com(func)))

    return decorate


def _json(data: object) -> str:
    # ensure_ascii=False keeps CJK note content readable in the tool result
    return json.dumps(data, ensure_ascii=False, indent=2)


# --- Read (Phase 2: wired to FixtureBackend on Linux) -----------------------


@logged_tool()
def list_notebooks() -> str:
    """List all open OneNote notebooks (name + ID)."""
    return _json(read.list_notebooks(get_backend()))


@logged_tool()
def list_sections(notebook_id: str) -> str:
    """List sections in a notebook (name + ID), preserving section-group nesting
    (one:SectionGroup containers appear as nested groups, not flattened)."""
    return _json(read.list_sections(get_backend(), notebook_id))


@logged_tool()
def list_pages(section_id: str) -> str:
    """List ALL pages in a section, in order, each with its subpage level (pageLevel). This is the
    COMPLETE page list — use it (NOT search_pages) to enumerate a section's pages or to find a
    page's SUBPAGES: a page's subpages are the consecutive pages that follow it at a DEEPER
    pageLevel, up to the next page at the same or a shallower level. To act on a page together with
    all its subpages (restyle, copy, …), take that run of pages from here and operate on each."""
    return _json(read.list_pages(get_backend(), section_id))


@logged_tool()
def search_pages(query: str, scope_id: str = "") -> str:
    """Full-text search: returns ONLY pages whose title/content MATCHES query (scope to a
    notebook/section ID — recommended). It is NOT a way to enumerate a section's pages or to gather
    a page's subpages — it SILENTLY MISSES every page that does not match the query. To list a
    section's pages, or to find a page and all its subpages, use list_pages instead."""
    return _json(read.search_pages(get_backend(), query, scope_id))


@logged_tool()
def get_page(page_id: str) -> str:
    """Read a page's full CONTENT — rich-text runs with resolved styles, structured tables, and
    the objectID of every content object. Includes both outline content (`outlines`) AND
    page-level objects (`page_level_images` / `page_level_files` — printout renders and page-level
    attachments that live outside any outline). This is the primary page read for working with
    TEXT and the source of the objectIDs that update_page_content / create_table /
    delete_page_content need. It returns images and attachments as lightweight references (object
    IDs + metadata), NOT their bytes — use get_page_images for image pixels and get_page_files /
    get_page_files_info for attachments. When you only need to know WHAT objects a page has and
    their IDs (e.g. to find/delete every image, including page-level printout renders) — not the
    full text — use get_page_info instead: it is a cheaper, flat, exhaustive object inventory.
    When you only need ONE table's contents (e.g. a big table-heavy page where this full read is
    large), use get_table — it returns just that table, compactly."""
    return _json(read.get_page(get_backend(), page_id))


@logged_tool()
def get_page_info(page_id: str) -> str:
    """Lightweight INVENTORY of every object on a page — each object's id, type, which delete tool
    removes it (`delete_with`), whether it is page-level, and light type metadata (image
    width/height/OCR-flag, table rows×cols, file name/kind, a short paragraph text preview). It
    does NOT return full text runs, the style table, or pixels — for the full content use get_page;
    for image pixels use get_page_images; for attachment content use get_page_files. Use this as the
    cheap first step of the two-step objectID rule, and especially to find ALL images to delete:
    the list is FLAT and EXHAUSTIVE — it includes images nested in table cells and page-level
    printout renders, which get_page's structured tree can bury. Each entry's `object_id` paired
    with its `delete_with` is DIRECTLY ACTIONABLE: pass that exact object_id to the named delete
    tool — you do NOT need to re-read with get_page to "verify" or translate the id (it is already
    the right target). Do NOT eyeball get_page's nested JSON to hunt for images; read this flat
    list, and across several pages call it per page rather than assuming pages with no images near
    the top have none lower down. The `preview` is a TRUNCATED identifying label (a trailing "…"
    means more text follows) — NOT the content. NEVER judge whether a paragraph is correct/clean or
    proofread for typos from `preview`: corruption can hide past the cutoff. To check text
    correctness read the full text with get_object (one object), get_page (whole page), or locate a
    known string with find_objects (it matches the FULL paragraph text)."""
    return _json(read.get_page_info(get_backend(), page_id))


@logged_tool()
def get_table(page_id: str, table_object_id: str) -> str:
    """Read ONE table's structured content — its columns, every cell (text, runs, shading color),
    and the row/cell objectIDs — without the rest of the page. Use this instead of get_page when you
    only care about a specific table, especially on a big table-heavy page (a long Guest List etc.)
    where get_page returns a large payload: this is the compact, table-only read. Get the
    table_object_id from get_page or get_page_info first (get_page_info lists each table with its id
    + rows×cols but NOT its cell contents; get_table is what returns the contents). It finds the
    table anywhere on the page, including one nested inside a cell. To then EDIT the table use
    modify_table (shape/bulk content) or update_page_content ("replace" for one cell)."""
    return _json(read.get_table(get_backend(), page_id, table_object_id))


@logged_tool()
def get_object(page_id: str, object_id: str) -> str:
    """Read ONE object on a page by its objectID — a paragraph's full text + runs + style, or a
    table / image / attachment — WITHOUT the rest of the page. The targeted companion to get_page:
    to inspect or fix a single paragraph (e.g. one find_objects pointed you at), fetch just it
    instead of the whole-page payload (which repeats every run under both "text" and "runs" and
    carries the page-wide style table). Get the object_id from get_page_info or find_objects. Finds
    the object anywhere on the page (inline, nested in a table cell, or page-level)."""
    return _json(read.get_object(get_backend(), page_id, object_id))


@logged_tool()
def find_objects(page_id: str, query: str) -> str:
    """Find the objects on a page whose TEXT contains a substring, returning their objectIDs — the
    within-page counterpart to search_pages (which returns whole PAGES, never a location on a page).
    Use it to pinpoint which paragraph(s) to edit, e.g. the one holding a typo, without dumping the
    whole page or reading get_page_info's TRUNCATED previews (it matches the FULL paragraph text).
    Searches body and table-cell paragraphs; returns each match's object_id (ready for get_object /
    update_page_content / find_and_replace) and a short preview. Case-sensitive (so 開鑿 ≠ 開逑)."""
    return _json(read.find_objects(get_backend(), page_id, query))


# structured_output=False: the return is image content, not a JSON schema — FastMCP can't
# build a pydantic output schema for Image, and we don't want one here.
@logged_tool(structured_output=False)
def get_page_images(page_id: str) -> list[Image]:
    """Return the PIXELS of a page's embedded images (one:Image) as viewable image content,
    so they can be recognized visually. This is bytes only — for an image's objectID,
    dimensions, or OCR text use get_page. For files attached to the page (PDFs, Office docs,
    text files — one:InsertedFile, shown as an icon or embedded preview, not an inline image)
    use get_page_files_info / get_page_files instead."""
    return [
        Image(data=base64.b64decode(img["data_base64"]), format=img["media_type"].split("/")[-1])
        for img in read.get_page_images(get_backend(), page_id)
    ]


@logged_tool()
def get_page_files_info(page_id: str) -> str:
    """File-EXTRACTION pre-check for a page's attachments/embedded objects (one:InsertedFile):
    per file it adds size_bytes, cache_available, extension, and media_class (text / image / pdf /
    unsupported) — i.e. whether and how get_page_files can extract its CONTENT. Use this right
    before get_page_files (it is that tool's prerequisite), NOT as the general "what's on this
    page" tool: to merely DISCOVER a page's files/objects and their IDs (e.g. to delete an
    attachment) use get_page_info, which is cheaper (no disk read), covers every object type, and
    includes page-level objects. Not for inline images (use get_page_images). Reads no content."""
    return _json(files.get_page_files_info(get_backend(), page_id))


# structured_output=False: entries may be MCP image content (see get_page_images).
@logged_tool(structured_output=False)
def get_page_files(page_id: str, object_id: str = "", max_chars: int = 50000) -> list[str | Image]:
    """Extract attachment CONTENT, for three supported types ONLY: text-class files (decoded
    text), image attachments (returned as viewable image content), and PDFs (server-side text
    extraction). Other types (docx/xlsx/pptx/…) return metadata + an explicit "unsupported" —
    run get_page_files_info first to check each file's type/size (and get_page_info for its
    objectID). This reads one:InsertedFile
    attachments, NOT inline page images (those are get_page_images). object_id narrows to one
    attachment; text is truncated at max_chars; a missing/unsynced cache is reported per file,
    never a crash."""
    out: list[str | Image] = []
    for entry in files.get_page_files(get_backend(), page_id, object_id, max_chars):
        image_b64 = entry.pop("data_base64", None)
        out.append(_json(entry))
        if image_b64 is not None:
            out.append(
                Image(data=base64.b64decode(image_b64), format=entry["media_type"].split("/")[-1])
            )
    return out


@logged_tool()
def get_current_context() -> str:
    """Where is the user right now? Returns the active OneNote window's current notebook /
    section group / section / page (IDs + names), so the user can say "this page" or
    "the current section". Granularity stops at the page — in-page cursor position and
    selected text are not available. Errors clearly if OneNote has no open window.
    Before acting on this context, report it back ("you're currently on page X") so the
    user can confirm they haven't switched pages since."""
    return _json(read.get_current_context(get_backend()))


# --- Create (service/create.py) ----------------------------------------------
# NOTE: there is deliberately no create_notebook tool. VM ground truth (2026-06-11): this
# M365 OneNote build refuses COM notebook creation — OpenHierarchy(cftNotebook) returns
# hrFileDoesNotExist for local paths AND OneDrive https parents alike. Notebooks are created
# in the OneNote UI; create_section covers everything below them.


@logged_tool()
def create_section(parent_id: str, name: str) -> str:
    """Create a new empty SECTION under parent_id — an existing notebook OR a section group —
    inheriting its sync. For a new PAGE use create_page; there is no create_notebook (make
    notebooks in the OneNote UI, then add sections here)."""
    return _json({"section_id": create.create_section(get_backend(), parent_id, name)})


@logged_tool()
def create_page(
    section_id: str,
    title: str,
    content: str | list[dict] = "",
    page_level: int = 1,
    after_page_id: str = "",
) -> str:
    """Create a new PAGE in a section. Use this to make a page; to add content to a page that
    already exists, use update_page_content. By DEFAULT the new page is placed right BELOW the page
    the user is currently viewing (from get_current_context) when that page is in this section —
    what you usually want when creating a page while reading one. Pass after_page_id to place it
    after a specific page instead; if there is no current page in this section the new page is added
    at the END (then reposition_page can move it). page_level (1/2/3) sets subpage indent. content
    optionally adds initial paragraphs — same shapes as update_page_content (plain text with
    newlines, or styled paragraph dicts)."""
    backend = get_backend()
    # Default anchor = the page the user is currently on — read BEFORE creating, so the anchor is
    # where they were (not the freshly created page). An explicit after_page_id wins; no open
    # window → no default (append at end).
    anchor = after_page_id
    if not anchor:
        try:
            anchor = backend.get_current_window_ids().page_id or ""
        except NoCurrentWindowError:
            anchor = ""
    new_id = create.create_page(backend, section_id, title, content, page_level)
    if anchor and anchor != new_id:
        try:
            hierarchy_edit.reposition_page(backend, section_id, new_id, after_page_id=anchor)
        except NodeNotFoundError:
            if after_page_id:
                raise  # an explicitly named anchor that is not in this section is a real error
            # the current page is in a DIFFERENT section → leave the new page at the section end
    return _json({"page_id": new_id})


# --- Modify (shared write core — service/page_edit.py) -----------------------


@logged_tool()
def update_page_content(
    page_id: str,
    content: str | list[dict],
    mode: Literal["append", "insert_before", "insert_after", "replace"] = "append",
    target_object_id: str = "",
    force: bool = False,
    return_ids: bool = False,
) -> str:
    """Edit a page's TEXT/paragraphs surgically — untouched paragraphs keep their formatting
    verbatim. Use this to add, insert, or rewrite text and styled paragraphs (size/font/color/
    highlight/hyperlink), and to edit a table CELL's text. NOT for: creating a table (use
    create_table); changing a table's row/column COUNT, i.e. adding/deleting rows or columns
    (use modify_table); removing a whole outline/image/attachment (use delete_page_content).
    (To add a picture: a VECTOR graphic via insert_svg_image, or a raster image that is a FILE on
    disk via insert_image_from_path.)

    mode (per value):
      "append"        — add paragraphs at the end of an outline; target_object_id optionally
                        names an outline objectID (default = the page's last outline).
      "insert_before" — insert new paragraphs just before target_object_id (a paragraph
                        objectID from get_page); they become its siblings.
      "insert_after"  — same, just after the target paragraph.
      "replace"       — swap target_object_id's text, keeping its paragraph style unless the
                        new content overrides it. Editing one table cell = "replace" on the
                        paragraph objectID inside that cell.

    content: plain text (newlines split paragraphs) OR a list of paragraph dicts —
    {"text": "...", "style": {...}} or {"runs": [{"text": "...", "style": {...}}, ...]},
    each optionally with "quick_style_index" / "alignment". Style keys are the CSS-like keys
    get_page returns: font-weight, font-style, text-decoration, color, background (highlight),
    font-family, font-size. A run (or a {"text": ...} paragraph) may carry "link": "https://…"
    to make that text a HYPERLINK. To restyle or hyperlink EXISTING text, read it with get_page,
    then "replace" the paragraph re-supplying its runs with the changed style/link (get_page
    reports each run's existing "link" so a replace round-trips it instead of dropping it).

    Concurrency-guarded: fails instead of clobbering if the page changed since it was read.
    force=True overwrites anyway — DESTRUCTIVE, only after explicit user confirmation.

    return_ids=True returns (at the cost of one extra read) the objectID(s) this edit affected —
    for "replace" the paragraph you edited, for append/insert the newly created paragraph(s) — plus
    the page's new last_modified_time, instead of a bare status; useful when a follow-up edit needs
    them. (To rewrite the SAME text in many spots, prefer find_and_replace; for several different
    edits to one page in one atomic write, prefer batch_update.)"""
    result = page_edit.edit_page_content(
        get_backend(),
        page_id,
        content,
        mode,
        target_object_id=target_object_id,
        force=force,
        return_ids=return_ids,
    )
    if not return_ids:
        return f"updated {page_id}"
    out: dict[str, object] = {"page_id": page_id, "mode": mode}
    if mode == "replace" and target_object_id:
        out["affected_object_ids"] = [target_object_id]
    if result:
        out["new_object_ids"] = result.get("new_object_ids", [])
        out["last_modified_time"] = result.get("last_modified_time")
    return _json(out)


@logged_tool()
def find_and_replace(
    page_id: str, find: str, replace: str, object_id: str = "", force: bool = False
) -> str:
    """Fix or change text occurrences IN PLACE without re-supplying the paragraph — replace every
    occurrence of `find` with `replace` across the whole page — title, body, and table cells — or
    within one object when object_id is given, in ONE guarded write; each text run keeps its own
    style (scope it with object_id when you mean only one paragraph). This is the right tool for a
    TYPO or a small wording change: it avoids re-emitting a whole paragraph (smaller, and it cannot
    re-introduce other character errors in text you did not mean to touch). Returns how many
    occurrences were replaced and which objectIDs changed.

    Limitation: matching is per text run. If the only occurrences straddle a run boundary (the
    matched text is split across differently-styled spans), they are NOT replaced and their
    objectIDs come back under "found_across_runs" — read one with get_object and rewrite it with
    update_page_content "replace". Distinct from update_page_content "replace" (rewrites a whole
    paragraph's text) and apply_text_style (changes style, never the words). Concurrency-guarded;
    force=True only after explicit user confirmation."""
    return _json(
        page_edit.find_and_replace(
            get_backend(), page_id, find, replace, object_id=object_id, force=force
        )
    )


@logged_tool()
def batch_update(
    page_id: str, operations: list[dict], force: bool = False, return_ids: bool = False
) -> str:
    """Apply SEVERAL text edits to one page in a single ATOMIC write — all the operations succeed
    together or, if any is invalid, none is written (and it is one round-trip, not many). Use this
    when you have multiple edits to the same page — several typo fixes, rewriting a few paragraphs,
    appending in more than one place — instead of a burst of update_page_content calls.

    operations: a list of dicts, each with an "op":
      "replace" / "append" / "insert_before" / "insert_after" — as in update_page_content:
        "content" (text or paragraph dicts) and "target_object_id".
      "find_replace" — "find" / "replace" (+ optional "object_id" to scope it to one object).
    Each operation targets objectIDs that ALREADY exist on the page (an object created by an earlier
    operation in the same batch has no id until the write completes — split such work across two
    calls). Returns a per-operation summary; set return_ids=True to also get the objectIDs created
    and the new last_modified_time (one extra read). Concurrency-guarded; force=True only after
    explicit user confirmation."""
    return _json(
        page_edit.batch_update(
            get_backend(), page_id, operations, force=force, return_ids=return_ids
        )
    )


@logged_tool()
def create_table(
    page_id: str,
    rows: list[list[str | dict]],
    borders_visible: bool = True,
    has_header_row: bool = False,
    target_object_id: str = "",
    force: bool = False,
) -> str:
    """Create a NEW table on a page. Use this ONLY to make a brand-new table — NOT to change a
    table that already exists. On an EXISTING table: change its shape (add/insert rows, add columns,
    delete rows/columns) or overwrite whole rows of content with modify_table (set_rows); edit the
    TEXT of ONE cell with update_page_content ("replace" on that cell's paragraph objectID).

    rows: cells are plain strings or dicts {"text" | "runs", "style", "shading_color",
    "alignment"} (short rows are padded). target_object_id: empty → new table at the end of the
    page's last outline; an outline objectID → new table in that outline. (Passing an existing
    table's objectID is an error — use modify_table.) Concurrency-guarded; force=True only after
    explicit user confirmation."""
    page_edit.add_table(
        get_backend(),
        page_id,
        rows,
        borders_visible=borders_visible,
        has_header_row=has_header_row,
        target_object_id=target_object_id,
        force=force,
    )
    return f"table added to {page_id}"


@logged_tool()
def modify_table(
    page_id: str,
    table_object_id: str,
    operation: Literal[
        "insert_columns",
        "insert_rows",
        "set_rows",
        "set_column",
        "reorder_columns",
        "reorder_rows",
        "delete_columns",
        "delete_rows",
    ],
    rows: list[list[str | dict | None]] | None = None,
    indices: list[int] | None = None,
    order: list[int] | None = None,
    values: list[str | dict | None] | None = None,
    at_index: int | None = None,
    count: int = 1,
    width: float | None = None,
    force: bool = False,
) -> str:
    """Change an EXISTING table in place — its SHAPE (row/column count or ORDER) or its CONTENT (a
    whole row, a whole column, or every row) — keeping the table's objectID and every untouched
    cell's identity. Pair with create_table (which only makes NEW tables) and update_page_content
    ("replace" to edit ONE cell's TEXT). Get table_object_id and the row/column layout from get_page
    (or get_table for just that one table) first.

    operation (row and column edits are symmetric):
      "insert_rows"    — insert rows at 0-based at_index (omit at_index → append at the end).
                         rows = cell content, same shape as create_table.
      "insert_columns" — insert count empty columns at 0-based at_index (omit → append at the
                         end); width defaults to the last column's. Every row gains an empty cell.
                         values (only with count=1) fills the new column in the SAME call — a flat
                         list, one cell value per row; otherwise fill later with set_column.
      "set_rows"       — OVERWRITE the content of rows that ALREADY exist with rows (cell content,
                         same shape as create_table) from at_index (omit → row 0), one input row per
                         existing row. set_rows only rewrites existing rows — to ADD new rows use
                         insert_rows. Fixed-shape: no row/column added or removed, every cell keeps
                         its ID; a short input row leaves trailing columns untouched, and a None
                         cell leaves THAT cell unchanged ([None,"",""] keeps column 0, clears the
                         rest — the "keep first column/header, clear the body" pattern). Pass one
                         row + at_index to overwrite a single row; pass every row to refresh the
                         whole table (vs update_page_content "replace", which rewrites ONE cell).
                         Writing past the last row, or a row wider than the table, is refused.
      "set_column"     — OVERWRITE the content of ONE column (at_index = the 0-based column) with
                         values, a flat list of one cell value per row from the top. The compact way
                         to rewrite a single column without re-supplying the whole table; same
                         fixed-shape/keep-identity rules as set_rows (short list leaves trailing
                         rows, None leaves that cell). To set a whole column's STYLE/COLOR (not its
                         text), use apply_text_style(columns=[j]).
      "reorder_columns"— reorder the columns into order: a COMPLETE permutation of the current
                         column indices ([2,0,1] puts column 2 first). The matching cell in every
                         row moves with its column; nothing is added, removed, or retyped. Use this
                         to rearrange columns instead of clearing+rewriting them with set_rows.
      "reorder_rows"   — reorder the rows into order: a COMPLETE permutation of the current row
                         indices. Each row keeps its objectID and content; only its position moves.
      "delete_rows"    — DESTRUCTIVE: remove the rows at indices (0-based list).
      "delete_columns" — DESTRUCTIVE: remove the columns at indices (0-based) plus the matching
                         cell in every row.
    Deleting every row/column is refused — remove the whole table with delete_page_content.
    Concurrency-guarded; force=True only after explicit user confirmation. DESTRUCTIVE
    operations should be proposed and confirmed with the user first."""
    page_edit.modify_table(
        get_backend(),
        page_id,
        table_object_id,
        operation,
        rows=rows,
        indices=indices,
        order=order,
        values=values,
        at_index=at_index,
        count=count,
        width=width,
        force=force,
    )
    return f"table {table_object_id} modified ({operation}) on {page_id}"


@logged_tool()
def insert_svg_image(
    page_id: str,
    svg: str,
    width: float | None = None,
    height: float | None = None,
    mode: Literal["append", "insert_before", "insert_after"] = "append",
    target_object_id: str = "",
    force: bool = False,
) -> str:
    """Insert a VECTOR graphic into a page from SVG markup — the server renders the SVG to an
    image and places it on the page. It takes SVG markup you generate directly — NOT a raster
    image, NOT a photo, NOT base64. Use it for diagrams, maps, charts, simple banners — anything
    expressible as vectors. For a RASTER image (a photo, or a PNG/JPG/GIF that exists as a FILE on
    the machine running the server) use insert_image_from_path instead; a raster image that exists
    ONLY in your context with no file on disk must be added BY HAND in the OneNote app (tell the
    user).

    Placement (like update_page_content):
      "append"        — (default) at the END of an outline; target_object_id optionally names an
                        outline objectID (default = the page's last outline).
      "insert_before" — right before target_object_id (a PARAGRAPH objectID from get_page), so the
                        picture lands MID-page instead of at the end.
      "insert_after"  — same, right after the target paragraph.

    svg: a complete <svg>…</svg> document. For Chinese/CJK text, set an explicit Windows
    font-family such as "Microsoft JhengHei" (微軟正黑體) — NOT the generic "sans-serif", which
    renders with the wrong font. Do NOT embed a raster image inside the SVG (an <image> with a
    data: URI is rejected — that just smuggles a photo back in and is slow). width/height (points)
    override the rendered size; omit to use the SVG's own size. Concurrency-guarded; force=True
    only after explicit user confirmation."""
    page_edit.insert_svg_image(
        get_backend(),
        page_id,
        svg,
        width=width,
        height=height,
        mode=mode,
        target_object_id=target_object_id,
        force=force,
    )
    return f"image inserted into {page_id}"


@logged_tool()
def insert_image_from_path(
    page_id: str,
    path: str,
    width: float | None = None,
    height: float | None = None,
    mode: Literal["append", "insert_before", "insert_after"] = "append",
    target_object_id: str = "",
    force: bool = False,
) -> str:
    """Insert a raster image (PNG/JPEG/GIF) into a page FROM A LOCAL FILE PATH. The server reads the
    file's bytes off disk, so they never pass through the model — this is how to add a real raster
    image efficiently (the old base64 insert was removed because emitting the bytes through the
    model was unusably slow). `path` is a file on the machine running the server: a photo the user
    pointed you at, or — in a client that can write files / run code (e.g. an agentic IDE) — an
    image you produced on disk (a chart rendered by code, a downloaded picture). Do NOT use it to
    smuggle bytes you only hold in context; if there is no file on disk, there is nothing to insert.
    For a VECTOR graphic (a diagram/chart you can express as markup) use insert_svg_image instead.

    Placement mirrors insert_svg_image: "append" (default) at an outline's end (target_object_id =
    an outline objectID, else the page's last outline); "insert_before" / "insert_after" relative to
    a paragraph objectID. width/height (points) override the image's size; omit to use its own. Only
    PNG/JPEG/GIF files are accepted (anything else is rejected). Concurrency-guarded; force=True
    only after explicit user confirmation."""
    page_edit.insert_image_from_path(
        get_backend(),
        page_id,
        path,
        width=width,
        height=height,
        mode=mode,
        target_object_id=target_object_id,
        force=force,
    )
    return f"image inserted into {page_id} from {path}"


@logged_tool()
def apply_text_style(
    page_id: str,
    font_family: str = "",
    size: float | None = None,
    color: str = "",
    highlight: str = "",
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    cell_shading: str = "",
    columns: list[int] | None = None,
    scope_object_id: str = "",
    force: bool = False,
) -> str:
    """Batch-change text styling — FONT, SIZE, COLOR, HIGHLIGHT, BOLD, ITALIC, UNDERLINE,
    STRIKETHROUGH — and/or TABLE-CELL background (cell_shading) across a page (or one part of it) in
    a single pass, leaving everything else intact: the emphasis/colors/sizes you did NOT change,
    hyperlinks, images and tables all survive. It does NOT change the text, structure, or which
    paragraphs exist. Use it for "make the whole page 微軟正黑體", "every heading 16pt", "the body
    blue", "make these pages bold + italic", "highlight pages 1-5 yellow", "clear the yellow
    backgrounds".

    Contrast: update_page_content("replace") rewrites ONE paragraph's text+style (and must
    re-supply its runs); modify_table(set_rows) overwrites whole CELLS. apply_text_style touches no
    content — only styling — so prefer it for restyling that should preserve the words. (To restyle
    a page AND its subpages, enumerate them with list_pages, then call this per page.)

    font_family: e.g. "微軟正黑體" / "Microsoft JhengHei". size: points (e.g. 12). color, highlight,
    cell_shading are colors — a name like "yellow" or hex "#FFFF00". highlight is the TEXT
    screen-marker (behind the words); cell_shading is the whole-CELL background of a table cell —
    these are DIFFERENT things. Either set to "none" to remove it; if a yellow background won't go
    away and you're unsure which it is, clear BOTH (highlight="none", cell_shading="none").
    bold/italic/underline/strikethrough are tri-state: true ON, false OFF, omit = leave. At least
    one styling argument is required. scope_object_id: omit = the WHOLE page (every outline + table;
    the page title is left alone); or pass an outline / table / paragraph / table-CELL / table-ROW
    objectID from get_page to restyle only that subtree. TABLE GRANULARITY: a whole ROW = pass that
    row's id from get_page's table.row_object_ids as scope_object_id; a whole COLUMN has no id, so
    pass columns=[j,...] (0-indexed) — both the text restyle and cell_shading then apply only to
    those columns of every row in scope (give a table objectID as scope). A single cell = its cell
    objectID; the whole table = the table objectID. Whole-page also updates the page's style
    baseline so future typing matches. Concurrency-guarded; force=True only after user confirms."""
    summary = page_edit.apply_text_style(
        get_backend(),
        page_id,
        font_family=font_family or None,
        size=size,
        color=color or None,
        highlight=highlight or None,
        bold=bold,
        italic=italic,
        underline=underline,
        strikethrough=strikethrough,
        cell_shading=cell_shading or None,
        columns=columns or None,
        scope_object_id=scope_object_id,
        force=force,
    )
    return _json(summary)


# --- Copy (Phase 5: raw-XML faithful transfer) ------------------------------


@logged_tool()
def copy_page(page_id: str, target_section_id: str, after_page_id: str = "") -> str:
    """DUPLICATE a page into a section (formatting, tables, inline images, attachments,
    pageLevel all preserved); the original stays put. This is a copy, NOT a move — to relocate
    a page without duplicating it, use move_page. This copies ONE page only: to copy several pages
    at once use copy_pages, to copy a page TOGETHER WITH its subpages use copy_page_subtree, and to
    copy a whole section use copy_section. By DEFAULT the copy is placed right BELOW the
    source page (a same-section duplicate appears immediately after its original — what you
    usually want when no position is given). Pass after_page_id to place it after a DIFFERENT page
    instead (it must be in target_section_id). Copying to a DIFFERENT section, where the original
    isn't present, leaves the copy at the end of that section. Returns the new page's name + level
    + id (refer to it by NAME, not the id). If the
    source is not fully downloaded on this machine (OneDrive files-on-demand), some
    images/files/embedded objects cannot be copied and come out blank/empty — sync_warning
    summarizes how many, and file_notes lists each. ALWAYS surface a non-null sync_warning to the
    user and suggest they fully sync the source in OneNote, then copy again (blanks don't
    self-heal)."""
    backend = get_backend()
    result = copy.transfer_page(backend, page_id, target_section_id)
    # transfer_page lands the copy at the section END. Default placement is right below the SOURCE
    # page; an explicit after_page_id wins. If the anchor isn't in the target section (a
    # cross-section copy with no explicit anchor), there is no "below the original" — leave it last.
    anchor = after_page_id or page_id
    try:
        hierarchy_edit.reposition_page(
            backend, target_section_id, result.page_id, after_page_id=anchor
        )
    except NodeNotFoundError:
        if after_page_id:
            raise  # an explicitly named anchor that is not in the section is a real error
    return _json(
        {
            "page_id": result.page_id,
            "name": result.name,  # report the copy by NAME, not the raw id
            "page_level": result.page_level,
            "sync_warning": copy.sync_warning(
                result.missing_images, result.missing_files, result.missing_objects
            ),
            "file_notes": result.file_notes,
        }
    )


@logged_tool()
def copy_section(section_id: str, target_parent_id: str) -> str:
    """DUPLICATE a whole section (all pages in order, subpage levels kept) into a notebook OR
    section group; the original stays put. The copy keeps the source name, de-collided with
    " (2)" if taken. For copying only SOME pages of a section, not the whole thing, use copy_pages
    (an explicit list) or copy_page_subtree (a page and its subpages). This is the largest copy
    unit (there is no copy_notebook — clone a whole
    notebook by copy_section per section into a manually-created notebook). Returns the new
    section's ID. If the source is not fully downloaded on this machine (OneDrive files-on-demand),
    images/files/embedded objects come out blank — sync_warning summarizes how many across all
    pages, file_notes details each (prefixed by page). ALWAYS surface a non-null sync_warning and
    tell the user to fully sync the source section in OneNote, then copy again (blanks do NOT
    self-heal)."""
    result = copy.transfer_section(get_backend(), section_id, target_parent_id)
    return _json(
        {
            "section_id": result.section_id,
            "sync_warning": copy.sync_warning(
                result.missing_images, result.missing_files, result.missing_objects
            ),
            "file_notes": result.file_notes,
        }
    )


@logged_tool()
def copy_pages(page_ids: list[str], target_section_id: str, after_page_id: str = "") -> str:
    """DUPLICATE SEVERAL pages into a section as ONE contiguous block, in the given order
    (each page's formatting, tables, images, attachments and subpage level preserved); the
    originals stay put. Use this when copying MORE THAN ONE page at once — copy_page is for a
    single page, copy_section is for a whole section, and to copy a page TOGETHER WITH its
    subpages use copy_page_subtree (it works out the subpage list for you). page_ids = the pages
    to copy, in the order you want them to end up. By DEFAULT the block lands at the END of the
    target section; pass after_page_id to place the whole block right after that page instead
    (it must be in target_section_id). Do NOT call copy_page repeatedly to copy a group of pages
    — that scatters each copy below its own original; this places them together. Returns each new
    page as name + level + id (report them by NAME, never as a raw id list). If the source is not
    fully downloaded on this machine, some images/files/embedded
    objects cannot be copied — sync_warning summarizes how many, file_notes lists each; ALWAYS
    surface a non-null sync_warning and tell the user to fully sync the source, then copy again."""
    backend = get_backend()
    result = copy.copy_pages(backend, page_ids, target_section_id)
    if after_page_id:
        # an explicitly named anchor must be a page in the target section, else it's a real error
        hierarchy_edit.reposition_pages(
            backend, target_section_id, result.page_ids, after_page_id=after_page_id
        )
    # no anchor → leave the block at the section end (already contiguous, in order)
    return _json(
        {
            # pages carry name + level so you report results BY NAME, not as a raw id list
            "pages": result.pages,
            "sync_warning": copy.sync_warning(
                result.missing_images, result.missing_files, result.missing_objects
            ),
            "file_notes": result.file_notes,
        }
    )


@logged_tool()
def copy_page_subtree(
    section_id: str,
    page_id: str,
    target_section_id: str = "",
    after_page_id: str = "",
) -> str:
    """DUPLICATE a page TOGETHER WITH its subpages (the more-indented pages that follow it) as one
    contiguous block — "copy this page and everything under it". section_id is where the source
    page lives; page_id is the page to copy (its subpages are found automatically).
    target_section_id is where the copies go (omit it to copy within the SAME section, the common
    case). By DEFAULT, a same-section copy lands right BELOW the source subtree (the copy appears
    just after the original, like copy_page); pass after_page_id to place the block right after a
    DIFFERENT page instead (this is how you do "copy ●ITIN and its subpages below the ●Local page":
    page_id=●ITIN, after_page_id=●Local). Copying to a different section, or with no same-section
    anchor, lands the block at that section's end. Distinct from copy_page (single page, no
    subpages), copy_pages (an explicit page list), and copy_section (the whole section). Returns
    each new page as name + level + id (report them by NAME, never as a raw id list); the same
    sync_warning / file_notes rules as copy_page apply — ALWAYS surface a non-null sync_warning."""
    backend = get_backend()
    target = target_section_id or section_id
    sub_ids = copy.subtree_page_ids(backend, section_id, page_id)
    result = copy.copy_pages(backend, sub_ids, target)
    # Default placement mirrors copy_page's "below the source": for a same-section copy, anchor on
    # the source subtree's LAST page so the copy block lands right after the original. An explicit
    # after_page_id wins. Cross-section (or an implicit anchor not in target) → leave at the end.
    anchor = after_page_id or (sub_ids[-1] if target == section_id else "")
    if anchor:
        try:
            hierarchy_edit.reposition_pages(backend, target, result.page_ids, after_page_id=anchor)
        except NodeNotFoundError:
            if after_page_id:
                raise  # an explicitly named anchor that is not in the section is a real error
    return _json(
        {
            # pages carry name + level so you report results BY NAME, not as a raw id list
            "pages": result.pages,
            "sync_warning": copy.sync_warning(
                result.missing_images, result.missing_files, result.missing_objects
            ),
            "file_notes": result.file_notes,
        }
    )


# NOTE: there is deliberately no copy_notebook tool — same ground truth as create_notebook
# (COM cannot create notebooks on this build). Whole-notebook cloning is done by copy_section
# into an existing notebook / section group, section by section.


# --- Restructure (Phase 4: whole-batch UpdateHierarchy — SPEC §5 discipline) -
# Structural changes are propose-then-confirm: suggest a clone backup (copy_section)
# first, and present the target order for user confirmation before applying.


@logged_tool()
def restructure_section(section_id: str, ordered_pages: list[dict]) -> str:
    """STRUCTURAL. Reorder MANY pages WITHIN ONE section and/or adjust their subpage levels, in one
    batch. Use this when reordering several pages at once, or to set pages' subpage level
    (pageLevel). To move just ONE page to a position, use reposition_page instead — it takes only
    the IDs, NOT the whole list. To move a page to a DIFFERENT section use move_page; to reorder the
    sections themselves use reorder_sections; to rename a page/section use rename_node.
    ordered_pages = the section's COMPLETE page list in target order, each entry {"page_id": str,
    "page_level": 1|2|3} (a partial list is rejected — build it by reordering the list_pages output
    in place; do NOT write it to an external scratch file). Propose the target order and confirm
    with the user first; suggest a copy_section backup."""
    hierarchy_edit.restructure_section(get_backend(), section_id, ordered_pages)
    return f"section {section_id} restructured"


@logged_tool()
def reposition_page(
    section_id: str, page_id: str, after_page_id: str = "", page_level: int | None = None
) -> str:
    """STRUCTURAL. Move ONE page to a new position WITHIN its section — placed right after
    after_page_id (leave empty to move it to the TOP of the section) — and optionally set its
    subpage level (page_level 1/2/3). You give only the IDs; you do NOT need the section's full page
    list (that's the difference from restructure_section, which is for reordering MANY pages or
    setting several levels at once). This is the tool for "put the copied/new page right below page
    X": copy_page / create_page append the page at the END of the section, then reposition_page
    moves it where you want. Same-section only — to move a page to a DIFFERENT section use
    move_page. Confirm with the user before applying."""
    hierarchy_edit.reposition_page(
        get_backend(), section_id, page_id, after_page_id=after_page_id, page_level=page_level
    )
    return f"page {page_id} repositioned in {section_id}"


@logged_tool()
def reorder_sections(notebook_id: str, ordered_section_ids: list[str]) -> str:
    """STRUCTURAL. Reorder the SECTIONS (and section groups) within a notebook or section group,
    in one batch. Use this for the order of sections themselves — NOT for pages inside a section
    (that's restructure_section) and NOT for renaming (rename_node). ordered_section_ids = the
    COMPLETE child list in target order, including BOTH sections and section groups exactly as
    list_sections shows them at that level (a partial list is rejected; the hidden recycle bin
    is handled automatically). Ordering the top-level notebook list itself is not supported.
    Propose the order and confirm with the user first; suggest a copy_section backup."""
    hierarchy_edit.reorder_sections(get_backend(), notebook_id, ordered_section_ids)
    return f"sections of {notebook_id} reordered"


@logged_tool()
def rename_node(parent_id: str, object_id: str, new_name: str) -> str:
    """STRUCTURAL. Change only the NAME of a page, section, or section group — nothing moves or
    reorders. Not for moving a page (move_page), reordering pages (restructure_section), or
    reordering sections (reorder_sections). parent_id = the containing section/notebook ID.
    (A page rename edits its title; the hierarchy name follows it.) Confirm with the user
    before applying."""
    hierarchy_edit.rename_node(get_backend(), parent_id, object_id, new_name)
    return f"{object_id} renamed to {new_name}"


@logged_tool()
def move_page(notebook_id: str, page_id: str, target_section_id: str) -> str:
    """STRUCTURAL. Move a page to a DIFFERENT section within the same notebook — the page leaves
    its current section (this is a move, not a copy_page duplicate) and lands at the end of the
    target section with its subpage level reset to 1. To move/position a page WITHIN its current
    section instead, use reposition_page (one page) or restructure_section (many pages). The moved
    page gets a NEW page ID — returned here; use it for any follow-up calls. Confirm with the user
    before applying."""
    new_id = hierarchy_edit.move_page(get_backend(), notebook_id, page_id, target_section_id)
    return _json({"new_page_id": new_id, "section_id": target_section_id})


# --- Delete (Phase 6: destructive — conservative) ---------------------------


@logged_tool()
def delete_node(object_id: str, permanent: bool = False) -> str:
    """DESTRUCTIVE. Delete a whole hierarchy NODE — a notebook, section group, section, or
    page — via DeleteHierarchy. This removes the entire node and everything under it; to
    remove an object from WITHIN a page (an outline, image, attachment) use
    delete_page_content instead. Defaults to the recycle bin (permanent=False, the undo net);
    permanent=True is unrecoverable. Confirm with the user before applying."""
    delete.delete_node(get_backend(), object_id, permanent=permanent)
    return f"deleted {object_id}" + (" permanently" if permanent else " to recycle bin")


@logged_tool()
def delete_page_content(page_id: str, object_id: str, force: bool = False) -> str:
    """DESTRUCTIVE. Delete ONE page-level content object from a page — a whole outline, a
    page-level image, or a page-level attachment/embedded object — via DeletePageContent.
    object_id comes from get_page / get_page_images / get_page_files_info. This removes a
    PAGE-LEVEL object only; it does NOT delete a whole page/section (use delete_node) and
    canNOT remove inline content (a single paragraph, or a table/image/attachment inside an
    outline) — use delete_inline_content for that, or to drop a table ROW/COLUMN use
    modify_table (delete_rows / delete_columns). Concurrency-guarded; force=True only after
    explicit user confirmation. Confirm with the user before applying."""
    delete.delete_page_content(get_backend(), page_id, object_id, force=force)
    return f"deleted content object {object_id} from {page_id}"


@logged_tool()
def delete_inline_content(page_id: str, object_id: str, force: bool = False) -> str:
    """DESTRUCTIVE. Delete ONE object from INSIDE an outline — a table, a single paragraph, or an
    inline image/attachment — by its objectID from get_page_info (or get_page). A whole table or a
    paragraph is ALWAYS inside an outline, so removing one ALWAYS uses THIS tool, never
    delete_page_content. This is the
    complement of delete_page_content: that tool removes PAGE-LEVEL objects (a whole outline, or a
    page-level — i.e. printout — image/attachment); this one removes content nested inside an
    outline, which DeletePageContent refuses. Sibling paragraphs in the same outline are kept, so
    "delete the table but keep the surrounding text" just works. objectID: pass the table's OWN
    objectID to drop a whole table; pass a paragraph's objectID to drop that paragraph. For an
    inline image/attachment, pass EXACTLY the object_id that get_page_info / get_page reports for it
    and it just works — that id is sometimes the image's own id and sometimes its enclosing
    paragraph's (OneNote puts the id on the OE when the image has none); both resolve correctly,
    and since an image/attachment OE holds nothing else, no sibling text is affected. Do NOT re-read
    with get_page to "verify" the id — get_page_info's id is directly usable here. To remove only
    SOME of a table's rows/columns use modify_table (delete_rows / delete_columns) instead of this;
    to delete a whole page or section use delete_node.
    Concurrency-guarded; force=True only after explicit user confirmation. Confirm with the user
    before applying."""
    page_edit.delete_inline_content(get_backend(), page_id, object_id, force=force)
    return f"deleted inline object {object_id} from {page_id}"


def main() -> None:
    """Console entry point. With no args: the stdio MCP server (logs to stderr, never stdout).
    With ``--configure``: register this server in Claude Desktop's config and exit (§8)."""
    import argparse

    parser = argparse.ArgumentParser(
        prog="onenote-com-mcp",
        description="COM-only OneNote MCP server. No args runs the stdio server.",
    )
    parser.add_argument(
        "--configure",
        action="store_true",
        help="register this server in every detected MCP client config "
        "(Claude Desktop regular + Store, Antigravity) and exit",
    )
    parser.add_argument(
        "--selftest",
        action="store_true",
        help="bind OneNote via COM and list notebooks, then exit (install health check)",
    )
    args = parser.parse_args()

    if args.configure:
        from onenote_com_mcp.configure import configure_mcp_clients, repair_onenote_typelib

        # Auto-repair a broken OneNote typelib registration (per-user HKCU shim, no admin) so a
        # cold OneNote launch doesn't fail with TYPE_E_LIBNOTREGISTERED. No-op when healthy.
        for note in repair_onenote_typelib():
            print(note)

        written = configure_mcp_clients()
        for path in written:
            print(f"configured: {path}")
        if not written:
            # No supported client installed — write NOTHING (a guessed config path = the tool is
            # installed but invisible). Not an error: the install itself still completes; the user
            # installs a client then re-runs --configure (SPEC §8).
            print(
                "No supported MCP client detected (Claude Desktop or Antigravity). "
                "Install one, then re-run configuration via the Start Menu shortcut "
                '"OneNoteMCP — 重新偵測並設定" (or run this exe with --configure).'
            )
            return
        print("Restart the client(s) (Claude Desktop / Antigravity) to load the OneNote server.")
        return

    if args.selftest:
        # Health check: exercise the real backend end to end (on a frozen build this also
        # proves win32com's gen_py cache regenerates and OneNote binds via COM). Not the
        # stdio server, so printing to stdout is fine here.
        configure_logging()
        try:
            backend = get_backend()
            notebooks = read.list_notebooks(backend)
        except Exception as exc:  # noqa: BLE001 — report any failure as a clean non-zero exit
            print(f"SELFTEST FAIL: {exc}", file=sys.stderr)
            # Surface the actual HRESULT + underlying com_error: "GetHierarchy failed" alone is
            # undiagnosable. On a frozen build a bind-OK-but-call-fails almost always means the
            # bundled gen_py (makepy) was generated for a DIFFERENT OneNote typelib version than
            # this machine's, so also report which OneNote type libraries are registered here.
            hresult = getattr(exc, "hresult", None)
            if hresult is not None:
                print(f"  HRESULT: {hresult} ({hresult & 0xFFFFFFFF:#010x})", file=sys.stderr)
            cause = exc.__cause__ or exc.__context__
            if cause is not None:
                print(f"  cause: {cause!r}", file=sys.stderr)
            try:
                from win32com.client import selecttlb  # noqa: PLC0415

                for t in selecttlb.EnumTlbs():
                    if "onenote" in t.desc.lower():
                        print(
                            f"  typelib: {t.desc!r} ver {t.major}.{t.minor} "
                            f"lcid={t.lcid} clsid={t.clsid}",
                            file=sys.stderr,
                        )
            except Exception as diag:  # noqa: BLE001 — diagnostics must never mask the real error
                print(f"  (typelib enumeration failed: {diag!r})", file=sys.stderr)
            # Known field trap (Phase 0b on the VM, again on Chris's PC 2026-06-12): a stale
            # version subkey under the OneNote libid with NO win32/win64 mapping (e.g. a
            # PIA-only "1.0" left by an interop installer) poisons LoadRegTypeLib in the COM
            # marshaling/server layer → TYPE_E_LIBNOTREGISTERED on the FIRST method call, even
            # though our client side never touches the registry (vendored makepy module). Name
            # the broken key and the exact fix instead of leaving an opaque HRESULT.
            try:
                import winreg  # noqa: PLC0415 — Windows-only stdlib; selftest runs on Windows

                libid = "{0EA692EE-BB50-4E3C-AEF0-356D91732725}"
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"TypeLib\{libid}") as root:
                    index = 0
                    while True:
                        try:
                            ver = winreg.EnumKey(root, index)
                        except OSError:
                            break
                        index += 1
                        resolvable = False
                        for plat in (r"0\win32", r"0\win64"):
                            try:
                                winreg.CloseKey(winreg.OpenKey(root, rf"{ver}\{plat}"))
                                resolvable = True
                            except OSError:
                                pass
                        if not resolvable:
                            print(
                                f"  BROKEN typelib subkey: HKCR\\TypeLib\\{libid}\\{ver} has "
                                "no win32/win64 mapping — this poisons LoadRegTypeLib and "
                                "causes TYPE_E_LIBNOTREGISTERED. Fix (admin cmd, after a "
                                "`reg export` backup):\n"
                                "    reg delete "
                                f'"HKLM\\SOFTWARE\\Classes\\TypeLib\\{libid}\\{ver}" /f',
                                file=sys.stderr,
                            )
            except Exception:  # noqa: BLE001, S110 — best-effort diagnostics only
                pass
            raise SystemExit(1) from exc
        via = getattr(backend, "_bind_method", None)
        suffix = f" (bound via: {via})" if via else ""
        print(f"SELFTEST OK: connected to OneNote, {len(notebooks)} notebook(s) visible{suffix}")
        return

    configure_logging()  # §7: reads ONENOTE_MCP_LOG_LEVEL/FILE; default OFF, never stdout
    print("onenote-mcp starting (stdio)", file=sys.stderr)
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
