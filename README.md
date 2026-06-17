# onenote-com-mcp

**Author:** Chris Lin · **License:** [MIT](LICENSE)

A **COM-only** [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that
gives Claude full CRUD over the live **OneNote desktop** app on Windows — reading and writing
notebooks, sections, pages, rich text (formatting preserved), tables, images, and attachments.

No Azure app registration. No Microsoft Graph. No tokens. It drives the OneNote desktop COM
API in-process via pywin32, inside the signed-in user's own OneNote session, so it works for
any user with synced OneNote and needs no admin setup.

> **Why not Graph?** Existing public OneNote MCP servers are almost all Graph-based and inherit
> a shared set of constraints: an Azure app + token, search/scale limits on large notebooks, and
> lossy HTML formatting. This project takes the COM route deliberately — full-fidelity
> round-trips, scoped queries, and zero cloud onboarding. See [docs/SPEC.md](docs/SPEC.md) §1.2.

## Features

~30 MCP tools across the OneNote object model, all preserving rich-text formatting:

- **Read** — list notebooks / sections / pages, full-text search, lossless page read (text runs +
  effective style + structured tables + object IDs), a lightweight per-page object inventory,
  single-table read, image extraction, attachment / embedded-file metadata & content
  (text / image / PDF), and the current-window context.
- **Edit** — in-place content edit (append / insert / replace, anchored by object ID), create &
  modify tables (insert/delete/reorder rows and columns, set cells, cell shading), batch text
  restyle (font / size / color / bold / italic / underline / highlight — by page, table, row or
  column), and SVG-to-image insertion (the model emits vector markup; the server rasterizes it).
- **Structure** — create sections & pages, reposition / reorder / rename, restructure a section,
  and move pages — whole-batch, ID-conserving hierarchy edits.
- **Copy** — faithful raw-XML copy of a single page, several pages, a page together with its
  subpages, or a whole section (inline image binaries and attachments carried).
- **Delete** — recycle-bin-by-default deletes of hierarchy nodes, page-level objects, and inline
  content.

Concurrency-safe (optimistic `lastModifiedTime` checks; COM calls serialized at the boundary) and
diagnostics-ready (opt-in rotating log, never on stdout).

## Requirements

- Windows with the **64-bit OneNote desktop** app installed and signed in (OneDrive-synced).
- The server runs inside that user's OneNote session — no admin rights, no cloud setup.

The pure XML/service layers are testable on any platform; only the live COM backend needs Windows.

## Install (end users)

Download and run `OneNoteMCP-Setup_<version>.exe` (per-user, no admin). It ships the server as a
standalone executable and, on install, auto-registers itself with every detected MCP client —
**Claude Desktop** (regular + Microsoft Store) and **Antigravity** (CLI / IDE). Restart the client
to load it. Installed a client later? Run **“OneNoteMCP — 重新偵測並設定”** from the Start Menu to
register it without reinstalling.

Health check: `OneNoteMCP.exe --selftest` binds OneNote and lists your notebooks.

## Develop (Linux or Windows host)

The XML and service layers are pure and fully tested on Linux against recorded fixtures; the COM
backend is a guarded import that only loads on Windows.

```bash
uv sync
uv run pytest                       # Tier 1; Windows-only COM tests auto-skip here
uv run ruff check . && uv run ruff format --check .
```

Run the server against fixtures (no OneNote needed):

```bash
ONENOTE_FIXTURES_DIR=tests/fixtures uv run python -m onenote_com_mcp
```

On Windows with OneNote installed, omit the env var to use live COM.

## Layout

```
src/onenote_com_mcp/   backend/ (OneNoteBackend + Fixture + Win32Com)  xmllayer/  service/  server.py
docs/                  SPEC.md  com-api-reference.md  onenote-xml-schema.md
packaging/             PyInstaller spec + Inno Setup installer
tests/                 Tier-1 unit tests + fixtures/
```

Testing is two-tier: **Tier 1** (pure logic + fixtures, runs anywhere) and **Tier 2**
(`@pytest.mark.windows`, real COM round-trips on a Windows VM). See
[CLAUDE.md](CLAUDE.md) for the architecture and the non-negotiable invariants.

## License

[MIT](LICENSE) © 2026 Chris Lin
