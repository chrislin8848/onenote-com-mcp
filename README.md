# onenote-com-mcp

A **COM-only** Model Context Protocol server that gives Claude full CRUD over the live
**OneNote desktop** app on Windows — reading and writing notebooks, sections, pages, rich
text (formatting preserved), tables, and images.

No Azure app registration. No Microsoft Graph. No tokens. It drives the OneNote desktop COM
API in-process via pywin32, inside the signed-in user's own OneNote session, so it works for
any employee with synced OneNote and needs no admin setup.

> Why not Graph? Existing public OneNote MCP servers are almost all Graph-based and inherit a
> shared set of constraints: an Azure app + token, search/scale problems across large
> notebooks, and lossy HTML formatting. This project takes the COM route deliberately —
> full-fidelity round-trips, scoped queries, and zero cloud onboarding. See `SPEC` §1.2.

## Status — 2026-06-10

| Phase | What | State |
|---|---|---|
| **0a** | Host scaffold: backend interface, guarded import, FixtureBackend, server tool catalog, CI | ✅ done |
| 0b | Windows 11 VM + autologon + secondary-session runner + `remote_test` + COM smoke | ⛔ **blocked: VM not built yet** |
| 1 | XML parse/build (lxml) + format preservation, TDD against **real VM fixtures** | ⛔ blocked on 0b (fixtures) |
| 2 | Read tools on `FixtureBackend` (Linux green) | ⛔ blocked on 0b (fixtures) |
| 3 | `Win32ComBackend` live + `dump_fixtures.py` | ⛔ blocked on 0b |
| 4 | Write tools + concurrency guard + format-preservation regression | ⛔ blocked |
| 5 | Copy/transfer (raw-XML faithful copy) | ⛔ blocked |
| 6 | Delete + retry hardening + **PyInstaller/Inno installer** | ⛔ blocked |

**Next bottleneck: build the Windows VM (Phase 0b).** Per the 2026-06-10 decision, fixtures
come from real OneNote (not synthetic), so Phase 1 onward needs the VM up and
`scripts/dump_fixtures.py` run once.

## Develop (Linux host)

```bash
uv sync
uv run pytest                       # Tier 1; Windows COM tests auto-skip here
uv run ruff check . && uv run ruff format --check .
```

## Run (against fixtures, Linux)

```bash
ONENOTE_FIXTURES_DIR=tests/fixtures uv run python -m onenote_com_mcp
```

On Windows with OneNote installed, omit the env var to use live COM.

## Layout

```
src/onenote_com_mcp/    backend/ (OneNoteBackend + Fixture + Win32Com)  xmllayer/  service/  server.py
docs/               com-api-reference.md   onenote-xml-schema.md
scripts/            remote_test.sh   dump_fixtures.py        (VM bridge — Phase 0b/3)
tests/              Tier 1 unit tests + fixtures/
refs/               mhzarem reference clone (gitignored)
```

See `CLAUDE.md` for the architecture and the non-negotiable invariants.
