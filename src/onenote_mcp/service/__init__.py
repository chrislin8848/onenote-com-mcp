"""Service layer — flow orchestration + input validation (SPEC §3).

Sits between the MCP tools and the XML/backend layers. Pure-Python and testable on Linux
against ``FixtureBackend``. Notably this is where the single shared
``GetPageContent → mutate tree in place → UpdatePageContent`` core lives (SPEC §4): the write
tools (update_page_content / create_table / insert_image) are thin facades over it. Copy
tools use the separate raw-XML transfer core. Implemented in Phases 2/4/5.
"""
