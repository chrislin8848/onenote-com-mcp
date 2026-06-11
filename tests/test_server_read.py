"""Server-layer wiring for the Phase 2 read tools.

The service-layer logic is covered in test_service_read.py; this only proves the MCP facades
are wired (get_backend via ONENOTE_FIXTURES_DIR, JSON formatting, the Image content facade)
and that the decorated tools stay directly callable.
"""

from __future__ import annotations

import json

import pytest
from mcp.server.fastmcp import Image

NOTEBOOK_ID = "{C94E632E-9829-45FF-914E-5E4031B2439D}{1}{B0}"
IMAGE_PAGE_ID = (
    "{65FA3E6E-E6E0-4610-B605-EE3D348FAD13}{1}{E1953306013858222940101982353039053288030011}"
)


@pytest.fixture(autouse=True)
def _use_fixture_backend(fixtures_dir, monkeypatch):
    monkeypatch.setenv("ONENOTE_FIXTURES_DIR", str(fixtures_dir))


def test_list_sections_returns_json_string():
    from onenote_com_mcp import server

    nodes = json.loads(server.list_sections(NOTEBOOK_ID))
    assert [n["name"] for n in nodes] == ["Phase 0 測試用", "節群組 測試用"]


def test_get_current_context_returns_json_string():
    from onenote_com_mcp import server

    ctx = json.loads(server.get_current_context())
    assert ctx["notebook"]["name"] == "MCP Test"
    assert ctx["page"]["name"] == "測試頁面4"


def test_get_page_images_returns_mcp_image_content():
    from onenote_com_mcp import server

    images = server.get_page_images(IMAGE_PAGE_ID)
    assert len(images) == 1
    assert isinstance(images[0], Image)
