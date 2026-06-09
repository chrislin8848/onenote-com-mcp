"""Pytest configuration.

Tier split (SPEC §2, §6): Tier 1 runs everywhere (pure logic + fixtures). Tier 2 is marked
``@pytest.mark.windows`` and requires live OneNote COM — auto-skipped off Windows so the host
loop stays green.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

FIXTURES_DIR = Path(__file__).parent / "fixtures"


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if sys.platform == "win32":
        return
    skip_windows = pytest.mark.skip(reason="requires live OneNote COM on Windows (Tier 2)")
    for item in items:
        if "windows" in item.keywords:
            item.add_marker(skip_windows)


@pytest.fixture
def fixtures_dir() -> Path:
    return FIXTURES_DIR
