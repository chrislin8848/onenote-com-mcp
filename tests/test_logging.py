"""Phase 6 Stage 2: diagnostic log (SPEC §7).

Tests assert via the rotating FILE handler (real output path) rather than caplog — the
onenote_com_mcp logger is non-propagating by design (it owns its output, never leaks to a
host root logger), so reading the file is the faithful check. Key contracts: default OFF
(level ERROR), never stdout, params/results redacted (no base64/note content), failures
recorded even with detailed logging off, and logging never throws or alters a tool's result.
"""

from __future__ import annotations

import inspect
import logging
import sys

import pytest

from onenote_com_mcp.errors import OneNoteComError
from onenote_com_mcp.logging_config import (
    RotatingFileHandler,
    configure_logging,
    log_tool_call,
)


@pytest.fixture
def log_file(tmp_path, monkeypatch):
    path = tmp_path / "onenote-mcp.log"
    monkeypatch.setenv("ONENOTE_MCP_LOG_FILE", str(path))
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    return path


def _flush(logger: logging.Logger) -> None:
    for handler in logger.handlers:
        handler.flush()


# --- configuration ----------------------------------------------------------------------


def test_default_level_is_error_and_no_file_without_path(monkeypatch):
    monkeypatch.delenv("ONENOTE_MCP_LOG_LEVEL", raising=False)
    monkeypatch.delenv("ONENOTE_MCP_LOG_FILE", raising=False)
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    logger = configure_logging()
    assert logger.level == logging.ERROR  # detailed logging OFF by default
    assert not any(isinstance(h, RotatingFileHandler) for h in logger.handlers)


def test_debug_level_from_env(monkeypatch, log_file):
    monkeypatch.setenv("ONENOTE_MCP_LOG_LEVEL", "DEBUG")
    logger = configure_logging()
    assert logger.level == logging.DEBUG
    assert any(isinstance(h, RotatingFileHandler) for h in logger.handlers)


def test_never_writes_stdout(monkeypatch, log_file):
    monkeypatch.setenv("ONENOTE_MCP_LOG_LEVEL", "DEBUG")
    logger = configure_logging()
    for handler in logger.handlers:
        assert getattr(handler, "stream", None) is not sys.stdout


def test_configure_is_idempotent(monkeypatch, log_file):
    monkeypatch.setenv("ONENOTE_MCP_LOG_LEVEL", "DEBUG")
    configure_logging()
    logger = configure_logging()
    # no handler pile-up across repeated configuration
    assert sum(isinstance(h, RotatingFileHandler) for h in logger.handlers) == 1
    assert sum(isinstance(h, logging.StreamHandler) for h in logger.handlers) == 2  # stderr+file


def test_bad_log_file_degrades_to_stderr(monkeypatch, tmp_path):
    # point the log file at a path whose parent can't be created (a file, not a dir)
    blocker = tmp_path / "blocker"
    blocker.write_text("x")
    monkeypatch.setenv("ONENOTE_MCP_LOG_LEVEL", "DEBUG")
    monkeypatch.setenv("ONENOTE_MCP_LOG_FILE", str(blocker / "sub" / "x.log"))
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    logger = configure_logging()  # must not raise
    assert not any(isinstance(h, RotatingFileHandler) for h in logger.handlers)
    assert any(isinstance(h, logging.StreamHandler) for h in logger.handlers)


# --- per-call logging --------------------------------------------------------------------


def test_debug_logs_call_and_result_redacted(monkeypatch, log_file):
    monkeypatch.setenv("ONENOTE_MCP_LOG_LEVEL", "DEBUG")
    logger = configure_logging()

    @log_tool_call
    def insert_image(page_id: str, image_base64: str) -> str:
        return "image inserted into " + page_id

    insert_image(page_id="{P}{1}{B0}", image_base64="A" * 5000)
    _flush(logger)
    text = log_file.read_text()

    assert "call insert_image(" in text
    assert "page_id='{P}{1}{B0}'" in text
    assert "A" * 200 not in text, "base64 must never be written in full"
    assert "<str len=5000>" in text  # redacted by length
    assert "insert_image -> 'image inserted into {P}{1}{B0}'" in text


def test_errors_recorded_even_with_detailed_logging_off(monkeypatch, log_file):
    monkeypatch.setenv("ONENOTE_MCP_LOG_LEVEL", "ERROR")  # detailed OFF
    logger = configure_logging()

    @log_tool_call
    def delete_page_content(page_id: str, object_id: str) -> str:
        raise OneNoteComError("DeletePageContent failed", hresult=-2147213296)

    with pytest.raises(OneNoteComError):
        delete_page_content(page_id="{P}{1}{B0}", object_id="{O}{1}{B0}")
    _flush(logger)
    text = log_file.read_text()

    assert "delete_page_content failed" in text
    assert "hresult=-2147213296" in text
    # the params line is DEBUG, so it must NOT appear when detailed logging is off
    assert "call delete_page_content(" not in text


def test_logging_never_throws_on_hostile_repr(monkeypatch, log_file):
    monkeypatch.setenv("ONENOTE_MCP_LOG_LEVEL", "DEBUG")
    configure_logging()

    class Hostile:
        def __repr__(self) -> str:
            raise RuntimeError("repr boom")

    @log_tool_call
    def tool(x) -> str:
        return "ok"

    # redaction of a hostile arg must not break the call or its result
    assert tool(x=Hostile()) == "ok"


def test_decorator_preserves_signature_and_annotations():
    @log_tool_call
    def get_page(page_id: str, force: bool = False) -> str:
        return ""

    sig = inspect.signature(get_page)
    assert list(sig.parameters) == ["page_id", "force"]
    assert get_page.__name__ == "get_page"
    # functools.wraps copies __annotations__ (PEP 563 strings under future-annotations)
    assert get_page.__annotations__["force"] == "bool"
