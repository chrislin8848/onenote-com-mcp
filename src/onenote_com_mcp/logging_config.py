"""Diagnostic logging (SPEC §7). Default OFF (level ERROR). NEVER writes stdout (JSON-RPC).

Switches, read at server startup (changing one needs a Claude Desktop restart — accepted, the
server is launched as a subprocess and reads env once):

  ``ONENOTE_MCP_LOG_LEVEL``  default ``ERROR`` (detailed per-call logging OFF; only failures are
                             recorded). Set ``DEBUG`` to log every tool call's params + result.
  ``ONENOTE_MCP_LOG_FILE``   override the rotating-file path; default
                             ``%LOCALAPPDATA%\\OneNoteMCP\\logs\\onenote-mcp.log``. When neither
                             the override nor ``LOCALAPPDATA`` is set (e.g. on the Linux host),
                             no file handler is added — stderr only.

Output goes to a rotating file and/or stderr — **never stdout** (§8: stdout is JSON-RPC). Claude
Desktop folds the server's stderr into its own ``mcp-server-<name>.log`` for cross-reference.

ROBUSTNESS (§7): logging must never throw or interrupt a tool. File setup degrades silently to
stderr; per-call redaction/formatting is wrapped so a hostile ``__repr__`` can't break a tool.

The ``--configure`` installer step deliberately does NOT write ``ONENOTE_MCP_LOG_LEVEL`` — the
default stays OFF; a user adds it to the server's ``env`` block when they need to debug.
"""

from __future__ import annotations

import contextlib
import functools
import inspect
import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any

_ROOT_LOGGER_NAME = "onenote_com_mcp"
_TOOL_LOGGER = logging.getLogger(f"{_ROOT_LOGGER_NAME}.tools")

# Redaction ceilings — keep note content and base64 blobs out of the log, bound every record.
_MAX_STR = 120
_MAX_ITEMS = 8

_LOG_FORMAT = "%(asctime)s %(levelname)s %(name)s: %(message)s"


def _default_log_path() -> Path | None:
    override = os.environ.get("ONENOTE_MCP_LOG_FILE")
    if override:
        return Path(override)
    base = os.environ.get("LOCALAPPDATA")
    return Path(base) / "OneNoteMCP" / "logs" / "onenote-mcp.log" if base else None


def configure_logging() -> logging.Logger:
    """Set up the ``onenote_com_mcp`` logger from the env. Idempotent (re-callable in tests)."""
    logger = logging.getLogger(_ROOT_LOGGER_NAME)
    for handler in list(logger.handlers):  # idempotent: drop our previous handlers
        logger.removeHandler(handler)

    level_name = os.environ.get("ONENOTE_MCP_LOG_LEVEL", "ERROR").upper()
    logger.setLevel(getattr(logging, level_name, logging.ERROR))
    logger.propagate = False  # we own this logger's output; never leak to a host root logger
    formatter = logging.Formatter(_LOG_FORMAT)

    stderr_handler = logging.StreamHandler(sys.stderr)  # NEVER sys.stdout (JSON-RPC)
    stderr_handler.setFormatter(formatter)
    logger.addHandler(stderr_handler)

    path = _default_log_path()
    if path is not None:
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            file_handler = RotatingFileHandler(
                path, maxBytes=2_000_000, backupCount=5, encoding="utf-8"
            )
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        except OSError as exc:  # write-fail → silent degrade to stderr (§7), never crash
            logger.warning("diagnostic log file unavailable (%s); using stderr only", exc)
    return logger


def _redact(value: Any) -> str:
    """A bounded, content-safe string for one value — long strings/blobs become a length tag."""
    try:
        if isinstance(value, str):
            return repr(value) if len(value) <= _MAX_STR else f"<str len={len(value)}>"
        if isinstance(value, (bytes, bytearray)):
            return f"<bytes len={len(value)}>"
        if isinstance(value, dict):
            shown = ", ".join(f"{k}: {_redact(v)}" for k, v in list(value.items())[:_MAX_ITEMS])
            return "{" + shown + ("…" if len(value) > _MAX_ITEMS else "") + "}"
        if isinstance(value, (list, tuple)):
            if len(value) > _MAX_ITEMS:
                return f"<{type(value).__name__} len={len(value)}>"
            return "[" + ", ".join(_redact(v) for v in value) + "]"
        text = repr(value)
        return text if len(text) <= _MAX_STR else f"<{type(value).__name__} len~{len(text)}>"
    except Exception:  # a hostile __repr__ must not break logging
        return f"<{type(value).__name__} unrenderable>"


def _hresult_suffix(exc: BaseException) -> str:
    hresult = getattr(exc, "hresult", None)
    return f" [hresult={hresult}]" if hresult is not None else ""


def _params(func: Any, args: tuple, kwargs: dict) -> str:
    try:
        bound = inspect.signature(func).bind_partial(*args, **kwargs)
        return ", ".join(f"{k}={_redact(v)}" for k, v in bound.arguments.items())
    except Exception:
        return ", ".join(f"{k}={_redact(v)}" for k, v in kwargs.items())


def log_tool_call(func):
    """Wrap an MCP tool so each call logs name + redacted params + result/COM error code.

    Detailed records are DEBUG (off by default); failures are ERROR (always recorded). Logging
    is fully guarded — it never raises and never changes the tool's result or exception.
    ``functools.wraps`` preserves the signature/annotations FastMCP introspects for the schema.
    """

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        with contextlib.suppress(Exception):  # logging must never break a tool
            _TOOL_LOGGER.debug("call %s(%s)", func.__name__, _params(func, args, kwargs))
        try:
            result = func(*args, **kwargs)
        except Exception as exc:
            with contextlib.suppress(Exception):
                _TOOL_LOGGER.error("%s failed: %s%s", func.__name__, exc, _hresult_suffix(exc))
            raise
        with contextlib.suppress(Exception):
            _TOOL_LOGGER.debug("%s -> %s", func.__name__, _redact(result))
        return result

    return wrapper
