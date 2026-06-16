"""The process-wide COM lock (``server._serialize_com``) serializes tool execution.

OneNote's COM is single-threaded; FastMCP runs sync tools in a worker-thread pool, so a burst of
tool calls can arrive in parallel and contend on the live COM server (which then rejects the
losers with RPC_E_SERVERCALL_RETRYLATER, burning the backend's finite busy-retry budget). The
lock turns that contention into an orderly queue — at most one tool body runs at a time. These
host tests prove the serialization and signature-preservation without needing COM; the live
"burst no longer fails" effect is Tier-2 (VM).
"""

from __future__ import annotations

import contextlib
import inspect
import threading
import time

from onenote_com_mcp.server import _COM_LOCK, _serialize_com


def test_serialize_com_allows_only_one_body_at_a_time():
    active = 0
    max_active = 0
    counter_lock = threading.Lock()  # guards the observation counters, NOT the thing under test

    @_serialize_com
    def work():
        nonlocal active, max_active
        with counter_lock:
            active += 1
            max_active = max(max_active, active)
        time.sleep(0.02)  # widen the window in which an overlap could be observed
        with counter_lock:
            active -= 1
        return "ok"

    threads = [threading.Thread(target=work) for _ in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert max_active == 1  # never two tool bodies running concurrently


def test_serialize_com_preserves_signature_and_result():
    @_serialize_com
    def tool(page_id: str, count: int = 1) -> str:
        return f"{page_id}:{count}"

    assert tool("p1", count=3) == "p1:3"  # result passes through unchanged
    # functools.wraps keeps name + signature — FastMCP introspects these to build the tool schema
    assert tool.__name__ == "tool"
    assert list(inspect.signature(tool).parameters) == ["page_id", "count"]


def test_serialize_com_releases_lock_on_exception():
    @_serialize_com
    def boom():
        raise ValueError("boom")

    with contextlib.suppress(ValueError):
        boom()
    # the context manager must release the lock even when the body raises
    assert _COM_LOCK.acquire(blocking=False)
    _COM_LOCK.release()
