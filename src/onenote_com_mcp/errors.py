"""Error types shared across layers. Pure Python — testable on Linux."""

from __future__ import annotations

# OneNote-is-busy HRESULTs worth retrying with backoff. Stored as both the unsigned
# 0x… form and the signed 32-bit form pywin32 surfaces, so callers can match either.
RPC_E_SERVERCALL_RETRYLATER = 0x8001010A  # -2147417846 signed
RPC_E_CALL_REJECTED = 0x80010001  # -2147418111 signed

_RETRYABLE_HRESULTS = frozenset(
    {
        RPC_E_SERVERCALL_RETRYLATER,
        RPC_E_SERVERCALL_RETRYLATER - 0x1_0000_0000,
        RPC_E_CALL_REJECTED,
        RPC_E_CALL_REJECTED - 0x1_0000_0000,
    }
)


def is_retryable_hresult(hresult: int | None) -> bool:
    """True if a COM HRESULT means 'OneNote is busy, try again'."""
    return hresult is not None and hresult in _RETRYABLE_HRESULTS


# hrLastModifiedDateDidNotMatch — the dateExpectedLastModified concurrency guard tripping
# on UpdatePageContent / DeletePageContent / DeleteHierarchy. Not a failure: the page/node
# changed since it was read. Both unsigned and pywin32's signed form.
HR_LAST_MODIFIED_MISMATCH = 0x80042010  # -2147213296 signed

_CONCURRENCY_HRESULTS = frozenset(
    {
        HR_LAST_MODIFIED_MISMATCH,
        HR_LAST_MODIFIED_MISMATCH - 0x1_0000_0000,
    }
)


def is_concurrency_hresult(hresult: int | None) -> bool:
    """True if a COM HRESULT means 'the target changed since you read it' (SPEC §5)."""
    return hresult is not None and hresult in _CONCURRENCY_HRESULTS


class OneNoteError(Exception):
    """Base class for all OneNote MCP errors."""


class BackendUnavailableError(OneNoteError):
    """The COM backend could not be reached (OneNote not running / not Windows)."""


class OneNoteComError(OneNoteError):
    """A COM call failed. Carries the HRESULT when known."""

    def __init__(self, message: str, hresult: int | None = None) -> None:
        super().__init__(message)
        self.hresult = hresult


class ConcurrencyError(OneNoteError):
    """The page changed since it was read; the guarded write was refused.

    Raised instead of clobbering the user's edits. Resolve by re-reading and retrying,
    or (explicit opt-in only) forcing the write.
    """


class NodeNotFoundError(OneNoteError):
    """A requested notebook/section/page/object ID was not found."""


class NoCurrentWindowError(OneNoteError):
    """OneNote has no open window, so the current viewing context cannot be read.

    SPEC §5: surface this clearly instead of guessing the user's location.
    """
