"""Retryable-HRESULT classification (the OneNote-is-busy backoff trigger)."""

from __future__ import annotations

from onenote_com_mcp.errors import (
    RPC_E_CALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER,
    is_retryable_hresult,
)


def test_retryable_unsigned_forms():
    assert is_retryable_hresult(RPC_E_SERVERCALL_RETRYLATER)
    assert is_retryable_hresult(RPC_E_CALL_REJECTED)


def test_retryable_signed_forms():
    # pywin32 typically surfaces HRESULTs as signed 32-bit ints.
    assert is_retryable_hresult(RPC_E_SERVERCALL_RETRYLATER - 0x1_0000_0000)
    assert is_retryable_hresult(RPC_E_CALL_REJECTED - 0x1_0000_0000)


def test_non_retryable():
    assert not is_retryable_hresult(None)
    assert not is_retryable_hresult(0)
    assert not is_retryable_hresult(0x80004005)  # E_FAIL — do not retry


def test_modal_ui_hresult_both_forms():
    from onenote_com_mcp.errors import HR_APP_IN_MODAL_UI, is_modal_ui_hresult

    assert is_modal_ui_hresult(HR_APP_IN_MODAL_UI)
    assert is_modal_ui_hresult(-2147213264)  # the signed form pywin32 surfaced on the VM
    assert not is_modal_ui_hresult(None)
    assert not is_modal_ui_hresult(0x80042010)
