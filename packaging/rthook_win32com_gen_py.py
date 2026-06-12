"""PyInstaller runtime hook: give win32com a WRITABLE gen_py cache in the frozen app.

Why this exists (the core freeze problem for THIS project): OneNote cannot be late-bound —
``Dispatch("OneNote.Application")`` yields a dynamic object whose ``GetIDsOfNames`` can't
resolve ``GetHierarchy`` (Phase 0b ground truth). So ``Win32ComBackend`` binds via
``gencache.EnsureModule(...)`` + coclass instantiation, which means gencache must be able to
GENERATE the OneNote type-library module at runtime and WRITE it somewhere. In a frozen build
``win32com.__gen_path__`` defaults to a directory inside the read-only bundle, so the write
fails. This hook (runs before any app code) re-points the gen_py cache at a per-user temp dir.

The OneNote typelib is registered on every target machine (OneNote desktop is a precondition),
and the makepy machinery is bundled (see the .spec hiddenimports), so EnsureModule regenerates
cleanly on first call. Fully guarded — a hook failure must never crash server startup; if the
cache truly can't be set up, the backend surfaces a clear BackendUnavailableError instead.
"""

import os
import tempfile

try:
    import win32com

    _gen_dir = os.path.join(tempfile.gettempdir(), "onenote_mcp_gen_py")
    os.makedirs(_gen_dir, exist_ok=True)
    # Must be set BEFORE win32com.gen_py / gencache cache their paths.
    win32com.__gen_path__ = _gen_dir
    import win32com.gen_py

    win32com.gen_py.__path__ = [_gen_dir]
except Exception:
    pass
