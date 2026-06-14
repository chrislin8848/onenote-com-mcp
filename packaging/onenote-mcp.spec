# PyInstaller spec — OneNote MCP server (console stdio exe). BUILT ON WINDOWS (the VM):
# PyInstaller does not cross-compile. Run from the repo root:
#     uv sync --group packaging
#     uv run pyinstaller packaging/onenote-mcp.spec --noconfirm
# Output: dist/OneNoteMCP/OneNoteMCP.exe (+ its dependency folder — onedir, see below).
#
# onedir (not onefile): more reliable for pywin32 + a runtime-writable gen_py cache, and
# faster startup (no per-launch temp extraction). The Inno installer ships the whole folder.
#
# stdout MUST stay clean (JSON-RPC): console=True gives a console subsystem exe, but the
# server itself writes only MCP protocol to stdout; logs go to stderr/file (§7/§8).

import os

from PyInstaller.utils.hooks import collect_all, collect_submodules

# SPECPATH is injected by PyInstaller = this spec's directory. Resolve our entry + runtime hook
# against it so the build works regardless of the CWD pyinstaller is invoked from.
_HERE = SPECPATH  # noqa: F821 (PyInstaller global)
_ENTRY = os.path.join(_HERE, "onenote_mcp_entry.py")
_RTHOOK = os.path.join(_HERE, "rthook_win32com_gen_py.py")

hiddenimports = [
    # win32 COM core
    "pywintypes",
    "pythoncom",
    "win32api",
    "win32con",
    "win32timezone",  # pywintypes pulls this in for some time conversions
    # the makepy machinery so gencache.EnsureModule() can (re)generate the OneNote typelib
    # module at runtime — OneNote can't be late-bound (Phase 0b), so EnsureModule stays and
    # must be able to build the module; the runtime hook gives it a writable gen_py dir.
    "win32com",
    "win32com.client",
    "win32com.client.gencache",
    "win32com.client.makepy",
    "win32com.client.genpy",
    "win32com.client.build",
    "win32com.client.selecttlb",
    # the vendored fully-generated OneNote 15.0 makepy module — imported lazily inside
    # Win32ComBackend._typelib_module(), so name it explicitly to guarantee it's bundled.
    "onenote_com_mcp.backend._gen_onenote15",
]
# MCP/FastMCP and pydantic resolve a lot dynamically — collect their submodules wholesale.
hiddenimports += collect_submodules("mcp")
hiddenimports += collect_submodules("pydantic")
hiddenimports += collect_submodules("pydantic_core")

# resvg_py — the SVG rasterizer behind insert_svg_image (a PyO3 extension bundling the resvg
# native lib). collect_all pulls its binary + data so the frozen exe can import it (freeze gate
# validated on the VM 2026-06-14).
_resvg_datas, _resvg_binaries, _resvg_hidden = collect_all("resvg_py")
hiddenimports += _resvg_hidden

a = Analysis(
    [_ENTRY],
    pathex=[],
    binaries=_resvg_binaries,
    datas=_resvg_datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[_RTHOOK],
    excludes=["tkinter", "matplotlib", "numpy", "PySide6", "PyQt5"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="OneNoteMCP",
    console=True,  # stdio server
    disable_windowed_traceback=False,
    upx=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    upx=False,
    name="OneNoteMCP",
)
