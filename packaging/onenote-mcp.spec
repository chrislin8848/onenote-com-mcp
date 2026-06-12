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

from PyInstaller.utils.hooks import collect_submodules

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
]
# MCP/FastMCP and pydantic resolve a lot dynamically — collect their submodules wholesale.
hiddenimports += collect_submodules("mcp")
hiddenimports += collect_submodules("pydantic")
hiddenimports += collect_submodules("pydantic_core")

a = Analysis(
    ["onenote_mcp_entry.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=["rthook_win32com_gen_py.py"],
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
