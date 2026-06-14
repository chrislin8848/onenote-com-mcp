# THROWAWAY PyInstaller spec — freeze gate for resvg_py (insert_svg_image). NOT shipped.
# Built on the VM (PyInstaller does not cross-compile). From the repo root:
#     uv run --with resvg-py --with pyinstaller pyinstaller packaging/spike_svg/spike.spec \
#         --noconfirm --distpath packaging/spike_svg/dist --workpath packaging/spike_svg/build
# Then run dist/spike_resvg/spike_resvg.exe — it must exit 0 and write spike_out.png with
# real Chinese glyphs (the Windows + CJK-fidelity half of the gate).

import os

from PyInstaller.utils.hooks import collect_all

_HERE = SPECPATH  # noqa: F821 (PyInstaller-injected = this spec's dir)
_ENTRY = os.path.join(_HERE, "spike_resvg.py")

# resvg_py is a PyO3 extension bundling the resvg native lib — collect its binary + any data so
# the frozen exe can import it. This is exactly what we must prove freezes cleanly on Windows.
datas, binaries, hiddenimports = collect_all("resvg_py")

a = Analysis(
    [_ENTRY],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "numpy", "PySide6", "PyQt5"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="spike_resvg",
    console=True,
    upx=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    upx=False,
    name="spike_resvg",
)
