"""THROWAWAY freeze spike (insert_svg_image gate) — NOT shipped, NOT imported by the package.

Proves the chosen resvg wheel can rasterize a Traditional-Chinese-bearing SVG to a valid PNG.
Stage A: run on the Linux host (cheap, retires "does the library work + render at all").
Stage B: PyInstaller-freeze this same script on the VM and run the frozen exe (the real gate:
         Windows packaging + CJK font fidelity). The written PNG is for Chris to eyeball.

Run (host):  uv run --with resvg-py python packaging/spike_svg/spike_resvg.py
"""

from __future__ import annotations

import sys
from pathlib import Path

# A tiny "route map"-ish SVG with Traditional Chinese labels — the real risk is whether the
# rasterizer finds a CJK font and renders glyphs (not tofu boxes).
SVG = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="480" height="200" viewBox="0 0 480 200">
  <rect x="0" y="0" width="480" height="200" fill="#f5f5f5"/>
  <line x1="60" y1="100" x2="420" y2="100" stroke="#3366cc" stroke-width="4"/>
  <circle cx="60" cy="100" r="10" fill="#cc3333"/>
  <circle cx="240" cy="100" r="10" fill="#cc3333"/>
  <circle cx="420" cy="100" r="10" fill="#cc3333"/>
  <text x="40" y="140" font-family="sans-serif" font-size="20" fill="#222">河內</text>
  <text x="205" y="140" font-family="sans-serif" font-size="20" fill="#222">下龍灣</text>
  <text x="395" y="140" font-family="sans-serif" font-size="20" fill="#222">寧平</text>
  <text x="120" y="40" font-family="sans-serif" font-size="24" fill="#3366cc">越南五日行程路線圖</text>
</svg>
"""

PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def rasterize(svg: str) -> bytes:
    import resvg_py  # the in-process PyO3 binding (pip: resvg-py)

    # mirror service/svg.py: pin the default + sans-serif family to the CJK-clean Windows font so
    # the eyeball reflects production behavior (fixes the handwriting-font substitution we saw).
    out = resvg_py.svg_to_bytes(
        svg_string=svg,
        font_family="Microsoft JhengHei",
        sans_serif_family="Microsoft JhengHei",
    )
    return bytes(out)


def main() -> int:
    try:
        png = rasterize(SVG)
    except Exception as exc:  # noqa: BLE001 — spike: report any failure plainly
        print(f"SPIKE FAIL: rasterize raised {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

    if not png.startswith(PNG_MAGIC):
        print(f"SPIKE FAIL: output is not a PNG (first bytes: {png[:8]!r})", file=sys.stderr)
        return 1
    if len(png) < 500:
        print(f"SPIKE FAIL: PNG suspiciously small ({len(png)} bytes)", file=sys.stderr)
        return 1

    # write to CWD (robust under a frozen exe, where __file__ points inside the bundle)
    out_path = Path.cwd() / "spike_out.png"
    out_path.write_bytes(png)
    print(f"SPIKE OK: {len(png)} byte PNG -> {out_path}")
    print("  (eyeball the PNG: the Chinese labels must be real glyphs, not tofu boxes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
