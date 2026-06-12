"""Generate the NEUTRAL InsertedFile cache-file fixtures (Phase 5b).

Unlike the XML/binary fixtures (VM dumps), attachment cache files are deliberately
hand-authored on the host: the read path is a plain disk read of ``pathCache``, so synthetic
neutral bytes are faithful — and the PII policy forbids committing real user files anyway.
Filenames follow the FixtureBackend convention: ``cachefile_<sanitized {GUID}.bin>.bin``,
keyed to the pathCache attributes in the two committed attachment pages (附件與嵌入物件-1/-2).

The docx cache (GUID 1020C5EC-…) is DELIBERATELY not generated — its absence replays the
"cache unavailable" case.

Run from the repo root:  uv run python scripts/make_cache_fixtures.py
"""

from __future__ import annotations

from pathlib import Path

FIXTURES = Path(__file__).resolve().parent.parent / "tests" / "fixtures"


def minimal_pdf(text: str) -> bytes:
    """A small valid one-page PDF whose text pypdf can extract (ASCII, Helvetica)."""
    stream = f"BT /F1 12 Tf 72 720 Td ({text}) Tj ET".encode("ascii")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
        b"/Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for i, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref_at = len(out)
    out += f"xref\n0 {len(objects) + 1}\n".encode()
    out += b"0000000000 65535 f \n"
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += (
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_at}\n%%EOF\n"
    ).encode()
    return bytes(out)


# The canonical smallest valid baseline JPEG (1x1, gray). Extension drives media typing in
# service/files.py; valid bytes keep the fixture honest for visual consumers.
TINY_JPEG = bytes.fromhex(
    "ffd8ffe000104a46494600010101004800480000ffdb004300"
    + "ff" * 64
    + "ffc2000b080001000101011100ffc40014100100000000000000000000000000000000"
    + "ffda0008010100013f10"
    + "ffd9"
)

FAKE_OFFICE_ZIP = b"PK\x03\x04" + b"\x00" * 26 + b"fixture: fake Office container, never parsed\n"

TXT_CONTENT = (
    "濁水溪發電之旅(測試用中性內容)\n"
    "第一站:上游攔河堰,看引水道與沉砂池。\n"
    "第二站:電廠展示館,認識水輪機與年發電量。\n"
    "第三站:下游堤防步道,結束行程。\n"
    "備註:本檔為 fixture,內容為虛構。\n"
)

CACHE_FILES: dict[str, bytes] = {
    # 附件與嵌入物件-1: inline txt icon (濁水溪發電之旅.txt)
    "F5154E7A_7DFD_40B1_9FB8_8C73704C7B7D": TXT_CONTENT.encode("utf-8"),
    # 附件與嵌入物件-1: inline pdf icon (丘山行問卷_中英對照.pdf)
    "C92D9F98_3937_4BC3_B249_8472FE6402D1": minimal_pdf("Sample questionnaire PDF fixture"),
    # 附件與嵌入物件-1: inline pdf printout (A4文宣-25.7.8月分享會.pdf)
    "52EBD52B_8E38_4E7A_8809_772E897B2150": minimal_pdf("Sample flyer PDF fixture"),
    # 附件與嵌入物件-1: page-level xlsx icon (2026 客人問卷_NEW.xlsx) — unsupported class
    "13FD92E1_C84F_463D_AB1F_3A1EAABB9821": FAKE_OFFICE_ZIP,
    # 附件與嵌入物件-2: inline jpg icon (捷斯山屋.jpg)
    "DBC4674A_857C_476E_8A5F_5BD55E5C5D3A": TINY_JPEG,
    # 附件與嵌入物件-2: inline embedded xlsx (Previews) — unsupported class
    "D5798D78_4196_4CD1_8476_7A011607D137": FAKE_OFFICE_ZIP,
    # 1020C5EC_2CA7_4461_B119_0222B5E0D6BC (page-level docx) deliberately ABSENT:
    # replays "cache unavailable"
}


def main() -> None:
    for guid, payload in CACHE_FILES.items():
        path = FIXTURES / f"cachefile_{guid}_bin.bin"
        path.write_bytes(payload)
        print(f"wrote {path.name} ({len(payload)} bytes)")


if __name__ == "__main__":
    main()
