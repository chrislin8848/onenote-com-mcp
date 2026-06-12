"""Attachment / embedded-object read tools (SPEC v0612-2 §5, Phase 5b).

``get_page_files_info`` lists EVERY ``one:InsertedFile`` (any type) — metadata only, never
content. ``get_page_files`` extracts content for exactly three classes: text-like files
(decoded), images (base64 for MCP image content), and PDFs (server-side text via pypdf);
everything else (docx/xlsx/pptx/…) reports metadata + an explicit "unsupported".

Binary path (SPEC §5): attachment content is the ``pathCache`` file on the OneNote machine's
disk, NOT ``GetBinaryPageContent`` (that is the image callback path). The cache may be missing
(unsynced / purged) — every entry then reports ``status: "cache_unavailable"`` instead of
crashing. File access lives behind the backend so this module stays Linux-testable.

The deletable ``object_id`` mirrors the image rule: an INLINE InsertedFile has no objectID of
its own — the enclosing ``one:OE`` carries it; a PAGE-LEVEL one (Position/Size, direct page
child) carries its own.
"""

from __future__ import annotations

import base64
import io
from typing import Any

from onenote_com_mcp.backend.base import OneNoteBackend
from onenote_com_mcp.enums import PageInfo
from onenote_com_mcp.xmllayer.models import InsertedFile
from onenote_com_mcp.xmllayer.parse import parse_page

# Content classes get_page_files extracts (SPEC §5 — deliberately ONLY these three).
_TEXT_EXTENSIONS = frozenset(
    {
        "txt", "csv", "tsv", "json", "md", "markdown", "log",
        "xml", "yaml", "yml", "toml", "ini", "html", "htm", "css",
        "py", "js", "ts", "sh", "bat", "ps1", "sql",
    }
)  # fmt: skip
_IMAGE_EXTENSIONS = frozenset({"png", "jpg", "jpeg", "gif", "webp"})

# Hard ceiling before reading content at all; text/PDF output is additionally truncated to
# the caller's max_chars. Images pass through whole (they are bounded by this cap).
_MAX_FILE_BYTES = 20 * 1024 * 1024
DEFAULT_MAX_CHARS = 50_000

_IMAGE_MEDIA_TYPES = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "webp": "image/webp",
}


def _extension(preferred_name: str | None) -> str | None:
    if not preferred_name or "." not in preferred_name:
        return None
    return preferred_name.rsplit(".", 1)[1].strip().lower() or None


def _media_class(extension: str | None) -> str:
    if extension in _TEXT_EXTENSIONS:
        return "text"
    if extension in _IMAGE_EXTENSIONS:
        return "image"
    if extension == "pdf":
        return "pdf"
    return "unsupported"


def _deletable_object_id(f: InsertedFile) -> str | None:
    if f.placement == "inline":
        parent = f.node.getparent()
        return parent.get("objectID") if parent is not None else None
    return f.object_id


def _info_dict(backend: OneNoteBackend, f: InsertedFile) -> dict[str, Any]:
    size = backend.stat_cache_file(f.path_cache) if f.path_cache else None
    extension = _extension(f.preferred_name)
    info: dict[str, Any] = {
        "object_id": _deletable_object_id(f),
        "placement": f.placement,
        "kind": f.kind,
        "preferred_name": f.preferred_name,
        "extension": extension,
        "media_class": _media_class(extension),
        "path_source": f.path_source,
        "size_bytes": size,  # from the cache file; None = cache unavailable
        "cache_available": size is not None,
    }
    if f.kind == "embedded_preview":
        info["preview_pages"] = f.preview_pages
    if f.last_modified_time:
        info["last_modified_time"] = f.last_modified_time
    return info


def get_page_files_info(backend: OneNoteBackend, page_id: str) -> list[dict[str, Any]]:
    """Metadata for ALL of a page's attachments / embedded objects. Never parses content."""
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    return [_info_dict(backend, f) for f in page.inserted_files]


def _decode_text(raw: bytes) -> tuple[str, str]:
    """Decode attachment bytes → (text, encoding). BOMs first, then cp950 (zh-TW Windows),
    last-resort UTF-8 with replacement so the tool never crashes on a weird file."""
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16"), "utf-16"
    for encoding in ("utf-8-sig", "cp950"):
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8 (lossy)"


def _extract_pdf_text(raw: bytes) -> tuple[str, int]:
    """PDF bytes → (text, page count). pypdf: pure Python, PyInstaller-safe (SPEC §5)."""
    from pypdf import PdfReader  # noqa: PLC0415 — only the PDF branch needs it

    reader = PdfReader(io.BytesIO(raw))
    text = "\n".join((p.extract_text() or "") for p in reader.pages)
    return text, len(reader.pages)


def _truncate(text: str, max_chars: int) -> tuple[str, bool]:
    return (text[:max_chars], True) if len(text) > max_chars else (text, False)


def _content_dict(backend: OneNoteBackend, f: InsertedFile, max_chars: int) -> dict[str, Any]:
    extension = _extension(f.preferred_name)
    media_class = _media_class(extension)
    entry: dict[str, Any] = {
        "object_id": _deletable_object_id(f),
        "kind": f.kind,
        "preferred_name": f.preferred_name,
        "media_class": media_class,
    }
    if media_class == "unsupported":
        entry["status"] = "unsupported"
        entry["note"] = (
            "content extraction supports only text-class, image, and PDF attachments; "
            "use get_page_files_info for metadata"
        )
        return entry

    size = backend.stat_cache_file(f.path_cache) if f.path_cache else None
    if size is None:
        entry["status"] = "cache_unavailable"
        entry["note"] = (
            "the OneNote cache file for this attachment is not on disk (not yet synced, or "
            "purged) — content cannot be read right now"
        )
        return entry
    if size > _MAX_FILE_BYTES:
        entry["status"] = "too_large"
        entry["size_bytes"] = size
        return entry

    raw = backend.read_cache_file(f.path_cache)
    if raw is None:  # vanished between stat and read — same graceful answer
        entry["status"] = "cache_unavailable"
        return entry
    entry["size_bytes"] = len(raw)

    if media_class == "text":
        text, encoding = _decode_text(raw)
        entry["text"], entry["truncated"] = _truncate(text, max_chars)
        entry["encoding"] = encoding
        entry["status"] = "ok"
    elif media_class == "image":
        entry["media_type"] = _IMAGE_MEDIA_TYPES[extension]
        entry["data_base64"] = base64.b64encode(raw).decode("ascii")
        entry["status"] = "ok"
    else:  # pdf
        try:
            text, pages = _extract_pdf_text(raw)
        except Exception as exc:  # noqa: BLE001 — a malformed PDF must not crash the tool
            entry["status"] = "extract_failed"
            entry["note"] = f"PDF text extraction failed: {exc}"
            return entry
        entry["text"], entry["truncated"] = _truncate(text, max_chars)
        entry["pdf_pages"] = pages
        entry["status"] = "ok"
    return entry


def get_page_files(
    backend: OneNoteBackend,
    page_id: str,
    object_id: str = "",
    max_chars: int = DEFAULT_MAX_CHARS,
) -> list[dict[str, Any]]:
    """Content of a page's attachments (text / image / PDF only — SPEC §5).

    ``object_id`` narrows to one attachment (the ID ``get_page_files_info`` reported);
    empty = all. Text and PDF text are truncated to ``max_chars`` (flagged ``truncated``).
    """
    page = parse_page(backend.get_page_content(page_id, PageInfo.piBasic))
    files = page.inserted_files
    if object_id:
        files = [f for f in files if _deletable_object_id(f) == object_id]
    return [_content_dict(backend, f, max_chars) for f in files]
