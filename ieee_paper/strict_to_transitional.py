"""Convert an OOXML 'strict' (ISO/IEC 29500) .docx to the 'transitional' flavour
that python-docx can read. The IEEE template ships in strict mode with purl.oclc.org
namespaces; we rewrite those in-place on a copy.

We do a targeted string-level rewrite over every .xml / .rels file in the
archive. This is safe because we're only touching namespace URIs and a few
content-type strings, all of which are already namespaced and not user content.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

_NAMESPACE_MAP: list[tuple[str, str]] = [
    (
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ),
    (
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ),
    (
        "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    ),
    (
        "http://purl.oclc.org/ooxml/drawingml/main",
        "http://schemas.openxmlformats.org/drawingml/2006/main",
    ),
    (
        "http://purl.oclc.org/ooxml/officeDocument/math",
        "http://schemas.openxmlformats.org/officeDocument/2006/math",
    ),
    (
        "http://purl.oclc.org/ooxml/officeDocument/sharedTypes",
        "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    ),
    (
        "http://purl.oclc.org/ooxml/officeDocument/extendedProperties",
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    ),
    (
        "http://purl.oclc.org/ooxml/officeDocument/customProperties",
        "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    ),
    (
        "http://purl.oclc.org/ooxml/officeDocument/customXml",
        "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
    ),
    (
        "http://purl.oclc.org/ooxml/schemas/package/2006/relationships",
        "http://schemas.openxmlformats.org/package/2006/relationships",
    ),
    (
        "http://purl.oclc.org/ooxml/schemas/package/2006/metadata/core-properties",
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    ),
]

_CONTENT_TYPE_MAP: list[tuple[str, str]] = [
    (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    ),
    (
        "application/vnd.ms-word.document.macroEnabled.main+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
    ),
]

_CONFORMANCE_RE = re.compile(r'\s+w:conformance="strict"')


def convert(src: Path, dst: Path) -> Path:
    """Read strict-OOXML docx at `src`, write a transitional copy to `dst`."""
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dst, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith((".xml", ".rels")):
                try:
                    text = data.decode("utf-8")
                except UnicodeDecodeError:
                    pass
                else:
                    for strict, transitional in _NAMESPACE_MAP:
                        text = text.replace(strict, transitional)
                    for strict_ct, transitional_ct in _CONTENT_TYPE_MAP:
                        if strict_ct != transitional_ct:
                            text = text.replace(strict_ct, transitional_ct)
                    text = _CONFORMANCE_RE.sub("", text)
                    data = text.encode("utf-8")
            zout.writestr(item, data)
    return dst


def ensure_transitional(template: Path) -> Path:
    """Return a transitional-flavoured copy of `template`, caching next to the
    original as `<stem>.transitional.docx`. Regenerates if the original is newer."""
    cached = template.with_suffix(".transitional.docx")
    if cached.is_file() and cached.stat().st_mtime >= template.stat().st_mtime:
        return cached
    return convert(template, cached)
