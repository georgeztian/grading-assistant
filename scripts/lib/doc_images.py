"""Heuristic detector for embedded images in a legacy .doc file.

The binary Word 97-2003 format (.doc) is, like .xls, an OLE2 compound file
— but there is no lightweight maintained pure-Python library that decodes
its picture/drawing records (the format is considerably more complex than
BIFF's). Instead of parsing that structure, this scans the relevant OLE
streams — "WordDocument", "Data", and anything inside an "ObjectPool"
storage (which holds OLE-embedded objects, commonly pictures) — for known
image file signatures (PNG, JPEG, GIF, placeable-WMF), and separately notes
whether an ObjectPool storage exists at all.

This is a presence heuristic, not an exact count or a content guarantee.
Whenever a `.doc`→`.docx` converter is available, `extract_docx.py`'s exact
`image_count` on the converted file is the authoritative signal; this
heuristic exists specifically so images aren't silently invisible when no
converter is installed (extraction fails before ever reaching extract_docx).
"""
from __future__ import annotations

from pathlib import Path

import olefile

# Only signatures distinctive enough to search for as raw substrings without
# meaningful false-positive risk in an arbitrary binary blob. BMP ("BM") and
# raw WMF/EMF headers are too short/common to include here.
IMAGE_SIGNATURES = {
    "png": b"\x89PNG\r\n\x1a\n",
    "jpeg": b"\xff\xd8\xff",
    "gif87a": b"GIF87a",
    "gif89a": b"GIF89a",
    "wmf_placeable": b"\xd7\xcd\xc6\x9a",
}


def _scan_for_signatures(data: bytes) -> set[str]:
    return {name for name, sig in IMAGE_SIGNATURES.items() if sig in data}


def detect_doc_images(path: Path) -> dict:
    try:
        ole = olefile.OleFileIO(str(path))
    except Exception as exc:
        return {"supported": False, "reason": f"not_an_ole_file: {exc}"}

    try:
        object_pool_present = ole.exists("ObjectPool")
        signatures_found: set[str] = set()

        for stream_name in ("WordDocument", "Data"):
            if ole.exists(stream_name):
                try:
                    signatures_found |= _scan_for_signatures(ole.openstream(stream_name).read())
                except Exception:
                    pass

        if object_pool_present:
            for entry in ole.listdir(streams=True, storages=False):
                if entry and entry[0] == "ObjectPool":
                    try:
                        signatures_found |= _scan_for_signatures(ole.openstream(entry).read())
                    except Exception:
                        continue
    finally:
        ole.close()

    return {
        "supported": True,
        "object_pool_present": object_pool_present,
        "image_signatures_found": sorted(signatures_found),
        "likely_has_images": object_pool_present or bool(signatures_found),
    }
