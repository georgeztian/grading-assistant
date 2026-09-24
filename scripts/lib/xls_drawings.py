"""Heuristic detector for embedded drawings/images in a legacy .xls file.

xlrd has no API for embedded images — Excel stores them as BIFF8
MSODRAWING/MSODRAWINGGROUP records (Escher/OfficeArt binary blobs) packed
directly inside the compound file's "Workbook" stream, not as separate,
named OLE streams. There is no maintained pure-Python library that decodes
these, so this does a lightweight linear scan of the raw stream: BIFF
records are laid out back-to-back as [type: u16][length: u16][payload], so
walking that layout and counting record types 0x00EB (MSODRAWINGGROUP),
0x00EC (MSODRAWING), and 0x005D (OBJ, embedded-object marker) is enough to
answer "does this file contain a drawing/image at all" without needing to
actually decode the Escher payload.

This is a presence heuristic, not an exact image count — a single
MSODRAWING record's Escher container can hold multiple shapes, so
`drawing_records` undercounts shapes and is not directly comparable to
extract_xlsx.py's exact `image_count`.
"""
from __future__ import annotations

from pathlib import Path

import olefile

MSODRAWINGGROUP = 0x00EB
MSODRAWING = 0x00EC
OBJ = 0x005D


def detect_xls_drawings(path: Path) -> dict:
    try:
        ole = olefile.OleFileIO(str(path))
    except Exception as exc:
        return {"supported": False, "reason": f"not_an_ole_file: {exc}"}

    try:
        stream_name = next(
            (name for name in ("Workbook", "Book") if ole.exists(name)), None
        )
        if stream_name is None:
            return {"supported": False, "reason": "no_workbook_stream_found"}
        data = ole.openstream(stream_name).read()
    finally:
        ole.close()

    drawing_group_records = 0
    drawing_records = 0
    obj_records = 0
    pos = 0
    n = len(data)
    # Safety cap: a well-formed stream has far fewer records than its byte
    # length; this just bounds worst-case iteration on a corrupt/garbage file.
    max_iterations = n
    iterations = 0

    while pos + 4 <= n and iterations < max_iterations:
        iterations += 1
        rec_type = data[pos] | (data[pos + 1] << 8)
        rec_len = data[pos + 2] | (data[pos + 3] << 8)
        if rec_type == MSODRAWINGGROUP:
            drawing_group_records += 1
        elif rec_type == MSODRAWING:
            drawing_records += 1
        elif rec_type == OBJ:
            obj_records += 1
        pos += 4 + rec_len

    return {
        "supported": True,
        "drawing_group_records": drawing_group_records,
        "drawing_records": drawing_records,
        "obj_records": obj_records,
        # OBJ (0x005D) also covers non-image embedded objects (buttons,
        # comments, checkboxes) — included anyway since this heuristic is
        # meant to err toward false positives ("go check") over false
        # negatives, same as the PDF equation-corruption flag.
        "likely_has_images": (
            drawing_group_records > 0 or drawing_records > 0 or obj_records > 0
        ),
    }
