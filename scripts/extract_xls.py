#!/usr/bin/env python3
"""Extract every tab of a legacy .xls workbook (cached values only).

Fallback only: extract.py normally converts .xls to .xlsx (Microsoft Excel
or LibreOffice, see convert_doc.py) and reads it with extract_xlsx.py, which
keeps formulas and images. This values-only reader is used when no converter
is installed, and its record is marked `degraded` so it is replaced once one
is.

openpyxl cannot open .xls at all, so this uses xlrd (pinned <2.0, since 2.0+
dropped .xls support — see .claude/skills/grading-instructions/extraction-fallback.md). xlrd only
exposes the last-cached value, never the formula text, so every record is
flagged `formula_available: false` rather than silently treating a computed
answer as a plain literal.

Usage:
    .venv/Scripts/python scripts/extract_xls.py <file.xls>
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

import xlrd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib.xls_drawings import detect_xls_drawings  # noqa: E402

VISIBILITY = {0: "visible", 1: "hidden", 2: "very_hidden"}


def extract_xls(path: Path) -> dict:
    book = xlrd.open_workbook(str(path), formatting_info=True)

    sheets = {}
    hidden_sheet_count = 0
    cells_with_comment = 0

    for sheet in book.sheets():
        visibility = VISIBILITY.get(getattr(sheet, "visibility", 0), "visible")
        if visibility != "visible":
            hidden_sheet_count += 1

        hidden_rows = [
            r for r, info in getattr(sheet, "rowinfo_map", {}).items() if info.hidden
        ]
        hidden_cols = [
            c for c, info in getattr(sheet, "colinfo_map", {}).items() if info.hidden
        ]

        # Cell comments ("notes"); xlrd only parses them with formatting_info=True.
        notes = {
            rc: (getattr(note, "text", "") or "")
            for rc, note in getattr(sheet, "cell_note_map", {}).items()
        }

        cells = {}
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):  # BLANK = formatting only
                    continue
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(value, book.datemode).isoformat()
                    except Exception:
                        pass
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    value = bool(value)
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    value = xlrd.error_text_from_code.get(value, "#ERROR")
                cells[f"{xlrd.colname(c)}{r + 1}"] = {"value": value}

        # Attach comments — including ones on otherwise-empty cells.
        for (r, c), text in notes.items():
            cells.setdefault(f"{xlrd.colname(c)}{r + 1}", {})["comment"] = text
        cells_with_comment += len(notes)

        sheets[sheet.name] = {
            "hidden": visibility != "visible",
            "hidden_rows": hidden_rows,
            "hidden_columns": hidden_cols,
            "cells": cells,
        }

    drawings = detect_xls_drawings(path)
    warnings = [
        "xlrd exposes only cached values for legacy .xls — no formula "
        "text is available; do not treat a missing formula as an "
        "absent answer."
    ]
    likely_has_images = bool(drawings.get("likely_has_images"))
    if likely_has_images:
        warnings.append(
            "This file appears to contain embedded drawing(s)/object(s)/"
            f"image(s) ({drawings['drawing_records']} MSODRAWING, "
            f"{drawings['drawing_group_records']} MSODRAWINGGROUP, "
            f"{drawings['obj_records']} OBJ record(s) detected via a "
            "heuristic byte scan — not an exact image count, may include "
            "non-image objects like comments/buttons, and xlrd cannot "
            "extract or render any of it). Their content is NOT in this "
            "extraction and cannot be viewed without a converter — report "
            "it, since installing Microsoft Excel or LibreOffice makes the "
            "images (and formulas) available."
        )
    elif not drawings.get("supported"):
        warnings.append(
            f"Could not check for embedded images/drawings: {drawings.get('reason')}"
        )

    return {
        "type": "xls",
        "formula_available": False,
        "sheet_order": book.sheet_names(),
        "sheets": sheets,
        "drawings": drawings,
        "warnings": warnings,
        "summary": {
            "sheet_count": book.nsheets,
            "hidden_sheet_count": hidden_sheet_count,
            "cells_with_formula": 0,
            "cells_with_comment": cells_with_comment,
            "likely_has_images": likely_has_images,
        },
    }


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) != 2 or not Path(sys.argv[1]).exists():
        print(json.dumps({"error": "usage: extract_xls.py <existing file.xls>"}))
        sys.exit(1)
    print(json.dumps(extract_xls(Path(sys.argv[1])), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
