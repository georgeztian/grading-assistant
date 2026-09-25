#!/usr/bin/env python3
"""Extract every tab of a legacy .xls workbook (cached values only).

openpyxl cannot open .xls at all, so this uses xlrd (pinned <2.0, since 2.0+
dropped .xls support — see grader.md "Spreadsheet Extraction"). xlrd only
exposes the last-cached value, never the formula text, so every record is
flagged `formula_available: false` rather than silently treating a computed
answer as a plain literal.

Usage:
    .venv/Scripts/python scripts/extract_xls.py <file.xls>
    .venv/Scripts/python scripts/extract_xls.py <file.xls> --cache
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

import xlrd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache  # noqa: E402
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
                if cell.ctype == xlrd.XL_CELL_EMPTY:
                    continue
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(value, book.datemode).isoformat()
                    except Exception:
                        pass
                coord = f"{xlrd.colname(c)}{r + 1}"
                cells[coord] = {"value": value, "formula_available": False}

        # Attach comments — including ones on otherwise-empty cells.
        for (r, c), text in notes.items():
            coord = f"{xlrd.colname(c)}{r + 1}"
            cells.setdefault(coord, {"formula_available": False})["comment"] = text
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
            "extraction. If a question depends on a figure/chart/pasted "
            "image, inspect it directly — e.g. convert the file to .xlsx "
            "with LibreOffice (`soffice --headless --convert-to xlsx "
            "--outdir .cache/converted <file.xls>`) and re-run "
            "extract_xlsx.py on the result, which can read and count images."
        )
    elif not drawings.get("supported"):
        warnings.append(
            f"Could not check for embedded images/drawings: {drawings.get('reason')}"
        )

    return {
        "type": "xls",
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
    parser = argparse.ArgumentParser(description=__doc__,
                                      formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("file", type=Path)
    parser.add_argument("--cache", action="store_true")
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    if not args.file.exists():
        print(json.dumps({"error": f"file not found: {args.file}"}))
        sys.exit(1)

    if args.cache and not args.force:
        cached = cache.load_cached(args.file)
        if cached is not None:
            print(json.dumps({
                "cache_hit": True,
                "cache_path": str(cache.cache_path_for(args.file)),
                "summary": cached.get("summary"),
            }, indent=2))
            return

    record = extract_xls(args.file)

    if args.cache:
        cache_file = cache.save_cache(args.file, record)
        print(json.dumps({
            "cache_hit": False,
            "cache_path": str(cache_file),
            "summary": record["summary"],
        }, indent=2))
    else:
        print(json.dumps(record, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
