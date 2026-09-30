#!/usr/bin/env python3
"""Extract every tab of a .xlsx workbook — formulas, cached values, hidden
sheets/rows/columns, cell comments, and merged-cell ranges.

Implements the spreadsheet-extraction method in .claude/skills/grading-instructions/extraction-fallback.md: a workbook is
loaded twice (data_only=False for formulas, data_only=True for cached
values) since neither load alone gives both, and every sheet is walked
(including hidden ones) rather than assuming a single active tab.

Normally used through extract.py (cached, with images exported); run
standalone it just prints the record.

Usage:
    .venv/Scripts/python scripts/extract_xlsx.py <file.xlsx>
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

import openpyxl
from openpyxl.chartsheet import Chartsheet
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib.images import export_image, image_ext  # noqa: E402
from lib.ooxml_extras import xlsx_drawing_objects  # noqa: E402


def formula_text(raw) -> str | None:
    """The formula a cell holds, as text — openpyxl returns array (CSE /
    dynamic-array) and data-table formulas as objects, not strings."""
    if isinstance(raw, str) and raw.startswith("="):
        return raw
    if isinstance(raw, ArrayFormula):
        return f"{{{raw.text}}}" if raw.text else "{=array formula}"
    if isinstance(raw, DataTableFormula):
        return f"{{=TABLE({raw.r1 or ''},{raw.r2 or ''})}}"
    return None


def export_sheet_images(ws, sheet_no: int, image_dir: Path) -> list[dict]:
    out = []
    for n, im in enumerate(getattr(ws, "_images", [])):
        start = getattr(im.anchor, "_from", None)
        anchor = f"{get_column_letter(start.col + 1)}{start.row + 1}" if start is not None else "absolute"
        try:
            blob = im._data()
        except Exception:
            continue
        name = f"sheet{sheet_no}_img{n}{image_ext(getattr(im, 'format', None))}"
        out.append({"anchor": anchor, "path": export_image(blob, name, image_dir)})
    return out


def extract_xlsx(path: Path, image_dir: Path | None = None) -> dict:
    """`image_dir`: when given, every sheet's images are exported there and
    listed (with their anchor cell) in the sheet's `images`."""
    wb_formulas = openpyxl.load_workbook(str(path), data_only=False)
    wb_values = openpyxl.load_workbook(str(path), data_only=True)
    # Text boxes and charts live in drawing parts openpyxl doesn't expose.
    drawing_objects = xlsx_drawing_objects(path)

    sheets = {}
    cells_with_formula = 0
    cells_with_comment = 0
    hidden_sheet_count = 0

    for sheet_name in wb_formulas.sheetnames:
        ws_f = wb_formulas[sheet_name]
        ws_v = wb_values[sheet_name]
        is_hidden = ws_f.sheet_state != "visible"
        if is_hidden:
            hidden_sheet_count += 1
        if isinstance(ws_f, Chartsheet):
            # A chart on its own tab: no cells; its chart is read from the
            # drawing part (below) and listed as chartN.
            sheets[sheet_name] = {"hidden": is_hidden, "chartsheet": True, "hidden_rows": [],
                                  "hidden_columns": [], "merged_ranges": [], "image_count": 0,
                                  "cells": {}, **drawing_objects.get(sheet_name, {})}
            continue

        hidden_rows = [r for r, dim in ws_f.row_dimensions.items() if dim.hidden]
        hidden_cols = [c for c, dim in ws_f.column_dimensions.items() if dim.hidden]
        merged_ranges = [str(r) for r in ws_f.merged_cells.ranges]
        image_count = len(getattr(ws_f, "_images", []))  # only populated if Pillow is installed

        cells: dict[str, dict] = {}
        max_row = max(ws_f.max_row, ws_v.max_row)
        max_col = max(ws_f.max_column, ws_v.max_column)
        for row_f, row_v in zip(
            ws_f.iter_rows(min_row=1, max_row=max_row, max_col=max_col),
            ws_v.iter_rows(min_row=1, max_row=max_row, max_col=max_col),
        ):
            for cell_f, cell_v in zip(row_f, row_v):
                raw = cell_f.value
                # A text cell may itself start with "=" (a label such as
                # "=rate (nper,pmt,pv,fv)"): only a formula-typed cell is one.
                formula = formula_text(raw) if cell_f.data_type == "f" else None
                value = cell_v.value if formula is not None else raw
                comment = cell_f.comment.text if cell_f.comment else None

                if formula is None and value is None and comment is None:
                    continue

                if formula is not None:
                    cells_with_formula += 1
                if comment is not None:
                    cells_with_comment += 1

                entry = {}
                if value is not None:
                    entry["value"] = value if not hasattr(value, "isoformat") else value.isoformat()
                if formula is not None:
                    entry["formula"] = formula
                    if value is None:
                        entry["value_uncached"] = True
                if comment is not None:
                    entry["comment"] = comment
                cells[cell_f.coordinate] = entry

        sheets[sheet_name] = {
            "hidden": is_hidden,
            "hidden_rows": hidden_rows,
            "hidden_columns": hidden_cols,
            "merged_ranges": merged_ranges,
            "image_count": image_count,
            "cells": cells,
            # Answers typed into text boxes, and chart data, anchored at a cell.
            **drawing_objects.get(sheet_name, {}),
        }
        if image_dir is not None and image_count:
            sheets[sheet_name]["images"] = export_sheet_images(
                ws_f, wb_formulas.sheetnames.index(sheet_name), image_dir)

    total_images = sum(s["image_count"] for s in sheets.values())
    total_charts = sum(len(s.get("charts", [])) for s in sheets.values())
    total_boxes = sum(len(s.get("text_boxes", [])) for s in sheets.values())

    return {
        "type": "xlsx",
        "sheet_order": wb_formulas.sheetnames,
        "sheets": sheets,
        "summary": {
            "sheet_count": len(wb_formulas.sheetnames),
            "hidden_sheet_count": hidden_sheet_count,
            "cells_with_formula": cells_with_formula,
            "cells_with_comment": cells_with_comment,
            "image_count": total_images,
            "chart_count": total_charts,
            "text_box_count": total_boxes,
            "image_warning": (
                f"{total_images} embedded image(s) — not in the cell text (no OCR). Each "
                "image's exported file is listed, with its anchor cell, in its sheet's header "
                "in the view — Read it there."
            ) if total_images else None,
        },
    }


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) != 2 or not Path(sys.argv[1]).exists():
        print(json.dumps({"error": "usage: extract_xlsx.py <existing file.xlsx>"}))
        sys.exit(1)
    print(json.dumps(extract_xlsx(Path(sys.argv[1])), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
