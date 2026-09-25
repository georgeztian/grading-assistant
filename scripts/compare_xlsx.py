#!/usr/bin/env python3
"""Deterministic pre-check for objective (single-number) spreadsheet answers.

For every solution cell whose cached value is a plain number, compare it to
the same cell reference in the student's workbook and classify it as
auto_correct / auto_incorrect / missing_answer — arithmetic comparison,
not judgment, so doing it in Python instead of by LLM reasoning changes
nothing about grading quality (if anything it removes a source of
inconsistency between the two independent LLM reads).

This is a HINT, not a verdict: it only ever covers cells where the solution
value is a bare number. Anything else (text, conceptual answers, formulas
the solution treats as the real answer, cells missing from the solution) is
returned with status "needs_review", and every cell of a solution tab that
has no same-named tab in the submission comes back
"sheet_missing_in_submission" — the grader/checker must judge those
themselves, exactly as before. Never skip a cell silently.

Works for both .xlsx and legacy .xls (each file's own extension picks the
right extractor independently, so a .xls submission can be compared against
an .xlsx solutions file or vice versa) — .xls cells never carry a `formula`
(xlrd exposes only cached values), so `likely_downstream_of` hints simply
won't appear for them, which is correct rather than a bug.

Usage:
    .venv/Scripts/python scripts/compare_xlsx.py <submission.xlsx> <solutions.xlsx> [--tolerance 1e-4]
    .venv/Scripts/python scripts/compare_xlsx.py <submission.xls> <solutions.xlsx> [--tolerance 1e-4]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache  # noqa: E402
import extract_xls  # noqa: E402
import extract_xlsx  # noqa: E402


def get_workbook_record(path: Path) -> dict:
    """Dispatch to the extractor matching this file's own extension."""
    if path.suffix.lower() == ".xls":
        return cache.get_extraction(path, extract_xls.extract_xls)
    return cache.get_extraction(path, extract_xlsx.extract_xlsx)

CELL_REF_RE = re.compile(r"\$?([A-Z]{1,3})\$?([0-9]{1,7})")


def referenced_cells(formula: str) -> set[str]:
    """Best-effort extraction of same-sheet cell references from a formula
    string, e.g. "=B2*0.05" -> {"B2"}. Used only to hint that one flagged
    cell's error may be a downstream consequence of another — not a
    substitute for the grader/checker actually tracing it."""
    if not formula:
        return set()
    return {f"{col}{row}" for col, row in CELL_REF_RE.findall(formula)}


def is_plain_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def values_match(a: float, b: float, tolerance: float) -> bool:
    return abs(a - b) <= max(1e-9, abs(a) * tolerance)


def compare_workbooks(submission_record: dict, solution_record: dict, tolerance: float) -> dict:
    results = []
    counts = {"auto_correct": 0, "auto_incorrect": 0, "missing_answer": 0,
              "needs_review": 0, "sheet_missing_in_submission": 0}

    for sheet_name, sol_sheet in solution_record["sheets"].items():
        sub_sheet = submission_record["sheets"].get(sheet_name)
        if sub_sheet is None:
            for coord in sol_sheet["cells"]:
                results.append({
                    "sheet": sheet_name, "cell": coord,
                    "status": "sheet_missing_in_submission",
                    "note": f"solution tab '{sheet_name}' has no matching "
                            "tab in the submission — needs manual review",
                })
                counts["sheet_missing_in_submission"] += 1
            continue

        for coord, sol_cell in sol_sheet["cells"].items():
            sol_value = sol_cell.get("value")
            sub_cell = sub_sheet["cells"].get(coord, {})
            sub_value = sub_cell.get("value")
            sub_formula = sub_cell.get("formula")

            if not is_plain_number(sol_value):
                status = "needs_review"
                note = "solution value is not a plain number — requires grader/checker judgment"
            elif sub_value is None:
                status = "missing_answer"
                note = "submission cell is blank or has no cached value"
            elif not is_plain_number(sub_value):
                status = "needs_review"
                note = "submission value is not a plain number — compare manually"
            elif values_match(sol_value, sub_value, tolerance):
                status = "auto_correct"
                note = None
            else:
                status = "auto_incorrect"
                note = None

            counts[status] += 1
            results.append({
                "sheet": sheet_name, "cell": coord,
                "solution_value": sol_value, "student_value": sub_value,
                "status": status, "note": note,
                "_formula": sub_formula,  # consumed below, stripped before output
            })

    # Second pass: hint when an auto_incorrect cell's formula references
    # another auto_incorrect cell on the same sheet — it may be a cascading
    # consequence rather than an independent mistake. A hint only: the
    # grader/checker must still confirm this themselves.
    incorrect_by_sheet: dict[str, set[str]] = {}
    for r in results:
        if r["status"] == "auto_incorrect":
            incorrect_by_sheet.setdefault(r["sheet"], set()).add(r["cell"])

    for r in results:
        formula = r.pop("_formula", None)
        if r["status"] != "auto_incorrect" or not formula:
            continue
        flagged_on_sheet = incorrect_by_sheet.get(r["sheet"], set())
        referenced = referenced_cells(formula) & flagged_on_sheet - {r["cell"]}
        if referenced:
            r["likely_downstream_of"] = sorted(referenced)
            r["note"] = ("formula references other flagged cell(s) "
                         f"{sorted(referenced)} — may be a cascading "
                         "consequence rather than a separate error; verify")

    return {"tolerance": tolerance, "counts": counts, "cells": results}


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__,
                                      formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("submission", type=Path)
    parser.add_argument("solutions", type=Path)
    parser.add_argument("--tolerance", type=float, default=1e-4,
                         help="relative tolerance for numeric match (default 1e-4)")
    args = parser.parse_args()

    for p in (args.submission, args.solutions):
        if not p.exists():
            print(json.dumps({"error": f"file not found: {p}"}))
            sys.exit(1)
        if p.suffix.lower() not in (".xlsx", ".xls"):
            print(json.dumps({"error": f"unsupported file type for compare_xlsx.py: "
                                        f"{p.suffix} (expected .xlsx or .xls)"}))
            sys.exit(1)

    submission_record = get_workbook_record(args.submission)
    solution_record = get_workbook_record(args.solutions)

    result = compare_workbooks(submission_record, solution_record, args.tolerance)
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
