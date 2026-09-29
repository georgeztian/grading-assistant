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
returned with status "needs_review" (or "text_match" when the student's text
is identical to the solution's, whitespace aside), and every cell of a solution tab that
has no same-named tab in the submission comes back
"sheet_missing_in_submission" — the grader/checker must judge those
themselves, exactly as before. Never skip a cell silently.

Works for .xlsx and legacy .xls in any combination: both are read through
extract.py's cached record (a .xls via its .xlsx conversion, so its formulas
— and therefore `likely_downstream_of` hints — are available too; only a
DEGRADED values-only .xls reading lacks them).

Tolerance: a question's `numeric_tolerance` from the verified rubric
({"abs": x} or {"rel": x}) applies to that question's cells; everywhere
else the relative `--tolerance` (default 1e-4). A mismatch that is really a
presentation difference is labelled rather than left as a bare
auto_incorrect: `rounded` (the solution rounded to the student's decimals),
`percent_scale` (×100 / ÷100). And when a solution value does appear
elsewhere on the student's sheet, `found_at` says where — several cells
shifted by the same offset are reported as a probable layout shift (an
inserted row/column), not as a column of wrong answers.

Output is compact by default — counts, then per sheet the auto_correct /
text_match cells as a list of refs, then one line per cell that still needs
attention (long text truncated; the full text is in both files' extraction
views). `--json` prints the full per-cell JSON instead.

Usage:
    .venv/Scripts/python scripts/compare_xlsx.py <submission.xlsx> <solutions.xlsx> [--tolerance 1e-4] [--json]
    .venv/Scripts/python scripts/compare_xlsx.py <submission.xls> <solutions.xlsx> [--tolerance 1e-4] [--json]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import views  # noqa: E402
from lib.views import fmt_value  # noqa: E402
import rubric  # noqa: E402
from openpyxl.utils.cell import (column_index_from_string, coordinate_to_tuple,  # noqa: E402
                                 get_column_letter)
import extract  # noqa: E402


def get_workbook_record(path: Path) -> dict:
    """The same cached record extract.py produces."""
    return extract.get_record(path)[0]


# A same-sheet cell reference or range: not preceded by "!" (another sheet's
# cell) or by a letter/underscore (part of a name), and not followed by "("
# or more name characters (a function such as LOG10(…)).
CELL_REF_RE = re.compile(
    r"(?<![!A-Za-z_$])\$?([A-Z]{1,3})\$?([0-9]{1,7})"
    r"(?::\$?([A-Z]{1,3})\$?([0-9]{1,7}))?(?![0-9A-Za-z_(])")
# Text literals and quoted sheet names ('Other sheet'!C5) hold no same-sheet refs.
QUOTED_RE = re.compile(r'"[^"]*"|\'(?:[^\']|\'\')*\'')
MAX_RANGE_CELLS = 400


def referenced_cells(formula: str) -> set[str]:
    """Best-effort extraction of same-sheet cells a formula reads, e.g.
    "=B2*0.05" -> {"B2"}, "=SUM(G11:G13)" -> {"G11", "G12", "G13"};
    other-sheet references and text literals are ignored. Used only to hint
    that one flagged cell's error may be a downstream consequence of another
    — not a substitute for the grader/checker actually tracing it."""
    if not formula:
        return set()
    refs = set()
    for c1, r1, c2, r2 in CELL_REF_RE.findall(QUOTED_RE.sub("", formula)):
        if not c2:
            refs.add(f"{c1}{r1}")
            continue
        lo_c, hi_c = sorted((column_index_from_string(c1), column_index_from_string(c2)))
        lo_r, hi_r = sorted((int(r1), int(r2)))
        if (hi_c - lo_c + 1) * (hi_r - lo_r + 1) > MAX_RANGE_CELLS:
            refs.update({f"{c1}{r1}", f"{c2}{r2}"})
            continue
        refs.update(f"{get_column_letter(c)}{r}"
                    for c in range(lo_c, hi_c + 1) for r in range(lo_r, hi_r + 1))
    return refs


def is_plain_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def values_match(a: float, b: float, tolerance) -> bool:
    """`tolerance`: a relative float, or {"abs": x} / {"rel": x}."""
    if isinstance(tolerance, dict):
        if "abs" in tolerance:
            return abs(a - b) <= tolerance["abs"] + 1e-12
        tolerance = tolerance["rel"]
    return abs(a - b) <= max(1e-9, abs(a) * tolerance)


def rubric_tolerances(solutions: Path) -> dict[str, dict]:
    """Cell id (Sheet!A1) -> numeric_tolerance, for questions of the
    verified rubric that set one. Empty if there is no usable rubric."""
    try:
        rubric.verified_rubric(solutions)
    except rubric.RubricError:
        return {}
    m = rubric.load_map(rubric.rubric_dir(solutions))
    record = get_workbook_record(solutions)
    units = views.units(record)
    pos = {u["id"]: i for i, u in enumerate(units)}
    out = {}
    for q in m["questions"]:
        if q.get("numeric_tolerance"):
            for i in rubric.resolve_list(q["blocks"], q["id"], record, units, pos, []):
                out[units[i]["id"]] = q["numeric_tolerance"]
    return out


def mismatch_hint(sol: float, stu: float, tolerance) -> str | None:
    for k in range(0, 7):
        if abs(round(sol, k) - stu) < 1e-9 * max(1.0, abs(stu)):
            return f"rounded: equals the solution rounded to {k} decimal(s)"
    if values_match(sol * 100, stu, tolerance) or values_match(sol, stu * 100, tolerance):
        return "percent_scale: equals the solution ×100 or ÷100"
    return None


def compare_workbooks(submission_record: dict, solution_record: dict, tolerance: float,
                      tolerances: dict[str, dict] | None = None) -> dict:
    tolerances = tolerances or {}
    results = []
    counts = {"auto_correct": 0, "auto_incorrect": 0, "missing_answer": 0,
              "text_match": 0, "needs_review": 0, "sheet_missing_in_submission": 0}

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

            if (isinstance(sol_value, str) and isinstance(sub_value, str)
                    and " ".join(sol_value.split()) == " ".join(sub_value.split())):
                status = "text_match"
                note = None
            elif not is_plain_number(sol_value):
                status = "needs_review"
                note = "solution value is not a plain number — requires grader/checker judgment"
            elif sub_value is None:
                status = "missing_answer"
                note = "submission cell is blank or has no cached value"
            elif not is_plain_number(sub_value):
                status = "needs_review"
                note = "submission value is not a plain number — compare manually"
            elif values_match(sol_value, sub_value, tolerances.get(f"{sheet_name}!{coord}", tolerance)):
                status = "auto_correct"
                note = None
            else:
                status = "auto_incorrect"
                note = mismatch_hint(sol_value, sub_value,
                                     tolerances.get(f"{sheet_name}!{coord}", tolerance))

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

    # Where a wrong/missing solution value sits elsewhere on the student's
    # sheet — and whether many of them sit at the same offset (layout shift).
    shifts: dict[str, dict[tuple, int]] = {}
    for r in results:
        if r["status"] not in ("auto_incorrect", "missing_answer"):
            continue
        sub_cells = submission_record["sheets"].get(r["sheet"], {}).get("cells", {})
        tol = tolerances.get(f"{r['sheet']}!{r['cell']}", tolerance)
        found = [c for c, cell in sub_cells.items() if c != r["cell"]
                 and is_plain_number(cell.get("value"))
                 and values_match(r["solution_value"], cell["value"], tol)]
        if found:
            r["found_at"] = found[:3]
            row, col = coordinate_to_tuple(r["cell"])
            for c in found[:3]:
                frow, fcol = coordinate_to_tuple(c)
                offset = (frow - row, fcol - col)
                sheet_shifts = shifts.setdefault(r["sheet"], {})
                sheet_shifts[offset] = sheet_shifts.get(offset, 0) + 1
    layout = {}
    for sheet, offsets in shifts.items():
        offset, n = max(offsets.items(), key=lambda kv: kv[1])
        if n >= 3:
            layout[sheet] = {"rows": offset[0], "cols": offset[1], "cells": n}
    return {"tolerance": tolerance, "rubric_tolerances": len(tolerances), "counts": counts,
            "layout_shift": layout, "cells": results}


def _short(value, limit: int = 80) -> str:
    text = fmt_value(value)
    return text if len(text) <= limit else text[:limit - 1] + "…"


def render_compact(result: dict) -> str:
    lines = [f"tolerance={result['tolerance']} (rubric tolerances on "
             f"{result['rubric_tolerances']} cell(s)) counts=" + json.dumps(result["counts"])]
    for sheet, shift in result["layout_shift"].items():
        lines.append(f"{json.dumps(sheet)} PROBABLE LAYOUT SHIFT: {shift['cells']} solution values "
                     f"sit {shift['rows']:+d} row(s) / {shift['cols']:+d} col(s) away on the "
                     "student's sheet — compare by content there, not by cell address")
    grouped: dict[str, dict[str, list[str]]] = {}
    attention = []
    missing_sheets: dict[str, int] = {}
    for r in result["cells"]:
        if r["status"] in ("auto_correct", "text_match"):
            grouped.setdefault(r["sheet"], {}).setdefault(r["status"], []).append(r["cell"])
        elif r["status"] == "sheet_missing_in_submission":
            missing_sheets[r["sheet"]] = missing_sheets.get(r["sheet"], 0) + 1
        else:
            attention.append(r)
    for sheet, by_status in grouped.items():
        for status, cells in by_status.items():
            lines.append(f"{json.dumps(sheet)} {status} ({len(cells)}): {' '.join(cells)}")
    for sheet, n in missing_sheets.items():
        lines.append(f"{json.dumps(sheet)} sheet_missing_in_submission: all {n} solution cells "
                     "— look for a renamed tab before calling them missing")
    for r in attention:
        line = (f"{r['sheet']}!{r['cell']} {r['status']} sol={_short(r.get('solution_value'))} "
                f"stu={_short(r.get('student_value'))}")
        if r.get("likely_downstream_of"):
            line += f" likely_downstream_of={r['likely_downstream_of']}"
        if r["status"] == "auto_incorrect" and r.get("note") and not r.get("likely_downstream_of"):
            line += f" [{r['note']}]"
        if r.get("found_at"):
            line += f" found_at={r['found_at']}"
        lines.append(line)
    return "\n".join(lines)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__,
                                      formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("submission", type=Path)
    parser.add_argument("solutions", type=Path)
    parser.add_argument("--tolerance", type=float, default=1e-4,
                         help="relative tolerance for numeric match (default 1e-4)")
    parser.add_argument("--json", action="store_true",
                         help="print the full per-cell JSON instead of the compact listing")
    args = parser.parse_args()

    for p in (args.submission, args.solutions):
        if not p.exists():
            print(json.dumps({"error": f"file not found: {p}"}))
            sys.exit(1)
        if p.suffix.lower() not in (".xlsx", ".xls"):
            print(json.dumps({"error": f"unsupported file type for compare_xlsx.py: "
                                        f"{p.suffix} (expected .xlsx or .xls)"}))
            sys.exit(1)

    try:
        submission_record = get_workbook_record(args.submission)
        solution_record = get_workbook_record(args.solutions)
        tolerances = rubric_tolerances(args.solutions)
    except Exception as exc:  # conversion failure, corrupt file, …
        print(json.dumps({"error": f"{type(exc).__name__}: {exc}"}, ensure_ascii=False))
        sys.exit(1)

    result = compare_workbooks(submission_record, solution_record, args.tolerance, tolerances)
    if args.json:
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        print(render_compact(result))


if __name__ == "__main__":
    main()
