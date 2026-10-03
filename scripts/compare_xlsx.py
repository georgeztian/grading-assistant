#!/usr/bin/env python3
"""Cell-by-cell grading of a workbook against its solution — the rule every
grader and checker applies, decided in code:

- A cell with a formula is judged by its formula. The student's formula is
  re-run on the solution's (correct) inputs: if that gives the solution's
  value, the formula is right, and a wrong displayed number only carries an
  upstream error over — `carried_over`, never marked. Otherwise
  `formula_incorrect` (INCORRECT).
- A cell without a formula (a typed value) is judged by its value: right
  (`correct`) or wrong (`value_incorrect`, INCORRECT).
- An empty cell where the solution has a number or formula is
  `missing_answer` (INCOMPLETE).
- Every wrong cell is marked on its own, however many sit in one question.

What code can't decide is `needs_review` (the grader/checker judges it by the
same rule): text answers that differ from the solution's, text the solution
has where the student left the cell empty (an answer, or just a label or
instructor note?), formulas outside the evaluator's subset
(lib/formula_eval.py), a sheet whose layout is shifted, and a DEGRADED .xls
read without its formulas. A solution tab with no same-named tab in the
submission is `sheet_missing_in_submission` (look for a renamed tab).
Cells the verified rubric leaves out of grading (`excluded_blocks`, and
`ungraded_blocks`: notes and labels students aren't asked to produce) are
skipped, so which cells count is decided once, in the answer key.

annotate.py enforces the result (every MARK cell annotated with its label, no
annotation on a correct or carried-over cell), and audit_graded.py re-checks
it in every graded file.

Tolerance: a question's `numeric_tolerance` from the verified rubric
({"abs": x} or {"rel": x}) applies to that question's cells; everywhere else
the relative `--tolerance` (default 1e-4). Mismatch hints: `rounded`,
`percent_scale`, and `found_at` (the solution value sits elsewhere on the
student's sheet — several at one offset is a probable layout shift).

Output is compact by default: counts, then MARK / DO NOT MARK / JUDGE lists
and the correct cells. `--json` prints the full per-cell JSON.

Usage:
    .venv/Scripts/python scripts/compare_xlsx.py <submission.xlsx|.xls> <solutions.xlsx|.xls> [--tolerance 1e-4] [--json]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from openpyxl.utils.cell import coordinate_to_tuple  # noqa: E402
from lib import formula_eval, views  # noqa: E402
import extract  # noqa: E402
import rubric  # noqa: E402

DEFAULT_TOLERANCE = 1e-4
MARK = {"formula_incorrect": "INCORRECT", "value_incorrect": "INCORRECT",
        "missing_answer": "INCOMPLETE"}
NO_MARK = ("correct", "carried_over")
JUDGE = ("needs_review", "sheet_missing_in_submission")
RULE = ("Rule: mark a cell only if its own formula is wrong or it holds a wrong typed value "
        "(INCORRECT), or it is empty (INCOMPLETE). A correct formula whose number is wrong only "
        "because of an upstream error is NOT marked. Mark every wrong cell, not one per question.")
NUMERIC_TEXT_RE = re.compile(r"^\s*\$?\s*(-?[0-9][0-9,]*\.?[0-9]*(?:[eE][-+]?[0-9]+)?)\s*(%?)\s*$")


def is_plain_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def values_match(a: float, b: float, tolerance) -> bool:
    """`tolerance`: a relative float, or {"abs": x} / {"rel": x}."""
    if isinstance(tolerance, dict):
        if "abs" in tolerance:
            return abs(a - b) <= tolerance["abs"] + 1e-12
        tolerance = tolerance["rel"]
    return abs(a - b) <= max(1e-9, abs(a) * tolerance)


def same_value(sol, stu, tolerance) -> bool:
    if is_plain_number(sol) and is_plain_number(stu):
        return values_match(sol, stu, tolerance)
    if isinstance(sol, str) and isinstance(stu, str):
        return " ".join(sol.split()).casefold() == " ".join(stu.split()).casefold()
    return type(sol) is type(stu) and sol == stu


def numeric_text(value) -> float | None:
    """A number typed as text ("3.53%", "$1,200"), else None."""
    m = NUMERIC_TEXT_RE.match(value) if isinstance(value, str) else None
    if not m:
        return None
    number = float(m.group(1).replace(",", ""))
    return number / 100 if m.group(2) else number


def rubric_cells(solutions: Path, record: dict) -> tuple[dict, set, dict, str | None]:
    """From the verified rubric: (cell id -> numeric_tolerance, the cell ids
    not graded (excluded, or `ungraded_blocks`: notes/labels kept only as
    context), cell id -> question id, None). Without a usable rubric: empty
    ones, and the reason it can't be used."""
    try:
        rubric.verified_rubric(solutions)
    except rubric.RubricError as exc:
        return {}, set(), {}, str(exc)
    m = rubric.load_map(rubric.rubric_dir(solutions))
    units = views.units(record)
    pos = {u["id"]: i for i, u in enumerate(units)}
    tolerances, question_of, excluded = {}, {}, set()
    for q in m["questions"]:
        for i in rubric.resolve_list(q["blocks"], q["id"], record, units, pos, []):
            # a cell in two sections takes the first one's question and tolerance
            if units[i]["id"] not in question_of:
                question_of[units[i]["id"]] = q["id"]
                if q.get("numeric_tolerance"):
                    tolerances[units[i]["id"]] = q["numeric_tolerance"]
    for ex in m.get("excluded_blocks", []) + m.get("ungraded_blocks", []):
        for i in rubric.resolve_list(ex.get("blocks"), "excluded", record, units, pos, []):
            excluded.add(units[i]["id"])
    return tolerances, excluded, question_of, None


def mismatch_hint(sol, stu, tolerance) -> str | None:
    if not (is_plain_number(sol) and is_plain_number(stu)):
        return None
    for k in range(0, 7):
        if abs(round(sol, k) - stu) < 1e-9 * max(1.0, abs(stu)):
            return f"rounded: equals the solution rounded to {k} decimal(s)"
    if values_match(sol * 100, stu, tolerance) or values_match(sol, stu * 100, tolerance):
        return "percent_scale: equals the solution ×100 or ÷100"
    return None


def classify(sheet: str, coord: str, sol_cell: dict, sub_cell: dict, tolerance,
             correct_inputs, sub_record: dict, sol_record: dict, degraded: bool) -> dict:
    sol_value, sol_formula = sol_cell.get("value"), sol_cell.get("formula")
    stu_value, stu_formula = sub_cell.get("value"), sub_cell.get("formula")
    out = {"sheet": sheet, "cell": coord, "solution_value": sol_value,
           "solution_formula": sol_formula, "student_value": stu_value,
           "student_formula": stu_formula}

    def result(status: str, note: str | None = None, **extra) -> dict:
        out.update({"status": status, "note": note, **extra})
        if status in MARK:
            out["mark"] = MARK[status]
        return out

    if sol_value is None:
        return result("needs_review", "the solution cell has no cached value to compare with")
    if stu_value is None and stu_formula is None:
        if is_plain_number(sol_value) or sol_formula:
            return result("missing_answer", "the cell is empty")
        return result("needs_review", "the solution has text here and the student's cell is "
                                      "empty — mark it INCOMPLETE only if it is an answer the "
                                      "student had to give (not a label or instructor note)")

    if stu_formula:
        if formula_eval.normalize(stu_formula) == formula_eval.normalize(sol_formula):
            formula_ok, evaluated = True, None
        else:
            try:
                evaluated = formula_eval.evaluate(stu_formula, sheet, correct_inputs)
            except (formula_eval.Unsupported, formula_eval.ExcelError) as exc:
                return result("needs_review", f"the student's formula could not be re-run on the "
                                              f"correct inputs ({exc}) — compare it with the "
                                              "solution's formula by hand, by the same rule")
            formula_ok = same_value(sol_value, evaluated, tolerance)
        if not formula_ok:
            return result("formula_incorrect",
                          "the formula gives the wrong result even on the solution's inputs",
                          evaluated_on_correct_inputs=evaluated,
                          hint=mismatch_hint(sol_value, evaluated, tolerance))
        if stu_value is None or same_value(sol_value, stu_value, tolerance):
            return result("correct")
        upstream = []
        for ref_sheet, ref in formula_eval.references(stu_formula, sheet):
            s = sol_record["sheets"].get(ref_sheet, {}).get("cells", {}).get(ref, {})
            t = sub_record["sheets"].get(ref_sheet, {}).get("cells", {}).get(ref, {})
            if "value" in s and not same_value(s["value"], t.get("value"), tolerance):
                upstream.append(ref if ref_sheet == sheet else f"{ref_sheet}!{ref}")
        return result("carried_over",
                      "the formula is right; the value differs only because of upstream cell(s)"
                      if upstream else "the formula and its inputs match the solution; the "
                                       "stored value is stale (the workbook was not recalculated)",
                      carried_from=upstream[:6])

    # A typed value (no formula).
    if same_value(sol_value, stu_value, tolerance):
        return result("correct")
    if degraded:
        return result("needs_review", "DEGRADED .xls read without formulas: this value may come "
                                      "from a formula — decide by hand")
    if is_plain_number(sol_value) and isinstance(stu_value, str):
        number = numeric_text(stu_value)
        if number is not None and values_match(sol_value, number, tolerance):
            return result("needs_review", "the right number, typed as text — judge it")
        return result("value_incorrect", "a typed value that differs from the solution")
    if is_plain_number(sol_value) and is_plain_number(stu_value):
        return result("value_incorrect", "a typed value that differs from the solution",
                      hint=mismatch_hint(sol_value, stu_value, tolerance))
    return result("needs_review", "text that differs from the solution's — judge its substance")


def compare_workbooks(submission_record: dict, solution_record: dict,
                      tolerance: float = DEFAULT_TOLERANCE, tolerances: dict | None = None,
                      excluded: set | None = None, question_of: dict | None = None) -> dict:
    tolerances, excluded, question_of = tolerances or {}, excluded or set(), question_of or {}
    degraded = bool(submission_record.get("degraded"))

    def correct_inputs(sheet: str, coord: str):
        """The solution's value where it has one; the student's own cells
        (helper cells the solution doesn't have) otherwise."""
        if sheet not in solution_record["sheets"] and sheet not in submission_record["sheets"]:
            raise formula_eval.Unsupported(f"reference to a sheet {sheet!r} neither workbook has")
        cell = solution_record["sheets"].get(sheet, {}).get("cells", {}).get(coord)
        if cell is not None and "value" in cell:
            return cell["value"]
        cell = submission_record["sheets"].get(sheet, {}).get("cells", {}).get(coord)
        return cell.get("value") if cell else None

    results = []
    for sheet, sol_sheet in solution_record["sheets"].items():
        sub_sheet = submission_record["sheets"].get(sheet)
        for coord, sol_cell in sol_sheet["cells"].items():
            cid = f"{sheet}!{coord}"
            if cid in excluded or ("value" not in sol_cell and "formula" not in sol_cell):
                continue  # excluded by the rubric, or a comment-only cell
            if sub_sheet is None:
                entry = {"sheet": sheet, "cell": coord, "solution_value": sol_cell.get("value"),
                         "status": "sheet_missing_in_submission",
                         "note": f"the submission has no tab named {sheet!r} — look for a "
                                 "renamed tab and grade it by the same rule"}
            else:
                entry = classify(sheet, coord, sol_cell, sub_sheet["cells"].get(coord, {}),
                                 tolerances.get(cid, tolerance), correct_inputs,
                                 submission_record, solution_record, degraded)
            if cid in question_of:
                entry["question"] = question_of[cid]
            results.append(entry)

    # Where a missing/wrong solution value sits elsewhere on the student's
    # sheet — and whether many sit at the same offset (a layout shift, which
    # makes every same-address comparison on that sheet meaningless).
    shifts: dict[str, dict[tuple, int]] = {}
    for r in results:
        if r["status"] not in MARK or not is_plain_number(r["solution_value"]):
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
    for r in results:
        if r["sheet"] in layout and (r["status"] in MARK or r["status"] == "carried_over"):
            r.pop("mark", None)
            r["note"] = (f"PROBABLE LAYOUT SHIFT on this sheet (was {r['status']}): the student's "
                         "cells sit at other addresses — grade by content, by the same rule")
            r["status"] = "needs_review"

    counts: dict[str, int] = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    return {"tolerance": tolerance, "rubric_tolerances": len(tolerances), "counts": counts,
            "degraded_submission": degraded, "layout_shift": layout, "cells": results,
            "not_graded": sorted(excluded)}


def grade_cells(submission: Path, solutions: Path, tolerance: float = DEFAULT_TOLERANCE) -> dict:
    """compare_workbooks for two workbook files, with the verified rubric's
    tolerances and exclusions."""
    submission_record = extract.get_record(submission)[0]
    solution_record = extract.get_record(solutions)[0]
    tolerances, excluded, question_of, rubric_error = rubric_cells(solutions, solution_record)
    result = compare_workbooks(submission_record, solution_record, tolerance, tolerances,
                               excluded, question_of)
    if rubric_error:
        result["rubric_error"] = rubric_error
    return result


def expected_marks(result: dict) -> tuple[dict, dict]:
    """(cell id -> required label for every MARK cell, cell id -> entry for
    every cell that must NOT be marked)."""
    must, must_not = {}, {}
    for r in result["cells"]:
        cid = f"{r['sheet']}!{r['cell']}"
        if r["status"] in MARK:
            must[cid] = r
        elif r["status"] in NO_MARK:
            must_not[cid] = r
    return must, must_not


def describe(r: dict) -> str:
    """One line on why a cell is (not) marked, for scripts' messages."""
    if r["status"] == "formula_incorrect":
        return (f"formula {r['student_formula']} gives {_short(r.get('evaluated_on_correct_inputs'))} on "
                f"the correct inputs; solution {r.get('solution_formula') or ''} = "
                f"{_short(r['solution_value'])}")
    if r["status"] == "value_incorrect":
        return f"typed {_short(r['student_value'])}; solution {_short(r['solution_value'])}"
    if r["status"] == "missing_answer":
        return f"empty; solution {r.get('solution_formula') or ''} = {_short(r['solution_value'])}"
    if r["status"] == "carried_over":
        return (f"formula {r['student_formula']} is correct; its value is wrong only because of "
                f"{', '.join(r.get('carried_from') or []) or 'a stale calculation'}")
    return "its answer is correct"


def _short(value, limit: int = 80) -> str:
    text = views.fmt_value(value)
    return text if len(text) <= limit else text[:limit - 1] + "…"


def render_compact(result: dict) -> str:
    lines = [f"tolerance={result['tolerance']} (rubric tolerances on "
             f"{result['rubric_tolerances']} cell(s)) counts=" + json.dumps(result["counts"]), RULE]
    if result.get("rubric_error"):
        lines.insert(0, "NO USABLE ANSWER KEY — do not grade (its tolerances and not-graded cells "
                        f"are not applied below): {result['rubric_error']}")
    if result.get("degraded_submission"):
        lines.append("DEGRADED submission: read without formulas — wrong typed values are "
                     "needs_review, not value_incorrect")
    for sheet, shift in result["layout_shift"].items():
        lines.append(f"{json.dumps(sheet)} PROBABLE LAYOUT SHIFT: {shift['cells']} solution values "
                     f"sit {shift['rows']:+d} row(s) / {shift['cols']:+d} col(s) away on the "
                     "student's sheet — its cells are listed under JUDGE")
    cells = result["cells"]
    tag = lambda r: f" [{r['question']}]" if r.get("question") else ""  # noqa: E731

    marks = [r for r in cells if r["status"] in MARK]
    lines.append(f"MARK ({len(marks)}) — annotate every one, anchored at the cell, with this label:")
    for r in marks:
        line = f"  {r['sheet']}!{r['cell']} {r['mark']} {r['status']}: {describe(r)}"
        for key in ("hint", "found_at"):
            if r.get(key):
                line += f" {key}={r[key]}"
        lines.append(line + tag(r))

    carried = [r for r in cells if r["status"] == "carried_over"]
    lines.append(f"DO NOT MARK ({len(carried)}) — correct formula, wrong value carried over from "
                 "upstream (name them in the root cell's explanation instead):")
    for r in carried:
        lines.append(f"  {r['sheet']}!{r['cell']} {r['student_formula']} "
                     f"(from {', '.join(r.get('carried_from') or []) or 'a stale calculation'})")

    judge = [r for r in cells if r["status"] == "needs_review"]
    missing_sheets: dict[str, int] = {}
    for r in cells:
        if r["status"] == "sheet_missing_in_submission":
            missing_sheets[r["sheet"]] = missing_sheets.get(r["sheet"], 0) + 1
    lines.append(f"JUDGE ({len(judge) + sum(missing_sheets.values())}) — decide each by the same rule:")
    for sheet, n in missing_sheets.items():
        lines.append(f"  {json.dumps(sheet)} sheet_missing_in_submission: all {n} solution cells "
                     "— look for a renamed tab")
    for r in judge:
        stu = r.get("student_formula") or (_short(r["student_value"])
                                            if r.get("student_value") is not None else "")
        sol = r.get("solution_formula") or _short(r.get("solution_value"))
        lines.append(f"  {r['sheet']}!{r['cell']} needs_review: {r['note']} | solution {sol} | "
                     f"student {stu or '(empty)'}{tag(r)}")

    grouped: dict[str, list[str]] = {}
    for r in cells:
        if r["status"] == "correct":
            grouped.setdefault(r["sheet"], []).append(r["cell"])
    for sheet, refs in grouped.items():
        lines.append(f"{json.dumps(sheet)} correct ({len(refs)}): {' '.join(refs)}")
    return "\n".join(lines)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("submission", type=Path)
    parser.add_argument("solutions", type=Path)
    parser.add_argument("--tolerance", type=float, default=DEFAULT_TOLERANCE,
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
        result = grade_cells(args.submission, args.solutions, args.tolerance)
    except Exception as exc:  # conversion failure, corrupt file, …
        print(json.dumps({"error": f"{type(exc).__name__}: {exc}"}, ensure_ascii=False))
        sys.exit(1)
    if args.json:
        print(json.dumps(result, indent=2, ensure_ascii=False, default=str))
    else:
        print(render_compact(result))


if __name__ == "__main__":
    main()
