---
name: grading-instructions
description: grading instructions to grade students' homework submissions
---

# Concurrent Grading Workflow

Orchestrates parallel **grader agent** (red annotations for incorrect/incomplete only) + **independent checker agent** (single-pass verification, no corrections).

**Key Requirements**: 
- **Only edit files in `graded-submissions/`** (the shared extraction cache under `.cache/` is the one exception — see below)
- Content extraction for ALL content (text, equations, pictures, special symbols, formatting, spreadsheet tabs/formulas) via the shared extraction toolkit in `scripts/`
- Supports .doc, .docx, .pdf, .xlsx, and .xls, for both submissions and solutions
- Single verification pass (no correction loops)
- **No scoring**: neither agent calculates or writes a total score/grade (e.g. "8/10", "80%", a letter grade) — only per-question verdicts and explanations; scoring is left to the human instructor

**Prerequisite (once per machine)**: `pip install -r requirements.txt` — installs `python-docx`, `pypdf`, `openpyxl`, `lxml`, `PyMuPDF`, `xlrd<2.0`, `olefile`. `.doc` conversion additionally needs `pandoc` or LibreOffice on PATH.

**Shared extraction toolkit** (`scripts/`, documented in `scripts/README.md`): both agents run `python scripts/extract.py <file>` instead of writing ad hoc parsing code (caching is automatic — no flag needed). It implements the exact methods `grader.md`'s "Equation Extraction" and "Spreadsheet Extraction" sections describe (Word OMML equation walking, dual formula/cached-value spreadsheet loads reading every tab, `.doc` conversion, PDF equation-region flagging) and caches the result by file content hash under `.cache/extraction/`. Since grader and checker are already required to use the identical method for comparable results, this means:
- A solutions file shared by N students is extracted **once**, not once per student per agent.
- The checker's independent verdict still comes from its own judgment — it just reads the grader's cached extraction record instead of re-parsing the file from scratch (see `grading-checker.md` Step 2).
- For spreadsheets, `python scripts/compare_xlsx.py <submission> <solutions>` adds a deterministic pre-check for purely numeric answer cells (`auto_correct`/`auto_incorrect`/`missing_answer`), so neither agent spends LLM reasoning re-deriving arithmetic that Python can verify exactly. Every other cell — text, conceptual, formula-as-answer — still requires the agent's own judgment (`needs_review`), and both agents still write every explanation themselves.
- Fall back to the manual method in `grader.md` only for a file the script can't handle well.

**Annotation convention**: documents get red feedback text immediately below the wrong answer; spreadsheets get two markers instead — the wrong cell itself highlighted in Excel's "Bad" style (light red fill/dark red font) plus a red explanation in the closest empty cell to it. Full placement rules are in `grader.md` Step 3/4.

**Conceptual/explanation questions**: for written-sentence answers, grading isn't just right/wrong — the grader must check the student's answer against every key word/point the solution relies on and name any that are missing, even if the answer otherwise sounds reasonable. Full method is in `grader.md`'s "Conceptual/Explanation/Interpretation Questions" section; grader and checker must use the same method.

## Step 0: Discover Submissions & Match Solutions (before invoking any agent)

`student-submissions/` and `reference-solutions/` may each be flat, or organized into per-homework subfolders (e.g. `HW1/`, `HW2/`) — not always, and not always the same way on both sides. Before invoking the grader for any submission, resolve which solutions file grades it:

- **Flat submission** (sits directly in `student-submissions/`): pair it with the solutions file that sits directly in `reference-solutions/`.
  - **If more than one flat solutions file exists**, filename alone won't tell you which one applies (they aren't required to match) — open each candidate and match by subject matter (title/case name, file type, worksheet tab names, etc.). Only proceed once one candidate is clearly the same assignment; if two or more still look equally plausible, stop and ask the user to confirm.
- **Subfoldered submission** (sits in `student-submissions/<X>/`): pair it with the solutions file(s) in `reference-solutions/<X>/` — **the subfolder name must match exactly**. A submission in `HW1/` does not match a solutions folder named `Homework 1/` or `hw1_solutions/`.
- **Mixed layouts** (some submissions flat, others in subfolders, at the same time) are expected — resolve each submission independently by the same rule.
- **No confident match found** (no exact-name subfolder, or no flat solutions file clearly matches by content): do NOT guess which solutions file applies — stop and ask the user to confirm or provide the correct solutions file before grading anything in it.
- **Output mirrors the input structure**: a submission at `student-submissions/HW1/name.docx` outputs to `graded-submissions/HW1/name_Graded.docx` (create the subfolder if it doesn't exist yet); a flat submission continues to output directly into `graded-submissions/`.

## Concurrent Workflow Overview

Both agents work independently on different files:
- **Grader**: processes submission queue, one at a time
- **Checker**: processes submissions with "Grading Completed" mark, one at a time
- Agents can be invoked concurrently (in separate processes/sessions), capped at 10 concurrent instances per agent type — up to 10 graders and up to 10 checkers may run at the same time, but never more than 10 of either type simultaneously. With more than 10 submissions in the queue, launch the first 10 graders, then launch the next as each earlier one finishes (same pattern for checkers).

## Grader: Grade Submission

**Invoke** for each submission, passing `submission_file`, the matched `solutions_file` (per Step 0), and `output_dir` — `graded-submissions/`, or its matching homework subfolder if the submission was subfoldered. Full process, extraction requirements, and output format are defined in `grader.md` — grading is final, no correction mode.

## Checker: Single-Pass Verification

**Invoke** once a submission has its "Grading Completed" mark, passing `submission_file`, `graded_file` (the `_Graded.docx`/`_Graded.xlsx` in `graded-submissions/`), and `solutions_file`. Full process is defined in `grading-checker.md` — grades independently first, then compares against the grader's annotations.

**Final verdict** (single pass — never sent back to the grader for correction):
- **"Review Passed"** (blue mark): grading complete and verified ✓
- **"Review FAILED"** (blue mark), immediately followed by blue text in the same graded file explicitly stating every problem (no separate file): user reviews manually and decides next steps

### Discrepancy Types & Severity

| Type | Meaning | Severity |
|------|---------|----------|
| `verdict_mismatch` | Grader/checker disagree on correct/incorrect | high |
| `missed_question` | Grader skipped a question | high |
| `annotation_placement` | Feedback not immediately below answer | medium |
| `explanation_error` | Explanation unclear/incomplete/wrong | medium |
| `incomplete_coverage` | Multi-part question not fully addressed | medium |
| `cell_highlight_missing` | (.xlsx) Incorrect/incomplete cell not highlighted red, or a correct cell wrongly highlighted | medium |
| `missing_keypoint_not_flagged` | Grader missed a key word/point absent from a conceptual answer, or flagged one that's actually present | medium |
