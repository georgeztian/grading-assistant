# Project Overview
Grade students' homework submissions using concurrent **grader agent** (red annotations) + **independent checker agent** (single-pass blue verification). Agents work in parallel on different files.

# Workflow Rules
- **Concurrency cap**: at most 10 concurrent agents per type (10 graders and 10 checkers may run together); queue the rest and launch more as earlier ones finish (see `SKILL.md`)
- **Only edit files in `graded-submissions/`** (the shared extraction cache in `.cache/` is the one exception — derived, git-ignored, not grading output)
- **Content extraction**: both agents run the shared toolkit in `scripts/` (`python scripts/extract.py <file>`) rather than writing ad hoc parsing code — it's cached by file hash automatically, so a solutions file or submission is only ever extracted once, even when reused across students or reread by the checker. See `scripts/README.md`.
- **Submissions**: Accepts .doc, .docx, .pdf, .xlsx, and .xls files
- **Annotations**: Only annotate incorrect/incomplete answers (correct answers need no feedback). Placement convention (documents vs. spreadsheets), content extraction requirements (equations, spreadsheet tabs), and conceptual-question key-point checking are detailed in `grader.md` and summarized in `SKILL.md` — both grader and checker must follow the same methods
- **No scoring**: Agents never calculate or write a total score/grade (e.g. "8/10", "80%", a letter grade) — only per-question correct/incorrect/partial verdicts and explanations. Scoring is left entirely to the human instructor.
- **Verification**: Single-pass check by grader-checker (no correction loops)
  - If "Review Passed" → grading complete ✓
  - If "Review FAILED" → explicitly state problems
- See SKILL.md for workflow details

# To Grade Submissions
Use the `/grading-instructions` skill for detailed workflow and invocation instructions.

# Project Structure
- The three data folders below are git-ignored (only `.gitkeep` is tracked) — never `git add` or force-add their contents; they hold student work
- `reference-solutions/` — Solution files (.doc, .docx, .pdf, .xlsx, .xls)
- `student-submissions/` — Student submissions (.doc, .docx, .pdf, .xlsx, or .xls, read-only)
- `graded-submissions/` — Output:
  - `[name]_Graded.docx` — annotated document (grader + checker marks), for .doc/.docx/.pdf input
  - `[name]_Graded.xlsx` — annotated workbook (grader + checker marks on a "Grading Summary" tab, cell-level annotations on original tabs), for .xlsx/.xls input
  - If verification fails, the checker states every problem in blue text directly below the "Review FAILED" mark in this same file — no separate file is created
- **Per-homework subfolders (optional)**: `reference-solutions/` and `student-submissions/` may each be flat or organized into subfolders (e.g. `HW1/`, `HW2/`), not always the same way on both sides. Agents must match a submission's subfolder to a solutions subfolder by exact name before grading (see `SKILL.md` Step 0), and output mirrors the input structure (e.g. `graded-submissions/HW1/`).
- `.claude/agents/` → `grader.md`, `grading-checker.md`
- `.claude/skills/grading-instructions/` → Full workflow
- `scripts/` — shared extraction toolkit (git-tracked code, not data): `extract.py` (unified entry point), `extract_docx.py`/`extract_pdf.py`/`extract_xlsx.py`/`extract_xls.py`/`convert_doc.py` (per-format extractors), `compare_xlsx.py` (deterministic numeric pre-check). See `scripts/README.md`.
- `.cache/` — git-ignored extraction cache written by `scripts/` (`.cache/extraction/`, `.cache/renders/`, `.cache/converted/`); safe to delete, everything re-extracts on next use
- `requirements.txt` — Python dependencies for `scripts/` (`pip install -r requirements.txt` once per machine)


