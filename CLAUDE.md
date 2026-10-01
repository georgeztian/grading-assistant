# Project Overview
Grade students' homework submissions using concurrent **grader agents** (red annotations) and **independent checker agents** (single-pass blue verification), which work in parallel on different files. Before any grading, each solutions file is turned once into a **verified answer key (rubric)**. A **rubric-builder** maps the solution, verbatim, question by question, and an independent **rubric-checker** approves it or has it corrected. Every grader and checker then reads that rubric instead of the raw solutions.

# Workflow Rules
- **Concurrency cap**: at most 10 concurrent agents per type (10 graders and 10 checkers may run together); queue the rest and launch more as earlier ones finish (see `SKILL.md`)
- **Only edit files in `graded-submissions/`** (the shared cache in `.cache/` — extraction, rubrics, per-file work dirs — is the one exception: derived, git-ignored, not grading output)
- **Answer key first**: no grader or checker starts on a homework until `scripts/rubric.py path <solutions>` succeeds: the rubric is built, independently verified with a per-question review record, unchanged, and not on hold. Any issue the rubric-checker finds must be corrected and re-verified first (see `SKILL.md` Step 1).
- **Solution concerns pause the homework**: a suspected error in the official solution (raised by the rubric-checker, a grader or a checker) puts the rubric on hold (`rubric.py hold`). Nobody grades against it until the user decides and the orchestrator runs `rubric.py release` (or the key is corrected and re-verified). Graders and checkers never grade around the rubric.
- **Grading records its answer key**: every graded file records which verified rubric it was graded against, and tags each annotation with its question. After any answer-key change, `rubric.py graded` lists outdated graded files to regrade.
- **Content extraction**: all agents run the shared toolkit in `scripts/` (`.venv/Scripts/python scripts/extract.py <file>`) rather than writing ad hoc parsing code. It's cached by file hash, so each file is extracted once. Agents read its compact **view** (`view_path`), not the JSON. `.doc`/`.xls` are converted by Microsoft Word/Excel (COM) or LibreOffice, so formulas, formatting and images survive. Embedded images are exported, and their paths shown inline in the view.
- **File writing is scripted**: graders run `scripts/annotate.py`, and checkers run `scripts/audit_graded.py` and `scripts/mark_review.py`. Agents never hand-write python-docx/openpyxl code for annotations or marks, so pinned colors, placement and wording come from code. See `scripts/README.md`.
- **Python environment**: always run Python via the project's private venv — `.venv/Scripts/python` (Windows; `.venv/bin/python` on macOS/Linux), never bare `python`/`pip` — for the `scripts/` toolkit *and* any ad hoc code. If `.venv/` is missing, run `python scripts/setup_env.py` once first (the only time a system `python` is used). Toolkit scripts relaunch themselves into `.venv` if started with the wrong interpreter, but ad hoc code does not.
- **Submissions**: Accepts .doc, .docx, .pdf, .xlsx, and .xls files
- **Annotations**: Only annotate incorrect/incomplete answers (correct answers need no feedback). Annotations are student-facing. They never mention the rubric, solution, answer key or key points (students have no access to them). They state the correct answer and working directly. `annotate.py` refuses such wording, and `audit_graded.py` flags it. Placement convention (documents vs. spreadsheets) and conceptual-question key-point checking (against the rubric's key points) are detailed in `grader.md` and summarized in `SKILL.md`. Grader and checker must follow the same methods. Manual extraction fallbacks are in `.claude/skills/grading-instructions/extraction-fallback.md`
- **Spreadsheets are graded cell by cell, by formula**: a cell is INCORRECT only if its own formula is wrong or it holds a wrong typed value (no formula), and INCOMPLETE if it is an empty answer cell. A correct formula whose number is wrong only because of an upstream error is never marked. Every wrong cell is marked, however many sit in one question. `compare_xlsx.py` decides this in code by re-running each student formula on the solution's inputs, `annotate.py` enforces it, and `audit_graded.py` re-checks it. Which cells are graded is settled once in the verified rubric: every cell of a section, except the notes and labels its map lists as `ungraded_blocks`.
- **No scoring**: Agents never calculate or write a total score/grade (e.g. "8/10", "80%", a letter grade) — only per-question (per-cell for spreadsheets) correct/incorrect/partial verdicts and explanations. Scoring is left entirely to the human instructor.
- **Verification**: Single-pass check by grader-checker (no correction loops). The checker commits its own blind verdicts to a file before the audit report (grader's verdicts plus an automatic comparison) can be produced, and `mark_review.py` refuses to mark a file without that blind audit.
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
- `.claude/agents/` → `rubric-builder.md`, `rubric-checker.md`, `grader.md`, `grading-checker.md`
- `.claude/skills/grading-instructions/` → Full workflow (`SKILL.md`) and `extraction-fallback.md` (manual extraction / image inspection, read only when needed)
- `scripts/` — shared grading toolkit (git-tracked code, not data). See `scripts/README.md`:
  - `extract.py`: unified entry point, which writes the compact view
  - `extract_docx.py` / `extract_pdf.py` / `extract_xlsx.py` / `extract_xls.py`: per-format extractors
  - `convert_doc.py`: `.doc`/`.xls` conversion via Word/Excel COM or LibreOffice
  - `rubric.py`: answer key build/verify gate
  - `compare_xlsx.py`: the spreadsheet cell rule (which cells to mark, which carry an upstream error)
  - `annotate.py`: grader output
  - `audit_graded.py`: the checker's mechanical audit
  - `mark_review.py`: checker marks
  - `setup_env.py`: creates `.venv/`
  - `_venv.py`: interpreter guard every script imports
  - `lib/`: shared helpers (views, PDF→docx rebuild, image export, pinned marks, cache, Excel formula evaluator)
- `.cache/` — git-ignored cache written by `scripts/`:
  - `.cache/extraction/`: records + views
  - `.cache/converted/`
  - `.cache/renders/`: page renders + exported images
  - `.cache/rubrics/`: answer keys + verification status
  - `.cache/work/`: per-file agent scratch JSON
  - `.cache/inspect/`: ad hoc manual inspection
  It's safe to delete, and everything re-extracts on next use, but deleting `.cache/rubrics/` means answer keys must be rebuilt and re-verified. Use it (never a made-up path in the project root or an OS temp dir) for any one-off inspection.
- `requirements.txt` — Python dependencies for `scripts/` (including `pywin32` on Windows for Word/Excel conversion), installed into `.venv/` by `scripts/setup_env.py`
- `.venv/` — private per-machine Python environment (git-ignored, excluded from Dropbox sync); create/refresh with `python scripts/setup_env.py`
