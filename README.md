# Grading Assistant

An AI-agent-based system for grading student homework submissions inside [Claude Code]. Two independent agents — a **grader** and a **checker** — extract content from submissions and solution files, compare answers, and produce annotated copies with color-coded feedback, without ever touching the original student files.

## How it works

1. **Grader agent** extracts every piece of content from a submission and its matching solution file (text, equations, embedded images, spreadsheet tabs) via the shared extraction toolkit in [`scripts/`](scripts/README.md) — no ad hoc parsing per submission — compares each answer against the solution, and produces an annotated copy. **Only incorrect or incomplete answers get feedback** — correct answers are left untouched. When grading is done, the grader marks the file **"Grading Completed."**
2. **Checker agent** independently re-grades the same submission from scratch — without looking at the grader's annotations first — then compares its own verdicts against the grader's. It adds a final **"Review Passed"** or **"Review FAILED"** mark. If it finds discrepancies, it writes them as blue text directly below the "Review FAILED" mark, explicitly listing every problem — no separate file is created.
3. **No auto-correction loop.** This is a single verification pass by design: if the checker fails a submission, a human reads the in-file discrepancy notes and decides what to do — the checker never sends work back to the grader.

The two agents can run concurrently across a batch of submissions since each works on its own file.

### Efficient by construction

Extraction (parsing Word equations, reading every spreadsheet tab, converting legacy `.doc`/`.xls`) is mechanical and, per the workflow rules, must already be identical between grader and checker — so it's implemented once in [`scripts/`](scripts/README.md) and cached by file content hash instead of being re-derived as ad hoc code on every run:
- A solutions file shared by a whole class is extracted **once**, not once per student per agent.
- The checker reuses the grader's cached extraction record rather than re-parsing the file — its **verdict** is still formed independently, without looking at the grader's annotations, exactly as before.
- `scripts/compare_xlsx.py` deterministically flags purely numeric spreadsheet answers (`auto_correct`/`auto_incorrect`) so neither agent spends reasoning re-checking arithmetic Python can verify exactly — every other cell (text, conceptual, formula-based) is still graded and explained by the agent, same as always.

None of this changes annotation format, placement, verdicts, or the single-pass verification workflow — see [`scripts/README.md`](scripts/README.md) for what changed and why it doesn't affect output quality.

## Folder structure

```
reference-solutions/      # Answer key file(s) — .doc, .docx, .pdf, .xlsx, or .xls
student-submissions/      # Student submissions (same file types) — read-only, never modified
graded-submissions/       # Output only — annotated copies (discrepancies, if any, are noted inline)
scripts/                  # Shared extraction toolkit (see scripts/README.md) — git-tracked code
.cache/                   # Extraction cache written by scripts/ — git-ignored, safe to delete
.claude/agents/           # grader.md, grading-checker.md — the two agents' full logic
.claude/skills/grading-instructions/  # SKILL.md — orchestration workflow
requirements.txt          # Python dependencies for scripts/
CLAUDE.md                 # Quick-reference project rules
```

The three data folders (`reference-solutions/`, `student-submissions/`, `graded-submissions/`) are git-ignored — their contents stay on your machine and are never committed or pushed, since they hold student work. Only a `.gitkeep` in each is tracked so the folders exist after cloning. `.cache/` is also git-ignored (it holds only derived extraction data, safe to delete at any time).

### Optional per-homework subfolders

Both `reference-solutions/` and `student-submissions/` can either be flat, or organized into per-assignment subfolders (e.g. `HW1/`, `HW2/`) — independently on each side, and mixed layouts are fine (some assignments flat, others subfoldered, at the same time).

- A submission in `student-submissions/HW1/` is matched to `reference-solutions/HW1/` by **exact folder name** — `HW1` will not match `Homework 1` or `hw1_solutions`.
- A flat submission is matched to a flat solutions file by content (subject/case name, file type), since filenames aren't required to match — if more than one flat solutions file could plausibly apply, the system asks rather than guessing.
- If no confident match is found, the system stops and asks rather than guessing.
- Output mirrors the input: a submission from `HW1/` produces its graded file in `graded-submissions/HW1/`.

## Supported file types

| Input | Output |
|---|---|
| `.doc`, `.docx`, `.pdf` | `[name]_Graded.docx` |
| `.xlsx`, `.xls` | `[name]_Graded.xlsx` (`.xls` is always upgraded to `.xlsx`, since the legacy format can't reliably round-trip formatting) |

## Reading the annotations

**Documents (.docx output):**
- Red text is inserted immediately below each incorrect/incomplete answer, formatted as `**INCORRECT**: [why, and what the correct answer is]` or `**INCOMPLETE**: [what's missing]`.
- `Grading Completed` appears in red at the end of the document once the grader is done.
- `Review Passed` or `Review FAILED` is added in **blue** as its own paragraph by the checker, right after the grader's red "Grading Completed" mark. If it's a "Review FAILED," every discrepancy is listed in blue text right below the mark, in the same file.

**Spreadsheets (.xlsx output):**
- Each incorrect/incomplete answer **cell itself** is highlighted using Excel's standard "Bad" style (light red fill, dark red font) — the value/formula is never changed, only its formatting.
- An explanation is written in red text into the closest empty cell to it (right, then below, then further out) — existing cell content is never overwritten.
- A dedicated **"Grading Summary"** tab (added as the last sheet) carries `Grading Completed` (red) and, once checked, `Review Passed`/`Review FAILED` (blue).

**Conceptual / short-answer / essay questions:** for written-sentence answers, the system doesn't just judge overall direction — it checks the student's answer against every key term or point the solution's explanation relies on. If the answer is coherent but missing a specific point, it's marked **INCOMPLETE** and every missing point is named explicitly (never a vague "explanation incomplete").

If nothing is annotated and the mark is "Review Passed," every question was answered correctly.

**No total score:** neither agent calculates or writes a total score/grade (e.g. "8/10", "80%", a letter grade) — only per-question verdicts and explanations are recorded. Totaling up a grade from the per-question verdicts is left to the human instructor.

## How to use it

### Prerequisites
- Claude Code with access to this repository.
- Python 3 with the extraction toolkit's dependencies installed: `pip install -r requirements.txt` (`python-docx`, `pypdf`, `openpyxl`, `lxml`, `PyMuPDF`, `xlrd<2.0`, `olefile`).
- A `.doc`-to-`.docx` converter for legacy Word files: `LibreOffice` (headless) or `pandoc` — either is picked up automatically by `scripts/convert_doc.py`.

### 1. Add your files
- Put the answer key in `reference-solutions/`.
- Put student submissions in `student-submissions/`.
- Optionally organize either folder into per-homework subfolders (see above).

### 2. Ask Claude to grade
Just ask, in plain language, for example:
- "grade all student submissions"
- "grade student-submissions/HW1/john_doe.docx against reference-solutions/HW1/answer_key.docx"

This runs the `/grading-instructions` skill, which will:
1. Inspect the folder structure and match each submission to its solution file (asking you to confirm if a match is ambiguous).
2. Run the grader agent on each submission.
3. Run the checker agent once a submission is marked "Grading Completed."

### 3. Review the output
- Open the corresponding file in `graded-submissions/`.
- Read the red annotations for what was marked wrong and why.
- Check the blue mark at the end (or on the "Grading Summary" tab) for the checker's final verdict.
- If it says **Review FAILED**, read the discrepancies listed in blue text right below the mark in that same graded file, and decide manually how to proceed — the system will not auto-correct or re-grade on its own.

## Important: this is AI-assisted grading

The grader and checker are independent, but both are AI agents interpreting free-form student work — treat "Review Passed" as a strong second opinion, not an infallible verdict. Spot-check a sample of graded files yourself, especially for high-stakes grading, ambiguous/creative answers, or anything the checker flagged as low-confidence or ambiguous.

## Where to look for more detail

- [`.claude/agents/grader.md`](.claude/agents/grader.md) — full grader logic and extraction methods
- [`.claude/agents/grading-checker.md`](.claude/agents/grading-checker.md) — full checker logic
- [`.claude/skills/grading-instructions/SKILL.md`](.claude/skills/grading-instructions/SKILL.md) — orchestration workflow, folder-matching rules, discrepancy types
- [`scripts/README.md`](scripts/README.md) — the shared extraction toolkit: what it does, why it's cached, how to extend it
- [`CLAUDE.md`](CLAUDE.md) — quick-reference project rules
