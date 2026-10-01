# Grading Assistant

An AI-agent-based system for grading student homework submissions. For each homework, a verified **answer key** is built once from the solution file. Independent **grader** and **checker** agents then grade every submission against it and produce annotated copies with color-coded feedback, without ever touching the original student files.

## How it works

1. **Answer key (once per homework).**
   - A **rubric-builder** agent turns the solution file into a question-by-question answer key (the "rubric").
   - The rubric is a **complete, verbatim reorganization of the solution**. The agent only decides which part of the solution belongs to which question, and a script copies the content word for word. That includes every calculation step, formula, intermediate value and explanation, so feedback quality doesn't depend on a paraphrase.
   - The script refuses a rubric that leaves any solution content unassigned.
   - For conceptual questions, the rubric also lists every key point a complete answer must cover.
   - For a spreadsheet, the rubric is a cell-level answer key, since spreadsheets are graded cell by cell. Every solution cell is copied with its formula and value. The rubric also settles once which cells are graded: instructor notes and labels a student isn't asked to produce stay in the key as context, marked "not graded", so no grader has to decide about them. Rounding tolerances are recorded in a form the cell check applies.
   - An independent **rubric-checker** agent then compares the rubric against the raw solution file and records a per-question review. Approval is refused without it. If it finds any issue, the builder must correct it and a fresh checker re-verifies it.
   - **If the official solution itself looks wrong**, the answer key is put **on hold**. That can come from the rubric-checker, or from any grader or checker who spots it. Grading of that homework pauses and you're asked to decide: the solution stands as written (grading resumes), or it gets corrected or another answer is also accepted (the key is updated and re-verified first).
   - **No grading starts until the answer key is verified.** After 3 rejected rounds, the system stops and asks you.
2. **Grader agent.**
   - It reads the verified answer key and the submission, and decides each question: correct, incorrect or incomplete. Spreadsheets are decided **cell by cell** instead.
   - It writes an explanation for every incorrect or incomplete answer; **correct answers get no feedback**.
   - A script writes the annotated copy and marks it **"Grading Completed"**, so colors, placement and wording are always identical.
   - The file invisibly records which version of the answer key (and of the solution file) it was graded against, and which question each annotation belongs to. If the key or the solution file later changes, outdated files are found and regraded; one that already carries a checker review is only regraded after asking you.
3. **Checker agent.**
   - It first grades the same submission **blind**, against the same answer key, and commits its verdicts to a file. The grader's verdicts can't be viewed through the workflow's audit until it has done so.
   - A script then compares both sets of verdicts question by question (cell by cell for spreadsheets), and audits the mechanics: mark present and exact, annotations placed correctly, correct colors, student content unaltered, current answer key.
   - It adds **"Review Passed"** or **"Review FAILED"**. For a failure, every problem is listed in blue text directly below the mark, in the same file.
4. **No auto-correction loop for grading.** This is a single verification pass by design. If the checker fails a submission, you read the in-file discrepancy notes and decide what to do; the checker never sends work back to the grader.

Graders and checkers run concurrently across a batch (up to 10 of each at a time), since each works on its own file.

### Efficient by construction

All mechanical work is done once, in code, by the toolkit in [`scripts/`](scripts/README.md):
- **Answer key once.** The solution file is read and organized once per homework. Every grader and checker reads the compact, verified rubric instead of the raw solution file, and all of them use the same key-point list, so verdicts on conceptual questions are consistent across students.
- **Extract once, read compactly.** Every file is extracted once, cached by content hash, and read through a compact text view, which is 60–90% smaller than the raw extraction on real homework. The checker reuses the grader's extraction; its *verdict* is still formed independently.
- **Scripts write the files.** Annotations, review marks and the checker's mechanical audit are single commands. Agents don't write and debug document-editing code for every submission.
- **Spreadsheet checks by Python.** `compare_xlsx.py` re-runs every student formula on the solution's inputs, so it decides in code which cells are wrong. No agent spends reasoning on what Python can verify exactly. Cells code can't decide (such as text answers) are judged by the agent, and every explanation is the agent's.

## Folder structure

```
reference-solutions/      # Answer key file(s) — .doc, .docx, .pdf, .xlsx, or .xls
student-submissions/      # Student submissions (same file types) — read-only, never modified
graded-submissions/       # Output only — annotated copies (discrepancies, if any, are noted inline)
scripts/                  # Shared grading toolkit (see scripts/README.md) — git-tracked code
.cache/                   # Written by scripts/ — extractions, verified answer keys, scratch; git-ignored
.claude/agents/           # rubric-builder.md, rubric-checker.md, grader.md, grading-checker.md
.claude/skills/grading-instructions/  # SKILL.md (orchestration workflow) + extraction-fallback.md
requirements.txt          # Python dependencies for scripts/ (installed into .venv/)
.venv/                    # Private Python environment — per machine, git-ignored, not synced by Dropbox
CLAUDE.md                 # Quick-reference project rules
```

The three data folders (`reference-solutions/`, `student-submissions/`, `graded-submissions/`) are git-ignored. Their contents stay on your machine and are never committed or pushed, since they hold student work. Only a `.gitkeep` in each is tracked, so the folders exist after cloning. 
`.cache/` and `.venv/` are also git-ignored. Both are safe to delete and rebuild. Deleting `.cache/rubrics/` means each homework's answer key is rebuilt and re-verified on the next run.

### Optional per-homework subfolders

Both `reference-solutions/` and `student-submissions/` can be flat or organized into per-assignment subfolders (e.g. `HW1/`, `HW2/`). The two sides are independent, and mixed layouts are fine (some assignments flat, others subfoldered, at the same time).

- A submission in `student-submissions/HW1/` is matched to `reference-solutions/HW1/` by **exact folder name**. `HW1` will not match `Homework 1` or `hw1_solutions`.
- A flat submission is matched to a flat solutions file by content (subject/case name, file type), since filenames aren't required to match. If more than one flat solutions file could plausibly apply, the system asks rather than guessing.
- If no confident match is found, the system stops and asks.
- Output mirrors the input: a submission from `HW1/` produces its graded file in `graded-submissions/HW1/`.

### Identical submissions

If two students hand in byte-identical files (a copied submission, or an untouched template), each is still treated as its own submission and identified by its **file name**. Each student gets their own graded file (`alice_hw1_Graded.docx`, `bob_hw1_Graded.docx`), their own working files, and their own grading and checking run. No student's file name appears in another student's grading. Only the mechanical reading of the file is shared behind the scenes, since the content is the same.

## Supported file types

| Input | Output |
|---|---|
| `.doc`, `.docx`, `.pdf` | `[name]_Graded.docx` |
| `.xlsx`, `.xls` | `[name]_Graded.xlsx` |

- **Legacy `.doc` / `.xls`** are converted to `.docx` / `.xlsx` by Microsoft Word/Excel (Windows) or LibreOffice, so formulas, formatting and images carry over into the graded copy.
- **`.pdf`** submissions are rebuilt as a `.docx`, one paragraph per line, so feedback can sit directly under the answer line. Embedded images and equations appear as pictures of the student's own work. Pages with images, garbled equation text or drawn charts are flagged and also read from a rendered picture of the whole page, so scanned or handwritten pages are still graded.
- **Embedded images** (pasted work, figures, equations stored as pictures) are exported so the agents can look at each one. There's no OCR, so image content is always read visually.
- **Charts and text boxes:**
  - A Word or Excel chart (including one on its own chart tab) is read as its data (type, title, every series).
  - A PDF chart drawn as vectors is flagged and pictured.
  - An answer typed into a spreadsheet text box is read, and kept in the graded copy.
- **Spreadsheets** are read tab by tab, hidden tabs included, with both each formula and its value.

## Reading the annotations

**Documents (.docx output):**
- Red text is inserted immediately below each incorrect/incomplete answer, formatted as **INCORRECT**: [the correct answer, the working behind it, and what specifically went wrong] or **INCOMPLETE**: [what's missing].
- `Grading Completed` appears in red at the end of the document once the grader is done.
- `Review Passed` or `Review FAILED` is added in **blue** by the checker, as its own paragraph right after "Grading Completed". For a "Review FAILED", every discrepancy is listed in blue right below the mark, in the same file.

**Spreadsheets (.xlsx output):**
- Every wrong cell is marked on its own (see [the rule below]). The **cell itself** is highlighted in Excel's standard "Bad" style (light red fill, dark red font). The value/formula is never changed, only its formatting.
- An explanation is written in red into the closest empty, visible cell (right, then below, then further out). Existing cell content is never overwritten, and hidden rows/columns are skipped.
- A dedicated **"Grading Summary"** tab (the last sheet) carries `Grading Completed` (red) and, once checked, `Review Passed` / `Review FAILED` (blue), plus a table of any discrepancies.

**Student-facing wording:** annotations never mention the rubric, solution or answer key, because students have no access to them. Each annotation states the correct answer and working directly. The annotation script refuses such wording, and the checker's audit flags any that gets through.

**No total score:** no agent calculates or writes a total score/grade. Only per-question (per-cell for spreadsheets) verdicts and explanations are recorded; totaling a grade is left to you.

### Spreadsheets: graded cell by cell

A spreadsheet is graded cell by cell, and every cell is judged by its **formula**, not just its number:
- **A cell with a formula** is marked **INCORRECT** only if the formula itself is wrong. If the formula is right but the number is wrong, the error came from an upstream cell, and the cell is **not** marked.
- **A cell without a formula** (a typed-in value) is judged by its value: a wrong value is marked **INCORRECT**.
- **An empty cell** that should hold an answer is marked **INCOMPLETE**.
- **Every** wrong cell is marked, even when several sit in the same question or section.

Which cells are graded is settled once in the verified answer key. Instructor notes and labels are marked "not graded", and every other solution cell counts. This is decided in code, not by the agent's judgment. `compare_xlsx.py` re-runs each student formula on the solution's correct inputs to see whether the formula itself is right, and lists which cells to mark and which carry an upstream error. The annotation script refuses grading that breaks the rule, and the checker's audit re-checks it. The agents judge only what code can't: text answers, a note the student left out, an unusual formula or a rearranged sheet.

### Conceptual questions

For conceptual, short-answer and essay questions (written-sentence answers), the system checks the student's answer against every key point the answer key lists for that question, not just its overall direction. If the answer is coherent but misses a specific point, it's marked **INCOMPLETE** and every missing point is named explicitly (never a vague "explanation incomplete").

## How to use it

### Prerequisites
- Claude Code with access to this repository.
- Python 3.9+ and a one-time environment setup: `python scripts/setup_env.py`.
  - It creates a private virtual environment in `.venv/` and installs the toolkit's dependencies from `requirements.txt` (including `pywin32` on Windows, for the Word/Excel conversion). Nothing is installed into your system Python.
  - `.venv/` is excluded from git and Dropbox sync, so each machine builds its own. Re-run the command after `requirements.txt` changes.
- For legacy `.doc` / `.xls` files, one of:
  - **Microsoft Word / Excel** (Windows), used automatically when installed; or
  - **LibreOffice**, found on PATH or in its standard install location.
  Without either, `.doc` files are flagged unreadable rather than guessed at, and `.xls` files are graded from cell values only (clearly marked **DEGRADED**, since formulas are then unavailable). Everything else needs no extra software.

### 1. Add your files
- Put the answer key in `reference-solutions/`.
- Put student submissions in `student-submissions/`.
- Optionally organize either folder into per-homework subfolders (see above).

### 2. Ask Claude to grade
Just ask, in plain language, for example:
- "grade all student submissions"
- "grade student-submissions/HW1/john_doe.docx against reference-solutions/HW1/answer_key.docx"

This runs the `/grading-instructions` skill, which will:
1. Inspect the folder structure and match each submission to its solution file, asking you to confirm if a match is ambiguous.
2. Build and independently verify the answer key for each solution file. A key that's already verified is reused, and if a key can't be verified after 3 correction rounds, you're asked how to proceed. If anyone questions the official solution, that homework pauses until you decide.
3. Run the grader agent on each submission.
4. Run the checker agent once a submission is marked "Grading Completed."
5. Before the final summary, check every graded file against the current answer key and regrade any that are outdated.

### 3. Review the output
- Open the corresponding file in `graded-submissions/`.
- Read the red annotations for what was marked wrong and why.
- Check the blue mark at the end (or on the "Grading Summary" tab) for the checker's final verdict.
- If it says **Review FAILED**, read the discrepancies listed in blue right below the mark in that same file, and decide manually how to proceed. The system will not auto-correct or re-grade on its own.
- Read Claude's end-of-run summary too. It reports any unreadable submissions (no graded file is created for those), any DEGRADED `.xls` grading, and any graded files that are outdated because the answer key changed. Concerns about the official solution are raised as soon as they come up, before grading continues.
- Avoid keeping a graded file open in Word or Excel while grading is still running: its checker writes the review mark into that same file.

## Important: AI-assisted grading

The rubric checker, grader and checker are independent, but all of them are AI agents interpreting free-form work. Treat "Review Passed" as a strong second opinion, not an infallible verdict. Spot-check a sample of graded files yourself, especially for high-stakes grading, ambiguous or creative answers, or anything the checker flagged as ambiguous. Glancing over the answer key (`.cache/rubrics/<id>/rubric.md`) before a large batch is also worthwhile, since every submission is graded against it.

Known limitations:
- Images and handwriting are read visually by the agents, never by OCR, so check answers that exist only as a picture with extra care.
- The data inside embedded objects (e.g. an Excel sheet pasted into a Word file) isn't read; only its preview picture is.
- A greyscale diagram drawn inside a PDF is flagged and read from the page picture, but isn't reproduced in the rebuilt graded copy.
- The graded copy of a PDF is a rebuilt Word document, so its layout won't match the original exactly.

## Where to look for more detail

- [`.claude/skills/grading-instructions/SKILL.md`](.claude/skills/grading-instructions/SKILL.md) — orchestration workflow, folder-matching rules, answer-key loop, discrepancy types
- [`.claude/agents/rubric-builder.md`](.claude/agents/rubric-builder.md) / [`rubric-checker.md`](.claude/agents/rubric-checker.md) — how the answer key is built and verified
- [`.claude/agents/grader.md`](.claude/agents/grader.md) — full grader logic and feedback standards
- [`.claude/agents/grading-checker.md`](.claude/agents/grading-checker.md) — full checker logic
- [`.claude/skills/grading-instructions/extraction-fallback.md`](.claude/skills/grading-instructions/extraction-fallback.md) — manual extraction and image-inspection methods
- [`scripts/README.md`](scripts/README.md) — the toolkit: what each script does, the cache, how to extend it
- [`CLAUDE.md`](CLAUDE.md) — quick-reference project rules
