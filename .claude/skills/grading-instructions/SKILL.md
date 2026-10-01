---
name: grading-instructions
description: grading instructions to grade students' homework submissions
---

# Concurrent Grading Workflow

Orchestrates the whole run:
- a **verified answer key (rubric)**, built once per solutions file;
- parallel **grader agents**, which annotate only incorrect/incomplete answers, in red;
- independent **checker agents**, which do a single verification pass with no corrections, in blue.

**Key Requirements**:
- **Only edit files in `graded-submissions/`.** The shared cache under `.cache/` (extraction, rubrics, per-file work dirs) is the one exception; it is derived, git-ignored data.
- All content extraction goes through the shared toolkit in `scripts/`. Agents read its compact **views**, not raw JSON.
- Supports .doc, .docx, .pdf, .xlsx and .xls, for both submissions and solutions.
- Single verification pass (no correction loops) for grading.
- **No scoring**: no agent calculates or writes a total score/grade (e.g. "8/10", "80%", a letter grade). Agents give only per-question verdicts (per-cell for workbooks) and explanations; scoring is left to the human instructor.

**Prerequisite (once per machine)**:
- If `.venv/` doesn't exist, run `python scripts/setup_env.py`. It creates the private environment and installs `requirements.txt`.
- `.doc` and `.xls` files are converted to `.docx` / `.xlsx` by Microsoft Word/Excel (Windows, via COM) or LibreOffice, whichever is installed.
- From then on, all Python runs with `.venv/Scripts/python` (Windows; `.venv/bin/python` on macOS/Linux), never a bare `python`/`pip`. Include this rule in every agent prompt.

**What the toolkit does for every agent** (details in `scripts/README.md`):
- `extract.py <file>`: cached by content hash, so a file is extracted once no matter how many agents read it. It prints:
  - `view_path`: a compact one-line-per-paragraph/cell view whose ids are the annotation anchors. Images are exported with their paths; charts, text boxes and vector PDF graphics are shown or flagged.
  - `work_dir`: that file's own scratch folder (per file path, so identical submissions never share one)
  A `.pdf` is rebuilt once, in code, as the `.docx` its graded copy is written from.
- `rubric.py`: the per-solutions-file answer key, and its gate. It serves the rubric only while it is verified, unchanged and not on hold. It also handles `hold`/`release` for solution concerns, and `graded` lists graded files made with an older answer key.
- `compare_xlsx.py`: the **cell rule** for workbooks, applied in code. It re-runs every student formula on the solution's inputs and lists three groups:
  - cells to **mark**: a wrong formula or wrong typed value is INCORRECT, an empty answer cell INCOMPLETE;
  - cells **not** to mark: a right formula whose value is wrong only because an upstream error carried over;
  - cells to **judge** (text answers, unusual formulas, shifted layouts).
  It uses the rubric's tolerances. `annotate.py` refuses verdicts that break the rule, and `audit_graded.py` re-checks it.
- `annotate.py` / `mark_review.py`: write the grader's and checker's marks with the pinned colors and wording. Agents never hand-write python-docx/openpyxl code for this. `annotate.py` also records in the graded file which answer key it used, and tags each annotation with its question.
- `audit_graded.py`: the mechanical half of the checker's audit, run in two steps:
  - a status check first;
  - after the checker commits its blind verdicts, the report: the grader's verdicts per question (per cell for workbooks), an automatic comparison with the checker's, and every mechanical problem (mark, placement, colors, highlights, altered content, outdated answer key).
- `.claude/skills/grading-instructions/extraction-fallback.md`: manual extraction and image-inspection methods. Agents open it only when a view is wrong or flags content outside the text.

**Annotation convention** (applied by `annotate.py`; full rules in `grader.md`):
- Documents: red feedback text immediately below the wrong answer.
- Workbooks, graded **cell by cell**: every wrong cell is highlighted in Excel's "Bad" style (`FFC7CE` fill / `9C0006` font), with a red explanation in the closest empty cell. A cell is wrong only if its own formula is wrong, it holds a wrong typed value, or it is empty. A right formula whose number is off only because of an upstream error is never marked; the root cell's explanation names the cells that inherit the error.
- **Annotations are student-facing**: they state the correct answer and working directly and never mention the rubric, solution or answer key, which students can't see. `annotate.py` refuses such wording, and `audit_graded.py` reports it as a `format_error`.
- **Colors and mark wording are pinned exactly**: grader red `FF0000` (annotations and the `Grading Completed` mark), checker blue `0000FF` (`Review Passed` / `Review FAILED` and discrepancy text), with no variants or decoration.

**Conceptual/explanation questions**: the rubric lists every key point each answer must cover. Graders and checkers check a student's answer against every one and name each missing point. Everyone uses the same list, so verdicts are consistent across students.

## Step 0: Discover Submissions & Match Solutions (before invoking any agent)

`student-submissions/` and `reference-solutions/` may each be flat, or organized into per-homework subfolders (e.g. `HW1/`, `HW2/`). They aren't always organized, and not always the same way on both sides. Before invoking anything for a submission, resolve which solutions file grades it:

- **Flat submission** (directly in `student-submissions/`): pair it with the solutions file directly in `reference-solutions/`.
  - **If more than one flat solutions file exists**, filename alone won't tell you which applies. Open each candidate and match by subject matter (title/case name, file type, worksheet tab names, etc.). Proceed only once one candidate is clearly the same assignment. If two or more still look equally plausible, stop and ask the user.
- **Subfoldered submission** (in `student-submissions/<X>/`): pair it with the solutions file(s) in `reference-solutions/<X>/`. **The subfolder name must match exactly.** `HW1/` does not match `Homework 1/` or `hw1_solutions/`.
- **Mixed layouts** are expected. Resolve each submission independently by the same rule.
- **No confident match**: do NOT guess. Stop and ask the user to confirm or provide the correct solutions file before grading anything in it.
- **Output mirrors the input structure**:
  - `student-submissions/HW1/name.docx` → `graded-submissions/HW1/name_Graded.docx` (create the subfolder if needed)
  - flat submissions → directly into `graded-submissions/`

## Step 1: Answer Key (once per solutions file, before any grader starts)

For each distinct solutions file matched in Step 0:

1. Run `.venv/Scripts/python scripts/rubric.py status <solutions_file>`.
   - **`verified`**: reuse it (e.g. from an earlier batch) and go straight to grading.
   - **`new`** or **`stale`** (the map, the rubric or the extraction version changed after approval): start the loop below at step 2 (build).
   - **`unverified`** (built, never reviewed — e.g. an interrupted run): start at step 3 (verify).
   - **`rejected`**: start at step 4 (correct), using the `last_issues` path that `status` prints.
   - **`on_hold`**: a solution concern is waiting for the user — go to **Holds** below.
2. **Build**: invoke the **rubric-builder** agent with `solutions_file`. It writes the question→content map and runs `rubric.py build`, which copies the solution **verbatim**. For a workbook, the map is a cell-level key: its sections restate no answers, tolerances are machine-readable, and `ungraded_blocks` names the notes and labels that are not graded. So which cells count is settled once, verified, and not left to each grader.
3. **Verify**: invoke a **fresh rubric-checker** agent with `solutions_file`. This must be a new instance, never the builder. It compares the rubric against the raw solutions and then does one of:
   - `rubric.py approve <review.json>`: needs its per-question review record;
   - `rubric.py reject <issues.json>`;
   - `rubric.py hold <concerns.json> <review.json>`: the rubric is faithful, but the official solution looks wrong.
4. **If rejected, the issues must be corrected before any grading.** Invoke the rubric-builder again with `solutions_file` and the `issues_path` that `reject` printed. It fixes the map and rebuilds. Then go back to step 3 with another fresh rubric-checker.
5. **Round limit**: when `reject` prints `escalate_to_user: true` (3 rejected rounds), stop. Show the user the latest issues and ask how to proceed. Don't grade that homework until the rubric is verified.

**Gate:** never launch a grader or checker for a homework until `rubric.py path <solutions_file>` succeeds. The graders and checkers enforce this too, refusing to grade against an unverified, stale or held rubric.

Rubrics for different homeworks can be built and verified in parallel (at most 10 concurrent agents per type).

### Holds: a concern about the official solution pauses that homework
A concern comes either from the rubric-checker's `hold`, or from a grader or checker that stopped and reported `RUBRIC CONCERN`.
1. If a grader or checker raised it, write their concern(s) to `concerns.json` in the rubric directory (`{"concerns": [{"question", "detail", "source"}]}`) and run `rubric.py hold <solutions_file> <concerns.json>`. This stops every other grader and checker from starting on that homework. If the rubric is already on hold, the same command adds the new concern(s) to the outstanding ones (`status` shows the file as `last_concerns`).
2. Launch no more agents for that homework, and let running ones finish.
3. Show the user each concern and ask for a decision:
   - **The solution stands as written** → `rubric.py release <solutions_file> --note "<the user's decision>"`. Resume grading.
   - **The solution is wrong / another answer is also acceptable** → the user corrects the solutions file (it then gets a new rubric, from Step 1), or the rubric-builder records the user's ruling in `grading_notes` and rebuilds. Either way a fresh rubric-checker re-verifies before grading resumes.

### After any answer-key change: find outdated grading
Every graded file records which verified rubric it was graded against. Whenever a rubric is rebuilt after grading has begun, and always before your final summary, run `rubric.py graded <solutions_file> <graded folder>` for each homework.
- **`outdated`** files were graded against an older answer key, or against an earlier version of a solutions file that has since been corrected in place. Regrade them: delete the old graded file and run the grader again.
- If an outdated file already carries a checker review, ask the user first, since deleting it discards that review.
- **`no_record`** files were made outside this workflow; tell the user.

## Concurrent Workflow Overview

Both agent types work independently on different files:
- **Grader**: processes the submission queue, one submission per agent invocation.
- **Checker**: processes submissions that have their "Grading Completed" mark, one per agent invocation.
- Graders and checkers may run concurrently, capped at **10 concurrent instances per agent type**: up to 10 graders and up to 10 checkers at once, never more than 10 of either. With more than 10 submissions queued, launch the first 10, then launch the next as each earlier one finishes. The same pattern applies to checkers.

## Grader: Grade Submission

**Invoke** for each submission with:
- `submission_file`
- the matched `solutions_file` (Step 0)
- `output_dir`: `graded-submissions/`, or its matching homework subfolder

The process is in `grader.md`. Grading is final; there is no correction mode. If the grader reports the submission **unreadable**, no graded file is created. Tell the user which file and why, including any image-content note. Also pass on any `DEGRADED` `.xls` report (graded from values only because no Excel/LibreOffice converter was available). If a grader reports `CELL RULE CONCERN` (it believes `compare_xlsx.py` classified a cell wrongly), no graded file was written: show the user the cell and the grader's reasons, and ask how to proceed.

## Checker: Single-Pass Verification

**Invoke** once a submission has its "Grading Completed" mark, with:
- `submission_file`
- `graded_file` (the `_Graded.docx` / `_Graded.xlsx`)
- `solutions_file`

The process is in `grading-checker.md`:
- grade blind against the same rubric;
- commit those verdicts to a file;
- only then produce the audit report and compare against the grader's annotations.
The scripts enforce that order, and `mark_review.py` won't mark a file without the blind audit. If a checker reports `RUBRIC CONCERN`, handle it as a hold (above). A `CELL RULE CONCERN` from a checker goes to the user the same way as a grader's; the file stays unmarked.

**Final verdict** (single pass — never sent back to the grader):
- **`Review Passed`** (blue `0000FF`, exact text): grading complete and verified ✓
- **`Review FAILED`** (blue `0000FF`, exact text), immediately followed, in the same graded file, by blue text stating every problem (no separate file): the user reviews manually and decides next steps.

### Discrepancy Types & Severity

| Type | Meaning | Severity |
|------|---------|----------|
| `verdict_mismatch` | Grader/checker disagree on correct vs. not correct (a question; a cell for workbooks) | high |
| `outdated_rubric` | The file was graded against an answer key that is no longer the verified one (or an earlier version of the solutions file) | high |
| `missed_question` | Grader skipped a question | high |
| `annotation_placement` | Feedback not immediately below its answer (docx) / not the closest empty visible cell (xlsx) | medium |
| `explanation_error` | Explanation unclear/incomplete/wrong | medium |
| `incomplete_coverage` | Multi-part question not fully addressed | medium |
| `cell_highlight_missing` | (.xlsx) A cell the cell rule says to mark is unmarked or has the other label, a correct / carried-over / `(not graded)` cell is marked, or a highlight is wrong or has no explanation beside it | medium |
| `missing_keypoint_not_flagged` | Grader missed a key point absent from a conceptual answer, or flagged one that's actually present | medium |
| `format_error` | Wrong color/mark wording, original student content altered, a missing grading record/question tag, or an annotation that mentions the rubric/solution/answer key | medium |
| `label_mismatch` | Both say "wrong", but one says INCORRECT and the other INCOMPLETE | low |
