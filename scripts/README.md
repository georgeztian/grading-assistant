# Grading toolkit

Deterministic Python that does every *mechanical* part of grading, so the
agents spend their tokens only on judgment. It covers:

- extracting content
- building and verifying the answer key
- writing annotations and review marks
- auditing a graded file

Grading judgment, annotation wording and the independent-verification
workflow are defined in [`SKILL.md`](../.claude/skills/grading-instructions/SKILL.md),
[`grader.md`](../.claude/agents/grader.md),
[`grading-checker.md`](../.claude/agents/grading-checker.md),
[`rubric-builder.md`](../.claude/agents/rubric-builder.md) and
[`rubric-checker.md`](../.claude/agents/rubric-checker.md).

## Why this exists

Each agent has its own context window, and every turn re-sends that whole
context. So cost is roughly **context size × number of turns**. The toolkit
cuts both:

- **Extract once, read compactly.** Every file is extracted once, cached by
  content hash, and read through a compact text *view*. On real homework the
  view is 60–90% smaller than the JSON record.
- **Answer key once.** Each solutions file becomes a verified rubric: a
  complete, verbatim, question-by-question copy. Graders and checkers read
  the rubric instead of each re-reading and re-decomposing the raw solutions.
  Every grader uses the same key points, so conceptual verdicts are
  consistent across students.
- **No hand-written file code.** Annotations, marks and audits are single
  commands. Agents no longer write and debug python-docx/openpyxl code (the
  most expensive part: output tokens and many turns), and the pinned colors
  and wording are guaranteed by code.

## Setup

```
python scripts/setup_env.py
```

This creates the project's private environment at `.venv/` and installs
`requirements.txt` into it. The script is stdlib only, so any Python 3.9+
can run it, and it's idempotent: re-run it after `requirements.txt` changes.
`.venv/` is git-ignored and marked ignored for Dropbox sync.

From then on, run everything with `.venv/Scripts/python` (Windows) or
`.venv/bin/python` (macOS/Linux), including ad hoc code. Every script
imports `_venv.py` first: under any other interpreter it re-runs itself
under `.venv`'s, and it exits with the setup command if `.venv/` is missing.

Pillow is required even though no script imports it directly. Without it,
openpyxl silently skips embedded images: `image_count` reads 0 and a saved
`_Graded.xlsx` loses the student's pictures.

**Legacy `.doc` / `.xls`** are converted to `.docx` / `.xlsx` by
`convert_doc.py`, using the first converter available:
1. **Microsoft Word / Excel via COM** (Windows, needs `pywin32`, which is
   listed in `requirements.txt`). These are the native applications, so the
   conversion is exact.
2. **LibreOffice** (`soffice`, on PATH or in its standard install location).

With either one, legacy files are graded with full fidelity: formulas,
formatting and images all survive. With neither:
- a `.doc` is unreadable (flagged, never guessed at);
- a `.xls` falls back to a values-only `xlrd` reading, marked **DEGRADED**,
  which is re-extracted automatically once a converter is installed.

`pandoc` can't read either binary format.

## Scripts

### Extraction
- **`extract.py <file> [--force] [--json]`** — the one entry point for
  reading any submission or solutions file. It dispatches by extension, goes
  through the shared cache, and prints:
  - `view_path`: the compact **view**. Agents Read this.
  - `cache_path`: the full JSON record, for scripts.
  - `work_dir`: this file's own scratch folder, `.cache/work/<sha12>-<path8>/`.
    It is keyed by content *and* path, so two students' byte-identical files
    never share one. Agents write `verdicts.json`, `checker_blind.json` and
    `discrepancies.json` there, and `audit_graded.py` writes `audit.json`.
  - `summary`
  - Per format:
    - `.doc` and `.xls` are converted first and then extracted as
      `.docx` / `.xlsx`.
    - A `.pdf` is extracted for its text/flags and also **rebuilt as a base
      `.docx`** (`lib/pdf_docx.py`) whose paragraphs are the annotation
      anchors.
    - For `.doc`, a raw-OLE image heuristic (`lib/doc_images.py`)
      cross-checks the conversion.
- **The view** (`lib/views.py`) shows one line per paragraph, table row or
  cell; empty paragraphs and default fields are dropped. Its ids are used
  everywhere: view, rubric and annotation anchors.
  - `[pN]` is top-level paragraph N (python-docx `document.paragraphs[N]`).
  - `[tN]` is table N (rows `r0, r1, …`; a table cell is `tN:rRcC`).
  - `[sN]` is body-level content control N (its text would otherwise be
    invisible: python-docx skips content controls).
  - `Sheet!B5` is a workbook cell. `Sheet!textboxN` / `Sheet!chartN` are
    text boxes and charts, listed "(at <anchor cell>)".
  - Warnings (images, flagged PDF pages, heuristics, DEGRADED) are listed
    at the top. Every embedded image is exported (`lib/images.py`, with
    EMF/WMF converted to PNG, rewritten whenever its bytes change) and its
    path shown inline: `(image: <path>)` on a paragraph, or
    `image at <cell>: <path>` in a sheet header. A chart shows its data
    inline: `(chart: …)`.
- **`extract_docx.py`** walks top-level body content in document order
  (`body_blocks`, shared with `annotate.py` and `audit_graded.py` so the ids
  always agree). It:
  - inlines Word equations (`<m:oMath>`) as `[EQ: …]` via `lib/omml.py`,
    because `python-docx` returns an empty string for them;
  - flattens nested tables into their cell's text;
  - summarizes native Word charts from their chart XML: type, title and
    every series' data (`lib/ooxml_extras.py`);
  - records `has_image`, exported `images`, and an exact `image_count`.
- **`extract_xlsx.py`** loads every sheet (hidden included) twice, for
  formulas (array/CSE formulas shown as `{=…}`) and for cached values. It
  also records hidden rows/columns, merged ranges, comments, and exported
  images (with anchor cell). Text boxes and charts, which openpyxl doesn't
  expose, are read straight from the package's drawing parts
  (`lib/ooxml_extras.py`); a chart sheet (a tab holding only a chart) is
  listed with its chart and no cells.
- **`extract_xls.py`** is the no-converter fallback only: cached values via
  `xlrd` (`formula_available: false`), and a heuristic `likely_has_images`
  (`lib/xls_drawings.py`).
- **`extract_pdf.py`** extracts text per page. A page is flagged, and
  rendered to `.cache/renders/<sha256>/page<N>.png`, for any of three
  independent reasons:
  - its text looks like a corrupted equation (Mathematical Alphanumeric
    Symbols, or a doubled token such as `PVPV`; plain numbers like `2020`
    don't count);
  - it holds an embedded image;
  - it holds a vector chart or diagram (`vector_graphics`).
- The per-format `extract_*.py` scripts can be run standalone to print a
  record, but only `extract.py` writes the cache (with images exported and
  the view built).
- **`lib/pdf_docx.py`** rebuilds a PDF as a `.docx`, one paragraph per text
  line, so annotations land directly under the answer line. Embedded
  images, corrupted equations and colour/curved vector graphics (with their
  labels) are re-inserted as pictures of their on-page region, so the graded
  copy shows the student's work as it looked.
- **`convert_doc.py <file.doc|file.xls>`** runs the converter chain above.
  The Office conversion runs in a child process with a timeout (a stuck
  dialog can't hang an agent), with macros force-disabled and the file
  opened read-only. Conversions are serialized across agents by a lock
  file and cached at `.cache/converted/<sha256>/converted.<ext>` (named by
  content, never after a student's file).

### Answer key
- **`rubric.py <command> <solutions> …`** manages the per-solutions-file
  rubric in `.cache/rubrics/<sha256>/`.
  - `init`, then the rubric-builder agent writes `map.json`: question →
    block ids, plus key points, restated answers, tolerances (free-text
    `tolerance`, machine-readable `numeric_tolerance`), image notes and
    grading notes.
  - `build` validates the map, **refusing any map that leaves a non-empty
    block of the solutions unassigned** (to a question, to shared context,
    or excluded with a reason). It then writes `rubric.md` with the solution
    content copied verbatim (tolerances, including `numeric_tolerance`, are
    shown on each question's header line, so the rubric-checker sees them).
  - The status moves `new → unverified → verified | rejected | on_hold`:
    - `approve <review.json>` (by an independent rubric-checker) makes it
      `verified`. It refuses a review record that doesn't cover every
      question with every applicable check passed, or that predates the
      build.
    - `reject <issues.json>` stores the issues for the builder's correction
      round. After 3 rejected rounds it reports `escalate_to_user`.
    - `hold <concerns.json> [<review.json>]` pauses a homework over a
      suspected error in the official solution (the review is required when
      holding a not-yet-approved rubric; a hold while already on hold adds
      to the outstanding concerns). `release --note "…"` records the user's
      decision and makes it verified again.
    - Any later change to the map, the rubric file, the solutions file or
      the extraction version makes it stale (as does a missing `map.json`).
      A rebuild keeps it verified only if the map is unchanged and the
      rubric text comes out identical; otherwise it needs a fresh review.
  - `path` prints `rubric.md` and the question ids **only** while the rubric
    is verified, unchanged and not on hold. Graders, checkers, `annotate.py`
    and `audit_graded.py` gate on it.
  - `graded <graded_dir>` lists graded files whose recorded answer key is
    not the current verified one, including files graded against an earlier
    version of a solutions file corrected in place (same path), with exit 1
    if any, so outdated grading is found after a key changes.

### Grading output
- **`compare_xlsx.py <submission> <solutions> [--json]`** is a
  deterministic numeric pre-check for workbooks. It classifies each
  solution cell as:
  - `auto_correct` / `auto_incorrect` / `missing_answer`: bare-number
    answers, compared with the rubric's `numeric_tolerance` for that
    question when set, otherwise `--tolerance`;
  - `text_match`: identical text;
  - `needs_review`: everything else;
  - `sheet_missing_in_submission`: a whole solution tab is missing.
  Hints:
  - `auto_incorrect` cells are labelled `rounded` or `percent_scale` when
    that is all that differs;
  - they get `likely_downstream_of` when their formula reads another flagged
    cell (directly or through a range);
  - they get `found_at` when the solution value sits elsewhere on the
    student's sheet;
  - many values found at the same offset produce a sheet-level
    `PROBABLE LAYOUT SHIFT`.
  Output is compact by default; `--json` gives the full per-cell output.
  These are hints only: agents still judge every non-automatic cell and
  write every explanation.
- **`annotate.py <submission> <verdicts.json> <output_dir> --solutions <solutions> [--overwrite]`**
  writes `[name]_Graded.docx|xlsx` from the grader's verdicts.
  - It needs the verified rubric, and every annotation's `question` must be
    one of its ids.
  - Documents: red `FF0000` paragraphs immediately below their `pN` / `tN` /
    `sN` anchor (or inside a `tN:rRcC` table cell), then the exact
    `Grading Completed` mark.
  - Workbooks: the answer cell gets the "Bad" style (`FFC7CE` fill /
    `9C0006` font, value untouched), the red explanation goes into the
    closest empty *visible* cell (right, then below, then outward), and a
    final "Grading Summary" tab holds the mark. The student's text boxes
    and shapes, which openpyxl drops on save, are put back
    (`lib/xlsx_shapes.py`).
  - **Provenance** (`lib/marks.py`): the file records, as custom document
    properties, the solutions file (path and hash) and verified rubric it
    was graded against. Each annotation carries a hidden tag (a bookmark, or a
    defined name) naming its rubric question.
  - `.doc` / `.xls` / `.pdf` are annotated from the same converted or
    rebuilt file their view describes.
  - `verdicts.json` is deleted once the graded file is written, so the
    checker can't see it before grading blind.
  - It refuses to overwrite an existing graded file unless `--overwrite`,
    and never overwrites one that already carries a checker review.
- **`audit_graded.py <graded> <submission> --solutions <solutions> [--report --blind <checker_blind.json>]`**
  audits the graded file against the original, in two steps:
  - Without `--report` it shows only the mark status and answer-key
    problems, nothing of the grader's verdicts.
  - `--report` requires the checker's committed blind verdicts for every
    rubric question. It then writes `audit.json` (with the blind file's
    hash) containing:
    - the grader's annotations, each with its question, anchor and the
      answer text it sits under;
    - `comparison`: per question, checker vs. grader —
      `agree` / `verdict_mismatch` / `label_mismatch`;
    - `problems`: altered or deleted student content, wrong colour shade, a
      reworded or misplaced mark, a highlight without an annotation, an
      annotation not in the closest empty visible cell, a missing question
      tag, and `outdated_rubric`.
- **`mark_review.py <graded> passed | failed <discrepancies.json> --submission <submission>`**
  writes the checker's exact blue `0000FF` `Review Passed` / `Review FAILED`,
  and for a failed review, one line or row per discrepancy plus a summary
  line. It refuses:
  - without a blind audit report of the file as it is now;
  - `passed` while that report shows any disagreement or problem;
  - `failed` that omits one of the report's disagreements (a skipped
    question's `verdict_mismatch` may be listed as `missed_question`), or
    any kind of problem the report lists;
  - a file with no `Grading Completed` mark, or one already reviewed.
  Verdict discrepancies need `question` and both verdicts; a file-level
  problem (e.g. `outdated_rubric`) may omit them but must explain itself in
  `notes`.

Shared constants and helpers live in `lib/marks.py`: colors, mark wording,
verdict and discrepancy vocabularies, and reading/writing the provenance
record.

## Cache

Everything lives in `.cache/`, which is git-ignored and outside the three
data folders:

- `.cache/extraction/<sha256>.json`: one record per distinct file content,
  shared by identical submissions, so it names no file. It is
  schema-versioned (`lib/cache.py`), so a script change invalidates stale
  entries automatically.
- `.cache/extraction/<sha256>.<path8>.view.txt`: one view per file path.
  Two students who hand in byte-identical files each get their own view,
  headed with (and referring to) their own file name, and their own scratch
  folder; only the extraction work is shared.
- `.cache/converted/<sha256>/`: `.doc` / `.xls` conversions and PDF base
  `.docx` files.
- `.cache/renders/<sha256>/`: PDF page renders and exported embedded images.
- `.cache/rubrics/<sha256>/`: answer keys and their verification status.
- `.cache/work/<sha12>-<path8>/`: per-file agent scratch (verdicts, blind verdicts, discrepancies,
  audit reports).
- `.cache/inspect/`: the only place for ad hoc manual inspection (e.g. the
  `unzip` commands the view suggests). Never use the project root or an OS
  temp dir.

Everything here is derived and safe to delete. It re-extracts on next use,
but deleting `.cache/rubrics/` means each answer key has to be rebuilt and
re-verified before the next grading run. Writes are atomic and retried,
because Dropbox or antivirus can briefly lock a file on Windows.

## Extending

If a script's output looks wrong or incomplete for a file it doesn't handle
well (an unusual equation shape, a spreadsheet layout the parser
mishandles), use the manual method in
[`extraction-fallback.md`](../.claude/skills/grading-instructions/extraction-fallback.md)
for that one file. If it's a recurring pattern, fix the script so every
agent benefits.
