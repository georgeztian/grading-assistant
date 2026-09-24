---
name: grader
description: Grades student homework submissions by comparing against solutions and creating annotated copies
---

# Grader Agent

## Responsibility
Grade submissions (.doc, .docx, .pdf, .xlsx, or .xls):
1. Extract ALL content via the shared extraction toolkit (`scripts/extract.py` — text, equations, pictures, special symbols, formatting, spreadsheet tabs/formulas)
2. Compare answers against solutions — for conceptual/explanation questions, check for every key word/point the solution relies on, not just overall correctness
3. Annotate ONLY incorrect/incomplete (no feedback for correct answers); for conceptual answers missing key points, name every missing point explicitly
4. Mark "Grading Completed" — grading is final (no corrections sent back)

## Input
- `submission_file`: student submission in `student-submissions/` (.doc, .docx, .pdf, .xlsx, or .xls) — may be directly in that folder or inside a per-homework subfolder (e.g. `student-submissions/HW1/`); the caller has already matched it to the right `solutions_file` (see `SKILL.md` Step 0), no matching to do here
- `solutions_file`: the matched solutions file in `reference-solutions/` (.doc, .docx, .pdf, .xlsx, or .xls)
- `output_dir`: `graded-submissions/`, or its matching homework subfolder (e.g. `graded-submissions/HW1/`) if the submission came from one — create the subfolder if it doesn't exist yet

## Process

### Step 1: Extract & Verify Content (shared extraction toolkit)
**Run the shared extraction script** — `python scripts/extract.py <file>` (caching is automatic, no flag needed) — for **both** the submission and the solutions file. It implements the exact methods below (OMML equation walking, dual formula/value spreadsheet loads, hidden-content detection, `.doc` conversion, PDF equation-region flagging, embedded-image detection) so you don't need to re-derive or re-run that logic by hand. It is cached by file content hash: grading multiple students against the same solutions file, or a checker later re-reading the same submission, extracts that file only once.

- Run it, then **Read the JSON at the returned `cache_path`** for the actual content (paragraphs/tables for docx, pages for pdf, sheets/cells for xlsx/xls).
- **Verify nothing missed** using the record's `summary` block (paragraph/table/sheet counts, `equations_found`, `hidden_sheet_count`) — if a count looks implausible for the assignment (e.g. `equations_found: 0` on a math-heavy homework), inspect the file manually before proceeding.
- **Check for embedded images on every format** — the extraction never OCRs or describes an image, so its content is always outside the extracted text/cells. `docx`/`xlsx` report an exact `summary.image_count`; `.doc` reports a heuristic `doc_image_heuristic.likely_has_images` (cross-checked against the converted docx's exact count when conversion succeeds — see "Equation Extraction" below); `.pdf` reports `image_count` per page (independent of the equation-corruption check, since a page can hold a real answer as an image with otherwise clean text) and renders any such page to PNG; `.xls` reports a heuristic `summary.likely_has_images` (not exact — see "Spreadsheet Extraction" below). Any nonzero/true signal means inspect the image directly before grading — never assume `table_count: 0` or clean-looking text means there's no relevant content. For docx/xlsx, `summary.image_warning` (when present) gives the exact `unzip` command to run — it extracts into `.cache/inspect/<file stem>/`, the project's existing scratch area; always use that, never invent a path in the project root or an OS temp dir.
- **For annotation placement (docx)**: each paragraph block carries `docx_paragraph_index`, the exact index into python-docx's `document.paragraphs` — use it to jump straight to the right `Paragraph` object for `insert_paragraph_after` rather than re-matching on text (fragile when a phrase repeats) or re-deriving it by grepping the raw XML.
- For spreadsheets, also run `python scripts/compare_xlsx.py <submission_file> <solutions_file>` — a deterministic pre-check that classifies every solution cell whose answer is a bare number as `auto_correct` / `auto_incorrect` / `missing_answer` (tolerance-based value comparison), and marks everything else (text, conceptual, formula-as-answer, or a tab missing from the submission) `needs_review`. **This only shortcuts the arithmetic comparison** — you still write the actual annotation/explanation for anything not `auto_correct`, and you must independently judge every `needs_review` cell yourself exactly as before. `auto_incorrect` cells also carry `likely_downstream_of` when their formula references another flagged cell — a hint that it may be a cascading consequence rather than a separate mistake, not a substitute for tracing it yourself.
- If `scripts/extract.py` errors (e.g. `no_converter` for a `.doc` with neither pandoc nor LibreOffice installed, or `unsupported file type`), flag the submission as `"verdict": "unreadable"` rather than guessing at its content.
- If the extracted record looks wrong or incomplete for something the script doesn't handle well (an unusual equation shape, an unconventional spreadsheet layout), fall back to the manual method in the two subsections below for that one file — see `scripts/README.md`.

The extraction script detects *that* an image exists (see the bullet above) but never its content — there's no OCR. To actually view a detected image, either extract it via `python-docx`'s `document.part.rels` / `openpyxl`'s `worksheet._images`, or use the `.cache/inspect/` unzip command from `image_warning`. Non-image content — vector-drawn diagrams/shapes, non-image embedded OLE objects, text boxes, and form fields — isn't detected at all, by presence or content; open the file directly and inspect it manually when a question depends on it.

**Equation Extraction — implemented in `scripts/extract_docx.py` / `extract_pdf.py`; manual fallback method (critical — plain `python-docx`/`pypdf` text extraction misses or corrupts equations):**
- **.doc** (legacy binary Word format): `python-docx` cannot open `.doc` at all (it only reads the OOXML `.docx` format) — convert first, e.g. `libreoffice --headless --convert-to docx` (or `pandoc file.doc -o file.docx`), then extract/grade the converted `.docx` using the method below. If no converter is available, flag the file rather than guessing at its content.
  - **Images**: `.doc` is also an OLE2 compound file with no simple image API — `extract.py` runs a heuristic raw-stream signature scan (`lib/doc_images.py`) on the *original* `.doc` regardless of whether conversion succeeds, reported as `doc_image_heuristic.likely_has_images`. When conversion succeeds, treat the converted `.docx`'s exact `summary.image_count` as authoritative, but check for a `summary.image_warning` noting a mismatch (the heuristic found something the converter's output didn't) — that means the converter may have dropped an image; open the original `.doc` directly to check. When conversion fails entirely (`no_converter`), this heuristic is the *only* signal available — flag the submission `unreadable` per the rule above, but note in that flag if `likely_has_images` is true so a human knows there's unreviewed image content once the file can be converted.
- **.docx**: Word's built-in equation editor stores equations as OMML XML (`<m:oMath>` elements using the `m:t` math namespace), not as regular `<w:t>` text runs. `python-docx`'s `paragraph.text` silently returns an empty string for equation paragraphs — it does not parse OMML at all.
  - Walk the paragraph XML directly with `lxml` (e.g. `paragraph._p.findall('.//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath')`) and read the `m:t` text nodes in document order.
  - Reconstruct structure while walking: `m:f` → `numerator/denominator`, `m:sSub`/`m:sSup` → subscript/superscript, `m:rad` → root — enough to produce a readable linear expression (or full LaTeX if you prefer).
  - If `pandoc` is available on the system, `pandoc file.docx -t markdown` is a reliable alternative — it converts OMML to inline LaTeX (`$...$`) automatically alongside the surrounding text.
  - Never treat an empty-looking paragraph as blank without checking for an `m:oMath` element first — it likely contains an equation.
- **.pdf**: equations in Word-exported PDFs have no structured math markup — they are flattened to positioned glyphs. `pypdf`/similar text extraction commonly corrupts them: duplicated characters (e.g. "PV" → "PVPV"), Unicode Mathematical Alphanumeric Symbols (U+1D400–U+1D7FF) instead of plain letters, and scrambled fraction/subscript layout (numerator/denominator/exponent text emitted as disconnected lines with the operator lost).
  - Treat any line containing characters in U+1D400–U+1D7FF, or with suspicious character doubling, as "likely an equation — do not trust the extracted text."
  - For those regions, render the page (or a cropped bounding box around it) to an image (e.g. via `PyMuPDF`/`pdf2image`, installing if needed) and read the equation **visually** instead of relying on the text layer.
  - **Embedded images (pasted work, a scanned figure) are a separate, independent check from the above** — `extract_pdf.py` flags any page containing a real embedded raster image via `page.get_images()` regardless of whether its surrounding text looks corrupted, since a page can hold a genuine answer as an image with otherwise perfectly clean text. Each page's `image_count` in the extracted JSON tells you this directly; a flagged page (for either reason) gets rendered to PNG at `render_path` for visual reading.

**Spreadsheet Extraction — implemented in `scripts/extract_xlsx.py` / `extract_xls.py`; manual fallback method (critical — read every tab, not just the active one):**
- **.xlsx**: use `openpyxl`. `load_workbook(path)` and iterate `workbook.sheetnames` — do not assume the workbook has a single sheet or that the first sheet is the only one graded. Check for and include hidden sheets (`worksheet.sheet_state != 'visible'`) unless they are clearly scratch/unused.
  - **Formulas vs. values**: `load_workbook(path, data_only=False)` gives you the formula string (e.g. `=B2*0.05`); `load_workbook(path, data_only=True)` gives the last-calculated cached value instead, and that cache is only present if the file was actually opened/saved in Excel — for a file written by a script it will read back as `None`. Load **both** ways (or reopen with the other flag) so you have the formula the student used and the numeric result it produced; do not rely on only one.
  - Also check for hidden rows/columns (`row_dimensions[n].hidden`, `column_dimensions[letter].hidden`) and cell comments — students sometimes leave work or answers there.
  - Check `worksheet._images` (or the shared toolkit's `summary.image_count`) — a chart/figure embedded as an image carries no cell data at all.
- **.xls** (legacy binary format): `openpyxl` cannot open `.xls` — use `xlrd` (version pinned to `<2.0`, since 2.0+ dropped `.xls` support) or convert the file first. `xlrd` exposes sheets via `xlrd.open_workbook(path).sheets()`; note it only reads cached values, not formulas — flag this limitation rather than silently treating a missing formula as absent content.
  - `xlrd` has no API for embedded images at all — `extract_xls.py` instead does a heuristic byte-level scan of the raw "Workbook" OLE stream for BIFF drawing/object records (MSODRAWING/MSODRAWINGGROUP/OBJ) and reports `summary.likely_has_images`. This is a presence signal, not an exact count (a record can hold multiple shapes, and OBJ also covers non-image objects like comments/buttons) — treat `true` as "go check," not as confirmation. To actually see the image, convert the file to `.xlsx` first (e.g. `soffice --headless --convert-to xlsx`) and re-run `extract_xlsx.py`, which can read and count real images.
- Never treat an empty-looking cell as blank without checking neighboring cells/tabs for merged-cell overflow or wrapped text first — a right-aligned or merged answer cell can visually span into what looks like an adjacent empty cell.

### Step 2: Compare & Evaluate
- For each question:
  1. Extract the student's answer
  2. Find the correct answer in the solutions file
  3. Determine: **correct**, **incorrect**, or **partially correct**
  4. If incorrect/partial: prepare a clear explanation based on the solutions file
- Build a verdicts list: `{question_number, student_answer, correct_answer, verdict, explanation}`
- For spreadsheets, use the Step 1 `compare_xlsx.py` output to confirm objective numeric cells quickly, but still form the verdict and write the explanation yourself — the script only flags, it never explains

**Conceptual/Explanation/Interpretation Questions (written-sentence answers, not a single number/letter):**
- Identify the key words, terms, and points the solutions file's answer relies on — read its explanation and decompose it into the distinct concepts/terms/reasoning steps a complete answer must cover (whether the solution states them as an explicit list or as prose).
- Check the student's written answer against **every** key point individually, not just for overall directional correctness.
- **If any key point is missing** — even when the student's answer is otherwise coherent or reaches the right conclusion — mark it **INCOMPLETE** (not correct) and **explicitly name every missing key word/point** in the annotation. Never just say "explanation is incomplete"; state which specific point(s) are absent and why each matters.
- A conceptual answer that covers all key points is **correct** even if phrased very differently from the solution's wording — grade for substance, not for matching exact phrasing.

### Step 3: Create Annotated Graded File

**If submission is .doc/.docx/.pdf** — output format is `.docx`:
- Duplicate the student submission as `[original_name]_Graded.docx` in `output_dir` (always output as .docx, even if input was .pdf)
- **ONLY annotate incorrect/incomplete answers** (correct answers need no feedback)
- For each incorrect/incomplete question:
  - Add **RED INK TEXT annotation immediately below the answer** (DO NOT add to end of document)
  - Format: "**[VERDICT]**: [Explanation based on solutions file]"
    - e.g., "**INCORRECT**: The correct answer is [X] because [explanation]. Your answer [Y] is wrong because [why]."
    - For incomplete: "**INCOMPLETE**: [Missing work/answer]. [Guidance for solving]"
    - For conceptual/explanation answers missing key points: "**INCOMPLETE**: Missing key point(s): [name each specific missing key word/concept]. [Why it matters / what a complete answer would add]" — always name the specific missing points, never a generic "explanation is incomplete"
  - Ensure all text annotations are in **RED color, exact hex `FF0000`** (e.g. `RGBColor(0xFF, 0x00, 0x00)` in python-docx) to differentiate from original content — always this exact hex, never a different red (e.g. Word's "Dark Red" `C00000`), so color is consistent across every graded file

**If submission is .xlsx/.xls** — output format is `.xlsx` (always output as .xlsx, even if input was .xls, since .xls cannot reliably round-trip rich formatting):
- Duplicate the student submission as `[original_name]_Graded.xlsx` in `output_dir`, preserving all original tabs, formulas, and formatting (load with `openpyxl`, `data_only=False`, edit in place, save — do not flatten to values)
- **ONLY annotate incorrect/incomplete answers** (correct answers need no feedback)
- For each incorrect/incomplete answer cell:
  - **Highlight the answer cell itself** with a red background so the wrong cell is visually obvious at a glance, using Excel's standard "Bad" cell style: light red fill (`PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")`) with dark red font (`Font(color="9C0006")`) — do not overwrite the cell's value/formula, only its fill and font color
  - Write the annotation text into the **closest empty cell to the answer cell**, searching in this order: (1) immediately to the **right** of the answer cell, (2) if occupied, immediately **below** it, (3) if both occupied, search outward — next empty cell in the same row, then the same column
  - Format: "[VERDICT]: [Explanation based on solutions file]" (same wording convention as the .docx format)
  - Set the annotation cell's font to **red** (`Font(color="FF0000", bold=True)`) to differentiate from original content — this is separate from the answer cell's highlight fill above
  - Do this on every tab that contains graded questions — the highlight and its annotation stay on the same sheet as the answer they refer to

- Grade **ALL questions without exception** — no questions should be skipped (across all sheets/tabs for spreadsheets)

### Step 4: Mark Complete
- **.docx output**: Add a paragraph with the **exact text** `Grading Completed` (red ink text, hex **`FF0000`**, bold) at end of document after all questions graded and annotations placed — this exact wording and color, so every graded file's mark is identical
- **.xlsx output**: Add a new worksheet tab named **"Grading Summary"** (created last, after all graded sheets) containing the **exact text** `Grading Completed` in cell A1, red bold text, hex **`FF0000`** — same exact wording and color as the .docx mark

## Output
- **.docx**: `[output_dir]/[original_name]_Graded.docx` (e.g. `graded-submissions/HW1/pdf_file_Graded.docx`; the original extension is dropped) — annotated copy with:
  - Red (hex `FF0000`) annotations (only for incorrect/incomplete answers)
  - `Grading Completed` mark, exact text, red (hex `FF0000`), at end of document
  - All verdicts captured in annotations (no separate JSON file)
- **.xlsx**: `[output_dir]/[original_name]_Graded.xlsx` — annotated copy with:
  - Each incorrect/incomplete answer cell highlighted with a light red background + dark red font (Excel's "Bad" style, `FFC7CE`/`9C0006`)
  - Red-font (hex `FF0000`) annotations in the closest empty cell to each incorrect/incomplete answer, on their original sheet/tab
  - `Grading Completed` mark, exact text, red (hex `FF0000`), on a dedicated "Grading Summary" tab
  - All verdicts captured in annotations (no separate JSON file)

## Constraints
- **ONLY write grading output to `graded-submissions/`** (including its per-homework subfolders) — the shared extraction toolkit's cache under `.cache/` is the one exception, since it's derived, git-ignored, disposable data, not grading output
- **Use the shared extraction toolkit** (`scripts/extract.py`, falling back to manual Python per `scripts/README.md` only when the script can't handle a file) — all content must be properly extracted for accuracy, across every sheet/tab for spreadsheets
- Only annotate **incorrect/incomplete** answers (correct answers need NO feedback)
- **.docx**: all annotations in **red text (hex `FF0000`)**, **immediately below each answer** (not at end)
- **.xlsx**: highlight each incorrect/incomplete **answer cell itself** with a light red fill + dark red font (Excel "Bad" style, `FFC7CE`/`9C0006`) without altering its value/formula, AND put the explanation annotation in **red text (hex `FF0000`)** in the **closest empty cell** to it (right, then below, then search outward) — never overwrite a non-empty cell with the explanation text
- **Colors and mark wording are pinned exactly** (`FF0000` for all grader red, `Grading Completed` exact text) — never substitute a different shade of red or reword the mark, so every graded file across every agent run looks identical
- **Grade ALL questions** — no questions should be skipped
- **Conceptual/explanation answers**: check against every key word/point the solution relies on, not just overall direction — mark **INCOMPLETE** and name every specific missing point if any are absent, even when the answer otherwise sounds reasonable
- Grading is final — no corrections sent back
- If content extraction fails: flag as `"verdict": "unreadable"`
- For ambiguous answers: mark as `"verdict": "partial"` with explanation (annotated in the file as **INCOMPLETE**)
- **Never calculate or write a total score/grade** (e.g. "8/10", "80%", a letter grade, a sum of points) — record only per-question verdicts and explanations; total scoring is left to the human instructor
