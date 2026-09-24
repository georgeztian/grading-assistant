# Extraction toolkit

Deterministic, tested Python extraction — the actual implementation of the
"Equation Extraction" / "Spreadsheet Extraction" methods described in
[`grader.md`](../.claude/agents/grader.md). Grader and checker agents run
these instead of writing ad hoc parsing code per submission, and instead of
each redoing the same mechanical parse independently.

## Why this exists

`grader.md` already requires the grader and checker to use the *identical*
extraction method so their results are comparable. Since the method is
identical by requirement, there is no reason to run it twice — this toolkit
makes it literally the same code, run once, and cached by content hash:

- The grader extracts a submission once; the checker reads the same cached
  record instead of re-parsing (it still forms its own verdict independently
  — only the mechanical text/formula extraction is shared).
- A solutions file used by N students is extracted once total, not 2N times
  (grader + checker, per student).

Nothing about grading judgment, annotation format, placement, or the
independent-verification workflow changes — see `grader.md` /
`grading-checker.md` / `SKILL.md` for that. This toolkit only changes *how*
content gets from a file into the agent's hands.

## Setup

```
pip install -r requirements.txt
```

`.doc` conversion additionally needs `pandoc` or LibreOffice (`soffice`) on
PATH — same as before, this is unchanged from `grader.md`'s requirement.

## Scripts

- **`extract.py <file> [--force] [--json]`** — the one entry point agents
  should call. Dispatches by extension (`.doc`/`.docx`/`.pdf`/`.xlsx`/`.xls`),
  converts `.doc` first, and always goes through the shared cache — caching
  is unconditional here (`--cache` is accepted as a harmless no-op, in case
  it's typed out of habit; it's the per-format `extract_*.py` scripts below
  that need `--cache` to opt in, since run standalone they default to a
  plain one-off dump). Prints `{cache_hit, cache_path, summary}` — read the
  JSON at `cache_path` (e.g. with the Read tool) for the actual extracted
  content, rather than passing `--json` and dumping a possibly-large record
  into the same turn. For `.doc`, it also runs `lib/doc_images.py`'s raw-OLE
  signature scan on the *original* file and attaches it as
  `doc_image_heuristic` — as a cross-check against the converted `.docx`'s
  exact `image_count` when conversion succeeds (mismatch → a warning is
  added, in case the converter dropped an image), and as the only available
  signal (in the printed error, uncached) when no converter is installed at
  all and conversion fails outright. Verified against real Word-produced
  `.doc` files (both with and without an embedded image) for all three
  cases: converter available + image present, converter available + no
  image, and no converter installed.
- **`extract_docx.py`** — walks paragraph/table XML in document order,
  inlining Word equations (`<m:oMath>`) as linearized `[EQ: ...]` text via
  `lib/omml.py` instead of the empty string `python-docx` returns for them.
  Each paragraph block carries `docx_paragraph_index` (its index into
  `document.paragraphs`, tables excluded) so an agent can jump straight to
  the right `Paragraph` object for `insert_paragraph_after` instead of
  re-matching on text or grepping the raw XML; each table block carries
  `docx_table_index` likewise. `summary.image_count` counts embedded images
  via the document's part relationships (inline *and* floating) — a nonzero
  count means real content (a figure, a data table baked into a screenshot)
  exists outside the extracted text entirely, since there's no OCR.
- **`extract_xlsx.py`** — loads every sheet (visible and hidden) with both
  `data_only=False` and `data_only=True` so each cell carries both its
  formula and cached value; also captures hidden rows/columns, merged
  ranges, cell comments, and per-sheet/total `image_count` (via
  `worksheet._images`) with the same "inspect it, it's not in the text"
  warning as docx.
- **`extract_xls.py`** — same shape of output for legacy `.xls` via `xlrd`,
  flagging `formula_available: false` since `xlrd` only exposes cached
  values. `xlrd` has no image API at all, so image detection here is a
  heuristic byte-scan of the raw OLE "Workbook" stream for BIFF drawing/
  object records (`lib/xls_drawings.py`), reported as `summary.
  likely_has_images` — a presence signal (not an exact count, and OBJ
  records also cover non-image objects), not a substitute for actually
  viewing the file (convert to `.xlsx` with LibreOffice and re-run
  `extract_xlsx.py` for that).
- **`extract_pdf.py`** — extracts text per page and flags pages whose text
  looks like a corrupted equation (Mathematical Alphanumeric Symbols,
  doubled characters); flagged pages get rendered to PNG under
  `.cache/renders/` for visual reading instead of trusting that text.
  Separately, every page also gets an exact `image_count` via
  `page.get_images()` — independent of the text-corruption check, since a
  page can hold a real embedded image (pasted work, a scanned figure) with
  otherwise perfectly clean surrounding text; any page with an image is
  flagged and rendered too.
- **`convert_doc.py`** — `.doc` → `.docx` via `pandoc` or LibreOffice; raises
  a clear `no_converter` error (instead of guessing) when neither is
  installed, matching `grader.md`'s "flag the file" instruction.
- **`compare_xlsx.py <submission> <solutions>`** — deterministic pre-check
  for spreadsheet cells where the solution's answer is a bare number:
  classifies each such cell `auto_correct` / `auto_incorrect` /
  `missing_answer`, tolerance-based. Works for `.xlsx` and `.xls` in any
  combination (each file's own extension picks its extractor); `.xls` cells
  never carry a formula, so they simply never get a `likely_downstream_of`
  hint, which is correct rather than a gap. Every other cell (text,
  conceptual, formula-as-the-answer, or a tab missing on one side) comes
  back `needs_review` — the grader/checker must still judge those, and must
  still write the actual explanation for anything flagged incorrect. An
  `auto_incorrect` cell whose formula references another `auto_incorrect`
  cell on the same sheet also gets `likely_downstream_of: [...]` — a hint
  that it may be one root cause cascading through several cells rather than
  several independent mistakes; still requires confirming, never skip
  tracing it yourself. This is a hint that saves LLM reasoning on
  unambiguous arithmetic, never a substitute for the required per-question
  judgment.

## Cache

`.cache/extraction/<sha256 of file>.json` — one record per distinct file
content, schema-versioned (`lib/cache.py`) so a script change invalidates
stale entries automatically. `.cache/renders/` holds PDF page images.
`.cache/converted/` holds `.doc` → `.docx` conversion output. The whole
`.cache/` directory is git-ignored and lives outside `student-submissions/`,
`reference-solutions/`, and `graded-submissions/` — it's derived, disposable
data, safe to delete at any time (everything just re-extracts on next use).

## Extending

If a script's output looks wrong or incomplete for a file it doesn't handle
well (an unusual equation shape, a spreadsheet layout the parser mishandles),
fall back to the manual method described in `grader.md`'s "Equation
Extraction" / "Spreadsheet Extraction" sections for that one file, and — if
it's a recurring pattern — fix the script so both agents benefit from it.
