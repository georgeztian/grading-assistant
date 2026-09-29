# Extraction fallback & image inspection (read only when needed)

The shared toolkit (`scripts/extract.py`) implements everything below and is always the first choice. Open this file only when:
- the extraction view looks wrong or incomplete for one file (an unusual equation shape, an unconventional spreadsheet layout) — use the manual method below for that one file (and, if it's a recurring pattern, the script should be fixed so everyone benefits — tell the user); or
- a view warning (`!` lines at its top, `(image)` tags, `FLAGGED` pages) says content exists outside the extracted text and you need to know how to look at it.

All ad hoc Python runs with `.venv/Scripts/python` (Windows; `.venv/bin/python` on macOS/Linux) — inline (`-c` / stdin) or under `.cache/inspect/`, never as a file elsewhere.

## Images and other non-text content

The extraction never OCRs or describes an image, so image content is never in the view text. Instead, every image is **exported to a PNG/JPG file whose path the view shows next to it**. Read that file with the Read tool. EMF/WMF pictures are converted to PNG first; these are common for equations in legacy `.doc` files, which store equations as pictures.
- **docx / .doc / pdf**: a paragraph holding a picture shows `(image: <path>)`. For an image the view can't place (inside a table, header or text box), the `!` warning gives an `unzip` command that extracts into `.cache/inspect/<content id>/` (images under `word/media/`). Always use that path, never one in the project root or an OS temp dir.
- **xlsx / .xls**: each sheet header lists `image at <anchor cell>: <path>`. Text boxes and charts are listed in the sheet as `textboxN` / `chartN` "(at <cell>)", with the box's text and the chart's type, title and series data.
- **pdf**: every page with an embedded image *or* a suspected corrupted equation is `FLAGGED`, with a rendered PNG of the whole page. The image check is independent of the equation check: a page can hold a real answer as an image with otherwise clean text. The base `.docx` rebuilt from the PDF carries such equations as pictures, each with its own `(image: …)` path.
- **.doc / .xls conversion**: legacy files are converted to `.docx` / `.xlsx` by Microsoft Word/Excel (Windows, via COM) or LibreOffice, so their images, formulas and formatting survive exactly.
  - For `.doc`, `doc_image_heuristic` is a raw-stream cross-check. A mismatch warning means the converter may have dropped an image, so open the original `.doc` to check.
  - If no converter is installed, a `.doc` is `unreadable`; note `likely_has_images` in the flag so a human knows image content awaits review.
  - If no converter is installed, a `.xls` falls back to a **DEGRADED** values-only reading with no formulas; the view says so at the top. Report it, since installing Excel or LibreOffice fixes it.
- **Charts** in Word documents appear on their paragraph as `(chart: …)`, with type, title and every series' data (read from the chart itself, not a picture). In PDFs, a chart or diagram drawn as vectors flags its page (`vector_graphics`, rendered to PNG). If it has colour or curves, it is also pictured in the rebuilt document as `(image: …)`, together with its labels.
- **Not captured:** the data inside embedded OLE objects (e.g. an embedded Excel sheet; only its preview picture is exported), and a greyscale vector diagram's shape in the rebuilt PDF document (its page is still flagged and rendered). Open the file directly when a question depends on one. Text boxes, form fields, content controls (`[sN]`) and nested tables *are* extracted: their text appears in the view.

## Equation extraction (manual method)

Plain `python-docx`/`pypdf` text extraction misses or corrupts equations.

- **.doc**: `python-docx` cannot open `.doc`. Convert it first with `.venv/Scripts/python scripts/convert_doc.py <file.doc>`, which uses Word or LibreOffice and writes to `.cache/converted/`. `pandoc` cannot read `.doc`. If no converter is available, flag the file rather than guess.
- **.docx**: Word equations are OMML (`<m:oMath>`, math namespace `http://schemas.openxmlformats.org/officeDocument/2006/math`), not `<w:t>` runs, so `paragraph.text` returns `""` for them.
  - Walk the paragraph XML with `lxml` (`paragraph._p.findall('.//{…/math}oMath')`) and read the `m:t` nodes in order.
  - Rebuild the structure while walking: `m:f` gives numerator/denominator, `m:sSub`/`m:sSup` give sub/superscripts, `m:rad` gives a root.
  - Alternatively, `pandoc file.docx -t markdown` renders OMML as inline LaTeX.
  - Never treat an empty-looking paragraph as blank without checking for `m:oMath`.
- **.pdf**: Word-exported equations are flattened to positioned glyphs. The text layer shows duplicated characters ("PV" → "PVPV"), Mathematical Alphanumeric Symbols (U+1D400–U+1D7FF), and scrambled fraction/subscript layout. Treat such lines as "equation — don't trust the text" and read them visually from the rendered page (PyMuPDF; `extract.py` already renders flagged pages to `.cache/renders/`). The base `.docx` the toolkit rebuilds from a PDF already carries such equations as pictures of their on-page region.

## Spreadsheet extraction (manual method)

Read every tab, not just the active one.

- **.xlsx** (`openpyxl`): iterate `workbook.sheetnames`, including hidden sheets (`sheet_state != 'visible'`) unless they are clearly scratch/unused.
  - **Formulas vs. values:** load twice. `data_only=False` gives the formula string, and `data_only=True` gives the cached value (which is `None` for a file never opened/saved in Excel). Keep both.
  - Check hidden rows/columns (`row_dimensions[n].hidden`, `column_dimensions[letter].hidden`) and cell comments.
  - Check `worksheet._images` / `worksheet._charts`. This needs Pillow, which is installed in `.venv`.
- **.xls** (`xlrd<2.0`, since 2.0+ dropped `.xls`): `xlrd.open_workbook(path).sheets()`. It exposes cached values only, never formulas, so don't treat a missing formula as absent content. It has no image API (see above).
- Never treat an empty-looking cell as blank without checking for merged-cell overflow or wrapped text from a neighbor.
