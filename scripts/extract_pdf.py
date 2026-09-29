#!/usr/bin/env python3
"""Extract text from a PDF, flag pages whose text layer likely corrupted an
equation, and flag pages containing embedded raster images, per the equation caveats in
.claude/skills/grading-instructions/extraction-fallback.md.

Word-exported PDFs flatten equations to positioned glyphs with no math
markup — plain text extraction commonly emits Unicode Mathematical
Alphanumeric Symbols (U+1D400-U+1D7FF) or doubled characters ("PV" ->
"PVPV") instead of the real expression. Rather than trust that text, this
renders any flagged page to a PNG (once, cached) so the calling agent can
read the equation visually instead of from the corrupted text layer.

Separately — and independent of that text-corruption check — every page is
also checked for embedded raster images (a pasted screenshot of handwritten
work, a scanned figure) via `page.get_images()`. A page can contain a real
answer as an image with otherwise perfectly clean surrounding text, so this
check does NOT depend on the text looking suspicious; any page with an
embedded image is flagged and rendered too.

Normally used through extract.py (cached; it also rebuilds the PDF as the
base .docx the graded copy is written from). Run standalone it just prints
the record.

Usage:
    .venv/Scripts/python scripts/extract_pdf.py <file.pdf>
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

import pymupdf
from pypdf import PdfReader

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache  # noqa: E402
from lib.pdf_docx import graphic_regions  # noqa: E402

MATH_ALPHANUMERIC_RE = re.compile("[\U0001D400-\U0001D7FF]")
# A short token immediately repeated with no separator: "PVPV", "12001200".
DOUBLED_TOKEN_RE = re.compile(r"\b(\w{2,8})\1\b")


def flag_reasons(text: str) -> list[str]:
    reasons = []
    if MATH_ALPHANUMERIC_RE.search(text):
        reasons.append("mathematical_alphanumeric_symbols")
    # Ignore short all-digit repeats — ordinary numbers like 2020 or 1010 —
    # which would otherwise flag (and render) nearly every page.
    if any(not (m.group(1).isdigit() and len(m.group(1)) <= 2)
           for m in DOUBLED_TOKEN_RE.finditer(text)):
        reasons.append("doubled_characters")
    return reasons


def extract_pdf(path: Path, render_dir: Path | None = None) -> dict:
    """`render_dir`: where flagged pages are rendered to PNG (page<N>.png);
    without it nothing is rendered."""
    reader = PdfReader(str(path))
    doc = pymupdf.open(str(path))

    pages = []
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        reasons = flag_reasons(text)
        image_count = len(doc[i].get_images(full=True))
        if image_count:
            reasons.append("embedded_image")
        if graphic_regions(doc[i])[1]:  # a chart/diagram drawn as vectors, not an image
            reasons.append("vector_graphics")

        render_path = None
        if reasons and render_dir is not None:
            render_dir.mkdir(parents=True, exist_ok=True)
            target = render_dir / f"page{i + 1}.png"
            if not target.exists():
                doc[i].get_pixmap(dpi=200).save(str(target))
            render_path = cache.rel_path(target)

        pages.append({"text": text, "flagged_reasons": reasons,
                      "image_count": image_count, "render_path": render_path})
    doc.close()

    flagged = sum(1 for pg in pages if pg["flagged_reasons"])
    return {
        "type": "pdf",
        "pages": pages,
        "summary": {
            "page_count": len(pages),
            "flagged_pages": flagged,
            "pages_with_images": sum(1 for pg in pages if pg["image_count"]),
            "image_count": sum(pg["image_count"] for pg in pages),
            "note": (
                "FLAGGED pages (suspected corrupted equation text, an embedded image, or a "
                "vector chart/diagram) "
                "have a rendered PNG — read it rather than trusting the text: a page can hold "
                "a real answer as an image while its surrounding text looks normal."
            ) if flagged else None,
        },
    }


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) != 2 or not Path(sys.argv[1]).exists():
        print(json.dumps({"error": "usage: extract_pdf.py <existing file.pdf>"}))
        sys.exit(1)
    print(json.dumps(extract_pdf(Path(sys.argv[1])), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
