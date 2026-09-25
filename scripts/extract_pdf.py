#!/usr/bin/env python3
"""Extract text from a PDF, flag pages whose text layer likely corrupted an
equation, and flag pages containing embedded raster images, per grader.md's
"Equation Extraction" caveats.

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

Usage:
    .venv/Scripts/python scripts/extract_pdf.py <file.pdf>
    .venv/Scripts/python scripts/extract_pdf.py <file.pdf> --cache
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

import pymupdf
from pypdf import PdfReader

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache  # noqa: E402

MATH_ALPHANUMERIC_RE = re.compile(r"[\U0001D400-\U0001D7FF]")
# A short word/token immediately repeated with no separator: "PVPV", "12001200"
DOUBLED_TOKEN_RE = re.compile(r"\b(\w{2,8})\1\b")


def flag_reasons(text: str) -> list[str]:
    reasons = []
    if MATH_ALPHANUMERIC_RE.search(text):
        reasons.append("mathematical_alphanumeric_symbols")
    if DOUBLED_TOKEN_RE.search(text):
        reasons.append("doubled_characters")
    return reasons


def extract_pdf(path: Path, render_dir: Path, file_hash: str) -> dict:
    reader = PdfReader(str(path))
    doc = pymupdf.open(str(path))

    pages = []
    flagged_count = 0
    pages_with_images = 0
    total_images = 0
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        reasons = flag_reasons(text)

        image_count = len(doc[i].get_images(full=True))
        if image_count:
            pages_with_images += 1
            total_images += image_count
            reasons = reasons + ["embedded_image"]

        render_path = None
        if reasons:
            flagged_count += 1
            render_dir.mkdir(parents=True, exist_ok=True)
            render_path = render_dir / f"{file_hash}_p{i}.png"
            if not render_path.exists():
                pix = doc[i].get_pixmap(dpi=200)
                pix.save(str(render_path))

        pages.append({
            "index": i,
            "text": text,
            "suspicious": bool(reasons),
            "flagged_reasons": reasons,
            "image_count": image_count,
            "render_path": str(render_path) if render_path else None,
        })

    doc.close()

    return {
        "type": "pdf",
        "pages": pages,
        "summary": {
            "page_count": len(pages),
            "flagged_pages": flagged_count,
            "pages_with_images": pages_with_images,
            "image_count": total_images,
            "note": (
                "Pages with suspicious text or embedded images "
                "(flagged_reasons non-empty) have a rendered PNG at "
                "render_path — read that image directly rather than "
                "trusting the text field for any equation OR any content "
                "that might be an embedded image (pasted work, a scanned "
                "figure) on that page; a page can hold a real answer as an "
                "image while its surrounding text looks perfectly normal."
            ) if flagged_count else None,
        },
    }


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__,
                                      formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("file", type=Path)
    parser.add_argument("--cache", action="store_true")
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    if not args.file.exists():
        print(json.dumps({"error": f"file not found: {args.file}"}))
        sys.exit(1)

    if args.cache and not args.force:
        cached = cache.load_cached(args.file)
        if cached is not None:
            print(json.dumps({
                "cache_hit": True,
                "cache_path": str(cache.cache_path_for(args.file)),
                "summary": cached.get("summary"),
            }, indent=2))
            return

    file_hash = cache.sha256_of(args.file)
    record = extract_pdf(args.file, cache.RENDER_DIR, file_hash)

    if args.cache:
        cache_file = cache.save_cache(args.file, record)
        print(json.dumps({
            "cache_hit": False,
            "cache_path": str(cache_file),
            "summary": record["summary"],
        }, indent=2))
    else:
        print(json.dumps(record, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
