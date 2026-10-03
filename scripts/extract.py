#!/usr/bin/env python3
"""Unified extraction entry point — the one command every agent runs to read
a submission or solutions file, instead of writing ad hoc parsing code.

Dispatches by file extension (.doc/.xls are converted to .docx/.xlsx first;
a .pdf is also rebuilt as the base .docx its graded copy is written from),
exports embedded images, and always goes through the shared cache keyed by
content hash: a file is parsed once, and whoever asks next (the checker
re-reading a submission, the next student's grader reading the same
solutions file) gets the cached record.

Usage:
    .venv/Scripts/python scripts/extract.py <file>          # print {view_path, cache_path, work_dir, summary}
    .venv/Scripts/python scripts/extract.py <file> --json   # also print the full record
    .venv/Scripts/python scripts/extract.py <file> --force  # ignore the cached record

Read the compact text view at view_path (e.g. with the Read tool): one line
per paragraph / table row / cell, with the ids used as annotation anchors —
the same content as the JSON record at cache_path in far fewer tokens.
work_dir is this file's scratch folder (.cache/work/<sha12>-<path8>/, keyed by
content and path) for the agent-written JSON the other scripts take
(verdicts, blind verdicts, discrepancies) and the audit report.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache  # noqa: E402
from lib.doc_images import detect_doc_images  # noqa: E402
import convert_doc  # noqa: E402
import extract_docx  # noqa: E402
import extract_pdf  # noqa: E402
import extract_xls  # noqa: E402
import extract_xlsx  # noqa: E402
from lib import views  # noqa: E402
from lib.pdf_docx import build_base_docx  # noqa: E402

SUPPORTED = {".doc", ".docx", ".pdf", ".xlsx", ".xls"}


def extract_any(path: Path) -> dict:
    ext = path.suffix.lower()
    file_hash = cache.sha256_of(path)
    # Page renders and exported images go here and are listed in the view.
    image_dir = cache.RENDER_DIR / file_hash

    if ext == ".doc":
        converted = convert_doc.convert(path, cache.CONVERTED_DIR, "docx")
        # -> .cache/converted/<sha256>/converted.docx (Word via COM, or LibreOffice)
        record = extract_docx.extract_docx(converted, image_dir)
        record["converted_from"] = ext  # the record is shared: no file name
        record["converter_output_path"] = str(converted)

        heuristic = detect_doc_images(path)
        record["doc_image_heuristic"] = heuristic
        if heuristic.get("likely_has_images") and not record["summary"].get("image_count"):
            record["summary"]["image_warning"] = (
                (record["summary"].get("image_warning") or "").rstrip() + " "
                "Additionally, a raw-file heuristic scan of the original "
                ".doc found signs of an embedded image/object "
                f"(signatures: {heuristic['image_signatures_found']}, "
                f"ObjectPool present: {heuristic['object_pool_present']}) "
                "that the converted .docx does NOT show (image_count: 0) — "
                "the converter may have dropped it; open the original .doc "
                "directly to check if this seems to leave a question "
                "unanswered."
            ).strip()
        return record

    if ext == ".docx":
        return extract_docx.extract_docx(path, image_dir)

    if ext == ".xlsx":
        return extract_xlsx.extract_xlsx(path, image_dir)

    if ext == ".xls":
        # Convert to .xlsx (Excel via COM, or LibreOffice) so formulas,
        # formatting and images are available — xlrd alone sees cached
        # values only. The graded copy is written from the same converted
        # workbook, so cell references line up exactly.
        try:
            converted = convert_doc.convert(path, cache.CONVERTED_DIR, "xlsx")
        except RuntimeError as exc:
            if not str(exc).startswith("no_converter"):
                raise
            record = extract_xls.extract_xls(path)
            record["degraded"] = True  # re-extracted once a converter exists
            record["warnings"].insert(0, (
                "DEGRADED: no .xls converter (Microsoft Excel or LibreOffice) is installed, so "
                "this is the values-only xlrd reading — formulas are unavailable and the graded "
                "copy can only be rebuilt from values. Install Excel or LibreOffice for full fidelity."))
            return record
        record = extract_xlsx.extract_xlsx(converted, image_dir)
        record["converted_from"] = ext
        record["converter_output_path"] = str(converted)
        return record

    if ext == ".pdf":
        record = extract_pdf.extract_pdf(path, image_dir)
        # The graded copy of a PDF is a .docx; build its base once, in code,
        # and expose that document's paragraphs as the annotation anchors.
        base = cache.CONVERTED_DIR / file_hash / "base.docx"  # shared by identical PDFs
        page_of_paragraph = build_base_docx(path, base)
        blocks = extract_docx.extract_docx(base, image_dir)["blocks"]
        for block in blocks:  # the rebuilt document holds only paragraphs
            block["page"] = page_of_paragraph[int(block["id"][1:])]
        record["base_docx_path"] = str(base)
        record["annotation_blocks"] = blocks
        record["summary"]["base_docx_paragraphs"] = len(page_of_paragraph)
        return record

    raise ValueError(f"unsupported file type: {ext}")


def view_path_for(cache_file: Path, path: Path) -> Path:
    """One view per file path (the record itself is shared by content), so
    each student's view is headed with — and refers to — their own file."""
    return cache_file.with_name(f"{cache_file.stem}.{cache.path_key(path)}.view.txt")


def get_record(path: Path, force: bool = False) -> tuple[dict, Path, Path, bool]:
    """Cached extraction for `path` plus its compact view (written if absent).
    Returns (record, cache_path, view_path, cache_hit). Raises on failure."""
    record = None if force else cache.load_cached(path)
    if record is not None and record.get("degraded") and convert_doc.available_converters("xlsx"):
        record = None  # a converter is available now — replace the values-only reading
    hit = record is not None
    if record is None:
        record = extract_any(path)
        cache_file = cache.save_cache(path, record)
    else:
        cache_file = cache.cache_path_for(path)
    view = view_path_for(cache_file, path)
    if not hit or not view.exists():
        cache.write_text_atomic(view, views.render_view(record, cache.rel_path(path)))
    return record, cache_file, view, hit


def source_docx_for(path: Path) -> Path:
    """The .docx a document submission's graded copy is written from — the
    file itself, its Word/LibreOffice conversion (.doc), or its rebuilt base
    (.pdf). Its paragraph indices are the view's [pN] ids."""
    ext = path.suffix.lower()
    if ext == ".docx":
        return path
    record, *_ = get_record(path)
    key = "converter_output_path" if ext == ".doc" else "base_docx_path"
    source = Path(record[key])
    if not source.exists():  # cache partly cleaned — rebuild
        record, *_ = get_record(path, force=True)
        source = Path(record[key])
    return source


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__,
                                      formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("file", type=Path)
    parser.add_argument("--force", action="store_true",
                         help="ignore any existing cache entry and re-extract")
    parser.add_argument("--json", action="store_true",
                         help="also print the full extracted record, not just the paths")
    args = parser.parse_args()

    path = args.file
    if not path.exists():
        print(json.dumps({"error": f"file not found: {path}"}))
        sys.exit(1)

    ext = path.suffix.lower()
    if ext not in SUPPORTED:
        print(json.dumps({"error": f"unsupported file type: {ext}",
                           "supported": sorted(SUPPORTED)}))
        sys.exit(1)

    try:
        record, cache_file, view, hit = get_record(path, force=args.force)
    except Exception as exc:  # failed conversion, corrupt/unreadable file, ...
        # Always a JSON error, never a bare traceback, so the grader can
        # report the file as unreadable.
        if isinstance(exc, (RuntimeError, ValueError)):
            out = {"error": str(exc)}
        else:
            out = {"error": f"unreadable: {type(exc).__name__}: {exc}"}
        if ext == ".doc":
            # Conversion failed (e.g. no_converter) — text extraction can't
            # proceed, but the image heuristic only needs the raw OLE file,
            # not a converter, so it still runs and is worth surfacing.
            # Not cached: unlike file content, converter availability can
            # change between runs, so a stale failure must not be served
            # once a converter is installed.
            try:
                out["doc_image_heuristic"] = detect_doc_images(path)
            except Exception:
                pass
        print(json.dumps(out))
        sys.exit(1)

    out = {
        "cache_hit": hit,
        "view_path": str(view),
        "cache_path": str(cache_file),
        "work_dir": str(cache.work_dir_for(path)),
        "summary": record.get("summary"),
    }
    if args.json:
        out["record"] = record
    print(json.dumps(out, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
