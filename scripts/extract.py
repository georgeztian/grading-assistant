#!/usr/bin/env python3
"""Unified extraction entry point — the one command grader/checker agents
should run instead of writing ad hoc parsing code per submission.

Dispatches by file extension to the right extractor (.doc is converted to
.docx first), and always goes through the shared cache: since grader.md
requires grader and checker to use the *exact same* extraction method for
comparable results, a file is only ever parsed once — whoever asks second
(same submission re-checked, or the same solutions file used by another
student's grading) gets the cached record instead of re-parsing.

Usage:
    python extract.py <file>              # cache + print {cache_path, summary}
    python extract.py <file> --json       # also print the full record
    python extract.py <file> --force      # ignore existing cache entry

Caching is always on for this script — there is no opt-in `--cache` flag to
pass (that flag exists on the per-format extract_*.py scripts, which default
to a plain one-off dump when run standalone). `--cache` is still accepted
here as a harmless no-op, purely so old muscle memory / docs referring to
`extract.py <file> --cache` keep working rather than erroring.

Then read the JSON at cache_path (e.g. with the Read tool) instead of
re-extracting — that keeps the extraction record out of this command's own
output for files where it would be large.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache  # noqa: E402
from lib.doc_images import detect_doc_images  # noqa: E402
import convert_doc  # noqa: E402
import extract_docx  # noqa: E402
import extract_pdf  # noqa: E402
import extract_xls  # noqa: E402
import extract_xlsx  # noqa: E402

SUPPORTED = {".doc", ".docx", ".pdf", ".xlsx", ".xls"}


def extract_any(path: Path) -> dict:
    ext = path.suffix.lower()

    if ext == ".doc":
        converted = convert_doc.convert_doc_to_docx(
            path, cache.CACHE_DIR.parent / "converted"
        )
        record = extract_docx.extract_docx(converted)
        record["converted_from"] = str(path)
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
        return extract_docx.extract_docx(path)

    if ext == ".xlsx":
        return extract_xlsx.extract_xlsx(path)

    if ext == ".xls":
        return extract_xls.extract_xls(path)

    if ext == ".pdf":
        file_hash = cache.sha256_of(path)
        return extract_pdf.extract_pdf(path, cache.RENDER_DIR, file_hash)

    raise ValueError(f"unsupported file type: {ext}")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__,
                                      formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("file", type=Path)
    parser.add_argument("--force", action="store_true",
                         help="ignore any existing cache entry and re-extract")
    parser.add_argument("--json", action="store_true",
                         help="also print the full extracted record, not just the cache path")
    parser.add_argument("--cache", action="store_true",
                         help="no-op — extract.py always caches by default; this flag exists "
                              "only so the documented `extract.py <file> --cache` invocation "
                              "doesn't error (the per-format extract_*.py scripts use --cache "
                              "to opt IN to caching, since run standalone they default to a "
                              "plain one-off dump)")
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

    if not args.force:
        cached = cache.load_cached(path)
        if cached is not None:
            out = {
                "cache_hit": True,
                "cache_path": str(cache.cache_path_for(path)),
                "summary": cached.get("summary"),
            }
            if args.json:
                out["record"] = cached
            print(json.dumps(out, indent=2, ensure_ascii=False))
            return

    try:
        record = extract_any(path)
    except (RuntimeError, ValueError) as exc:
        out = {"error": str(exc)}
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

    cache_file = cache.save_cache(path, record)
    out = {
        "cache_hit": False,
        "cache_path": str(cache_file),
        "summary": record.get("summary"),
    }
    if args.json:
        out["record"] = record
    print(json.dumps(out, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
