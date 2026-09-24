"""Shared extraction cache.

Every extractor produces a deterministic JSON record for a given input file.
Since grader and checker are required to use the exact same extraction method
(see grader.md), there is no reason to ever run that extraction twice for the
same file: the grader computes it once, the checker (and any re-run) reuses
the cached record instead of re-parsing.

Cache lives outside the three protected data folders so it is never mistaken
for graded output and never needs `graded-submissions/` write access.
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[2]
CACHE_DIR = REPO_ROOT / ".cache" / "extraction"
RENDER_DIR = REPO_ROOT / ".cache" / "renders"

# Bump when an extractor's output schema/logic changes, to invalidate stale
# cache entries produced by an older version of the script.
# v2: added docx_paragraph_index/docx_table_index and image_count/image_warning
# (docx + xlsx) so agents don't have to re-open source files to locate
# annotation anchors or discover image-only content.
# v3: added per-page image_count to extract_pdf.py (independent of the
# text-corruption heuristic) and a drawing-record heuristic to extract_xls.py
# (summary.likely_has_images), closing the same image blind spot for pdf/xls.
# v4: added a raw-OLE-stream image-signature heuristic for .doc
# (doc_image_heuristic in extract.py's .doc branch, lib/doc_images.py) as a
# cross-check against the docx-converter's image_count and a fallback signal
# when no converter is installed at all.
# v5: docx/xlsx image_warning now gives a concrete, shell-safe unzip command
# into .cache/inspect/<file stem>/ instead of an ad hoc /tmp/x path — a
# prior run followed the vague version and left a stray folder in the
# project root.
SCHEMA_VERSION = 5


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def cache_path_for(path: Path) -> Path:
    digest = sha256_of(path)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return CACHE_DIR / f"{digest}.json"


def load_cached(path: Path) -> dict[str, Any] | None:
    cache_file = cache_path_for(path)
    if not cache_file.exists():
        return None
    try:
        with open(cache_file, "r", encoding="utf-8") as f:
            record = json.load(f)
    except (json.JSONDecodeError, OSError):
        return None
    if record.get("schema_version") != SCHEMA_VERSION:
        return None
    return record


def get_extraction(path: Path, extractor_fn, force: bool = False) -> dict[str, Any]:
    """Return the cached extraction record for `path`, running
    `extractor_fn(path) -> dict` and caching the result only on a miss."""
    if not force:
        cached = load_cached(path)
        if cached is not None:
            return cached
    record = extractor_fn(path)
    save_cache(path, record)
    return record


def save_cache(path: Path, record: dict[str, Any]) -> Path:
    record = dict(record)
    record["schema_version"] = SCHEMA_VERSION
    record["source_path"] = str(path)
    record["source_sha256"] = sha256_of(path)
    cache_file = cache_path_for(path)
    tmp_file = cache_file.with_suffix(".json.tmp")
    with open(tmp_file, "w", encoding="utf-8") as f:
        json.dump(record, f, indent=2, ensure_ascii=False)
    os.replace(tmp_file, cache_file)
    return cache_file
