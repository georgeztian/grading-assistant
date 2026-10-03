"""Shared cache: locations, content hashing, atomic writes.

Every extractor produces a deterministic JSON record for a given input file,
so a file is extracted once and every later request — the checker re-reading
a submission, the next student's grader reading the same solutions file — is
served the cached record (see extract.get_record, the only reader/writer of
extraction records).

Everything lives under `.cache/`, outside the three data folders, so it is
never mistaken for graded output:
    extraction/  <sha256>.json record (shared by identical files) +
                 <sha256>.<path8>.view.txt compact view (one per file path)
    converted/   <sha256>/ — .doc/.xls conversions and PDF base .docx files
    renders/     <sha256>/ — PDF page renders and exported embedded images
    rubrics/     <sha256>/ — answer keys and their verification status
    work/        <sha12>-<path8>/ — per-file agent scratch (verdicts, audit reports)
    inspect/     ad hoc manual inspection
"""
from __future__ import annotations

import hashlib
import json
import os
import time
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[2]
CACHE_DIR = REPO_ROOT / ".cache" / "extraction"
RENDER_DIR = REPO_ROOT / ".cache" / "renders"
CONVERTED_DIR = REPO_ROOT / ".cache" / "converted"
RUBRIC_DIR = REPO_ROOT / ".cache" / "rubrics"
WORK_DIR = REPO_ROOT / ".cache" / "work"

# Bump when an extractor's output schema/logic changes, to invalidate stale
# cache entries produced by an older version of the script.
# v2: docx_paragraph_index/docx_table_index and image_count/image_warning.
# v3: per-page image_count for pdf; drawing-record heuristic for xls.
# v4: raw-OLE image-signature heuristic for .doc.
# v5: image_warning gives a concrete unzip command into .cache/inspect/.
# v6: conversions under .cache/converted/<sha256>/; xlsx chart_count; xls comments.
# v7: docx has_image; pdf base_docx_path + annotation_blocks; compact views.
# v8: .xls extracted through its .xlsx conversion; embedded images exported.
# v9: body-level content controls (sN) and nested-table text extracted; PDF
#     renders under renders/<sha>/; fewer false "doubled character" flags;
#     OMML delimiters/n-ary defaults; redundant per-block fields dropped.
# v10: docx charts summarized (`charts`); xlsx text boxes + charts read from
#      drawing parts; PDF vector graphics flagged and pictured.
# v11: records name no file (shared by identical submissions); views are per
#      file path; conversions/PDF bases have content-neutral names.
# v12: an xlsx text cell starting with "=" is text, not a formula.
# v13: a converted .doc/.xls record keeps only the original's extension
#      (`converted_from`), never the path of whoever extracted it first.
SCHEMA_VERSION = 13

_sha_memo: dict[tuple, str] = {}


def sha256_of(path: Path) -> str:
    """Content hash, memoized per (path, size, mtime) within a process."""
    st = os.stat(path)
    key = (str(Path(path).resolve()), st.st_size, st.st_mtime_ns)
    if key not in _sha_memo:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        _sha_memo[key] = h.hexdigest()
    return _sha_memo[key]


def path_key(path) -> str:
    """Short, stable key for a file's *location* (not content). Anything
    per-student — a view naming the file, a scratch folder — is keyed by
    content + path_key, so two students who hand in byte-identical files
    each get their own, named after their own file."""
    return hashlib.sha256(str(Path(path).resolve()).lower().encode("utf-8")).hexdigest()[:8]


def rel_path(path) -> str:
    """Repo-relative, forward-slash path for display (absolute if outside)."""
    try:
        return Path(path).resolve().relative_to(REPO_ROOT).as_posix()
    except ValueError:
        return str(path)


def replace_with_retry(tmp: Path, target: Path, attempts: int = 8) -> None:
    """os.replace, retried: on Windows a target briefly held open by
    Dropbox's sync client, an antivirus scan, or another agent reading it
    raises PermissionError — transient, so back off and retry."""
    for attempt in range(attempts):
        try:
            os.replace(tmp, target)
            return
        except PermissionError:
            if attempt == attempts - 1:
                raise
            time.sleep(0.05 * 2 ** attempt)


def write_text_atomic(target: Path, text: str) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    # pid-unique temp name: concurrent agents writing the same file must not
    # write through each other's half-finished temp file.
    tmp = target.with_name(f"{target.name}.{os.getpid()}.tmp")
    tmp.write_bytes(text.encode("utf-8"))  # no newline translation: hashes must match
    replace_with_retry(tmp, target)


def cache_path_for(path: Path) -> Path:
    return CACHE_DIR / f"{sha256_of(path)}.json"


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


def save_cache(path: Path, record: dict[str, Any]) -> Path:
    record = dict(record)
    # The record is shared by every file with this content, so it names no
    # file: per-file details (the name in a view header) are added when the
    # view for a particular path is rendered.
    record["schema_version"] = SCHEMA_VERSION
    record["source_sha256"] = sha256_of(path)
    cache_file = cache_path_for(path)
    write_text_atomic(cache_file, json.dumps(record, indent=1, ensure_ascii=False))
    return cache_file


def work_dir_for(path: Path) -> Path:
    """Scratch folder for one submission *file*: keyed by content and path,
    so two students who hand in byte-identical files (a copied submission,
    an untouched template) never share — and overwrite — each other's
    verdicts."""
    d = WORK_DIR / f"{sha256_of(path)[:12]}-{path_key(path)}"
    d.mkdir(parents=True, exist_ok=True)
    return d
