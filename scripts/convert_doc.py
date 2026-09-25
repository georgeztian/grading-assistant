#!/usr/bin/env python3
"""Convert a legacy .doc file to .docx so extract_docx.py can read it.

python-docx cannot open .doc (OOXML-only), and neither can pandoc (it reads
.docx but not the legacy binary .doc format), so this uses LibreOffice's
`soffice --headless` — looked up on PATH and in its standard install
locations. If LibreOffice isn't installed, it fails loudly with a clear
"flag this file" message instead of guessing at the file's content.

Output goes to `<out-dir>/<sha256 of the .doc>/<stem>.docx`, so two different
files that share a name (e.g. `HW1/john.doc` and `HW2/john.doc`) never
overwrite each other's converted copy.

Usage:
    .venv/Scripts/python scripts/convert_doc.py <file.doc> [--out-dir DIR]
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache  # noqa: E402

# LibreOffice is often installed without being put on PATH (always so on
# Windows), so also check its standard install locations.
_SOFFICE_CANDIDATES = [
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
]


def find_converter() -> str | None:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in _SOFFICE_CANDIDATES:
        if candidate.exists():
            return str(candidate)
    return None


def convert_doc_to_docx(path: Path, out_dir: Path) -> Path:
    converter = find_converter()
    if converter is None:
        raise RuntimeError(
            "no_converter: LibreOffice (soffice) is not installed — flag "
            f"{path.name} for manual conversion rather than guessing at its "
            "content."
        )

    target_dir = out_dir / cache.sha256_of(path)
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / (path.stem + ".docx")

    try:
        subprocess.run(
            [converter, "--headless", "--convert-to", "docx",
             "--outdir", str(target_dir), str(path)],
            check=True, capture_output=True, timeout=300,
        )
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(f"conversion_failed: {converter} failed on {path.name}: {exc}") from exc

    if not target.exists():
        raise RuntimeError(f"conversion_failed: {converter} did not produce {target}")
    return target


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__,
                                      formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("file", type=Path)
    parser.add_argument("--out-dir", type=Path,
                         default=Path(__file__).resolve().parents[1] / ".cache" / "converted")
    args = parser.parse_args()

    if not args.file.exists():
        print(json.dumps({"error": f"file not found: {args.file}"}))
        sys.exit(1)

    try:
        result_path = convert_doc_to_docx(args.file, args.out_dir)
    except RuntimeError as exc:
        print(json.dumps({"error": str(exc)}))
        sys.exit(1)

    print(json.dumps({"converted_path": str(result_path), "converter": find_converter()}))


if __name__ == "__main__":
    main()
