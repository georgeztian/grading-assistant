#!/usr/bin/env python3
"""Convert a legacy .doc file to .docx so extract_docx.py can read it.

python-docx cannot open .doc (OOXML-only). This tries `pandoc` first, then
LibreOffice's `soffice --headless`, matching grader.md's converter list. If
neither is installed, it fails loudly with a clear "flag this file" message
instead of guessing at the file's content.

Usage:
    python convert_doc.py <file.doc> [--out-dir DIR]
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path


def find_converter() -> str | None:
    for name in ("pandoc", "soffice", "libreoffice"):
        if shutil.which(name):
            return name
    return None


def convert_doc_to_docx(path: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    converter = find_converter()
    target = out_dir / (path.stem + ".docx")

    if converter is None:
        raise RuntimeError(
            "no_converter: neither pandoc nor LibreOffice (soffice) is "
            f"installed — flag {path.name} for manual conversion rather "
            "than guessing at its content."
        )

    if converter == "pandoc":
        subprocess.run(
            ["pandoc", str(path), "-o", str(target)],
            check=True, capture_output=True,
        )
    else:  # soffice / libreoffice
        subprocess.run(
            [converter, "--headless", "--convert-to", "docx",
             "--outdir", str(out_dir), str(path)],
            check=True, capture_output=True,
        )

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
    except (RuntimeError, subprocess.CalledProcessError) as exc:
        print(json.dumps({"error": str(exc)}))
        sys.exit(1)

    print(json.dumps({"converted_path": str(result_path), "converter": find_converter()}))


if __name__ == "__main__":
    main()
