#!/usr/bin/env python3
"""Create (or refresh) the project's private Python environment at `.venv/`.

Run once per machine with any Python 3 interpreter:

    python scripts/setup_env.py

Stdlib only, so it works before any dependency is installed. Idempotent:
re-running reuses an existing `.venv/` and just re-syncs `requirements.txt`
into it. Every other script in `scripts/` — and any ad hoc Python an agent
writes (e.g. python-docx/openpyxl annotation code) — must then run with the
venv's interpreter: `.venv/Scripts/python` on Windows, `.venv/bin/python` on
macOS/Linux (see `_venv.py`).

Because this project lives in a Dropbox folder, the new `.venv/` is marked
"ignored" for Dropbox sync (a best-effort no-op when Dropbox isn't in use),
so thousands of interpreter/package files never get uploaded or conflict
across machines — each machine builds its own.
"""
from __future__ import annotations

import os
import subprocess
import sys
import venv
from pathlib import Path

from _venv import ROOT, VENV_DIR, venv_python

REQUIREMENTS = ROOT / "requirements.txt"


def mark_dropbox_ignored(path: Path) -> None:
    """Best-effort: tell Dropbox not to sync `path` (per Dropbox's documented
    ignore mechanism — an NTFS alternate data stream on Windows, an xattr on
    macOS/Linux). Silently skipped if unsupported."""
    try:
        if os.name == "nt":
            with open(f"{path}:com.dropbox.ignored", "w") as fh:
                fh.write("1")
        elif sys.platform == "darwin":
            subprocess.run(["xattr", "-w", "com.dropbox.ignored", "1", str(path)],
                           check=True, capture_output=True)
        else:
            subprocess.run(["attr", "-s", "com.dropbox.ignored", "-V", "1", str(path)],
                           check=True, capture_output=True)
    except (OSError, subprocess.CalledProcessError):
        pass


def main() -> int:
    if sys.version_info < (3, 9):
        print("error: Python 3.9+ is required to create the environment", file=sys.stderr)
        return 1

    py = venv_python()
    if py.exists():
        print(f"Reusing existing environment: {VENV_DIR}")
    else:
        print(f"Creating environment: {VENV_DIR}")
        VENV_DIR.mkdir()
        # Mark before populating so Dropbox never starts uploading it.
        mark_dropbox_ignored(VENV_DIR)
        venv.EnvBuilder(with_pip=True).create(VENV_DIR)
    mark_dropbox_ignored(VENV_DIR)

    subprocess.run([str(py), "-m", "pip", "install", "--upgrade", "pip", "--quiet"], check=True)
    subprocess.run([str(py), "-m", "pip", "install", "-r", str(REQUIREMENTS)], check=True)

    print(f"\nDone. Project interpreter: {py.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
